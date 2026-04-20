"""使用遺傳算法計算工廠訂單排程，目標為最小化總加班分鐘數。"""

from __future__ import annotations

import argparse
import csv
import math
import random
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Iterable


EXCEL_EPOCH = datetime(1899, 12, 30)
XML_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

LABOR_CONFIG = {
    "manufacturing": {"people": 2, "regular_minutes": 2 * 8 * 60, "max_minutes": 2 * 10 * 60},
    "bottling": {"people": 2, "regular_minutes": 2 * 8 * 60, "max_minutes": 2 * 10 * 60},
    "filling": {"people": 5, "regular_minutes": 5 * 8 * 60, "max_minutes": 5 * 10 * 60},
}

OVERTIME_TARGETS = {
    "manufacturing": 2500.0,
    "filling": 2585.0,
}
TARGET_EXCESS_WEIGHT = 1_000.0
WEEKEND_OVERTIME_WEIGHT = 5.0
MANUFACTURING_SWITCH_CLEANING_WEIGHT = 10.0
PACKAGING_TEAM = {
    "regular_minutes": 5 * 8 * 60,
    "max_minutes": 5 * 10 * 60,
    "monthly_overtime_limit": 5 * 46 * 60,
}
BOTTLING_MAX_PEOPLE = 2
PROCESS_OVERTIME_MACHINES = {
    "bottling": "排瓶機",
    "filling": "充填機",
}
PROCESS_DAILY_REGULAR_MINUTES = 8 * 60
PROCESS_OVERTIME_REGULAR_MINUTES = {
    "bottling": 8 * 60,
    "filling": 8 * 60,
}
SLOT_MINUTES = 5

MACHINE_DAILY_CAPACITY = 10 * 60
MACHINE_COUNTS = {
    "預攪機": 1,
    "精機": 1,
    "小乳機": 1,
    "大乳機": 1,
    "大機": 1,
    "排瓶機": 1,
    "充填機": 1,
}

MACHINE_LABEL_TRANSLATIONS = {
    "預攪機": "Premix Mixer",
    "精機": "Essence Mixer",
    "小乳機": "Small Emulsion Mixer",
    "大乳機": "Large Emulsion Mixer",
    "大機": "Large Tank Mixer",
    "排瓶機": "Bottle Preparation Station",
    "充填機": "Filling Machine",
    "大型攪拌機": "Large Mixer",
    "小型攪拌機": "Small Mixer",
    "50L真空乳化機": "50L Vacuum Emulsifier",
    "100L真空乳化機": "100L Vacuum Emulsifier",
    "100L升溫桶槽": "100L Heating Tank",
    "400L大桶槽": "400L Tank",
    "500L大桶槽": "500L Tank",
    "落地型均質機": "Floor Homogenizer",
    "超音波震盪機": "Ultrasonic Mixer",
    "泵浦式充填機": "Pump Filling Machine",
    "定量缸式充填機": "Piston Filling Machine",
    "大定量缸式充填機": "Large Piston Filling Machine",
    "噴印機": "Printer",
    "噴印機-小": "Small Printer",
    "面膜機": "Mask Machine",
    "烘箱": "Oven",
    "自動鎖蓋機": "Auto Capping Machine",
}


@dataclass(frozen=True)
class ManufacturingItem:
    name: str
    process_name: str
    mix_machine: str
    batches: int
    premix_minutes_per_batch: float
    mix_clean_minutes_per_batch: float


@dataclass(frozen=True)
class FillingItem:
    bottling_machine: str
    filling_machine: str
    bottling_machine_minutes: float
    filling_machine_minutes: float
    bottling_people: int
    filling_people: int


@dataclass(frozen=True)
class Order:
    order_id: str
    product_name: str
    quantity: int
    due_date: date
    raw_ready_date: date
    material_ready_date: date
    manufacturing_items: tuple[ManufacturingItem, ...]
    filling_item: FillingItem
    source_order_id: str = ""
    source_product_code: str = ""
    source_core_code: str = ""
    manufacturing_source_override: bool = False


@dataclass(frozen=True)
class OrderAnomaly:
    order_id: str
    product_name: str
    severity: str
    anomaly_type: str
    message: str
    exclude_from_optimization: bool


@dataclass(frozen=True)
class ScheduledTask:
    order_id: str
    task_type: str
    detail: str
    scheduled_date: date
    machine_type: str
    machine_minutes: float
    labor_group: str
    labor_minutes: float


@dataclass
class ScheduleResult:
    total_overtime_minutes: float
    infeasible_penalty: float
    target_penalty: float = 0.0
    overtime_by_group: dict[str, float] = field(default_factory=dict)
    weekend_overtime_minutes: float = 0.0
    task_records: list[ScheduledTask] = field(default_factory=list)
    daily_labor_usage: dict[tuple[str, date], float] = field(default_factory=dict)
    daily_machine_usage: dict[tuple[str, date], float] = field(default_factory=dict)

    @property
    def objective(self) -> float:
        return self.total_overtime_minutes + self.infeasible_penalty + self.target_penalty


@dataclass(frozen=True)
class ScheduleValidationIssue:
    order_id: str
    issue_type: str
    severity: str
    message: str


def build_schedule_summary(schedule: ScheduleResult) -> dict[str, float]:
    total_bottling_minutes = sum(
        task.machine_minutes for task in schedule.task_records if task.task_type == "Cat3"
    )
    total_filling_minutes = sum(
        task.machine_minutes for task in schedule.task_records if task.task_type == "Cat4"
    )
    return {
        "total_overtime_minutes": schedule.total_overtime_minutes,
        "manufacturing_overtime_minutes": schedule.overtime_by_group.get("manufacturing", 0.0),
        "bottling_overtime_minutes": schedule.overtime_by_group.get("bottling", 0.0),
        "filling_overtime_minutes": schedule.overtime_by_group.get("filling", 0.0),
        "weekend_overtime_minutes": schedule.weekend_overtime_minutes,
        "total_bottling_minutes": total_bottling_minutes,
        "total_filling_minutes": total_filling_minutes,
    }


OUTPUT_FIELD_NOTES: dict[str, dict[str, object]] = {
    "best_schedule_tasks.csv": {
        "table_note": "任務明細總表。用來逐筆檢查每張訂單在什麼日期做了哪一個工序，以及該工序占用了多少機器分鐘與人力分鐘。",
        "fields": [
            ("訂單", "GA 內部排程訂單編號。", "同一訂單的所有工序應可在這張表串成完整流程。", "若不是「總計」，就代表這一列真的有被安排進排程。"),
            ("類別", "工序類別，例如預攪、攪料清洗、排瓶、充填。", "可快速辨認此列屬於製造段或充填段。", "一張正常要充填的訂單，通常會看到製造工序後接排瓶/充填工序。"),
            ("明細", "更細的作業內容，通常包含產品、製程別、批次資訊。", "用來檢查同一批次是否有被拆得太碎，或流程命名是否合理。", "若同一產品同一天出現非常多重複零碎列，就值得回頭檢查排法。"),
            ("日期", "此筆作業實際被安排的日曆日期。", "用來檢查是否晚於交期、是否符合前後工序順序。", "預攪應早於攪料清洗；排瓶原則上要早於或緊鄰充填。"),
            ("機器類型", "此工序占用的主要機器或設備類型。", "可搭配每日製造/充填排程表檢查機器是否過載。", "若同一天同機器分鐘過高，就要再看是否超過可用產能。"),
            ("機器分鐘", "這筆作業占用機器的總分鐘數。", "目前對外統一用整數顯示。", "若數值異常偏小或偏大，應回頭對照效能表與需求量。"),
            ("人力類型", "此工序歸屬的人力池，例如調製、排瓶、充填。", "可對照每日負荷表，理解哪一種人力是瓶頸。", "排瓶與充填必須和調製分開看，不能混為同一組人力。"),
            ("人力分鐘", "這筆作業消耗的人力分鐘。", "排瓶與調製通常看人力分鐘；充填摘要另外採機台分鐘口徑統計超時。", "若工序相同但人力分鐘和機器分鐘比例怪異，需回頭檢查標準配置人數。"),
        ],
    },
    "best_schedule_daily_load.csv": {
        "table_note": "每日負荷摘要表。這份表目前已改成和正式摘要一致的超時口徑，方便直接檢查哪一天、哪一類資源發生加班。",
        "fields": [
            ("日期", "每日負荷的統計日期。", "以天為單位檢查壅塞日。", "可用來找出加班集中在哪些日期。"),
            ("人力類型", "統計對象，分為調製、排瓶、充填。", "調製看調製人力分鐘；排瓶看排瓶作業分鐘基準；充填看充填機台分鐘基準。", "這個欄位是理解超時口徑的關鍵。"),
            ("已用人力分鐘", "該日該類資源實際被排進去的使用量。", "代表當日真正承載的工作量。", "若明顯高於正常人力分鐘，就表示發生加班。"),
            ("正常人力分鐘", "模型設定下該類資源每日不加班時可承載的分鐘數。", "是計算超時的基準。", "排瓶與充填目前以流程瓶頸基準計算，不是單純總人頭乘 8 小時。"),
            ("超時分鐘", "已用量扣掉正常量後的超出部分，不足則為 0。", "這欄加總後應和正式摘要的各類超時數接近或一致。", "若和 summary 對不起來，就表示統計口徑有問題。"),
        ],
    },
    "schedule_summary.csv": {
        "table_note": "正式摘要表。這是最適合報告時直接引用的總量結果。",
        "fields": [
            ("總加班分鐘數", "本次排程所有加班分鐘數的總和。", "通常等於調製超時 + 排瓶超時 + 充填超時。", "若和三者合計差太多，就要檢查摘要邏輯。"),
            ("調製超時分鐘數", "調製端超出正常可承載量的分鐘數。", "反映製造是否壅塞。", "目前若是 0，代表調製端在現有模型下未超過日基準。"),
            ("排瓶超時分鐘數", "排瓶端超出的分鐘數。", "常用來判斷是否是排瓶人力不足造成瓶頸。", "若遠大於充填超時，通常代表瓶頸集中在排瓶。"),
            ("充填超時分鐘數", "充填端超出的分鐘數。", "反映充填機與流程的緊度。", "若很高，表示充填機台天數或產能安排需要再優化。"),
            ("週末加班分鐘數", "落在週末的作業分鐘數。", "是額外風險指標，代表排程對假日的依賴。", "若此數值過高，代表平日容量可能不足。"),
            ("總排瓶分鐘數", "整個排程所有排瓶工序的機器分鐘總和。", "可用來和排瓶超時一起判讀負荷強度。", "總量高但超時低，代表攤得開；總量高且超時高，代表排得很滿。"),
            ("總充填分鐘數", "整個排程所有充填工序的機器分鐘總和。", "可與充填超時一起判斷充填段緊迫程度。", "適合和業主提供的充填 baseline 比較。"),
        ],
    },
    "daily_manufacturing_schedule.csv": {
        "table_note": "每日調製排程表。用來查每一天有哪些製造工序占用哪些機器。",
        "fields": [
            ("日期", "製造作業發生日期。", "可用來追工序先後與機器負荷。", "若製造太集中在少數幾天，通常容易形成清洗或人力壓力。"),
            ("訂單", "該製造作業所屬訂單。", "可與任務明細表互相追蹤。", "應能在任務表找到同一筆作業。"),
            ("類別", "預攪或攪料清洗。", "表示製造流程中的哪一段。", "預攪應先於攪料清洗。"),
            ("明細", "含製程別與批次的詳細名稱。", "用來檢查是否確實按批次安排。", "若批次順序混亂，需要再對照業務規則。"),
            ("機器類型", "當日此筆製造作業所占用的機器。", "用來看同機器是否塞太滿。", "同一台關鍵機器若常滿載，代表它可能是真瓶頸。"),
            ("機器分鐘", "此筆製造工序占用機器多久。", "反映該批次作業時間。", "應與製造效能表的對應批次時間相近。"),
            ("人力分鐘", "此筆製造工序占用調製人力多久。", "目前製造視為顧機，因此人力時間通常與作業時間同尺度。", "若出現極端不對稱，需要再查模型假設。"),
            ("當日機器總分鐘", "同日同機器全部製造作業合計分鐘。", "用來看某台機器當天是否超量。", "適合和機器日容量或可用台數交叉檢查。"),
            ("當日訂單總分鐘", "同日同訂單的製造分鐘合計。", "可用來看某張訂單是否把當天塞得太滿。", "若單日集中過多，容易後續影響前置與充填。"),
            ("當日調製人力總分鐘", "該日整體調製人力被占用的總分鐘。", "可快速看調製是否接近滿載。", "若超過正常基準，就會對應到調製超時。"),
        ],
    },
    "daily_packaging_schedule.csv": {
        "table_note": "每日排瓶與充填排程表。這張表是找充填瓶頸最重要的檢查表。",
        "fields": [
            ("日期", "排瓶或充填發生日期。", "用來檢查前後工序是否合理。", "排瓶通常應在充填前一天或同一極緊迫情況下緊鄰處理。"),
            ("訂單", "當日這筆排瓶/充填作業所屬訂單。", "可追蹤單張訂單的充填流程。", "搭配 reference format 可檢查時段切分是否合理。"),
            ("類別", "GA 中的工序類別，對應排瓶或充填。", "快速識別此列屬於哪個瓶頸段。", "目前對外應以排瓶/充填中文理解。"),
            ("工序", "流程名稱，通常是排瓶或充填。", "是給業主閱讀的流程欄。", "應與類別一致，不應出現互相矛盾。"),
            ("明細", "產品名稱加上工序描述。", "用來辨認是哪一個產品在做排瓶或充填。", "若同一天同訂單同工序出現多列，應確認是否為必要分段。"),
            ("機器類型", "該作業使用的充填設備或排瓶設備。", "用來檢查同機台每天有沒有塞滿。", "若特定機台長期高負荷，可能需優先優化。"),
            ("機器分鐘", "該作業占用設備的分鐘數。", "目前對外顯示為整數。", "可與時段數乘 5 分鐘做交叉檢查。"),
            ("人力分鐘", "該作業消耗的人力分鐘。", "排瓶通常較接近人工負擔；充填則常反映標準配置人數乘工時。", "若人力分鐘異常高，要回頭看標準配置人數。"),
            ("當日工序總分鐘", "同一天同工序的總分鐘。", "顯示這一天排瓶或充填的總負荷。", "是找壅塞日的重要欄位。"),
            ("當日機器總分鐘", "同一天同機器類型的總分鐘。", "可看設備是否集中爆量。", "若超過可用時數，代表非常緊。"),
            ("當日訂單總分鐘", "同一天同一訂單在包裝段的總分鐘。", "用來看單張訂單是否在一天內塞入過多排瓶/充填。", "若非常高，可能表示排程過度集中。"),
            ("當日工序人力總分鐘", "同一天同工序的人力分鐘總和。", "排瓶超時會特別受這欄影響。", "若大於正常分鐘很多，表示加班壓力大。"),
            ("當日正常分鐘", "該工序當日不加班時的正常可承載分鐘。", "是超時判斷基準。", "排瓶與充填的正常分鐘目前都用流程瓶頸基準。"),
            ("當日超時分鐘", "當日工序總量超出正常分鐘的部分。", "是最直接的單日超時訊號。", "把這欄加總後，應能對應到 summary 的排瓶/充填超時。"),
        ],
    },
    "manufacturing_cleaning_audit.csv": {
        "table_note": "調製清洗稽核表。用來檢查同機器同一天是否涉及多種不同製程，進而推估換產品清洗次數。",
        "fields": [
            ("日期", "清洗稽核的日期。", "以天為單位檢查換線/清洗風險。", "連續多天都有換產品時，要特別注意現場可行性。"),
            ("機器類型", "被稽核的調製機器。", "例如預攪機、大乳機、小乳機等。", "關鍵機器出現多次換產品清洗時，代表排程切換頻繁。"),
            ("當日機器總分鐘", "該日該機器被安排的總分鐘數。", "可反映機器負荷強度。", "若負荷很高又換產品多，實務風險會上升。"),
            ("當日訂單數", "該日該機器碰到幾張不同訂單。", "可粗看當天是否切很多單。", "訂單數愈多，通常換產品風險愈高。"),
            ("當日不同製程數", "該日該機器涉及幾種不同製程群組。", "用來估計清洗需求。", "如果大於 1，通常代表有切換。"),
            ("推估換產品清洗次數", "用不同製程數推估的換產品清洗次數。", "目前是稽核指標，不是完整硬限制。", "若這欄偏高，應視為後續優化重點。"),
            ("訂單清單", "當天在該機器上處理過的訂單集合。", "方便追查是哪幾張訂單造成切換。", "可和製程明細一起看切換來源。"),
            ("製程明細", "當日該機器涉及的製程群組。", "可辨識是大乳/小乳/精等哪幾種流程在切換。", "若同一天混太多類型，合理性要特別檢查。"),
        ],
    },
    "best_schedule_reference_format.csv": {
        "table_note": "業主簡報格式對照表。這張表把排程轉成較接近簡報與檢查用途的欄位。",
        "fields": [
            ("訂單編號", "排程中的訂單識別碼。", "用來和其他所有表單串接。", "應可對應到 audit、tasks、anomalies。"),
            ("子任務編號", "同一訂單下的子工序識別碼。", "用來辨識同一筆工序在 reference format 中的列。", "同一批次或同一工序通常會共用相近編號。"),
            ("工序代碼", "內部工序碼，例如 Cat1~Cat4。", "Cat1=預攪、Cat2=攪料清洗、Cat3=排瓶、Cat4=充填。", "若要和業主溝通，建議搭配作業工序欄一起看。"),
            ("工序順序", "模型設定的流程順序代碼。", "數字愈小表示流程愈前面。", "可用來快速檢查前後順序是否顛倒。"),
            ("作業項目", "較大的作業分類，例如調製項、排瓶項、充填項。", "方便業主按大流程閱讀。", "應與工序代碼、作業工序一致。"),
            ("作業工序", "實際流程名稱。", "這是簡報時最容易解讀的欄位。", "若和 tasks 表中的類別對不起來，就要查映射。"),
            ("人力別", "此工序所屬人力類型。", "通常區分調製類型或充填類型。", "有助於理解人力資源是誰在負責。"),
            ("機台別", "此子任務使用的設備。", "可對照每日排程與機器限制。", "同日同機台排滿時要看時段範圍是否合理。"),
            ("標準配置人數", "來源效能或模型基準設定的人數。", "代表這個工序在資料來源中的基準配置。", "若與業主原始資料不一致，要先修資料或映射邏輯。"),
            ("實際安排人數", "本次排程實際採用的人數。", "若為排瓶，可能因模型允許同時最多兩人而大於標準配置人數。", "若標準與實際差很多，要回頭檢查是否真的符合現場規則。"),
            ("總需求量", "此工序整體需求的總單位數。", "製造工序通常表示批次需求，充填工序通常表示 5 分鐘 slot 需求量。", "適合和當日安排量一起看完成比例。"),
            ("當日安排量", "這一天實際安排進去的單位數。", "是將總需求切到單日後的安排量。", "若總需求很大但安排量很小，表示排程攤得很開。"),
            ("每單位時間(min)", "單位作業時間。", "製造通常是一整批的時間；排瓶/充填目前以 5 分鐘 slot 顯示。", "可用來理解總時間是怎麼累積出來的。"),
            ("排程日期", "此子任務被安排的日期。", "搭配時段範圍可看到更細的日內位置。", "應滿足前後工序順序與交期限制。"),
            ("時段範圍", "把一天切成 5 分鐘後，此任務在日內占用的起訖時段。", "主要用來做簡報視覺化與日內檢查。", "若同機台時段明顯重疊，就要再查 slot 指派邏輯。"),
            ("時段數", "此任務總共占用幾個 5 分鐘時段。", "乘以 5 分鐘後應接近作業總時間。", "可用來檢查時段與分鐘是否一致。"),
            ("排程類型", "目前輸出多為正常。", "保留給後續標記急單、例外處理等用途。", "若未來出現特殊類型，要另外定義業務含義。"),
            ("作業總時間(min)", "此列實際占用的總分鐘數。", "是 reference format 內最直接的時間量。", "應與 tasks 表的機器分鐘一致或接近。"),
            ("需求單位數", "再次表達此工序總需求單位。", "主要是為了讓簡報表可單獨閱讀。", "應和總需求量口徑一致。"),
            ("最早時段索引", "此任務在該機台當日排到的最早 slot 編號。", "0 代表 09:00 開始的第一個 5 分鐘時段。", "可用來判斷作業在一天中的相對前後位置。"),
            ("最晚時段索引", "此任務在該機台當日排到的最晚 slot 編號。", "用來界定任務在日內占用到哪個時間點。", "若最晚索引過高，表示該天已排到很晚。"),
        ],
    },
    "order_anomalies.csv": {
        "table_note": "異常訂單總表。列出所有被偵測到的 exclude 或 warning 訂單與原因。",
        "fields": [
            ("order_id", "訂單編號。", "對應排程中的 order_id。", "可回頭查是否有出現在主排程內。"),
            ("product_name", "產品名稱。", "方便業務與現場辨識。", "若產品名稱與訂單對不起來，要先檢查資料映射。"),
            ("severity", "異常嚴重程度，常見為 exclude 或 warning。", "exclude 代表不進 GA 主優化集合；warning 代表仍可排但要注意。", "是判讀是否漏排的第一關鍵。"),
            ("anomaly_type", "異常類型代碼。", "方便快速分類問題來源。", "若同一訂單有多種異常，會有多列。"),
            ("exclude_from_optimization", "是否從主優化集合排除。", "Y 代表不會進正式排程；N 代表只是提醒。", "所有沒排進去的訂單都應先在這欄找到 Y。"),
            ("message", "異常的文字說明。", "直接描述為什麼不合理或風險高。", "是最適合向老師或業主說明的欄位。"),
            ("due_date", "交期。", "用來和 ready date 比較。", "若 ready 在 due 後，通常就是硬性無解。"),
            ("raw_ready_date", "原料齊備可開始製造的日期。", "製造最早可動工的依據之一。", "若晚於交期，代表製造根本來不及。"),
            ("material_ready_date", "物料/包材齊備可進包裝的日期。", "排瓶與充填是否可開始的重要日期。", "若晚於交期，通常直接無法排。"),
            ("manufacturing_batch_count", "此訂單需要的總批次數。", "可看製造壓力。", "批次很多但窗期很短時，通常會被標 warning。"),
            ("filling_minutes", "估算的充填分鐘。", "用來理解此單在充填段的負荷。", "分鐘太大時常是瓶頸來源。"),
            ("bottling_minutes", "估算的排瓶分鐘。", "用來理解此單在排瓶段的負荷。", "若排瓶分鐘大，排瓶加班風險也會較高。"),
        ],
    },
    "order_anomalies_exclude.csv": {"table_note": "只保留 exclude 異常的清單。", "fields": []},
    "order_anomalies_warning.csv": {"table_note": "只保留 warning 異常的清單。", "fields": []},
    "schedule_validation.csv": {
        "table_note": "排程驗證明細表。只要這張表有內容，就表示至少有一個核心規則被違反。",
        "fields": [
            ("order_id", "被驗證出問題的訂單。", "用來定位哪一張單違規。", "若正式版要交付，理想上這張表應該是空的。"),
            ("severity", "問題嚴重程度。", "常見為 error 或 warning。", "正式版最少要做到 error = 0。"),
            ("issue_type", "違規類型代碼。", "例如無排程、先後順序錯誤、缺工序等。", "可快速聚焦是哪種邏輯出問題。"),
            ("message", "驗證訊息。", "更完整描述規則違反內容。", "最適合拿來 debug。"),
        ],
    },
    "schedule_validation_summary.csv": {
        "table_note": "排程驗證摘要表。比起 validation 明細，這張更適合快速確認正式結果是否可交付。",
        "fields": [
            ("metric", "驗證指標名稱。", "例如 total_issues、error_count 等。", "用來快速看排程是否仍有硬性違規。"),
            ("value", "該指標的數值。", "代表 validation 的統計結果。", "正式版通常至少要確保 total_issues 與 error_count 為 0。"),
        ],
    },
    "order_source_links.csv": {
        "table_note": "來源資料對照表。用來看每張 P 訂單是怎麼去對到原始資料來源的。",
        "fields": [
            ("p_order_id", "GA 內部使用的 P 訂單編號。", "是目前模型主鍵之一。", "應可和 tasks、audit 表串接。"),
            ("product_name", "產品名稱。", "幫助人工確認映射合理性。", "若名稱相差很大，可能配對錯誤。"),
            ("time_ready_core_candidates", "由 ready 時間與核心碼推估出的候選資訊。", "反映模型初步比對線索。", "可用來回查為何會配到某張來源單。"),
            ("due_serial", "從訂單編號或資料中抽出的交期序列資訊。", "作為比對輔助。", "若和實際交期差很多，需再清洗資料。"),
            ("quantity", "訂單需求數量。", "常作為來源匹配分數的一部分。", "若數量落差很大，映射可信度通常偏低。"),
            ("best_match_score", "目前最佳配對分數。", "分數愈高通常代表匹配愈可信。", "很低時應人工複核。"),
            ("match_count", "候選配對筆數。", "代表這張單有多少可能來源。", "太多時表示資料歧義高。"),
            ("bom_match_count", "在 BOM 層成功對到的候選數。", "反映 BOM 映射完整度。", "若為 0，通常代表模型是用 fallback。"),
            ("manufacturing_efficiency_match_count", "在製造效能表成功對到的候選數。", "用來看製造工時資料是否充足。", "若為 0，製造工時可能不夠可靠。"),
            ("filling_efficiency_match_count", "在充填效能表成功對到的候選數。", "用來看排瓶/充填工時資料是否充足。", "若為 0，包裝段估算可信度較低。"),
            ("source_order_ids", "實際匹配到的來源訂單編號集合。", "方便人工追資料。", "可直接拿去原始檔搜尋。"),
            ("source_product_codes", "匹配到的來源產品代碼集合。", "幫助核對核心碼。", "若代碼分散太多，表示映射不穩。"),
            ("source_core_codes", "匹配到的來源核心碼集合。", "是產品辨識的重要依據。", "核心碼若與預期不符，應視為高風險。"),
        ],
    },
    "machine_counts.csv": {
        "table_note": "模型使用的機器台數表。用來確認程式實際採用了哪些機器台數設定。",
        "fields": [
            ("machine_type", "機器類型名稱。", "對應排程內的設備名稱。", "應和日排程中的機器名稱一致。"),
            ("count", "模型假設的可用台數。", "直接影響該設備每日總容量。", "若業主提供台數不同，這張表一定要先對。"),
        ],
    },
    "order_model_audit.csv": {
        "table_note": "建模稽核表。用來看每張訂單被模型轉成了哪些參數，是 debug 與報告的重要依據。",
        "fields": [
            ("order_id", "訂單編號。", "模型主鍵。", "應和其他所有主表一致。"),
            ("product_name", "產品名稱。", "方便人工辨識。", "若與來源資料不一致，要先處理資料清洗。"),
            ("source_order_id", "對到的原始來源訂單編號。", "追溯資料來源用。", "找不到時通常代表用 fallback 推估。"),
            ("source_product_code", "對到的原始產品代碼。", "用來追核心碼與 ml。", "可拿去比對 BOM 與效能表。"),
            ("source_core_code", "對到的核心碼。", "是跨表映射的重要識別欄。", "錯了通常整張單的建模都會歪掉。"),
            ("due_date", "交期。", "排程終點限制。", "必須晚於主要工序安排日。"),
            ("raw_ready_date", "原料 ready 日期。", "製造最早起點之一。", "若晚於交期則為硬性無解。"),
            ("material_ready_date", "物料 ready 日期。", "排瓶與充填最早起點。", "若太晚會直接擠壓充填窗。"),
            ("quantity", "需求量。", "決定製造/充填總量基礎。", "若和原始資料差異大，後續所有時間都會失真。"),
            ("manufacturing_item_count", "該訂單被拆成幾個製造子流程。", "雙層劑型等通常會大於 1。", "若明顯不合理，表示 BOM 拆解有問題。"),
            ("manufacturing_processes", "模型識別出的製造流程名稱集合。", "用來看這單究竟走了哪些製程。", "可直接核對業主對產品製程的理解。"),
            ("manufacturing_machines", "這些製造流程對應到的機器集合。", "反映機器需求。", "若關鍵機器抓錯，排程容量會整個偏掉。"),
            ("manufacturing_batch_total", "總製造批次數。", "用來看製造負荷。", "不應該出現 1.5 這種小數批次。"),
            ("manufacturing_premix_total_minutes", "所有預攪加總分鐘數。", "反映製造前段耗時。", "通常應與批次數與效能表一致。"),
            ("manufacturing_mix_clean_total_minutes", "所有攪料清洗加總分鐘數。", "反映製造主體負荷。", "是檢查調製端是否合理的重要欄。"),
            ("filling_machine", "模型採用的充填設備。", "影響充填容量。", "應與產品實務相符。"),
            ("bottling_machine", "模型採用的排瓶設備。", "影響排瓶容量。", "若這欄錯，排瓶時間也會錯。"),
            ("filling_minutes", "估算的充填總分鐘。", "是充填段工作量。", "和實際 tasks 表加總應大致一致。"),
            ("bottling_minutes", "估算的排瓶總分鐘。", "是排瓶段工作量。", "和實際 tasks 表加總應大致一致。"),
            ("filling_people", "模型採用的充填配置人數。", "會影響充填人力分鐘。", "若和現場認知不同，需調整假設。"),
            ("bottling_people", "模型採用的排瓶配置人數。", "會影響排瓶人力分鐘。", "理論上不應出現小數。"),
            ("source_match_count", "來源匹配候選數。", "反映映射確定性。", "太多時表示需人工確認。"),
            ("source_best_match_score", "最佳映射分數。", "數值愈高通常可信度愈高。", "低分訂單應列為後續檢查重點。"),
            ("bom_match_count", "BOM 成功對應次數。", "看 BOM 映射完整度。", "若太低，製程拆解可能不穩。"),
            ("manufacturing_efficiency_match_count", "製造效能對應次數。", "看製造工時資料完整度。", "太低時調製分鐘可信度不足。"),
            ("filling_efficiency_match_count", "充填效能對應次數。", "看包裝工時資料完整度。", "太低時排瓶/充填分鐘可信度不足。"),
        ],
    },
}

OUTPUT_FIELD_NOTES["order_anomalies_exclude.csv"]["fields"] = OUTPUT_FIELD_NOTES["order_anomalies.csv"]["fields"]
OUTPUT_FIELD_NOTES["order_anomalies_warning.csv"]["fields"] = OUTPUT_FIELD_NOTES["order_anomalies.csv"]["fields"]


def write_output_field_notes(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    note_dir = output_dir / "欄位備註表"
    note_dir.mkdir(parents=True, exist_ok=True)

    master_path = note_dir / "欄位備註總表.csv"
    with master_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["表單檔名", "表單用途", "欄位名稱", "欄位說明", "如何判讀", "合理性檢查重點"])
        for filename, config in OUTPUT_FIELD_NOTES.items():
            table_note = str(config.get("table_note", ""))
            fields = list(config.get("fields", []))
            for field_name, meaning, reading, check in fields:
                writer.writerow([filename, table_note, field_name, meaning, reading, check])

    for filename, config in OUTPUT_FIELD_NOTES.items():
        safe_stem = Path(filename).stem
        note_path = note_dir / f"{safe_stem}_欄位備註表.csv"
        with note_path.open("w", newline="", encoding="utf-8-sig") as file:
            writer = csv.writer(file)
            writer.writerow(["表單檔名", "表單用途", "欄位名稱", "欄位說明", "如何判讀", "合理性檢查重點"])
            table_note = str(config.get("table_note", ""))
            fields = list(config.get("fields", []))
            for field_name, meaning, reading, check in fields:
                writer.writerow([filename, table_note, field_name, meaning, reading, check])


def allows_same_day_bottling(order: Order, fill_day: date) -> bool:
    return (
        order.filling_item.bottling_machine_minutes > 0
        and fill_day >= order.material_ready_date
        and order.due_date <= order.material_ready_date + timedelta(days=1)
    )


def validate_schedule(
    schedule: ScheduleResult,
    schedulable_orders: list[Order],
) -> list[ScheduleValidationIssue]:
    issues: list[ScheduleValidationIssue] = []
    order_map = {order.order_id: order for order in schedulable_orders}
    tasks_by_order: dict[str, list[ScheduledTask]] = defaultdict(list)
    for task in schedule.task_records:
        tasks_by_order[task.order_id].append(task)

    for order in schedulable_orders:
        tasks = tasks_by_order.get(order.order_id, [])
        if not tasks:
            issues.append(
                ScheduleValidationIssue(
                    order_id=order.order_id,
                    issue_type="order_not_scheduled",
                    severity="error",
                    message="Schedulable order has no scheduled tasks.",
                )
            )
            continue

        meaningful_tasks = [task for task in tasks if task.machine_minutes > 1e-2 or task.labor_minutes > 1e-2]
        cat1_days = sorted(task.scheduled_date for task in meaningful_tasks if task.task_type == "Cat1")
        cat2_days = sorted(task.scheduled_date for task in meaningful_tasks if task.task_type == "Cat2")
        cat3_days = sorted(task.scheduled_date for task in meaningful_tasks if task.task_type == "Cat3")
        cat4_days = sorted(task.scheduled_date for task in meaningful_tasks if task.task_type == "Cat4")

        if cat4_days and not cat2_days:
            issues.append(
                ScheduleValidationIssue(
                    order_id=order.order_id,
                    issue_type="fill_without_manufacturing",
                    severity="error",
                    message="Order has filling tasks but no manufacturing tasks.",
                )
            )

        actual_fill_minutes = sum(task.machine_minutes for task in tasks if task.task_type == "Cat4")
        if abs(actual_fill_minutes - order.filling_item.filling_machine_minutes) > 1e-2:
            issues.append(
                ScheduleValidationIssue(
                    order_id=order.order_id,
                    issue_type="filling_minutes_mismatch",
                    severity="error",
                    message=(
                        f"Filled {actual_fill_minutes:.2f} min, expected "
                        f"{order.filling_item.filling_machine_minutes:.2f} min."
                    ),
                )
            )

        actual_bottling_work = sum(task.labor_minutes for task in tasks if task.task_type == "Cat3")
        expected_bottling_minutes = order.filling_item.bottling_machine_minutes
        expected_bottling_work = expected_bottling_minutes * order.filling_item.bottling_people
        if abs(actual_bottling_work - expected_bottling_work) > 1e-2:
            issues.append(
                ScheduleValidationIssue(
                    order_id=order.order_id,
                    issue_type="bottling_minutes_mismatch",
                    severity="error",
                    message=(
                        f"Bottling work {actual_bottling_work:.2f} min, expected "
                        f"{expected_bottling_work:.2f} min."
                    ),
                )
            )

        cat1_day_set = set(cat1_days)
        for mix_day in cat2_days:
            if mix_day - timedelta(days=1) not in cat1_day_set:
                issues.append(
                    ScheduleValidationIssue(
                        order_id=order.order_id,
                        issue_type="premix_not_previous_day",
                        severity="error",
                        message=f"Mix day {mix_day.isoformat()} has no premix on the previous day.",
                    )
                )
                break

        if expected_bottling_minutes > 1e-2:
            cat3_day_set = set(cat3_days)
            for fill_day in cat4_days:
                has_previous_day_bottling = fill_day - timedelta(days=1) in cat3_day_set
                has_same_day_bottling = fill_day in cat3_day_set and allows_same_day_bottling(order, fill_day)
                if not has_previous_day_bottling and not has_same_day_bottling:
                    issues.append(
                        ScheduleValidationIssue(
                            order_id=order.order_id,
                            issue_type="bottling_not_previous_day",
                            severity="error",
                            message=f"Fill day {fill_day.isoformat()} has no bottling on the previous day.",
                        )
                    )
                    break

        for task in tasks:
            if task.task_type in {"Cat1", "Cat2"} and task.scheduled_date < order.raw_ready_date:
                issues.append(
                    ScheduleValidationIssue(
                        order_id=order.order_id,
                        issue_type="manufacturing_before_raw_ready",
                        severity="error",
                        message=(
                            f"{task.task_type} on {task.scheduled_date.isoformat()} is earlier than "
                            f"raw_ready_date {order.raw_ready_date.isoformat()}."
                        ),
                    )
                )
                break
            if task.task_type in {"Cat3", "Cat4"} and task.scheduled_date < order.material_ready_date:
                issues.append(
                    ScheduleValidationIssue(
                        order_id=order.order_id,
                        issue_type="packaging_before_material_ready",
                        severity="error",
                        message=(
                            f"{task.task_type} on {task.scheduled_date.isoformat()} is earlier than "
                            f"material_ready_date {order.material_ready_date.isoformat()}."
                        ),
                    )
                )
                break
            if task.scheduled_date > order.due_date:
                issues.append(
                    ScheduleValidationIssue(
                        order_id=order.order_id,
                        issue_type="task_after_due_date",
                        severity="error",
                        message=f"{task.task_type} is scheduled after due_date {order.due_date.isoformat()}.",
                    )
                )
                break

    return issues


class SimpleXlsxReader:
    """用標準函式庫讀取 xlsx。"""

    def __init__(self, workbook_path: Path) -> None:
        self.workbook_path = workbook_path

    def list_sheet_names(self) -> list[str]:
        with zipfile.ZipFile(self.workbook_path) as zf:
            workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        return [sheet.attrib["name"] for sheet in workbook.findall("main:sheets/main:sheet", XML_NS)]

    def read_sheet(self, sheet_name: str) -> list[dict[str, str]]:
        with zipfile.ZipFile(self.workbook_path) as zf:
            shared_strings = self._read_shared_strings(zf)
            target = self._resolve_sheet_target(zf, sheet_name)
            xml_root = ET.fromstring(zf.read(target))

        rows: list[list[str]] = []
        for row in xml_root.findall("main:sheetData/main:row", XML_NS):
            values: dict[int, str] = {}
            for cell in row.findall("main:c", XML_NS):
                values[self._column_index(cell.attrib.get("r", "A1"))] = self._cell_value(
                    cell=cell,
                    shared_strings=shared_strings,
                )
            if values:
                max_index = max(values)
                rows.append([values.get(index, "") for index in range(max_index + 1)])
            else:
                rows.append([])

        if not rows:
            return []

        header = [str(item).strip() for item in rows[0]]
        records: list[dict[str, str]] = []
        for raw_row in rows[1:]:
            if not any(str(item).strip() for item in raw_row):
                continue
            record: dict[str, str] = {}
            for index, column_name in enumerate(header):
                if not column_name:
                    continue
                value = raw_row[index] if index < len(raw_row) else ""
                record[column_name] = "" if value is None else str(value).strip()
            records.append(record)
        return records

    def _read_shared_strings(self, zf: zipfile.ZipFile) -> list[str]:
        if "xl/sharedStrings.xml" not in zf.namelist():
            return []
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
        return [
            "".join(node.text or "" for node in item.iterfind(".//main:t", XML_NS))
            for item in root.findall("main:si", XML_NS)
        ]

    def _resolve_sheet_target(self, zf: zipfile.ZipFile, sheet_name: str) -> str:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        relationships = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        relation_map = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in relationships.findall("rel:Relationship", XML_NS)
        }
        requested_name = self._normalize_sheet_name(sheet_name)
        for sheet in workbook.findall("main:sheets/main:sheet", XML_NS):
            actual_name = sheet.attrib["name"]
            if actual_name == sheet_name or self._normalize_sheet_name(actual_name) == requested_name:
                relation_id = sheet.attrib[
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                ]
                return "xl/" + relation_map[relation_id]
        raise KeyError(f"找不到工作表：{sheet_name}")

    @staticmethod
    def _normalize_sheet_name(sheet_name: str) -> str:
        return re.sub(r"\s+", "", sheet_name).casefold()

    @staticmethod
    def _column_index(cell_ref: str) -> int:
        match = re.match(r"[A-Z]+", cell_ref)
        if not match:
            return 0
        index = 0
        for character in match.group(0):
            index = index * 26 + ord(character) - 64
        return index - 1

    @staticmethod
    def _cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
        cell_type = cell.attrib.get("t")
        if cell_type == "s":
            value_node = cell.find("main:v", XML_NS)
            if value_node is None:
                return ""
            return shared_strings[int(value_node.text)]
        if cell_type == "inlineStr":
            return "".join(node.text or "" for node in cell.iterfind(".//main:t", XML_NS))
        value_node = cell.find("main:v", XML_NS)
        return value_node.text if value_node is not None else ""


class WorkbookLoader:
    """把 Excel 的時間準備表整理成 GA 可用的訂單資料。"""

    def __init__(self, workbook_path: Path) -> None:
        self.reader = SimpleXlsxReader(workbook_path)
        self.sheet_names = self.reader.list_sheet_names()
        self.source_tables = self._load_source_tables()
        self.machine_counts = self._build_machine_counts()
        self.source_indexes = self._build_source_indexes()

    def _load_source_tables(self) -> dict[str, list[dict[str, str]]]:
        source_sheet_map = {
            "time_ready": ("時間準備表V1",),
            "orders_2023": ("X2023\u8a02\u55ae\u8cc7\u8a0a\u542b\u5165\u5eab\u65e5", "2023\u8a02\u55ae\u8cc7\u8a0a"),
            "bom": ("BOM",),
            "manufacturing_efficiency": ("\u88fd\u9020\u6548\u80fd_BOM\u56de\u63a8", "\u88fd\u9020\u6548\u80fd"),
            "filling_efficiency": ("\u5145\u586b\u6548\u80fd_BOM\u56de\u63a8", "\u5145\u586b\u6548\u80fd"),
            "machine_limits": ("機器限制",),
            "overtime_calc": ("加班計算",),
            "field_guide": ("欄位說明",),
        }
        tables: dict[str, list[dict[str, str]]] = {}
        for key, candidates in source_sheet_map.items():
            tables[key] = self._read_first_available_sheet(*candidates)
        return tables

    def _read_first_available_sheet(self, *sheet_names: str) -> list[dict[str, str]]:
        normalized_available = {
            SimpleXlsxReader._normalize_sheet_name(name): name for name in self.sheet_names
        }
        for sheet_name in sheet_names:
            actual_name = normalized_available.get(SimpleXlsxReader._normalize_sheet_name(sheet_name))
            if actual_name is not None:
                return self.reader.read_sheet(actual_name)
        return []

    def _build_machine_counts(self) -> dict[str, int]:
        counts = dict(MACHINE_COUNTS)
        for row in self.source_tables["machine_limits"]:
            for value in row.values():
                if not value:
                    continue
                for machine_name, machine_count in self._extract_machine_counts(value).items():
                    counts[machine_name] = max(counts.get(machine_name, 0), machine_count)

        # Map source-machine names into the simplified machine names currently used by the scheduler.
        counts["精機"] = max(counts.get("精機", 1), counts.get("大型攪拌機", 1))
        counts["小乳機"] = max(counts.get("小乳機", 1), counts.get("50L真空乳化機", 1))
        counts["大乳機"] = max(counts.get("大乳機", 1), counts.get("100L真空乳化機", 1))
        counts["大機"] = max(
            counts.get("大機", 1),
            counts.get("400L大桶槽", 0) + counts.get("500L大桶槽", 0),
        )
        counts["充填機"] = max(
            counts.get("充填機", 1),
            counts.get("泵浦式充填機", 0)
            + counts.get("定量缸式充填機", 0)
            + counts.get("大定量缸式充填機", 0)
            + counts.get("面膜機", 0),
        )
        counts["排瓶機"] = max(counts.get("排瓶機", 1), 1)
        return counts

    def _build_source_indexes(self) -> dict[str, dict[str, list[dict[str, str]]]]:
        return {
            "bom": self._index_rows_by_code_candidates(
                self.source_tables["bom"],
                ("核心碼SN",),
            ),
            "manufacturing_efficiency": self._index_rows_by_code_candidates(
                self.source_tables["manufacturing_efficiency"],
                ("核心碼", "料號"),
            ),
            "filling_efficiency": self._index_rows_by_code_candidates(
                self.source_tables["filling_efficiency"],
                ("核心碼", "料號"),
            ),
        }

    def _index_rows_by_code_candidates(
        self,
        rows: list[dict[str, str]],
        code_fields: tuple[str, ...],
    ) -> dict[str, list[dict[str, str]]]:
        index: dict[str, list[dict[str, str]]] = defaultdict(list)
        for row in rows:
            row_candidates: list[str] = []
            for field in code_fields:
                for candidate in self._core_code_candidates(self._get_value(row, field)):
                    if candidate and candidate not in row_candidates:
                        row_candidates.append(candidate)
            for candidate in row_candidates:
                index[candidate].append(row)
        return index

    @classmethod
    def _extract_machine_counts(cls, text: str) -> dict[str, int]:
        normalized = str(text).replace("*", "x").replace("X", "x")
        found: dict[str, int] = {}
        for match in re.finditer(r"([^\[\]\n、,，]+?)x(\d+)", normalized):
            machine_name = match.group(1).strip(" -")
            if not machine_name:
                continue
            found[machine_name] = max(found.get(machine_name, 0), int(match.group(2)))
        return found

    @staticmethod
    def _get_value(row: dict[str, str], *column_names: str) -> str:
        normalized_row = {
            SimpleXlsxReader._normalize_sheet_name(key): value for key, value in row.items()
        }
        for column_name in column_names:
            direct = row.get(column_name)
            if direct not in (None, ""):
                return direct
            normalized = normalized_row.get(SimpleXlsxReader._normalize_sheet_name(column_name))
            if normalized not in (None, ""):
                return normalized
        return ""

    @staticmethod
    def _normalize_code(code: str) -> str:
        return re.sub(r"[^A-Za-z0-9]", "", code or "").upper()

    @staticmethod
    def _normalize_name(name: str) -> str:
        return re.sub(r"[\s\u200b\ufeff]+", "", name or "").casefold()

    @classmethod
    def _core_code_candidates(cls, code: str) -> tuple[str, ...]:
        normalized = cls._normalize_code(code)
        candidates: list[str] = []

        def add(candidate: str) -> None:
            if candidate and candidate not in candidates:
                candidates.append(candidate)

        add(normalized)
        if len(normalized) >= 6:
            add(normalized[:6])
        if len(normalized) >= 7:
            add(normalized[1:7])
            add(normalized[:7])
        if len(normalized) >= 8:
            add(normalized[1:8])
        if normalized and normalized[-1:] in {"P", "T", "N"}:
            trimmed = normalized[:-1]
            add(trimmed)
            if len(trimmed) >= 6:
                add(trimmed[:6])
            if len(trimmed) >= 7:
                add(trimmed[1:7])

        return tuple(candidates)

    @classmethod
    def _order_row_code_candidates(cls, row: dict[str, str]) -> tuple[str, ...]:
        product_code = cls._get_value(row, "產品代號", "核心六碼", "核心碼SN")
        core_six = cls._get_value(row, "核心六碼")
        candidates: list[str] = []
        for source in (product_code, core_six):
            for candidate in cls._core_code_candidates(source):
                if candidate not in candidates:
                    candidates.append(candidate)
        return tuple(candidates)

    @classmethod
    def _time_ready_row_code_candidates(cls, row: dict[str, str]) -> tuple[str, ...]:
        source_code = cls._get_value(row, "核心碼SN", "核心六碼", "產品代號")
        return cls._core_code_candidates(source_code)

    @classmethod
    def _score_source_match(
        cls,
        time_ready_row: dict[str, str],
        order_row: dict[str, str],
    ) -> int:
        score = 0

        time_codes = set(cls._time_ready_row_code_candidates(time_ready_row))
        order_codes = set(cls._order_row_code_candidates(order_row))
        if time_codes.intersection(order_codes):
            score += 5

        time_due = cls._get_value(time_ready_row, "到客戶倉庫日")
        order_due = cls._get_value(order_row, "到客戶倉庫日")
        if time_due and order_due and int(float(time_due)) == int(float(order_due)):
            score += 3

        time_qty = cls._get_value(time_ready_row, "訂單需求的數量(pcs)", "訂單數量")
        order_qty = cls._get_value(order_row, "訂單數量", "訂單需求的數量(pcs)")
        if time_qty and order_qty and int(float(time_qty)) == int(float(order_qty)):
            score += 2

        time_name = cls._normalize_name(cls._get_value(time_ready_row, "中文名稱", "產品名稱"))
        order_name = cls._normalize_name(cls._get_value(order_row, "產品名稱", "中文名稱"))
        if time_name and order_name and time_name == order_name:
            score += 2

        return score

    def _best_source_order_row(self, time_ready_row: dict[str, str]) -> dict[str, str] | None:
        best_row: dict[str, str] | None = None
        best_score = 0
        for order_row in self.source_tables["orders_2023"]:
            score = self._score_source_match(time_ready_row, order_row)
            if score > best_score:
                best_score = score
                best_row = order_row
        if best_score < 5:
            return None
        return best_row

    def _matched_source_rows(
        self,
        candidates: tuple[str, ...],
        index_name: str,
    ) -> list[dict[str, str]]:
        matched: list[dict[str, str]] = []
        seen_ids: set[int] = set()
        source_index = self.source_indexes[index_name]
        for candidate in candidates:
            for row in source_index.get(candidate, []):
                row_id = id(row)
                if row_id in seen_ids:
                    continue
                seen_ids.add(row_id)
                matched.append(row)
        return matched

    def load_orders(self) -> list[Order]:
        records = self.source_tables["time_ready"]
        grouped_records: dict[str, list[dict[str, str]]] = defaultdict(list)
        for record in records:
            order_id = self._get_value(record, "訂單編號")
            if order_id:
                grouped_records[order_id].append(record)

        orders: list[Order] = []
        for order_id, rows in grouped_records.items():
            first = rows[0]
            source_row = self._best_source_order_row(first)
            manufacturing_items, manufacturing_source_override = self._build_manufacturing_items(
                rows,
                source_row,
            )
            filling_item = self._build_filling_item(rows, source_row)
            if not manufacturing_items or filling_item is None:
                continue

            orders.append(
                Order(
                    order_id=order_id,
                    product_name=self._get_value(first, "中文名稱", "產品名稱"),
                    quantity=self._to_int(self._get_value(first, "訂單需求的數量(pcs)", "訂單數量")),
                    due_date=self._excel_date(self._get_value(first, "到客戶倉庫日")),
                    raw_ready_date=self._excel_date(self._get_value(first, "原料齊料點")),
                    material_ready_date=self._excel_date(self._get_value(first, "物料齊料點")),
                    manufacturing_items=tuple(manufacturing_items),
                    filling_item=filling_item,
                    source_order_id=self._get_value(source_row or {}, "訂單編號"),
                    source_product_code=self._get_value(source_row or {}, "產品代號"),
                    source_core_code=self._get_value(source_row or {}, "核心六碼"),
                )
            )
        return sorted(orders, key=lambda item: (item.due_date, item.order_id))

    def build_order_source_links(self) -> list[dict[str, str]]:
        time_ready_records = self.source_tables["time_ready"]
        order_rows = self.source_tables["orders_2023"]

        linked_rows: list[dict[str, str]] = []
        seen_p_orders: set[str] = set()
        for row in time_ready_records:
            p_order_id = self._get_value(row, "訂單編號")
            if not p_order_id or p_order_id in seen_p_orders:
                continue
            seen_p_orders.add(p_order_id)

            product_name = self._get_value(row, "中文名稱", "產品名稱").strip()
            due_value = self._get_value(row, "到客戶倉庫日")
            qty_value = self._get_value(row, "訂單需求的數量(pcs)", "訂單數量")
            due_serial = int(float(due_value)) if due_value else None
            quantity = int(float(qty_value)) if qty_value else None
            time_ready_candidates = self._time_ready_row_code_candidates(row)
            bom_matches = self._matched_source_rows(time_ready_candidates, "bom")
            manufacturing_efficiency_matches = self._matched_source_rows(
                time_ready_candidates,
                "manufacturing_efficiency",
            )
            filling_efficiency_matches = self._matched_source_rows(
                time_ready_candidates,
                "filling_efficiency",
            )

            scored_matches: list[tuple[int, dict[str, str]]] = []
            for order_row in order_rows:
                score = self._score_source_match(row, order_row)
                if score > 0:
                    scored_matches.append((score, order_row))

            scored_matches.sort(
                key=lambda item: (
                    -item[0],
                    self._get_value(item[1], "訂單編號"),
                )
            )
            best_score = scored_matches[0][0] if scored_matches else 0
            matches = [order_row for score, order_row in scored_matches if score == best_score and score >= 5]
            linked_rows.append(
                {
                    "p_order_id": p_order_id,
                    "product_name": product_name,
                    "time_ready_core_candidates": " | ".join(time_ready_candidates),
                    "due_serial": "" if due_serial is None else str(due_serial),
                    "quantity": "" if quantity is None else str(quantity),
                    "best_match_score": str(best_score),
                    "match_count": str(len(matches)),
                    "bom_match_count": str(len(bom_matches)),
                    "manufacturing_efficiency_match_count": str(len(manufacturing_efficiency_matches)),
                    "filling_efficiency_match_count": str(len(filling_efficiency_matches)),
                    "source_order_ids": " | ".join(self._get_value(match, "訂單編號") for match in matches),
                    "source_product_codes": " | ".join(self._get_value(match, "產品代號") for match in matches),
                    "source_core_codes": " | ".join(self._get_value(match, "核心六碼") for match in matches),
                }
            )
        return linked_rows

    @staticmethod
    def _split_process_names(value: str) -> list[str]:
        return [item.strip() for item in re.split(r"[,、/&]+", value or "") if item.strip()]

    @staticmethod
    def _simplify_manufacturing_machine_name(machine_name: str) -> str:
        normalized = str(machine_name or "")
        if "50L真空乳化機" in normalized or "小乳機" in normalized:
            return "å°ä¹³æ©Ÿ"
        if "100L真空乳化機" in normalized or "大乳機" in normalized:
            return "å¤§ä¹³æ©Ÿ"
        if "大型攪拌機" in normalized or "精機" in normalized:
            return "ç²¾æ©Ÿ"
        if "400L大桶槽" in normalized or "500L大桶槽" in normalized or "大機" in normalized:
            return "å¤§æ©Ÿ"
        return machine_name

    def _manufacturing_machine_from_source(self, source_row: dict[str, str]) -> str:
        process_equipment = self._get_value(source_row, "\u88fd\u7a0b\u8a2d\u5099")
        if process_equipment in {"\u5c0f\u4e73", "\u5c0f\u4e73\u6a5f"}:
            return "\u5c0f\u4e73\u6a5f"
        if process_equipment in {"\u5927\u4e73", "\u5927\u4e73\u6a5f"}:
            return "\u5927\u4e73\u6a5f"
        if process_equipment in {"\u7cbe", "\u7cbe\u6a5f"}:
            return "\u7cbe\u6a5f"
        if process_equipment in {"\u5927", "\u5927\u6a5f"}:
            return "\u5927\u6a5f"

        machine_columns = (
            ("50L真空乳化機", "å°ä¹³æ©Ÿ"),
            ("100L真空乳化機", "å¤§ä¹³æ©Ÿ"),
            ("大型攪拌機", "ç²¾æ©Ÿ"),
            ("400L大桶槽", "å¤§æ©Ÿ"),
            ("500L大桶槽", "å¤§æ©Ÿ"),
        )
        for column_name, machine_name in machine_columns:
            if self._to_int(self._get_value(source_row, column_name)) > 0:
                return machine_name
        return ""

    def _manufacturing_machine_from_efficiency_rows(
        self,
        rows: list[dict[str, str]],
        process_name: str,
    ) -> str:
        preferred_rows = [
            row
            for row in rows
            if self._get_value(row, "工作") == "攪料"
            and (not process_name or self._get_value(row, "攪料程序") == process_name)
        ]
        for row in preferred_rows or rows:
            machine_name = self._get_value(row, "機器")
            bracket_match = re.search(r"\[([^\]]+)\]", machine_name)
            if bracket_match:
                machine_name = bracket_match.group(1)
            machine_name = machine_name.split("&", 1)[0]
            machine_name = re.sub(r"x\d+", "", machine_name).strip(" -")
            simplified = self._simplify_manufacturing_machine_name(machine_name)
            if simplified:
                return simplified
        return ""

    def _build_manufacturing_items(
        self,
        rows: list[dict[str, str]],
        source_row: dict[str, str] | None = None,
    ) -> tuple[list[ManufacturingItem], bool]:
        grouped: dict[tuple[str, str], list[dict[str, str]]] = defaultdict(list)
        for row in rows:
            category = self._get_value(row, "Cat")
            if category not in {"Cat1", "Cat2"}:
                continue
            key = (self._get_value(row, "項目", "項目類型"), self._get_value(row, "攪料程序"))
            grouped[key].append(row)

        items: list[ManufacturingItem] = []
        for (_, process_name), group_rows in grouped.items():
            batches = max(self._to_int(self._get_value(row, "需求批次", "製造批次")) for row in group_rows)
            premix_minutes = sum(
                self._to_float(self._get_value(row, "工作(min)", "工作 min"))
                for row in group_rows
                if self._get_value(row, "Cat") == "Cat1"
            )
            mix_clean_minutes = sum(
                self._to_float(self._get_value(row, "工作(min)", "工作 min"))
                for row in group_rows
                if self._get_value(row, "Cat") == "Cat2"
            )
            mix_machine = next(
                (
                    self._get_value(row, "機器類型", "機器")
                    for row in group_rows
                    if self._get_value(row, "Cat") == "Cat2"
                ),
                "",
            )
            item_name = self._get_value(group_rows[0], "項目", "項目類型") or process_name

            if batches <= 0 or premix_minutes <= 0 or mix_clean_minutes <= 0 or not mix_machine:
                continue

            items.append(
                ManufacturingItem(
                    name=item_name,
                    process_name=process_name,
                    mix_machine=mix_machine,
                    batches=batches,
                    premix_minutes_per_batch=self._to_slot_minutes(premix_minutes),
                    mix_clean_minutes_per_batch=self._to_slot_minutes(mix_clean_minutes),
                )
            )
        if not items or source_row is None:
            return items, False

        source_core_code = self._get_value(source_row, "æ ¸å¿ƒå…­ç¢¼")
        manufacturing_efficiency_rows = self._matched_source_rows(
            self._core_code_candidates(source_core_code),
            "manufacturing_efficiency",
        )
        source_batches = self._to_int(self._get_value(source_row, "è£½é€ æ‰¹æ¬¡"))
        source_premix_total = self._to_float(self._get_value(source_row, "é æ”ªç¸½åˆ†é˜æ•¸"))
        source_mix_total = (
            self._to_float(self._get_value(source_row, "è£½é€ ç¸½æ™‚åˆ†é˜æ•¸*æ‰¹æ¬¡"))
            + self._to_float(self._get_value(source_row, "æ¸…æ´—ç¸½åˆ†é˜æ•¸"))
        )
        source_batches = source_batches or self._to_int(self._get_value(source_row, "\u9700\u6c42\u6279\u6b21"))
        source_premix_per_batch = self._to_float(self._get_value(source_row, "\u9810\u652a"))
        source_mix_per_batch = self._to_float(self._get_value(source_row, "\u652a\u6599"))
        source_clean_per_batch = self._to_float(self._get_value(source_row, "\u6e05\u6d17"))
        if source_batches > 0 and source_premix_total <= 0 and source_premix_per_batch > 0:
            source_premix_total = source_premix_per_batch * source_batches
        if source_batches > 0 and source_mix_total <= 0 and (source_mix_per_batch > 0 or source_clean_per_batch > 0):
            source_mix_total = (source_mix_per_batch + source_clean_per_batch) * source_batches
        source_machine = self._manufacturing_machine_from_source(source_row)
        source_process_names = self._split_process_names(self._get_value(source_row, "æ”ªæ–™ç¨‹åº"))
        source_single_item = self._get_value(source_row, "å–®ä¸€è£½é€ é …ç›®")

        if len(items) == 1:
            item = items[0]
            process_name = source_single_item or (
                source_process_names[0] if len(source_process_names) == 1 else item.process_name
            )
            updated_item = ManufacturingItem(
                name=source_single_item or item.name,
                process_name=process_name,
                mix_machine=(
                    source_machine
                    or self._manufacturing_machine_from_efficiency_rows(
                        manufacturing_efficiency_rows,
                        process_name,
                    )
                    or item.mix_machine
                ),
                batches=source_batches or item.batches,
                premix_minutes_per_batch=self._to_slot_minutes(
                    source_premix_total / source_batches
                    if source_batches > 0 and source_premix_total > 0
                    else item.premix_minutes_per_batch
                ),
                mix_clean_minutes_per_batch=self._to_slot_minutes(
                    source_mix_total / source_batches
                    if source_batches > 0 and source_mix_total > 0
                    else item.mix_clean_minutes_per_batch
                ),
            )
            return [updated_item], updated_item != item

        updated_items: list[ManufacturingItem] = []
        manufacturing_source_override = False
        for item in items:
            updated_machine = (
                self._manufacturing_machine_from_efficiency_rows(
                    manufacturing_efficiency_rows,
                    item.process_name,
                )
                or item.mix_machine
            )
            updated_item = ManufacturingItem(
                name=item.name,
                process_name=item.process_name,
                mix_machine=updated_machine,
                batches=item.batches,
                premix_minutes_per_batch=item.premix_minutes_per_batch,
                mix_clean_minutes_per_batch=item.mix_clean_minutes_per_batch,
            )
            updated_items.append(updated_item)
            manufacturing_source_override = manufacturing_source_override or updated_item != item
        return updated_items, manufacturing_source_override

    def _build_filling_item(
        self,
        rows: list[dict[str, str]],
        source_row: dict[str, str] | None = None,
    ) -> FillingItem | None:
        bottling_rows = [row for row in rows if self._get_value(row, "Cat") == "Cat3"]
        filling_rows = [row for row in rows if self._get_value(row, "Cat") == "Cat4"]
        if not filling_rows:
            return None

        time_ready_bottling_minutes = sum(
            self._to_float(self._get_value(row, "充填類型工序耗時 (h)", "排瓶時數")) for row in bottling_rows
        )
        time_ready_filling_minutes = sum(
            self._to_float(self._get_value(row, "充填類型工序耗時 (h)", "充填時數")) for row in filling_rows
        )

        source_bottling_minutes = self._to_float(self._get_value(source_row or {}, "排瓶時數")) * 60.0
        source_filling_minutes = self._to_float(self._get_value(source_row or {}, "充填時數")) * 60.0
        bottling_minutes = (
            source_bottling_minutes
            if source_bottling_minutes > 0 or not bottling_rows
            else time_ready_bottling_minutes
        )
        filling_minutes = source_filling_minutes if source_filling_minutes > 0 else time_ready_filling_minutes
        bottling_minutes = self._to_slot_minutes(bottling_minutes)
        filling_minutes = self._to_slot_minutes(filling_minutes)

        source_core_code = self._get_value(source_row or {}, "核心六碼")
        filling_efficiency_rows = self._matched_source_rows(
            self._core_code_candidates(source_core_code),
            "filling_efficiency",
        )
        bottling_efficiency_rows = [
            row for row in filling_efficiency_rows if self._get_value(row, "工作") == "排瓶"
        ]
        filling_efficiency_rows = [
            row for row in filling_efficiency_rows if self._get_value(row, "工作") == "充填"
        ]

        return FillingItem(
            bottling_machine=next(
                (
                    self._get_value(row, "機器", "機器類型")
                    for row in bottling_efficiency_rows + bottling_rows
                    if self._get_value(row, "機器", "機器類型")
                ),
                "排瓶機",
            ),
            filling_machine=next(
                (
                    self._get_value(row, "機器", "機器類型")
                    for row in filling_efficiency_rows + filling_rows
                    if self._get_value(row, "機器", "機器類型")
                ),
                "充填機",
            ),
            bottling_machine_minutes=bottling_minutes,
            filling_machine_minutes=filling_minutes,
            bottling_people=max(
                [
                    self._to_int(self._get_value(row, "排瓶產能人數", "表定產能人數"))
                    for row in ([source_row] if source_row else []) + bottling_efficiency_rows + bottling_rows
                ]
                or [1]
            ),
            filling_people=max(
                [
                    self._to_int(self._get_value(row, "充填產能人數", "表定產能人數"))
                    for row in ([source_row] if source_row else []) + filling_efficiency_rows + filling_rows
                ]
                or [4]
            ),
        )

    @staticmethod
    def _to_float(value: str) -> float:
        try:
            return float(value)
        except (TypeError, ValueError):
            return 0.0

    @staticmethod
    def _to_int(value: str) -> int:
        return int(round(WorkbookLoader._to_float(value)))

    @staticmethod
    def _to_slot_minutes(value: float) -> int:
        if value <= 0:
            return 0
        return int(math.ceil(value / SLOT_MINUTES) * SLOT_MINUTES)

    @staticmethod
    def _excel_date(value: str) -> date:
        if not value:
            raise ValueError("日期欄位不可為空")
        serial = float(value)
        return (EXCEL_EPOCH + timedelta(days=serial)).date()


class GreedyScheduleDecoder:
    """把訂單排序解碼成可評估的排程。"""

    infeasible_weight = 100_000.0
    minute_epsilon = 1e-3

    def __init__(self) -> None:
        self.daily_labor_usage: dict[tuple[str, date], float] = defaultdict(float)
        self.daily_machine_usage: dict[tuple[str, date], float] = defaultdict(float)
        self.task_records: list[ScheduledTask] = []

    @staticmethod
    def _assigned_bottling_people(filling_item: FillingItem) -> int:
        return max(filling_item.bottling_people, BOTTLING_MAX_PEOPLE)

    @staticmethod
    def _minutes_to_slots(minutes: float) -> int:
        if minutes <= 0:
            return 0
        return int(math.ceil(minutes / SLOT_MINUTES - 1e-9))

    @staticmethod
    def _slots_to_minutes(slots: int) -> int:
        return max(slots, 0) * SLOT_MINUTES

    @staticmethod
    def _aligned_bottling_slot_count(
        fill_slots: int,
        remaining_fill_slots: int,
        remaining_bottling_slots: int,
    ) -> int:
        if fill_slots <= 0 or remaining_bottling_slots <= 0:
            return 0
        if fill_slots >= remaining_fill_slots:
            return remaining_bottling_slots
        proportional_slots = math.ceil(fill_slots * remaining_bottling_slots / remaining_fill_slots)
        return max(1, min(remaining_bottling_slots, proportional_slots))

    def decode(self, chromosome: list[Order]) -> ScheduleResult:
        for order in chromosome:
            self._schedule_order(order)

        overtime_by_group = self._overtime_by_group()

        return ScheduleResult(
            total_overtime_minutes=self._total_overtime_minutes(),
            infeasible_penalty=self._capacity_penalty(),
            target_penalty=self._target_penalty(overtime_by_group),
            overtime_by_group=overtime_by_group,
            weekend_overtime_minutes=self._weekend_overtime_minutes(),
            task_records=sorted(
                self.task_records,
                key=lambda item: (item.scheduled_date, item.order_id, item.task_type, item.detail),
            ),
            daily_labor_usage=dict(self.daily_labor_usage),
            daily_machine_usage=dict(self.daily_machine_usage),
        )

    def _schedule_order(self, order: Order) -> None:
        labor_snapshot = dict(self.daily_labor_usage)
        machine_snapshot = dict(self.daily_machine_usage)
        task_count_snapshot = len(self.task_records)

        first_fill_day = self._schedule_filling(order)
        if first_fill_day is None:
            self._restore_snapshot(labor_snapshot, machine_snapshot, task_count_snapshot)
            self.daily_labor_usage[("filling", date.min)] += (
                order.filling_item.filling_machine_minutes * self.infeasible_weight
            )
            return

        earliest_manufacturing_day = max(
            order.raw_ready_date,
            order.due_date - timedelta(days=30),
        )
        latest_mix_day = first_fill_day - timedelta(days=1)

        for item in order.manufacturing_items:
            for batch_index in range(1, item.batches + 1):
                ok = self._schedule_manufacturing_batch(
                    order=order,
                    item=item,
                    batch_index=batch_index,
                    earliest_manufacturing_day=earliest_manufacturing_day,
                    latest_mix_day=latest_mix_day,
                )
                if not ok:
                    self._restore_snapshot(labor_snapshot, machine_snapshot, task_count_snapshot)
                    self.daily_labor_usage[("manufacturing", date.min)] += self.infeasible_weight
                    return

    def _restore_snapshot(
        self,
        labor_snapshot: dict[tuple[str, date], float],
        machine_snapshot: dict[tuple[str, date], float],
        task_count_snapshot: int,
    ) -> None:
        self.daily_labor_usage = defaultdict(float, labor_snapshot)
        self.daily_machine_usage = defaultdict(float, machine_snapshot)
        del self.task_records[task_count_snapshot:]

    def _schedule_filling(self, order: Order) -> date | None:
        filling_item = order.filling_item
        assigned_bottling_people = self._assigned_bottling_people(filling_item)
        remaining_fill_minutes = filling_item.filling_machine_minutes
        fill_to_bottling_ratio = (
            (
                filling_item.bottling_machine_minutes
                * filling_item.bottling_people
                / assigned_bottling_people
            )
            / filling_item.filling_machine_minutes
            if filling_item.filling_machine_minutes > 0
            else 0
        )

        first_fill_day: date | None = None
        candidate_days = self._date_range(order.material_ready_date, order.due_date)

        for use_max_capacity in (False, True):
            while remaining_fill_minutes > self.minute_epsilon:
                best_option: tuple[float, int, float, date, float] | None = None
                for current_day in reversed(candidate_days):
                    bottling_day = current_day - timedelta(days=1)
                    if bottling_day < order.material_ready_date:
                        if allows_same_day_bottling(order, current_day):
                            bottling_day = current_day
                        else:
                            continue

                    capacity = self._max_fill_chunk(
                        current_day=current_day,
                        bottling_day=bottling_day,
                        fill_machine=filling_item.filling_machine,
                        bottling_machine=filling_item.bottling_machine,
                        fill_people=filling_item.filling_people,
                        bottling_people=assigned_bottling_people,
                        fill_to_bottling_ratio=fill_to_bottling_ratio,
                        use_max_capacity=use_max_capacity,
                    )
                    chunk = min(remaining_fill_minutes, capacity)
                    if chunk <= self.minute_epsilon:
                        continue

                    added_overtime = self._added_overtime_for_fill(
                        current_day=current_day,
                        bottling_day=bottling_day,
                        fill_machine_minutes=chunk,
                        filling_item=filling_item,
                        bottling_people=assigned_bottling_people,
                        fill_to_bottling_ratio=fill_to_bottling_ratio,
                    )
                    if added_overtime is None:
                        continue

                    candidate = (
                        added_overtime,
                        -current_day.toordinal(),
                        -chunk,
                        current_day,
                        chunk,
                    )
                    if best_option is None or candidate < best_option:
                        best_option = candidate

                if best_option is None:
                    break

                _, _, _, current_day, chunk = best_option
                bottling_day = current_day - timedelta(days=1)
                if bottling_day < order.material_ready_date and allows_same_day_bottling(order, current_day):
                    bottling_day = current_day
                self._commit_filling_chunk(
                    order,
                    current_day,
                    bottling_day,
                    chunk,
                    filling_item,
                    assigned_bottling_people,
                    fill_to_bottling_ratio,
                )
                remaining_fill_minutes -= chunk
                if remaining_fill_minutes < self.minute_epsilon:
                    remaining_fill_minutes = 0.0
                if first_fill_day is None or current_day < first_fill_day:
                    first_fill_day = current_day

        if remaining_fill_minutes > self.minute_epsilon:
            return None
        return first_fill_day

    def _schedule_manufacturing_batch(
        self,
        order: Order,
        item: ManufacturingItem,
        batch_index: int,
        earliest_manufacturing_day: date,
        latest_mix_day: date,
    ) -> bool:
        best_day: date | None = None
        best_increment = math.inf

        for mix_day in reversed(self._date_range(earliest_manufacturing_day + timedelta(days=1), latest_mix_day)):
            premix_day = mix_day - timedelta(days=1)
            increment = self._added_overtime_for_manufacturing(item, premix_day, mix_day)
            if increment is None:
                continue
            if increment < best_increment:
                best_increment = increment
                best_day = mix_day

        if best_day is None:
            return False

        premix_day = best_day - timedelta(days=1)
        self._commit_task(
            order_id=order.order_id,
            task_type="Cat1",
            detail=f"{item.name}-{item.process_name}-第{batch_index}批預攪",
            scheduled_date=premix_day,
            machine_type="預攪機",
            machine_minutes=item.premix_minutes_per_batch,
            labor_group="manufacturing",
            labor_minutes=item.premix_minutes_per_batch,
        )
        self._commit_task(
            order_id=order.order_id,
            task_type="Cat2",
            detail=f"{item.name}-{item.process_name}-第{batch_index}批攪料清洗",
            scheduled_date=best_day,
            machine_type=item.mix_machine,
            machine_minutes=item.mix_clean_minutes_per_batch,
            labor_group="manufacturing",
            labor_minutes=item.mix_clean_minutes_per_batch,
        )
        return True

    def _max_fill_chunk(
        self,
        current_day: date,
        bottling_day: date,
        fill_machine: str,
        bottling_machine: str,
        fill_people: int,
        bottling_people: int,
        fill_to_bottling_ratio: float,
        use_max_capacity: bool,
    ) -> float:
        fill_labor_cap = self._remaining_labor_capacity("filling", current_day, use_max_capacity)
        bottling_labor_cap = self._remaining_labor_capacity("bottling", bottling_day, use_max_capacity)
        packaging_labor_cap = self._remaining_packaging_capacity(current_day, use_max_capacity)
        fill_machine_cap = self._remaining_machine_capacity_for_phase(
            fill_machine,
            current_day,
            use_max_capacity,
        )
        bottling_machine_cap = self._remaining_machine_capacity_for_phase(
            bottling_machine,
            bottling_day,
            use_max_capacity,
        )

        candidates = [fill_machine_cap, fill_labor_cap / max(fill_people, 1)]
        if fill_to_bottling_ratio > 0:
            candidates.append(bottling_machine_cap / fill_to_bottling_ratio)
            candidates.append(bottling_labor_cap / (fill_to_bottling_ratio * max(bottling_people, 1)))
            candidates.append(
                packaging_labor_cap
                / (fill_people + fill_to_bottling_ratio * max(bottling_people, 1))
            )
        else:
            candidates.append(packaging_labor_cap / max(fill_people, 1))
        return max(min(candidates), 0.0)

    def _added_overtime_for_manufacturing(
        self,
        item: ManufacturingItem,
        premix_day: date,
        mix_day: date,
    ) -> float | None:
        if self._remaining_machine_capacity("預攪機", premix_day) < item.premix_minutes_per_batch:
            return None
        if self._remaining_machine_capacity(item.mix_machine, mix_day) < item.mix_clean_minutes_per_batch:
            return None

        premix_increment = self._incremental_labor_overtime(
            labor_group="manufacturing",
            current_day=premix_day,
            add_minutes=item.premix_minutes_per_batch,
        )
        mix_increment = self._incremental_labor_overtime(
            labor_group="manufacturing",
            current_day=mix_day,
            add_minutes=item.mix_clean_minutes_per_batch,
        )
        if premix_increment is None or mix_increment is None:
            return None
        return premix_increment + mix_increment

    def _added_overtime_for_fill(
        self,
        current_day: date,
        bottling_day: date,
        fill_machine_minutes: float,
        filling_item: FillingItem,
        bottling_people: int,
        fill_to_bottling_ratio: float,
    ) -> float | None:
        fill_increment = self._incremental_machine_overtime(
            machine_type=filling_item.filling_machine,
            current_day=current_day,
            add_minutes=fill_machine_minutes,
        )
        if fill_increment is None:
            return None

        bottling_minutes = fill_machine_minutes * fill_to_bottling_ratio
        bottling_increment = self._incremental_machine_overtime(
            machine_type=filling_item.bottling_machine,
            current_day=bottling_day,
            add_minutes=bottling_minutes,
        )
        if bottling_increment is None:
            return None

        packaging_increment = self._incremental_packaging_overtime(
            current_day=current_day,
            bottling_labor_minutes=(
                bottling_minutes * bottling_people if bottling_day == current_day else 0.0
            ),
            filling_labor_minutes=fill_machine_minutes * filling_item.filling_people,
        )
        if packaging_increment is None:
            return None

        next_day_packaging_increment = 0.0
        if bottling_day != current_day:
            next_day_packaging_increment = self._incremental_packaging_overtime(
                current_day=bottling_day,
                bottling_labor_minutes=bottling_minutes * bottling_people,
                filling_labor_minutes=0.0,
            )
            if next_day_packaging_increment is None:
                return None

        return fill_increment + bottling_increment + packaging_increment + next_day_packaging_increment

    def _commit_filling_chunk(
        self,
        order: Order,
        current_day: date,
        bottling_day: date,
        fill_machine_minutes: float,
        filling_item: FillingItem,
        bottling_people: int,
        fill_to_bottling_ratio: float,
    ) -> None:
        if fill_machine_minutes <= self.minute_epsilon:
            return
        bottling_machine_minutes = fill_machine_minutes * fill_to_bottling_ratio

        if bottling_machine_minutes > self.minute_epsilon:
            self._commit_task(
                order_id=order.order_id,
                task_type="Cat3",
                detail=f"{order.product_name}-排瓶",
                scheduled_date=bottling_day,
                machine_type=filling_item.bottling_machine,
                machine_minutes=bottling_machine_minutes,
                labor_group="bottling",
                labor_minutes=bottling_machine_minutes * bottling_people,
            )

        self._commit_task(
            order_id=order.order_id,
            task_type="Cat4",
            detail=f"{order.product_name}-充填",
            scheduled_date=current_day,
            machine_type=filling_item.filling_machine,
            machine_minutes=fill_machine_minutes,
            labor_group="filling",
            labor_minutes=fill_machine_minutes * filling_item.filling_people,
        )

    def _commit_task(
        self,
        order_id: str,
        task_type: str,
        detail: str,
        scheduled_date: date,
        machine_type: str,
        machine_minutes: float,
        labor_group: str,
        labor_minutes: float,
    ) -> None:
        if machine_minutes <= self.minute_epsilon and labor_minutes <= self.minute_epsilon:
            return
        self.daily_machine_usage[(machine_type, scheduled_date)] += machine_minutes
        self.daily_labor_usage[(labor_group, scheduled_date)] += labor_minutes
        self.task_records.append(
            ScheduledTask(
                order_id=order_id,
                task_type=task_type,
                detail=detail,
                scheduled_date=scheduled_date,
                machine_type=machine_type,
                machine_minutes=machine_minutes,
                labor_group=labor_group,
                labor_minutes=labor_minutes,
            )
        )

    def _remaining_labor_capacity(
        self,
        labor_group: str,
        current_day: date,
        use_max_capacity: bool,
    ) -> float:
        config = LABOR_CONFIG[labor_group]
        used = self.daily_labor_usage[(labor_group, current_day)]
        capacity = config["max_minutes"] if use_max_capacity else config["regular_minutes"]
        return max(capacity - used, 0.0)

    def _remaining_machine_capacity(self, machine_type: str, current_day: date) -> float:
        used = self.daily_machine_usage[(machine_type, current_day)]
        machine_count = MACHINE_COUNTS.get(machine_type, 1)
        return max(machine_count * MACHINE_DAILY_CAPACITY - used, 0.0)

    def _remaining_machine_capacity_for_phase(
        self,
        machine_type: str,
        current_day: date,
        use_max_capacity: bool,
    ) -> float:
        used = self.daily_machine_usage[(machine_type, current_day)]
        machine_count = MACHINE_COUNTS.get(machine_type, 1)
        phase_capacity = MACHINE_DAILY_CAPACITY if use_max_capacity else PROCESS_DAILY_REGULAR_MINUTES
        return max(machine_count * phase_capacity - used, 0.0)

    def _incremental_machine_overtime(
        self,
        machine_type: str,
        current_day: date,
        add_minutes: float,
    ) -> float | None:
        used = self.daily_machine_usage[(machine_type, current_day)]
        machine_count = MACHINE_COUNTS.get(machine_type, 1)
        max_minutes = machine_count * MACHINE_DAILY_CAPACITY
        regular_minutes = machine_count * PROCESS_DAILY_REGULAR_MINUTES
        if used + add_minutes > max_minutes + 1e-6:
            return None
        before = max(used - regular_minutes, 0.0)
        after = max(used + add_minutes - regular_minutes, 0.0)
        return after - before

    def _packaging_labor_used(self, current_day: date) -> float:
        return (
            self.daily_labor_usage[("bottling", current_day)]
            + self.daily_labor_usage[("filling", current_day)]
        )

    def _remaining_packaging_capacity(self, current_day: date, use_max_capacity: bool) -> float:
        capacity = (
            PACKAGING_TEAM["max_minutes"]
            if use_max_capacity
            else PACKAGING_TEAM["regular_minutes"]
        )
        return max(capacity - self._packaging_labor_used(current_day), 0.0)

    def _incremental_packaging_overtime(
        self,
        current_day: date,
        bottling_labor_minutes: float,
        filling_labor_minutes: float,
    ) -> float | None:
        used = self._packaging_labor_used(current_day)
        add_minutes = bottling_labor_minutes + filling_labor_minutes
        if used + add_minutes > PACKAGING_TEAM["max_minutes"] + 1e-6:
            return None
        before = max(used - PACKAGING_TEAM["regular_minutes"], 0.0)
        after = max(used + add_minutes - PACKAGING_TEAM["regular_minutes"], 0.0)
        return after - before

    def _incremental_labor_overtime(
        self,
        labor_group: str,
        current_day: date,
        add_minutes: float,
    ) -> float | None:
        config = LABOR_CONFIG[labor_group]
        used = self.daily_labor_usage[(labor_group, current_day)]
        if used + add_minutes > config["max_minutes"] + 1e-6:
            return None
        before = max(used - config["regular_minutes"], 0.0)
        after = max(used + add_minutes - config["regular_minutes"], 0.0)
        return after - before

    def _process_overtime_by_group(self) -> dict[str, float]:
        overtime = {process: 0.0 for process in PROCESS_OVERTIME_MACHINES}
        process_usage: dict[tuple[str, date], float] = defaultdict(float)
        for task in self.task_records:
            if task.task_type == "Cat3":
                process_usage[("bottling", task.scheduled_date)] += task.labor_minutes
            elif task.task_type == "Cat4":
                process_usage[("filling", task.scheduled_date)] += task.machine_minutes

        for (process, current_day), used_minutes in process_usage.items():
            if current_day == date.min:
                continue
            regular_minutes = PROCESS_OVERTIME_REGULAR_MINUTES.get(
                process,
                PROCESS_DAILY_REGULAR_MINUTES,
            )
            overtime[process] += max(used_minutes - regular_minutes, 0.0)
        return overtime

    def _total_overtime_minutes(self) -> float:
        total = 0.0
        for (labor_group, current_day), used_minutes in self.daily_labor_usage.items():
            if current_day == date.min or labor_group in {"bottling", "filling"}:
                continue
            regular = LABOR_CONFIG[labor_group]["regular_minutes"]
            total += max(used_minutes - regular, 0.0)
        process_overtime = self._process_overtime_by_group()
        total += sum(process_overtime.values())
        return total

    def _overtime_by_group(self) -> dict[str, float]:
        overtime_by_group = {"manufacturing": 0.0, "filling": 0.0, "bottling": 0.0}
        for (labor_group, current_day), used_minutes in self.daily_labor_usage.items():
            if current_day == date.min or labor_group in {"bottling", "filling"}:
                continue
            regular = LABOR_CONFIG[labor_group]["regular_minutes"]
            overtime_by_group[labor_group] += max(used_minutes - regular, 0.0)
        process_overtime = self._process_overtime_by_group()
        overtime_by_group["bottling"] = process_overtime["bottling"]
        overtime_by_group["filling"] = process_overtime["filling"]
        return overtime_by_group

    def _weekend_overtime_minutes(self) -> float:
        total = 0.0
        for (labor_group, current_day), used_minutes in self.daily_labor_usage.items():
            if current_day == date.min or current_day.weekday() < 5 or labor_group in {"bottling", "filling"}:
                continue
            regular = LABOR_CONFIG[labor_group]["regular_minutes"]
            total += max(used_minutes - regular, 0.0)
        process_usage: dict[tuple[str, date], float] = defaultdict(float)
        for task in self.task_records:
            if task.task_type == "Cat3":
                process_usage[("bottling", task.scheduled_date)] += task.labor_minutes
            elif task.task_type == "Cat4":
                process_usage[("filling", task.scheduled_date)] += task.machine_minutes
        for (process, current_day), used_minutes in process_usage.items():
            if current_day == date.min or current_day.weekday() < 5:
                continue
            regular_minutes = PROCESS_OVERTIME_REGULAR_MINUTES.get(
                process,
                PROCESS_DAILY_REGULAR_MINUTES,
            )
            total += max(used_minutes - regular_minutes, 0.0)
        return total

    def _manufacturing_switch_cleaning_penalty(self) -> float:
        machine_day_processes: dict[tuple[date, str], set[str]] = defaultdict(set)
        for task in self.task_records:
            if task.task_type not in {"Cat1", "Cat2"}:
                continue
            if task.machine_type == "é æ”ªæ©Ÿ":
                continue
            process_key = task.detail.split("-ç¬¬", 1)[0]
            machine_day_processes[(task.scheduled_date, task.machine_type)].add(process_key)

        penalty = 0.0
        for process_keys in machine_day_processes.values():
            normalized_process_keys = {
                _normalize_manufacturing_detail(process_key) for process_key in process_keys
            }
            switch_count = max(len(normalized_process_keys) - 1, 0)
            penalty += switch_count * MANUFACTURING_SWITCH_CLEANING_WEIGHT
        return penalty

    def _target_penalty(self, overtime_by_group: dict[str, float]) -> float:
        penalty = 0.0
        for labor_group, target_minutes in OVERTIME_TARGETS.items():
            overtime = overtime_by_group.get(labor_group, 0.0)
            if overtime > target_minutes:
                penalty += (overtime - target_minutes) * TARGET_EXCESS_WEIGHT
        penalty += self._weekend_overtime_minutes() * WEEKEND_OVERTIME_WEIGHT
        penalty += self._manufacturing_switch_cleaning_penalty()
        return penalty

    def _capacity_penalty(self) -> float:
        penalty = 0.0

        for (labor_group, current_day), used_minutes in self.daily_labor_usage.items():
            if current_day == date.min:
                penalty += used_minutes
                continue
            max_minutes = LABOR_CONFIG[labor_group]["max_minutes"]
            if used_minutes > max_minutes + 1e-6:
                penalty += (used_minutes - max_minutes) * self.infeasible_weight

        packaging_days = {
            current_day
            for labor_group, current_day in self.daily_labor_usage
            if labor_group in {"bottling", "filling"} and current_day != date.min
        }
        for current_day in packaging_days:
            packaging_used = self._packaging_labor_used(current_day)
            if packaging_used > PACKAGING_TEAM["max_minutes"] + 1e-6:
                penalty += (packaging_used - PACKAGING_TEAM["max_minutes"]) * self.infeasible_weight

        for (machine_type, _current_day), used_minutes in self.daily_machine_usage.items():
            max_minutes = MACHINE_COUNTS.get(machine_type, 1) * MACHINE_DAILY_CAPACITY
            if used_minutes > max_minutes + 1e-6:
                penalty += (used_minutes - max_minutes) * self.infeasible_weight

        monthly_usage: dict[tuple[str, int, int], float] = defaultdict(float)
        for (labor_group, current_day), used_minutes in self.daily_labor_usage.items():
            if current_day == date.min or labor_group in {"bottling", "filling"}:
                continue
            overtime = max(used_minutes - LABOR_CONFIG[labor_group]["regular_minutes"], 0.0)
            monthly_usage[(labor_group, current_day.year, current_day.month)] += overtime

        packaging_monthly_usage: dict[tuple[int, int], float] = defaultdict(float)
        for current_day in packaging_days:
            overtime = max(self._packaging_labor_used(current_day) - PACKAGING_TEAM["regular_minutes"], 0.0)
            packaging_monthly_usage[(current_day.year, current_day.month)] += overtime

        monthly_limits = {
            "manufacturing": 2 * 46 * 60,
        }
        for (labor_group, _year, _month), overtime in monthly_usage.items():
            if overtime > monthly_limits[labor_group]:
                penalty += (overtime - monthly_limits[labor_group]) * self.infeasible_weight

        for (_year, _month), overtime in packaging_monthly_usage.items():
            if overtime > PACKAGING_TEAM["monthly_overtime_limit"]:
                penalty += (overtime - PACKAGING_TEAM["monthly_overtime_limit"]) * self.infeasible_weight

        return penalty

    @staticmethod
    def _date_range(start_day: date, end_day: date) -> list[date]:
        if end_day < start_day:
            return []
        days = (end_day - start_day).days
        return [start_day + timedelta(days=offset) for offset in range(days + 1)]


@dataclass
class GeneticScheduler:
    population_size: int = 40
    generations: int = 120
    mutation_rate: float = 0.15
    crossover_rate: float = 0.85
    elite_count: int = 4
    tournament_size: int = 3
    seed: int = 42
    seeded_order_ids: tuple[str, ...] = ()

    def optimize(self, orders: list[Order]) -> ScheduleResult:
        random.seed(self.seed)
        population = self._initialize_population(orders)
        scored_population = self._score_population(population)
        best_ordering, best_result = min(scored_population, key=lambda item: item[1].objective)

        for generation in range(1, self.generations + 1):
            next_population = [chromosome[:] for chromosome, _ in scored_population[: self.elite_count]]
            while len(next_population) < self.population_size:
                parent_a = self._select_parent(scored_population)
                parent_b = self._select_parent(scored_population)
                child_a, child_b = self._crossover(parent_a, parent_b)
                self._mutate(child_a)
                self._mutate(child_b)
                next_population.append(child_a)
                if len(next_population) < self.population_size:
                    next_population.append(child_b)

            scored_population = self._score_population(next_population)
            generation_best_ordering, generation_best_result = scored_population[0]
            if generation_best_result.objective < best_result.objective:
                best_ordering = generation_best_ordering[:]
                best_result = generation_best_result

            print(
                f"Generation {generation:03d} | "
                f"best objective={generation_best_result.objective:.1f} | "
                f"overtime={generation_best_result.total_overtime_minutes:.1f} | "
                f"penalty={generation_best_result.infeasible_penalty:.1f} | "
                f"target_penalty={generation_best_result.target_penalty:.1f}"
            )

        return GreedyScheduleDecoder().decode(best_ordering)

    def _initialize_population(self, orders: list[Order]) -> list[list[Order]]:
        base = sorted(orders, key=lambda item: (item.due_date, item.material_ready_date, item.order_id))
        population = [base]
        seeded = self._build_seeded_chromosome(orders)
        if seeded is not None and seeded != base:
            population.append(seeded)
        population.append(sorted(orders, key=lambda item: (item.material_ready_date, item.due_date)))
        population.append(sorted(orders, key=lambda item: (item.raw_ready_date, item.due_date)))
        population.append(
            sorted(
                orders,
                key=lambda item: (
                    item.due_date,
                    sum(m.batches for m in item.manufacturing_items),
                    item.filling_item.filling_machine_minutes,
                ),
            )
        )
        while len(population) < self.population_size:
            chromosome = base[:]
            random.shuffle(chromosome)
            population.append(chromosome)
        return population

    def _build_seeded_chromosome(self, orders: list[Order]) -> list[Order] | None:
        if not self.seeded_order_ids:
            return None

        order_map = {order.order_id: order for order in orders}
        seeded_orders: list[Order] = []
        seen: set[str] = set()
        for order_id in self.seeded_order_ids:
            order = order_map.get(order_id)
            if order is None or order_id in seen:
                continue
            seeded_orders.append(order)
            seen.add(order_id)

        if not seeded_orders:
            return None

        remaining_orders = [
            order
            for order in sorted(orders, key=lambda item: (item.due_date, item.material_ready_date, item.order_id))
            if order.order_id not in seen
        ]
        return seeded_orders + remaining_orders

    def _score_population(
        self,
        population: Iterable[list[Order]],
    ) -> list[tuple[list[Order], ScheduleResult]]:
        scored = []
        for chromosome in population:
            scored.append((chromosome[:], GreedyScheduleDecoder().decode(chromosome)))
        scored.sort(key=lambda item: item[1].objective)
        return scored

    def _select_parent(self, scored_population: list[tuple[list[Order], ScheduleResult]]) -> list[Order]:
        candidates = random.sample(scored_population, k=min(self.tournament_size, len(scored_population)))
        return min(candidates, key=lambda item: item[1].objective)[0][:]

    def _crossover(
        self,
        parent_a: list[Order],
        parent_b: list[Order],
    ) -> tuple[list[Order], list[Order]]:
        if random.random() > self.crossover_rate:
            return parent_a[:], parent_b[:]
        start = random.randint(0, len(parent_a) - 2)
        end = random.randint(start + 1, len(parent_a) - 1)
        return (
            self._ordered_crossover(parent_a, parent_b, start, end),
            self._ordered_crossover(parent_b, parent_a, start, end),
        )

    def _ordered_crossover(
        self,
        donor: list[Order],
        receiver: list[Order],
        start: int,
        end: int,
    ) -> list[Order]:
        middle = donor[start : end + 1]
        remainder = [item for item in receiver if item not in middle]
        return remainder[:start] + middle + remainder[start:]

    def _mutate(self, chromosome: list[Order]) -> None:
        if random.random() >= self.mutation_rate:
            return
        left = random.randint(0, len(chromosome) - 1)
        right = random.randint(0, len(chromosome) - 1)
        chromosome[left], chromosome[right] = chromosome[right], chromosome[left]


def write_schedule_outputs(schedule: ScheduleResult, output_dir: Path) -> None:
    """把任務明細與每日負荷輸出成 CSV。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    summary = build_schedule_summary(schedule)

    task_path = output_dir / "best_schedule_tasks.csv"
    with task_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "訂單",
                "類別",
                "明細",
                "日期",
                "機器類型",
                "機器分鐘",
                "人力類型",
                "人力分鐘",
            ]
        )
        for task in schedule.task_records:
            writer.writerow(
                [
                    task.order_id,
                    _display_task_stage(task.task_type),
                    _display_task_detail(task),
                    task.scheduled_date.isoformat(),
                    _display_machine_label(task.machine_type),
                    _display_numeric(task.machine_minutes),
                    _display_labor_group(task.labor_group),
                    _display_numeric(task.labor_minutes),
                ]
            )
        writer.writerow(
            [
                "總計",
                "摘要",
                (
                    f"總加班分鐘數={_display_numeric(summary['total_overtime_minutes'])}; "
                    f"排瓶超時分鐘數={_display_numeric(summary['bottling_overtime_minutes'])}; "
                    f"充填超時分鐘數={_display_numeric(summary['filling_overtime_minutes'])}; "
                    f"總排瓶分鐘數={_display_numeric(summary['total_bottling_minutes'])}; "
                    f"總充填分鐘數={_display_numeric(summary['total_filling_minutes'])}"
                ),
                "",
                "",
                "",
                "",
                "",
            ]
        )

    daily_path = output_dir / "best_schedule_daily_load.csv"
    with daily_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "日期",
                "人力類型",
                "已用人力分鐘",
                "正常人力分鐘",
                "超時分鐘",
            ]
        )
        for (labor_group, current_day), used_minutes in sorted(schedule.daily_labor_usage.items()):
            if current_day == date.min:
                continue
            regular_minutes = LABOR_CONFIG[labor_group]["regular_minutes"]
            writer.writerow(
                [
                    current_day.isoformat(),
                    _display_labor_group(labor_group),
                    _display_integer(used_minutes),
                    regular_minutes,
                    _display_integer(max(used_minutes - regular_minutes, 0.0)),
                ]
            )

    summary_path = output_dir / "schedule_summary.csv"
    with summary_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "總加班分鐘數",
                "調製超時分鐘數",
                "排瓶超時分鐘數",
                "充填超時分鐘數",
                "週末加班分鐘數",
                "總排瓶分鐘數",
                "總充填分鐘數",
            ]
        )
        writer.writerow(
            [
                _display_integer(summary["total_overtime_minutes"]),
                _display_integer(summary["manufacturing_overtime_minutes"]),
                _display_integer(summary["bottling_overtime_minutes"]),
                _display_integer(summary["filling_overtime_minutes"]),
                _display_integer(summary["weekend_overtime_minutes"]),
                _display_integer(summary["total_bottling_minutes"]),
                _display_integer(summary["total_filling_minutes"]),
            ]
        )


def write_daily_manufacturing_schedule_output(
    schedule: ScheduleResult,
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    manufacturing_tasks = [
        task for task in schedule.task_records if task.task_type in {"Cat1", "Cat2"}
    ]
    machine_day_totals: dict[tuple[date, str], float] = defaultdict(float)
    order_day_totals: dict[tuple[date, str], float] = defaultdict(float)
    labor_day_totals: dict[date, float] = defaultdict(float)

    for task in manufacturing_tasks:
        machine_day_totals[(task.scheduled_date, task.machine_type)] += task.machine_minutes
        order_day_totals[(task.scheduled_date, task.order_id)] += task.machine_minutes
        labor_day_totals[task.scheduled_date] += task.labor_minutes

    path = output_dir / "daily_manufacturing_schedule.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "日期",
                "訂單",
                "類別",
                "明細",
                "機器類型",
                "機器分鐘",
                "人力分鐘",
                "當日機器總分鐘",
                "當日訂單總分鐘",
                "當日調製人力總分鐘",
            ]
        )
        for task in sorted(
            manufacturing_tasks,
            key=lambda item: (
                item.scheduled_date,
                item.machine_type,
                item.order_id,
                item.task_type,
                item.detail,
            ),
        ):
            writer.writerow(
                [
                    task.scheduled_date.isoformat(),
                    task.order_id,
                    _display_task_stage(task.task_type),
                    _display_task_detail(task),
                    _display_machine_label(task.machine_type),
                    _display_integer(task.machine_minutes),
                    _display_integer(task.labor_minutes),
                    _display_integer(machine_day_totals[(task.scheduled_date, task.machine_type)]),
                    _display_integer(order_day_totals[(task.scheduled_date, task.order_id)]),
                    _display_integer(labor_day_totals[task.scheduled_date]),
                ]
            )


def write_daily_packaging_schedule_output(
    schedule: ScheduleResult,
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    packaging_tasks = [
        task for task in schedule.task_records if task.task_type in {"Cat3", "Cat4"}
    ]
    process_day_totals: dict[tuple[date, str], float] = defaultdict(float)
    machine_day_totals: dict[tuple[date, str], float] = defaultdict(float)
    order_day_totals: dict[tuple[date, str], float] = defaultdict(float)
    labor_day_totals: dict[tuple[date, str], float] = defaultdict(float)

    for task in packaging_tasks:
        process = "bottling" if task.task_type == "Cat3" else "filling"
        process_day_totals[(task.scheduled_date, process)] += task.machine_minutes
        machine_day_totals[(task.scheduled_date, task.machine_type)] += task.machine_minutes
        order_day_totals[(task.scheduled_date, task.order_id)] += task.machine_minutes
        labor_day_totals[(task.scheduled_date, process)] += task.labor_minutes

    path = output_dir / "daily_packaging_schedule.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "日期",
                "訂單",
                "類別",
                "工序",
                "明細",
                "機器類型",
                "機器分鐘",
                "人力分鐘",
                "當日工序總分鐘",
                "當日機器總分鐘",
                "當日訂單總分鐘",
                "當日工序人力總分鐘",
                "當日正常分鐘",
                "當日超時分鐘",
            ]
        )
        for task in sorted(
            packaging_tasks,
            key=lambda item: (
                item.scheduled_date,
                item.task_type,
                item.machine_type,
                item.order_id,
                item.detail,
            ),
        ):
            process = "bottling" if task.task_type == "Cat3" else "filling"
            process_day_total = process_day_totals[(task.scheduled_date, process)]
            process_day_total_labor = labor_day_totals[(task.scheduled_date, process)]
            regular_minutes = PROCESS_OVERTIME_REGULAR_MINUTES.get(
                process,
                PROCESS_DAILY_REGULAR_MINUTES,
            )
            overtime_base_minutes = process_day_total_labor if process == "bottling" else process_day_total
            writer.writerow(
                [
                    task.scheduled_date.isoformat(),
                    task.order_id,
                    _display_task_stage(task.task_type),
                    "排瓶" if process == "bottling" else "充填",
                    _display_task_detail(task),
                    _display_machine_label(task.machine_type),
                    _display_numeric(task.machine_minutes),
                    _display_numeric(task.labor_minutes),
                    _display_numeric(process_day_total),
                    _display_numeric(machine_day_totals[(task.scheduled_date, task.machine_type)]),
                    _display_numeric(order_day_totals[(task.scheduled_date, task.order_id)]),
                    _display_numeric(process_day_total_labor),
                    regular_minutes,
                    _display_numeric(max(overtime_base_minutes - regular_minutes, 0.0)),
                ]
            )


def write_manufacturing_cleaning_audit_output(
    schedule: ScheduleResult,
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    machine_day_orders: dict[tuple[date, str], set[str]] = defaultdict(set)
    machine_day_details: dict[tuple[date, str], set[str]] = defaultdict(set)
    machine_day_minutes: dict[tuple[date, str], float] = defaultdict(float)

    for task in schedule.task_records:
        if task.task_type not in {"Cat1", "Cat2"}:
            continue
        key = (task.scheduled_date, task.machine_type)
        machine_day_orders[key].add(task.order_id)
        machine_day_details[key].add(task.detail.split("-第", 1)[0])
        machine_day_minutes[key] += task.machine_minutes

    path = output_dir / "manufacturing_cleaning_audit.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "日期",
                "機器類型",
                "當日機器總分鐘",
                "當日訂單數",
                "當日不同製程數",
                "推估換產品清洗次數",
                "訂單清單",
                "製程明細",
            ]
        )
        for (scheduled_date, machine_type), total_minutes in sorted(machine_day_minutes.items()):
            distinct_orders = sorted(machine_day_orders[(scheduled_date, machine_type)])
            distinct_details = sorted(
                {_normalize_output_text(detail) for detail in machine_day_details[(scheduled_date, machine_type)]}
            )
            distinct_process_count = len(distinct_details)
            writer.writerow(
                [
                    scheduled_date.isoformat(),
                    _display_machine_label(machine_type),
                    _display_numeric(total_minutes),
                    len(distinct_orders),
                    distinct_process_count,
                    max(distinct_process_count - 1, 0),
                    " | ".join(distinct_orders),
                    " | ".join(distinct_details),
                ]
            )


def _slot_index_to_time(slot_index: int) -> str:
    start_minutes = 9 * 60 + slot_index * 5
    hours, minutes = divmod(start_minutes, 60)
    return f"{hours:02d}:{minutes:02d}"


OUTPUT_TEXT_NORMALIZATION_MAP = {
    "ç¬¬": "第",
    "æ‰¹": "批",
    "é æ”ª": "預攪",
    "æ”ªæ–™æ¸…æ´—": "攪料清洗",
    "æŽ’ç“¶": "排瓶",
    "å……å¡«": "充填",
    "é æ”ªæ©Ÿ": "預攪機",
    "ç²¾æ©Ÿ": "精機",
    "å°ä¹³æ©Ÿ": "小乳機",
    "å¤§ä¹³æ©Ÿ": "大乳機",
    "å¤§æ©Ÿ": "大機",
    "æŽ’ç“¶æ©Ÿ": "排瓶機",
    "å……å¡«æ©Ÿ": "充填機",
    "å¤§åž‹æ”ªæ‹Œæ©Ÿ": "大型攪拌機",
    "å°åž‹æ”ªæ‹Œæ©Ÿ": "小型攪拌機",
    "50LçœŸç©ºä¹³åŒ–æ©Ÿ": "50L真空乳化機",
    "100LçœŸç©ºä¹³åŒ–æ©Ÿ": "100L真空乳化機",
    "100Lå‡æº«æ¡¶æ§½": "100L升溫桶槽",
    "400Lå¤§æ¡¶æ§½": "400L大桶槽",
    "500Lå¤§æ¡¶æ§½": "500L大桶槽",
    "è½åœ°åž‹å‡è³ªæ©Ÿ": "落地型均質機",
    "è¶…éŸ³æ³¢éœ‡ç›ªæ©Ÿ": "超音波震盪機",
}


def _normalize_output_text(text: str) -> str:
    normalized = str(text or "")
    for source, target in OUTPUT_TEXT_NORMALIZATION_MAP.items():
        normalized = normalized.replace(source, target)
    return normalized


def _display_machine_label(machine_type: str) -> str:
    return _normalize_output_text(machine_type)


def _display_task_detail(task: ScheduledTask) -> str:
    return _normalize_output_text(task.detail)


def _display_labor_group(labor_group: str) -> str:
    return {
        "manufacturing": "調製",
        "bottling": "排瓶",
        "filling": "充填",
    }.get(labor_group, labor_group)


def _display_task_item(task_type: str) -> str:
    return {
        "Cat1": "調製項",
        "Cat2": "調製項",
        "Cat3": "排瓶項",
        "Cat4": "充填項",
    }.get(task_type, task_type)


def _display_task_stage(task_type: str) -> str:
    return {
        "Cat1": "預攪",
        "Cat2": "攪料清洗",
        "Cat3": "排瓶",
        "Cat4": "充填",
    }.get(task_type, task_type)


def _display_labor_type(task_type: str) -> str:
    return "調製類型" if task_type in {"Cat1", "Cat2"} else "充填類型"


def _display_numeric(value: float | int) -> int | float:
    if isinstance(value, int):
        return value
    rounded_int = round(value)
    if abs(value - rounded_int) < 1e-6:
        return int(rounded_int)
    return round(value, 2)


def _display_integer(value: float | int) -> int:
    return int(round(float(value)))


def _infer_actual_assigned_people(task: ScheduledTask) -> int:
    if abs(task.machine_minutes) <= 1e-6:
        return 0
    return max(1, int(round(task.labor_minutes / task.machine_minutes)))


def _group_output_tasks(tasks: list[ScheduledTask]) -> list[ScheduledTask]:
    grouped: dict[tuple[str, str, str, date, str, str], ScheduledTask] = {}
    ordered_keys: list[tuple[str, str, str, date, str, str]] = []
    for task in tasks:
        normalized_detail = _display_task_detail(task)
        key = (
            task.order_id,
            task.task_type,
            normalized_detail,
            task.scheduled_date,
            task.machine_type,
            task.labor_group,
        )
        existing = grouped.get(key)
        if existing is None:
            grouped[key] = ScheduledTask(
                order_id=task.order_id,
                task_type=task.task_type,
                detail=normalized_detail,
                scheduled_date=task.scheduled_date,
                machine_type=task.machine_type,
                machine_minutes=task.machine_minutes,
                labor_group=task.labor_group,
                labor_minutes=task.labor_minutes,
            )
            ordered_keys.append(key)
            continue
        grouped[key] = ScheduledTask(
            order_id=existing.order_id,
            task_type=existing.task_type,
            detail=existing.detail,
            scheduled_date=existing.scheduled_date,
            machine_type=existing.machine_type,
            machine_minutes=existing.machine_minutes + task.machine_minutes,
            labor_group=existing.labor_group,
            labor_minutes=existing.labor_minutes + task.labor_minutes,
        )
    return [grouped[key] for key in ordered_keys]


def _english_machine_label(machine_type: str) -> str:
    label = str(machine_type or "")
    for zh, en in MACHINE_LABEL_TRANSLATIONS.items():
        label = label.replace(zh, en)
    return label


def _english_task_detail(task: ScheduledTask) -> str:
    detail = str(task.detail or "")
    detail = detail.replace("第", "Batch ")
    detail = detail.replace("批預攪", " Premix")
    detail = detail.replace("批攪料清洗", " Mix+Clean")
    detail = detail.replace("-排瓶", " - Bottling Prep")
    detail = detail.replace("-充填", " - Filling")
    detail = detail.replace("預攪", "Premix")
    detail = detail.replace("攪料清洗", "Mix+Clean")
    return detail


def _format_chinese_date(current_day: date) -> str:
    return f"{current_day.year}年{current_day.month}月{current_day.day}日"


def _normalize_manufacturing_detail(detail: str) -> str:
    return re.sub(r"-第\d+批(預攪|攪料清洗)$", "", detail)


def _normalized_manufacturing_group_key(detail: str) -> str:
    normalized = _normalize_output_text(detail)
    return re.sub(r"-第\d+批(預攪|攪料清洗)$", "", normalized)


def write_reference_format_output(schedule: ScheduleResult, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    summary = build_schedule_summary(schedule)
    grouped_tasks = _group_output_tasks(schedule.task_records)
    grouped_tasks = _group_output_tasks(schedule.task_records)
    grouped_tasks = _group_output_tasks(schedule.task_records)

    manufacturing_group_totals: dict[tuple[str, str], float] = defaultdict(float)
    packaging_group_totals: dict[tuple[str, str], float] = defaultdict(float)
    manufacturing_group_ids: dict[tuple[str, str], int] = {}
    machine_day_slot_cursor: dict[tuple[date, str], int] = defaultdict(int)

    for task in grouped_tasks:
        if task.task_type in {"Cat1", "Cat2"}:
            group_key = (task.order_id, _normalized_manufacturing_group_key(task.detail))
            manufacturing_group_totals[group_key] += 1.0
            if group_key not in manufacturing_group_ids:
                manufacturing_group_ids[group_key] = len(manufacturing_group_ids) + 1
        else:
            packaging_group_totals[(task.order_id, task.task_type)] += task.machine_minutes / 5.0

    reference_path = output_dir / "best_schedule_reference_format.csv"
    with reference_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "訂單",
                "子任務ID",
                "Cat",
                "Seq",
                "項目",
                "工序",
                "人力類型",
                "機器類型",
                "表定產能人數",
                "需求總單位",
                "當日安排單位",
                "單位時間(min)",
                "日期",
                "Slot 範圍",
                "Slots 數",
                "類型",
                "總時間(min)",
                "Need Unit",
                "最早Slot索引",
                "最晚Slot索引",
            ]
        )

        for task in grouped_tasks:
            slot_count = max(1, math.ceil(task.machine_minutes / 5.0))
            start_slot = machine_day_slot_cursor[(task.scheduled_date, task.machine_type)]
            end_slot = start_slot + slot_count - 1
            machine_day_slot_cursor[(task.scheduled_date, task.machine_type)] = end_slot + 1

            if task.task_type == "Cat1":
                group_key = (task.order_id, _normalize_manufacturing_detail(task.detail))
                writer.writerow(
                    [
                        task.order_id,
                        f"{task.order_id}_{manufacturing_group_ids[group_key]}",
                        "Cat1",
                        "10",
                        "製造項",
                        "預攪",
                        "製造類型",
                        task.machine_type,
                        "1",
                        int(manufacturing_group_totals[group_key]),
                        "1",
                        _display_numeric(task.machine_minutes),
                        _format_chinese_date(task.scheduled_date),
                        f"{_slot_index_to_time(start_slot)}-{_slot_index_to_time(end_slot)}",
                        slot_count,
                        "正常",
                        _display_numeric(task.machine_minutes),
                        int(manufacturing_group_totals[group_key]),
                        start_slot,
                        end_slot,
                    ]
                )
            elif task.task_type == "Cat2":
                group_key = (task.order_id, _normalize_manufacturing_detail(task.detail))
                writer.writerow(
                    [
                        task.order_id,
                        f"{task.order_id}_{manufacturing_group_ids[group_key]}",
                        "Cat2",
                        "12",
                        "製造項",
                        "攪料+清洗",
                        "製造類型",
                        task.machine_type,
                        "1",
                        int(manufacturing_group_totals[group_key]),
                        "1",
                        _display_numeric(task.machine_minutes),
                        _format_chinese_date(task.scheduled_date),
                        f"{_slot_index_to_time(start_slot)}-{_slot_index_to_time(end_slot)}",
                        slot_count,
                        "正常",
                        _display_numeric(task.machine_minutes),
                        int(manufacturing_group_totals[group_key]),
                        start_slot,
                        end_slot,
                    ]
                )
            elif task.task_type == "Cat3":
                demand_units = packaging_group_totals[(task.order_id, "Cat3")]
                arranged_units = task.machine_minutes / 5.0
                writer.writerow(
                    [
                        task.order_id,
                        f"{task.order_id}_C3_18",
                        "Cat3",
                        "18",
                        "排瓶項",
                        "排瓶",
                        "充填類型",
                        task.machine_type,
                        "1",
                        _display_numeric(demand_units),
                        _display_numeric(arranged_units),
                        5,
                        _format_chinese_date(task.scheduled_date),
                        f"{_slot_index_to_time(start_slot)}-{_slot_index_to_time(end_slot)}",
                        slot_count,
                        "正常",
                        _display_numeric(task.machine_minutes),
                        _display_numeric(demand_units),
                        start_slot,
                        end_slot,
                    ]
                )
            elif task.task_type == "Cat4":
                demand_units = packaging_group_totals[(task.order_id, "Cat4")]
                arranged_units = task.machine_minutes / 5.0
                writer.writerow(
                    [
                        task.order_id,
                        f"{task.order_id}_C4_19",
                        "Cat4",
                        "19",
                        "充填項",
                        "充填",
                        "充填類型",
                        task.machine_type,
                        "4",
                        _display_numeric(demand_units),
                        _display_numeric(arranged_units),
                        5,
                        _format_chinese_date(task.scheduled_date),
                        f"{_slot_index_to_time(start_slot)}-{_slot_index_to_time(end_slot)}",
                        slot_count,
                        "正常",
                        _display_numeric(task.machine_minutes),
                        _display_numeric(demand_units),
                        start_slot,
                        end_slot,
                    ]
                )
        writer.writerow(
            [
                "TOTAL",
                "SUMMARY",
                "",
                "",
                "摘要",
                (
                    f"總加班={_display_numeric(summary['total_overtime_minutes'])}; "
                    f"排瓶加班={_display_numeric(summary['bottling_overtime_minutes'])}; "
                    f"充填加班={_display_numeric(summary['filling_overtime_minutes'])}; "
                    f"總排瓶={_display_numeric(summary['total_bottling_minutes'])}; "
                    f"總充填={_display_numeric(summary['total_filling_minutes'])}"
                ),
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ]
        )


def write_schedule_outputs(schedule: ScheduleResult, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    summary = build_schedule_summary(schedule)
    grouped_tasks = _group_output_tasks(schedule.task_records)

    task_path = output_dir / "best_schedule_tasks.csv"
    with task_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "訂單",
                "類別",
                "明細",
                "日期",
                "機器類型",
                "機器分鐘",
                "人力類型",
                "人力分鐘",
            ]
        )
        for task in grouped_tasks:
            writer.writerow(
                [
                    task.order_id,
                    _display_task_stage(task.task_type),
                    _display_task_detail(task),
                    task.scheduled_date.isoformat(),
                    _display_machine_label(task.machine_type),
                    _display_integer(task.machine_minutes),
                    _display_labor_group(task.labor_group),
                    _display_integer(task.labor_minutes),
                ]
            )
        writer.writerow(
            [
                "總計",
                "摘要",
                (
                    f"總加班分鐘數={_display_numeric(summary['total_overtime_minutes'])}; "
                    f"調製超時分鐘數={_display_numeric(summary['manufacturing_overtime_minutes'])}; "
                    f"排瓶超時分鐘數={_display_numeric(summary['bottling_overtime_minutes'])}; "
                    f"充填超時分鐘數={_display_numeric(summary['filling_overtime_minutes'])}; "
                    f"總排瓶分鐘數={_display_numeric(summary['total_bottling_minutes'])}; "
                    f"總充填分鐘數={_display_numeric(summary['total_filling_minutes'])}"
                ),
                "",
                "",
                "",
                "",
                "",
            ]
        )

    daily_path = output_dir / "best_schedule_daily_load.csv"
    with daily_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["日期", "人力類型", "已用人力分鐘", "正常人力分鐘", "超時分鐘"])
        for (labor_group, current_day), used_minutes in sorted(schedule.daily_labor_usage.items()):
            if current_day == date.min:
                continue
            regular_minutes = LABOR_CONFIG[labor_group]["regular_minutes"]
            writer.writerow(
                [
                    current_day.isoformat(),
                    _display_labor_group(labor_group),
                    _display_numeric(used_minutes),
                    regular_minutes,
                    _display_numeric(max(used_minutes - regular_minutes, 0.0)),
                ]
            )

    summary_path = output_dir / "schedule_summary.csv"
    with summary_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "總加班分鐘數",
                "調製超時分鐘數",
                "排瓶超時分鐘數",
                "充填超時分鐘數",
                "週末加班分鐘數",
                "總排瓶分鐘數",
                "總充填分鐘數",
            ]
        )
        writer.writerow(
            [
                _display_numeric(summary["total_overtime_minutes"]),
                _display_numeric(summary["manufacturing_overtime_minutes"]),
                _display_numeric(summary["bottling_overtime_minutes"]),
                _display_numeric(summary["filling_overtime_minutes"]),
                _display_numeric(summary["weekend_overtime_minutes"]),
                _display_numeric(summary["total_bottling_minutes"]),
                _display_numeric(summary["total_filling_minutes"]),
            ]
        )


def write_daily_manufacturing_schedule_output(
    schedule: ScheduleResult,
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    manufacturing_tasks = [task for task in schedule.task_records if task.task_type in {"Cat1", "Cat2"}]
    machine_day_totals: dict[tuple[date, str], float] = defaultdict(float)
    order_day_totals: dict[tuple[date, str], float] = defaultdict(float)
    labor_day_totals: dict[date, float] = defaultdict(float)

    for task in manufacturing_tasks:
        machine_day_totals[(task.scheduled_date, task.machine_type)] += task.machine_minutes
        order_day_totals[(task.scheduled_date, task.order_id)] += task.machine_minutes
        labor_day_totals[task.scheduled_date] += task.labor_minutes

    path = output_dir / "daily_manufacturing_schedule.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "日期",
                "訂單",
                "類別",
                "明細",
                "機器類型",
                "機器分鐘",
                "人力分鐘",
                "當日機器總分鐘",
                "當日訂單總分鐘",
                "當日調製人力總分鐘",
            ]
        )
        for task in sorted(
            manufacturing_tasks,
            key=lambda item: (item.scheduled_date, item.machine_type, item.order_id, item.task_type, item.detail),
        ):
            writer.writerow(
                [
                    task.scheduled_date.isoformat(),
                    task.order_id,
                    _display_task_stage(task.task_type),
                    _display_task_detail(task),
                    _display_machine_label(task.machine_type),
                    _display_numeric(task.machine_minutes),
                    _display_numeric(task.labor_minutes),
                    _display_numeric(machine_day_totals[(task.scheduled_date, task.machine_type)]),
                    _display_numeric(order_day_totals[(task.scheduled_date, task.order_id)]),
                    _display_numeric(labor_day_totals[task.scheduled_date]),
                ]
            )


def write_daily_packaging_schedule_output(
    schedule: ScheduleResult,
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    packaging_tasks = _group_output_tasks(
        [task for task in schedule.task_records if task.task_type in {"Cat3", "Cat4"}]
    )
    process_day_totals: dict[tuple[date, str], float] = defaultdict(float)
    machine_day_totals: dict[tuple[date, str], float] = defaultdict(float)
    order_day_totals: dict[tuple[date, str], float] = defaultdict(float)
    labor_day_totals: dict[tuple[date, str], float] = defaultdict(float)

    for task in packaging_tasks:
        process = "排瓶" if task.task_type == "Cat3" else "充填"
        process_day_totals[(task.scheduled_date, process)] += task.machine_minutes
        machine_day_totals[(task.scheduled_date, task.machine_type)] += task.machine_minutes
        order_day_totals[(task.scheduled_date, task.order_id)] += task.machine_minutes
        labor_day_totals[(task.scheduled_date, process)] += task.labor_minutes

    path = output_dir / "daily_packaging_schedule.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "日期",
                "訂單",
                "類別",
                "工序",
                "明細",
                "機器類型",
                "機器分鐘",
                "人力分鐘",
                "當日工序總分鐘",
                "當日機器總分鐘",
                "當日訂單總分鐘",
                "當日工序人力總分鐘",
                "當日正常分鐘",
                "當日超時分鐘",
            ]
        )
        for task in sorted(
            packaging_tasks,
            key=lambda item: (item.scheduled_date, item.task_type, item.machine_type, item.order_id, item.detail),
        ):
            process = "排瓶" if task.task_type == "Cat3" else "充填"
            process_day_total = process_day_totals[(task.scheduled_date, process)]
            process_day_total_labor = labor_day_totals[(task.scheduled_date, process)]
            regular_minutes = PROCESS_OVERTIME_REGULAR_MINUTES.get(
                "bottling" if process == "排瓶" else "filling",
                PROCESS_DAILY_REGULAR_MINUTES,
            )
            overtime_base_minutes = process_day_total_labor if process == "排瓶" else process_day_total
            writer.writerow(
                [
                    task.scheduled_date.isoformat(),
                    task.order_id,
                    _display_task_stage(task.task_type),
                    process,
                    _display_task_detail(task),
                    _display_machine_label(task.machine_type),
                    _display_integer(task.machine_minutes),
                    _display_integer(task.labor_minutes),
                    _display_integer(process_day_total),
                    _display_integer(machine_day_totals[(task.scheduled_date, task.machine_type)]),
                    _display_integer(order_day_totals[(task.scheduled_date, task.order_id)]),
                    _display_integer(process_day_total_labor),
                    regular_minutes,
                    _display_integer(max(overtime_base_minutes - regular_minutes, 0.0)),
                ]
            )


def write_manufacturing_cleaning_audit_output(
    schedule: ScheduleResult,
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    machine_day_orders: dict[tuple[date, str], set[str]] = defaultdict(set)
    machine_day_details: dict[tuple[date, str], set[str]] = defaultdict(set)
    machine_day_minutes: dict[tuple[date, str], float] = defaultdict(float)

    for task in schedule.task_records:
        if task.task_type not in {"Cat1", "Cat2"}:
            continue
        key = (task.scheduled_date, task.machine_type)
        machine_day_orders[key].add(task.order_id)
        machine_day_details[key].add(_normalized_manufacturing_group_key(task.detail))
        machine_day_minutes[key] += task.machine_minutes

    path = output_dir / "manufacturing_cleaning_audit.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "日期",
                "機器類型",
                "當日機器總分鐘",
                "當日訂單數",
                "當日不同製程數",
                "推估換產品清洗次數",
                "訂單清單",
                "製程明細",
            ]
        )
        for (scheduled_date, machine_type), total_minutes in sorted(machine_day_minutes.items()):
            distinct_orders = sorted(machine_day_orders[(scheduled_date, machine_type)])
            distinct_details = sorted(machine_day_details[(scheduled_date, machine_type)])
            distinct_process_count = len(distinct_details)
            writer.writerow(
                [
                    scheduled_date.isoformat(),
                    _display_machine_label(machine_type),
                    _display_numeric(total_minutes),
                    len(distinct_orders),
                    distinct_process_count,
                    max(distinct_process_count - 1, 0),
                    " | ".join(distinct_orders),
                    " | ".join(_normalize_output_text(detail) for detail in distinct_details),
                ]
            )


def write_reference_format_output(schedule: ScheduleResult, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    summary = build_schedule_summary(schedule)

    manufacturing_group_totals: dict[tuple[str, str], float] = defaultdict(float)
    packaging_group_totals: dict[tuple[str, str], float] = defaultdict(float)
    manufacturing_group_ids: dict[tuple[str, str], int] = {}
    machine_day_slot_cursor: dict[tuple[date, str], int] = defaultdict(int)

    for task in schedule.task_records:
        if task.task_type in {"Cat1", "Cat2"}:
            group_key = (task.order_id, _normalized_manufacturing_group_key(task.detail))
            manufacturing_group_totals[group_key] += 1.0
            if group_key not in manufacturing_group_ids:
                manufacturing_group_ids[group_key] = len(manufacturing_group_ids) + 1
        else:
            packaging_group_totals[(task.order_id, task.task_type)] += task.machine_minutes / 5.0

    reference_path = output_dir / "best_schedule_reference_format.csv"
    with reference_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "訂單",
                "子任務ID",
                "Cat",
                "Seq",
                "項目",
                "工序",
                "人力類型",
                "機器類型",
                "表定產能人數",
                "需求總單位",
                "當日安排單位",
                "單位時間(min)",
                "日期",
                "Slot 範圍",
                "Slots 數",
                "類型",
                "總時間(min)",
                "需求單位",
                "最早Slot索引",
                "最晚Slot索引",
            ]
        )

        for task in grouped_tasks:
            slot_count = max(1, math.ceil(task.machine_minutes / 5.0))
            start_slot = machine_day_slot_cursor[(task.scheduled_date, task.machine_type)]
            end_slot = start_slot + slot_count - 1
            machine_day_slot_cursor[(task.scheduled_date, task.machine_type)] = end_slot + 1

            if task.task_type in {"Cat1", "Cat2"}:
                group_key = (task.order_id, _normalized_manufacturing_group_key(task.detail))
                seq = "10" if task.task_type == "Cat1" else "12"
                arranged_units = 1
                demand_units = int(manufacturing_group_totals[group_key])
                planned_people = "1"
            else:
                demand_units = round(packaging_group_totals[(task.order_id, task.task_type)], 2)
                arranged_units = round(task.machine_minutes / 5.0, 2)
                seq = "18" if task.task_type == "Cat3" else "19"
                planned_people = "1" if task.task_type == "Cat3" else "4"

            writer.writerow(
                [
                    task.order_id,
                    (
                        f"{task.order_id}_{manufacturing_group_ids[(task.order_id, _normalized_manufacturing_group_key(task.detail))]}"
                        if task.task_type in {"Cat1", "Cat2"}
                        else f"{task.order_id}_{task.task_type}_{seq}"
                    ),
                    task.task_type,
                    seq,
                    _display_task_item(task.task_type),
                    _display_task_stage(task.task_type),
                    _display_labor_type(task.task_type),
                    _display_machine_label(task.machine_type),
                    planned_people,
                    demand_units,
                    arranged_units,
                    _display_numeric(task.machine_minutes if task.task_type in {"Cat1", "Cat2"} else 5),
                    task.scheduled_date.isoformat(),
                    f"{_slot_index_to_time(start_slot)}-{_slot_index_to_time(end_slot)}",
                    slot_count,
                    "正常",
                    _display_numeric(task.machine_minutes),
                    demand_units,
                    start_slot,
                    end_slot,
                ]
            )

        writer.writerow(
            [
                "總計",
                "摘要",
                "",
                "",
                "摘要",
                (
                    f"總加班分鐘數={_display_numeric(summary['total_overtime_minutes'])}; "
                    f"調製超時分鐘數={_display_numeric(summary['manufacturing_overtime_minutes'])}; "
                    f"排瓶超時分鐘數={_display_numeric(summary['bottling_overtime_minutes'])}; "
                    f"充填超時分鐘數={_display_numeric(summary['filling_overtime_minutes'])}; "
                    f"總排瓶分鐘數={_display_numeric(summary['total_bottling_minutes'])}; "
                    f"總充填分鐘數={_display_numeric(summary['total_filling_minutes'])}"
                ),
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ]
        )


def write_reference_format_output(schedule: ScheduleResult, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    summary = build_schedule_summary(schedule)

    manufacturing_group_totals: dict[tuple[str, str], float] = defaultdict(float)
    packaging_group_totals: dict[tuple[str, str], float] = defaultdict(float)
    manufacturing_group_ids: dict[tuple[str, str], int] = {}
    machine_day_slot_cursor: dict[tuple[date, str], int] = defaultdict(int)

    for task in schedule.task_records:
        if task.task_type in {"Cat1", "Cat2"}:
            group_key = (task.order_id, _normalized_manufacturing_group_key(task.detail))
            manufacturing_group_totals[group_key] += 1.0
            if group_key not in manufacturing_group_ids:
                manufacturing_group_ids[group_key] = len(manufacturing_group_ids) + 1
        else:
            packaging_group_totals[(task.order_id, task.task_type)] += task.machine_minutes / 5.0

    reference_path = output_dir / "best_schedule_reference_format.csv"
    with reference_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "訂單編號",
                "子任務編號",
                "工序代碼",
                "工序順序",
                "作業項目",
                "作業工序",
                "人力別",
                "機台別",
                "標準配置人數",
                "總需求量",
                "當日安排量",
                "每單位時間(min)",
                "排程日期",
                "時段範圍",
                "時段數",
                "排程類型",
                "作業總時間(min)",
                "需求單位數",
                "最早時段索引",
                "最晚時段索引",
            ]
        )

        for task in grouped_tasks:
            slot_count = max(1, math.ceil(task.machine_minutes / 5.0))
            start_slot = machine_day_slot_cursor[(task.scheduled_date, task.machine_type)]
            end_slot = start_slot + slot_count - 1
            machine_day_slot_cursor[(task.scheduled_date, task.machine_type)] = end_slot + 1

            if task.task_type in {"Cat1", "Cat2"}:
                group_key = (task.order_id, _normalized_manufacturing_group_key(task.detail))
                seq = "10" if task.task_type == "Cat1" else "12"
                arranged_units = 1
                demand_units = int(manufacturing_group_totals[group_key])
                planned_people = "1"
                subtask_id = f"{task.order_id}_{manufacturing_group_ids[group_key]}"
            else:
                demand_units = _display_integer(packaging_group_totals[(task.order_id, task.task_type)])
                arranged_units = _display_integer(task.machine_minutes / 5.0)
                seq = "18" if task.task_type == "Cat3" else "19"
                planned_people = "1" if task.task_type == "Cat3" else "4"
                subtask_id = f"{task.order_id}_{task.task_type}_{seq}"

            writer.writerow(
                [
                    task.order_id,
                    subtask_id,
                    task.task_type,
                    seq,
                    _display_task_item(task.task_type),
                    _display_task_stage(task.task_type),
                    _display_labor_type(task.task_type),
                    _display_machine_label(task.machine_type),
                    planned_people,
                    demand_units,
                    arranged_units,
                    _display_integer(task.machine_minutes if task.task_type in {"Cat1", "Cat2"} else 5),
                    task.scheduled_date.isoformat(),
                    f"{_slot_index_to_time(start_slot)}-{_slot_index_to_time(end_slot)}",
                    slot_count,
                    "正常",
                    _display_integer(task.machine_minutes),
                    _display_integer(demand_units),
                    start_slot,
                    end_slot,
                ]
            )

        writer.writerow(
            [
                "總計",
                "摘要",
                "",
                "",
                "摘要",
                (
                    f"總加班分鐘數={_display_numeric(summary['total_overtime_minutes'])}; "
                    f"調製超時分鐘數={_display_numeric(summary['manufacturing_overtime_minutes'])}; "
                    f"排瓶超時分鐘數={_display_numeric(summary['bottling_overtime_minutes'])}; "
                    f"充填超時分鐘數={_display_numeric(summary['filling_overtime_minutes'])}; "
                    f"總排瓶分鐘數={_display_numeric(summary['total_bottling_minutes'])}; "
                    f"總充填分鐘數={_display_numeric(summary['total_filling_minutes'])}"
                ),
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ]
        )


def write_schedule_outputs(schedule: ScheduleResult, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    summary = build_schedule_summary(schedule)
    grouped_tasks = _group_output_tasks(schedule.task_records)

    task_path = output_dir / "best_schedule_tasks.csv"
    with task_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["訂單", "類別", "明細", "日期", "機器類型", "機器分鐘", "人力類型", "人力分鐘"])
        for task in grouped_tasks:
            writer.writerow(
                [
                    task.order_id,
                    _display_task_stage(task.task_type),
                    _display_task_detail(task),
                    task.scheduled_date.isoformat(),
                    _display_machine_label(task.machine_type),
                    _display_integer(task.machine_minutes),
                    _display_labor_group(task.labor_group),
                    _display_integer(task.labor_minutes),
                ]
            )
        writer.writerow(
            [
                "總計",
                "摘要",
                (
                    f"總加班分鐘數={_display_integer(summary['total_overtime_minutes'])}; "
                    f"調製超時分鐘數={_display_integer(summary['manufacturing_overtime_minutes'])}; "
                    f"排瓶超時分鐘數={_display_integer(summary['bottling_overtime_minutes'])}; "
                    f"充填超時分鐘數={_display_integer(summary['filling_overtime_minutes'])}; "
                    f"總排瓶分鐘數={_display_integer(summary['total_bottling_minutes'])}; "
                    f"總充填分鐘數={_display_integer(summary['total_filling_minutes'])}"
                ),
                "",
                "",
                "",
                "",
                "",
            ]
        )

    daily_path = output_dir / "best_schedule_daily_load.csv"
    with daily_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["日期", "人力類型", "已用人力分鐘", "正常人力分鐘", "超時分鐘"])
        daily_base_usage: dict[tuple[str, date], float] = defaultdict(float)
        for task in schedule.task_records:
            if task.task_type in {"Cat1", "Cat2"}:
                daily_base_usage[("manufacturing", task.scheduled_date)] += task.labor_minutes
            elif task.task_type == "Cat3":
                daily_base_usage[("bottling", task.scheduled_date)] += task.labor_minutes
            elif task.task_type == "Cat4":
                daily_base_usage[("filling", task.scheduled_date)] += task.machine_minutes

        regular_by_group = {
            "manufacturing": LABOR_CONFIG["manufacturing"]["regular_minutes"],
            "bottling": PROCESS_OVERTIME_REGULAR_MINUTES["bottling"],
            "filling": PROCESS_OVERTIME_REGULAR_MINUTES["filling"],
        }
        for (labor_group, current_day), used_minutes in sorted(daily_base_usage.items()):
            if current_day == date.min or used_minutes <= 0:
                continue
            regular_minutes = regular_by_group[labor_group]
            writer.writerow(
                [
                    current_day.isoformat(),
                    _display_labor_group(labor_group),
                    _display_integer(used_minutes),
                    regular_minutes,
                    _display_integer(max(used_minutes - regular_minutes, 0.0)),
                ]
            )

    summary_path = output_dir / "schedule_summary.csv"
    with summary_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "總加班分鐘數",
                "調製超時分鐘數",
                "排瓶超時分鐘數",
                "充填超時分鐘數",
                "週末加班分鐘數",
                "總排瓶分鐘數",
                "總充填分鐘數",
            ]
        )
        writer.writerow(
            [
                _display_integer(summary["total_overtime_minutes"]),
                _display_integer(summary["manufacturing_overtime_minutes"]),
                _display_integer(summary["bottling_overtime_minutes"]),
                _display_integer(summary["filling_overtime_minutes"]),
                _display_integer(summary["weekend_overtime_minutes"]),
                _display_integer(summary["total_bottling_minutes"]),
                _display_integer(summary["total_filling_minutes"]),
            ]
        )


def write_reference_format_output(
    schedule: ScheduleResult,
    orders: list[Order],
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    summary = build_schedule_summary(schedule)
    grouped_tasks = _group_output_tasks(schedule.task_records)
    order_map = {order.order_id: order for order in orders}

    manufacturing_group_totals: dict[tuple[str, str], int] = defaultdict(int)
    packaging_group_totals: dict[tuple[str, str], float] = defaultdict(float)
    manufacturing_group_ids: dict[tuple[str, str], int] = {}
    machine_day_slot_cursor: dict[tuple[date, str], int] = defaultdict(int)
    order_day_bottling_next_slot: dict[tuple[str, date], int] = defaultdict(int)

    for task in grouped_tasks:
        if task.task_type in {"Cat1", "Cat2"}:
            group_key = (task.order_id, _normalized_manufacturing_group_key(task.detail))
            manufacturing_group_totals[group_key] += 1
            if group_key not in manufacturing_group_ids:
                manufacturing_group_ids[group_key] = len(manufacturing_group_ids) + 1
        else:
            packaging_group_totals[(task.order_id, task.task_type)] += task.machine_minutes / 5.0

    reference_path = output_dir / "best_schedule_reference_format.csv"
    with reference_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "訂單編號",
                "子任務編號",
                "工序代碼",
                "工序順序",
                "作業項目",
                "作業工序",
                "人力別",
                "機台別",
                "標準配置人數",
                "實際安排人數",
                "總需求量",
                "當日安排量",
                "每單位時間(min)",
                "排程日期",
                "時段範圍",
                "時段數",
                "排程類型",
                "作業總時間(min)",
                "需求單位數",
                "最早時段索引",
                "最晚時段索引",
            ]
        )

        task_priority = {"Cat1": 1, "Cat2": 2, "Cat3": 3, "Cat4": 4}
        for task in sorted(
            grouped_tasks,
            key=lambda item: (
                item.scheduled_date,
                task_priority.get(item.task_type, 99),
                item.order_id,
                item.machine_type,
                item.detail,
            ),
        ):
            slot_count = max(1, math.ceil(task.machine_minutes / 5.0))
            dependency_slot = 0
            if task.task_type == "Cat4":
                dependency_slot = order_day_bottling_next_slot[(task.order_id, task.scheduled_date)]
            start_slot = max(
                machine_day_slot_cursor[(task.scheduled_date, task.machine_type)],
                dependency_slot,
            )
            end_slot = start_slot + slot_count - 1
            machine_day_slot_cursor[(task.scheduled_date, task.machine_type)] = end_slot + 1

            if task.task_type in {"Cat1", "Cat2"}:
                group_key = (task.order_id, _normalized_manufacturing_group_key(task.detail))
                seq = "10" if task.task_type == "Cat1" else "12"
                demand_units = manufacturing_group_totals[group_key]
                arranged_units = 1
                standard_people = 1
                actual_people = 1
                subtask_id = f"{task.order_id}_{manufacturing_group_ids[group_key]}"
                unit_minutes = _display_integer(task.machine_minutes)
            else:
                seq = "18" if task.task_type == "Cat3" else "19"
                demand_units = _display_integer(packaging_group_totals[(task.order_id, task.task_type)])
                arranged_units = _display_integer(task.machine_minutes / 5.0)
                source_order = order_map.get(task.order_id)
                if task.task_type == "Cat3":
                    standard_people = (
                        source_order.filling_item.bottling_people
                        if source_order is not None
                        else _infer_actual_assigned_people(task)
                    )
                else:
                    standard_people = (
                        source_order.filling_item.filling_people
                        if source_order is not None
                        else _infer_actual_assigned_people(task)
                    )
                actual_people = _infer_actual_assigned_people(task)
                subtask_id = f"{task.order_id}_{task.task_type}_{seq}"
                unit_minutes = 5
                if task.task_type == "Cat3":
                    order_day_bottling_next_slot[(task.order_id, task.scheduled_date)] = end_slot + 1

            writer.writerow(
                [
                    task.order_id,
                    subtask_id,
                    task.task_type,
                    seq,
                    _display_task_item(task.task_type),
                    _display_task_stage(task.task_type),
                    _display_labor_type(task.task_type),
                    _display_machine_label(task.machine_type),
                    standard_people,
                    actual_people,
                    _display_integer(demand_units),
                    _display_integer(arranged_units),
                    unit_minutes,
                    task.scheduled_date.isoformat(),
                    f"{_slot_index_to_time(start_slot)}-{_slot_index_to_time(end_slot)}",
                    slot_count,
                    "正常",
                    _display_integer(task.machine_minutes),
                    _display_integer(demand_units),
                    start_slot,
                    end_slot,
                ]
            )

        writer.writerow(
            [
                "總計",
                "摘要",
                "",
                "",
                "摘要",
                (
                    f"總加班分鐘數={_display_integer(summary['total_overtime_minutes'])}; "
                    f"調製超時分鐘數={_display_integer(summary['manufacturing_overtime_minutes'])}; "
                    f"排瓶超時分鐘數={_display_integer(summary['bottling_overtime_minutes'])}; "
                    f"充填超時分鐘數={_display_integer(summary['filling_overtime_minutes'])}; "
                    f"總排瓶分鐘數={_display_integer(summary['total_bottling_minutes'])}; "
                    f"總充填分鐘數={_display_integer(summary['total_filling_minutes'])}"
                ),
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ]
        )


def write_order_anomalies_output(
    orders: list[Order],
    anomalies: list[OrderAnomaly],
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    order_map = {order.order_id: order for order in orders}
    anomaly_path = output_dir / "order_anomalies.csv"
    with anomaly_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "order_id",
                "product_name",
                "severity",
                "anomaly_type",
                "exclude_from_optimization",
                "message",
                "due_date",
                "raw_ready_date",
                "material_ready_date",
                "manufacturing_batch_count",
                "filling_minutes",
                "bottling_minutes",
            ]
        )
        for anomaly in anomalies:
            order = order_map[anomaly.order_id]
            writer.writerow(
                [
                    anomaly.order_id,
                    anomaly.product_name,
                    anomaly.severity,
                    anomaly.anomaly_type,
                    "Y" if anomaly.exclude_from_optimization else "N",
                    anomaly.message,
                    order.due_date.isoformat(),
                    order.raw_ready_date.isoformat(),
                    order.material_ready_date.isoformat(),
                    sum(item.batches for item in order.manufacturing_items),
                    round(order.filling_item.filling_machine_minutes, 2),
                    round(order.filling_item.bottling_machine_minutes, 2),
                ]
            )

    grouped_paths = {
        "exclude": output_dir / "order_anomalies_exclude.csv",
        "warning": output_dir / "order_anomalies_warning.csv",
    }
    for severity, path in grouped_paths.items():
        with path.open("w", newline="", encoding="utf-8-sig") as file:
            writer = csv.writer(file)
            writer.writerow(
                [
                    "order_id",
                    "product_name",
                    "severity",
                    "anomaly_type",
                    "exclude_from_optimization",
                    "message",
                    "due_date",
                    "raw_ready_date",
                    "material_ready_date",
                    "manufacturing_batch_count",
                    "filling_minutes",
                    "bottling_minutes",
                ]
            )
            for anomaly in anomalies:
                if anomaly.severity != severity:
                    continue
                order = order_map[anomaly.order_id]
                writer.writerow(
                    [
                        anomaly.order_id,
                        anomaly.product_name,
                        anomaly.severity,
                        anomaly.anomaly_type,
                        "Y" if anomaly.exclude_from_optimization else "N",
                        anomaly.message,
                        order.due_date.isoformat(),
                        order.raw_ready_date.isoformat(),
                        order.material_ready_date.isoformat(),
                        sum(item.batches for item in order.manufacturing_items),
                        round(order.filling_item.filling_machine_minutes, 2),
                        round(order.filling_item.bottling_machine_minutes, 2),
                    ]
                )


def write_schedule_validation_output(
    issues: list[ScheduleValidationIssue],
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    validation_path = output_dir / "schedule_validation.csv"
    with validation_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["order_id", "severity", "issue_type", "message"])
        for issue in issues:
            writer.writerow([issue.order_id, issue.severity, issue.issue_type, issue.message])

    summary_path = output_dir / "schedule_validation_summary.csv"
    by_type: dict[str, int] = defaultdict(int)
    by_severity: dict[str, int] = defaultdict(int)
    for issue in issues:
        by_type[issue.issue_type] += 1
        by_severity[issue.severity] += 1

    with summary_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["metric", "value"])
        writer.writerow(["total_issues", len(issues)])
        writer.writerow(["error_count", by_severity.get("error", 0)])
        writer.writerow(["warning_count", by_severity.get("warning", 0)])
        for issue_type, count in sorted(by_type.items()):
            writer.writerow([f"issue_type:{issue_type}", count])


def write_source_link_output(
    source_links: list[dict[str, str]],
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / "order_source_links.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "p_order_id",
                "product_name",
                "time_ready_core_candidates",
                "due_serial",
                "quantity",
                "best_match_score",
                "match_count",
                "bom_match_count",
                "manufacturing_efficiency_match_count",
                "filling_efficiency_match_count",
                "source_order_ids",
                "source_product_codes",
                "source_core_codes",
            ]
        )
        for row in source_links:
            writer.writerow(
                [
                    row.get("p_order_id", ""),
                    row.get("product_name", ""),
                    row.get("time_ready_core_candidates", ""),
                    row.get("due_serial", ""),
                    row.get("quantity", ""),
                    row.get("best_match_score", ""),
                    row.get("match_count", ""),
                    row.get("bom_match_count", ""),
                    row.get("manufacturing_efficiency_match_count", ""),
                    row.get("filling_efficiency_match_count", ""),
                    row.get("source_order_ids", ""),
                    row.get("source_product_codes", ""),
                    row.get("source_core_codes", ""),
                ]
            )


def write_machine_counts_output(
    machine_counts: dict[str, int],
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / "machine_counts.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["machine_type", "count"])
        for machine_type, count in sorted(machine_counts.items()):
            writer.writerow([machine_type, count])


def write_order_model_audit_output(
    orders: list[Order],
    source_links: list[dict[str, str]],
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    source_link_map = {row.get("p_order_id", ""): row for row in source_links}
    path = output_dir / "order_model_audit.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "order_id",
                "product_name",
                "source_order_id",
                "source_product_code",
                "source_core_code",
                "due_date",
                "raw_ready_date",
                "material_ready_date",
                "quantity",
                "manufacturing_item_count",
                "manufacturing_processes",
                "manufacturing_machines",
                "manufacturing_batch_total",
                "manufacturing_premix_total_minutes",
                "manufacturing_mix_clean_total_minutes",
                "filling_machine",
                "bottling_machine",
                "filling_minutes",
                "bottling_minutes",
                "filling_people",
                "bottling_people",
                "source_match_count",
                "source_best_match_score",
                "bom_match_count",
                "manufacturing_efficiency_match_count",
                "filling_efficiency_match_count",
            ]
        )
        for order in orders:
            source_link = source_link_map.get(order.order_id, {})
            writer.writerow(
                [
                    order.order_id,
                    order.product_name,
                    order.source_order_id,
                    order.source_product_code,
                    order.source_core_code,
                    order.due_date.isoformat(),
                    order.raw_ready_date.isoformat(),
                    order.material_ready_date.isoformat(),
                    order.quantity,
                    len(order.manufacturing_items),
                    " | ".join(item.process_name for item in order.manufacturing_items),
                    " | ".join(item.mix_machine for item in order.manufacturing_items),
                    sum(item.batches for item in order.manufacturing_items),
                    round(sum(item.premix_minutes_per_batch * item.batches for item in order.manufacturing_items), 2),
                    round(sum(item.mix_clean_minutes_per_batch * item.batches for item in order.manufacturing_items), 2),
                    order.filling_item.filling_machine,
                    order.filling_item.bottling_machine,
                    round(order.filling_item.filling_machine_minutes, 2),
                    round(order.filling_item.bottling_machine_minutes, 2),
                    order.filling_item.filling_people,
                    order.filling_item.bottling_people,
                    source_link.get("match_count", ""),
                    source_link.get("best_match_score", ""),
                    source_link.get("bom_match_count", ""),
                    source_link.get("manufacturing_efficiency_match_count", ""),
                    source_link.get("filling_efficiency_match_count", ""),
                ]
            )


def detect_order_anomalies(orders: list[Order]) -> list[OrderAnomaly]:
    anomalies: list[OrderAnomaly] = []
    for order in orders:
        if order.raw_ready_date > order.due_date:
            anomalies.append(
                OrderAnomaly(
                    order_id=order.order_id,
                    product_name=order.product_name,
                    severity="exclude",
                    anomaly_type="raw_ready_after_due",
                    message=f"raw_ready_date {order.raw_ready_date} is later than due_date {order.due_date}",
                    exclude_from_optimization=True,
                )
            )
        if order.material_ready_date > order.due_date:
            anomalies.append(
                OrderAnomaly(
                    order_id=order.order_id,
                    product_name=order.product_name,
                    severity="exclude",
                    anomaly_type="material_ready_after_due",
                    message=f"material_ready_date {order.material_ready_date} is later than due_date {order.due_date}",
                    exclude_from_optimization=True,
                )
            )
        if (
            order.filling_item.bottling_machine_minutes > 0
            and order.material_ready_date + timedelta(days=1) > order.due_date
        ):
            anomalies.append(
                OrderAnomaly(
                    order_id=order.order_id,
                    product_name=order.product_name,
                    severity="exclude",
                    anomaly_type="no_bottling_day_before_due",
                    message=(
                        f"requires one bottling day before filling, but material_ready_date "
                        f"{order.material_ready_date} leaves no valid day before due_date {order.due_date}"
                    ),
                    exclude_from_optimization=True,
                )
            )

        manufacturing_window_days = max(
            (min(order.due_date, order.material_ready_date) - max(order.raw_ready_date, order.due_date - timedelta(days=30))).days,
            -1,
        )
        total_batches = sum(item.batches for item in order.manufacturing_items)
        if total_batches >= 8 and manufacturing_window_days <= 3:
            anomalies.append(
                OrderAnomaly(
                    order_id=order.order_id,
                    product_name=order.product_name,
                    severity="warning",
                    anomaly_type="tight_manufacturing_window",
                    message=(
                        f"manufacturing window is only {manufacturing_window_days + 1} day(s) for "
                        f"{total_batches} batch(es)"
                    ),
                    exclude_from_optimization=False,
                )
            )

    return anomalies


def split_schedulable_orders(
    orders: list[Order],
    exclude_warnings: bool = False,
) -> tuple[list[Order], list[OrderAnomaly]]:
    anomalies = detect_order_anomalies(orders)
    excluded_order_ids = {
        anomaly.order_id
        for anomaly in anomalies
        if anomaly.exclude_from_optimization or (exclude_warnings and anomaly.severity == "warning")
    }
    schedulable_orders = [order for order in orders if order.order_id not in excluded_order_ids]
    return schedulable_orders, anomalies


def load_seed_order_ids(schedule_csv_path: Path) -> tuple[str, ...]:
    order_first_seen: dict[str, tuple[str, int, int, str]] = {}
    task_rank = {"Cat1": 1, "Cat2": 2, "Cat3": 3, "Cat4": 4}
    with schedule_csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row_index, row in enumerate(reader):
            order_id = (row.get("order_id") or "").strip()
            if not order_id or order_id == "TOTAL":
                continue
            scheduled_date = (row.get("scheduled_date") or "").strip()
            task_type = (row.get("task_type") or "").strip()
            sort_key = (
                scheduled_date or "9999-12-31",
                task_rank.get(task_type, 99),
                row_index,
                order_id,
            )
            existing = order_first_seen.get(order_id)
            if existing is None or sort_key < existing:
                order_first_seen[order_id] = sort_key

    return tuple(
        order_id
        for order_id, _sort_key in sorted(order_first_seen.items(), key=lambda item: item[1])
    )


def build_timestamped_output_dir(base_output_dir: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return base_output_dir.parent / f"{base_output_dir.name}_{timestamp}"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="根據 Excel 訂單資料執行遺傳算法排程。")
    parser.add_argument(
        "--workbook",
        type=Path,
        required=True,
        help="Excel 檔案路徑，例如 2023訂單資訊&產能表-0319.xlsx",
    )
    parser.add_argument("--output-dir", type=Path, default=Path("ga_output"))
    parser.add_argument("--population-size", type=int, default=40)
    parser.add_argument("--generations", type=int, default=120)
    parser.add_argument("--mutation-rate", type=float, default=0.15)
    parser.add_argument("--crossover-rate", type=float, default=0.85)
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument(
        "--seed-schedule-csv",
        type=Path,
        help="Use an existing best_schedule_tasks.csv as a seed ordering for the initial GA population.",
    )
    parser.add_argument(
        "--exclude-warning-anomalies",
        action="store_true",
        help="將 warning 類型的異常訂單也排除出 GA 主優化集合。",
    )
    return parser.parse_args()


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

    args = parse_args()
    args.output_dir = build_timestamped_output_dir(args.output_dir)
    loader = WorkbookLoader(args.workbook)
    MACHINE_COUNTS.update(loader.machine_counts)
    orders = loader.load_orders()
    source_links = loader.build_order_source_links()
    if not orders:
        raise RuntimeError("沒有載入到可排程的訂單資料。")
    schedulable_orders, anomalies = split_schedulable_orders(
        orders,
        exclude_warnings=args.exclude_warning_anomalies,
    )
    if not schedulable_orders:
        raise RuntimeError("沒有可排程的訂單資料；請先修正日期窗衝突。")

    seeded_order_ids: tuple[str, ...] = ()
    if args.seed_schedule_csv is not None:
        seeded_order_ids = load_seed_order_ids(args.seed_schedule_csv)

    scheduler = GeneticScheduler(
        population_size=args.population_size,
        generations=args.generations,
        mutation_rate=args.mutation_rate,
        crossover_rate=args.crossover_rate,
        seed=args.seed,
        seeded_order_ids=seeded_order_ids,
    )
    best_schedule = scheduler.optimize(schedulable_orders)
    write_schedule_outputs(best_schedule, args.output_dir)
    write_daily_manufacturing_schedule_output(best_schedule, args.output_dir)
    write_daily_packaging_schedule_output(best_schedule, args.output_dir)
    write_manufacturing_cleaning_audit_output(best_schedule, args.output_dir)
    write_reference_format_output(best_schedule, orders, args.output_dir)
    write_order_anomalies_output(orders, anomalies, args.output_dir)
    validation_issues = validate_schedule(best_schedule, schedulable_orders)
    write_schedule_validation_output(validation_issues, args.output_dir)
    write_source_link_output(source_links, args.output_dir)
    write_machine_counts_output(loader.machine_counts, args.output_dir)
    write_order_model_audit_output(orders, source_links, args.output_dir)
    write_output_field_notes(args.output_dir)

    excluded_anomalies = [anomaly for anomaly in anomalies if anomaly.exclude_from_optimization]
    warning_anomalies = [anomaly for anomaly in anomalies if not anomaly.exclude_from_optimization]
    if excluded_anomalies:
        print("資料警示：部分訂單日期窗本身不可行，已先排除出 GA 主優化集合。")
        for anomaly in excluded_anomalies:
            print(f"  - {anomaly.order_id}: {anomaly.message}")
        print()

    if warning_anomalies:
        print("資料提醒：偵測到可能影響排程品質的異常資訊，已輸出 anomaly 清單。")
        for anomaly in warning_anomalies[:10]:
            print(f"  - {anomaly.order_id}: {anomaly.message}")
        if len(warning_anomalies) > 10:
            print(f"  - ... 其餘 {len(warning_anomalies) - 10} 筆請見 order_anomalies.csv")
        print()

    print("\n最佳排程摘要")
    print(f"原始訂單數量：{len(orders)}")
    print(f"實際參與 GA 的訂單數量：{len(schedulable_orders)}")
    print(f"異常訂單筆數：{len({anomaly.order_id for anomaly in anomalies})}")
    print(f"排程驗證問題數：{len(validation_issues)}")
    matched_source_links = sum(1 for row in source_links if row.get("match_count") == "1")
    print(f"來源訂單唯一對應數：{matched_source_links}/{len(source_links)}")
    print(f"總加班分鐘數：{best_schedule.total_overtime_minutes:.1f}")
    print(f"不可行懲罰：{best_schedule.infeasible_penalty:.1f}")
    print(f"目標懲罰：{best_schedule.target_penalty:.1f}")
    for labor_group in ("manufacturing", "filling", "bottling"):
        group_overtime = best_schedule.overtime_by_group.get(labor_group, 0.0)
        target = OVERTIME_TARGETS.get(labor_group)
        metric_name = f"{labor_group}_overtime_minutes"
        if target is None:
            print(f"{metric_name}: {group_overtime:.1f}")
            continue
        print(f"{metric_name}: {group_overtime:.1f} / target {target:.1f}")
        continue
        if target is None:
            print(f"{labor_group} 加班分鐘數：{group_overtime:.1f}")
        else:
            print(f"{labor_group} 加班分鐘數：{group_overtime:.1f} / 目標 {target:.1f}")
    print(f"週末加班分鐘數：{best_schedule.weekend_overtime_minutes:.1f}")
    print(f"目標值：{best_schedule.objective:.1f}")
    print(f"輸出資料夾：{args.output_dir.resolve()}")


if __name__ == "__main__":
    main()
