"""使用遺傳算法計算工廠訂單排程，目標為最小化總加班分鐘數。"""

from __future__ import annotations

import argparse
import csv
import html
import math
import random
import re
import shutil
import sys
import time
import xml.etree.ElementTree as ET
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Iterable, Mapping

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
FILLING_NONCONTIGUOUS_ORDER_WEIGHT = 25.0
VALIDATION_ISSUE_LABELS = {
    "order_not_scheduled": "可排訂單未排入",
    "fill_without_manufacturing": "充填缺少前置製造",
    "filling_minutes_mismatch": "充填工時不一致",
    "bottling_minutes_mismatch": "排瓶工時不一致",
    "premix_not_previous_day": "預攪未在攪料前一工作日",
    "bottling_not_previous_day": "排瓶未在充填前一工作日",
    "task_on_holiday": "排到非工作日",
    "manufacturing_before_raw_ready": "製造早於原料齊料日",
    "packaging_before_material_ready": "排瓶/充填早於物料齊料日",
    "task_after_due_date": "晚於交期",
    "intraday_resource_unallocatable": "日內確切時段資源不足",
}
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
DISPATCH_DAY_START_MINUTES = 9 * 60
DEFAULT_BREAK_START_TIME = "12:00"
DEFAULT_BREAK_END_TIME = "13:00"

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
    intraday_unallocated_task_count: int = 0
    filling_noncontiguous_order_count: int = 0
    overtime_by_group: dict[str, float] = field(default_factory=dict)
    weekend_overtime_minutes: float = 0.0
    line_change_minutes: float = 0.0
    line_change_minutes_per_switch: int = 0
    line_change_by_day_machine: dict[tuple[date, str], float] = field(default_factory=dict)
    break_start_time: str = DEFAULT_BREAK_START_TIME
    break_end_time: str = DEFAULT_BREAK_END_TIME
    frozen_schedule_entries: tuple[FrozenScheduleEntry, ...] = ()
    calendar_capacity: CalendarCapacity | None = None
    task_records: list[ScheduledTask] = field(default_factory=list)
    daily_labor_usage: dict[tuple[str, date], float] = field(default_factory=dict)
    daily_machine_usage: dict[tuple[str, date], float] = field(default_factory=dict)

    @property
    def objective(self) -> float:
        return self.total_overtime_minutes + self.infeasible_penalty + self.target_penalty


@dataclass
class SchedulerRunResult:
    output_dir: Path
    best_schedule: ScheduleResult
    orders: list[Order]
    schedulable_orders: list[Order]
    anomalies: list[OrderAnomaly]
    validation_issues: list[ScheduleValidationIssue]
    source_links: list[dict[str, str]]
    elapsed_seconds: float

@dataclass(frozen=True)
class ScheduleValidationIssue:
    order_id: str
    issue_type: str
    severity: str
    message: str

@dataclass(frozen=True)
class IntradayTaskAllocation:
    """A dispatch-time placement for one task after day-level scheduling."""

    task: ScheduledTask
    labor_pool: str
    actual_people: int
    start_slot: int | None
    end_slot: int | None

    @property
    def is_allocated(self) -> bool:
        return self.start_slot is not None and self.end_slot is not None

@dataclass(frozen=True)
class FrozenScheduleEntry:
    order_id: str
    task_type: str
    scheduled_date: date
    source_stage: str = ""


@dataclass(frozen=True)
class CalendarDayCapacity:
    current_day: date
    is_holiday: bool
    filling_people: int
    manufacturing_people: int
    note: str = ""


@dataclass(frozen=True)
class CalendarCapacity:
    days: Mapping[date, CalendarDayCapacity] = field(default_factory=dict)
    exclude_weekends: bool = True

    def is_workday(self, current_day: date) -> bool:
        day_config = self.days.get(current_day)
        if day_config is not None:
            if day_config.is_holiday:
                return False
            if self.exclude_weekends and current_day.weekday() >= 5:
                return False
            return True
        if self.exclude_weekends and current_day.weekday() >= 5:
            return False
        return True

    def previous_workday(self, before_day: date, min_day: date) -> date | None:
        current_day = before_day - timedelta(days=1)
        while current_day >= min_day:
            if self.is_workday(current_day):
                return current_day
            current_day -= timedelta(days=1)
        return None

    def workdays_between(self, start_day: date, end_day: date) -> list[date]:
        if end_day < start_day:
            return []
        return [
            start_day + timedelta(days=offset)
            for offset in range((end_day - start_day).days + 1)
            if self.is_workday(start_day + timedelta(days=offset))
        ]

    def regular_labor_minutes(self, labor_group: str, current_day: date, default_minutes: float) -> float:
        if not self.is_workday(current_day):
            return 0.0
        day_config = self.days.get(current_day)
        if day_config is None:
            return default_minutes
        if labor_group == "manufacturing":
            return day_config.manufacturing_people * 8 * 60
        if labor_group == "filling":
            return day_config.filling_people * 8 * 60
        if labor_group == "bottling":
            return min(day_config.filling_people, BOTTLING_MAX_PEOPLE) * 8 * 60
        return default_minutes

    def max_labor_minutes(self, labor_group: str, current_day: date, default_minutes: float) -> float:
        if not self.is_workday(current_day):
            return 0.0
        day_config = self.days.get(current_day)
        if day_config is None:
            return default_minutes
        if labor_group == "manufacturing":
            return day_config.manufacturing_people * 10 * 60
        if labor_group == "filling":
            return day_config.filling_people * 10 * 60
        if labor_group == "bottling":
            return min(day_config.filling_people, BOTTLING_MAX_PEOPLE) * 10 * 60
        return default_minutes


@dataclass(frozen=True)
class SchedulerAdvancedOptions:
    """Normalized options shared by CLI, Tkinter, metadata, and output writers."""

    raw_options: Mapping[str, object] = field(default_factory=dict)
    line_change_minutes: int = 0
    frozen_schedule_csv: Path | None = None
    manual_frozen_schedule_entries: tuple[FrozenScheduleEntry, ...] = ()
    calendar_capacity_path: Path | None = None
    exclude_weekends: bool = True
    output_week_view: bool = True
    schedule_mode: str = "full_ga"
    keep_split_days_contiguous: bool = True
    manufacturing_efficiency_path: Path | None = None
    filling_efficiency_path: Path | None = None
    break_start_time: str = DEFAULT_BREAK_START_TIME
    break_end_time: str = DEFAULT_BREAK_END_TIME

    @classmethod
    def from_mapping(
        cls,
        advanced_options: Mapping[str, object] | None,
    ) -> "SchedulerAdvancedOptions":
        raw_options = dict(advanced_options or {})
        schedule_mode = str(raw_options.get("schedule_mode") or "full_ga")
        if schedule_mode == "filling_only_planning":
            schedule_mode = "full_ga"
        return cls(
            raw_options=raw_options,
            line_change_minutes=_option_int(raw_options, "line_change_minutes", 0),
            frozen_schedule_csv=_option_path(raw_options, "frozen_schedule_csv"),
            manual_frozen_schedule_entries=_manual_frozen_entries(raw_options.get("frozen_schedule_entries")),
            calendar_capacity_path=_option_path(raw_options, "calendar_capacity_path"),
            exclude_weekends=_option_bool(raw_options, "exclude_weekends", True),
            output_week_view=_option_bool(raw_options, "output_week_view", True),
            schedule_mode=schedule_mode,
            keep_split_days_contiguous=_option_bool(raw_options, "keep_split_days_contiguous", True),
            manufacturing_efficiency_path=_option_path(raw_options, "manufacturing_efficiency_path"),
            filling_efficiency_path=_option_path(raw_options, "filling_efficiency_path"),
            break_start_time=_option_text(raw_options, "break_start_time", DEFAULT_BREAK_START_TIME),
            break_end_time=_option_text(raw_options, "break_end_time", DEFAULT_BREAK_END_TIME),
        )

    @property
    def independent_sections(self) -> bool:
        return self.schedule_mode == "independent_sections"

    def as_dict(self) -> dict[str, object]:
        normalized = dict(self.raw_options)
        normalized["line_change_minutes"] = self.line_change_minutes
        normalized["frozen_schedule_csv"] = str(self.frozen_schedule_csv or "")
        normalized["calendar_capacity_path"] = str(self.calendar_capacity_path or "")
        normalized["exclude_weekends"] = self.exclude_weekends
        normalized["manual_frozen_schedule_count"] = len(self.manual_frozen_schedule_entries)
        normalized["output_week_view"] = self.output_week_view
        normalized["schedule_mode"] = self.schedule_mode
        normalized["keep_split_days_contiguous"] = self.keep_split_days_contiguous
        normalized["manufacturing_efficiency_path"] = str(self.manufacturing_efficiency_path or "")
        normalized["filling_efficiency_path"] = str(self.filling_efficiency_path or "")
        normalized["break_start_time"] = self.break_start_time
        normalized["break_end_time"] = self.break_end_time
        return normalized

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
        "line_change_minutes": schedule.line_change_minutes,
        "intraday_unallocated_task_count": float(schedule.intraday_unallocated_task_count),
        "filling_noncontiguous_order_count": float(schedule.filling_noncontiguous_order_count),
        "total_bottling_minutes": total_bottling_minutes,
        "total_filling_minutes": total_filling_minutes,
        "total_filling_with_line_change_minutes": total_filling_minutes + schedule.line_change_minutes,
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
            ("充填換線分鐘數", "依同日同充填機不同產品數推估的換線時間。", "用來回答加入一小時換線後對充填負荷的影響。", "若數值高，表示同日同機台產品切換太多。"),
            ("總排瓶分鐘數", "整個排程所有排瓶工序的機器分鐘總和。", "可用來和排瓶超時一起判讀負荷強度。", "總量高但超時低，代表攤得開；總量高且超時高，代表排得很滿。"),
            ("總充填分鐘數", "整個排程所有充填工序的機器分鐘總和。", "可與充填超時一起判斷充填段緊迫程度。", "適合和業主提供的充填 baseline 比較。"),
            ("總充填含換線分鐘數", "充填作業分鐘加上推估換線分鐘。", "用來看加入換線後充填端總負荷。", "若比總充填分鐘數高很多，代表換線是重要瓶頸。"),
            ("日內未配置任務數", "無法在單日確切時段內同時滿足人力與機台的任務數。", "此值必須為 0 才能直接作為派工時間表。", "大於 0 時請先重排，不可直接交付。"),
            ("非連續充填訂單數", "同一訂單被拆到不相鄰工作日充填的訂單數。", "勾選連續排程時會列入軟性目標。", "不一定阻擋交付，但應由現場確認是否可接受。"),
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
    "filling_line_change_audit.csv": {
        "table_note": "充填換線稽核表。用來檢查同一天同一台充填機是否切換多種產品，並估算換線分鐘。",
        "fields": [
            ("日期", "充填換線稽核日期。", "以天為單位看換線集中在哪些日期。", "若同一天換線分鐘高，代表派工可能太碎。"),
            ("充填機", "被稽核的充填設備。", "用來看哪台充填機切換最多。", "可和 daily packaging schedule 對照。"),
            ("充填分鐘", "當日該充填機原本的充填作業分鐘。", "不含換線，代表原作業負荷。", "若加上換線後超過正常容量，就需要調整排序。"),
            ("不同產品數", "當日同機台涉及的不同產品數。", "用來估算最少切換次數。", "不同產品數愈多，換線風險愈高。"),
            ("推估換線次數", "以不同產品數 - 1 推估的最少換線次數。", "目前是可解釋近似，尚非完整日內 sequencing。", "若要更精準，下一版需接日內排序。"),
            ("換線分鐘", "推估換線次數乘以每次換線分鐘。", "此分鐘已併入充填端超時與摘要。", "若設定 line_change_minutes=60，這欄就是一小時換線情境。"),
            ("訂單清單", "造成該日該機台充填與換線的訂單集合。", "方便回頭追查來源訂單。", "可用於和業主討論是否合併或凍結。"),
            ("產品清單", "造成切換的產品集合。", "方便看是否為產品差異導致換線。", "若產品清單太長，可考慮同品項連續排。"),
        ],
    },
    "frozen_schedule_audit.csv": {
        "table_note": "凍結排程稽核表。用來確認手動指定的訂單、工序與日期是否成功套用。",
        "fields": [
            ("訂單", "使用者指定要凍結的訂單編號。", "用來和輸入 CSV 與排程任務明細串接。", "若訂單不存在，狀態會顯示未排入。"),
            ("指定工序", "使用者輸入的工序文字。", "可接受 Cat1/Cat2/Cat3/Cat4 或中文工序。", "若工序名稱不清楚，需回頭修凍結 CSV。"),
            ("模型工序", "程式轉換後的內部工序代碼。", "用來確認中文工序是否被正確映射。", "Cat3=排瓶，Cat4=充填。"),
            ("指定日期", "使用者希望鎖定的排程日期。", "排程時會嘗試將該訂單工序排在這一天。", "若違反容量或時間窗，可能無法成功套用。"),
            ("實際日期", "模型最後排出的日期集合。", "用來確認是否真的固定成功。", "若和指定日期不同，就要看狀態說明。"),
            ("狀態", "凍結套用結果。", "已套用代表所有該工序任務都在指定日期。", "未套用或部分套用需人工檢查。"),
            ("說明", "狀態的補充文字。", "協助理解失敗原因。", "常見原因包含訂單未進入 GA、容量不足或日期窗衝突。"),
        ],
    },
    "dispatch_week_view.csv": {
        "table_note": "派工週視圖。以一週一列、星期一到星期日為欄位，呈現接近 calendar 的派工摘要。",
        "fields": [
            ("週起始日", "該列週曆的星期一起始日期。", "方便用週為單位瀏覽排程。", "可直接對應業主希望看的 weekly calendar。"),
            ("星期一到星期日", "每天的派工摘要格。", "每格包含日期、工序、訂單、產品摘要與時間量。", "若格內有換線或超時提醒，代表該天需優先檢查。"),
            ("凍結", "若當天有手動凍結任務，會顯示凍結標記。", "方便確認人工指定的排程是否出現在週視圖上。", "若凍結稽核失敗，需回看 17_凍結排程稽核。"),
            ("換線", "若當天有推估充填換線，會顯示換線分鐘。", "方便判斷同日同機台切換是否過多。", "若換線高，可能要把同品項排近一點。"),
            ("超時", "若當天有任一資源超時，會顯示超時提醒。", "方便派工時先看風險日。", "若超時過高，需調整工序日期或人工凍結。"),
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
            ("時段範圍", "把一天切成 5 分鐘後，此任務在日內占用的起訖時段。", "已同時檢查機台與共用人力；結束時間為最後占用 slot 的結束邊界。", "若顯示未配置，代表當日無法產生可派工時段。"),
            ("時段數", "此任務總共占用幾個 5 分鐘時段。", "乘以 5 分鐘後應接近作業總時間。", "可用來檢查時段與分鐘是否一致。"),
            ("排程類型", "時段配置狀態。", "正常代表已取得符合日內機台與人力限制的時段；阻擋代表配置失敗。", "出現阻擋時，validation 會產生 error，不可直接派工。"),
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
            ("severity", "問題嚴重程度。", "error 代表硬性違規，warning 代表仍可交付但建議人工注意。", "正式版最少要做到 error = 0。"),
            ("issue_type", "違規類型代碼。", "例如無排程、先後順序錯誤、缺工序等。", "可快速聚焦是哪種邏輯出問題。"),
            ("message", "驗證訊息。", "更完整描述規則違反內容。", "最適合拿來 debug。"),
        ],
    },
    "schedule_validation_summary.csv": {
        "table_note": "排程驗證摘要表。比起 validation 明細，這張更適合快速確認正式結果是否可交付。",
        "fields": [
            ("metric", "驗證指標名稱。", "例如 blocking_issue_count、error_count、warning_count 等。", "用來快速看排程是否仍有硬性違規。"),
            ("value", "該指標的數值。", "代表 validation 的統計結果。", "正式版通常至少要確保 blocking_issue_count 與 error_count 為 0。"),
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
    "input_model_issues.csv": {
        "table_note": "輸入建模缺漏表。列出在輸入中存在、但缺少可建模工序而未能進入排程的訂單。",
        "fields": [
            ("order_id", "輸入資料中的訂單編號。", "用來追查被漏建模的訂單。", "demo 前理想上只有欄名，沒有資料列。"),
            ("product_name", "產品名稱。", "方便回原始 Excel 查找。", "需確認該品項是否本來就不需某工序。"),
            ("issue_type", "缺漏類型代碼。", "目前會標示 missing_model_operation。", "代表不能靜默視為已排程。"),
            ("message", "缺漏文字說明。", "說明欠缺製造或充填工序。", "若有資料列，需先補來源或確認業務規則。"),
        ],
    },
    "machine_counts.csv": {
        "table_note": "模型使用的機器台數表。用來確認程式實際採用了哪些機器台數設定。",
        "fields": [
            ("machine_type", "機器類型名稱。", "對應排程內的設備名稱。", "應和日排程中的機器名稱一致。"),
            ("count", "模型假設的可用台數。", "直接影響該設備每日總容量。", "若業主提供台數不同，這張表一定要先對。"),
        ],
    },
    "run_metadata.csv": {
        "table_note": "本次執行資訊表。用來記錄一次正式版或 quick run 實際跑了多久。",
        "fields": [
            ("欄位", "執行資訊的項目名稱。", "例如執行開始時間、總執行秒數、GA 參數。", "可用來快速找到這次 run 的關鍵資訊。"),
            ("數值", "對應項目的內容。", "會包含時間戳記、總秒數、時分秒格式與參數設定。", "總執行秒數與 HH:MM:SS 應該互相對得起來。"),
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

FIELD_NOTE_HEADER = ["表單檔名", "表單用途", "欄位名稱", "欄位說明", "如何判讀", "合理性檢查重點"]

def _write_field_note_csv(path: Path, rows: list[list[str]]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(FIELD_NOTE_HEADER)
        writer.writerows(rows)

def write_output_field_notes(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    note_dir = output_dir / "欄位備註表"
    note_dir.mkdir(parents=True, exist_ok=True)

    master_rows: list[list[str]] = []
    for filename, config in OUTPUT_FIELD_NOTES.items():
        table_note = str(config.get("table_note", ""))
        fields = list(config.get("fields", []))
        for field_name, meaning, reading, check in fields:
            master_rows.append([filename, table_note, field_name, meaning, reading, check])

    _write_field_note_csv(output_dir / "欄位備註總表.csv", master_rows)
    _write_field_note_csv(note_dir / "欄位備註總表.csv", master_rows)

    for filename, config in OUTPUT_FIELD_NOTES.items():
        safe_stem = Path(filename).stem
        note_path = note_dir / f"{safe_stem}_欄位備註表.csv"
        table_note = str(config.get("table_note", ""))
        fields = list(config.get("fields", []))
        rows = [[filename, table_note, field_name, meaning, reading, check] for field_name, meaning, reading, check in fields]
        _write_field_note_csv(note_path, rows)


DELIVERY_OUTPUT_FILES = [
    ("schedule_check_report.xlsx", "01_排程檢查報告.xlsx", "固定格式 Excel 檢查報告，對齊業主提供的檢查表欄位。"),
    ("daily_schedule_exact_time.xlsx", "02_當日排程確切時間.xlsx", "含日內機台/人力時段檢查；僅在 validation 無阻擋錯誤時可作派工表。"),
    ("google_calendar_schedule.ics", "03_Google日曆匯入檔.ics", "可匯入 Google Calendar 的排程事件檔。"),
    ("supporting_outputs.xlsx", "04_輔助檢查資料.xlsx", "把一頁式總覽、驗證摘要、瓶頸、異常、稽核與欄位備註集中在同一本 Excel。"),
]

def _filling_product_key(task: ScheduledTask) -> str:
    detail = task.detail.rsplit("-充填", 1)[0]
    return _normalize_output_text(detail.strip())

def build_line_change_audit_rows(schedule: ScheduleResult) -> list[dict[str, object]]:
    machine_day_products: dict[tuple[date, str], set[str]] = defaultdict(set)
    machine_day_orders: dict[tuple[date, str], set[str]] = defaultdict(set)
    machine_day_minutes: dict[tuple[date, str], float] = defaultdict(float)
    for task in schedule.task_records:
        if task.task_type != "Cat4":
            continue
        key = (task.scheduled_date, task.machine_type)
        machine_day_products[key].add(_filling_product_key(task))
        machine_day_orders[key].add(task.order_id)
        machine_day_minutes[key] += task.machine_minutes

    rows: list[dict[str, object]] = []
    for key, products in sorted(machine_day_products.items()):
        scheduled_date, machine_type = key
        switch_count = max(len(products) - 1, 0)
        line_change_minutes = schedule.line_change_by_day_machine.get(key, 0.0)
        rows.append(
            {
                "date": scheduled_date,
                "machine": machine_type,
                "filling_minutes": machine_day_minutes[key],
                "distinct_product_count": len(products),
                "estimated_switch_count": switch_count,
                "line_change_minutes": line_change_minutes,
                "orders": " | ".join(sorted(machine_day_orders[key])),
                "products": " | ".join(sorted(products)),
            }
        )
    return rows


def _daily_resource_load_rows(schedule: ScheduleResult) -> list[dict[str, object]]:
    daily_base_usage: dict[tuple[str, date], float] = defaultdict(float)
    for task in schedule.task_records:
        if task.task_type in {"Cat1", "Cat2"}:
            daily_base_usage[("manufacturing", task.scheduled_date)] += task.labor_minutes
        elif task.task_type == "Cat3":
            daily_base_usage[("bottling", task.scheduled_date)] += task.labor_minutes
        elif task.task_type == "Cat4":
            daily_base_usage[("filling", task.scheduled_date)] += task.machine_minutes
    for (current_day, _machine_type), line_change_minutes in schedule.line_change_by_day_machine.items():
        daily_base_usage[("filling", current_day)] += line_change_minutes

    regular_by_group = {
        "manufacturing": LABOR_CONFIG["manufacturing"]["regular_minutes"],
        "bottling": PROCESS_OVERTIME_REGULAR_MINUTES["bottling"],
        "filling": PROCESS_OVERTIME_REGULAR_MINUTES["filling"],
    }
    rows: list[dict[str, object]] = []
    for (labor_group, current_day), used_minutes in sorted(daily_base_usage.items()):
        if current_day == date.min or used_minutes <= 0:
            continue
        regular_minutes = regular_by_group[labor_group]
        overtime = max(used_minutes - regular_minutes, 0.0)
        rows.append(
            {
                "date": current_day,
                "resource": _display_labor_group(labor_group),
                "used": used_minutes,
                "regular": regular_minutes,
                "overtime": overtime,
            }
        )
    return rows


def _ics_escape(value: object) -> str:
    text = str(value or "")
    return (
        text.replace("\\", "\\\\")
        .replace(";", r"\;")
        .replace(",", r"\,")
        .replace("\r\n", r"\n")
        .replace("\n", r"\n")
    )


def _ics_datetime(current_day: date, hour: int, minute: int = 0) -> str:
    return f"{current_day.strftime('%Y%m%d')}T{hour:02d}{minute:02d}00"


def _ics_add_minutes(current_day: date, hour: int, minute: int, duration_minutes: int) -> tuple[date, int, int]:
    start = datetime.combine(current_day, datetime.min.time()).replace(hour=hour, minute=minute)
    end = start + timedelta(minutes=max(duration_minutes, SLOT_MINUTES))
    return end.date(), end.hour, end.minute


def write_google_calendar_ics(schedule: ScheduleResult, orders: list[Order], output_dir: Path) -> None:
    """Write an .ics file that can be imported into Google Calendar."""
    output_dir.mkdir(parents=True, exist_ok=True)
    grouped_tasks = _group_output_tasks(schedule.task_records)
    order_map = {order.order_id: order for order in orders}
    daily_stage: dict[tuple[date, str], list[ScheduledTask]] = defaultdict(list)
    stage_order = {"預攪": 1, "攪料清洗": 2, "排瓶": 3, "充填": 4}
    stage_start = {
        "預攪": (9, 0),
        "攪料清洗": (10, 30),
        "排瓶": (13, 0),
        "充填": (15, 0),
    }

    for task in grouped_tasks:
        daily_stage[(task.scheduled_date, _display_task_stage(task.task_type))].append(task)

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Factory GA Scheduler//Factory Schedule//ZH-TW",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:工廠訂單排程",
        "X-WR-TIMEZONE:Asia/Taipei",
    ]
    dtstamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    for event_index, ((scheduled_date, stage), tasks) in enumerate(
        sorted(
            daily_stage.items(),
            key=lambda item: (item[0][0], stage_order.get(item[0][1], 99), item[0][1]),
        ),
        start=1,
    ):
        order_ids = sorted({task.order_id for task in tasks})
        product_names = []
        for order_id in order_ids[:5]:
            order = order_map.get(order_id)
            if order is not None:
                product_names.append(f"{order_id} {order.product_name}")
        if len(order_ids) > 5:
            product_names.append(f"...其餘 {len(order_ids) - 5} 張")

        total_machine_minutes = _display_integer(sum(task.machine_minutes for task in tasks))
        total_labor_minutes = _display_integer(sum(task.labor_minutes for task in tasks))
        visual_duration = min(max(total_machine_minutes, 30), 240)
        start_hour, start_minute = stage_start.get(stage, (9, 0))
        end_day, end_hour, end_minute = _ics_add_minutes(
            scheduled_date,
            start_hour,
            start_minute,
            visual_duration,
        )

        summary = f"{stage}：{len(order_ids)}張訂單"
        description = (
            f"工序：{stage}\\n"
            f"訂單：{' | '.join(order_ids)}\\n"
            f"品項：{' | '.join(product_names)}\\n"
            f"機器分鐘：{total_machine_minutes}\\n"
            f"人力分鐘：{total_labor_minutes}\\n"
            "此事件由工廠訂單排程 GA 自動產生。"
        )
        uid = f"{scheduled_date.strftime('%Y%m%d')}-{stage_order.get(stage, 99)}-{event_index}@factory-ga-scheduler"
        lines.extend(
            [
                "BEGIN:VEVENT",
                f"UID:{_ics_escape(uid)}",
                f"DTSTAMP:{dtstamp}",
                f"DTSTART;TZID=Asia/Taipei:{_ics_datetime(scheduled_date, start_hour, start_minute)}",
                f"DTEND;TZID=Asia/Taipei:{_ics_datetime(end_day, end_hour, end_minute)}",
                f"SUMMARY:{_ics_escape(summary)}",
                f"DESCRIPTION:{_ics_escape(description)}",
                "LOCATION:工廠",
                "END:VEVENT",
            ]
        )

    lines.append("END:VCALENDAR")
    (output_dir / "google_calendar_schedule.ics").write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_intuitive_output_tables(
    schedule: ScheduleResult,
    orders: list[Order],
    schedulable_orders: list[Order],
    anomalies: list[OrderAnomaly],
    validation_issues: list[ScheduleValidationIssue],
    output_dir: Path,
) -> None:
    """Write stakeholder-friendly tables with fewer model-specific columns."""
    output_dir.mkdir(parents=True, exist_ok=True)
    summary = build_schedule_summary(schedule)
    grouped_tasks = _group_output_tasks(schedule.task_records)
    daily_load_rows = _daily_resource_load_rows(schedule)
    order_map = {order.order_id: order for order in orders}
    tasks_by_order: dict[str, list[ScheduledTask]] = defaultdict(list)
    anomalies_by_order: dict[str, list[OrderAnomaly]] = defaultdict(list)
    validation_by_order: dict[str, list[ScheduleValidationIssue]] = defaultdict(list)

    for task in grouped_tasks:
        tasks_by_order[task.order_id].append(task)
    for anomaly in anomalies:
        anomalies_by_order[anomaly.order_id].append(anomaly)
    for issue in validation_issues:
        validation_by_order[issue.order_id].append(issue)

    max_overtime_row = max(daily_load_rows, key=lambda row: float(row["overtime"]), default=None)
    bottleneck_text = "無明顯超時日"
    if max_overtime_row and float(max_overtime_row["overtime"]) > 0:
        bottleneck_text = (
            f"{max_overtime_row['date']} "
            f"{max_overtime_row['resource']}超時{_display_integer(float(max_overtime_row['overtime']))}分鐘"
        )
    validation_error_count = sum(1 for issue in validation_issues if issue.severity == "error")
    validation_warning_count = sum(1 for issue in validation_issues if issue.severity == "warning")
    scheduled_order_ids = {task.order_id for task in schedule.task_records}
    missing_schedulable_order_ids = sorted(
        order.order_id for order in schedulable_orders if order.order_id not in scheduled_order_ids
    )
    schedulable_coverage_text = (
        "全部可排訂單已排入"
        if not missing_schedulable_order_ids
        else "未排入：" + "、".join(missing_schedulable_order_ids[:8])
    )

    dashboard_rows = [
        ["分類", "指標", "數值", "判讀"],
        ["整體", "原始訂單數量", len(orders), "資料中讀到的訂單總數。"],
        ["整體", "實際參與 GA 的訂單數量", len(schedulable_orders), "扣除明顯無解訂單後，實際進入排程的數量。"],
        ["整體", "可排訂單排入狀態", schedulable_coverage_text, "正式 demo 要能說明：合理訂單全部排入；不合理訂單列在異常清單。"],
        ["整體", "排除型異常訂單數", len({item.order_id for item in anomalies if item.exclude_from_optimization}), "這些訂單通常是資料窗本身不可行，需要人工處理。"],
        ["驗證", "硬性違規 error_count", validation_error_count, "正式交付前應為 0。"],
        ["驗證", "人工提醒 warning_count", validation_warning_count, "通常不阻擋交付，但建議檢查是否符合現場習慣。"],
        ["驗證", "日內未配置任務數", schedule.intraday_unallocated_task_count, "大於 0 代表確切時段不可直接派工。"],
        ["連續性", "非連續充填訂單數", schedule.filling_noncontiguous_order_count, "勾選連續排程時，此數值會列入軟性目標。"],
        ["加班", "總加班分鐘數", _display_integer(summary["total_overtime_minutes"]), "本次排程最主要的目標值。"],
        ["加班", "調製超時分鐘數", _display_integer(summary["manufacturing_overtime_minutes"]), "調製端是否為瓶頸。"],
        ["加班", "排瓶超時分鐘數", _display_integer(summary["bottling_overtime_minutes"]), "排瓶端是否為瓶頸。"],
        ["加班", "充填超時分鐘數", _display_integer(summary["filling_overtime_minutes"]), "充填端是否為瓶頸。"],
        ["加班", "週末加班分鐘數", _display_integer(summary["weekend_overtime_minutes"]), "是否依賴假日排程。"],
        ["換線", "充填換線分鐘數", _display_integer(summary["line_change_minutes"]), "同日同機台不同產品造成的推估換線時間。"],
        ["瓶頸", "最大單日瓶頸", bottleneck_text, "若有超時，建議先從這一天往下查。"],
    ]
    with (output_dir / "intuitive_dashboard.csv").open("w", newline="", encoding="utf-8-sig") as file:
        csv.writer(file).writerows(dashboard_rows)

    with (output_dir / "intuitive_bottleneck_alerts.csv").open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["日期", "資源", "使用分鐘", "正常分鐘", "超時分鐘", "提醒"])
        for row in daily_load_rows:
            overtime = float(row["overtime"])
            if overtime <= 0:
                continue
            writer.writerow(
                [
                    row["date"].isoformat(),
                    row["resource"],
                    _display_integer(float(row["used"])),
                    _display_integer(float(row["regular"])),
                    _display_integer(overtime),
                    "需要優先檢查是否能提前、拆分或人工調整。",
                ]
            )

    with (output_dir / "intuitive_daily_schedule.csv").open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["日期", "工序", "訂單數", "訂單清單", "品項摘要", "機器分鐘", "人力分鐘"])
        daily_stage: dict[tuple[date, str], list[ScheduledTask]] = defaultdict(list)
        stage_order = {"預攪": 1, "攪料清洗": 2, "排瓶": 3, "充填": 4}
        for task in grouped_tasks:
            daily_stage[(task.scheduled_date, _display_task_stage(task.task_type))].append(task)
        for (scheduled_date, stage), tasks in sorted(
            daily_stage.items(),
            key=lambda item: (item[0][0], stage_order.get(item[0][1], 99), item[0][1]),
        ):
            order_ids = sorted({task.order_id for task in tasks})
            product_names = []
            for order_id in order_ids[:3]:
                order = order_map.get(order_id)
                if order is not None:
                    product_names.append(order.product_name)
            if len(order_ids) > 3:
                product_names.append(f"...等{len(order_ids)}張")
            writer.writerow(
                [
                    scheduled_date.isoformat(),
                    stage,
                    len(order_ids),
                    " | ".join(order_ids),
                    " | ".join(product_names),
                    _display_integer(sum(task.machine_minutes for task in tasks)),
                    _display_integer(sum(task.labor_minutes for task in tasks)),
                ]
            )

    with (output_dir / "intuitive_order_schedule.csv").open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "訂單",
                "產品名稱",
                "數量",
                "交期",
                "原料齊料日",
                "包材/物料齊料日",
                "調製日期",
                "排瓶日期",
                "充填日期",
                "最後作業日",
                "狀態",
                "提醒",
            ]
        )
        for order in sorted(orders, key=lambda item: (item.due_date, item.order_id)):
            tasks = tasks_by_order.get(order.order_id, [])
            task_days = [task.scheduled_date for task in tasks]
            manufacturing_days = sorted({task.scheduled_date for task in tasks if task.task_type in {"Cat1", "Cat2"}})
            bottling_days = sorted({task.scheduled_date for task in tasks if task.task_type == "Cat3"})
            filling_days = sorted({task.scheduled_date for task in tasks if task.task_type == "Cat4"})
            last_day = max(task_days) if task_days else None
            order_anomalies = anomalies_by_order.get(order.order_id, [])
            order_issues = validation_by_order.get(order.order_id, [])
            if order_issues:
                status = "需檢查"
            elif any(item.exclude_from_optimization for item in order_anomalies):
                status = "未排入-資料/日期窗異常"
            elif tasks:
                status = "已排入"
            else:
                status = "未排入"
            reminders = [item.message for item in order_anomalies[:2]] + [issue.message for issue in order_issues[:2]]
            writer.writerow(
                [
                    order.order_id,
                    order.product_name,
                    order.quantity,
                    order.due_date.isoformat(),
                    order.raw_ready_date.isoformat(),
                    order.material_ready_date.isoformat(),
                    " | ".join(day.isoformat() for day in manufacturing_days),
                    " | ".join(day.isoformat() for day in bottling_days),
                    " | ".join(day.isoformat() for day in filling_days),
                    last_day.isoformat() if last_day else "",
                    status,
                    " | ".join(reminders),
                ]
            )


def write_delivery_output_pack(output_dir: Path) -> None:
    """Create a user-facing subset of the most useful outputs."""
    delivery_dir = output_dir / "交付用輸出"
    delivery_dir.mkdir(parents=True, exist_ok=True)

    index_rows = [["順序", "檔案名稱", "來源檔案", "用途"]]
    for index, (source_name, target_name, description) in enumerate(DELIVERY_OUTPUT_FILES, start=1):
        source_path = output_dir / source_name
        target_path = delivery_dir / target_name
        if source_path.exists():
            shutil.copy2(source_path, target_path)
            status = description
        else:
            status = f"來源檔案不存在：{source_name}"
        index_rows.append([str(index), target_name, source_name, status])

    with (delivery_dir / "00_輸出閱讀順序.csv").open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerows(index_rows)

    guide_lines = [
        "# 工廠排程輸出閱讀順序",
        "",
        "這個資料夾只保留最適合交付業主的精簡輸出。",
        "完整原始 CSV 仍保留在上一層資料夾，主要給開發者 debug 與追查模型細節。",
        "",
        "建議閱讀順序：",
    ]
    for index, (_source_name, target_name, description) in enumerate(DELIVERY_OUTPUT_FILES, start=1):
        guide_lines.append(f"{index}. `{target_name}`：{description}")
    guide_lines.extend(
        [
            "",
            "判讀重點：",
            "- 業主主要看 `01_排程檢查報告.xlsx` 與 `02_當日排程確切時間.xlsx`。",
            "- 若要放進 Google Calendar，匯入 `03_Google日曆匯入檔.ics`。",
            "- 若要查驗證摘要、瓶頸、異常訂單或凍結排程稽核，打開 `04_輔助檢查資料.xlsx`。",
            "- 正式交付前，先確認輔助檢查資料中的 `驗證摘要` 工作表，`error_count` / `blocking_issue_count` 應為 0。",
            "- 若 `02_當日排程確切時間.xlsx` 出現 `未配置：日內資源不足`，代表當日仍需重排，不可直接派工。",
            "- 若只有 `filling_days_not_contiguous` warning，排程未違反硬性限制，但仍應請現場確認拆單日期是否可接受。",
        ]
    )
    (delivery_dir / "README_先看我.md").write_text("\n".join(guide_lines) + "\n", encoding="utf-8")


def _xlsx_column_name(index: int) -> str:
    name = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _xlsx_text_width(value: object) -> int:
    text = "" if value is None else str(value)
    display_width = sum(2 if ord(character) > 127 else 1 for character in text)
    return min(max(display_width + 2, 8), 38)


def _xlsx_sheet_xml(rows: list[list[object]]) -> str:
    max_columns = max((len(row) for row in rows), default=1)
    last_column = _xlsx_column_name(max_columns - 1)
    has_title_row = (
        len(rows) > 1
        and sum(1 for value in rows[0] if str(value or "").strip()) == 1
        and len(rows[1]) > 1
    )
    header_row = 2 if has_title_row else 1

    widths = [8 for _ in range(max_columns)]
    for row in rows[:1500]:
        for index, value in enumerate(row):
            widths[index] = max(widths[index], _xlsx_text_width(value))
    columns_xml = "".join(
        f'<col min="{index}" max="{index}" width="{width}" customWidth="1"/>'
        for index, width in enumerate(widths, start=1)
    )

    row_xml: list[str] = []
    for row_index, row in enumerate(rows, start=1):
        style_index = 2 if has_title_row and row_index == 1 else 1 if row_index == header_row else 3
        row_attributes = ' ht="30" customHeight="1"' if row_index in {1, header_row} else ""
        cell_xml: list[str] = []
        for column_index, value in enumerate(row):
            text = html.escape("" if value is None else str(value), quote=False)
            ref = f"{_xlsx_column_name(column_index)}{row_index}"
            cell_xml.append(f'<c r="{ref}" s="{style_index}" t="inlineStr"><is><t>{text}</t></is></c>')
        row_xml.append(f'<row r="{row_index}"{row_attributes}>{"".join(cell_xml)}</row>')

    frozen_view = (
        '<sheetViews><sheetView workbookViewId="0">'
        f'<pane ySplit="{header_row}" topLeftCell="A{header_row + 1}" activePane="bottomLeft" state="frozen"/>'
        '</sheetView></sheetViews>'
    )
    merge_xml = (
        f'<mergeCells count="1"><mergeCell ref="A1:{last_column}1"/></mergeCells>'
        if has_title_row and max_columns > 1
        else ""
    )
    filter_xml = (
        f'<autoFilter ref="A{header_row}:{last_column}{len(rows)}"/>'
        if len(rows) > header_row
        else ""
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'{frozen_view}<cols>{columns_xml}</cols><sheetData>{"".join(row_xml)}</sheetData>'
        f'{merge_xml}{filter_xml}</worksheet>'
    )


def _xlsx_styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="3">'
        '<font><sz val="11"/><name val="Microsoft JhengHei"/></font>'
        '<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Microsoft JhengHei"/></font>'
        '<font><b/><color rgb="FF134E4A"/><sz val="14"/><name val="Microsoft JhengHei"/></font>'
        '</fonts>'
        '<fills count="4">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF0F766E"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFCCFBF1"/><bgColor indexed="64"/></patternFill></fill>'
        '</fills>'
        '<borders count="2">'
        '<border><left/><right/><top/><bottom/><diagonal/></border>'
        '<border><left style="thin"><color rgb="FFE5E7EB"/></left>'
        '<right style="thin"><color rgb="FFE5E7EB"/></right>'
        '<top style="thin"><color rgb="FFE5E7EB"/></top>'
        '<bottom style="thin"><color rgb="FFE5E7EB"/></bottom><diagonal/></border>'
        '</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="4">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1">'
        '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyAlignment="1">'
        '<alignment vertical="center"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1">'
        '<alignment vertical="top" wrapText="1"/></xf>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )


def write_simple_xlsx(path: Path, sheets: list[tuple[str, list[list[object]]]]) -> None:
    workbook_sheets = []
    workbook_rels = []
    content_types = [
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for index, (sheet_name, _rows) in enumerate(sheets, start=1):
        safe_name = html.escape(sheet_name[:31], quote=True)
        workbook_sheets.append(f'<sheet name="{safe_name}" sheetId="{index}" r:id="rId{index}"/>')
        workbook_rels.append(
            f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>'
        )
        content_types.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    workbook_rels.append(
        f'<Relationship Id="rId{len(sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets>{"".join(workbook_sheets)}</sheets></workbook>'
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{"".join(workbook_rels)}</Relationships>'
    )
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        f'{"".join(content_types)}</Types>'
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        zf.writestr("xl/styles.xml", _xlsx_styles_xml())
        for index, (_sheet_name, rows) in enumerate(sheets, start=1):
            zf.writestr(f"xl/worksheets/sheet{index}.xml", _xlsx_sheet_xml(rows))


def _csv_rows_for_xlsx(path: Path) -> list[list[object]]:
    if not path.exists():
        return [["提示", f"來源檔案不存在：{path.name}"]]
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        rows = [list(row) for row in csv.reader(file)]
    return rows or [["提示", f"來源檔案是空檔：{path.name}"]]


def write_compact_excel_outputs(output_dir: Path) -> None:
    """Bundle stakeholder-facing details into a few clean Excel workbooks."""
    daily_schedule_sheets = [
        ("確切時間總表", _csv_rows_for_xlsx(output_dir / "best_schedule_reference_format.csv")),
        ("每日調製明細", _csv_rows_for_xlsx(output_dir / "daily_manufacturing_schedule.csv")),
        ("每日排瓶充填明細", _csv_rows_for_xlsx(output_dir / "daily_packaging_schedule.csv")),
        ("派工週視圖", _csv_rows_for_xlsx(output_dir / "dispatch_week_view.csv")),
    ]
    write_simple_xlsx(output_dir / "daily_schedule_exact_time.xlsx", daily_schedule_sheets)

    supporting_sheets = [
        ("一頁式總覽", _csv_rows_for_xlsx(output_dir / "intuitive_dashboard.csv")),
        ("驗證摘要", _csv_rows_for_xlsx(output_dir / "schedule_validation_summary.csv")),
        ("輸入建模缺漏", _csv_rows_for_xlsx(output_dir / "input_model_issues.csv")),
        ("來源映射稽核", _csv_rows_for_xlsx(output_dir / "order_source_links.csv")),
        ("瓶頸提醒", _csv_rows_for_xlsx(output_dir / "intuitive_bottleneck_alerts.csv")),
        ("訂單排程總表", _csv_rows_for_xlsx(output_dir / "intuitive_order_schedule.csv")),
        ("每日負荷", _csv_rows_for_xlsx(output_dir / "best_schedule_daily_load.csv")),
        ("異常訂單", _csv_rows_for_xlsx(output_dir / "order_anomalies.csv")),
        ("排程警告明細", _csv_rows_for_xlsx(output_dir / "schedule_validation.csv")),
        ("充填換線稽核", _csv_rows_for_xlsx(output_dir / "filling_line_change_audit.csv")),
        ("凍結排程稽核", _csv_rows_for_xlsx(output_dir / "frozen_schedule_audit.csv")),
        ("執行參數與時間", _csv_rows_for_xlsx(output_dir / "run_metadata.csv")),
        ("欄位備註", _csv_rows_for_xlsx(output_dir / "欄位備註總表.csv")),
    ]
    write_simple_xlsx(output_dir / "supporting_outputs.xlsx", supporting_sheets)


def _task_date_range_text(tasks: list[ScheduledTask], task_type: str) -> tuple[str, str]:
    days = sorted({task.scheduled_date for task in tasks if task.task_type == task_type})
    if not days:
        return "", ""
    return days[0].isoformat(), days[-1].isoformat()


def write_fixed_check_report_xlsx(
    schedule: ScheduleResult,
    orders: list[Order],
    schedulable_orders: list[Order],
    anomalies: list[OrderAnomaly],
    validation_issues: list[ScheduleValidationIssue],
    source_links: list[dict[str, str]],
    output_dir: Path,
) -> None:
    tasks_by_order: dict[str, list[ScheduledTask]] = defaultdict(list)
    for task in schedule.task_records:
        tasks_by_order[task.order_id].append(task)
    source_by_order = {row.get("ga_order_id", ""): row for row in source_links}
    schedulable_ids = {order.order_id for order in schedulable_orders}

    completed_rows: list[list[object]] = [
        ["2023 化妝品工廠排程結果 — 訂單完成檢查表"],
        [
            "訂單編號",
            "產品代號",
            "產品名稱",
            "原料齊料點",
            "物料齊料點",
            "截止日",
            "預攪 開始",
            "預攪 結束",
            "攪料 開始",
            "攪料 結束",
            "清洗 日期",
            "排瓶 開始",
            "排瓶 結束",
            "充填 開始",
            "充填 結束",
            "緩衝天數",
        ],
    ]
    for order in sorted(orders, key=lambda item: (item.due_date, item.order_id)):
        tasks = tasks_by_order.get(order.order_id, [])
        cat1_start, cat1_end = _task_date_range_text(tasks, "Cat1")
        cat2_start, cat2_end = _task_date_range_text(tasks, "Cat2")
        cat3_start, cat3_end = _task_date_range_text(tasks, "Cat3")
        cat4_start, cat4_end = _task_date_range_text(tasks, "Cat4")
        last_fill = max((task.scheduled_date for task in tasks if task.task_type == "Cat4"), default=None)
        completed_rows.append(
            [
                order.source_order_id or order.order_id,
                order.source_product_code or source_by_order.get(order.order_id, {}).get("source_product_code", ""),
                order.product_name,
                order.raw_ready_date.isoformat(),
                order.material_ready_date.isoformat(),
                order.due_date.isoformat(),
                cat1_start,
                cat1_end,
                cat2_start,
                cat2_end,
                cat2_end,
                cat3_start,
                cat3_end,
                cat4_start,
                cat4_end,
                (order.due_date - last_fill).days if last_fill else "",
            ]
        )

    warning_rows: list[list[object]] = [
        ["訂單編號", "產品代號", "產品名稱", "排程日期", "物料齊料點", "截止日", "警告原因", "排瓶工時(min)", "充填工時(min)", "當日固定充填工時(min)"]
    ]
    order_map = {order.order_id: order for order in orders}
    for issue in validation_issues:
        order = order_map.get(issue.order_id)
        warning_rows.append(
            [
                issue.order_id,
                order.source_product_code if order else "",
                order.product_name if order else "",
                "",
                order.material_ready_date.isoformat() if order else "",
                order.due_date.isoformat() if order else "",
                f"{issue.severity}:{issue.issue_type} {issue.message}",
                "",
                "",
                "",
            ]
        )
    if len(warning_rows) == 1:
        warning_rows.append(["本次排程沒有排程警告。", "", "", "", "", "", "", "", "", ""])

    excluded_rows: list[list[object]] = [["訂單編號", "製造類型", "工序", "排除原因"]]
    for anomaly in anomalies:
        if anomaly.exclude_from_optimization or anomaly.order_id not in schedulable_ids:
            excluded_rows.append([anomaly.order_id, "", "", anomaly.message])
    if len(excluded_rows) == 1:
        excluded_rows.append(["本次排程沒有排除訂單。", "", "", ""])

    weekday_names = "一二三四五六日"
    daily_rows: list[list[object]] = [
        ["日期", "星期", "假日", "區塊", "工序", "項目", "訂單編號", "產品代號", "產品名稱", "數量(批/次/pcs)", "機器", "單批工時(min)/產能(pcs/hr)", "表定人數", "調製人員", "建議人數", "工序總工時(min)", "截止日"]
    ]
    for task in schedule.task_records:
        order = order_map.get(task.order_id)
        if order is None:
            continue
        is_holiday = schedule.calendar_capacity is not None and not schedule.calendar_capacity.is_workday(task.scheduled_date)
        daily_rows.append(
            [
                task.scheduled_date.isoformat(),
                weekday_names[task.scheduled_date.weekday()],
                "假日" if is_holiday else "平日",
                "調製" if task.task_type in {"Cat1", "Cat2"} else "充填",
                _display_task_stage(task.task_type),
                task.detail,
                order.source_order_id or order.order_id,
                order.source_product_code,
                order.product_name,
                int(order.quantity),
                _normalize_output_text(task.machine_type),
                _display_integer(task.machine_minutes),
                _display_integer(task.labor_minutes / max(task.machine_minutes, 1)) if task.machine_minutes else "",
                "",
                "",
                _display_integer(max(task.machine_minutes, task.labor_minutes)),
                order.due_date.isoformat(),
            ]
        )

    daily_load_rows: list[list[object]] = [
        ["日期", "星期", "假日", "當日調製總工時(min)", "當日充填總工時(min)", "調製加班(min)", "充填加班(min)", "總加班(min)"]
    ]
    all_days = sorted({task.scheduled_date for task in schedule.task_records})
    for current_day in all_days:
        manufacturing_minutes = sum(
            task.labor_minutes for task in schedule.task_records if task.scheduled_date == current_day and task.task_type in {"Cat1", "Cat2"}
        )
        filling_minutes = sum(
            max(task.machine_minutes, task.labor_minutes)
            for task in schedule.task_records
            if task.scheduled_date == current_day and task.task_type in {"Cat3", "Cat4"}
        )
        manufacturing_regular = (
            schedule.calendar_capacity.regular_labor_minutes("manufacturing", current_day, LABOR_CONFIG["manufacturing"]["regular_minutes"])
            if schedule.calendar_capacity
            else LABOR_CONFIG["manufacturing"]["regular_minutes"]
        )
        filling_regular = (
            schedule.calendar_capacity.regular_labor_minutes("filling", current_day, PACKAGING_TEAM["regular_minutes"])
            if schedule.calendar_capacity
            else PACKAGING_TEAM["regular_minutes"]
        )
        manufacturing_overtime = max(manufacturing_minutes - manufacturing_regular, 0.0)
        filling_overtime = max(filling_minutes - filling_regular, 0.0)
        is_holiday = schedule.calendar_capacity is not None and not schedule.calendar_capacity.is_workday(current_day)
        daily_load_rows.append(
            [
                current_day.isoformat(),
                weekday_names[current_day.weekday()],
                "假日" if is_holiday else "平日",
                _display_integer(manufacturing_minutes),
                _display_integer(filling_minutes),
                _display_integer(manufacturing_overtime),
                _display_integer(filling_overtime),
                _display_integer(manufacturing_overtime + filling_overtime),
            ]
        )

    write_simple_xlsx(
        output_dir / "schedule_check_report.xlsx",
        [
            ("訂單完成檢查", completed_rows),
            ("排程警告", warning_rows),
            ("訂單排除清單", excluded_rows),
            ("每日工作排程", daily_rows),
            ("每日工時", daily_load_rows),
        ],
    )


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
        premix_required_prefixes = tuple(
            f"{item.name}-{item.process_name}-"
            for item in order.manufacturing_items
            if item.premix_minutes_per_batch > 1e-2
        )
        for mix_task in (task for task in meaningful_tasks if task.task_type == "Cat2"):
            if not any(mix_task.detail.startswith(prefix) for prefix in premix_required_prefixes):
                continue
            mix_day = mix_task.scheduled_date
            expected_premix_day = (
                schedule.calendar_capacity.previous_workday(mix_day, order.raw_ready_date)
                if schedule.calendar_capacity is not None
                else mix_day - timedelta(days=1)
            )
            if expected_premix_day not in cat1_day_set:
                issues.append(
                    ScheduleValidationIssue(
                        order_id=order.order_id,
                        issue_type="premix_not_previous_day",
                        severity="error",
                        message=f"Mix day {mix_day.isoformat()} has no premix on the previous workday.",
                    )
                )
                break

        if expected_bottling_minutes > 1e-2:
            cat3_day_set = set(cat3_days)
            for fill_day in cat4_days:
                expected_bottling_day = (
                    schedule.calendar_capacity.previous_workday(fill_day, order.material_ready_date)
                    if schedule.calendar_capacity is not None
                    else fill_day - timedelta(days=1)
                )
                has_previous_day_bottling = expected_bottling_day in cat3_day_set
                has_same_day_bottling = fill_day in cat3_day_set and allows_same_day_bottling(order, fill_day)
                if not has_previous_day_bottling and not has_same_day_bottling:
                    issues.append(
                        ScheduleValidationIssue(
                            order_id=order.order_id,
                            issue_type="bottling_not_previous_day",
                            severity="error",
                            message=f"Fill day {fill_day.isoformat()} has no bottling on the previous workday.",
                        )
                    )
                    break

        if schedule.calendar_capacity is not None:
            for task in tasks:
                if not schedule.calendar_capacity.is_workday(task.scheduled_date):
                    issues.append(
                        ScheduleValidationIssue(
                            order_id=order.order_id,
                            issue_type="task_on_holiday",
                            severity="error",
                            message=f"{task.task_type} is scheduled on non-workday {task.scheduled_date.isoformat()}.",
                        )
                    )
                    break

        if len(cat4_days) > 1:
            expected_days = (
                schedule.calendar_capacity.workdays_between(cat4_days[0], cat4_days[-1])
                if schedule.calendar_capacity is not None
                else [cat4_days[0] + timedelta(days=offset) for offset in range((cat4_days[-1] - cat4_days[0]).days + 1)]
            )
            if set(cat4_days) != set(expected_days):
                issues.append(
                    ScheduleValidationIssue(
                        order_id=order.order_id,
                        issue_type="filling_days_not_contiguous",
                        severity="warning",
                        message="Order filling is split across non-contiguous workdays.",
                    )
                )

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

    def maximum_labor_minutes(labor_group: str, current_day: date) -> float:
        default_minutes = LABOR_CONFIG[labor_group]["max_minutes"]
        if schedule.calendar_capacity is None:
            return default_minutes
        return schedule.calendar_capacity.max_labor_minutes(labor_group, current_day, default_minutes)

    def maximum_packaging_minutes(current_day: date) -> float:
        if schedule.calendar_capacity is None:
            return PACKAGING_TEAM["max_minutes"]
        return schedule.calendar_capacity.max_labor_minutes(
            "filling",
            current_day,
            PACKAGING_TEAM["max_minutes"],
        )

    capacity_issue_count = 0
    for (labor_group, current_day), used_minutes in sorted(
        schedule.daily_labor_usage.items(),
        key=lambda item: (item[0][1], item[0][0]),
    ):
        if current_day == date.min:
            continue
        maximum = maximum_labor_minutes(labor_group, current_day)
        if used_minutes > maximum + 1e-6:
            capacity_issue_count += 1
            issues.append(
                ScheduleValidationIssue(
                    order_id="RESOURCE",
                    issue_type="labor_capacity_exceeded",
                    severity="error",
                    message=(
                        f"{current_day.isoformat()} {labor_group} 人力使用 {_display_integer(used_minutes)} 分鐘，"
                        f"超過最大容量 {_display_integer(maximum)} 分鐘。"
                    ),
                )
            )

    packaging_days = sorted(
        {
            current_day
            for labor_group, current_day in schedule.daily_labor_usage
            if labor_group in {"bottling", "filling"} and current_day != date.min
        }
    )
    for current_day in packaging_days:
        used_minutes = (
            schedule.daily_labor_usage.get(("bottling", current_day), 0.0)
            + schedule.daily_labor_usage.get(("filling", current_day), 0.0)
        )
        maximum = maximum_packaging_minutes(current_day)
        if used_minutes > maximum + 1e-6:
            capacity_issue_count += 1
            issues.append(
                ScheduleValidationIssue(
                    order_id="RESOURCE",
                    issue_type="packaging_team_capacity_exceeded",
                    severity="error",
                    message=(
                        f"{current_day.isoformat()} 充填團隊使用 {_display_integer(used_minutes)} 分鐘，"
                        f"超過最大容量 {_display_integer(maximum)} 分鐘。"
                    ),
                )
            )

    machine_usage = defaultdict(float, schedule.daily_machine_usage)
    for (current_day, machine_type), change_minutes in schedule.line_change_by_day_machine.items():
        machine_usage[(machine_type, current_day)] += change_minutes
    for (machine_type, current_day), used_minutes in sorted(
        machine_usage.items(),
        key=lambda item: (item[0][1], item[0][0]),
    ):
        is_workday = (
            schedule.calendar_capacity.is_workday(current_day)
            if schedule.calendar_capacity is not None
            else True
        )
        maximum = MACHINE_COUNTS.get(machine_type, 1) * MACHINE_DAILY_CAPACITY if is_workday else 0.0
        if used_minutes > maximum + 1e-6:
            capacity_issue_count += 1
            issues.append(
                ScheduleValidationIssue(
                    order_id="RESOURCE",
                    issue_type="machine_capacity_exceeded",
                    severity="error",
                    message=(
                        f"{current_day.isoformat()} {machine_type} 使用 {_display_integer(used_minutes)} 分鐘，"
                        f"超過最大容量 {_display_integer(maximum)} 分鐘。"
                    ),
                )
            )

    monthly_manufacturing_overtime: dict[tuple[int, int], float] = defaultdict(float)
    for (labor_group, current_day), used_minutes in schedule.daily_labor_usage.items():
        if labor_group != "manufacturing" or current_day == date.min:
            continue
        regular_minutes = (
            schedule.calendar_capacity.regular_labor_minutes(
                labor_group,
                current_day,
                LABOR_CONFIG[labor_group]["regular_minutes"],
            )
            if schedule.calendar_capacity is not None
            else LABOR_CONFIG[labor_group]["regular_minutes"]
        )
        monthly_manufacturing_overtime[(current_day.year, current_day.month)] += max(
            used_minutes - regular_minutes,
            0.0,
        )
    manufacturing_monthly_max = 2 * 46 * 60
    for (year, month), overtime_minutes in sorted(monthly_manufacturing_overtime.items()):
        if overtime_minutes > manufacturing_monthly_max + 1e-6:
            capacity_issue_count += 1
            issues.append(
                ScheduleValidationIssue(
                    order_id="RESOURCE",
                    issue_type="monthly_overtime_limit_exceeded",
                    severity="error",
                    message=(
                        f"{year:04d}-{month:02d} 製造加班 {_display_integer(overtime_minutes)} 分鐘，"
                        f"超過月上限 {_display_integer(manufacturing_monthly_max)} 分鐘。"
                    ),
                )
            )

    monthly_packaging_overtime: dict[tuple[int, int], float] = defaultdict(float)
    for current_day in packaging_days:
        used_minutes = (
            schedule.daily_labor_usage.get(("bottling", current_day), 0.0)
            + schedule.daily_labor_usage.get(("filling", current_day), 0.0)
        )
        regular_minutes = (
            schedule.calendar_capacity.regular_labor_minutes(
                "filling",
                current_day,
                PACKAGING_TEAM["regular_minutes"],
            )
            if schedule.calendar_capacity is not None
            else PACKAGING_TEAM["regular_minutes"]
        )
        monthly_packaging_overtime[(current_day.year, current_day.month)] += max(
            used_minutes - regular_minutes,
            0.0,
        )
    for (year, month), overtime_minutes in sorted(monthly_packaging_overtime.items()):
        if overtime_minutes > PACKAGING_TEAM["monthly_overtime_limit"] + 1e-6:
            capacity_issue_count += 1
            issues.append(
                ScheduleValidationIssue(
                    order_id="RESOURCE",
                    issue_type="monthly_overtime_limit_exceeded",
                    severity="error",
                    message=(
                        f"{year:04d}-{month:02d} 充填團隊加班 {_display_integer(overtime_minutes)} 分鐘，"
                        f"超過月上限 {_display_integer(PACKAGING_TEAM['monthly_overtime_limit'])} 分鐘。"
                    ),
                )
            )

    unallocated_by_pool_day: dict[tuple[date, str], list[IntradayTaskAllocation]] = defaultdict(list)
    for allocation in build_intraday_allocations(schedule):
        if not allocation.is_allocated:
            unallocated_by_pool_day[(allocation.task.scheduled_date, allocation.labor_pool)].append(allocation)
    for (current_day, labor_pool), allocations in sorted(unallocated_by_pool_day.items()):
        capacity_issue_count += 1
        failed_orders = sorted({allocation.task.order_id for allocation in allocations})
        displayed_orders = "、".join(failed_orders[:5])
        if len(failed_orders) > 5:
            displayed_orders += f" 等 {len(failed_orders)} 張"
        issues.append(
            ScheduleValidationIssue(
                order_id="RESOURCE",
                issue_type="intraday_resource_unallocatable",
                severity="error",
                message=(
                    f"{current_day.isoformat()} {_display_intraday_labor_pool(labor_pool)}無法配置全部確切時段；"
                    f"未配置訂單：{displayed_orders}。請重排後再作為派工表交付。"
                ),
            )
        )

    if schedule.infeasible_penalty > 1e-6 and capacity_issue_count == 0:
        issues.append(
            ScheduleValidationIssue(
                order_id="RESOURCE",
                issue_type="capacity_penalty_present",
                severity="error",
                message="排程仍含不可行懲罰，請檢查容量或凍結排程限制後再交付。",
            )
        )

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
                target = relation_map[relation_id]
                if target.startswith("/"):
                    return target.lstrip("/")
                if target.startswith("xl/"):
                    return target
                return "xl/" + target
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

    def __init__(
        self,
        workbook_path: Path,
        manufacturing_efficiency_path: Path | None = None,
        filling_efficiency_path: Path | None = None,
    ) -> None:
        self.reader = SimpleXlsxReader(workbook_path)
        self.sheet_names = self.reader.list_sheet_names()
        self.manufacturing_efficiency_path = manufacturing_efficiency_path
        self.filling_efficiency_path = filling_efficiency_path
        self.source_tables = self._load_source_tables()
        self.machine_counts = self._build_machine_counts()
        self.source_indexes = self._build_source_indexes()
        self.model_build_issues: list[dict[str, str]] = []

    def _load_source_tables(self) -> dict[str, list[dict[str, str]]]:
        source_sheet_map = {
            "time_ready": ("時間準備表V1",),
            "orders_2023": ("X2023訂單資訊含入庫日", "2023訂單資訊"),
            "bom": ("BOM",),
            "manufacturing_efficiency": ("製造效能_BOM回推", "製造效能"),
            "filling_efficiency": ("充填效能_BOM回推", "充填效能"),
            "machine_limits": ("機器限制",),
            "overtime_calc": ("加班計算",),
            "field_guide": ("欄位說明",),
        }
        tables: dict[str, list[dict[str, str]]] = {}
        for key, candidates in source_sheet_map.items():
            tables[key] = self._read_first_available_sheet(*candidates)
        if self.manufacturing_efficiency_path is not None:
            override_rows = self._read_first_available_sheet_from_path(
                self.manufacturing_efficiency_path,
                "製造效能_BOM回推",
                "製造效能",
            )
            if not override_rows:
                raise ValueError(
                    f"外部製造效能表必須包含「製造效能_BOM回推」或「製造效能」工作表："
                    f"{self.manufacturing_efficiency_path}"
                )
            tables["manufacturing_efficiency"] = override_rows
        if self.filling_efficiency_path is not None:
            override_rows = self._read_first_available_sheet_from_path(
                self.filling_efficiency_path,
                "充填效能_BOM回推",
                "充填效能",
            )
            if not override_rows:
                raise ValueError(
                    f"外部充填效能表必須包含「充填效能_BOM回推」或「充填效能」工作表："
                    f"{self.filling_efficiency_path}"
                )
            tables["filling_efficiency"] = override_rows
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

    @staticmethod
    def _read_first_available_sheet_from_path(path: Path, *sheet_names: str) -> list[dict[str, str]]:
        reader = SimpleXlsxReader(path)
        normalized_available = {
            SimpleXlsxReader._normalize_sheet_name(name): name for name in reader.list_sheet_names()
        }
        for sheet_name in sheet_names:
            actual_name = normalized_available.get(SimpleXlsxReader._normalize_sheet_name(sheet_name))
            if actual_name is not None:
                return reader.read_sheet(actual_name)
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
        self.model_build_issues = []
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
                missing_parts: list[str] = []
                if not manufacturing_items:
                    missing_parts.append("製造工序")
                if filling_item is None:
                    missing_parts.append("充填工序")
                self.model_build_issues.append(
                    {
                        "order_id": order_id,
                        "product_name": self._get_value(first, "中文名稱", "產品名稱"),
                        "issue_type": "missing_model_operation",
                        "message": f"無法建立{'與'.join(missing_parts)}；該訂單未進入排程。",
                    }
                )
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
            return "小乳機"
        if "100L真空乳化機" in normalized or "大乳機" in normalized:
            return "大乳機"
        if "大型攪拌機" in normalized or "精機" in normalized:
            return "精機"
        if "400L大桶槽" in normalized or "500L大桶槽" in normalized or "大機" in normalized:
            return "大機"
        return machine_name

    def _manufacturing_machine_from_source(self, source_row: dict[str, str]) -> str:
        process_equipment = self._get_value(source_row, "製程設備")
        if process_equipment in {"小乳", "小乳機"}:
            return "小乳機"
        if process_equipment in {"大乳", "大乳機"}:
            return "大乳機"
        if process_equipment in {"精", "精機"}:
            return "精機"
        if process_equipment in {"大", "大機"}:
            return "大機"

        machine_columns = (
            ("50L真空乳化機", "小乳機"),
            ("100L真空乳化機", "大乳機"),
            ("大型攪拌機", "精機"),
            ("400L大桶槽", "大機"),
            ("500L大桶槽", "大機"),
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
        bom_rows = (
            self._matched_source_rows(self._time_ready_row_code_candidates(rows[0]), "bom")
            if rows
            else []
        )
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
            item_name = self._get_value(group_rows[0], "項目", "項目類型") or process_name
            premix_minutes = sum(
                self._to_float(self._get_value(row, "工作(min)", "工作 min"))
                for row in group_rows
                if self._get_value(row, "Cat") == "Cat1"
            )
            if premix_minutes <= 0:
                # 有些整理表只留下 Cat2；若 BOM 明確有預攪，補回該必要工序。
                premix_minutes = sum(
                    self._to_float(self._get_value(row, "工作(min)", "工作 min"))
                    for row in bom_rows
                    if self._get_value(row, "工序", "工作") == "預攪"
                    and (
                        not item_name
                        or self._get_value(row, "項目", "項目類型") == item_name
                    )
                    and (
                        not process_name
                        or not self._get_value(row, "攪料程序")
                        or self._get_value(row, "攪料程序") == process_name
                    )
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

            if batches <= 0 or mix_clean_minutes <= 0 or not mix_machine:
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

        source_core_code = self._get_value(source_row, "核心六碼")
        manufacturing_efficiency_rows = self._matched_source_rows(
            self._core_code_candidates(source_core_code),
            "manufacturing_efficiency",
        )
        source_batches = self._to_int(self._get_value(source_row, "製造批次"))
        source_premix_total = self._to_float(self._get_value(source_row, "預攪總分鐘數"))
        source_mix_total = (
            self._to_float(self._get_value(source_row, "製造總時分鐘數*批次"))
            + self._to_float(self._get_value(source_row, "清洗總分鐘數"))
        )
        source_batches = source_batches or self._to_int(self._get_value(source_row, "需求批次"))
        source_premix_per_batch = self._to_float(self._get_value(source_row, "預攪"))
        source_mix_per_batch = self._to_float(self._get_value(source_row, "攪料"))
        source_clean_per_batch = self._to_float(self._get_value(source_row, "清洗"))
        if source_batches > 0 and source_premix_total <= 0 and source_premix_per_batch > 0:
            source_premix_total = source_premix_per_batch * source_batches
        if source_batches > 0 and source_mix_total <= 0 and (source_mix_per_batch > 0 or source_clean_per_batch > 0):
            source_mix_total = (source_mix_per_batch + source_clean_per_batch) * source_batches
        source_machine = self._manufacturing_machine_from_source(source_row)
        source_process_names = self._split_process_names(self._get_value(source_row, "攪料程序"))
        source_single_item = self._get_value(source_row, "單一製造項目")

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
        bom_rows = (
            self._matched_source_rows(self._time_ready_row_code_candidates(rows[0]), "bom")
            if rows
            else []
        )
        if not bottling_rows:
            bottling_rows = [
                row for row in bom_rows if self._get_value(row, "工序", "工作") == "排瓶"
            ]
        if not filling_rows:
            filling_rows = [
                row for row in bom_rows if self._get_value(row, "工序", "工作") == "充填"
            ]
        if not filling_rows:
            return None

        time_ready_bottling_minutes = sum(
            self._to_float(self._get_value(row, "充填類型工序耗時 (h)", "排瓶時數", "工作(min)", "工作 min"))
            for row in bottling_rows
        )
        time_ready_filling_minutes = sum(
            self._to_float(self._get_value(row, "充填類型工序耗時 (h)", "充填時數", "工作(min)", "工作 min"))
            for row in filling_rows
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

    def __init__(
        self,
        line_change_minutes: int = 0,
        frozen_schedule_entries: tuple[FrozenScheduleEntry, ...] = (),
        calendar_capacity: CalendarCapacity | None = None,
        independent_sections: bool = False,
        keep_split_days_contiguous: bool = True,
        break_start_time: str = DEFAULT_BREAK_START_TIME,
        break_end_time: str = DEFAULT_BREAK_END_TIME,
    ) -> None:
        self.line_change_minutes = max(0, int(line_change_minutes))
        self.frozen_schedule_entries = frozen_schedule_entries
        self.calendar_capacity = calendar_capacity
        self.independent_sections = independent_sections
        self.keep_split_days_contiguous = keep_split_days_contiguous
        self.break_start_time = break_start_time
        self.break_end_time = break_end_time
        self.break_slots = _break_slots_from_times(break_start_time, break_end_time)
        regular_slot_count = int(MACHINE_DAILY_CAPACITY // SLOT_MINUTES)
        self.intraday_slot_count = _slot_limit_for_work_slots(regular_slot_count, self.break_slots)
        self.frozen_dates = {
            (entry.order_id, entry.task_type): entry.scheduled_date
            for entry in frozen_schedule_entries
        }
        self.daily_labor_usage: dict[tuple[str, date], float] = defaultdict(float)
        self.daily_machine_usage: dict[tuple[str, date], float] = defaultdict(float)
        self.task_records: list[ScheduledTask] = []
        self.packaging_labor_slots: dict[date, list[int]] = defaultdict(lambda: [0] * self.intraday_slot_count)
        self.packaging_machine_slots: dict[tuple[date, str], list[int]] = defaultdict(
            lambda: [0] * self.intraday_slot_count
        )
        self.manufacturing_labor_slots: dict[date, list[int]] = defaultdict(
            lambda: [0] * self.intraday_slot_count
        )
        self.manufacturing_machine_slots: dict[tuple[date, str], list[int]] = defaultdict(
            lambda: [0] * self.intraday_slot_count
        )

    @staticmethod
    def _assigned_bottling_people(filling_item: FillingItem) -> int:
        if filling_item.bottling_machine_minutes <= 0:
            return 1
        return min(max(filling_item.bottling_people, BOTTLING_MAX_PEOPLE), BOTTLING_MAX_PEOPLE)

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

        line_change_by_day_machine = self._line_change_by_day_machine()
        overtime_by_group = self._overtime_by_group(line_change_by_day_machine)
        filling_noncontiguous_order_count = self._filling_noncontiguous_order_count()

        result = ScheduleResult(
            total_overtime_minutes=self._total_overtime_minutes(line_change_by_day_machine),
            infeasible_penalty=self._capacity_penalty(line_change_by_day_machine),
            target_penalty=self._target_penalty(overtime_by_group, filling_noncontiguous_order_count),
            filling_noncontiguous_order_count=filling_noncontiguous_order_count,
            overtime_by_group=overtime_by_group,
            weekend_overtime_minutes=self._weekend_overtime_minutes(line_change_by_day_machine),
            line_change_minutes=sum(line_change_by_day_machine.values()),
            line_change_minutes_per_switch=self.line_change_minutes,
            line_change_by_day_machine=dict(line_change_by_day_machine),
            break_start_time=self.break_start_time,
            break_end_time=self.break_end_time,
            frozen_schedule_entries=self.frozen_schedule_entries,
            calendar_capacity=self.calendar_capacity,
            task_records=sorted(
                self.task_records,
                key=lambda item: (item.scheduled_date, item.order_id, item.task_type, item.detail),
            ),
            daily_labor_usage=dict(self.daily_labor_usage),
            daily_machine_usage=dict(self.daily_machine_usage),
        )
        result.intraday_unallocated_task_count = sum(
            1 for allocation in build_intraday_allocations(result) if not allocation.is_allocated
        )
        result.infeasible_penalty += result.intraday_unallocated_task_count * self.infeasible_weight
        return result

    def _schedule_order(self, order: Order) -> None:
        labor_snapshot = dict(self.daily_labor_usage)
        machine_snapshot = dict(self.daily_machine_usage)
        task_count_snapshot = len(self.task_records)
        packaging_labor_snapshot = {
            current_day: slots[:]
            for current_day, slots in self.packaging_labor_slots.items()
        }
        packaging_machine_snapshot = {
            key: slots[:]
            for key, slots in self.packaging_machine_slots.items()
        }
        manufacturing_labor_snapshot = {
            current_day: slots[:]
            for current_day, slots in self.manufacturing_labor_slots.items()
        }
        manufacturing_machine_snapshot = {
            key: slots[:]
            for key, slots in self.manufacturing_machine_slots.items()
        }

        first_fill_day = self._schedule_filling(order)
        if first_fill_day is None:
            self._restore_snapshot(
                labor_snapshot,
                machine_snapshot,
                task_count_snapshot,
                packaging_labor_snapshot,
                packaging_machine_snapshot,
                manufacturing_labor_snapshot,
                manufacturing_machine_snapshot,
            )
            first_fill_day = self._force_schedule_filling(order)

        earliest_manufacturing_day = max(
            order.raw_ready_date,
            order.due_date - timedelta(days=30),
        ) if self.independent_sections else order.raw_ready_date
        latest_mix_day = (
            order.due_date
            if self.independent_sections
            else first_fill_day - timedelta(days=1)
        )

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
                    self._force_schedule_manufacturing_batch(
                        order=order,
                        item=item,
                        batch_index=batch_index,
                        earliest_manufacturing_day=earliest_manufacturing_day,
                        latest_mix_day=latest_mix_day,
                    )

    def _restore_snapshot(
        self,
        labor_snapshot: dict[tuple[str, date], float],
        machine_snapshot: dict[tuple[str, date], float],
        task_count_snapshot: int,
        packaging_labor_snapshot: dict[date, list[int]],
        packaging_machine_snapshot: dict[tuple[date, str], list[int]],
        manufacturing_labor_snapshot: dict[date, list[int]],
        manufacturing_machine_snapshot: dict[tuple[date, str], list[int]],
    ) -> None:
        self.daily_labor_usage = defaultdict(float, labor_snapshot)
        self.daily_machine_usage = defaultdict(float, machine_snapshot)
        self.packaging_labor_slots = defaultdict(
            lambda: [0] * self.intraday_slot_count,
            {current_day: slots[:] for current_day, slots in packaging_labor_snapshot.items()},
        )
        self.packaging_machine_slots = defaultdict(
            lambda: [0] * self.intraday_slot_count,
            {key: slots[:] for key, slots in packaging_machine_snapshot.items()},
        )
        self.manufacturing_labor_slots = defaultdict(
            lambda: [0] * self.intraday_slot_count,
            {current_day: slots[:] for current_day, slots in manufacturing_labor_snapshot.items()},
        )
        self.manufacturing_machine_slots = defaultdict(
            lambda: [0] * self.intraday_slot_count,
            {key: slots[:] for key, slots in manufacturing_machine_snapshot.items()},
        )
        del self.task_records[task_count_snapshot:]

    def _resolve_bottling_day(
        self,
        order: Order,
        fill_day: date,
        frozen_bottling_day: date | None,
    ) -> date | None:
        if frozen_bottling_day is not None:
            if frozen_bottling_day < order.material_ready_date:
                return None
            if not self._is_workday(frozen_bottling_day):
                return None
            if frozen_bottling_day == fill_day - timedelta(days=1):
                return frozen_bottling_day
            if frozen_bottling_day == fill_day and allows_same_day_bottling(order, fill_day):
                return frozen_bottling_day
            return None

        bottling_day = fill_day - timedelta(days=1)
        if self.calendar_capacity is not None:
            bottling_day = self.calendar_capacity.previous_workday(fill_day, order.material_ready_date)
            if bottling_day is None:
                if allows_same_day_bottling(order, fill_day) and self._is_workday(fill_day):
                    return fill_day
                return None
        if bottling_day < order.material_ready_date:
            if allows_same_day_bottling(order, fill_day):
                return fill_day
            return None
        return bottling_day

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
        selected_fill_days: set[date] = set()
        frozen_fill_day = self._frozen_date(order.order_id, "Cat4")
        frozen_bottling_day = self._frozen_date(order.order_id, "Cat3")
        if frozen_fill_day is not None:
            if frozen_fill_day < order.material_ready_date or frozen_fill_day > order.due_date or not self._is_workday(frozen_fill_day):
                return None
            candidate_days = [frozen_fill_day]
        else:
            candidate_days = self._date_range(order.material_ready_date, order.due_date)

        for use_max_capacity in (False, True):
            while remaining_fill_minutes > self.minute_epsilon:
                candidate_options: list[tuple[float, int, float, date, float]] = []
                for current_day in reversed(candidate_days):
                    bottling_day = self._resolve_bottling_day(order, current_day, frozen_bottling_day)
                    if bottling_day is None:
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
                        added_overtime + self._filling_day_selection_penalty(selected_fill_days, current_day),
                        -current_day.toordinal(),
                        -chunk,
                        current_day,
                        chunk,
                    )
                    candidate_options.append(candidate)

                if not candidate_options:
                    break

                selected_option: tuple[date, date, float] | None = None
                for _overtime, _reverse_day, _negative_chunk, current_day, candidate_chunk in sorted(candidate_options):
                    bottling_day = self._resolve_bottling_day(order, current_day, frozen_bottling_day)
                    if bottling_day is None:
                        continue
                    chunk = self._max_intraday_feasible_fill_chunk(
                        current_day=current_day,
                        bottling_day=bottling_day,
                        maximum_chunk=candidate_chunk,
                        filling_item=filling_item,
                        bottling_people=assigned_bottling_people,
                        fill_to_bottling_ratio=fill_to_bottling_ratio,
                    )
                    if chunk > self.minute_epsilon:
                        selected_option = (current_day, bottling_day, chunk)
                        break

                if selected_option is None:
                    break
                current_day, bottling_day, chunk = selected_option
                self._commit_filling_chunk(
                    order,
                    current_day,
                    bottling_day,
                    chunk,
                    filling_item,
                    assigned_bottling_people,
                    fill_to_bottling_ratio,
                )
                selected_fill_days.add(current_day)
                remaining_fill_minutes -= chunk
                if remaining_fill_minutes < self.minute_epsilon:
                    remaining_fill_minutes = 0.0
                if first_fill_day is None or current_day < first_fill_day:
                    first_fill_day = current_day

        if remaining_fill_minutes > self.minute_epsilon:
            return None
        return first_fill_day

    def _filling_day_selection_penalty(
        self,
        selected_days: set[date],
        candidate_day: date,
    ) -> float:
        if not self.keep_split_days_contiguous or not selected_days or candidate_day in selected_days:
            return 0.0

        def missing_days(days: set[date]) -> int:
            expected_days = set(self._date_range(min(days), max(days)))
            return len(expected_days - days)

        before_missing = missing_days(selected_days)
        after_missing = missing_days(selected_days | {candidate_day})
        return (after_missing - before_missing) * FILLING_NONCONTIGUOUS_ORDER_WEIGHT

    def _force_schedule_filling(self, order: Order) -> date:
        filling_item = order.filling_item
        frozen_fill_day = self._frozen_date(order.order_id, "Cat4")
        fill_day = frozen_fill_day or self._latest_workday(order.material_ready_date, order.due_date)
        if fill_day is None:
            fill_day = order.due_date
        frozen_bottling_day = self._frozen_date(order.order_id, "Cat3")
        bottling_day = frozen_bottling_day or self._resolve_bottling_day(order, fill_day, None)
        if bottling_day is None:
            bottling_day = fill_day if allows_same_day_bottling(order, fill_day) else fill_day - timedelta(days=1)
        self._commit_filling_chunk(
            order=order,
            current_day=fill_day,
            bottling_day=bottling_day,
            fill_machine_minutes=filling_item.filling_machine_minutes,
            filling_item=filling_item,
            bottling_people=self._assigned_bottling_people(filling_item),
            fill_to_bottling_ratio=(
                filling_item.bottling_machine_minutes
                * filling_item.bottling_people
                / self._assigned_bottling_people(filling_item)
                / filling_item.filling_machine_minutes
                if filling_item.filling_machine_minutes > 0
                else 0.0
            ),
        )
        return fill_day

    def _schedule_manufacturing_batch(
        self,
        order: Order,
        item: ManufacturingItem,
        batch_index: int,
        earliest_manufacturing_day: date,
        latest_mix_day: date,
    ) -> bool:
        requires_premix = item.premix_minutes_per_batch > self.minute_epsilon
        frozen_premix_day = self._frozen_date(order.order_id, "Cat1")
        frozen_mix_day = self._frozen_date(order.order_id, "Cat2")
        if frozen_premix_day is not None or frozen_mix_day is not None:
            if not requires_premix and frozen_premix_day is not None:
                return False
            mix_day = frozen_mix_day or (frozen_premix_day + timedelta(days=1))  # type: ignore[operator]
            premix_day = (
                self._previous_workday(mix_day, earliest_manufacturing_day) or (mix_day - timedelta(days=1))
                if requires_premix
                else None
            )
            if frozen_premix_day is not None and premix_day != frozen_premix_day:
                return False
            first_mix_day = earliest_manufacturing_day + timedelta(days=1) if requires_premix else earliest_manufacturing_day
            if mix_day < first_mix_day or mix_day > latest_mix_day:
                return False
            if self._added_overtime_for_manufacturing(item, premix_day, mix_day) is None:
                return False
            intraday_tasks = [(mix_day, item.mix_machine, item.mix_clean_minutes_per_batch)]
            if requires_premix and premix_day is not None:
                intraday_tasks.insert(0, (premix_day, "預攪機", item.premix_minutes_per_batch))
            if not self._manufacturing_tasks_are_intraday_feasible(intraday_tasks, commit=True):
                return False
            if requires_premix and premix_day is not None:
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
                scheduled_date=mix_day,
                machine_type=item.mix_machine,
                machine_minutes=item.mix_clean_minutes_per_batch,
                labor_group="manufacturing",
                labor_minutes=item.mix_clean_minutes_per_batch,
            )
            return True

        best_day: date | None = None
        best_increment = math.inf

        first_mix_day = earliest_manufacturing_day + timedelta(days=1) if requires_premix else earliest_manufacturing_day
        for mix_day in reversed(self._date_range(first_mix_day, latest_mix_day)):
            premix_day = self._previous_workday(mix_day, earliest_manufacturing_day) if requires_premix else None
            if requires_premix and premix_day is None:
                continue
            increment = self._added_overtime_for_manufacturing(item, premix_day, mix_day)
            if increment is None:
                continue
            intraday_tasks = [(mix_day, item.mix_machine, item.mix_clean_minutes_per_batch)]
            if requires_premix and premix_day is not None:
                intraday_tasks.insert(0, (premix_day, "預攪機", item.premix_minutes_per_batch))
            if not self._manufacturing_tasks_are_intraday_feasible(intraday_tasks):
                continue
            if increment < best_increment:
                best_increment = increment
                best_day = mix_day

        if best_day is None:
            return False

        premix_day = (
            self._previous_workday(best_day, earliest_manufacturing_day) or (best_day - timedelta(days=1))
            if requires_premix
            else None
        )
        intraday_tasks = [(best_day, item.mix_machine, item.mix_clean_minutes_per_batch)]
        if requires_premix and premix_day is not None:
            intraday_tasks.insert(0, (premix_day, "預攪機", item.premix_minutes_per_batch))
        self._manufacturing_tasks_are_intraday_feasible(intraday_tasks, commit=True)
        if requires_premix and premix_day is not None:
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

    def _force_schedule_manufacturing_batch(
        self,
        order: Order,
        item: ManufacturingItem,
        batch_index: int,
        earliest_manufacturing_day: date,
        latest_mix_day: date,
    ) -> None:
        requires_premix = item.premix_minutes_per_batch > self.minute_epsilon
        first_mix_day = earliest_manufacturing_day + timedelta(days=1) if requires_premix else earliest_manufacturing_day
        mix_day = self._frozen_date(order.order_id, "Cat2") or self._latest_workday(
            first_mix_day,
            latest_mix_day,
        )
        if mix_day is None:
            mix_day = max(first_mix_day, latest_mix_day)
        premix_day: date | None = None
        if requires_premix:
            premix_day = self._frozen_date(order.order_id, "Cat1") or self._previous_workday(
                mix_day,
                earliest_manufacturing_day,
            )
            if premix_day is None:
                for candidate_mix_day in self._date_range(first_mix_day, latest_mix_day):
                    candidate_premix_day = self._previous_workday(
                        candidate_mix_day,
                        earliest_manufacturing_day,
                    )
                    if candidate_premix_day is not None:
                        mix_day = candidate_mix_day
                        premix_day = candidate_premix_day
                        break
            if premix_day is None:
                premix_day = self._previous_workday(mix_day, earliest_manufacturing_day - timedelta(days=14))
            if premix_day is None:
                premix_day = mix_day - timedelta(days=1)
            self._manufacturing_tasks_are_intraday_feasible(
                [(premix_day, "預攪機", item.premix_minutes_per_batch)],
                commit=True,
            )
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
        self._manufacturing_tasks_are_intraday_feasible(
            [(mix_day, item.mix_machine, item.mix_clean_minutes_per_batch)],
            commit=True,
        )
        self._commit_task(
            order_id=order.order_id,
            task_type="Cat2",
            detail=f"{item.name}-{item.process_name}-第{batch_index}批攪料清洗",
            scheduled_date=mix_day,
            machine_type=item.mix_machine,
            machine_minutes=item.mix_clean_minutes_per_batch,
            labor_group="manufacturing",
            labor_minutes=item.mix_clean_minutes_per_batch,
        )

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
        bottling_work_cap = self._remaining_bottling_work_capacity(bottling_day, use_max_capacity)
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
            candidates.append(bottling_work_cap / (fill_to_bottling_ratio * max(bottling_people, 1)))
            candidates.append(
                packaging_labor_cap
                / (fill_people + fill_to_bottling_ratio * max(bottling_people, 1))
            )
        else:
            candidates.append(packaging_labor_cap / max(fill_people, 1))
        return max(min(candidates), 0.0)

    def _max_intraday_feasible_fill_chunk(
        self,
        current_day: date,
        bottling_day: date,
        maximum_chunk: float,
        filling_item: FillingItem,
        bottling_people: int,
        fill_to_bottling_ratio: float,
    ) -> float:
        if maximum_chunk <= self.minute_epsilon:
            return 0.0
        if self._packaging_chunk_is_intraday_feasible(
            current_day,
            bottling_day,
            maximum_chunk,
            filling_item,
            bottling_people,
            fill_to_bottling_ratio,
        ):
            return maximum_chunk

        maximum_slots = int(math.floor(maximum_chunk / SLOT_MINUTES + 1e-9))
        low, high = 1, maximum_slots
        best_chunk = 0.0
        while low <= high:
            middle = (low + high) // 2
            candidate_chunk = middle * SLOT_MINUTES
            if self._packaging_chunk_is_intraday_feasible(
                current_day,
                bottling_day,
                candidate_chunk,
                filling_item,
                bottling_people,
                fill_to_bottling_ratio,
            ):
                best_chunk = candidate_chunk
                low = middle + 1
            else:
                high = middle - 1
        return best_chunk

    def _packaging_chunk_is_intraday_feasible(
        self,
        current_day: date,
        bottling_day: date,
        fill_machine_minutes: float,
        filling_item: FillingItem,
        bottling_people: int,
        fill_to_bottling_ratio: float,
        commit: bool = False,
    ) -> bool:
        bottling_machine_minutes = fill_machine_minutes * fill_to_bottling_ratio
        affected_days = {current_day, bottling_day}
        labor_slots = {
            day: self.packaging_labor_slots[day][:]
            for day in affected_days
        }
        machine_keys = {
            (bottling_day, filling_item.bottling_machine),
            (current_day, filling_item.filling_machine),
        }
        machine_slots = {
            key: self.packaging_machine_slots[key][:]
            for key in machine_keys
        }

        def place_task(
            scheduled_day: date,
            machine_type: str,
            machine_minutes: float,
            people: int,
            earliest_start: int = 0,
            prefer_late: bool = False,
        ) -> int | None:
            required_slots = self._minutes_to_slots(machine_minutes)
            if required_slots <= 0:
                return earliest_start - 1
            day_config = self.calendar_capacity.days.get(scheduled_day) if self.calendar_capacity is not None else None
            available_people = (
                day_config.filling_people
                if day_config is not None
                else LABOR_CONFIG["filling"]["people"]
            )
            if people > available_people or not self._is_workday(scheduled_day):
                return None
            slot_limit = self.intraday_slot_count
            machine_count = max(MACHINE_COUNTS.get(machine_type, 1), 1)
            latest_start = slot_limit - 1
            candidate_starts = (
                range(latest_start, earliest_start - 1, -1)
                if prefer_late
                else range(earliest_start, latest_start + 1)
            )
            for start_slot in candidate_starts:
                occupied_slots = _work_slots_from_start(
                    start_slot,
                    required_slots,
                    slot_limit,
                    self.break_slots,
                )
                if occupied_slots is not None and all(
                    labor_slots[scheduled_day][slot] + people <= available_people
                    and machine_slots[(scheduled_day, machine_type)][slot] < machine_count
                    for slot in occupied_slots
                ):
                    for slot in occupied_slots:
                        labor_slots[scheduled_day][slot] += people
                        machine_slots[(scheduled_day, machine_type)][slot] += 1
                    return occupied_slots[-1]
            return None

        earliest_fill_slot = 0
        if bottling_machine_minutes > self.minute_epsilon:
            bottling_end = place_task(
                bottling_day,
                filling_item.bottling_machine,
                bottling_machine_minutes,
                bottling_people,
                prefer_late=bottling_day != current_day,
            )
            if bottling_end is None:
                return False
            if bottling_day == current_day:
                earliest_fill_slot = bottling_end + 1
        if place_task(
            current_day,
            filling_item.filling_machine,
            fill_machine_minutes,
            filling_item.filling_people,
            earliest_start=earliest_fill_slot,
        ) is None:
            return False

        if commit:
            for day, slots in labor_slots.items():
                self.packaging_labor_slots[day] = slots
            for key, slots in machine_slots.items():
                self.packaging_machine_slots[key] = slots
        return True

    def _manufacturing_tasks_are_intraday_feasible(
        self,
        tasks: list[tuple[date, str, float]],
        commit: bool = False,
    ) -> bool:
        affected_days = {scheduled_day for scheduled_day, _machine_type, _minutes in tasks}
        labor_slots = {
            scheduled_day: self.manufacturing_labor_slots[scheduled_day][:]
            for scheduled_day in affected_days
        }
        machine_keys = {(scheduled_day, machine_type) for scheduled_day, machine_type, _minutes in tasks}
        machine_slots = {
            key: self.manufacturing_machine_slots[key][:]
            for key in machine_keys
        }
        slot_limit = self.intraday_slot_count

        for scheduled_day, machine_type, machine_minutes in tasks:
            required_slots = self._minutes_to_slots(machine_minutes)
            if required_slots <= 0:
                continue
            day_config = self.calendar_capacity.days.get(scheduled_day) if self.calendar_capacity is not None else None
            available_people = (
                day_config.manufacturing_people
                if day_config is not None
                else LABOR_CONFIG["manufacturing"]["people"]
            )
            if available_people < 1 or not self._is_workday(scheduled_day):
                return False
            machine_count = max(MACHINE_COUNTS.get(machine_type, 1), 1)
            placed = False
            for start_slot in range(0, slot_limit):
                occupied_slots = _work_slots_from_start(
                    start_slot,
                    required_slots,
                    slot_limit,
                    self.break_slots,
                )
                if occupied_slots is not None and all(
                    labor_slots[scheduled_day][slot] + 1 <= available_people
                    and machine_slots[(scheduled_day, machine_type)][slot] < machine_count
                    for slot in occupied_slots
                ):
                    for slot in occupied_slots:
                        labor_slots[scheduled_day][slot] += 1
                        machine_slots[(scheduled_day, machine_type)][slot] += 1
                    placed = True
                    break
            if not placed:
                return False

        if commit:
            for scheduled_day, slots in labor_slots.items():
                self.manufacturing_labor_slots[scheduled_day] = slots
            for key, slots in machine_slots.items():
                self.manufacturing_machine_slots[key] = slots
        return True

    def _added_overtime_for_manufacturing(
        self,
        item: ManufacturingItem,
        premix_day: date | None,
        mix_day: date,
    ) -> float | None:
        if (
            premix_day is not None
            and self._remaining_machine_capacity("預攪機", premix_day) < item.premix_minutes_per_batch
        ):
            return None
        if self._remaining_machine_capacity(item.mix_machine, mix_day) < item.mix_clean_minutes_per_batch:
            return None

        premix_increment = (
            self._incremental_labor_overtime(
                labor_group="manufacturing",
                current_day=premix_day,
                add_minutes=item.premix_minutes_per_batch,
            )
            if premix_day is not None
            else 0.0
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
        bottling_machine_increment = self._incremental_machine_overtime(
            machine_type=filling_item.bottling_machine,
            current_day=bottling_day,
            add_minutes=bottling_minutes,
        )
        if bottling_machine_increment is None:
            return None

        bottling_increment = self._incremental_bottling_work_overtime(
            current_day=bottling_day,
            add_minutes=bottling_minutes * bottling_people,
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

    def _frozen_date(self, order_id: str, task_type: str) -> date | None:
        return self.frozen_dates.get((order_id, task_type))

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
        self._packaging_chunk_is_intraday_feasible(
            current_day=current_day,
            bottling_day=bottling_day,
            fill_machine_minutes=fill_machine_minutes,
            filling_item=filling_item,
            bottling_people=bottling_people,
            fill_to_bottling_ratio=fill_to_bottling_ratio,
            commit=True,
        )

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
        capacity = (
            self._max_labor_minutes(labor_group, current_day)
            if use_max_capacity
            else self._regular_labor_minutes(labor_group, current_day)
        )
        return max(capacity - used, 0.0)

    def _remaining_machine_capacity(self, machine_type: str, current_day: date) -> float:
        if not self._is_workday(current_day):
            return 0.0
        used = self.daily_machine_usage[(machine_type, current_day)]
        machine_count = MACHINE_COUNTS.get(machine_type, 1)
        return max(machine_count * MACHINE_DAILY_CAPACITY - used, 0.0)

    def _remaining_machine_capacity_for_phase(
        self,
        machine_type: str,
        current_day: date,
        use_max_capacity: bool,
    ) -> float:
        if not self._is_workday(current_day):
            return 0.0
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
        if not self._is_workday(current_day):
            return None
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
        capacity = self._packaging_max_minutes(current_day) if use_max_capacity else self._packaging_regular_minutes(current_day)
        return max(capacity - self._packaging_labor_used(current_day), 0.0)

    def _remaining_bottling_work_capacity(self, current_day: date, use_max_capacity: bool) -> float:
        capacity = self._max_labor_minutes("bottling", current_day) if use_max_capacity else self._bottling_regular_minutes(current_day)
        used = self.daily_labor_usage[("bottling", current_day)]
        return max(capacity - used, 0.0)

    def _incremental_bottling_work_overtime(self, current_day: date, add_minutes: float) -> float | None:
        used = self.daily_labor_usage[("bottling", current_day)]
        max_minutes = self._max_labor_minutes("bottling", current_day)
        regular_minutes = self._bottling_regular_minutes(current_day)
        if used + add_minutes > max_minutes + 1e-6:
            return None
        before = max(used - regular_minutes, 0.0)
        after = max(used + add_minutes - regular_minutes, 0.0)
        return after - before

    def _incremental_packaging_overtime(
        self,
        current_day: date,
        bottling_labor_minutes: float,
        filling_labor_minutes: float,
    ) -> float | None:
        used = self._packaging_labor_used(current_day)
        add_minutes = bottling_labor_minutes + filling_labor_minutes
        if used + add_minutes > self._packaging_max_minutes(current_day) + 1e-6:
            return None
        before = max(used - self._packaging_regular_minutes(current_day), 0.0)
        after = max(used + add_minutes - self._packaging_regular_minutes(current_day), 0.0)
        return after - before

    def _incremental_labor_overtime(
        self,
        labor_group: str,
        current_day: date,
        add_minutes: float,
    ) -> float | None:
        config = LABOR_CONFIG[labor_group]
        used = self.daily_labor_usage[(labor_group, current_day)]
        max_minutes = self._max_labor_minutes(labor_group, current_day)
        regular_minutes = self._regular_labor_minutes(labor_group, current_day)
        if used + add_minutes > max_minutes + 1e-6:
            return None
        before = max(used - regular_minutes, 0.0)
        after = max(used + add_minutes - regular_minutes, 0.0)
        return after - before

    def _is_workday(self, current_day: date) -> bool:
        if self.calendar_capacity is None:
            return True
        return self.calendar_capacity.is_workday(current_day)

    def _previous_workday(self, before_day: date, min_day: date) -> date | None:
        if self.calendar_capacity is not None:
            return self.calendar_capacity.previous_workday(before_day, min_day)
        candidate = before_day - timedelta(days=1)
        return candidate if candidate >= min_day else None

    def _latest_workday(self, start_day: date, end_day: date) -> date | None:
        for current_day in reversed(self._date_range(start_day, end_day)):
            return current_day
        return None

    def _regular_labor_minutes(self, labor_group: str, current_day: date) -> float:
        default_minutes = LABOR_CONFIG[labor_group]["regular_minutes"]
        if self.calendar_capacity is None:
            return default_minutes
        return self.calendar_capacity.regular_labor_minutes(labor_group, current_day, default_minutes)

    def _max_labor_minutes(self, labor_group: str, current_day: date) -> float:
        default_minutes = LABOR_CONFIG[labor_group]["max_minutes"]
        if self.calendar_capacity is None:
            return default_minutes
        return self.calendar_capacity.max_labor_minutes(labor_group, current_day, default_minutes)

    def _bottling_regular_minutes(self, current_day: date) -> float:
        if not self._is_workday(current_day):
            return 0.0
        return PROCESS_OVERTIME_REGULAR_MINUTES["bottling"]

    def _packaging_regular_minutes(self, current_day: date) -> float:
        if self.calendar_capacity is None:
            return PACKAGING_TEAM["regular_minutes"]
        return self.calendar_capacity.regular_labor_minutes(
            "filling",
            current_day,
            PACKAGING_TEAM["regular_minutes"],
        )

    def _packaging_max_minutes(self, current_day: date) -> float:
        if self.calendar_capacity is None:
            return PACKAGING_TEAM["max_minutes"]
        return self.calendar_capacity.max_labor_minutes(
            "filling",
            current_day,
            PACKAGING_TEAM["max_minutes"],
        )

    def _line_change_by_day_machine(self) -> dict[tuple[date, str], float]:
        if self.line_change_minutes <= 0:
            return {}

        machine_day_products: dict[tuple[date, str], set[str]] = defaultdict(set)
        for task in self.task_records:
            if task.task_type != "Cat4" or task.scheduled_date == date.min:
                continue
            machine_day_products[(task.scheduled_date, task.machine_type)].add(_filling_product_key(task))

        return {
            key: max(len(products) - 1, 0) * self.line_change_minutes
            for key, products in machine_day_products.items()
            if len(products) > 1
        }

    def _process_overtime_by_group(
        self,
        line_change_by_day_machine: Mapping[tuple[date, str], float] | None = None,
    ) -> dict[str, float]:
        overtime = {process: 0.0 for process in PROCESS_OVERTIME_MACHINES}
        process_usage: dict[tuple[str, date], float] = defaultdict(float)
        for task in self.task_records:
            if task.task_type == "Cat3":
                process_usage[("bottling", task.scheduled_date)] += task.labor_minutes
            elif task.task_type == "Cat4":
                process_usage[("filling", task.scheduled_date)] += task.machine_minutes
        for (current_day, _machine_type), line_change_minutes in (line_change_by_day_machine or {}).items():
            process_usage[("filling", current_day)] += line_change_minutes

        for (process, current_day), used_minutes in process_usage.items():
            if current_day == date.min:
                continue
            regular_minutes = PROCESS_OVERTIME_REGULAR_MINUTES.get(
                process,
                PROCESS_DAILY_REGULAR_MINUTES,
            )
            if self.calendar_capacity is not None and not self.calendar_capacity.is_workday(current_day):
                regular_minutes = 0.0
            overtime[process] += max(used_minutes - regular_minutes, 0.0)
        return overtime

    def _total_overtime_minutes(
        self,
        line_change_by_day_machine: Mapping[tuple[date, str], float] | None = None,
    ) -> float:
        total = 0.0
        for (labor_group, current_day), used_minutes in self.daily_labor_usage.items():
            if current_day == date.min or labor_group in {"bottling", "filling"}:
                continue
            regular = self._regular_labor_minutes(labor_group, current_day)
            total += max(used_minutes - regular, 0.0)
        process_overtime = self._process_overtime_by_group(line_change_by_day_machine)
        total += sum(process_overtime.values())
        return total

    def _overtime_by_group(
        self,
        line_change_by_day_machine: Mapping[tuple[date, str], float] | None = None,
    ) -> dict[str, float]:
        overtime_by_group = {"manufacturing": 0.0, "filling": 0.0, "bottling": 0.0}
        for (labor_group, current_day), used_minutes in self.daily_labor_usage.items():
            if current_day == date.min or labor_group in {"bottling", "filling"}:
                continue
            regular = self._regular_labor_minutes(labor_group, current_day)
            overtime_by_group[labor_group] += max(used_minutes - regular, 0.0)
        process_overtime = self._process_overtime_by_group(line_change_by_day_machine)
        overtime_by_group["bottling"] = process_overtime["bottling"]
        overtime_by_group["filling"] = process_overtime["filling"]
        return overtime_by_group

    def _weekend_overtime_minutes(
        self,
        line_change_by_day_machine: Mapping[tuple[date, str], float] | None = None,
    ) -> float:
        total = 0.0
        for (labor_group, current_day), used_minutes in self.daily_labor_usage.items():
            if current_day == date.min or current_day.weekday() < 5 or labor_group in {"bottling", "filling"}:
                continue
            regular = self._regular_labor_minutes(labor_group, current_day)
            total += max(used_minutes - regular, 0.0)
        process_usage: dict[tuple[str, date], float] = defaultdict(float)
        for task in self.task_records:
            if task.task_type == "Cat3":
                process_usage[("bottling", task.scheduled_date)] += task.labor_minutes
            elif task.task_type == "Cat4":
                process_usage[("filling", task.scheduled_date)] += task.machine_minutes
        for (current_day, _machine_type), line_change_minutes in (line_change_by_day_machine or {}).items():
            process_usage[("filling", current_day)] += line_change_minutes
        for (process, current_day), used_minutes in process_usage.items():
            if current_day == date.min or current_day.weekday() < 5:
                continue
            regular_minutes = PROCESS_OVERTIME_REGULAR_MINUTES.get(
                process,
                PROCESS_DAILY_REGULAR_MINUTES,
            )
            if self.calendar_capacity is not None and not self.calendar_capacity.is_workday(current_day):
                regular_minutes = 0.0
            total += max(used_minutes - regular_minutes, 0.0)
        return total

    def _manufacturing_switch_cleaning_penalty(self) -> float:
        machine_day_processes: dict[tuple[date, str], set[str]] = defaultdict(set)
        for task in self.task_records:
            if task.task_type not in {"Cat1", "Cat2"}:
                continue
            if task.machine_type == "預攪機":
                continue
            process_key = task.detail.split("-第", 1)[0]
            machine_day_processes[(task.scheduled_date, task.machine_type)].add(process_key)

        penalty = 0.0
        for process_keys in machine_day_processes.values():
            normalized_process_keys = {
                _normalize_manufacturing_detail(process_key) for process_key in process_keys
            }
            switch_count = max(len(normalized_process_keys) - 1, 0)
            penalty += switch_count * MANUFACTURING_SWITCH_CLEANING_WEIGHT
        return penalty

    def _filling_noncontiguous_order_count(self) -> int:
        filling_days_by_order: dict[str, set[date]] = defaultdict(set)
        for task in self.task_records:
            if task.task_type == "Cat4":
                filling_days_by_order[task.order_id].add(task.scheduled_date)

        noncontiguous_count = 0
        for filling_days in filling_days_by_order.values():
            if len(filling_days) <= 1:
                continue
            first_day = min(filling_days)
            last_day = max(filling_days)
            expected_days = set(self._date_range(first_day, last_day))
            if filling_days != expected_days:
                noncontiguous_count += 1
        return noncontiguous_count

    def _target_penalty(
        self,
        overtime_by_group: dict[str, float],
        filling_noncontiguous_order_count: int,
    ) -> float:
        penalty = 0.0
        for labor_group, target_minutes in OVERTIME_TARGETS.items():
            overtime = overtime_by_group.get(labor_group, 0.0)
            if overtime > target_minutes:
                penalty += (overtime - target_minutes) * TARGET_EXCESS_WEIGHT
        penalty += self._weekend_overtime_minutes() * WEEKEND_OVERTIME_WEIGHT
        penalty += self._manufacturing_switch_cleaning_penalty()
        if self.keep_split_days_contiguous:
            penalty += filling_noncontiguous_order_count * FILLING_NONCONTIGUOUS_ORDER_WEIGHT
        return penalty

    def _capacity_penalty(
        self,
        line_change_by_day_machine: Mapping[tuple[date, str], float] | None = None,
    ) -> float:
        penalty = 0.0

        for (labor_group, current_day), used_minutes in self.daily_labor_usage.items():
            if current_day == date.min:
                penalty += used_minutes
                continue
            max_minutes = self._max_labor_minutes(labor_group, current_day)
            if used_minutes > max_minutes + 1e-6:
                penalty += (used_minutes - max_minutes) * self.infeasible_weight

        packaging_days = {
            current_day
            for labor_group, current_day in self.daily_labor_usage
            if labor_group in {"bottling", "filling"} and current_day != date.min
        }
        for current_day in packaging_days:
            packaging_used = self._packaging_labor_used(current_day)
            packaging_max = self._packaging_max_minutes(current_day)
            if packaging_used > packaging_max + 1e-6:
                penalty += (packaging_used - packaging_max) * self.infeasible_weight

        machine_usage = defaultdict(float, self.daily_machine_usage)
        for (current_day, machine_type), line_change_minutes in (line_change_by_day_machine or {}).items():
            machine_usage[(machine_type, current_day)] += line_change_minutes

        for (machine_type, current_day), used_minutes in machine_usage.items():
            max_minutes = 0.0 if not self._is_workday(current_day) else MACHINE_COUNTS.get(machine_type, 1) * MACHINE_DAILY_CAPACITY
            if used_minutes > max_minutes + 1e-6:
                penalty += (used_minutes - max_minutes) * self.infeasible_weight

        monthly_usage: dict[tuple[str, int, int], float] = defaultdict(float)
        for (labor_group, current_day), used_minutes in self.daily_labor_usage.items():
            if current_day == date.min or labor_group in {"bottling", "filling"}:
                continue
            overtime = max(used_minutes - self._regular_labor_minutes(labor_group, current_day), 0.0)
            monthly_usage[(labor_group, current_day.year, current_day.month)] += overtime

        packaging_monthly_usage: dict[tuple[int, int], float] = defaultdict(float)
        for current_day in packaging_days:
            overtime = max(self._packaging_labor_used(current_day) - self._packaging_regular_minutes(current_day), 0.0)
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

    def _date_range(self, start_day: date, end_day: date) -> list[date]:
        if end_day < start_day:
            return []
        if self.calendar_capacity is not None:
            return self.calendar_capacity.workdays_between(start_day, end_day)
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
    line_change_minutes: int = 0
    frozen_schedule_entries: tuple[FrozenScheduleEntry, ...] = ()
    calendar_capacity: CalendarCapacity | None = None
    independent_sections: bool = False
    keep_split_days_contiguous: bool = True
    break_start_time: str = DEFAULT_BREAK_START_TIME
    break_end_time: str = DEFAULT_BREAK_END_TIME

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
                f"target_penalty={generation_best_result.target_penalty:.1f} | "
                f"intraday_unallocated={generation_best_result.intraday_unallocated_task_count} | "
                f"noncontiguous_fill_orders={generation_best_result.filling_noncontiguous_order_count}"
            )

        return GreedyScheduleDecoder(
            line_change_minutes=self.line_change_minutes,
            frozen_schedule_entries=self.frozen_schedule_entries,
            calendar_capacity=self.calendar_capacity,
            independent_sections=self.independent_sections,
            keep_split_days_contiguous=self.keep_split_days_contiguous,
            break_start_time=self.break_start_time,
            break_end_time=self.break_end_time,
        ).decode(best_ordering)

    def _initialize_population(self, orders: list[Order]) -> list[list[Order]]:
        base = sorted(orders, key=lambda item: (item.due_date, item.material_ready_date, item.order_id))
        population = [base]
        seeded = self._build_seeded_chromosome(orders)
        if seeded is not None and seeded != base:
            population.append(seeded)
        population.append(
            sorted(
                orders,
                key=lambda item: (
                    self._packaging_pull_slack_minutes(item),
                    item.due_date,
                    item.material_ready_date,
                    item.order_id,
                ),
            )
        )
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

    def _packaging_pull_slack_minutes(self, order: Order) -> float:
        if self.calendar_capacity is None:
            available_days = (order.due_date - order.material_ready_date).days + 1
        else:
            available_days = len(
                self.calendar_capacity.workdays_between(order.material_ready_date, order.due_date)
            )
        available_minutes = max(available_days, 0) * MACHINE_DAILY_CAPACITY
        required_minutes = (
            order.filling_item.filling_machine_minutes
            + order.filling_item.bottling_machine_minutes
        )
        return available_minutes - required_minutes

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
            scored.append(
                (
                    chromosome[:],
                    GreedyScheduleDecoder(
                        line_change_minutes=self.line_change_minutes,
                        frozen_schedule_entries=self.frozen_schedule_entries,
                        calendar_capacity=self.calendar_capacity,
                        independent_sections=self.independent_sections,
                        keep_split_days_contiguous=self.keep_split_days_contiguous,
                        break_start_time=self.break_start_time,
                        break_end_time=self.break_end_time,
                    ).decode(chromosome),
                )
            )
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


def _slot_index_to_time(slot_index: int) -> str:
    start_minutes = DISPATCH_DAY_START_MINUTES + slot_index * SLOT_MINUTES
    day_offset, minute_of_day = divmod(start_minutes, 24 * 60)
    hours, minutes = divmod(minute_of_day, 60)
    if day_offset:
        return f"+{day_offset}天 {hours:02d}:{minutes:02d}"
    return f"{hours:02d}:{minutes:02d}"


def _format_slot_time_ranges(start_slot: int, end_slot: int, break_slots: set[int]) -> str:
    work_slots = [slot for slot in range(start_slot, end_slot + 1) if slot not in break_slots]
    if not work_slots:
        return f"{_slot_index_to_time(start_slot)}-{_slot_index_to_time(end_slot + 1)}"

    segments: list[tuple[int, int]] = []
    segment_start = work_slots[0]
    previous_slot = work_slots[0]
    for slot in work_slots[1:]:
        if slot == previous_slot + 1:
            previous_slot = slot
            continue
        segments.append((segment_start, previous_slot))
        segment_start = slot
        previous_slot = slot
    segments.append((segment_start, previous_slot))
    return "; ".join(
        f"{_slot_index_to_time(start)}-{_slot_index_to_time(end + 1)}"
        for start, end in segments
    )


def _parse_hhmm_minutes(value: str) -> int | None:
    text = str(value or "").strip()
    if not text:
        return None
    try:
        hour_text, minute_text = text.split(":", 1)
        hour = int(hour_text)
        minute = int(minute_text)
    except ValueError:
        return None
    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        return None
    return hour * 60 + minute


def _break_slots_from_times(start_time: str, end_time: str) -> set[int]:
    start_minutes = _parse_hhmm_minutes(start_time)
    end_minutes = _parse_hhmm_minutes(end_time)
    if start_minutes is None or end_minutes is None or end_minutes <= start_minutes:
        return set()
    start_slot = max(0, math.floor((start_minutes - DISPATCH_DAY_START_MINUTES) / SLOT_MINUTES))
    end_slot = max(start_slot, math.ceil((end_minutes - DISPATCH_DAY_START_MINUTES) / SLOT_MINUTES))
    return set(range(start_slot, end_slot))


def _break_slots_from_schedule(schedule: ScheduleResult) -> set[int]:
    return _break_slots_from_times(schedule.break_start_time, schedule.break_end_time)


def _range_avoids_break(slot_range: range, break_slots: set[int]) -> bool:
    return not any(slot in break_slots for slot in slot_range)


def _work_slots_from_start(
    start_slot: int,
    required_work_slots: int,
    slot_limit: int,
    break_slots: set[int],
) -> list[int] | None:
    """Return occupied work slots for a task that pauses during break slots."""
    if required_work_slots <= 0:
        return []
    if start_slot in break_slots:
        return None
    occupied_slots: list[int] = []
    cursor = start_slot
    while cursor < slot_limit and len(occupied_slots) < required_work_slots:
        if cursor not in break_slots:
            occupied_slots.append(cursor)
        cursor += 1
    return occupied_slots if len(occupied_slots) == required_work_slots else None


def _slot_limit_for_work_slots(work_slots: int, break_slots: set[int]) -> int:
    """Return timeline slots needed to fit working slots while skipping breaks."""
    slot_limit = max(0, int(work_slots))
    while True:
        blocked_slots = sum(1 for slot in break_slots if 0 <= slot < slot_limit)
        adjusted_limit = max(0, int(work_slots)) + blocked_slots
        if adjusted_limit == slot_limit:
            return slot_limit
        slot_limit = adjusted_limit

OUTPUT_TEXT_NORMALIZATION_MAP = {
    # 保留舊版亂碼修正表，避免早期中間字串流進輸出報表。
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

def _intraday_labor_pool(task: ScheduledTask) -> str:
    return "manufacturing" if task.task_type in {"Cat1", "Cat2"} else "packaging"

def _display_intraday_labor_pool(labor_pool: str) -> str:
    return "調製人力" if labor_pool == "manufacturing" else "排瓶/充填共用人力"

def _intraday_slot_limit(tasks: list[ScheduledTask], break_slots: set[int] | None = None) -> int:
    """Allow exact dispatch output to extend into overtime instead of dropping tasks."""
    regular_slots = int(MACHINE_DAILY_CAPACITY // SLOT_MINUTES)
    break_slots = break_slots or set()
    if not tasks:
        return _slot_limit_for_work_slots(regular_slots, break_slots)
    sequential_slots = sum(
        max(1, math.ceil(task.machine_minutes / SLOT_MINUTES))
        for task in tasks
    )
    longest_task_slots = max(
        max(1, math.ceil(task.machine_minutes / SLOT_MINUTES))
        for task in tasks
    )
    return max(
        _slot_limit_for_work_slots(regular_slots, break_slots),
        _slot_limit_for_work_slots(sequential_slots, break_slots),
        _slot_limit_for_work_slots(longest_task_slots, break_slots),
    )

def _intraday_people_capacity(schedule: ScheduleResult, labor_pool: str, current_day: date) -> int:
    if schedule.calendar_capacity is not None:
        if not schedule.calendar_capacity.is_workday(current_day):
            return 0
        day_capacity = schedule.calendar_capacity.days.get(current_day)
        if day_capacity is not None:
            return (
                day_capacity.manufacturing_people
                if labor_pool == "manufacturing"
                else day_capacity.filling_people
            )
    return LABOR_CONFIG["manufacturing" if labor_pool == "manufacturing" else "filling"]["people"]

def _intraday_priority(task: ScheduledTask, mode: str) -> tuple[object, ...]:
    people = _infer_actual_assigned_people(task)
    slots = max(1, math.ceil(task.machine_minutes / SLOT_MINUTES))
    stage_order = {"Cat1": 1, "Cat2": 2, "Cat3": 3, "Cat4": 4}
    if mode == "stage_first":
        return (stage_order.get(task.task_type, 99), -slots, task.order_id, task.machine_type)
    if mode == "long_first":
        return (-slots, -people, stage_order.get(task.task_type, 99), task.order_id)
    if mode == "light_first":
        return (people, -slots, stage_order.get(task.task_type, 99), task.order_id)
    return (-people, -slots, stage_order.get(task.task_type, 99), task.order_id)

def _allocate_intraday_group(
    schedule: ScheduleResult,
    tasks: list[ScheduledTask],
    labor_pool: str,
    same_day_dependency_orders: set[str] | None = None,
) -> dict[ScheduledTask, IntradayTaskAllocation]:
    """Try several deterministic packings and retain the most complete dispatch plan."""
    if not tasks:
        return {}

    same_day_dependency_orders = same_day_dependency_orders or set()
    capacity_people = _intraday_people_capacity(schedule, labor_pool, tasks[0].scheduled_date)
    break_slots = _break_slots_from_schedule(schedule)
    slot_limit = _intraday_slot_limit(tasks, break_slots)
    modes = ("stage_first", "heavy_first", "long_first", "light_first")
    best_allocations: dict[ScheduledTask, IntradayTaskAllocation] = {}
    best_score: tuple[int, int] = (-1, -1)

    for mode in modes:
        for direction in ("early", "late"):
            labor_slots = [0] * slot_limit
            machine_slots: dict[str, list[int]] = defaultdict(lambda: [0] * slot_limit)
            pending = list(tasks)
            allocations: dict[ScheduledTask, IntradayTaskAllocation] = {}

            while pending:
                eligible: list[tuple[ScheduledTask, int]] = []
                for task in pending:
                    dependency_start = 0
                    if task.task_type == "Cat4" and task.order_id in same_day_dependency_orders:
                        same_day_bottling = [
                            other
                            for other in tasks
                            if other.order_id == task.order_id and other.task_type == "Cat3"
                        ]
                        if any(
                            other not in allocations or not allocations[other].is_allocated
                            for other in same_day_bottling
                        ):
                            continue
                        dependency_start = max(
                            (
                                allocations[other].end_slot + 1
                                for other in same_day_bottling
                                if allocations[other].end_slot is not None
                            ),
                            default=0,
                        )
                    eligible.append((task, dependency_start))

                if not eligible:
                    for task in pending:
                        allocations[task] = IntradayTaskAllocation(
                            task=task,
                            labor_pool=labor_pool,
                            actual_people=_infer_actual_assigned_people(task),
                            start_slot=None,
                            end_slot=None,
                        )
                    break

                eligible.sort(key=lambda item: _intraday_priority(item[0], mode))
                task, earliest_start = eligible[0]
                pending.remove(task)
                people = _infer_actual_assigned_people(task)
                required_slots = max(1, math.ceil(task.machine_minutes / SLOT_MINUTES))
                machine_count = max(MACHINE_COUNTS.get(task.machine_type, 1), 1)
                start_slot: int | None = None
                end_slot: int | None = None
                latest_start = slot_limit - 1
                candidate_starts: Iterable[int]
                if direction == "early":
                    candidate_starts = range(earliest_start, latest_start + 1)
                else:
                    candidate_starts = range(latest_start, earliest_start - 1, -1)

                if people <= capacity_people and required_slots <= slot_limit:
                    for candidate_start in candidate_starts:
                        required_range = _work_slots_from_start(
                            candidate_start,
                            required_slots,
                            slot_limit,
                            break_slots,
                        )
                        if required_range is not None and all(
                            machine_slots[task.machine_type][slot] < machine_count
                            and labor_slots[slot] + people <= capacity_people
                            for slot in required_range
                        ):
                            start_slot = candidate_start
                            end_slot = required_range[-1]
                            for slot in required_range:
                                machine_slots[task.machine_type][slot] += 1
                                labor_slots[slot] += people
                            break

                allocations[task] = IntradayTaskAllocation(
                    task=task,
                    labor_pool=labor_pool,
                    actual_people=people,
                    start_slot=start_slot,
                    end_slot=end_slot,
                )

            allocated_count = sum(allocation.is_allocated for allocation in allocations.values())
            allocated_slots = sum(
                max(1, math.ceil(task.machine_minutes / SLOT_MINUTES))
                for task, allocation in allocations.items()
                if allocation.is_allocated
            )
            score = (allocated_count, allocated_slots)
            if score > best_score:
                best_score = score
                best_allocations = allocations

    if (
        best_score[0] < len(tasks)
        and schedule.infeasible_penalty <= 1e-6
        and labor_pool == "manufacturing"
        and len(tasks) <= 10
    ):
        repaired_allocations = _repair_intraday_group_with_backtracking(
            schedule,
            tasks,
            labor_pool,
            same_day_dependency_orders,
        )
        if repaired_allocations is not None:
            return repaired_allocations
    return best_allocations

def _repair_intraday_group_with_backtracking(
    schedule: ScheduleResult,
    tasks: list[ScheduledTask],
    labor_pool: str,
    same_day_dependency_orders: set[str],
) -> dict[ScheduledTask, IntradayTaskAllocation] | None:
    """Find a complete small daily dispatch when greedy packing leaves a false conflict."""
    if len(tasks) > 16:
        return None

    capacity_people = _intraday_people_capacity(schedule, labor_pool, tasks[0].scheduled_date)
    break_slots = _break_slots_from_schedule(schedule)
    slot_limit = _intraday_slot_limit(tasks, break_slots)
    labor_slots = [0] * slot_limit
    machine_slots: dict[str, list[int]] = defaultdict(lambda: [0] * slot_limit)
    allocations: dict[ScheduledTask, IntradayTaskAllocation] = {}
    nodes_visited = 0
    max_nodes = 200_000

    def search(pending: list[ScheduledTask]) -> bool:
        nonlocal nodes_visited
        nodes_visited += 1
        if nodes_visited > max_nodes:
            return False
        if not pending:
            return True

        eligible: list[tuple[ScheduledTask, int]] = []
        for task in pending:
            earliest_start = 0
            if task.task_type == "Cat4" and task.order_id in same_day_dependency_orders:
                bottling_tasks = [
                    other
                    for other in tasks
                    if other.order_id == task.order_id and other.task_type == "Cat3"
                ]
                if any(other not in allocations for other in bottling_tasks):
                    continue
                earliest_start = max(
                    (allocations[other].end_slot + 1 for other in bottling_tasks),
                    default=0,
                )
            eligible.append((task, earliest_start))
        if not eligible:
            return False

        eligible.sort(key=lambda item: _intraday_priority(item[0], "long_first"))
        task, earliest_start = eligible[0]
        people = _infer_actual_assigned_people(task)
        required_slots = max(1, math.ceil(task.machine_minutes / SLOT_MINUTES))
        if people > capacity_people or required_slots > slot_limit:
            return False
        machine_count = max(MACHINE_COUNTS.get(task.machine_type, 1), 1)
        latest_start = slot_limit - 1

        for start_slot in range(earliest_start, latest_start + 1):
            occupied_slots = _work_slots_from_start(
                start_slot,
                required_slots,
                slot_limit,
                break_slots,
            )
            if occupied_slots is None or not all(
                labor_slots[slot] + people <= capacity_people
                and machine_slots[task.machine_type][slot] < machine_count
                for slot in occupied_slots
            ):
                continue
            for slot in occupied_slots:
                labor_slots[slot] += people
                machine_slots[task.machine_type][slot] += 1
            allocations[task] = IntradayTaskAllocation(
                task=task,
                labor_pool=labor_pool,
                actual_people=people,
                start_slot=start_slot,
                end_slot=occupied_slots[-1],
            )
            next_pending = pending[:]
            next_pending.remove(task)
            if search(next_pending):
                return True
            del allocations[task]
            for slot in occupied_slots:
                labor_slots[slot] -= people
                machine_slots[task.machine_type][slot] -= 1
        return False

    return allocations if search(tasks[:]) else None

def build_intraday_allocations(schedule: ScheduleResult) -> list[IntradayTaskAllocation]:
    """Allocate exact dispatch slots while respecting simultaneous labor and machines."""
    grouped_tasks = _group_output_tasks(schedule.task_records)
    tasks_by_day_pool: dict[tuple[date, str], list[ScheduledTask]] = defaultdict(list)
    for task in grouped_tasks:
        tasks_by_day_pool[(task.scheduled_date, _intraday_labor_pool(task))].append(task)

    allocations: list[IntradayTaskAllocation] = []
    for (current_day, labor_pool), tasks in sorted(tasks_by_day_pool.items()):
        same_day_dependency_orders: set[str] = set()
        if labor_pool == "packaging":
            previous_workday = (
                schedule.calendar_capacity.previous_workday(current_day, date.min)
                if schedule.calendar_capacity is not None
                else current_day - timedelta(days=1)
            )
            same_day_bottling_orders = {task.order_id for task in tasks if task.task_type == "Cat3"}
            previous_day_bottling_orders = {
                task.order_id
                for task in grouped_tasks
                if task.task_type == "Cat3" and task.scheduled_date == previous_workday
            }
            same_day_dependency_orders = same_day_bottling_orders - previous_day_bottling_orders
        task_allocations = _allocate_intraday_group(
            schedule,
            tasks,
            labor_pool,
            same_day_dependency_orders=same_day_dependency_orders,
        )
        allocations.extend(task_allocations[task] for task in tasks)
    return allocations

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
                    f"充填換線分鐘數={_display_integer(summary['line_change_minutes'])}; "
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
        for (current_day, _machine_type), line_change_minutes in schedule.line_change_by_day_machine.items():
            daily_base_usage[("filling", current_day)] += line_change_minutes

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
                "充填換線分鐘數",
                "總排瓶分鐘數",
                "總充填分鐘數",
                "總充填含換線分鐘數",
                "日內未配置任務數",
                "非連續充填訂單數",
            ]
        )
        writer.writerow(
            [
                _display_integer(summary["total_overtime_minutes"]),
                _display_integer(summary["manufacturing_overtime_minutes"]),
                _display_integer(summary["bottling_overtime_minutes"]),
                _display_integer(summary["filling_overtime_minutes"]),
                _display_integer(summary["weekend_overtime_minutes"]),
                _display_integer(summary["line_change_minutes"]),
                _display_integer(summary["total_bottling_minutes"]),
                _display_integer(summary["total_filling_minutes"]),
                _display_integer(summary["total_filling_with_line_change_minutes"]),
                _display_integer(summary["intraday_unallocated_task_count"]),
                _display_integer(summary["filling_noncontiguous_order_count"]),
            ]
        )

def write_filling_line_change_audit_output(schedule: ScheduleResult, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / "filling_line_change_audit.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "日期",
                "充填機",
                "充填分鐘",
                "不同產品數",
                "推估換線次數",
                "換線分鐘",
                "訂單清單",
                "產品清單",
            ]
        )
        for row in build_line_change_audit_rows(schedule):
            if float(row["line_change_minutes"]) <= 0 and schedule.line_change_minutes_per_switch <= 0:
                continue
            writer.writerow(
                [
                    row["date"].isoformat(),
                    _display_machine_label(str(row["machine"])),
                    _display_integer(float(row["filling_minutes"])),
                    row["distinct_product_count"],
                    row["estimated_switch_count"],
                    _display_integer(float(row["line_change_minutes"])),
                    row["orders"],
                    row["products"],
                ]
            )

def write_frozen_schedule_audit_output(schedule: ScheduleResult, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    tasks_by_key: dict[tuple[str, str], list[ScheduledTask]] = defaultdict(list)
    for task in schedule.task_records:
        tasks_by_key[(task.order_id, task.task_type)].append(task)

    path = output_dir / "frozen_schedule_audit.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["訂單", "指定工序", "模型工序", "指定日期", "實際日期", "狀態", "說明"])
        for entry in schedule.frozen_schedule_entries:
            tasks = tasks_by_key.get((entry.order_id, entry.task_type), [])
            actual_dates = sorted({task.scheduled_date for task in tasks})
            if not tasks:
                status = "未套用"
                message = "沒有找到此訂單工序的排程任務，可能是訂單未進入 GA 或凍結日期不可行。"
            elif all(task.scheduled_date == entry.scheduled_date for task in tasks):
                status = "已套用"
                message = "所有對應任務皆排在指定日期。"
            elif any(task.scheduled_date == entry.scheduled_date for task in tasks):
                status = "部分套用"
                message = "部分任務排在指定日期，但同訂單同工序仍有其他日期。"
            else:
                status = "日期不符"
                message = "模型未能把此工序排在指定日期，請檢查容量、ready date 與 due date。"
            writer.writerow(
                [
                    entry.order_id,
                    entry.source_stage or _display_task_stage(entry.task_type),
                    _display_task_stage(entry.task_type),
                    entry.scheduled_date.isoformat(),
                    " | ".join(day.isoformat() for day in actual_dates),
                    status,
                    message,
                ]
            )

WEEKDAY_LABELS = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]

def _week_start(current_day: date) -> date:
    return current_day - timedelta(days=current_day.weekday())

def _compact_order_list(order_ids: list[str], limit: int = 5) -> str:
    if len(order_ids) <= limit:
        return "、".join(order_ids)
    return "、".join(order_ids[:limit]) + f"、...等{len(order_ids)}張"

def _compact_product_list(tasks: list[ScheduledTask], order_map: dict[str, Order], limit: int = 3) -> str:
    product_names: list[str] = []
    seen: set[str] = set()
    for task in tasks:
        order = order_map.get(task.order_id)
        product_name = order.product_name if order is not None else _display_task_detail(task)
        if product_name in seen:
            continue
        seen.add(product_name)
        product_names.append(product_name)
        if len(product_names) >= limit:
            break
    if len(seen) < len({task.order_id for task in tasks}):
        product_names.append("...")
    return "、".join(product_names)

def _dispatch_day_cell(
    current_day: date,
    tasks_by_day_stage: dict[tuple[date, str], list[ScheduledTask]],
    order_map: dict[str, Order],
    overtime_by_day: dict[date, list[str]],
    line_change_by_day: dict[date, float],
    frozen_by_day: dict[date, list[str]],
) -> str:
    lines = [f"{current_day.isoformat()}（{WEEKDAY_LABELS[current_day.weekday()].replace('星期', '')}）"]
    stage_order = ["預攪", "攪料清洗", "排瓶", "充填"]
    has_work = False
    for stage in stage_order:
        tasks = tasks_by_day_stage.get((current_day, stage), [])
        if not tasks:
            continue
        has_work = True
        order_ids = sorted({task.order_id for task in tasks})
        machine_minutes = sum(task.machine_minutes for task in tasks)
        labor_minutes = sum(task.labor_minutes for task in tasks)
        products = _compact_product_list(tasks, order_map)
        lines.append(
            f"{stage}: {_compact_order_list(order_ids)} ({len(order_ids)}單, 機{_display_integer(machine_minutes)}分, 人{_display_integer(labor_minutes)}分)"
        )
        if products:
            lines.append(f"  品項: {products}")
    if frozen_by_day.get(current_day):
        lines.append("凍結: " + "、".join(frozen_by_day[current_day]))
    if line_change_by_day.get(current_day, 0.0) > 0:
        lines.append(f"換線: {_display_integer(line_change_by_day[current_day])}分")
    if overtime_by_day.get(current_day):
        lines.append("超時: " + "、".join(overtime_by_day[current_day]))
    if not has_work:
        lines.append("無排程")
    return "\n".join(lines)

def write_dispatch_week_view_output(schedule: ScheduleResult, orders: list[Order], output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    grouped_tasks = _group_output_tasks(schedule.task_records)
    order_map = {order.order_id: order for order in orders}
    tasks_by_day_stage: dict[tuple[date, str], list[ScheduledTask]] = defaultdict(list)
    for task in grouped_tasks:
        tasks_by_day_stage[(task.scheduled_date, _display_task_stage(task.task_type))].append(task)

    scheduled_days = sorted({task.scheduled_date for task in grouped_tasks if task.scheduled_date != date.min})
    if not scheduled_days:
        (output_dir / "dispatch_week_view.csv").write_text("週起始日,星期一,星期二,星期三,星期四,星期五,星期六,星期日\n", encoding="utf-8-sig")
        return

    overtime_by_day: dict[date, list[str]] = defaultdict(list)
    for row in _daily_resource_load_rows(schedule):
        overtime = float(row["overtime"])
        if overtime > 0:
            overtime_by_day[row["date"]].append(f"{row['resource']}{_display_integer(overtime)}分")

    line_change_by_day: dict[date, float] = defaultdict(float)
    for (current_day, _machine_type), minutes in schedule.line_change_by_day_machine.items():
        line_change_by_day[current_day] += minutes

    frozen_by_day: dict[date, list[str]] = defaultdict(list)
    for entry in schedule.frozen_schedule_entries:
        frozen_by_day[entry.scheduled_date].append(f"{entry.order_id}-{_display_task_stage(entry.task_type)}")

    first_week = _week_start(scheduled_days[0])
    last_week = _week_start(scheduled_days[-1])
    week_starts = [
        first_week + timedelta(days=offset)
        for offset in range(0, (last_week - first_week).days + 1, 7)
    ]

    path = output_dir / "dispatch_week_view.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["週起始日", *WEEKDAY_LABELS])
        for week_start in week_starts:
            row = [week_start.isoformat()]
            for offset in range(7):
                current_day = week_start + timedelta(days=offset)
                row.append(
                    _dispatch_day_cell(
                        current_day,
                        tasks_by_day_stage,
                        order_map,
                        overtime_by_day,
                        line_change_by_day,
                        frozen_by_day,
                    )
                )
            writer.writerow(row)

def _format_elapsed_hms(elapsed_seconds: float) -> str:
    total_seconds = max(0, int(round(elapsed_seconds)))
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

def _metadata_value(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "是" if value else "否"
    return str(value)

def write_developer_options_output(
    output_dir: Path,
    advanced_options: Mapping[str, object] | None,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    option_labels = {
        "data_source_mode": "資料來源模式",
        "google_sheet_id": "Google Sheet ID",
        "schedule_mode": "排程模式",
        "calendar_capacity_path": "工作日曆與人力容量檔",
        "manufacturing_efficiency_path": "外部製造效能表",
        "filling_efficiency_path": "外部充填效能表",
        "exclude_weekends": "排除週六日",
        "frozen_schedule_csv": "凍結排程 CSV",
        "daily_rerun_time": "每日重新排程時間",
        "line_change_minutes": "換線分鐘",
        "break_start_time": "每日休息開始",
        "break_end_time": "每日休息結束",
        "split_unit": "拆單單位",
        "minimum_fill_liters": "最小充填量(L)",
        "keep_split_days_contiguous": "拆單日期盡量連續",
        "small_printer_count": "小噴印機數量",
        "normal_printer_count": "正常噴印機數量",
        "output_week_view": "輸出派工週視圖",
    }
    implemented_options = {
        "schedule_mode",
        "calendar_capacity_path",
        "manufacturing_efficiency_path",
        "filling_efficiency_path",
        "exclude_weekends",
        "frozen_schedule_csv",
        "line_change_minutes",
        "break_start_time",
        "break_end_time",
        "keep_split_days_contiguous",
        "output_week_view",
    }
    path = output_dir / "developer_options.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["設定代碼", "設定名稱", "設定值", "目前狀態"])
        for key, label in option_labels.items():
            value = "" if advanced_options is None else advanced_options.get(key, "")
            if value in ("", None):
                status = "未設定"
            elif key in implemented_options:
                status = "已接入排程運算或輸出"
            else:
                status = "已記錄；後續可逐項接入排程限制"
            writer.writerow([key, label, _metadata_value(value), status])

def write_run_metadata(
    output_dir: Path,
    *,
    started_at: datetime,
    finished_at: datetime,
    elapsed_seconds: float,
    workbook_path: Path,
    population_size: int,
    generations: int,
    mutation_rate: float,
    crossover_rate: float,
    seed: int,
    advanced_options: Mapping[str, object] | None = None,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    metadata_path = output_dir / "run_metadata.csv"
    with metadata_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["欄位", "數值"])
        writer.writerow(["執行開始時間", started_at.isoformat(timespec="seconds")])
        writer.writerow(["執行結束時間", finished_at.isoformat(timespec="seconds")])
        writer.writerow(["總執行秒數", f"{elapsed_seconds:.2f}"])
        writer.writerow(["總執行時間(HH:MM:SS)", _format_elapsed_hms(elapsed_seconds)])
        writer.writerow(["輸入活頁簿", str(workbook_path)])
        writer.writerow(["族群大小", population_size])
        writer.writerow(["世代數", generations])
        writer.writerow(["突變率", mutation_rate])
        writer.writerow(["交配率", crossover_rate])
        writer.writerow(["隨機種子", seed])
        if advanced_options:
            writer.writerow([])
            writer.writerow(["下一版開發設定", ""])
            for key, value in advanced_options.items():
                writer.writerow([key, _metadata_value(value)])

def write_reference_format_output(
    schedule: ScheduleResult,
    orders: list[Order],
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    summary = build_schedule_summary(schedule)
    intraday_allocations = build_intraday_allocations(schedule)
    allocation_by_task = {allocation.task: allocation for allocation in intraday_allocations}
    grouped_tasks = [allocation.task for allocation in intraday_allocations]
    order_map = {order.order_id: order for order in orders}

    manufacturing_group_totals: dict[tuple[str, str], int] = defaultdict(int)
    packaging_group_totals: dict[tuple[str, str], float] = defaultdict(float)
    manufacturing_group_ids: dict[tuple[str, str], int] = {}

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
        break_slots = _break_slots_from_schedule(schedule)
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
            allocation = allocation_by_task[task]
            slot_count = max(1, math.ceil(task.machine_minutes / 5.0))
            start_slot = allocation.start_slot
            end_slot = allocation.end_slot

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
                actual_people = allocation.actual_people
                subtask_id = f"{task.order_id}_{task.task_type}_{seq}"
                unit_minutes = 5

            time_range = (
                _format_slot_time_ranges(start_slot, end_slot, break_slots)
                if start_slot is not None and end_slot is not None
                else "未配置：日內資源不足"
            )
            scheduling_type = "正常" if allocation.is_allocated else "阻擋：日內資源不足"

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
                    time_range,
                    slot_count,
                    scheduling_type,
                    _display_integer(task.machine_minutes),
                    _display_integer(demand_units),
                    start_slot if start_slot is not None else "",
                    end_slot if end_slot is not None else "",
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
    error_count = by_severity.get("error", 0)
    warning_count = by_severity.get("warning", 0)

    with summary_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["metric", "value"])
        writer.writerow(["total_issues", len(issues)])
        writer.writerow(["blocking_issue_count", error_count])
        writer.writerow(["error_count", error_count])
        writer.writerow(["warning_count", warning_count])
        writer.writerow(["delivery_ready_if_error_count_zero", "yes" if error_count == 0 else "no"])
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


def write_input_model_issues_output(
    issues: list[dict[str, str]],
    output_dir: Path,
) -> None:
    """Write source orders that were not usable by the scheduling model."""
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / "input_model_issues.csv"
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["order_id", "product_name", "issue_type", "message"])
        for issue in issues:
            writer.writerow(
                [
                    issue.get("order_id", ""),
                    issue.get("product_name", ""),
                    issue.get("issue_type", ""),
                    issue.get("message", ""),
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

def detect_order_anomalies(
    orders: list[Order],
    independent_sections: bool = False,
    calendar_capacity: CalendarCapacity | None = None,
) -> list[OrderAnomaly]:
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
            and not allows_same_day_bottling(order, order.due_date)
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

        if order.material_ready_date <= order.due_date:
            packaging_days = (
                calendar_capacity.workdays_between(order.material_ready_date, order.due_date)
                if calendar_capacity is not None
                else [
                    order.material_ready_date + timedelta(days=offset)
                    for offset in range((order.due_date - order.material_ready_date).days + 1)
                ]
            )
            if packaging_days:
                filling_item = order.filling_item
                assigned_bottling_people = min(
                    max(filling_item.bottling_people, BOTTLING_MAX_PEOPLE),
                    BOTTLING_MAX_PEOPLE,
                )
                actual_bottling_minutes = (
                    filling_item.bottling_machine_minutes
                    * max(filling_item.bottling_people, 1)
                    / max(assigned_bottling_people, 1)
                )
                machine_window_minutes = len(packaging_days) * MACHINE_DAILY_CAPACITY
                max_packaging_people = max(
                    (
                        calendar_capacity.days.get(current_day).filling_people
                        if (
                            calendar_capacity is not None
                            and calendar_capacity.days.get(current_day) is not None
                        )
                        else LABOR_CONFIG["filling"]["people"]
                    )
                    for current_day in packaging_days
                )
                cannot_overlap = (
                    filling_item.filling_people + assigned_bottling_people > max_packaging_people
                )
                if cannot_overlap and (
                    filling_item.filling_machine_minutes + actual_bottling_minutes
                    > machine_window_minutes + 1e-6
                ):
                    anomalies.append(
                        OrderAnomaly(
                            order_id=order.order_id,
                            product_name=order.product_name,
                            severity="exclude",
                            anomaly_type="packaging_sequence_window_too_short",
                            message=(
                                f"bottling and filling require {_display_integer(actual_bottling_minutes + filling_item.filling_machine_minutes)} "
                                f"sequential minutes, exceeding {_display_integer(machine_window_minutes)} available "
                                f"minutes from {order.material_ready_date} to {order.due_date}"
                            ),
                            exclude_from_optimization=True,
                        )
                    )

        manufacturing_latest_day = (
            order.due_date
            if independent_sections
            else order.due_date - timedelta(days=1)
        )
        manufacturing_window_start = (
            max(order.raw_ready_date, order.due_date - timedelta(days=30))
            if independent_sections
            else order.raw_ready_date
        )
        if calendar_capacity is not None and not independent_sections:
            for item in order.manufacturing_items:
                requires_premix = item.premix_minutes_per_batch > 1e-2
                if requires_premix:
                    first_mix_day = manufacturing_window_start + timedelta(days=1)
                    has_feasible_day = any(
                        calendar_capacity.previous_workday(mix_day, manufacturing_window_start) is not None
                        for mix_day in calendar_capacity.workdays_between(first_mix_day, manufacturing_latest_day)
                    )
                else:
                    has_feasible_day = bool(
                        calendar_capacity.workdays_between(manufacturing_window_start, manufacturing_latest_day)
                    )
                if not has_feasible_day:
                    anomalies.append(
                        OrderAnomaly(
                            order_id=order.order_id,
                            product_name=order.product_name,
                            severity="exclude",
                            anomaly_type="no_manufacturing_workday_before_filling",
                            message=(
                                f"no feasible manufacturing workday for {item.process_name} between "
                                f"{manufacturing_window_start} and {manufacturing_latest_day} before filling"
                            ),
                            exclude_from_optimization=True,
                        )
                    )
                    break
        manufacturing_window_days = max(
            (manufacturing_latest_day - manufacturing_window_start).days,
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
    independent_sections: bool = False,
    calendar_capacity: CalendarCapacity | None = None,
) -> tuple[list[Order], list[OrderAnomaly]]:
    anomalies = detect_order_anomalies(
        orders,
        independent_sections=independent_sections,
        calendar_capacity=calendar_capacity,
    )
    excluded_order_ids = {
        anomaly.order_id
        for anomaly in anomalies
        if anomaly.exclude_from_optimization or (exclude_warnings and anomaly.severity == "warning")
    }
    schedulable_orders = [order for order in orders if order.order_id not in excluded_order_ids]
    return schedulable_orders, anomalies


def load_seed_order_ids(schedule_csv_path: Path) -> tuple[str, ...]:
    order_first_seen: dict[str, tuple[str, int, int, str]] = {}
    task_rank = {
        "Cat1": 1,
        "Cat2": 2,
        "Cat3": 3,
        "Cat4": 4,
        "預攪": 1,
        "攪料清洗": 2,
        "排瓶": 3,
        "充填": 4,
    }
    with schedule_csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row_index, row in enumerate(reader):
            order_id = (
                row.get("order_id")
                or row.get("訂單")
                or row.get("訂單編號")
                or ""
            ).strip()
            if not order_id or order_id in {"TOTAL", "總計"}:
                continue
            scheduled_date = (
                row.get("scheduled_date")
                or row.get("日期")
                or row.get("排程日期")
                or ""
            ).strip()
            task_type = (
                row.get("task_type")
                or row.get("工序代碼")
                or row.get("類別")
                or row.get("作業工序")
                or ""
            ).strip()
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

def _normalize_frozen_task_type(value: str) -> str | None:
    text = (value or "").strip()
    if not text:
        return None
    upper_text = text.upper()
    if upper_text in {"CAT1", "CAT2", "CAT3", "CAT4"}:
        return upper_text.title()
    if "預攪" in text:
        return "Cat1"
    if "攪料" in text or "清洗" in text or "製造" in text or "調製" in text:
        return "Cat2"
    if "排瓶" in text:
        return "Cat3"
    if "充填" in text:
        return "Cat4"
    return None

def _parse_iso_date_text(value: str) -> date:
    return _parse_date_text(value)


def _parse_date_text(value: str) -> date:
    text = str(value or "").strip()
    if not text:
        raise ValueError("日期欄位不可為空。")
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    try:
        return (EXCEL_EPOCH + timedelta(days=float(text))).date()
    except ValueError as exc:
        raise ValueError(f"無法辨識日期：{value}") from exc


def _manual_frozen_entries(value: object) -> tuple[FrozenScheduleEntry, ...]:
    if not value:
        return ()
    entries: list[FrozenScheduleEntry] = []
    if isinstance(value, tuple) and all(isinstance(item, FrozenScheduleEntry) for item in value):
        return value
    if not isinstance(value, list):
        return ()
    for index, item in enumerate(value, start=1):
        if isinstance(item, FrozenScheduleEntry):
            entries.append(item)
            continue
        if not isinstance(item, Mapping):
            continue
        order_id = str(item.get("order_id") or item.get("訂單") or "").strip()
        stage_text = str(item.get("stage") or item.get("工序") or item.get("task_type") or "").strip()
        date_text = str(item.get("date") or item.get("日期") or item.get("scheduled_date") or "").strip()
        if not order_id or not stage_text or not date_text:
            raise ValueError(f"手動凍結排程第 {index} 筆缺少訂單、工序或日期。")
        task_type = _normalize_frozen_task_type(stage_text)
        if task_type is None:
            raise ValueError(f"手動凍結排程第 {index} 筆工序無法辨識：{stage_text}")
        entries.append(
            FrozenScheduleEntry(
                order_id=order_id,
                task_type=task_type,
                scheduled_date=_parse_date_text(date_text),
                source_stage=stage_text,
            )
        )
    return tuple(entries)


def load_calendar_capacity(
    calendar_capacity_path: Path | None,
    *,
    exclude_weekends: bool = True,
) -> CalendarCapacity | None:
    if calendar_capacity_path is None:
        return None
    reader = SimpleXlsxReader(calendar_capacity_path)
    rows = reader.read_sheet("calendar_capacity")
    days: dict[date, CalendarDayCapacity] = {}
    for row in rows:
        current_day = _parse_date_text(row.get("日期", ""))
        holiday_text = str(row.get("是否假日", "")).strip()
        is_holiday = "假" in holiday_text or holiday_text.casefold() in {"holiday", "true", "1", "yes"}
        days[current_day] = CalendarDayCapacity(
            current_day=current_day,
            is_holiday=is_holiday,
            filling_people=max(0, WorkbookLoader._to_int(row.get("充填人力", ""))),
            manufacturing_people=max(0, WorkbookLoader._to_int(row.get("製造人力", ""))),
            note=str(row.get("註解", "")).strip(),
        )
    return CalendarCapacity(days=days, exclude_weekends=exclude_weekends)

def load_frozen_schedule_entries(schedule_csv_path: Path | None) -> tuple[FrozenScheduleEntry, ...]:
    if schedule_csv_path is None:
        return ()

    entries: list[FrozenScheduleEntry] = []
    with schedule_csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row_index, row in enumerate(reader, start=2):
            order_id = (
                row.get("order_id")
                or row.get("訂單")
                or row.get("訂單編號")
                or ""
            ).strip()
            if not order_id or order_id in {"TOTAL", "總計"}:
                continue
            stage_text = (
                row.get("stage")
                or row.get("task_type")
                or row.get("工序")
                or row.get("工序代碼")
                or row.get("類別")
                or row.get("作業工序")
                or ""
            ).strip()
            task_type = _normalize_frozen_task_type(stage_text)
            if task_type is None:
                raise ValueError(f"凍結排程 CSV 第 {row_index} 列工序無法辨識：{stage_text}")
            date_text = (
                row.get("date")
                or row.get("scheduled_date")
                or row.get("日期")
                or row.get("排程日期")
                or ""
            ).strip()
            if not date_text:
                raise ValueError(f"凍結排程 CSV 第 {row_index} 列缺少日期。")
            entries.append(
                FrozenScheduleEntry(
                    order_id=order_id,
                    task_type=task_type,
                    scheduled_date=_parse_iso_date_text(date_text),
                    source_stage=stage_text,
                )
            )
    return tuple(entries)

def build_timestamped_output_dir(base_output_dir: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return base_output_dir.parent / f"{base_output_dir.name}_{timestamp}"

def _option_int(
    options: Mapping[str, object] | None,
    key: str,
    default: int = 0,
) -> int:
    if not options:
        return default
    value = options.get(key, default)
    if value in ("", None):
        return default
    return max(0, int(value))


def _option_path(
    options: Mapping[str, object] | None,
    key: str,
) -> Path | None:
    if not options:
        return None
    value = options.get(key)
    if value in ("", None):
        return None
    return Path(str(value))


def _option_bool(
    options: Mapping[str, object] | None,
    key: str,
    default: bool = False,
) -> bool:
    if not options or key not in options:
        return default
    value = options.get(key)
    if isinstance(value, bool):
        return value
    if value in ("", None):
        return default
    return str(value).strip().lower() in {"1", "true", "yes", "y", "是", "輸出"}


def _option_text(
    options: Mapping[str, object] | None,
    key: str,
    default: str = "",
) -> str:
    if not options or key not in options:
        return default
    value = options.get(key)
    if value in ("", None):
        return default
    return str(value).strip()


def run_scheduler(
    *,
    workbook_path: Path,
    output_dir: Path,
    population_size: int = 40,
    generations: int = 120,
    mutation_rate: float = 0.15,
    crossover_rate: float = 0.85,
    seed: int = 42,
    seed_schedule_csv: Path | None = None,
    exclude_warning_anomalies: bool = False,
    advanced_options: Mapping[str, object] | None = None,
) -> SchedulerRunResult:
    run_started_at = datetime.now()
    run_started_perf = time.perf_counter()
    timestamped_output_dir = build_timestamped_output_dir(output_dir)
    scheduler_options = SchedulerAdvancedOptions.from_mapping(advanced_options)
    normalized_advanced_options = scheduler_options.as_dict()
    frozen_schedule_entries = (
        load_frozen_schedule_entries(scheduler_options.frozen_schedule_csv)
        + scheduler_options.manual_frozen_schedule_entries
    )
    calendar_capacity = load_calendar_capacity(
        scheduler_options.calendar_capacity_path,
        exclude_weekends=scheduler_options.exclude_weekends,
    )

    loader = WorkbookLoader(
        workbook_path,
        manufacturing_efficiency_path=scheduler_options.manufacturing_efficiency_path,
        filling_efficiency_path=scheduler_options.filling_efficiency_path,
    )
    MACHINE_COUNTS.update(loader.machine_counts)
    orders = loader.load_orders()
    source_links = loader.build_order_source_links()
    if not orders:
        raise RuntimeError("沒有載入到可排程的訂單資料。")

    schedulable_orders, anomalies = split_schedulable_orders(
        orders,
        exclude_warnings=exclude_warning_anomalies,
        independent_sections=scheduler_options.independent_sections,
        calendar_capacity=calendar_capacity,
    )
    if not schedulable_orders:
        raise RuntimeError("沒有可排程的訂單資料；請先修正日期窗衝突。")

    seeded_order_ids: tuple[str, ...] = ()
    if seed_schedule_csv is not None:
        seeded_order_ids = load_seed_order_ids(seed_schedule_csv)

    scheduler = GeneticScheduler(
        population_size=population_size,
        generations=generations,
        mutation_rate=mutation_rate,
        crossover_rate=crossover_rate,
        seed=seed,
        seeded_order_ids=seeded_order_ids,
        line_change_minutes=scheduler_options.line_change_minutes,
        frozen_schedule_entries=frozen_schedule_entries,
        calendar_capacity=calendar_capacity,
        independent_sections=scheduler_options.independent_sections,
        keep_split_days_contiguous=scheduler_options.keep_split_days_contiguous,
        break_start_time=scheduler_options.break_start_time,
        break_end_time=scheduler_options.break_end_time,
    )
    best_schedule = scheduler.optimize(schedulable_orders)

    write_schedule_outputs(best_schedule, timestamped_output_dir)
    write_daily_manufacturing_schedule_output(best_schedule, timestamped_output_dir)
    write_daily_packaging_schedule_output(best_schedule, timestamped_output_dir)
    write_manufacturing_cleaning_audit_output(best_schedule, timestamped_output_dir)
    write_filling_line_change_audit_output(best_schedule, timestamped_output_dir)
    write_frozen_schedule_audit_output(best_schedule, timestamped_output_dir)
    if scheduler_options.output_week_view:
        write_dispatch_week_view_output(best_schedule, orders, timestamped_output_dir)
    write_reference_format_output(best_schedule, orders, timestamped_output_dir)
    write_order_anomalies_output(orders, anomalies, timestamped_output_dir)
    validation_issues = validate_schedule(best_schedule, schedulable_orders)
    write_schedule_validation_output(validation_issues, timestamped_output_dir)
    write_source_link_output(source_links, timestamped_output_dir)
    write_input_model_issues_output(loader.model_build_issues, timestamped_output_dir)
    write_machine_counts_output(loader.machine_counts, timestamped_output_dir)
    write_order_model_audit_output(orders, source_links, timestamped_output_dir)

    run_finished_at = datetime.now()
    elapsed_seconds = time.perf_counter() - run_started_perf
    write_run_metadata(
        timestamped_output_dir,
        started_at=run_started_at,
        finished_at=run_finished_at,
        elapsed_seconds=elapsed_seconds,
        workbook_path=workbook_path,
        population_size=population_size,
        generations=generations,
        mutation_rate=mutation_rate,
        crossover_rate=crossover_rate,
        seed=seed,
        advanced_options=normalized_advanced_options,
    )
    write_developer_options_output(timestamped_output_dir, normalized_advanced_options)
    write_output_field_notes(timestamped_output_dir)
    write_intuitive_output_tables(
        best_schedule,
        orders,
        schedulable_orders,
        anomalies,
        validation_issues,
        timestamped_output_dir,
    )
    write_fixed_check_report_xlsx(
        best_schedule,
        orders,
        schedulable_orders,
        anomalies,
        validation_issues,
        source_links,
        timestamped_output_dir,
    )
    write_google_calendar_ics(best_schedule, orders, timestamped_output_dir)
    write_compact_excel_outputs(timestamped_output_dir)
    write_delivery_output_pack(timestamped_output_dir)

    return SchedulerRunResult(
        output_dir=timestamped_output_dir,
        best_schedule=best_schedule,
        orders=orders,
        schedulable_orders=schedulable_orders,
        anomalies=anomalies,
        validation_issues=validation_issues,
        source_links=source_links,
        elapsed_seconds=elapsed_seconds,
    )


def print_run_summary(result: SchedulerRunResult) -> None:
    excluded_anomalies = [anomaly for anomaly in result.anomalies if anomaly.exclude_from_optimization]
    warning_anomalies = [anomaly for anomaly in result.anomalies if not anomaly.exclude_from_optimization]
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
    print(f"原始訂單數量：{len(result.orders)}")
    print(f"實際參與 GA 的訂單數量：{len(result.schedulable_orders)}")
    print(f"異常訂單筆數：{len({anomaly.order_id for anomaly in result.anomalies})}")
    validation_error_count = sum(1 for issue in result.validation_issues if issue.severity == "error")
    validation_warning_count = sum(1 for issue in result.validation_issues if issue.severity == "warning")
    print(f"排程驗證硬性錯誤：{validation_error_count}")
    print(f"排程驗證人工提醒：{validation_warning_count}")
    print(f"本次執行時間：{_format_elapsed_hms(result.elapsed_seconds)} ({result.elapsed_seconds:.2f} 秒)")
    matched_source_links = sum(1 for row in result.source_links if row.get("match_count") == "1")
    print(f"來源訂單唯一對應數：{matched_source_links}/{len(result.source_links)}")
    print(f"總加班分鐘數：{_display_integer(result.best_schedule.total_overtime_minutes)}")
    print(f"日內未配置任務數：{result.best_schedule.intraday_unallocated_task_count}")
    print(f"非連續充填訂單數：{result.best_schedule.filling_noncontiguous_order_count}")
    print(f"充填換線分鐘數：{_display_integer(result.best_schedule.line_change_minutes)}")
    print(f"不可行懲罰：{result.best_schedule.infeasible_penalty:.1f}")
    print(f"目標懲罰：{result.best_schedule.target_penalty:.1f}")
    overtime_labels = {
        "manufacturing": "調製超時分鐘數",
        "filling": "充填超時分鐘數",
        "bottling": "排瓶超時分鐘數",
    }
    for labor_group in ("manufacturing", "filling", "bottling"):
        group_overtime = result.best_schedule.overtime_by_group.get(labor_group, 0.0)
        target = OVERTIME_TARGETS.get(labor_group)
        metric_name = overtime_labels[labor_group]
        if target is None:
            print(f"{metric_name}：{group_overtime:.1f}")
            continue
        print(f"{metric_name}：{group_overtime:.1f} / 目標 {target:.1f}")
    print(f"週末加班分鐘數：{result.best_schedule.weekend_overtime_minutes:.1f}")
    print(f"目標值：{result.best_schedule.objective:.1f}")
    print(f"輸出資料夾：{result.output_dir.resolve()}")


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
    parser.add_argument(
        "--line-change-minutes",
        type=int,
        default=0,
        help="同一天同一台充填機切換不同產品時，推估每次換線要加入的分鐘數。",
    )
    parser.add_argument(
        "--break-start",
        default=DEFAULT_BREAK_START_TIME,
        help="每日不可排班休息開始時間，格式 HH:MM，預設 12:00。",
    )
    parser.add_argument(
        "--break-end",
        default=DEFAULT_BREAK_END_TIME,
        help="每日不可排班休息結束時間，格式 HH:MM，預設 13:00。",
    )
    parser.add_argument(
        "--frozen-schedule-csv",
        type=Path,
        help="手動凍結排程 CSV，欄位可用 order_id/訂單、stage/工序、date/日期。",
    )
    parser.add_argument(
        "--calendar-capacity",
        type=Path,
        help="固定輸入格式 calendar_capacity.xlsx，用來排除假日並讀取每日製造/充填人力。",
    )
    parser.add_argument(
        "--allow-weekends",
        action="store_true",
        help="若 calendar_capacity 未標為假日，允許週六日排程。",
    )
    parser.add_argument(
        "--no-week-view",
        action="store_true",
        help="Do not generate dispatch_week_view.csv.",
    )
    parser.add_argument(
        "--manufacturing-efficiency",
        type=Path,
        default=None,
        help="Optional external workbook containing 製造效能_BOM回推 or 製造效能.",
    )
    parser.add_argument(
        "--filling-efficiency",
        type=Path,
        default=None,
        help="Optional external workbook containing 充填效能_BOM回推 or 充填效能.",
    )
    return parser.parse_args()

def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

    args = parse_args()
    result = run_scheduler(
        workbook_path=args.workbook,
        output_dir=args.output_dir,
        population_size=args.population_size,
        generations=args.generations,
        mutation_rate=args.mutation_rate,
        crossover_rate=args.crossover_rate,
        seed=args.seed,
        seed_schedule_csv=args.seed_schedule_csv,
        exclude_warning_anomalies=args.exclude_warning_anomalies,
        advanced_options={
            "schedule_mode": "full_ga",
            "line_change_minutes": args.line_change_minutes,
            "break_start_time": args.break_start,
            "break_end_time": args.break_end,
            "frozen_schedule_csv": str(args.frozen_schedule_csv) if args.frozen_schedule_csv else "",
            "calendar_capacity_path": str(args.calendar_capacity) if args.calendar_capacity else "",
            "manufacturing_efficiency_path": str(args.manufacturing_efficiency) if args.manufacturing_efficiency else "",
            "filling_efficiency_path": str(args.filling_efficiency) if args.filling_efficiency else "",
            "exclude_weekends": not args.allow_weekends,
            "output_week_view": not args.no_week_view,
        },
    )
    print_run_summary(result)

if __name__ == "__main__":
    main()
