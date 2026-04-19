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


def write_reference_format_output(schedule: ScheduleResult, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    summary = build_schedule_summary(schedule)
    grouped_tasks = _group_output_tasks(schedule.task_records)

    manufacturing_group_totals: dict[tuple[str, str], int] = defaultdict(int)
    packaging_group_totals: dict[tuple[str, str], float] = defaultdict(float)
    manufacturing_group_ids: dict[tuple[str, str], int] = {}
    machine_day_slot_cursor: dict[tuple[date, str], int] = defaultdict(int)

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
                demand_units = manufacturing_group_totals[group_key]
                arranged_units = 1
                planned_people = "1"
                subtask_id = f"{task.order_id}_{manufacturing_group_ids[group_key]}"
                unit_minutes = _display_integer(task.machine_minutes)
            else:
                seq = "18" if task.task_type == "Cat3" else "19"
                demand_units = _display_integer(packaging_group_totals[(task.order_id, task.task_type)])
                arranged_units = _display_integer(task.machine_minutes / 5.0)
                planned_people = "1" if task.task_type == "Cat3" else "4"
                subtask_id = f"{task.order_id}_{task.task_type}_{seq}"
                unit_minutes = 5

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
    write_reference_format_output(best_schedule, args.output_dir)
    write_order_anomalies_output(orders, anomalies, args.output_dir)
    validation_issues = validate_schedule(best_schedule, schedulable_orders)
    write_schedule_validation_output(validation_issues, args.output_dir)
    write_source_link_output(source_links, args.output_dir)
    write_machine_counts_output(loader.machine_counts, args.output_dir)
    write_order_model_audit_output(orders, source_links, args.output_dir)

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
