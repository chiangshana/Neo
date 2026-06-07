"""Run a fresh demo schedule and audit whether the output is presentation-ready."""

from __future__ import annotations

import argparse
import csv
import re
import sys
from datetime import date
from pathlib import Path

from genetic_algorithm import (
    SimpleXlsxReader,
    load_calendar_capacity,
    print_run_summary,
    run_scheduler,
)


DELIVERY_DIR_NAME = "交付用輸出"
DELIVERY_FILES = [
    Path(DELIVERY_DIR_NAME) / "00_輸出閱讀順序.csv",
    Path(DELIVERY_DIR_NAME) / "README_先看我.md",
    Path(DELIVERY_DIR_NAME) / "01_排程檢查報告.xlsx",
    Path(DELIVERY_DIR_NAME) / "02_當日排程確切時間.xlsx",
    Path(DELIVERY_DIR_NAME) / "03_Google日曆匯入檔.ics",
    Path(DELIVERY_DIR_NAME) / "04_輔助檢查資料.xlsx",
]
INPUT_SHEET_ALIASES = {
    "訂單資訊": ("2023訂單資訊", "X2023訂單資訊含入庫日"),
    "BOM": ("BOM",),
    "製造效能": ("製造效能_BOM回推", "製造效能"),
    "充填效能": ("充填效能_BOM回推", "充填效能"),
}
BUSINESS_CSVS = [
    "schedule_summary.csv",
    "best_schedule_tasks.csv",
    "best_schedule_reference_format.csv",
    "best_schedule_daily_load.csv",
    "daily_manufacturing_schedule.csv",
    "daily_packaging_schedule.csv",
    "dispatch_week_view.csv",
]
HISTORICAL_OVERTIME_REFERENCES = [
    ("舊版可交付檢查 ga_end_to_end_check_20260521_234103", 1514),
    ("週報草稿正式版紀錄", 1774),
    ("早期輸入輸出檢查 ga_demo_input_output_check_20260524_002620", 2811),
    ("全部合理訂單排入且避開午休 ga_demo_all_reasonable_break_check_20260526_182626", 1607),
]
EXPECTED_FORMAL_OVERTIME_RANGE = (1000, 3000)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="從固定輸入重新跑 demo，並檢查輸出是否適合展示。")
    parser.add_argument("--workbook", type=Path, default=Path("workbook_bom_refresh.xlsx"))
    parser.add_argument("--calendar-capacity", type=Path, default=Path("..") / "calendar_capacity.xlsx")
    parser.add_argument("--output-dir", type=Path, default=Path("ga_present_demo_check"))
    parser.add_argument("--population-size", type=int, default=20)
    parser.add_argument("--generations", type=int, default=40)
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--manufacturing-efficiency", type=Path, default=None)
    parser.add_argument("--filling-efficiency", type=Path, default=None)
    parser.add_argument("--break-start", default="12:00")
    parser.add_argument("--break-end", default="13:00")
    return parser.parse_args()


def read_csv_rows(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file))


def read_metric_csv(path: Path) -> dict[str, str]:
    metrics: dict[str, str] = {}
    for row in read_csv_rows(path):
        metrics[row.get("metric", "")] = row.get("value", "")
    return metrics


def sheet_exists(available: list[str], aliases: tuple[str, ...]) -> bool:
    normalized_available = {
        SimpleXlsxReader._normalize_sheet_name(sheet_name) for sheet_name in available
    }
    return any(SimpleXlsxReader._normalize_sheet_name(alias) in normalized_available for alias in aliases)


def count_decimal_cells(output_dir: Path) -> list[tuple[str, int, str, str]]:
    decimal_pattern = re.compile(r"^-?\d+\.\d+$")
    decimal_cells: list[tuple[str, int, str, str]] = []
    for relative_path in BUSINESS_CSVS:
        path = output_dir / relative_path
        if not path.exists():
            continue
        for row_index, row in enumerate(read_csv_rows(path), start=2):
            for column_name, value in row.items():
                text = (value or "").strip()
                if decimal_pattern.match(text):
                    decimal_cells.append((relative_path, row_index, column_name, text))
    return decimal_cells


def parse_hhmm_minutes(value: str) -> int | None:
    try:
        hour_text, minute_text = str(value).strip().split(":", 1)
        hour = int(hour_text)
        minute = int(minute_text)
    except ValueError:
        return None
    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        return None
    return hour * 60 + minute


def time_range_overlaps_break(time_range: str, break_start: int, break_end: int) -> bool:
    for segment in str(time_range or "").split(";"):
        time_pairs = re.findall(r"(?:\+\d+天\s*)?(\d{2}):(\d{2})", segment)
        if len(time_pairs) < 2:
            continue
        start = int(time_pairs[0][0]) * 60 + int(time_pairs[0][1])
        end = int(time_pairs[1][0]) * 60 + int(time_pairs[1][1])
        intervals = [(start, end)] if end > start else [(start, 24 * 60), (0, end)]
        if any(interval_start < break_end and interval_end > break_start for interval_start, interval_end in intervals):
            return True
    return False


def audit_output(
    output_dir: Path,
    workbook_path: Path,
    calendar_path: Path,
    break_start_time: str,
    break_end_time: str,
) -> tuple[list[str], list[str]]:
    failures: list[str] = []
    warnings: list[str] = []

    if not workbook_path.exists():
        failures.append(f"找不到訂單/效能 workbook：{workbook_path}")
    else:
        workbook_sheets = SimpleXlsxReader(workbook_path).list_sheet_names()
        missing_sources = [
            label for label, aliases in INPUT_SHEET_ALIASES.items() if not sheet_exists(workbook_sheets, aliases)
        ]
        if missing_sources:
            failures.append(f"輸入 workbook 缺少必要資料來源：{', '.join(missing_sources)}")

    if not calendar_path.exists():
        failures.append(f"找不到 calendar_capacity.xlsx：{calendar_path}")
    else:
        calendar_sheets = SimpleXlsxReader(calendar_path).list_sheet_names()
        if not sheet_exists(calendar_sheets, ("calendar_capacity",)):
            failures.append("calendar_capacity.xlsx 缺少 calendar_capacity 工作表")

    missing_delivery = [str(path) for path in DELIVERY_FILES if not (output_dir / path).exists()]
    if missing_delivery:
        failures.append(f"交付用輸出缺少檔案：{', '.join(missing_delivery)}")

    validation_summary = read_metric_csv(output_dir / "schedule_validation_summary.csv")
    error_count = int(validation_summary.get("error_count", "999"))
    blocking_count = int(validation_summary.get("blocking_issue_count", "999"))
    warning_count = int(validation_summary.get("warning_count", "0"))
    if error_count:
        failures.append(f"排程驗證仍有硬性錯誤：{error_count}")
    if blocking_count:
        failures.append(f"交付阻擋問題仍存在：{blocking_count}")
    if warning_count:
        warnings.append(f"排程驗證有人工提醒：{warning_count}")

    schedule_summary = read_csv_rows(output_dir / "schedule_summary.csv")[0]
    summary_total_overtime = int(schedule_summary.get("總加班分鐘數", "0"))
    expected_min, expected_max = EXPECTED_FORMAL_OVERTIME_RANGE
    if not expected_min <= summary_total_overtime <= expected_max:
        failures.append(
            f"正式版總加班分鐘數 {summary_total_overtime} 不在合理驗收區間 "
            f"{expected_min}-{expected_max}；請檢查加班口徑是否又被放寬。"
        )
    if int(schedule_summary.get("日內未配置任務數", "999")):
        failures.append(f"仍有日內未配置任務：{schedule_summary.get('日內未配置任務數')}")
    if int(schedule_summary.get("週末加班分鐘數", "0")):
        failures.append(f"週末/假日仍有加班分鐘：{schedule_summary.get('週末加班分鐘數')}")

    break_start = parse_hhmm_minutes(break_start_time)
    break_end = parse_hhmm_minutes(break_end_time)
    if break_start is None or break_end is None or break_end <= break_start:
        failures.append(f"休息時間設定不合法：{break_start_time}-{break_end_time}")
    else:
        break_overlaps: list[tuple[str, str, str, str]] = []
        for row in read_csv_rows(output_dir / "best_schedule_reference_format.csv"):
            time_range = row.get("時段範圍", "")
            if time_range_overlaps_break(time_range, break_start, break_end):
                break_overlaps.append(
                    (
                        row.get("訂單編號", ""),
                        row.get("作業工序", ""),
                        row.get("排程日期", ""),
                        time_range,
                    )
                )
        if break_overlaps:
            sample = ", ".join(
                f"{order}/{stage}/{day}/{time_range}"
                for order, stage, day, time_range in break_overlaps[:5]
            )
            failures.append(
                f"有任務排到休息時間 {break_start_time}-{break_end_time}："
                f"{len(break_overlaps)} 筆，例如 {sample}"
            )

    if calendar_path.exists():
        calendar_capacity = load_calendar_capacity(calendar_path, exclude_weekends=True)
        task_on_nonworkday: list[tuple[str, str, str]] = []
        for row in read_csv_rows(output_dir / "best_schedule_tasks.csv"):
            scheduled_date = row.get("日期", "")
            if scheduled_date and not calendar_capacity.is_workday(date.fromisoformat(scheduled_date)):
                task_on_nonworkday.append((row.get("訂單", ""), row.get("類別", ""), scheduled_date))
        if task_on_nonworkday:
            sample = ", ".join(f"{order}/{stage}/{day}" for order, stage, day in task_on_nonworkday[:5])
            failures.append(f"有任務排到非工作日：{len(task_on_nonworkday)} 筆，例如 {sample}")

    decimal_cells = count_decimal_cells(output_dir)
    if decimal_cells:
        sample = ", ".join(
            f"{file_name}:{row_index}:{column_name}={value}"
            for file_name, row_index, column_name, value in decimal_cells[:5]
        )
        failures.append(f"業主主要 CSV 仍有小數格式：{len(decimal_cells)} 格，例如 {sample}")

    daily_load_overtime = 0
    for row in read_csv_rows(output_dir / "best_schedule_daily_load.csv"):
        daily_load_overtime += int(row.get("超時分鐘", "0") or 0)
    if abs(daily_load_overtime - summary_total_overtime) > 1:
        failures.append(
            f"摘要總加班 {summary_total_overtime} 與每日負荷加總 {daily_load_overtime} 不一致。"
        )

    issue_types = {
        row.get("issue_type", ""): 0 for row in read_csv_rows(output_dir / "schedule_validation.csv")
    }
    for row in read_csv_rows(output_dir / "schedule_validation.csv"):
        issue_type = row.get("issue_type", "")
        issue_types[issue_type] = issue_types.get(issue_type, 0) + 1
    if issue_types and set(issue_types) != {"filling_days_not_contiguous"}:
        warnings.append(f"驗證提醒種類不只充填非連續：{issue_types}")

    ics_path = output_dir / DELIVERY_DIR_NAME / "03_Google日曆匯入檔.ics"
    if ics_path.exists():
        event_count = ics_path.read_text(encoding="utf-8").count("BEGIN:VEVENT")
        if event_count == 0:
            failures.append("Google 日曆 ICS 沒有任何事件")
    return failures, warnings


def print_audit_summary(output_dir: Path, failures: list[str], warnings: list[str]) -> None:
    validation_summary = read_metric_csv(output_dir / "schedule_validation_summary.csv")
    schedule_summary = read_csv_rows(output_dir / "schedule_summary.csv")[0]
    excluded_ids: list[str] = []
    for row in read_csv_rows(output_dir / "order_anomalies_exclude.csv"):
        order_id = row.get("order_id", "")
        if order_id and order_id not in excluded_ids:
            excluded_ids.append(order_id)

    print("\nDemo 驗收摘要")
    print(f"輸出資料夾：{output_dir.resolve()}")
    print(f"error_count：{validation_summary.get('error_count')}")
    print(f"blocking_issue_count：{validation_summary.get('blocking_issue_count')}")
    print(f"warning_count：{validation_summary.get('warning_count')}")
    print(f"總加班分鐘數：{schedule_summary.get('總加班分鐘數')}")
    print(f"日內未配置任務數：{schedule_summary.get('日內未配置任務數')}")
    print(f"排除訂單：{', '.join(excluded_ids) if excluded_ids else '無'}")
    print(f"交付資料夾：{(output_dir / DELIVERY_DIR_NAME).resolve()}")
    print("加班口徑提醒：")
    print("- 這次使用保守正式口徑：summary 必須與每日負荷表加總一致。")
    print("- 歷史合理區間參考：")
    for label, overtime_minutes in HISTORICAL_OVERTIME_REFERENCES:
        print(f"  {label}：{overtime_minutes} 分鐘")
    print("- 若結果低於 1000 分鐘，通常代表加班口徑被放太鬆，腳本會直接 FAIL。")
    for warning in warnings:
        print(f"提醒：{warning}")
    if failures:
        print("\nFAIL：這份輸出不建議直接 demo")
        for failure in failures:
            print(f"- {failure}")
    else:
        print("\nPASS：這份輸出可作為 demo 候選結果")


def main() -> int:
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
        seed=args.seed,
        advanced_options={
            "schedule_mode": "full_ga",
            "calendar_capacity_path": str(args.calendar_capacity),
            "exclude_weekends": True,
            "output_week_view": True,
            "line_change_minutes": 0,
            "manufacturing_efficiency_path": str(args.manufacturing_efficiency or ""),
            "filling_efficiency_path": str(args.filling_efficiency or ""),
            "break_start_time": args.break_start,
            "break_end_time": args.break_end,
        },
    )
    print_run_summary(result)
    failures, warnings = audit_output(
        result.output_dir,
        args.workbook,
        args.calendar_capacity,
        args.break_start,
        args.break_end,
    )
    print_audit_summary(result.output_dir, failures, warnings)
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
