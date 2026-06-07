"""Tkinter GUI for the factory GA scheduler.

這個介面保留原本 `genetic_algorithm.py` 的命令列流程，只負責把
輸入資料、GA 參數、輸出格式與執行結果拆成比較好操作的畫面。
"""

from __future__ import annotations

import os
import csv
import json
import queue
import subprocess
import sys
import threading
import tkinter as tk
from contextlib import redirect_stdout
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


PROJECT_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
BUILD_SCRIPT = PROJECT_DIR / "build_installer.py"

OUTPUT_FILES = [
    ("交付用輸出/01_排程檢查報告.xlsx", "固定格式 Excel 檢查報告，對齊業主提供的檢查表。"),
    ("交付用輸出/02_當日排程確切時間.xlsx", "日內機台/人力時段表；僅在驗證無阻擋錯誤時可直接派工。"),
    ("交付用輸出/03_Google日曆匯入檔.ics", "可匯入 Google Calendar 的排程事件檔。"),
    ("交付用輸出/04_輔助檢查資料.xlsx", "一頁式總覽、驗證摘要、瓶頸、異常、稽核與欄位備註集中在同一本 Excel。"),
    ("交付用輸出/README_先看我.md", "交付資料夾的閱讀說明。"),
]

INPUT_FORMAT_GROUPS = [
    (
        "訂單時間",
        "時間準備表V1、X2023訂單資訊含入庫日 或 2023訂單資訊",
        "訂單編號、產品名稱、數量、交期、原料齊料日、包材/物料齊料日",
        "決定每張訂單可以從哪天開始做，以及最晚必須何時完成。",
    ),
    (
        "製造效能",
        "製造效能_BOM回推 或 製造效能",
        "核心六碼/產品代號、攪料程序、製程設備、製造批次、預攪分鐘、攪料分鐘、清洗分鐘",
        "決定預攪、攪料清洗要花多久，以及使用哪一台調製設備。",
    ),
    (
        "充填效能",
        "充填效能_BOM回推 或 充填效能",
        "核心六碼/產品代號、工作類型、機器、排瓶時數、充填時數、表定產能人數",
        "決定排瓶與充填要花多久，以及包裝段人力與設備負荷。",
    ),
]


class QueueTextWriter:
    def __init__(self, message_queue: queue.Queue[tuple[str, object]]) -> None:
        self.message_queue = message_queue

    def write(self, text: str) -> int:
        if text:
            self.message_queue.put(("log", text))
        return len(text)

    def flush(self) -> None:
        pass


class SchedulerGui(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("工廠訂單排程 GA 工具")
        self.geometry("1120x760")
        self.minsize(980, 680)

        self.message_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.running_thread: threading.Thread | None = None
        self.last_output_dir: Path | None = None
        self.last_validation_errors: int | None = None
        self.last_validation_warnings: int | None = None

        default_workbook = PROJECT_DIR / "workbook_bom_refresh.xlsx"
        self.workbook_var = tk.StringVar(value=str(default_workbook if default_workbook.exists() else ""))
        self.seed_schedule_var = tk.StringVar()
        self.output_dir_var = tk.StringVar(value=str(PROJECT_DIR / "ga_gui_output"))
        self.population_size_var = tk.StringVar(value="40")
        self.generations_var = tk.StringVar(value="120")
        self.mutation_rate_var = tk.StringVar(value="0.15")
        self.crossover_rate_var = tk.StringVar(value="0.85")
        self.seed_var = tk.StringVar(value="42")
        self.exclude_warning_var = tk.BooleanVar(value=False)
        self.data_source_mode_var = tk.StringVar(value="本機 Excel")
        self.google_sheet_id_var = tk.StringVar()
        self.schedule_mode_var = tk.StringVar(value="完整連動排程（製造先於充填，排瓶於充填前一工作日）")
        self.manufacturing_efficiency_path_var = tk.StringVar()
        self.filling_efficiency_path_var = tk.StringVar()
        calendar_candidates = (PROJECT_DIR / "calendar_capacity.xlsx", PROJECT_DIR.parent / "calendar_capacity.xlsx")
        default_calendar = next((path for path in calendar_candidates if path.exists()), calendar_candidates[0])
        self.calendar_capacity_var = tk.StringVar(value=str(default_calendar if default_calendar.exists() else ""))
        self.exclude_weekends_var = tk.BooleanVar(value=True)
        self.frozen_schedule_var = tk.StringVar()
        self.freeze_order_var = tk.StringVar()
        self.freeze_stage_var = tk.StringVar(value="充填")
        self.freeze_date_var = tk.StringVar()
        self.freeze_time_var = tk.StringVar()
        self.daily_rerun_time_var = tk.StringVar(value="16:00")
        self.line_change_minutes_var = tk.StringVar(value="0")
        self.break_start_time_var = tk.StringVar(value="12:00")
        self.break_end_time_var = tk.StringVar(value="13:00")
        self.split_unit_var = tk.StringVar(value="10")
        self.minimum_fill_liters_var = tk.StringVar(value="50")
        self.keep_split_days_contiguous_var = tk.BooleanVar(value=True)
        self.small_printer_count_var = tk.StringVar(value="1")
        self.normal_printer_count_var = tk.StringVar(value="1")
        self.output_week_view_var = tk.BooleanVar(value=True)
        self.preview_var = tk.StringVar(value="尚未檢查輸入資料。")
        self.status_var = tk.StringVar(value="待命")
        self.output_dir_status_var = tk.StringVar(value="尚未產生輸出資料夾。")
        self.developer_summary_var = tk.StringVar(value="尚未刷新下一版設定摘要。")
        self.overview_summary_var = tk.StringVar(value="")
        self.next_step_var = tk.StringVar(value="建議先檢查輸入資料，再執行 quick run。")

        self._configure_style()
        self._build_layout()
        self._attach_overview_traces()
        self._refresh_overview()
        self.after(100, self._poll_messages)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Microsoft JhengHei UI", 17, "bold"), foreground="#1f2d2f")
        style.configure("Subtitle.TLabel", font=("Microsoft JhengHei UI", 10), foreground="#52676d")
        style.configure("Section.TLabelframe.Label", font=("Microsoft JhengHei UI", 10, "bold"))
        style.configure("Primary.TButton", font=("Microsoft JhengHei UI", 10, "bold"))
        style.configure("Muted.TLabel", foreground="#66757a")
        style.configure("Status.TLabel", font=("Microsoft JhengHei UI", 10, "bold"), foreground="#24505a")

    def _build_layout(self) -> None:
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=12)

        self.home_tab = ttk.Frame(self.notebook, padding=12)
        self.input_tab = ttk.Frame(self.notebook, padding=12)
        self.params_tab = ttk.Frame(self.notebook, padding=12)
        self.output_tab = ttk.Frame(self.notebook, padding=12)
        self.developer_tab = ttk.Frame(self.notebook, padding=12)
        self.run_tab = ttk.Frame(self.notebook, padding=12)
        self.package_tab = ttk.Frame(self.notebook, padding=12)

        self.notebook.add(self.home_tab, text="總覽")
        self.notebook.add(self.input_tab, text="1. 輸入產能日曆")
        self.notebook.add(self.params_tab, text="2. GA 參數")
        self.notebook.add(self.output_tab, text="3. 輸出格式")
        self.notebook.add(self.developer_tab, text="4. 開發者設定")
        self.notebook.add(self.run_tab, text="5. 執行結果")
        self.notebook.add(self.package_tab, text="6. 打包安裝包")

        self._build_home_tab(self.home_tab)
        self._build_input_tab(self.input_tab)
        self._build_params_tab(self.params_tab)
        self._build_output_tab(self.output_tab)
        self._build_developer_tab(self.developer_tab)
        self._build_run_tab(self.run_tab)
        self._build_package_tab(self.package_tab)

    def _build_home_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(3, weight=1)

        header = ttk.Frame(parent)
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="工廠訂單排程控制台", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="從輸入檢查、GA 調參、進階限制到交付輸出，照流程走就能完成一次排程。",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        status_card = ttk.LabelFrame(parent, text="目前狀態", padding=12, style="Section.TLabelframe")
        status_card.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=8)
        status_card.columnconfigure(0, weight=1)
        ttk.Label(status_card, textvariable=self.status_var, style="Status.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(status_card, textvariable=self.output_dir_status_var, style="Muted.TLabel", wraplength=470).grid(
            row=1, column=0, sticky="w", pady=(6, 0)
        )
        ttk.Label(status_card, textvariable=self.next_step_var, wraplength=470).grid(row=2, column=0, sticky="w", pady=(10, 0))

        summary_card = ttk.LabelFrame(parent, text="本次設定摘要", padding=12, style="Section.TLabelframe")
        summary_card.grid(row=1, column=1, sticky="nsew", padx=(8, 0), pady=8)
        summary_card.columnconfigure(0, weight=1)
        ttk.Label(summary_card, textvariable=self.overview_summary_var, justify="left", wraplength=470).grid(
            row=0, column=0, sticky="w"
        )

        actions = ttk.LabelFrame(parent, text="常用操作", padding=12, style="Section.TLabelframe")
        actions.grid(row=2, column=0, columnspan=2, sticky="ew", pady=8)
        for column in range(7):
            actions.columnconfigure(column, weight=1)
        ttk.Button(actions, text="檢查輸入資料", command=self._start_preview, style="Primary.TButton").grid(
            row=0, column=0, sticky="ew", padx=4, pady=4
        )
        ttk.Button(actions, text="Quick Run", command=self._quick_run_from_home).grid(row=0, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(actions, text="Demo 參數", command=self._demo_preset_from_home).grid(row=0, column=2, sticky="ew", padx=4, pady=4)
        ttk.Button(actions, text="正式版參數", command=self._formal_preset_from_home).grid(row=0, column=3, sticky="ew", padx=4, pady=4)
        ttk.Button(actions, text="執行 GA", command=self._start_run, style="Primary.TButton").grid(
            row=0, column=4, sticky="ew", padx=4, pady=4
        )
        ttk.Button(actions, text="開交付輸出", command=self._open_delivery_dir).grid(row=0, column=5, sticky="ew", padx=4, pady=4)
        ttk.Button(actions, text="開日曆檔", command=self._open_calendar_file).grid(row=0, column=6, sticky="ew", padx=4, pady=4)

        workflow = ttk.LabelFrame(parent, text="建議流程", padding=12, style="Section.TLabelframe")
        workflow.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        workflow.columnconfigure(1, weight=1)
        steps = [
            ("1", "輸入產能日曆", "選擇 calendar_capacity.xlsx；訂單、製造效能、充填效能使用內建基礎資料。", self.input_tab),
            ("2", "GA 參數", "先用 quick run 檢查格式，再用正式版參數產出交付結果。", self.params_tab),
            ("3", "開發者設定", "設定換線、凍結排程、派工週視圖與下一版限制入口。", self.developer_tab),
            ("4", "執行與交付", "執行後優先看交付用輸出、一頁式總覽、驗證摘要與週視圖。", self.run_tab),
        ]
        for row, (number, title, description, tab) in enumerate(steps):
            ttk.Label(workflow, text=number, style="Status.TLabel", width=3).grid(row=row, column=0, sticky="n", pady=6)
            ttk.Label(workflow, text=f"{title}：{description}", wraplength=800).grid(
                row=row, column=1, sticky="w", pady=6
            )
            ttk.Button(workflow, text="前往", command=lambda tab=tab: self.notebook.select(tab)).grid(
                row=row, column=2, sticky="e", padx=(8, 0), pady=6
            )

    def _attach_overview_traces(self) -> None:
        watched_vars: tuple[tk.Variable, ...] = (
            self.workbook_var,
            self.manufacturing_efficiency_path_var,
            self.filling_efficiency_path_var,
            self.output_dir_var,
            self.population_size_var,
            self.generations_var,
            self.mutation_rate_var,
            self.crossover_rate_var,
            self.seed_var,
            self.line_change_minutes_var,
            self.frozen_schedule_var,
            self.calendar_capacity_var,
            self.exclude_weekends_var,
            self.output_week_view_var,
        )
        for variable in watched_vars:
            variable.trace_add("write", lambda *_args: self._refresh_overview())

    def _refresh_overview(self) -> None:
        workbook_text = self.workbook_var.get().strip()
        workbook_status = "未找到內建基礎資料"
        if workbook_text:
            workbook_status = "內建 OK" if Path(workbook_text).expanduser().exists() else "找不到內建基礎資料"
        frozen_status = "已指定" if self.frozen_schedule_var.get().strip() else "未指定"
        manual_freeze_count = len(getattr(self, "frozen_tree", ()).get_children()) if hasattr(self, "frozen_tree") else 0
        calendar_text = self.calendar_capacity_var.get().strip()
        calendar_status = "未選擇"
        if calendar_text:
            calendar_status = "OK" if Path(calendar_text).expanduser().exists() else "找不到檔案"
        week_view_status = "輸出" if self.output_week_view_var.get() else "不輸出"
        self.overview_summary_var.set(
            "\n".join(
                [
                    f"產能日曆 calendar_capacity：{calendar_status}；排除週六日：{'是' if self.exclude_weekends_var.get() else '否'}",
                    f"訂單基礎資料：{workbook_status}",
                    f"外部效能表：製造 {'已指定' if self.manufacturing_efficiency_path_var.get().strip() else '使用內建'}；充填 {'已指定' if self.filling_efficiency_path_var.get().strip() else '使用內建'}",
                    f"輸出基底：{self.output_dir_var.get()}",
                    f"GA：population={self.population_size_var.get()}；generations={self.generations_var.get()}；seed={self.seed_var.get()}",
                    f"交配/突變：crossover={self.crossover_rate_var.get()}；mutation={self.mutation_rate_var.get()}",
                    f"進階：換線 {self.line_change_minutes_var.get()} 分鐘；凍結 CSV {frozen_status}；手動凍結 {manual_freeze_count} 筆；週視圖 {week_view_status}",
                ]
            )
        )
        if self.last_output_dir and self.last_validation_errors:
            self.next_step_var.set(
                f"已產生輸出，但仍有 {self.last_validation_errors} 項阻擋錯誤。請先看 04_輔助檢查資料，02 不可直接派工。"
            )
        elif self.last_output_dir:
            self.next_step_var.set("已產生輸出。建議先開啟 01_排程檢查報告，再看 04_輔助檢查資料的驗證摘要。")
        elif calendar_status == "OK" and workbook_status == "內建 OK":
            self.next_step_var.set("產能日曆已就緒。建議先檢查輸入資料，或直接做 quick run。")
        elif workbook_status != "內建 OK":
            self.next_step_var.set("找不到內建基礎資料 workbook_bom_refresh.xlsx，請確認程式資料夾內容完整。")
        else:
            self.next_step_var.set("請先在「輸入產能日曆」選擇有效的 calendar_capacity.xlsx。")

    def _quick_run_from_home(self) -> None:
        self._use_quick_preset()
        self.notebook.select(self.run_tab)
        self._start_run()

    def _demo_preset_from_home(self) -> None:
        self._use_demo_preset()
        self.notebook.select(self.params_tab)
        self._refresh_overview()

    def _formal_preset_from_home(self) -> None:
        self._use_formal_preset()
        self.notebook.select(self.params_tab)
        self._refresh_overview()

    def _build_input_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(8, weight=1)

        ttk.Label(parent, text="產能日曆 Excel：").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(parent, textvariable=self.calendar_capacity_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(parent, text="選擇 calendar_capacity", command=self._choose_calendar_capacity).grid(row=0, column=2, pady=6)

        ttk.Checkbutton(parent, text="排除週六日", variable=self.exclude_weekends_var).grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(2, 4)
        )

        base_data_note = (
            "訂單資訊、製造效能與充填效能使用程式內建的 workbook_bom_refresh.xlsx；"
            "業主日常操作只需要選擇 calendar_capacity.xlsx。"
        )
        ttk.Label(parent, text=base_data_note, foreground="#555555", wraplength=900).grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(2, 8)
        )

        ttk.Label(parent, text="Seed 排程 CSV（選填）：").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Entry(parent, textvariable=self.seed_schedule_var).grid(row=3, column=1, sticky="ew", padx=6)
        ttk.Button(parent, text="選擇 CSV", command=self._choose_seed_schedule).grid(row=3, column=2, pady=6)

        ttk.Checkbutton(
            parent,
            text="warning 類型異常訂單也排除出 GA 主優化集合",
            variable=self.exclude_warning_var,
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 4))

        ttk.Button(parent, text="檢查輸入資料", command=self._start_preview).grid(row=5, column=0, sticky="w", pady=8)
        ttk.Button(parent, text="匯出輸入格式說明", command=self._export_input_format_guide).grid(
            row=5, column=1, sticky="w", padx=6, pady=8
        )

        preview = ttk.LabelFrame(parent, text="輸入檢查結果", padding=10)
        preview.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=8)
        preview.columnconfigure(0, weight=1)
        ttk.Label(preview, textvariable=self.preview_var, justify="left").grid(row=0, column=0, sticky="w")

        note = (
            "建議流程：先選 calendar_capacity.xlsx，再按「檢查輸入資料」。\n"
            "如果排除異常後仍有可排訂單，就可以切到 GA 參數頁開始調整。"
        )
        ttk.Label(parent, text=note, foreground="#555555").grid(row=7, column=0, columnspan=3, sticky="w", pady=12)

        input_format = ttk.LabelFrame(parent, text="內建基礎資料格式（開發者維護；業主日常只需提供 calendar_capacity）", padding=8)
        input_format.grid(row=8, column=0, columnspan=3, sticky="nsew", pady=(4, 0))
        input_format.columnconfigure(0, weight=1)
        input_format.rowconfigure(0, weight=1)

        columns = ("group", "sheets", "fields", "purpose")
        tree = ttk.Treeview(input_format, columns=columns, show="headings", height=5)
        tree.heading("group", text="資料類型")
        tree.heading("sheets", text="建議工作表")
        tree.heading("fields", text="必要欄位概念")
        tree.heading("purpose", text="用途")
        tree.column("group", width=100, anchor="w")
        tree.column("sheets", width=250, anchor="w")
        tree.column("fields", width=360, anchor="w")
        tree.column("purpose", width=260, anchor="w")
        tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(input_format, orient="vertical", command=tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)
        for row in INPUT_FORMAT_GROUPS:
            tree.insert("", "end", values=row)

    def _build_params_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)

        fields = [
            ("族群大小 population_size", self.population_size_var, "越大搜尋越廣，但執行時間也會增加。"),
            ("世代數 max_generations", self.generations_var, "越大越有機會改善解，但會跑更久。"),
            ("突變率 mutation_rate", self.mutation_rate_var, "建議 0.05 到 0.30；太高會讓搜尋不穩。"),
            ("交配率 crossover_rate", self.crossover_rate_var, "建議 0.70 到 0.95；用來交換兩個排序方案。"),
            ("隨機種子 seed", self.seed_var, "固定 seed 可重現正式版結果。"),
        ]

        for row, (label, variable, help_text) in enumerate(fields):
            ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=7)
            ttk.Entry(parent, textvariable=variable, width=16).grid(row=row, column=1, sticky="w", padx=6)
            ttk.Label(parent, text=help_text, foreground="#555555").grid(row=row, column=2, sticky="w", pady=7)

        button_row = len(fields)
        ttk.Button(parent, text="套用 quick run 參數", command=self._use_quick_preset).grid(
            row=button_row, column=0, sticky="w", pady=14
        )
        ttk.Button(parent, text="套用 demo 參數", command=self._use_demo_preset).grid(
            row=button_row, column=1, sticky="w", pady=14
        )
        ttk.Button(parent, text="套用正式版參數", command=self._use_formal_preset).grid(
            row=button_row, column=2, sticky="w", pady=14
        )

        explanation = ttk.LabelFrame(parent, text="參數調整建議", padding=10)
        explanation.grid(row=button_row + 1, column=0, columnspan=3, sticky="ew", pady=8)
        ttk.Label(
            explanation,
            text=(
                "quick run：適合檢查格式與資料是否可跑。\n"
                "demo：20 個體、40 代，採保守正式加班口徑，適合 present 前驗收。\n"
                "正式版：適合交付與報告，會跑完整 120 世代。\n"
                "如果想追求更低加班，可先固定 seed，再逐步增加 generations 或 population_size 比較結果。"
            ),
            justify="left",
        ).grid(row=0, column=0, sticky="w")

    def _build_output_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(2, weight=1)

        ttk.Label(parent, text="輸出資料夾基底：").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(parent, textvariable=self.output_dir_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(parent, text="選擇位置", command=self._choose_output_base).grid(row=0, column=2, pady=6)

        ttk.Label(
            parent,
            text="程式會自動在資料夾名稱後加上時間戳記，避免覆蓋前一次結果。",
            foreground="#555555",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 8))

        table_frame = ttk.LabelFrame(parent, text="完整輸出格式", padding=8)
        table_frame.grid(row=2, column=0, columnspan=3, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = ("filename", "purpose")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=13)
        tree.heading("filename", text="檔案")
        tree.heading("purpose", text="用途")
        tree.column("filename", width=250, anchor="w")
        tree.column("purpose", width=620, anchor="w")
        tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)

        for filename, purpose in OUTPUT_FILES:
            tree.insert("", "end", values=(filename, purpose))

    def _build_developer_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(6, weight=1)
        parent.rowconfigure(8, weight=1)

        intro = (
            "這個頁籤把業主規則與情境測試做成可視化設定面板。\n"
            "正式排程固定為製造先於充填、排瓶在充填前一個工作日；這裡主要調整企業資料表、產能日曆、凍結排程與批次/資源假設。"
        )
        ttk.Label(parent, text=intro, justify="left", foreground="#444444").grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 10)
        )

        ttk.Label(parent, text="資料來源").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Combobox(
            parent,
            textvariable=self.data_source_mode_var,
            values=("本機 Excel", "Google Sheets（下一版串接）"),
            state="readonly",
            width=24,
        ).grid(row=1, column=1, sticky="w", padx=6, pady=5)
        ttk.Label(parent, text="Google Sheet ID").grid(row=1, column=2, sticky="w", padx=(18, 0), pady=5)
        ttk.Entry(parent, textvariable=self.google_sheet_id_var).grid(row=1, column=3, sticky="ew", padx=6, pady=5)

        ttk.Label(parent, text="排程邏輯").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Label(parent, textvariable=self.schedule_mode_var, foreground="#1f5f5b").grid(row=2, column=1, sticky="w", padx=6, pady=5)
        ttk.Label(parent, text="每日重排時間").grid(row=2, column=2, sticky="w", padx=(18, 0), pady=5)
        ttk.Entry(parent, textvariable=self.daily_rerun_time_var, width=10).grid(row=2, column=3, sticky="w", padx=6, pady=5)

        ttk.Label(parent, text="訂單基礎 workbook").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=self.workbook_var).grid(row=3, column=1, columnspan=2, sticky="ew", padx=6, pady=5)
        ttk.Button(parent, text="選擇企業 workbook", command=self._choose_workbook).grid(row=3, column=3, sticky="w", padx=6, pady=5)

        ttk.Label(parent, text="製造效能表（選填）").grid(row=4, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=self.manufacturing_efficiency_path_var).grid(row=4, column=1, sticky="ew", padx=6, pady=5)
        ttk.Button(parent, text="選擇 Excel", command=self._choose_manufacturing_efficiency).grid(row=4, column=2, sticky="w", padx=6, pady=5)
        ttk.Label(parent, text="空白則使用基礎 workbook 內建表").grid(row=4, column=3, sticky="w", pady=5)

        ttk.Label(parent, text="充填效能表（選填）").grid(row=5, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=self.filling_efficiency_path_var).grid(row=5, column=1, sticky="ew", padx=6, pady=5)
        ttk.Button(parent, text="選擇 Excel", command=self._choose_filling_efficiency).grid(row=5, column=2, sticky="w", padx=6, pady=5)
        ttk.Label(parent, text="空白則使用基礎 workbook 內建表").grid(row=5, column=3, sticky="w", pady=5)

        ttk.Label(parent, text="產能日曆 calendar_capacity").grid(row=6, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=self.calendar_capacity_var).grid(row=6, column=1, columnspan=2, sticky="ew", padx=6, pady=5)
        calendar_button_frame = ttk.Frame(parent)
        calendar_button_frame.grid(row=6, column=3, sticky="w", padx=6, pady=5)
        ttk.Button(calendar_button_frame, text="選擇 Excel", command=self._choose_calendar_capacity).pack(side="left")
        ttk.Checkbutton(calendar_button_frame, text="排除週六日", variable=self.exclude_weekends_var).pack(side="left", padx=8)

        ttk.Label(parent, text="凍結排程 CSV（選填）").grid(row=7, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=self.frozen_schedule_var).grid(row=7, column=1, columnspan=2, sticky="ew", padx=6, pady=5)
        frozen_button_frame = ttk.Frame(parent)
        frozen_button_frame.grid(row=7, column=3, sticky="w", padx=6, pady=5)
        ttk.Button(frozen_button_frame, text="選擇 CSV", command=self._choose_frozen_schedule).pack(side="left")
        ttk.Button(frozen_button_frame, text="匯出範本", command=self._export_frozen_schedule_template).pack(side="left", padx=6)

        manual_freeze = ttk.LabelFrame(parent, text="手動凍結排程（可直接新增多筆，不必寫 CSV）", padding=8)
        manual_freeze.grid(row=8, column=0, columnspan=4, sticky="nsew", pady=8)
        manual_freeze.columnconfigure(7, weight=1)
        ttk.Label(manual_freeze, text="訂單").grid(row=0, column=0, sticky="w")
        ttk.Entry(manual_freeze, textvariable=self.freeze_order_var, width=18).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(manual_freeze, text="工序").grid(row=0, column=2, sticky="w")
        ttk.Combobox(
            manual_freeze,
            textvariable=self.freeze_stage_var,
            values=("預攪", "攪料清洗", "排瓶", "充填"),
            state="readonly",
            width=10,
        ).grid(row=0, column=3, sticky="w", padx=4)
        ttk.Label(manual_freeze, text="日期").grid(row=0, column=4, sticky="w")
        ttk.Entry(manual_freeze, textvariable=self.freeze_date_var, width=12).grid(row=0, column=5, sticky="w", padx=4)
        ttk.Label(manual_freeze, text="時間/備註").grid(row=0, column=6, sticky="w")
        ttk.Entry(manual_freeze, textvariable=self.freeze_time_var).grid(row=0, column=7, sticky="ew", padx=4)
        ttk.Button(manual_freeze, text="新增", command=self._add_manual_freeze).grid(row=0, column=8, sticky="ew", padx=4)
        ttk.Button(manual_freeze, text="刪除選取", command=self._remove_manual_freeze).grid(row=0, column=9, sticky="ew")

        freeze_columns = ("order_id", "stage", "date", "time_note")
        self.frozen_tree = ttk.Treeview(manual_freeze, columns=freeze_columns, show="headings", height=5)
        for column, label, width in (
            ("order_id", "訂單", 150),
            ("stage", "工序", 90),
            ("date", "日期", 110),
            ("time_note", "時間/備註", 260),
        ):
            self.frozen_tree.heading(column, text=label)
            self.frozen_tree.column(column, width=width, anchor="w")
        self.frozen_tree.grid(row=1, column=0, columnspan=10, sticky="nsew", pady=(8, 0))

        settings_frame = ttk.LabelFrame(parent, text="批次與資源假設", padding=8)
        settings_frame.grid(row=9, column=0, columnspan=4, sticky="nsew", pady=10)
        settings_frame.columnconfigure(1, weight=1)
        settings_frame.columnconfigure(3, weight=1)

        rows = [
            ("換線分鐘", self.line_change_minutes_var, "同機台換產品或換型態時的額外時間，可用 60 測試一小時換線。"),
            ("休息開始", self.break_start_time_var, "每日不可排班休息開始時間，預設 12:00。"),
            ("休息結束", self.break_end_time_var, "每日不可排班休息結束時間，預設 13:00。"),
            ("拆單單位", self.split_unit_var, "製造與充填盡量以整批/固定倍數處理；目前先用 10 的倍數。"),
            ("最小充填量(L)", self.minimum_fill_liters_var, "業主提到 50L 桶量；目前先記錄，正式排程盡量避免切得太碎。"),
            ("小噴印機數量", self.small_printer_count_var, "目前業主回饋為 1 台。"),
            ("正常噴印機數量", self.normal_printer_count_var, "目前業主回饋為 1 台。"),
        ]
        for row_index, (label, variable, help_text) in enumerate(rows):
            row = row_index // 2
            column = (row_index % 2) * 2
            ttk.Label(settings_frame, text=label).grid(row=row, column=column, sticky="w", pady=6)
            ttk.Entry(settings_frame, textvariable=variable, width=12).grid(row=row, column=column + 1, sticky="w", padx=6, pady=6)
            ttk.Label(settings_frame, text=help_text, foreground="#666666").grid(
                row=row + 3, column=column, columnspan=2, sticky="w", pady=(0, 4)
            )

        ttk.Checkbutton(
            settings_frame,
            text="拆單後日期盡量連續（已接入 GA 軟性目標）",
            variable=self.keep_split_days_contiguous_var,
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=8)
        ttk.Checkbutton(
            settings_frame,
            text="輸出派工週視圖",
            variable=self.output_week_view_var,
        ).grid(row=6, column=2, columnspan=2, sticky="w", pady=8)

        action_frame = ttk.Frame(parent)
        action_frame.grid(row=10, column=0, columnspan=4, sticky="ew", pady=8)
        ttk.Button(action_frame, text="刷新設定摘要", command=self._refresh_developer_summary).pack(side="left")
        ttk.Button(action_frame, text="匯出開發設定 JSON", command=self._export_developer_options).pack(side="left", padx=8)

        summary_frame = ttk.LabelFrame(parent, text="目前設定摘要", padding=8)
        summary_frame.grid(row=11, column=0, columnspan=4, sticky="ew", pady=(4, 0))
        ttk.Label(summary_frame, textvariable=self.developer_summary_var, justify="left").grid(row=0, column=0, sticky="w")
        self._refresh_developer_summary()

    def _build_run_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(3, weight=1)

        button_frame = ttk.Frame(parent)
        button_frame.grid(row=0, column=0, sticky="ew")
        ttk.Button(button_frame, text="開始執行 GA", command=self._start_run, style="Primary.TButton").pack(side="left")
        ttk.Button(button_frame, text="停止執行", command=self._stop_run).pack(side="left", padx=8)
        ttk.Button(button_frame, text="開啟輸出資料夾", command=self._open_output_dir).pack(side="left")
        ttk.Button(button_frame, text="開啟交付資料夾", command=self._open_delivery_dir).pack(side="left", padx=8)
        ttk.Button(button_frame, text="開啟 Google 日曆匯入檔", command=self._open_calendar_file).pack(side="left")

        status_frame = ttk.LabelFrame(parent, text="狀態", padding=8)
        status_frame.grid(row=1, column=0, sticky="ew", pady=8)
        status_frame.columnconfigure(0, weight=1)
        ttk.Label(status_frame, textvariable=self.status_var).grid(row=0, column=0, sticky="w")
        ttk.Label(status_frame, textvariable=self.output_dir_status_var).grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.run_progress = ttk.Progressbar(status_frame, mode="indeterminate")
        self.run_progress.grid(row=2, column=0, sticky="ew", pady=(8, 0))

        quick_output_frame = ttk.LabelFrame(parent, text="交付輸出快速開啟", padding=8)
        quick_output_frame.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        for column in range(5):
            quick_output_frame.columnconfigure(column, weight=1)
        ttk.Button(quick_output_frame, text="01 排程檢查報告", command=self._open_check_report).grid(
            row=0, column=0, sticky="ew", padx=4
        )
        ttk.Button(quick_output_frame, text="02 當日確切時間", command=self._open_exact_time_schedule).grid(
            row=0, column=1, sticky="ew", padx=4
        )
        ttk.Button(quick_output_frame, text="03 Google 日曆", command=self._open_calendar_file).grid(
            row=0, column=2, sticky="ew", padx=4
        )
        ttk.Button(quick_output_frame, text="04 輔助檢查資料", command=self._open_supporting_outputs).grid(
            row=0, column=3, sticky="ew", padx=4
        )
        ttk.Button(quick_output_frame, text="README 先看我", command=self._open_delivery_readme).grid(
            row=0, column=4, sticky="ew", padx=4
        )

        log_frame = ttk.LabelFrame(parent, text="執行紀錄", padding=8)
        log_frame.grid(row=3, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, wrap="word", height=20)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _build_package_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)

        intro = (
            "這個頁籤是給專案維護者使用的打包工具。\n"
            "一般使用者只需要拿到建好的 installer 或 dist 資料夾，不需要自己執行這一步。"
        )
        ttk.Label(parent, text=intro, justify="left").grid(row=0, column=0, sticky="w")

        button_frame = ttk.Frame(parent)
        button_frame.grid(row=1, column=0, sticky="ew", pady=12)
        ttk.Button(button_frame, text="建立 Windows 安裝包/可攜版", command=self._start_package_build).pack(side="left")
        ttk.Button(button_frame, text="開啟 dist 資料夾", command=self._open_dist_dir).pack(side="left", padx=8)

        package_frame = ttk.LabelFrame(parent, text="打包紀錄", padding=8)
        package_frame.grid(row=2, column=0, sticky="nsew")
        package_frame.columnconfigure(0, weight=1)
        package_frame.rowconfigure(0, weight=1)

        self.package_log_text = tk.Text(package_frame, wrap="word", height=18)
        self.package_log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(package_frame, orient="vertical", command=self.package_log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.package_log_text.configure(yscrollcommand=scrollbar.set)

    def _choose_workbook(self) -> None:
        path = filedialog.askopenfilename(
            title="選擇訂單基礎 workbook",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=str(PROJECT_DIR),
        )
        if path:
            self.workbook_var.set(path)

    def _choose_manufacturing_efficiency(self) -> None:
        path = filedialog.askopenfilename(
            title="選擇製造效能表 Excel",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=str(PROJECT_DIR),
        )
        if path:
            self.manufacturing_efficiency_path_var.set(path)
            self._refresh_developer_summary()
            self._refresh_overview()

    def _choose_filling_efficiency(self) -> None:
        path = filedialog.askopenfilename(
            title="選擇充填效能表 Excel",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=str(PROJECT_DIR),
        )
        if path:
            self.filling_efficiency_path_var.set(path)
            self._refresh_developer_summary()
            self._refresh_overview()

    def _choose_seed_schedule(self) -> None:
        path = filedialog.askopenfilename(
            title="選擇 seed 排程 CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialdir=str(PROJECT_DIR),
        )
        if path:
            self.seed_schedule_var.set(path)

    def _choose_output_base(self) -> None:
        directory = filedialog.askdirectory(title="選擇輸出資料夾位置", initialdir=str(PROJECT_DIR))
        if directory:
            self.output_dir_var.set(str(Path(directory) / "ga_gui_output"))

    def _choose_calendar_capacity(self) -> None:
        current = self.calendar_capacity_var.get().strip()
        initial_dir = Path(current).expanduser().parent if current else PROJECT_DIR
        path = filedialog.askopenfilename(
            title="選擇 calendar_capacity.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=str(initial_dir if initial_dir.exists() else PROJECT_DIR),
        )
        if path:
            self.calendar_capacity_var.set(path)
            self._refresh_developer_summary()
            self._refresh_overview()

    def _choose_frozen_schedule(self) -> None:
        path = filedialog.askopenfilename(
            title="選擇凍結排程 CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialdir=str(PROJECT_DIR),
        )
        if path:
            self.frozen_schedule_var.set(path)
            self._refresh_developer_summary()

    def _add_manual_freeze(self) -> None:
        order_id = self.freeze_order_var.get().strip()
        stage = self.freeze_stage_var.get().strip()
        freeze_date = self.freeze_date_var.get().strip()
        time_note = self.freeze_time_var.get().strip()
        if not order_id or not stage or not freeze_date:
            messagebox.showerror("凍結排程缺資料", "請至少輸入訂單、工序與日期。")
            return
        try:
            from genetic_algorithm import _parse_date_text

            normalized_date = _parse_date_text(freeze_date).isoformat()
        except Exception as exc:
            messagebox.showerror("日期格式錯誤", f"日期請輸入 YYYY-MM-DD，例如 2023-05-15。\n{exc}")
            return
        self.frozen_tree.insert("", "end", values=(order_id, stage, normalized_date, time_note))
        self.freeze_order_var.set("")
        self.freeze_date_var.set("")
        self.freeze_time_var.set("")
        self._refresh_developer_summary()
        self._refresh_overview()

    def _remove_manual_freeze(self) -> None:
        for item_id in self.frozen_tree.selection():
            self.frozen_tree.delete(item_id)
        self._refresh_developer_summary()
        self._refresh_overview()

    def _manual_frozen_entries(self) -> list[dict[str, str]]:
        if not hasattr(self, "frozen_tree"):
            return []
        entries: list[dict[str, str]] = []
        for item_id in self.frozen_tree.get_children():
            order_id, stage, freeze_date, time_note = self.frozen_tree.item(item_id, "values")
            entries.append(
                {
                    "order_id": str(order_id),
                    "stage": str(stage),
                    "date": str(freeze_date),
                    "note": str(time_note),
                }
            )
        return entries

    def _export_frozen_schedule_template(self) -> None:
        path = filedialog.asksaveasfilename(
            title="匯出凍結排程 CSV 範本",
            defaultextension=".csv",
            initialfile="凍結排程範本.csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialdir=str(PROJECT_DIR),
        )
        if not path:
            return
        with Path(path).open("w", newline="", encoding="utf-8-sig") as file:
            writer = csv.writer(file)
            writer.writerow(["order_id", "stage", "date", "note"])
            writer.writerow(["P001", "充填", "2023-05-15", "範例：將 P001 充填固定在 2023-05-15"])
            writer.writerow(["P001", "排瓶", "2023-05-14", "範例：若要同時固定排瓶日，可再加一列"])
            writer.writerow(["P002", "攪料清洗", "2023-05-10", "範例：製造端鎖攪料日，預攪會自動排前一天"])
        messagebox.showinfo("已匯出", f"已匯出凍結排程範本：\n{path}")

    def _validate_time_text(self, value: str, label: str = "時間") -> str:
        value = value.strip()
        try:
            hour_text, minute_text = value.split(":", 1)
            hour = int(hour_text)
            minute = int(minute_text)
        except ValueError as exc:
            raise ValueError(f"{label}必須是 HH:MM，例如 16:00。") from exc
        if not (0 <= hour <= 23 and 0 <= minute <= 59):
            raise ValueError(f"{label}必須介於 00:00 和 23:59。")
        return f"{hour:02d}:{minute:02d}"

    def _time_text_to_minutes(self, value: str, label: str) -> int:
        normalized = self._validate_time_text(value, label)
        hour_text, minute_text = normalized.split(":", 1)
        return int(hour_text) * 60 + int(minute_text)

    def _read_non_negative_int(self, variable: tk.StringVar, label: str) -> int:
        try:
            value = int(variable.get())
        except ValueError as exc:
            raise ValueError(f"{label} 必須是整數。") from exc
        if value < 0:
            raise ValueError(f"{label} 不能小於 0。")
        return value

    def _read_int_at_least(self, variable: tk.StringVar, label: str, minimum: int) -> int:
        try:
            value = int(variable.get())
        except ValueError as exc:
            raise ValueError(f"{label} 必須是整數。") from exc
        if value < minimum:
            raise ValueError(f"{label} 必須大於等於 {minimum}。")
        return value

    def _read_rate(self, variable: tk.StringVar, label: str) -> float:
        try:
            value = float(variable.get())
        except ValueError as exc:
            raise ValueError(f"{label} 必須是 0 到 1 之間的小數。") from exc
        if not 0 <= value <= 1:
            raise ValueError(f"{label} 必須介於 0 和 1 之間。")
        return value

    def _build_developer_options(self, *, validate_paths: bool) -> dict[str, object]:
        frozen_schedule = self.frozen_schedule_var.get().strip()
        if frozen_schedule and validate_paths and not Path(frozen_schedule).expanduser().exists():
            raise ValueError(f"找不到凍結排程 CSV：\n{frozen_schedule}")
        calendar_capacity = self.calendar_capacity_var.get().strip()
        if calendar_capacity and validate_paths and not Path(calendar_capacity).expanduser().exists():
            raise ValueError(f"找不到 calendar_capacity.xlsx：\n{calendar_capacity}")
        manufacturing_efficiency = self.manufacturing_efficiency_path_var.get().strip()
        if manufacturing_efficiency and validate_paths and not Path(manufacturing_efficiency).expanduser().exists():
            raise ValueError(f"找不到製造效能表 Excel：\n{manufacturing_efficiency}")
        filling_efficiency = self.filling_efficiency_path_var.get().strip()
        if filling_efficiency and validate_paths and not Path(filling_efficiency).expanduser().exists():
            raise ValueError(f"找不到充填效能表 Excel：\n{filling_efficiency}")
        break_start_time = self._validate_time_text(self.break_start_time_var.get(), "休息開始時間")
        break_end_time = self._validate_time_text(self.break_end_time_var.get(), "休息結束時間")
        if self._time_text_to_minutes(break_end_time, "休息結束時間") <= self._time_text_to_minutes(break_start_time, "休息開始時間"):
            raise ValueError("休息結束時間必須晚於休息開始時間。")

        schedule_mode_label = "完整連動排程（製造先於充填，排瓶於充填前一工作日）"

        return {
            "data_source_mode": self.data_source_mode_var.get(),
            "google_sheet_id": self.google_sheet_id_var.get().strip(),
            "schedule_mode": "full_ga",
            "schedule_mode_label": schedule_mode_label,
            "calendar_capacity_path": calendar_capacity,
            "manufacturing_efficiency_path": manufacturing_efficiency,
            "filling_efficiency_path": filling_efficiency,
            "exclude_weekends": self.exclude_weekends_var.get(),
            "frozen_schedule_csv": frozen_schedule,
            "frozen_schedule_entries": self._manual_frozen_entries(),
            "daily_rerun_time": self._validate_time_text(self.daily_rerun_time_var.get(), "每日重排時間"),
            "line_change_minutes": self._read_non_negative_int(self.line_change_minutes_var, "換線分鐘"),
            "break_start_time": break_start_time,
            "break_end_time": break_end_time,
            "split_unit": self._read_non_negative_int(self.split_unit_var, "拆單單位"),
            "minimum_fill_liters": self._read_non_negative_int(self.minimum_fill_liters_var, "最小充填量"),
            "keep_split_days_contiguous": self.keep_split_days_contiguous_var.get(),
            "small_printer_count": self._read_non_negative_int(self.small_printer_count_var, "小噴印機數量"),
            "normal_printer_count": self._read_non_negative_int(self.normal_printer_count_var, "正常噴印機數量"),
            "output_week_view": self.output_week_view_var.get(),
        }

    def _refresh_developer_summary(self) -> None:
        try:
            options = self._build_developer_options(validate_paths=False)
        except ValueError as exc:
            self.developer_summary_var.set(f"設定錯誤：{exc}")
            return

        frozen_text = options["frozen_schedule_csv"] or "未指定"
        calendar_text = options["calendar_capacity_path"] or "未指定"
        sheet_text = options["google_sheet_id"] or "未指定"
        manufacturing_text = options["manufacturing_efficiency_path"] or "使用基礎 workbook"
        filling_text = options["filling_efficiency_path"] or "使用基礎 workbook"
        manual_count = len(options["frozen_schedule_entries"])
        self.developer_summary_var.set(
            "\n".join(
                [
                    f"資料來源：{options['data_source_mode']}；Google Sheet ID：{sheet_text}",
                    f"製造效能表：{manufacturing_text}",
                    f"充填效能表：{filling_text}",
                    f"產能日曆：{calendar_text}；排除週六日：{'是' if options['exclude_weekends'] else '否'}",
                    f"排程模式：{options['schedule_mode_label']}；每日重排：{options['daily_rerun_time']}",
                    f"凍結排程 CSV：{frozen_text}；手動凍結：{manual_count} 筆",
                    f"換線：{options['line_change_minutes']} 分鐘；休息不可排班：{options['break_start_time']}-{options['break_end_time']}",
                    f"拆單單位：{options['split_unit']}；最小充填量：{options['minimum_fill_liters']} L",
                    f"連續充填偏好：{'啟用' if options['keep_split_days_contiguous'] else '停用'}；週視圖：{'輸出' if options['output_week_view'] else '不輸出'}",
                    f"噴印機：小 {options['small_printer_count']} 台、正常 {options['normal_printer_count']} 台（目前僅記錄，尚未限制排程）",
                ]
            )
        )

    def _export_developer_options(self) -> None:
        try:
            options = self._build_developer_options(validate_paths=False)
        except ValueError as exc:
            messagebox.showerror("設定錯誤", str(exc))
            return
        path = filedialog.asksaveasfilename(
            title="匯出開發設定 JSON",
            defaultextension=".json",
            initialfile="developer_options.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir=str(PROJECT_DIR),
        )
        if not path:
            return
        Path(path).write_text(json.dumps(options, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        messagebox.showinfo("已匯出", f"已匯出開發設定：\n{path}")

    def _export_input_format_guide(self) -> None:
        path = filedialog.asksaveasfilename(
            title="匯出輸入格式說明",
            defaultextension=".csv",
            initialfile="輸入格式說明.csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialdir=str(PROJECT_DIR),
        )
        if not path:
            return

        with Path(path).open("w", newline="", encoding="utf-8-sig") as file:
            writer = csv.writer(file)
            writer.writerow(["資料類型", "建議工作表", "必要欄位概念", "用途"])
            writer.writerows(INPUT_FORMAT_GROUPS)
        messagebox.showinfo("匯出完成", f"已匯出輸入格式說明：\n{path}")

    def _use_quick_preset(self) -> None:
        self.population_size_var.set("3")
        self.generations_var.set("1")
        self.mutation_rate_var.set("0.15")
        self.crossover_rate_var.set("0.85")
        self.seed_var.set("42")

    def _use_demo_preset(self) -> None:
        self.population_size_var.set("20")
        self.generations_var.set("40")
        self.mutation_rate_var.set("0.15")
        self.crossover_rate_var.set("0.85")
        self.seed_var.set("42")
        self.schedule_mode_var.set("完整連動排程（製造先於充填，排瓶於充填前一工作日）")
        self.line_change_minutes_var.set("0")
        self.break_start_time_var.set("12:00")
        self.break_end_time_var.set("13:00")

    def _use_formal_preset(self) -> None:
        self.population_size_var.set("40")
        self.generations_var.set("120")
        self.mutation_rate_var.set("0.15")
        self.crossover_rate_var.set("0.85")
        self.seed_var.set("42")

    def _start_preview(self) -> None:
        workbook = Path(self.workbook_var.get()).expanduser()
        calendar_text = self.calendar_capacity_var.get().strip()
        calendar_path = Path(calendar_text).expanduser() if calendar_text else None
        if calendar_path is None or not calendar_path.exists():
            messagebox.showerror("產能日曆不存在", f"請先選擇 calendar_capacity.xlsx：\n{calendar_path or ''}")
            return
        if not workbook.exists():
            messagebox.showerror(
                "內建基礎資料不存在",
                f"找不到內建基礎資料 workbook_bom_refresh.xlsx：\n{workbook}\n\n"
                "請確認程式資料夾內有附帶這份基礎 workbook。",
            )
            return
        self.preview_var.set("正在檢查 calendar_capacity 與內建基礎資料...")
        self.status_var.set("正在檢查 calendar_capacity 與內建基礎資料...")
        exclude_warnings = self.exclude_warning_var.get()
        try:
            advanced_options = self._build_developer_options(validate_paths=True)
        except ValueError as exc:
            messagebox.showerror("設定錯誤", str(exc))
            return
        thread = threading.Thread(
            target=self._preview_worker,
            args=(workbook, exclude_warnings, advanced_options, calendar_path, self.exclude_weekends_var.get()),
            daemon=True,
        )
        thread.start()

    def _preview_worker(
        self,
        workbook: Path,
        exclude_warnings: bool,
        advanced_options: dict[str, object],
        calendar_path: Path | None,
        exclude_weekends: bool,
    ) -> None:
        try:
            from genetic_algorithm import WorkbookLoader, load_calendar_capacity, split_schedulable_orders

            loader = WorkbookLoader(
                workbook,
                manufacturing_efficiency_path=Path(str(advanced_options["manufacturing_efficiency_path"])).expanduser()
                if advanced_options.get("manufacturing_efficiency_path")
                else None,
                filling_efficiency_path=Path(str(advanced_options["filling_efficiency_path"])).expanduser()
                if advanced_options.get("filling_efficiency_path")
                else None,
            )
            orders = loader.load_orders()
            calendar_capacity = (
                load_calendar_capacity(calendar_path, exclude_weekends=exclude_weekends)
                if calendar_path is not None and calendar_path.exists()
                else None
            )
            schedulable_orders, anomalies = split_schedulable_orders(
                orders,
                exclude_warnings=exclude_warnings,
                independent_sections=False,
                calendar_capacity=calendar_capacity,
            )
            excluded = [item for item in anomalies if item.exclude_from_optimization]
            warnings = [item for item in anomalies if not item.exclude_from_optimization]
            available_sheets = set(loader.sheet_names)
            required_groups = {
                "訂單時間": ["時間準備表V1"],
                "製造效能": ["製造效能_BOM回推", "製造效能"],
                "充填效能": ["充填效能_BOM回推", "充填效能"],
            }
            format_status = []
            for group_name, candidates in required_groups.items():
                matched = [sheet for sheet in candidates if sheet in available_sheets]
                if matched:
                    format_status.append(f"{group_name}：OK（{matched[0]}）")
                else:
                    format_status.append(f"{group_name}：缺少建議工作表")
            text = (
                f"業主輸入檔：{calendar_path.name if calendar_path else '未指定'}\n"
                f"產能日曆筆數：{len(calendar_capacity.days) if calendar_capacity else 0}\n"
                f"內建基礎資料：{workbook.name}\n"
                f"製造效能來源：{Path(str(advanced_options['manufacturing_efficiency_path'])).name if advanced_options.get('manufacturing_efficiency_path') else '基礎 workbook'}\n"
                f"充填效能來源：{Path(str(advanced_options['filling_efficiency_path'])).name if advanced_options.get('filling_efficiency_path') else '基礎 workbook'}\n"
                "排程邏輯：製造先於充填；排瓶安排在充填前一個工作日\n"
                f"休息不可排班：{advanced_options['break_start_time']}-{advanced_options['break_end_time']}\n"
                f"原始訂單數量：{len(orders)}\n"
                f"可進入 GA 的訂單數量：{len(schedulable_orders)}\n"
                f"排除型異常訂單筆數：{len({item.order_id for item in excluded})}\n"
                f"提醒型異常訂單筆數：{len({item.order_id for item in warnings})}\n"
                f"無法建模訂單筆數：{len(loader.model_build_issues)}\n"
                f"內建基礎資料格式檢查：{'；'.join(format_status)}\n"
                f"機器設定：{', '.join(f'{name}={count}' for name, count in loader.machine_counts.items())}"
            )
            self.message_queue.put(("preview", text))
        except Exception as exc:  # pragma: no cover - GUI error path
            self.message_queue.put(("preview_error", str(exc)))

    def _start_run(self) -> None:
        if self.running_thread is not None and self.running_thread.is_alive():
            messagebox.showwarning("仍在執行", "目前已有 GA 正在執行，請先停止或等待完成。")
            return

        try:
            config = self._read_run_config()
        except ValueError as exc:
            messagebox.showerror("參數錯誤", str(exc))
            return

        self.log_text.delete("1.0", "end")
        self.status_var.set("執行中...")
        self.output_dir_status_var.set("尚未產生輸出資料夾。")
        self.last_output_dir = None
        self.last_validation_errors = None
        self.last_validation_warnings = None
        self.run_progress.start(12)
        self._refresh_overview()

        self.running_thread = threading.Thread(target=self._run_worker, args=(config,), daemon=True)
        self.running_thread.start()

    def _read_run_config(self) -> dict[str, object]:
        workbook = Path(self.workbook_var.get()).expanduser()
        output_dir = Path(self.output_dir_var.get()).expanduser()
        seed_schedule = self.seed_schedule_var.get().strip()

        if not workbook.exists():
            raise ValueError(
                f"找不到內建基礎資料 workbook_bom_refresh.xlsx：\n{workbook}\n\n"
                "業主日常只需要提供 calendar_capacity.xlsx；這份基礎資料應隨程式一起打包。"
            )

        population_size = self._read_int_at_least(self.population_size_var, "population_size", 2)
        generations = self._read_int_at_least(self.generations_var, "generations", 1)
        mutation_rate = self._read_rate(self.mutation_rate_var, "mutation_rate")
        crossover_rate = self._read_rate(self.crossover_rate_var, "crossover_rate")
        seed = self._read_int_at_least(self.seed_var, "seed", 0)
        if seed_schedule and not Path(seed_schedule).expanduser().exists():
            raise ValueError(f"找不到 seed schedule CSV：\n{seed_schedule}")
        advanced_options = self._build_developer_options(validate_paths=True)
        if not advanced_options["calendar_capacity_path"]:
            raise ValueError("請先選擇 calendar_capacity.xlsx。")

        return {
            "workbook": workbook,
            "output_dir": output_dir,
            "population_size": population_size,
            "generations": generations,
            "mutation_rate": mutation_rate,
            "crossover_rate": crossover_rate,
            "seed": seed,
            "seed_schedule": Path(seed_schedule).expanduser() if seed_schedule else None,
            "exclude_warnings": self.exclude_warning_var.get(),
            "advanced_options": advanced_options,
        }

    def _run_worker(self, config: dict[str, object]) -> None:
        try:
            from genetic_algorithm import print_run_summary, run_scheduler

            self.message_queue.put(
                (
                    "log",
                    (
                        "開始執行 GA\n"
                        f"base_workbook={config['workbook']}\n"
                        f"calendar_capacity={config['advanced_options']['calendar_capacity_path']}\n"
                        f"output_dir={config['output_dir']}\n"
                        f"population_size={config['population_size']}, generations={config['generations']}\n"
                        f"mutation_rate={config['mutation_rate']}, crossover_rate={config['crossover_rate']}, seed={config['seed']}\n\n"
                        "下一版開發設定會寫入 run_metadata.csv 與 15_下一版開發設定.csv。\n\n"
                    ),
                )
            )

            writer = QueueTextWriter(self.message_queue)
            with redirect_stdout(writer):
                result = run_scheduler(
                    workbook_path=Path(config["workbook"]),
                    output_dir=Path(config["output_dir"]),
                    population_size=int(config["population_size"]),
                    generations=int(config["generations"]),
                    mutation_rate=float(config["mutation_rate"]),
                    crossover_rate=float(config["crossover_rate"]),
                    seed=int(config["seed"]),
                    seed_schedule_csv=config["seed_schedule"],  # type: ignore[arg-type]
                    exclude_warning_anomalies=bool(config["exclude_warnings"]),
                    advanced_options=config["advanced_options"],  # type: ignore[arg-type]
                )
                print_run_summary(result)

            validation_errors = sum(1 for issue in result.validation_issues if issue.severity == "error")
            validation_warnings = sum(1 for issue in result.validation_issues if issue.severity == "warning")
            self.message_queue.put(("output_dir", result.output_dir))
            self.message_queue.put(("done", (0, validation_errors, validation_warnings)))
        except Exception as exc:  # pragma: no cover - GUI error path
            self.message_queue.put(("run_error", str(exc)))
        finally:
            self.running_thread = None

    def _stop_run(self) -> None:
        if self.running_thread is not None and self.running_thread.is_alive():
            messagebox.showinfo("無法強制停止", "目前版本以安全背景執行緒運算，請等待本次 GA 結束。")
            return
        self.status_var.set("目前沒有正在執行的 GA。")

    def _open_path(self, path: Path) -> None:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])

    def _open_output_dir(self) -> None:
        if self.last_output_dir is None or not self.last_output_dir.exists():
            messagebox.showinfo("尚無輸出", "目前還沒有可開啟的輸出資料夾。")
            return
        self._open_path(self.last_output_dir)

    def _delivery_dir(self) -> Path | None:
        if self.last_output_dir is None:
            return None
        return self.last_output_dir / "交付用輸出"

    def _open_delivery_dir(self) -> None:
        delivery_dir = self._delivery_dir()
        if delivery_dir is None or not delivery_dir.exists():
            messagebox.showinfo("尚無交付輸出", "請先執行 GA，完成後才會產生交付用輸出資料夾。")
            return
        self._open_path(delivery_dir)

    def _open_delivery_file(self, filename: str, label: str) -> None:
        delivery_dir = self._delivery_dir()
        if delivery_dir is None:
            messagebox.showinfo("尚無輸出", "請先執行 GA，完成後才會產生交付用輸出。")
            return
        path = delivery_dir / filename
        if not path.exists():
            messagebox.showinfo("找不到檔案", f"找不到 {label}：\n{path}")
            return
        self._open_path(path)

    def _open_primary_dashboard(self) -> None:
        self._open_delivery_file("04_輔助檢查資料.xlsx", "輔助檢查資料")

    def _open_validation_summary(self) -> None:
        self._open_delivery_file("04_輔助檢查資料.xlsx", "輔助檢查資料")

    def _open_daily_schedule(self) -> None:
        self._open_delivery_file("02_當日排程確切時間.xlsx", "當日排程確切時間")

    def _open_week_view(self) -> None:
        self._open_delivery_file("02_當日排程確切時間.xlsx", "當日排程確切時間")

    def _open_check_report(self) -> None:
        self._open_delivery_file("01_排程檢查報告.xlsx", "排程檢查報告")

    def _open_exact_time_schedule(self) -> None:
        self._open_delivery_file("02_當日排程確切時間.xlsx", "當日排程確切時間")

    def _open_supporting_outputs(self) -> None:
        self._open_delivery_file("04_輔助檢查資料.xlsx", "輔助檢查資料")

    def _open_delivery_readme(self) -> None:
        self._open_delivery_file("README_先看我.md", "交付輸出 README")

    def _open_calendar_file(self) -> None:
        if self.last_output_dir is None:
            messagebox.showinfo("尚無輸出", "請先執行 GA，完成後才會產生 Google 日曆匯入檔。")
            return
        calendar_path = self.last_output_dir / "交付用輸出" / "03_Google日曆匯入檔.ics"
        if not calendar_path.exists():
            calendar_path = self.last_output_dir / "google_calendar_schedule.ics"
        if not calendar_path.exists():
            messagebox.showinfo("找不到檔案", "找不到 Google 日曆匯入檔，請重新執行 GA。")
            return
        self._open_path(calendar_path)

    def _start_package_build(self) -> None:
        if not BUILD_SCRIPT.exists():
            messagebox.showerror("找不到打包腳本", f"找不到：\n{BUILD_SCRIPT}")
            return
        self.package_log_text.delete("1.0", "end")
        thread = threading.Thread(target=self._package_worker, daemon=True)
        thread.start()

    def _package_worker(self) -> None:
        command = [sys.executable, "-u", "-B", str(BUILD_SCRIPT)]
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        self.message_queue.put(("package_log", "執行打包指令：\n" + " ".join(command) + "\n\n"))
        try:
            process = subprocess.Popen(
                command,
                cwd=str(PROJECT_DIR),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=env,
            )
            assert process.stdout is not None
            for line in process.stdout:
                self.message_queue.put(("package_log", line))
            self.message_queue.put(("package_done", process.wait()))
        except Exception as exc:  # pragma: no cover - GUI error path
            self.message_queue.put(("package_log", f"打包失敗：{exc}\n"))
            self.message_queue.put(("package_done", 1))

    def _open_dist_dir(self) -> None:
        dist_dir = PROJECT_DIR / "dist"
        if not dist_dir.exists():
            messagebox.showinfo("尚無 dist", "目前還沒有 dist 資料夾，請先建立安裝包/可攜版。")
            return
        self._open_path(dist_dir)

    def _poll_messages(self) -> None:
        while True:
            try:
                message_type, payload = self.message_queue.get_nowait()
            except queue.Empty:
                break

            if message_type == "log":
                self._append_log(str(payload))
            elif message_type == "preview":
                self.preview_var.set(str(payload))
                self.status_var.set("輸入檢查完成")
                self._refresh_overview()
            elif message_type == "preview_error":
                self.preview_var.set("輸入檢查失敗。")
                self.status_var.set("輸入檢查失敗")
                self._refresh_overview()
                messagebox.showerror("輸入檢查失敗", str(payload))
            elif message_type == "output_dir":
                self.last_output_dir = Path(payload)  # type: ignore[arg-type]
                self.output_dir_status_var.set(f"輸出資料夾：{self.last_output_dir}")
                self._refresh_overview()
            elif message_type == "done":
                return_code, validation_errors, validation_warnings = payload  # type: ignore[misc]
                self.last_validation_errors = int(validation_errors)
                self.last_validation_warnings = int(validation_warnings)
                self.run_progress.stop()
                if int(return_code) != 0:
                    self.status_var.set(f"執行失敗，return code = {return_code}")
                elif self.last_validation_errors:
                    self.status_var.set(
                        f"執行完成，但有 {self.last_validation_errors} 項阻擋錯誤，不可直接派工"
                    )
                else:
                    self.status_var.set(f"執行完成，可交付檢查；提醒 {self.last_validation_warnings or 0} 項")
                self._refresh_overview()
            elif message_type == "run_error":
                self.run_progress.stop()
                self.status_var.set("執行失敗")
                self._refresh_overview()
                messagebox.showerror("執行失敗", str(payload))
            elif message_type == "package_log":
                self.package_log_text.insert("end", str(payload))
                self.package_log_text.see("end")
            elif message_type == "package_done":
                return_code = int(payload)
                if return_code == 0:
                    self.package_log_text.insert("end", "\n打包完成。\n")
                else:
                    self.package_log_text.insert("end", f"\n打包未完成，return code = {return_code}。\n")
                self.package_log_text.see("end")

        self.after(100, self._poll_messages)

    def _append_log(self, text: str) -> None:
        self.log_text.insert("end", text)
        self.log_text.see("end")


def main() -> None:
    app = SchedulerGui()
    app.mainloop()


if __name__ == "__main__":
    main()
