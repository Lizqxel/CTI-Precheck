import csv
import queue
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional, Set, Tuple

from core.cancellation import request_cancel_service
from core.csv_processing import read_csv, validate_rows
from core.judgement_runner import run_judgement
from core.settings_store import SETTINGS_PATH, load_browser_settings, save_browser_settings
from ui.update_manager import UpdateManager
from version import APP_NAME, VERSION


EventQueue = queue.Queue[Tuple[str, object]]


class DesktopApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"提供判定CSVツール（デスクトップ版） - {APP_NAME} {VERSION}")
        self.root.geometry("1160x760")

        self.rows_data: List[Dict[str, str]] = []
        self.event_queue: EventQueue = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.current_stop_event: threading.Event | None = None
        self.running = False
        self.judgement_started_at: float | None = None
        self.judgement_started_datetime: datetime | None = None
        self.elapsed_timer_job: str | None = None
        self.start_time_label = tk.StringVar(value="開始時刻: -")
        self.elapsed_label = tk.StringVar(value="実行時間: -")

        self.total_label = tk.StringVar(value="総行数: 0")
        self.file_label = tk.StringVar(value="未選択")
        self.result_label = tk.StringVar(value="CSVファイルを選択してください。")
        self.progress_label = tk.StringVar(value="進捗: -")
        self.monitor_browser_var = tk.BooleanVar(value=False)
        self.show_popup_var = tk.BooleanVar(value=True)
        self.parallel_count_var = tk.IntVar(value=1)
        self.parallel_count_values = (1, 2, 3, 4, 5, 6, 7, 8)
        self.run_scope_var = tk.StringVar(value="全行")
        self.target_line_var = tk.StringVar(value="対象行: 未選択")
        self.execution_target_line: Optional[int] = None

        self.worker_log_texts: List[tk.Text] = []
        self.worker_logs_container: ttk.Frame | None = None
        self.main_canvas: tk.Canvas | None = None
        self.main_scrollbar: ttk.Scrollbar | None = None
        self.main_content: ttk.Frame | None = None
        self.main_canvas_window_id: int | None = None
        self.update_manager = UpdateManager(root=self.root, log_callback=self._append_log)

        self._load_settings_to_ui()
        self._build_ui()
        self._try_restore_autosave_on_startup()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close_requested)
        self.root.after(150, self._drain_event_queue)
        self.root.after(1000, self.check_for_updates_on_startup)

    def _build_ui(self) -> None:
        self._build_menu()

        content_host = ttk.Frame(self.root)
        content_host.pack(fill=tk.BOTH, expand=True)

        self.main_canvas = tk.Canvas(content_host, highlightthickness=0, borderwidth=0)
        self.main_scrollbar = ttk.Scrollbar(content_host, orient=tk.VERTICAL, command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)

        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.main_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.main_content = ttk.Frame(self.main_canvas)
        self.main_canvas_window_id = self.main_canvas.create_window((0, 0), window=self.main_content, anchor="nw")
        self.main_content.bind("<Configure>", self._on_main_content_configure)
        self.main_canvas.bind("<Configure>", self._on_main_canvas_configure)
        self.root.bind_all("<MouseWheel>", self._on_main_mousewheel)

        if self.main_content is None:
            return

        top_frame = ttk.Frame(self.main_content, padding=12)
        top_frame.pack(fill=tk.X)

        self.select_button = ttk.Button(top_frame, text="CSVファイルを選択", command=self.load_csv)
        self.select_button.pack(side=tk.LEFT)

        self.start_button = ttk.Button(top_frame, text="提供判定開始", command=self.start_judgement)
        self.start_button.pack(side=tk.LEFT, padx=(8, 0))

        self.scope_combo = ttk.Combobox(
            top_frame,
            textvariable=self.run_scope_var,
            values=["全行", "選択行のみ", "選択行以降"],
            state="readonly",
            width=14,
        )
        self.scope_combo.pack(side=tk.LEFT, padx=(12, 0))

        self.set_target_button = ttk.Button(top_frame, text="選択行を対象にセット", command=self.set_target_from_selection)
        self.set_target_button.pack(side=tk.LEFT, padx=(8, 0))

        ttk.Label(top_frame, textvariable=self.target_line_var).pack(side=tk.LEFT, padx=(8, 0))

        self.stop_button = ttk.Button(top_frame, text="停止", command=self.stop_judgement, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=(8, 0))

        self.save_button = ttk.Button(top_frame, text="結果CSV保存", command=self.save_result_csv)
        self.save_button.pack(side=tk.LEFT, padx=(8, 0))

        ttk.Label(top_frame, textvariable=self.file_label).pack(side=tk.LEFT, padx=(12, 0))

        setting_frame = ttk.LabelFrame(self.main_content, text="設定", padding=10)
        setting_frame.pack(fill=tk.X, padx=12, pady=(0, 8))

        ttk.Checkbutton(setting_frame, text="ブラウザ表示で監視する", variable=self.monitor_browser_var).pack(side=tk.LEFT)
        ttk.Checkbutton(setting_frame, text="判定結果ポップアップを有効化", variable=self.show_popup_var).pack(
            side=tk.LEFT, padx=(12, 0)
        )
        ttk.Label(setting_frame, text="並列数").pack(side=tk.LEFT, padx=(12, 4))
        self.parallel_count_combo = ttk.Combobox(
            setting_frame,
            values=self.parallel_count_values,
            width=3,
            state="readonly",
            textvariable=self.parallel_count_var,
        )
        self.parallel_count_combo.pack(side=tk.LEFT)
        self.parallel_count_combo.bind("<<ComboboxSelected>>", self._on_parallel_count_changed)
        ttk.Button(setting_frame, text="設定保存", command=self.save_settings).pack(side=tk.LEFT, padx=(16, 0))
        ttk.Button(setting_frame, text="更新チェック", command=self.check_for_updates_manual).pack(side=tk.LEFT, padx=(8, 0))

        info_frame = ttk.Frame(self.main_content, padding=(12, 0, 12, 8))
        info_frame.pack(fill=tk.X)
        ttk.Label(info_frame, textvariable=self.total_label).pack(side=tk.LEFT)
        ttk.Label(info_frame, textvariable=self.result_label).pack(side=tk.LEFT, padx=(16, 0))
        ttk.Label(info_frame, textvariable=self.progress_label).pack(side=tk.LEFT, padx=(16, 0))
        ttk.Label(info_frame, textvariable=self.start_time_label).pack(side=tk.LEFT, padx=(16, 0))
        ttk.Label(info_frame, textvariable=self.elapsed_label).pack(side=tk.LEFT, padx=(16, 0))

        table_frame = ttk.Frame(self.main_content)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 8))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        columns = ("行", "郵便番号", "住所", "状態", "判定結果", "備考")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=10)
        for col in columns:
            self.tree.heading(col, text=col)

        self._tree_column_layout = {
            "行": {"ratio": 0.06, "min": 50, "max": 90, "anchor": tk.CENTER},
            "郵便番号": {"ratio": 0.11, "min": 90, "max": 150, "anchor": tk.CENTER},
            "住所": {"ratio": 0.33, "min": 180, "max": 640, "anchor": tk.W},
            "状態": {"ratio": 0.10, "min": 90, "max": 180, "anchor": tk.CENTER},
            "判定結果": {"ratio": 0.11, "min": 100, "max": 180, "anchor": tk.CENTER},
            "備考": {"ratio": 0.29, "min": 180, "max": 720, "anchor": tk.W},
        }
        self._configure_tree_columns(1160)
        self.tree.bind("<Configure>", self._on_tree_configure)

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=h_scrollbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_selection)

        note_frame = ttk.LabelFrame(self.main_content, text="備考詳細", padding=8)
        note_frame.pack(fill=tk.BOTH, expand=False, padx=12, pady=(0, 8))

        self.note_text = tk.Text(note_frame, height=4, wrap=tk.WORD)
        note_scroll = ttk.Scrollbar(note_frame, orient=tk.VERTICAL, command=self.note_text.yview)
        self.note_text.configure(yscrollcommand=note_scroll.set)
        self.note_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        note_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.note_text.configure(state=tk.DISABLED)

        global_log_frame = ttk.LabelFrame(self.main_content, text="全体ログ", padding=8)
        global_log_frame.pack(fill=tk.BOTH, expand=False, padx=12, pady=(0, 8))

        self.log_text = tk.Text(global_log_frame, height=6, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state=tk.DISABLED)

        worker_log_frame = ttk.LabelFrame(self.main_content, text="ワーカー別ログ（提供判定実行中）", padding=8)
        worker_log_frame.pack(fill=tk.BOTH, expand=False, padx=12, pady=(0, 12))
        self.worker_logs_container = ttk.Frame(worker_log_frame)
        self.worker_logs_container.pack(fill=tk.BOTH, expand=True)

        self._rebuild_worker_log_panels()

    def _build_menu(self) -> None:
        menu_bar = tk.Menu(self.root)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="更新チェック", command=self.check_for_updates_manual)
        help_menu.add_separator()
        help_menu.add_command(label=f"バージョン: {VERSION}", state=tk.DISABLED)

        menu_bar.add_cascade(label="ヘルプ", menu=help_menu)
        self.root.configure(menu=menu_bar)

    def check_for_updates_on_startup(self) -> None:
        self.update_manager.check_for_updates(interactive=False, auto=True)

    def check_for_updates_manual(self) -> None:
        self.update_manager.check_for_updates(interactive=True, auto=False)

    def _load_settings_to_ui(self) -> None:
        browser_settings = load_browser_settings(SETTINGS_PATH)
        self.monitor_browser_var.set(not browser_settings.get("headless", True))
        self.show_popup_var.set(bool(browser_settings.get("show_popup", True)))

    def _on_tree_configure(self, event: tk.Event) -> None:
        width = int(getattr(event, "width", 0) or 0)
        if width > 0:
            self._configure_tree_columns(width)

    def _configure_tree_columns(self, total_width: int) -> None:
        visible_width = max(total_width - 8, 360)
        logical_width = max(visible_width + 220, 1100)
        preferred: Dict[str, int] = {}

        for col, conf in self._tree_column_layout.items():
            width = int(logical_width * float(conf["ratio"]))
            width = max(int(conf["min"]), min(int(conf["max"]), width))
            preferred[col] = width

        adjusted_total = sum(preferred.values())
        if adjusted_total < logical_width:
            preferred["備考"] += logical_width - adjusted_total

        preferred["備考"] += 120

        for col, conf in self._tree_column_layout.items():
            self.tree.column(
                col,
                width=preferred[col],
                minwidth=40,
                anchor=conf["anchor"],
                stretch=False,
            )

    def _build_browser_settings_from_ui(self) -> Dict[str, object]:
        return {
            "headless": not self.monitor_browser_var.get(),
            "show_popup": self.show_popup_var.get(),
            "auto_close": True,
            "page_load_timeout": 60,
            "script_timeout": 60,
        }

    def save_settings(self) -> None:
        browser_settings = self._build_browser_settings_from_ui()
        save_browser_settings(browser_settings, SETTINGS_PATH)
        self._append_log("設定を保存しました")
        self.result_label.set("設定を保存しました")

    def load_csv(self) -> None:
        if self.running:
            messagebox.showwarning("実行中", "提供判定の実行中はCSVを変更できません。")
            return

        selected = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
        )
        if not selected:
            return

        file_path = Path(selected)
        self.file_label.set(str(file_path))

        try:
            rows, removed_blank_rows = read_csv(file_path)
        except Exception as exc:
            self.result_label.set("読み込み失敗")
            messagebox.showerror("エラー", f"CSVの読み込みに失敗しました\n{exc}")
            return

        if not rows:
            self._clear_tree()
            self.rows_data = []
            self.total_label.set("総行数: 0")
            self.result_label.set("CSVにデータがありません。")
            messagebox.showwarning("警告", "CSVにデータがありません。")
            return

        parsed_rows, invalid_line_numbers = validate_rows(rows)
        self.rows_data = parsed_rows
        self._render_rows(self.rows_data)
        self.execution_target_line = None
        self.target_line_var.set("対象行: 未選択")
        self.run_scope_var.set("全行")

        self.total_label.set(f"総行数: {len(self.rows_data)}")
        if removed_blank_rows > 0:
            self.result_label.set(f"CSV読み込み完了（空行 {removed_blank_rows} 行を除外）")
        else:
            self.result_label.set("CSV読み込み完了")
        self.progress_label.set("進捗: -")
        if removed_blank_rows > 0:
            self._append_log(f"CSVを読み込みました: {file_path.name}（空行 {removed_blank_rows} 行を除外）")
        else:
            self._append_log(f"CSVを読み込みました: {file_path.name}")
        if invalid_line_numbers:
            messagebox.showwarning(
                "入力不備のある行",
                f"次の行に入力不備があります: {', '.join(map(str, invalid_line_numbers))}",
            )

    def _try_restore_autosave_on_startup(self) -> None:
        autosave_path = self._get_autosave_path()
        if not autosave_path.exists():
            return

        should_restore = messagebox.askyesno(
            "前回進捗の復元",
            f"前回の自動保存データが見つかりました。\n復元しますか？\n\n{autosave_path.name}",
        )
        if not should_restore:
            self._append_log("前回進捗の自動復元をキャンセルしました")
            return

        try:
            rows, removed_blank_rows = read_csv(autosave_path)
        except Exception as exc:
            self._append_log(f"自動保存CSVの読み込みに失敗しました: {exc}")
            return

        if not rows:
            return

        parsed_rows, invalid_line_numbers = validate_rows(rows)
        if not parsed_rows:
            return

        for index, parsed in enumerate(parsed_rows):
            source = rows[index] if index < len(rows) else []
            result_value = source[2].strip() if len(source) >= 3 and source[2] is not None else ""
            note_value = source[3].strip() if len(source) >= 4 and source[3] is not None else ""
            if result_value:
                parsed["判定結果"] = result_value
            if note_value:
                parsed["備考"] = note_value

        self.rows_data = parsed_rows
        self._render_rows(self.rows_data)
        self.execution_target_line = None
        self.target_line_var.set("対象行: 未選択")
        self.run_scope_var.set("全行")

        self.file_label.set(f"自動復元: {autosave_path.name}")
        self.total_label.set(f"総行数: {len(self.rows_data)}")
        if removed_blank_rows > 0:
            self.result_label.set(f"前回の進捗を自動読み込みしました（空行 {removed_blank_rows} 行を除外）")
        else:
            self.result_label.set("前回の進捗を自動読み込みしました")
        self.progress_label.set("進捗: -")
        if removed_blank_rows > 0:
            self._append_log(f"前回の進捗を自動読み込みしました: {autosave_path}（空行 {removed_blank_rows} 行を除外）")
        else:
            self._append_log(f"前回の進捗を自動読み込みしました: {autosave_path}")

        if invalid_line_numbers:
            self._append_log(f"自動復元データに入力不備の行があります: {', '.join(map(str, invalid_line_numbers))}")

    def _resolve_target_lines(self) -> Optional[Set[int]]:
        scope = self.run_scope_var.get().strip()
        if scope == "全行":
            return None

        if self.execution_target_line is None:
            return set()

        if scope == "選択行のみ":
            return {self.execution_target_line}

        if scope == "選択行以降":
            return {
                int(row["行"])
                for row in self.rows_data
                if int(row["行"]) >= self.execution_target_line
            }

        return None

    def _set_execution_target_line(self, line_number: int) -> None:
        self.execution_target_line = line_number
        self.target_line_var.set(f"対象行: {line_number}")

    def _find_first_unfinished_line(self) -> Optional[int]:
        for row in self.rows_data:
            line_number = int(row.get("行", "0") or "0")
            result = str(row.get("判定結果", "") or "")
            if result in ("未実行", "停止"):
                return line_number
        return None

    def set_target_from_selection(self) -> None:
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("未選択", "テーブルから対象行を選択してください。")
            return

        try:
            line_number = int(selected[0])
        except Exception:
            messagebox.showwarning("選択エラー", "対象行の取得に失敗しました。")
            return

        self._set_execution_target_line(line_number)
        self._append_log(f"対象行を {line_number} に設定しました")

    def _on_tree_double_click(self, event: tk.Event) -> None:
        if self.running:
            return

        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return

        try:
            line_number = int(row_id)
        except Exception:
            return

        self._set_execution_target_line(line_number)
        choice = messagebox.askyesnocancel(
            "実行範囲を選択",
            f"{line_number}行目を選択しました。\n"
            "[はい] この行だけ実行\n"
            "[いいえ] この行以降を実行\n"
            "[キャンセル] 対象設定のみ",
        )

        if choice is True:
            self.run_scope_var.set("選択行のみ")
            self.start_judgement()
        elif choice is False:
            self.run_scope_var.set("選択行以降")
            self.start_judgement()

    def start_judgement(self) -> None:
        if self.running:
            return
        if not self.rows_data:
            messagebox.showwarning("未読み込み", "先にCSVファイルを選択してください。")
            return

        current_scope = self.run_scope_var.get().strip()
        if current_scope == "全行":
            resume_line = self._find_first_unfinished_line()
            if resume_line is not None and resume_line > 1:
                should_resume = messagebox.askyesno(
                    "実行範囲の確認",
                    f"{resume_line}行目以降に未実行データがあります。\n途中から再開しますか？\n\n"
                    "[はい] 未実行の先頭行以降を実行\n"
                    "[いいえ] 1行目から全行を実行",
                )
                if should_resume:
                    self._set_execution_target_line(resume_line)
                    self.run_scope_var.set("選択行以降")
                    self._append_log(f"未実行先頭の {resume_line} 行目から再開します")

        if self.execution_target_line is not None and self.run_scope_var.get().strip() == "全行":
            self.run_scope_var.set("選択行以降")
            self._append_log(f"対象行 {self.execution_target_line} が設定済みのため、実行範囲を『選択行以降』に変更しました")

        target_lines = self._resolve_target_lines()
        if target_lines is not None and len(target_lines) == 0:
            messagebox.showwarning("対象未設定", "実行対象の行を選択してください。")
            return

        total_targets = len(self.rows_data) if target_lines is None else len(target_lines)

        self.save_settings()
        run_queue: EventQueue = queue.Queue()
        run_stop_event = threading.Event()
        self.event_queue = run_queue
        self.current_stop_event = run_stop_event
        self.running = True
        self._set_running_ui_state(True)
        self.result_label.set("提供判定を実行中...")
        self.progress_label.set(f"進捗: 0/{total_targets}")
        self.judgement_started_datetime = datetime.now()
        self.start_time_label.set(f"開始時刻: {self._format_datetime(self.judgement_started_datetime)}")
        self.elapsed_label.set("実行時間: 00:00")
        self._append_log("提供判定を開始しました")
        self._rebuild_worker_log_panels(clear_existing=True)
        self.judgement_started_at = time.perf_counter()
        self._start_elapsed_timer()

        self.worker_thread = threading.Thread(
            target=self._run_judgement,
            args=(target_lines, run_queue, run_stop_event),
            daemon=True,
        )
        self.worker_thread.start()

    def stop_judgement(self) -> None:
        if not self.running:
            return

        if self.current_stop_event is not None:
            self.current_stop_event.set()

        self._append_log("停止要求を受け付けました（即時終了）")
        self.running = False
        self._set_running_ui_state(False)
        self.result_label.set("提供判定を停止しました")
        self.progress_label.set("進捗: 停止")
        request_cancel_service()
        self.event_queue = queue.Queue()
        self.current_stop_event = None
        self._stop_elapsed_timer()
        self.elapsed_label.set(f"実行時間: {self._get_elapsed_time_text()}")

    def _run_judgement(
        self,
        target_lines: Optional[Set[int]] = None,
        run_queue: Optional[EventQueue] = None,
        run_stop_event: Optional[threading.Event] = None,
    ) -> None:
        active_queue = run_queue if run_queue is not None else self.event_queue
        active_stop_event = run_stop_event if run_stop_event is not None else threading.Event()
        run_judgement(
            rows_data=self.rows_data,
            event_queue=active_queue,
            stop_requested=active_stop_event.is_set,
            parallel_count=self._get_parallel_count(),
            target_lines=target_lines,
        )

    def _drain_event_queue(self) -> None:
        while True:
            try:
                event, payload = self.event_queue.get_nowait()
            except queue.Empty:
                break

            if event == "row":
                self._update_row(payload)
            elif event == "log":
                self._append_log(str(payload))
            elif event == "worker_log":
                self._append_worker_log(payload)
            elif event == "progress":
                current, total = payload
                self.progress_label.set(f"進捗: {current}/{total}")
            elif event == "done":
                self._on_worker_done(payload)

        self.root.after(150, self._drain_event_queue)

    def _on_worker_done(self, payload: Dict[str, object]) -> None:
        self.running = False
        self._set_running_ui_state(False)

        failed_rows = payload.get("failed_rows", [])
        cancelled = bool(payload.get("cancelled", False))
        start_datetime_text = self._format_datetime(self.judgement_started_datetime)
        end_datetime = datetime.now()
        end_datetime_text = self._format_datetime(end_datetime)
        elapsed_text = self._get_elapsed_time_text()
        self._stop_elapsed_timer()
        self.elapsed_label.set(f"実行時間: {elapsed_text}")
        self.judgement_started_at = None
        self.judgement_started_datetime = None

        if cancelled:
            self.result_label.set("提供判定を停止しました")
            self._append_log("提供判定を停止しました")
            messagebox.showinfo(
                "停止",
                "提供判定を停止しました。\n"
                f"開始時刻: {start_datetime_text}\n"
                f"終了時刻: {end_datetime_text}\n"
                f"実行時間: {elapsed_text}",
            )
            return

        if failed_rows:
            self.result_label.set("提供判定完了（失敗あり）")
            lines = ", ".join(map(str, failed_rows))
            self._append_log(f"提供判定完了: 失敗行 {lines}")
            messagebox.showwarning(
                "失敗行",
                f"以下の行が失敗しました: {lines}\n"
                f"開始時刻: {start_datetime_text}\n"
                f"終了時刻: {end_datetime_text}\n"
                f"実行時間: {elapsed_text}",
            )
        else:
            self.result_label.set("提供判定完了")
            self._append_log("提供判定が完了しました")
            messagebox.showinfo(
                "完了",
                "提供判定が完了しました。\n"
                f"開始時刻: {start_datetime_text}\n"
                f"終了時刻: {end_datetime_text}\n"
                f"実行時間: {elapsed_text}",
            )

    def _format_datetime(self, value: datetime | None) -> str:
        if value is None:
            return "-"
        return value.strftime("%Y-%m-%d %H:%M:%S")

    def _get_elapsed_time_text(self) -> str:
        if self.judgement_started_at is None:
            return "不明"

        elapsed_seconds = max(0, int(time.perf_counter() - self.judgement_started_at))
        minutes, seconds = divmod(elapsed_seconds, 60)
        hours, minutes = divmod(minutes, 60)

        if hours > 0:
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        return f"{minutes:02d}:{seconds:02d}"

    def _start_elapsed_timer(self) -> None:
        self._stop_elapsed_timer()
        self._update_elapsed_timer()

    def _stop_elapsed_timer(self) -> None:
        if self.elapsed_timer_job is not None:
            try:
                self.root.after_cancel(self.elapsed_timer_job)
            except Exception:
                pass
            self.elapsed_timer_job = None

    def _update_elapsed_timer(self) -> None:
        if not self.running or self.judgement_started_at is None:
            self.elapsed_timer_job = None
            return

        self.elapsed_label.set(f"実行時間: {self._get_elapsed_time_text()}")
        self.elapsed_timer_job = self.root.after(500, self._update_elapsed_timer)

    def _set_running_ui_state(self, is_running: bool) -> None:
        self.select_button.configure(state=tk.DISABLED if is_running else tk.NORMAL)
        self.start_button.configure(state=tk.DISABLED if is_running else tk.NORMAL)
        self.stop_button.configure(state=tk.NORMAL if is_running else tk.DISABLED)
        self.parallel_count_combo.configure(state=tk.DISABLED if is_running else "readonly")
        self.scope_combo.configure(state=tk.DISABLED if is_running else "readonly")
        self.set_target_button.configure(state=tk.DISABLED if is_running else tk.NORMAL)

    def _on_parallel_count_changed(self, _event: object | None = None) -> None:
        if self.running:
            return
        self._rebuild_worker_log_panels()

    def _get_parallel_count(self) -> int:
        value = self.parallel_count_var.get()
        if value not in self.parallel_count_values:
            return 2
        return int(value)

    def _on_main_content_configure(self, _event: tk.Event) -> None:
        if self.main_canvas is None:
            return
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

    def _on_main_canvas_configure(self, event: tk.Event) -> None:
        if self.main_canvas is None or self.main_canvas_window_id is None:
            return
        self.main_canvas.itemconfigure(self.main_canvas_window_id, width=event.width)

    def _is_widget_or_descendant(self, widget: tk.Misc | None, ancestor: tk.Misc | None) -> bool:
        if widget is None or ancestor is None:
            return False

        current: tk.Misc | None = widget
        while current is not None:
            if current == ancestor:
                return True
            current = getattr(current, "master", None)

        return False

    def _is_inner_scrollable_area(self, widget: tk.Misc | None) -> bool:
        if widget is None:
            return False

        if self._is_widget_or_descendant(widget, self.tree):
            return True
        if self._is_widget_or_descendant(widget, self.note_text):
            return True
        if self._is_widget_or_descendant(widget, self.log_text):
            return True

        for worker_text in self.worker_log_texts:
            if self._is_widget_or_descendant(widget, worker_text):
                return True

        return False

    def _on_main_mousewheel(self, event: tk.Event) -> None:
        if self.main_canvas is None:
            return

        widget = getattr(event, "widget", None)
        if isinstance(widget, tk.Misc) and self._is_inner_scrollable_area(widget):
            return

        delta = int(getattr(event, "delta", 0) or 0)
        if delta == 0:
            return

        first, last = self.main_canvas.yview()
        if delta > 0 and first <= 0.0:
            return
        if delta < 0 and last >= 1.0:
            return

        self.main_canvas.yview_scroll(int(-delta / 120), "units")

    def _rebuild_worker_log_panels(self, clear_existing: bool = False) -> None:
        if self.worker_logs_container is None:
            return

        panel_count = self._get_parallel_count()
        for child in self.worker_logs_container.winfo_children():
            child.destroy()

        self.worker_log_texts = []

        columns = 2 if panel_count > 1 else 1
        rows = (panel_count + columns - 1) // columns

        for row_index in range(rows):
            self.worker_logs_container.rowconfigure(row_index, weight=1)
        for col_index in range(columns):
            self.worker_logs_container.columnconfigure(col_index, weight=1)

        for worker_index in range(panel_count):
            row_index = worker_index // columns
            col_index = worker_index % columns
            panel = ttk.LabelFrame(self.worker_logs_container, text=f"ワーカー {worker_index + 1}", padding=6)
            panel.grid(row=row_index, column=col_index, sticky="nsew", padx=4, pady=4)

            panel.grid_rowconfigure(0, weight=1)
            panel.grid_columnconfigure(0, weight=1)

            text = tk.Text(panel, height=9, wrap=tk.WORD)
            worker_scrollbar = ttk.Scrollbar(panel, orient=tk.VERTICAL, command=text.yview)
            text.configure(yscrollcommand=worker_scrollbar.set)

            text.grid(row=0, column=0, sticky="nsew")
            worker_scrollbar.grid(row=0, column=1, sticky="ns")
            text.configure(state=tk.DISABLED)

            if clear_existing:
                text.configure(state=tk.NORMAL)
                text.delete("1.0", tk.END)
                text.configure(state=tk.DISABLED)

            self.worker_log_texts.append(text)

        if self.main_canvas is not None:
            self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _append_worker_log(self, payload: object) -> None:
        if not isinstance(payload, dict):
            return

        worker = payload.get("worker")
        message = payload.get("message")
        if not isinstance(worker, int) or not isinstance(message, str):
            return
        if worker < 0 or worker >= len(self.worker_log_texts):
            return

        target = self.worker_log_texts[worker]
        target.configure(state=tk.NORMAL)
        target.insert(tk.END, f"{message}\n")
        target.see(tk.END)
        target.configure(state=tk.DISABLED)

    def _clear_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _render_rows(self, rows: List[Dict[str, str]]) -> None:
        self._clear_tree()
        for row in rows:
            self.tree.insert(
                "",
                tk.END,
                iid=row["行"],
                values=(row["行"], row["郵便番号"], row["住所"], row["状態"], row["判定結果"], row.get("備考", "")),
            )

        self._refresh_note_detail()

    def _update_row(self, row: Dict[str, str]) -> None:
        row_id = row["行"]
        if self.tree.exists(row_id):
            self.tree.item(
                row_id,
                values=(row["行"], row["郵便番号"], row["住所"], row["状態"], row["判定結果"], row.get("備考", "")),
            )

        self._refresh_note_detail()

    def _on_tree_selection(self, _event: tk.Event) -> None:
        self._refresh_note_detail()

    def _refresh_note_detail(self) -> None:
        selected = self.tree.selection()
        note = ""
        if selected:
            selected_id = selected[0]
            for row in self.rows_data:
                if row.get("行") == selected_id:
                    note = row.get("備考", "")
                    break

        self.note_text.configure(state=tk.NORMAL)
        self.note_text.delete("1.0", tk.END)
        if note:
            self.note_text.insert(tk.END, note)
        self.note_text.configure(state=tk.DISABLED)

    def _on_close_requested(self) -> None:
        self._auto_save_result_csv()
        self.root.destroy()

    def _get_runtime_base_dir(self) -> Path:
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parents[1]

    def _get_autosave_path(self) -> Path:
        return self._get_runtime_base_dir() / "result_autosave.csv"

    def _write_result_csv(self, save_path: Path) -> None:
        with save_path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            for row in self.rows_data:
                result_value = row.get("判定結果", "未実行")
                note_value = row.get("備考", "")
                writer.writerow([row["郵便番号"], row["住所"], result_value, note_value])

    def _auto_save_result_csv(self) -> None:
        if not self.rows_data:
            return

        save_path = self._get_autosave_path()
        try:
            self._write_result_csv(save_path)
            self._append_log(f"終了時に結果CSVを自動保存しました: {save_path}")
        except Exception as exc:
            self._append_log(f"終了時CSV自動保存に失敗しました: {exc}")

    def save_result_csv(self) -> None:
        if not self.rows_data:
            messagebox.showwarning("未読み込み", "先にCSVファイルを読み込んでください。")
            return

        selected = filedialog.asksaveasfilename(
            title="結果CSVを保存",
            defaultextension=".csv",
            filetypes=[("CSVファイル", "*.csv")],
            initialfile="result.csv",
        )
        if not selected:
            return

        save_path = Path(selected)
        self._write_result_csv(save_path)

        self.result_label.set(f"結果CSV保存: {save_path.name}")
        self._append_log(f"結果CSVを保存しました: {save_path}")
        messagebox.showinfo("保存完了", f"結果CSVを保存しました\n{save_path}")
