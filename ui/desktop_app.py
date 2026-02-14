import csv
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Tuple

from core.cancellation import request_cancel_service
from core.csv_processing import read_csv, validate_rows
from core.judgement_runner import run_judgement
from core.settings_store import SETTINGS_PATH, load_browser_settings, save_browser_settings


EventQueue = queue.Queue[Tuple[str, object]]


class DesktopApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("提供判定CSVツール（デスクトップ版）")
        self.root.geometry("1160x760")

        self.rows_data: List[Dict[str, str]] = []
        self.event_queue: EventQueue = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.running = False
        self.stop_requested_flag = False

        self.total_label = tk.StringVar(value="総行数: 0")
        self.file_label = tk.StringVar(value="未選択")
        self.result_label = tk.StringVar(value="CSVファイルを選択してください。")
        self.progress_label = tk.StringVar(value="進捗: -")
        self.monitor_browser_var = tk.BooleanVar(value=False)
        self.show_popup_var = tk.BooleanVar(value=True)
        self.enable_screenshots_var = tk.BooleanVar(value=True)
        self.parallel_count_var = tk.IntVar(value=2)
        self.parallel_count_values = (1, 2, 3, 4)

        self.worker_log_texts: List[tk.Text] = []
        self.worker_logs_container: ttk.Frame | None = None

        self._load_settings_to_ui()
        self._build_ui()
        self.root.after(150, self._drain_event_queue)

    def _build_ui(self) -> None:
        top_frame = ttk.Frame(self.root, padding=12)
        top_frame.pack(fill=tk.X)

        self.select_button = ttk.Button(top_frame, text="CSVファイルを選択", command=self.load_csv)
        self.select_button.pack(side=tk.LEFT)

        self.start_button = ttk.Button(top_frame, text="提供判定開始", command=self.start_judgement)
        self.start_button.pack(side=tk.LEFT, padx=(8, 0))

        self.stop_button = ttk.Button(top_frame, text="停止", command=self.stop_judgement, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=(8, 0))

        self.save_button = ttk.Button(top_frame, text="結果CSV保存", command=self.save_result_csv)
        self.save_button.pack(side=tk.LEFT, padx=(8, 0))

        ttk.Label(top_frame, textvariable=self.file_label).pack(side=tk.LEFT, padx=(12, 0))

        setting_frame = ttk.LabelFrame(self.root, text="設定", padding=10)
        setting_frame.pack(fill=tk.X, padx=12, pady=(0, 8))

        ttk.Checkbutton(setting_frame, text="ブラウザ表示で監視する", variable=self.monitor_browser_var).pack(side=tk.LEFT)
        ttk.Checkbutton(setting_frame, text="判定結果ポップアップを有効化", variable=self.show_popup_var).pack(
            side=tk.LEFT, padx=(12, 0)
        )
        ttk.Checkbutton(setting_frame, text="スクリーンショット保存", variable=self.enable_screenshots_var).pack(
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

        info_frame = ttk.Frame(self.root, padding=(12, 0, 12, 8))
        info_frame.pack(fill=tk.X)
        ttk.Label(info_frame, textvariable=self.total_label).pack(side=tk.LEFT)
        ttk.Label(info_frame, textvariable=self.result_label).pack(side=tk.LEFT, padx=(16, 0))
        ttk.Label(info_frame, textvariable=self.progress_label).pack(side=tk.LEFT, padx=(16, 0))

        table_frame = ttk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 8))

        columns = ("行", "郵便番号", "住所", "状態", "判定結果")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=10)
        for col in columns:
            self.tree.heading(col, text=col)

        self.tree.column("行", width=70, anchor=tk.CENTER)
        self.tree.column("郵便番号", width=130, anchor=tk.CENTER)
        self.tree.column("住所", width=560, anchor=tk.W)
        self.tree.column("状態", width=140, anchor=tk.CENTER)
        self.tree.column("判定結果", width=140, anchor=tk.CENTER)

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.LEFT, fill=tk.Y)

        global_log_frame = ttk.LabelFrame(self.root, text="全体ログ", padding=8)
        global_log_frame.pack(fill=tk.BOTH, expand=False, padx=12, pady=(0, 8))

        self.log_text = tk.Text(global_log_frame, height=6, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state=tk.DISABLED)

        worker_log_frame = ttk.LabelFrame(self.root, text="ワーカー別ログ（提供判定実行中）", padding=8)
        worker_log_frame.pack(fill=tk.BOTH, expand=False, padx=12, pady=(0, 12))
        self.worker_logs_container = ttk.Frame(worker_log_frame)
        self.worker_logs_container.pack(fill=tk.BOTH, expand=True)
        self._rebuild_worker_log_panels()

    def _load_settings_to_ui(self) -> None:
        browser_settings = load_browser_settings(SETTINGS_PATH)
        self.monitor_browser_var.set(not browser_settings.get("headless", True))
        self.show_popup_var.set(bool(browser_settings.get("show_popup", True)))
        self.enable_screenshots_var.set(bool(browser_settings.get("enable_screenshots", True)))

    def _build_browser_settings_from_ui(self) -> Dict[str, object]:
        return {
            "headless": not self.monitor_browser_var.get(),
            "show_popup": self.show_popup_var.get(),
            "auto_close": False,
            "page_load_timeout": 60,
            "script_timeout": 60,
            "enable_screenshots": self.enable_screenshots_var.get(),
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
            rows = read_csv(file_path)
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

        self.total_label.set(f"総行数: {len(self.rows_data)}")
        self.result_label.set("CSV読み込み完了")
        self.progress_label.set("進捗: -")
        self._append_log(f"CSVを読み込みました: {file_path.name}")
        if invalid_line_numbers:
            messagebox.showwarning(
                "入力不備のある行",
                f"次の行に入力不備があります: {', '.join(map(str, invalid_line_numbers))}",
            )

    def start_judgement(self) -> None:
        if self.running:
            return
        if not self.rows_data:
            messagebox.showwarning("未読み込み", "先にCSVファイルを選択してください。")
            return

        self.save_settings()
        self.stop_requested_flag = False
        self.running = True
        self._set_running_ui_state(True)
        self.result_label.set("提供判定を実行中...")
        self.progress_label.set(f"進捗: 0/{len(self.rows_data)}")
        self._append_log("提供判定を開始しました")
        self._rebuild_worker_log_panels(clear_existing=True)

        self.worker_thread = threading.Thread(target=self._run_judgement, daemon=True)
        self.worker_thread.start()

    def stop_judgement(self) -> None:
        if not self.running:
            return
        self.stop_requested_flag = True
        self._append_log("停止要求を受け付けました")
        self.result_label.set("停止処理中...")
        request_cancel_service()

    def _run_judgement(self) -> None:
        run_judgement(
            rows_data=self.rows_data,
            event_queue=self.event_queue,
            stop_requested=lambda: self.stop_requested_flag,
            parallel_count=self._get_parallel_count(),
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

        if cancelled:
            self.result_label.set("提供判定を停止しました")
            self._append_log("提供判定を停止しました")
            messagebox.showinfo("停止", "提供判定を停止しました。")
            return

        if failed_rows:
            self.result_label.set("提供判定完了（失敗あり）")
            lines = ", ".join(map(str, failed_rows))
            self._append_log(f"提供判定完了: 失敗行 {lines}")
            messagebox.showwarning("失敗行", f"以下の行が失敗しました: {lines}")
        else:
            self.result_label.set("提供判定完了")
            self._append_log("提供判定が完了しました")
            messagebox.showinfo("完了", "提供判定が完了しました。")

    def _set_running_ui_state(self, is_running: bool) -> None:
        self.select_button.configure(state=tk.DISABLED if is_running else tk.NORMAL)
        self.start_button.configure(state=tk.DISABLED if is_running else tk.NORMAL)
        self.stop_button.configure(state=tk.NORMAL if is_running else tk.DISABLED)
        self.parallel_count_combo.configure(state=tk.DISABLED if is_running else "readonly")

    def _on_parallel_count_changed(self, _event: object | None = None) -> None:
        if self.running:
            return
        self._rebuild_worker_log_panels()

    def _get_parallel_count(self) -> int:
        value = self.parallel_count_var.get()
        if value not in self.parallel_count_values:
            return 2
        return int(value)

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
                values=(row["行"], row["郵便番号"], row["住所"], row["状態"], row["判定結果"]),
            )

    def _update_row(self, row: Dict[str, str]) -> None:
        row_id = row["行"]
        if self.tree.exists(row_id):
            self.tree.item(
                row_id,
                values=(row["行"], row["郵便番号"], row["住所"], row["状態"], row["判定結果"]),
            )

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
        with save_path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            for row in self.rows_data:
                result_value = row.get("判定結果", "未実行")
                writer.writerow([row["郵便番号"], row["住所"], result_value])

        self.result_label.set(f"結果CSV保存: {save_path.name}")
        self._append_log(f"結果CSVを保存しました: {save_path}")
        messagebox.showinfo("保存完了", f"結果CSVを保存しました\n{save_path}")
