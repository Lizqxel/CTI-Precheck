import csv
import json
import queue
import re
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional, Set, Tuple

from services.area_search import clear_cancel_flag as clear_cancel_flag_west
from services.area_search import search_service_area
from services.area_search import set_cancel_flag as set_cancel_flag_west
from services import area_search_east
from utils.address_utils import normalize_address

ZIP_PATTERN = re.compile(r"^\d{3}-?\d{4}$")
SETTINGS_PATH = Path("settings.json")


def _decode_csv_bytes(file_bytes: bytes) -> str:
    for encoding in ("utf-8-sig", "cp932", "utf-8"):
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("unknown", b"", 0, 1, "サポートされていないエンコーディングです")


def _normalize_zipcode(value: str) -> str:
    cleaned = (value or "").strip().replace("－", "-").replace("ー", "-")
    digits_only = re.sub(r"\D", "", cleaned)
    if len(digits_only) == 7:
        return f"{digits_only[:3]}-{digits_only[3:]}"
    return cleaned


def _validate_rows(rows: List[List[str]]) -> Tuple[List[Dict[str, str]], List[int]]:
    parsed: List[Dict[str, str]] = []
    invalid_line_numbers: List[int] = []

    for index, row in enumerate(rows, start=1):
        zipcode = row[0].strip() if len(row) >= 1 and row[0] is not None else ""
        address = row[1].strip() if len(row) >= 2 and row[1] is not None else ""

        normalized_zipcode = _normalize_zipcode(zipcode)
        normalized_address = normalize_address(address) if address else ""

        status = "OK"
        if not zipcode and not address:
            status = "空行"
        elif not zipcode or not address:
            status = "入力不足"
            invalid_line_numbers.append(index)
        elif not ZIP_PATTERN.match(normalized_zipcode):
            status = "郵便番号形式エラー"
            invalid_line_numbers.append(index)

        parsed.append(
            {
                "行": str(index),
                "郵便番号": normalized_zipcode,
                "住所": normalized_address,
                "状態": status,
                "判定結果": "未実行",
                "備考": "",
            }
        )

    return parsed, invalid_line_numbers


def _read_csv(file_path: Path) -> List[List[str]]:
    file_bytes = file_path.read_bytes()
    text = _decode_csv_bytes(file_bytes)
    reader = csv.reader(text.splitlines())
    return [row for row in reader]


def _map_result(result: Dict[str, object]) -> str:
    status = str(result.get("status", "")).lower()
    message = str(result.get("message", ""))
    details = result.get("details") if isinstance(result.get("details"), dict) else {}
    note = str(details.get("備考", "")) if isinstance(details, dict) else ""
    area_text = str(details.get("提供エリア", "")) if isinstance(details, dict) else ""

    if "要手動再検索" in message or "調査" in message or "調査" in note or "調査" in area_text:
        return "要調査"

    if status == "available":
        return "提供可能"
    if status == "unavailable":
        return "未提供"
    if status == "cancelled":
        return "停止"
    if "未提供" in message:
        return "未提供"
    return "失敗"


def _extract_note(result: Dict[str, object]) -> str:
    search_notes = result.get("search_notes")
    if isinstance(search_notes, list):
        merged = " / ".join(str(item).strip() for item in search_notes if str(item).strip())
        if merged:
            return merged

    details = result.get("details")
    if isinstance(details, dict):
        note = str(details.get("備考", "")).strip()
        if note:
            return note

    return ""


class DesktopApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("提供判定CSVツール（デスクトップ版）")
        self.root.geometry("1320x760")

        self.rows_data: List[Dict[str, str]] = []
        self.event_queue: queue.Queue[Tuple[str, object]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.running = False
        self.stop_requested = False

        self.total_label = tk.StringVar(value="総行数: 0")
        self.file_label = tk.StringVar(value="未選択")
        self.result_label = tk.StringVar(value="CSVファイルを選択してください。")
        self.progress_label = tk.StringVar(value="進捗: -")
        self.monitor_browser_var = tk.BooleanVar(value=False)
        self.show_popup_var = tk.BooleanVar(value=True)
        self.run_scope_var = tk.StringVar(value="全行")
        self.target_line_var = tk.StringVar(value="対象行: 未選択")
        self.execution_target_line: Optional[int] = None

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

        setting_frame = ttk.LabelFrame(self.root, text="設定", padding=10)
        setting_frame.pack(fill=tk.X, padx=12, pady=(0, 8))

        ttk.Checkbutton(setting_frame, text="ブラウザ表示で監視する", variable=self.monitor_browser_var).pack(side=tk.LEFT)
        ttk.Checkbutton(setting_frame, text="判定結果ポップアップを有効化", variable=self.show_popup_var).pack(side=tk.LEFT, padx=(12, 0))
        ttk.Button(setting_frame, text="設定保存", command=self.save_settings).pack(side=tk.LEFT, padx=(16, 0))

        info_frame = ttk.Frame(self.root, padding=(12, 0, 12, 8))
        info_frame.pack(fill=tk.X)
        ttk.Label(info_frame, textvariable=self.total_label).pack(side=tk.LEFT)
        ttk.Label(info_frame, textvariable=self.result_label).pack(side=tk.LEFT, padx=(16, 0))
        ttk.Label(info_frame, textvariable=self.progress_label).pack(side=tk.LEFT, padx=(16, 0))

        table_frame = ttk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 8))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        columns = ("行", "郵便番号", "住所", "状態", "判定結果", "備考")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=20)
        for col in columns:
            self.tree.heading(col, text=col)

        self._tree_column_layout = {
            "行": {"ratio": 0.06, "min": 50, "max": 90, "anchor": tk.CENTER},
            "郵便番号": {"ratio": 0.11, "min": 90, "max": 150, "anchor": tk.CENTER},
            "住所": {"ratio": 0.36, "min": 220, "max": 640, "anchor": tk.W},
            "状態": {"ratio": 0.11, "min": 90, "max": 180, "anchor": tk.CENTER},
            "判定結果": {"ratio": 0.12, "min": 100, "max": 180, "anchor": tk.CENTER},
            "備考": {"ratio": 0.24, "min": 180, "max": 640, "anchor": tk.W},
        }
        self._configure_tree_columns(1320)
        self.tree.bind("<Configure>", self._on_tree_configure)

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=h_scrollbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        self.tree.bind("<Double-1>", self._on_tree_double_click)

        self._bind_tree_scroll()
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_selection)

        note_frame = ttk.LabelFrame(self.root, text="備考詳細", padding=8)
        note_frame.pack(fill=tk.BOTH, expand=False, padx=12, pady=(0, 8))

        self.note_text = tk.Text(note_frame, height=4, wrap=tk.WORD)
        note_scroll = ttk.Scrollbar(note_frame, orient=tk.VERTICAL, command=self.note_text.yview)
        self.note_text.configure(yscrollcommand=note_scroll.set)
        self.note_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        note_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.note_text.configure(state=tk.DISABLED)

        log_frame = ttk.LabelFrame(self.root, text="監視ログ", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=False, padx=12, pady=(0, 12))

        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state=tk.DISABLED)

    def _bind_tree_scroll(self) -> None:
        def on_mousewheel(event: tk.Event) -> str:
            delta = int(event.delta / 120) if event.delta else 0
            if delta == 0:
                return "break"

            if event.state & 0x0001:  # Shift押下時は横スクロール
                self.tree.xview_scroll(-delta * 6, "units")
            else:
                self.tree.yview_scroll(-delta * 4, "units")
            return "break"

        self.tree.bind("<MouseWheel>", on_mousewheel)

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

        # 右端で文字が切れないよう、スクロール終端に余白を追加
        preferred["備考"] += 120

        for col, conf in self._tree_column_layout.items():
            self.tree.column(
                col,
                width=preferred[col],
                minwidth=40,
                anchor=conf["anchor"],
                stretch=False,
            )

    def _load_settings_to_ui(self) -> None:
        if not SETTINGS_PATH.exists():
            return
        try:
            settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
            browser_settings = settings.get("browser_settings", {})
            self.monitor_browser_var.set(not browser_settings.get("headless", True))
            self.show_popup_var.set(browser_settings.get("show_popup", True))
        except Exception:
            pass

    def _build_settings_payload(self) -> Dict[str, Dict[str, object]]:
        return {
            "browser_settings": {
                "headless": not self.monitor_browser_var.get(),
                "show_popup": self.show_popup_var.get(),
                "auto_close": True,
                "page_load_timeout": 60,
                "script_timeout": 60,
            }
        }

    def save_settings(self) -> None:
        payload = self._build_settings_payload()
        SETTINGS_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
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
            rows = _read_csv(file_path)
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

        parsed_rows, invalid_line_numbers = _validate_rows(rows)
        self.rows_data = parsed_rows
        self._render_rows(self.rows_data)
        self.execution_target_line = None
        self.target_line_var.set("対象行: 未選択")
        self.run_scope_var.set("全行")

        self.total_label.set(f"総行数: {len(self.rows_data)}")
        self.result_label.set("CSV読み込み完了")
        self.progress_label.set("進捗: -")
        self._append_log(f"CSVを読み込みました: {file_path.name}")
        if invalid_line_numbers:
            messagebox.showwarning(
                "入力不備のある行",
                f"次の行に入力不備があります: {', '.join(map(str, invalid_line_numbers))}",
            )

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

        target_lines = self._resolve_target_lines()
        if target_lines is not None and len(target_lines) == 0:
            messagebox.showwarning("対象未設定", "実行対象の行を選択してください。")
            return

        total_targets = len(self.rows_data) if target_lines is None else len(target_lines)

        self.save_settings()
        self.stop_requested = False
        self.running = True
        self._set_running_ui_state(True)
        self.result_label.set("提供判定を実行中...")
        self.progress_label.set(f"進捗: 0/{total_targets}")
        self._append_log("提供判定を開始しました")

        self.worker_thread = threading.Thread(target=self._run_judgement, args=(target_lines,), daemon=True)
        self.worker_thread.start()

    def stop_judgement(self) -> None:
        if not self.running:
            return
        self.stop_requested = True
        self._append_log("停止要求を受け付けました")
        self.result_label.set("停止処理中...")
        self._request_cancel_service()

    def _request_cancel_service(self) -> None:
        try:
            set_cancel_flag_west(True)
        except Exception:
            pass
        try:
            area_search_east.set_cancel_flag(True)
        except Exception:
            pass
        self._append_log("停止要求: キャンセルフラグを送信しました（ドライバーは処理側で安全停止）")

    def _clear_cancel_flags(self) -> None:
        try:
            clear_cancel_flag_west()
        except Exception:
            pass
        try:
            area_search_east.clear_cancel_flag()
        except Exception:
            pass

    def _run_judgement(self, target_lines: Optional[Set[int]] = None) -> None:
        failed_rows: List[int] = []
        processed = 0
        total_targets = len(self.rows_data) if target_lines is None else len(target_lines)

        self._clear_cancel_flags()

        for row in self.rows_data:
            line_number = int(row["行"])
            if self.stop_requested:
                break

            if target_lines is not None and line_number not in target_lines:
                continue

            if row["状態"] != "OK":
                row["判定結果"] = "失敗"
                row["備考"] = f"入力不備: {row['状態']}"
                failed_rows.append(line_number)
                processed += 1
                self.event_queue.put(("row", row.copy()))
                self.event_queue.put(("progress", (processed, total_targets)))
                continue

            postal_code = row["郵便番号"]
            address = row["住所"]
            self.event_queue.put(("log", f"{line_number}行目を判定中: {postal_code} {address}"))

            def progress_callback(message: str, row_no: int = line_number) -> None:
                self.event_queue.put(("log", f"{row_no}行目: {message}"))

            try:
                self._clear_cancel_flags()
                result = search_service_area(postal_code, address, progress_callback=progress_callback)
                judgement = _map_result(result if isinstance(result, dict) else {})
                row["判定結果"] = judgement
                row["備考"] = _extract_note(result if isinstance(result, dict) else {})
                if judgement == "失敗":
                    failed_rows.append(line_number)
            except Exception as exc:
                row["判定結果"] = "失敗"
                row["備考"] = f"実行時エラー: {exc}"
                failed_rows.append(line_number)
                self.event_queue.put(("log", f"{line_number}行目: エラー {exc}"))

            processed += 1
            self.event_queue.put(("row", row.copy()))
            self.event_queue.put(("progress", (processed, total_targets)))

        cancelled = self.stop_requested
        self.event_queue.put(("done", {"failed_rows": failed_rows, "cancelled": cancelled}))
        self._clear_cancel_flags()

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
        self.scope_combo.configure(state=tk.DISABLED if is_running else "readonly")
        self.set_target_button.configure(state=tk.DISABLED if is_running else tk.NORMAL)

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _clear_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _render_rows(self, rows: List[Dict[str, str]]) -> None:
        self._clear_tree()
        for row in rows:
            note_full = row.get("備考", "")
            note_cell = note_full if len(note_full) <= 48 else f"{note_full[:48]}…"
            self.tree.insert(
                "",
                tk.END,
                iid=row["行"],
                values=(row["行"], row["郵便番号"], row["住所"], row["状態"], row["判定結果"], note_cell),
            )

        self._refresh_note_detail()

    def _update_row(self, row: Dict[str, str]) -> None:
        row_id = row["行"]
        if self.tree.exists(row_id):
            note_full = row.get("備考", "")
            note_cell = note_full if len(note_full) <= 48 else f"{note_full[:48]}…"
            self.tree.item(
                row_id,
                values=(row["行"], row["郵便番号"], row["住所"], row["状態"], row["判定結果"], note_cell),
            )

        self._refresh_note_detail()

    def _on_tree_selection(self, event: tk.Event) -> None:
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
                note_value = row.get("備考", "")
                writer.writerow([row["郵便番号"], row["住所"], result_value, note_value])

        self.result_label.set(f"結果CSV保存: {save_path.name}")
        self._append_log(f"結果CSVを保存しました: {save_path}")
        messagebox.showinfo("保存完了", f"結果CSVを保存しました\n{save_path}")


def main() -> None:
    root = tk.Tk()
    app = DesktopApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
