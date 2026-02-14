import csv
import json
import queue
import re
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Tuple

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
    if status == "available":
        return "提供可能"
    if status == "unavailable":
        return "未提供"
    if status == "cancelled":
        return "停止"
    if "未提供" in message:
        return "未提供"
    return "失敗"


class DesktopApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("提供判定CSVツール（デスクトップ版）")
        self.root.geometry("1160x760")

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
        self.enable_screenshots_var = tk.BooleanVar(value=True)

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
        ttk.Checkbutton(setting_frame, text="判定結果ポップアップを有効化", variable=self.show_popup_var).pack(side=tk.LEFT, padx=(12, 0))
        ttk.Checkbutton(setting_frame, text="スクリーンショット保存", variable=self.enable_screenshots_var).pack(side=tk.LEFT, padx=(12, 0))
        ttk.Button(setting_frame, text="設定保存", command=self.save_settings).pack(side=tk.LEFT, padx=(16, 0))

        info_frame = ttk.Frame(self.root, padding=(12, 0, 12, 8))
        info_frame.pack(fill=tk.X)
        ttk.Label(info_frame, textvariable=self.total_label).pack(side=tk.LEFT)
        ttk.Label(info_frame, textvariable=self.result_label).pack(side=tk.LEFT, padx=(16, 0))
        ttk.Label(info_frame, textvariable=self.progress_label).pack(side=tk.LEFT, padx=(16, 0))

        table_frame = ttk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 8))

        columns = ("行", "郵便番号", "住所", "状態", "判定結果")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=20)
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

        log_frame = ttk.LabelFrame(self.root, text="監視ログ", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=False, padx=12, pady=(0, 12))

        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state=tk.DISABLED)

    def _load_settings_to_ui(self) -> None:
        if not SETTINGS_PATH.exists():
            return
        try:
            settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
            browser_settings = settings.get("browser_settings", {})
            self.monitor_browser_var.set(not browser_settings.get("headless", True))
            self.show_popup_var.set(browser_settings.get("show_popup", True))
            self.enable_screenshots_var.set(browser_settings.get("enable_screenshots", True))
        except Exception:
            pass

    def _build_settings_payload(self) -> Dict[str, Dict[str, object]]:
        return {
            "browser_settings": {
                "headless": not self.monitor_browser_var.get(),
                "show_popup": self.show_popup_var.get(),
                "auto_close": False,
                "page_load_timeout": 60,
                "script_timeout": 60,
                "enable_screenshots": self.enable_screenshots_var.get(),
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
        self.stop_requested = False
        self.running = True
        self._set_running_ui_state(True)
        self.result_label.set("提供判定を実行中...")
        self.progress_label.set(f"進捗: 0/{len(self.rows_data)}")
        self._append_log("提供判定を開始しました")

        self.worker_thread = threading.Thread(target=self._run_judgement, daemon=True)
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

        for module in (None, area_search_east):
            try:
                driver = None
                if module is None:
                    from services import area_search
                    driver = getattr(area_search, "global_driver", None)
                else:
                    driver = getattr(module, "global_driver", None)
                if driver:
                    driver.quit()
            except Exception:
                pass

    def _clear_cancel_flags(self) -> None:
        try:
            clear_cancel_flag_west()
        except Exception:
            pass
        try:
            area_search_east.clear_cancel_flag()
        except Exception:
            pass

    def _run_judgement(self) -> None:
        failed_rows: List[int] = []
        processed = 0

        self._clear_cancel_flags()

        for row in self.rows_data:
            line_number = int(row["行"])
            if self.stop_requested:
                break

            if row["状態"] != "OK":
                row["判定結果"] = "失敗"
                failed_rows.append(line_number)
                processed += 1
                self.event_queue.put(("row", row.copy()))
                self.event_queue.put(("progress", (processed, len(self.rows_data))))
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
                if judgement == "失敗":
                    failed_rows.append(line_number)
            except Exception as exc:
                row["判定結果"] = "失敗"
                failed_rows.append(line_number)
                self.event_queue.put(("log", f"{line_number}行目: エラー {exc}"))

            processed += 1
            self.event_queue.put(("row", row.copy()))
            self.event_queue.put(("progress", (processed, len(self.rows_data))))

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
            self.tree.insert(
                "",
                tk.END,
                iid=row["行"],
                values=(row["行"], row["郵便番号"], row["住所"], row["状態"], row["判定結果"]),
            )

    def _update_row(self, row: Dict[str, str]) -> None:
        row_id = row["行"]
        if self.tree.exists(row_id):
            self.tree.item(row_id, values=(row["行"], row["郵便番号"], row["住所"], row["状態"], row["判定結果"]))

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


def main() -> None:
    root = tk.Tk()
    app = DesktopApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
