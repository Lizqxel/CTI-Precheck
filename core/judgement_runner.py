import queue
import threading
from typing import Callable, Dict, List, Optional, Set, Tuple

from services.area_search import CancellationError
from services.area_search import search_service_area

from core.cancellation import clear_cancel_flags
from core.result_mapping import extract_note, map_result


EventQueue = queue.Queue[Tuple[str, object]]


def run_judgement(
    rows_data: List[Dict[str, str]],
    event_queue: EventQueue,
    stop_requested: Callable[[], bool],
    parallel_count: int = 1,
    target_lines: Optional[Set[int]] = None,
) -> None:
    failed_rows: List[int] = []
    failed_rows_lock = threading.Lock()
    processed_lock = threading.Lock()
    processed = 0
    total = len(rows_data) if target_lines is None else len(target_lines)

    effective_parallel = max(1, min(int(parallel_count or 1), 8))

    task_queue: queue.Queue[Dict[str, str]] = queue.Queue()
    for row in rows_data:
        line_number = int(row["行"])
        if target_lines is not None and line_number not in target_lines:
            continue
        task_queue.put(row)

    clear_cancel_flags()

    def process_row(row: Dict[str, str], worker_id: int) -> None:
        nonlocal processed

        line_number = int(row["行"])

        if row["状態"] != "OK":
            row["判定結果"] = "失敗"
            row["備考"] = f"入力不備: {row['状態']}"
            with failed_rows_lock:
                failed_rows.append(line_number)
        else:
            postal_code = row["郵便番号"]
            address = row["住所"]
            event_queue.put(("worker_log", {"worker": worker_id, "message": f"{line_number}行目を判定中: {postal_code} {address}"}))

            def progress_callback(message: str, row_no: int = line_number) -> None:
                event_queue.put(("worker_log", {"worker": worker_id, "message": f"{row_no}行目: {message}"}))

            try:
                result = search_service_area(postal_code, address, progress_callback=progress_callback)
                judgement = map_result(result if isinstance(result, dict) else {})
                row["判定結果"] = judgement
                row["備考"] = extract_note(result if isinstance(result, dict) else {})
                if judgement == "失敗":
                    with failed_rows_lock:
                        failed_rows.append(line_number)
            except CancellationError:
                row["判定結果"] = "停止"
                row["備考"] = "停止要求により中断"
                event_queue.put(("worker_log", {"worker": worker_id, "message": f"{line_number}行目: 停止要求により中断"}))
            except Exception as exc:
                row["判定結果"] = "失敗"
                row["備考"] = f"実行時エラー: {exc}"
                with failed_rows_lock:
                    failed_rows.append(line_number)
                event_queue.put(("worker_log", {"worker": worker_id, "message": f"{line_number}行目: エラー {exc}"}))

        with processed_lock:
            processed += 1
            current = processed

        event_queue.put(("row", row.copy()))
        event_queue.put(("progress", (current, total)))

    def worker_loop(worker_id: int) -> None:
        while True:
            if stop_requested():
                return
            try:
                row = task_queue.get_nowait()
            except queue.Empty:
                return

            try:
                process_row(row, worker_id)
            finally:
                task_queue.task_done()

    workers: List[threading.Thread] = []
    for worker_id in range(effective_parallel):
        thread = threading.Thread(target=worker_loop, args=(worker_id,), daemon=True)
        workers.append(thread)
        thread.start()

    for thread in workers:
        thread.join()

    cancelled = stop_requested()
    failed_rows_sorted = sorted(failed_rows)
    event_queue.put(("done", {"failed_rows": failed_rows_sorted, "cancelled": cancelled}))
    clear_cancel_flags()
