import queue
import threading
import time
from typing import Callable, Dict, List, Optional, Set, Tuple

from services.area_search import CancellationError
from services.area_search import close_active_drivers
from services.area_search import RetryableWebDriverError
from services.area_search import search_service_area

from core.cancellation import clear_cancel_flags
from core.result_mapping import extract_note, map_result


EventQueue = queue.Queue[Tuple[str, object]]


def _is_retryable_driver_failure(note: str) -> bool:
    message = (note or "").lower()
    keywords = (
        "stacktrace",
        "message:",
        "検索結果確認ボタンを検出できず",
        "画面状態から処理を再開できませんでした",
        "webdriverセッションが切断されました",
        "max retries exceeded",
        "failed to establish a new connection",
        "winerror 10061",
        "remote end closed connection",
        "localhost",
        "chrome not reachable",
        "invalid session id",
        "session not created",
        "disconnected",
        "target window already closed",
        "no such window",
        "webview not found",
    )
    return any(keyword in message for keyword in keywords)


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
    retryable_failure_streak_lock = threading.Lock()
    retryable_failure_streak = 0
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
        nonlocal retryable_failure_streak

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
                retry_limit = 3
                result: Dict[str, object] | object = {}
                judgement = "失敗"
                note = ""

                for attempt in range(1, retry_limit + 1):
                    try:
                        result = search_service_area(postal_code, address, progress_callback=progress_callback)
                        mapped_result = result if isinstance(result, dict) else {}
                        judgement = map_result(mapped_result)
                        note = extract_note(mapped_result)
                    except RetryableWebDriverError as retryable_exc:
                        judgement = "失敗"
                        note = "ブラウザ通信エラー（自動再試行対象）"
                        if attempt < retry_limit:
                            event_queue.put((
                                "worker_log",
                                {
                                    "worker": worker_id,
                                    "message": f"{line_number}行目: セッション断を検出したため高速リトライします（{attempt}/{retry_limit}）",
                                },
                            ))
                            if effective_parallel == 1:
                                close_active_drivers()
                            time.sleep(0.25 * attempt)
                            continue
                        note = "ブラウザ通信エラーにより判定できませんでした（再試行後も失敗）"
                        break

                    if judgement != "失敗" or not _is_retryable_driver_failure(note):
                        break

                    with retryable_failure_streak_lock:
                        current_streak = retryable_failure_streak

                    if attempt < retry_limit and current_streak < 3:
                        event_queue.put((
                            "worker_log",
                            {
                                "worker": worker_id,
                                "message": f"{line_number}行目: WebDriverエラーを検出したため再試行します（{attempt}/{retry_limit}）",
                            },
                        ))
                        if effective_parallel == 1:
                            close_active_drivers()
                        time.sleep(0.3 * attempt)
                        continue

                    break

                row["判定結果"] = judgement
                row["備考"] = note

                if judgement == "失敗" and _is_retryable_driver_failure(note):
                    with retryable_failure_streak_lock:
                        retryable_failure_streak += 1
                        if effective_parallel == 1 and retryable_failure_streak % 10 == 0:
                            close_active_drivers()
                else:
                    with retryable_failure_streak_lock:
                        retryable_failure_streak = 0

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
