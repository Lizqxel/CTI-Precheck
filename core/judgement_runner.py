import queue
from typing import Callable, Dict, List, Tuple

from services.area_search import search_service_area

from core.cancellation import clear_cancel_flags
from core.result_mapping import map_result


EventQueue = queue.Queue[Tuple[str, object]]


def run_judgement(
    rows_data: List[Dict[str, str]],
    event_queue: EventQueue,
    stop_requested: Callable[[], bool],
) -> None:
    failed_rows: List[int] = []
    processed = 0

    clear_cancel_flags()

    for row in rows_data:
        line_number = int(row["行"])
        if stop_requested():
            break

        if row["状態"] != "OK":
            row["判定結果"] = "失敗"
            failed_rows.append(line_number)
            processed += 1
            event_queue.put(("row", row.copy()))
            event_queue.put(("progress", (processed, len(rows_data))))
            continue

        postal_code = row["郵便番号"]
        address = row["住所"]
        event_queue.put(("log", f"{line_number}行目を判定中: {postal_code} {address}"))

        def progress_callback(message: str, row_no: int = line_number) -> None:
            event_queue.put(("log", f"{row_no}行目: {message}"))

        try:
            clear_cancel_flags()
            result = search_service_area(postal_code, address, progress_callback=progress_callback)
            judgement = map_result(result if isinstance(result, dict) else {})
            row["判定結果"] = judgement
            if judgement == "失敗":
                failed_rows.append(line_number)
        except Exception as exc:
            row["判定結果"] = "失敗"
            failed_rows.append(line_number)
            event_queue.put(("log", f"{line_number}行目: エラー {exc}"))

        processed += 1
        event_queue.put(("row", row.copy()))
        event_queue.put(("progress", (processed, len(rows_data))))

    cancelled = stop_requested()
    event_queue.put(("done", {"failed_rows": failed_rows, "cancelled": cancelled}))
    clear_cancel_flags()
