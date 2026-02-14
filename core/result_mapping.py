from typing import Dict


def map_result(result: Dict[str, object]) -> str:
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


def extract_note(result: Dict[str, object]) -> str:
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
