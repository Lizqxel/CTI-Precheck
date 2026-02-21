from typing import Dict


INVESTIGATION_IMAGE_MESSAGE = "要手動再検索（住所をご確認ください）"
INVESTIGATION_IMAGE_NOTE = "「住所を特定できないため、担当者がお調べします」の画像有"


def _append_unique(parts: list[str], value: str) -> None:
    normalized = (value or "").strip()
    if not normalized:
        return

    segments = [segment.strip() for segment in normalized.split("/") if segment.strip()]
    if not segments:
        return

    for segment in segments:
        if segment not in parts:
            parts.append(segment)


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
    note_parts: list[str] = []

    details = result.get("details")
    if isinstance(details, dict):
        _append_unique(note_parts, str(details.get("備考", "")))

    search_notes = result.get("search_notes")
    if isinstance(search_notes, list):
        for item in search_notes:
            _append_unique(note_parts, str(item))

    message = str(result.get("message", "")).strip()
    if INVESTIGATION_IMAGE_MESSAGE in message:
        _append_unique(note_parts, INVESTIGATION_IMAGE_NOTE)

    if note_parts:
        return " / ".join(note_parts)

    return ""
