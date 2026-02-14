from typing import Dict


def map_result(result: Dict[str, object]) -> str:
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
