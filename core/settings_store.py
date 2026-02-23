import json
from pathlib import Path
from typing import Any, Dict

SETTINGS_PATH = Path("settings.json")

DEFAULT_BROWSER_SETTINGS: Dict[str, Any] = {
    "headless": True,
    "show_popup": True,
    "auto_close": True,
    "page_load_timeout": 60,
    "script_timeout": 60,
    "disable_images": True,
    "page_load_strategy": "eager",
}

DEFAULT_UPDATE_SETTINGS: Dict[str, Any] = {
    "channel": "stable",
    "etag": "",
    "cached_release": {},
    "last_checked_at": "",
    "last_latest_version": "",
    "skipped_version": "",
    "auto_check_interval_hours": 24,
    "last_result": {},
}

UPDATE_HISTORY_MAX_LENGTH = 30


def _load_all_settings(settings_path: Path = SETTINGS_PATH) -> Dict[str, Any]:
    if not settings_path.exists():
        return {}

    try:
        loaded = json.loads(settings_path.read_text(encoding="utf-8"))
        if isinstance(loaded, dict):
            return loaded
        return {}
    except Exception:
        return {}


def _save_all_settings(payload: Dict[str, Any], settings_path: Path = SETTINGS_PATH) -> None:
    settings_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def load_browser_settings(settings_path: Path = SETTINGS_PATH) -> Dict[str, Any]:
    settings = _load_all_settings(settings_path)
    browser_settings = settings.get("browser_settings", {})
    merged = dict(DEFAULT_BROWSER_SETTINGS)
    if isinstance(browser_settings, dict):
        merged.update(browser_settings)
    return merged


def build_settings_payload(browser_settings: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    merged = dict(DEFAULT_BROWSER_SETTINGS)
    merged.update(browser_settings)
    return {"browser_settings": merged}


def save_browser_settings(browser_settings: Dict[str, Any], settings_path: Path = SETTINGS_PATH) -> None:
    payload = _load_all_settings(settings_path)
    payload.update(build_settings_payload(browser_settings))
    _save_all_settings(payload, settings_path)


def load_update_settings(settings_path: Path = SETTINGS_PATH) -> Dict[str, Any]:
    settings = _load_all_settings(settings_path)
    update_settings = settings.get("update_settings", {})
    merged = dict(DEFAULT_UPDATE_SETTINGS)
    if isinstance(update_settings, dict):
        merged.update(update_settings)
    return merged


def save_update_settings(update_settings: Dict[str, Any], settings_path: Path = SETTINGS_PATH) -> None:
    payload = _load_all_settings(settings_path)
    merged = dict(DEFAULT_UPDATE_SETTINGS)
    merged.update(update_settings)
    payload["update_settings"] = merged
    _save_all_settings(payload, settings_path)


def append_update_history(history_item: Dict[str, Any], settings_path: Path = SETTINGS_PATH) -> None:
    payload = _load_all_settings(settings_path)
    history = payload.get("update_history", [])
    if not isinstance(history, list):
        history = []

    history.append(history_item)
    payload["update_history"] = history[-UPDATE_HISTORY_MAX_LENGTH:]
    _save_all_settings(payload, settings_path)
