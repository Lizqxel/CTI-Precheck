import json
from pathlib import Path
from typing import Any, Dict

SETTINGS_PATH = Path("settings.json")

DEFAULT_BROWSER_SETTINGS: Dict[str, Any] = {
    "headless": True,
    "show_popup": True,
    "auto_close": False,
    "page_load_timeout": 60,
    "script_timeout": 60,
    "enable_screenshots": True,
}


def load_browser_settings(settings_path: Path = SETTINGS_PATH) -> Dict[str, Any]:
    if not settings_path.exists():
        return dict(DEFAULT_BROWSER_SETTINGS)

    try:
        settings = json.loads(settings_path.read_text(encoding="utf-8"))
        browser_settings = settings.get("browser_settings", {})
        merged = dict(DEFAULT_BROWSER_SETTINGS)
        merged.update(browser_settings)
        return merged
    except Exception:
        return dict(DEFAULT_BROWSER_SETTINGS)


def build_settings_payload(browser_settings: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    merged = dict(DEFAULT_BROWSER_SETTINGS)
    merged.update(browser_settings)
    return {"browser_settings": merged}


def save_browser_settings(browser_settings: Dict[str, Any], settings_path: Path = SETTINGS_PATH) -> None:
    payload = build_settings_payload(browser_settings)
    settings_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
