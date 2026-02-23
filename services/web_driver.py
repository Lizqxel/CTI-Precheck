"""
WebDriverの作成と管理を行うモジュール

このモジュールは、Selenium WebDriverの作成と管理を
担当します。
"""

import logging
import json
import os
import threading
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager


_DRIVER_PATH_LOCK = threading.Lock()
_CACHED_DRIVER_PATH = None


def _apply_render_optimizations(driver, aggressive_blocking: bool = True) -> None:
    try:
        driver.execute_cdp_cmd("Network.enable", {})
    except Exception:
        return

    if aggressive_blocking:
        blocked_patterns = [
            "*.png", "*.jpg", "*.jpeg", "*.gif", "*.webp", "*.svg", "*.ico",
            "*.woff", "*.woff2", "*.ttf", "*.otf",
            "*.mp4", "*.webm", "*.mp3", "*.m4a", "*.wav",
            "data:image/*",
        ]
        try:
            driver.execute_cdp_cmd("Network.setBlockedURLs", {"urls": blocked_patterns})
        except Exception:
            pass

    animation_killer_script = """
        (() => {
            try {
                const style = document.createElement('style');
                style.id = '__cti_render_light_style';
                style.textContent = `
                    *, *::before, *::after {
                        animation: none !important;
                        transition: none !important;
                        caret-color: transparent !important;
                    }
                    html { scroll-behavior: auto !important; }
                `;
                document.documentElement.appendChild(style);
            } catch (e) {}
        })();
    """

    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {"source": animation_killer_script},
        )
    except Exception:
        pass


def _resolve_chromedriver_path() -> str:
    global _CACHED_DRIVER_PATH
    if _CACHED_DRIVER_PATH:
        return _CACHED_DRIVER_PATH

    with _DRIVER_PATH_LOCK:
        if _CACHED_DRIVER_PATH:
            return _CACHED_DRIVER_PATH
        _CACHED_DRIVER_PATH = ChromeDriverManager().install()
        return _CACHED_DRIVER_PATH


def create_driver(headless=False):
    """
    Chrome WebDriverを作成する
    
    Args:
        headless (bool): ヘッドレスモードで実行するかどうか
        
    Returns:
        WebDriver: 作成されたWebDriverインスタンス
    """
    try:
        # キャンセルチェック（ドライバー作成開始時）
        try:
            from services.area_search import check_cancellation
            check_cancellation()
        except (ImportError, NameError):
            pass  # area_searchモジュールが利用できない場合はスキップ
        
        browser_settings = load_browser_settings()
        disable_images = bool(browser_settings.get("disable_images", True))
        aggressive_resource_blocking = bool(browser_settings.get("aggressive_resource_blocking", True))
        page_load_strategy = str(browser_settings.get("page_load_strategy", "eager")).strip().lower() or "eager"
        if page_load_strategy not in {"normal", "eager", "none"}:
            page_load_strategy = "eager"

        # Chromeオプションの設定
        chrome_options = Options()
        chrome_options.page_load_strategy = page_load_strategy

        if headless:
            chrome_options.add_argument('--headless=new')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--disable-infobars')
        chrome_options.add_argument('--disable-notifications')
        chrome_options.add_argument('--disable-popup-blocking')
        chrome_options.add_argument('--disable-save-password-bubble')
        chrome_options.add_argument('--disable-translate')
        chrome_options.add_argument('--disable-web-security')
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--ignore-ssl-errors')
        chrome_options.add_argument('--disable-background-networking')
        chrome_options.add_argument('--disable-backgrounding-occluded-windows')
        chrome_options.add_argument('--disable-renderer-backgrounding')
        chrome_options.add_argument('--disable-component-update')
        chrome_options.add_argument('--metrics-recording-only')
        chrome_options.add_argument('--mute-audio')
        chrome_options.add_argument('--no-first-run')
        chrome_options.add_argument('--no-default-browser-check')
        chrome_options.add_argument('--password-store=basic')
        chrome_options.add_argument('--window-size=1280,720')
        chrome_options.add_argument('--log-level=3')

        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        prefs = {
            'profile.default_content_setting_values.notifications': 2,
            'profile.default_content_setting_values.geolocation': 2,
        }
        if disable_images:
            prefs['profile.managed_default_content_settings.images'] = 2
            prefs['profile.default_content_setting_values.images'] = 2
        chrome_options.add_experimental_option('prefs', prefs)
        
        # キャンセルチェック（オプション設定後）
        try:
            from services.area_search import check_cancellation
            check_cancellation()
        except (ImportError, NameError):
            pass
        
        # ChromeDriverManagerの設定（解決結果をキャッシュ）
        service = Service(_resolve_chromedriver_path())
        
        # キャンセルチェック（ドライバー起動直前）
        try:
            from services.area_search import check_cancellation
            check_cancellation()
        except (ImportError, NameError):
            pass
        
        # WebDriverの作成
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        # キャンセルチェック（ドライバー作成直後）
        try:
            from services.area_search import check_cancellation
            check_cancellation()
        except (ImportError, NameError):
            pass
        
        # タイムアウト設定
        driver.set_page_load_timeout(60)
        driver.implicitly_wait(0)
        _apply_render_optimizations(driver, aggressive_blocking=aggressive_resource_blocking)
        
        return driver
        
    except Exception as e:
        logging.error(f"WebDriverの作成に失敗: {str(e)}")
        raise

def load_browser_settings():
    """
    ブラウザ設定をファイルから読み込む
    
    Returns:
        dict: ブラウザ設定
    """
    default_settings = {
        "headless": True,
        "page_load_timeout": 30,
        "script_timeout": 30,
        "disable_images": True,
        "aggressive_resource_blocking": True,
        "show_popup": False,
        "auto_close": False
    }
    
    try:
        if os.path.exists("settings.json"):
            with open("settings.json", "r", encoding="utf-8") as f:
                settings = json.load(f)
                # ブラウザ設定が含まれていない場合はデフォルト値を使用
                browser_settings = settings.get("browser_settings", {})
                merged_settings = dict(default_settings)
                if isinstance(browser_settings, dict):
                    merged_settings.update(browser_settings)
                
                # auto_closeが設定に含まれていない場合はデフォルト値を使用
                if "auto_close" not in merged_settings:
                    merged_settings["auto_close"] = default_settings["auto_close"]
                    
                return merged_settings
    except Exception as e:
        logging.warning(f"ブラウザ設定の読み込みに失敗しました: {str(e)}")
    
    return default_settings 