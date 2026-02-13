"""
WebDriverの作成と管理を行うモジュール

このモジュールは、Selenium WebDriverの作成と管理を
担当します。
"""

import logging
import json
import os
import time
import sys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.os_manager import ChromeType


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
        
        # Chromeオプションの設定
        chrome_options = Options()
        if headless:
            chrome_options.add_argument('--headless')
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
        chrome_options.add_argument('--start-maximized')
        chrome_options.add_argument('--window-size=1920,1080')
        
        # 画像を無効化
        prefs = {
            'profile.managed_default_content_settings.images': 2,
            'profile.default_content_setting_values.images': 2
        }
        chrome_options.add_experimental_option('prefs', prefs)
        
        # キャンセルチェック（オプション設定後）
        try:
            from services.area_search import check_cancellation
            check_cancellation()
        except (ImportError, NameError):
            pass
        
        # ChromeDriverManagerの設定
        driver_manager = ChromeDriverManager()
        service = Service(driver_manager.install())
        
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
        driver.implicitly_wait(10)
        
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
        "show_popup": False,
        "auto_close": False
    }
    
    try:
        if os.path.exists("settings.json"):
            with open("settings.json", "r", encoding="utf-8") as f:
                settings = json.load(f)
                # ブラウザ設定が含まれていない場合はデフォルト値を使用
                browser_settings = settings.get("browser_settings", default_settings)
                
                # auto_closeが設定に含まれていない場合はデフォルト値を使用
                if "auto_close" not in browser_settings:
                    browser_settings["auto_close"] = default_settings["auto_close"]
                    
                return browser_settings
    except Exception as e:
        logging.warning(f"ブラウザ設定の読み込みに失敗しました: {str(e)}")
    
    return default_settings 