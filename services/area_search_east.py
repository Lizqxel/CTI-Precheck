"""
NTT東日本の提供エリア検索サービス

このモジュールは、NTT東日本の提供エリア検索を
自動化するための機能を提供します。

主な機能：
- 郵便番号による住所検索
- 住所の自動選択
- 番地・号の入力
- 建物情報の選択
- 提供エリア判定

制限事項：
- キャンセル機能は使用しません
- エラー発生時は詳細なログを出力します
"""

import logging
import time
import re
import os
import json
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from bs4 import BeautifulSoup
from datetime import datetime

from services.web_driver import create_driver, load_browser_settings
from utils.string_utils import normalize_string, calculate_similarity
from utils.address_utils import normalize_address
from services.area_search import take_full_page_screenshot, check_cancellation, CancellationError

# グローバル変数でブラウザドライバーを保持
global_driver = None

# キャンセルフラグ
_global_cancel_flag = False

def set_cancel_flag(value=True):
    """キャンセルフラグを設定
    
    Args:
        value (bool): 設定する値（デフォルト: True）
    """
    global _global_cancel_flag
    _global_cancel_flag = value
    if value:
        logging.info("★★★ エリア検索のキャンセルフラグを設定しました ★★★")
    else:
        logging.info("エリア検索のキャンセルフラグをクリアしました")

def clear_cancel_flag():
    """キャンセルフラグをクリア"""
    global _global_cancel_flag
    _global_cancel_flag = False
    logging.info("エリア検索のキャンセルフラグをクリアしました")

def is_cancelled():
    """キャンセルフラグの状態を確認（互換性のため）"""
    global _global_cancel_flag
    return _global_cancel_flag

def check_cancellation():
    """キャンセル要求をチェックし、必要に応じて例外を発生"""
    global _global_cancel_flag
    if _global_cancel_flag:
        logging.info("★★★ キャンセル要求を検出：検索を中断します ★★★")
        raise CancellationError("検索がキャンセルされました")

def split_address(address):
    """
    住所を分割して各要素を抽出する

    Args:
        address (str): 分割する住所
        
    Returns:
        dict: 分割された住所の要素
            - prefecture: 都道府県
            - city: 市区町村
            - town: 町名（大字・字を除く）
            - block: 丁目
            - number: 番地
            - building_id: 建物ID
    """
    try:
        # 住所を正規化
        address = normalize_address(address)
        logging.info(f"正規化後の住所: {address}")
        
        # 都道府県を抽出
        prefecture_match = re.match(r'^(.+?[都道府県])', address)
        prefecture = prefecture_match.group(1) if prefecture_match else None
        block = None
        number_part = None
        town = ""
        
        if prefecture:
            remaining_address = address[len(prefecture):].strip()
            # 市区町村を抽出（郡がある場合も考慮）
            city_match = re.match(r'^(.+?郡.+?[町村]|.+?[市区町村])', remaining_address)
            city = city_match.group(1) if city_match else None
            
            if city:
                # 残りの住所から基本住所と番地を分離
                remaining = remaining_address[len(city):].strip()
                
                # 特殊な表記（大字、字、甲乙丙丁）を含む部分を抽出
                special_location_match = re.search(r'(大字.+?字.*?|大字.*?|字.*?)([甲乙丙丁])?(\d+)', remaining)
                
                if special_location_match:
                    # 番地のみを抽出（甲乙丙丁は除外）
                    number_part = special_location_match.group(3)
                    # 基本住所は市区町村までとする
                    town = ""
                else:
                    # 丁目を含む場合は、丁目の後ろの番地を抽出
                    chome_match = re.search(r'(\d+)丁目', remaining)
                    if chome_match:
                        # 丁目より後ろの部分から番地を探す
                        after_chome = remaining[remaining.find('丁目') + 2:].strip()
                        # ハイフンを含む番地のパターンを優先的に検索
                        number_match = re.search(r'(\d+(?:[-－]\d+)?)', after_chome)
                        if number_match:
                            number_part = number_match.group(1)
                            town = remaining[:remaining.find('丁目') - len(chome_match.group(1))].strip()
                        else:
                            number_part = None
                            town = remaining[:remaining.find('丁目')].strip()
                        block = chome_match.group(1)
                    else:
                        # ハイフンが2つある場合は、最初の数字を丁目として扱う
                        double_hyphen_match = re.search(r'(\d+)-(\d+)-(\d+)', remaining)
                        if double_hyphen_match:
                            block = double_hyphen_match.group(1)
                            number_part = f"{double_hyphen_match.group(2)}-{double_hyphen_match.group(3)}"
                            town = remaining[:double_hyphen_match.start()].strip()
                        else:
                            # 通常の番地パターンを検索
                            number_match = re.search(r'(\d+(?:[-－]\d+)?)', remaining)
                            if number_match:
                                number_part = number_match.group(1)
                                town = remaining[:number_match.start()].strip()
                            else:
                                number_part = None
                                town = remaining
                
                result = {
                    'prefecture': prefecture,
                    'city': city,
                    'town': town if town else "",
                    'block': block,
                    'number': number_part,
                    'building_id': None
                }
                
                # 分割結果をログ出力
                logging.info("住所分割結果:")
                logging.info(f"  都道府県: {result['prefecture']}")
                logging.info(f"  市区町村: {result['city']}")
                logging.info(f"  町名: {result['town']}")
                logging.info(f"  丁目: {result['block']}")
                logging.info(f"  番地: {result['number']}")
                
                return result
            
            return {
                'prefecture': prefecture,
                'city': remaining_address if prefecture else None,
                'town': "",
                'block': None,
                'number': None,
                'building_id': None
            }
        
        return {
            'prefecture': None,
            'city': None,
            'town': "",
            'block': None,
            'number': None,
            'building_id': None
        }
        
    except Exception as e:
        logging.error(f"住所の分割中にエラー: {str(e)}")
        return None

def normalize_address(address):
    """
    住所文字列を正規化する関数
    
    Args:
        address (str): 正規化する住所文字列
        
    Returns:
        str: 正規化された住所文字列
    """
    # 空白文字の正規化
    address = address.replace('　', ' ').strip()
    
    # ハイフンの正規化
    address = address.replace('−', '-').replace('ー', '-').replace('－', '-')
    
    # 数字の正規化（全角→半角）
    zen_to_han = str.maketrans('０１２３４５６７８９', '0123456789')
    address = address.translate(zen_to_han)
    
    return address

def is_address_match(input_address, candidate_address):
    """
    入力された住所と候補の住所が一致するかを判定する
    
    Args:
        input_address (str): 入力された住所
        candidate_address (str): 候補の住所
        
    Returns:
        tuple: (bool, float) 一致判定と類似度
    """
    # 両方の住所を正規化
    normalized_input = normalize_string(input_address)
    normalized_candidate = normalize_string(candidate_address)
    
    logging.info(f"住所比較 - 入力: {normalized_input} vs 候補: {normalized_candidate}")
    
    # 完全一致の場合
    if normalized_input == normalized_candidate:
        logging.info("完全一致しました")
        return True, 1.0
    
    # 類似度を計算
    similarity = calculate_similarity(normalized_input, normalized_candidate)
    logging.info(f"類似度: {similarity}")
    
    # 類似度が0.8以上なら一致とみなす
    if similarity >= 0.8:
        logging.info(f"十分な類似度（{similarity}）で一致と判定")
        return True, similarity
    
    # 基本的な住所部分の比較
    input_parts = set(normalized_input.split())
    candidate_parts = set(normalized_candidate.split())
    
    # 共通部分の割合を計算
    common_parts = input_parts & candidate_parts
    if len(input_parts) > 0:
        part_similarity = len(common_parts) / len(input_parts)
        logging.info(f"共通部分の割合: {part_similarity}")
        
        if part_similarity >= 0.8:
            logging.info(f"共通部分の割合（{part_similarity}）で一致と判定")
            return True, part_similarity
    
    logging.info(f"不一致と判定（類似度: {similarity}）")
    return False, similarity

def find_best_address_match(input_address, candidates):
    """
    最も一致度の高い住所を見つける
    
    Args:
        input_address (str): 入力された住所
        candidates (list): 候補の住所リスト
        
    Returns:
        tuple: (最も一致する候補, 類似度)
    """
    best_candidate = None
    best_similarity = -1
    
    # 入力住所を分割
    input_parts = split_address(input_address)
    if not input_parts:
        logging.error("入力住所の分割に失敗しました")
        return None, -1
        
    # 基本住所を構築（番地と号を除く）
    base_input_address = f"{input_parts['prefecture']}{input_parts['city']}{input_parts['town']}"
    if input_parts['block']:
        base_input_address += f"{input_parts['block']}丁目"
    
    logging.info(f"基本住所（比較用）: {base_input_address}")
    
    for candidate in candidates:
        try:
            candidate_text = candidate.text.strip()
            # 候補の住所も分割して基本住所を取得
            candidate_parts = split_address(candidate_text)
            if not candidate_parts:
                continue
                
            base_candidate_address = f"{candidate_parts['prefecture']}{candidate_parts['city']}{candidate_parts['town']}"
            if candidate_parts['block']:
                base_candidate_address += f"{candidate_parts['block']}丁目"
            
            # 基本住所での比較
            _, similarity = is_address_match(base_input_address, base_candidate_address)
            
            logging.info(f"候補 '{candidate_text}' の基本住所: {base_candidate_address}")
            logging.info(f"基本住所での類似度: {similarity}")
            
            if similarity > best_similarity:
                best_similarity = similarity
                best_candidate = candidate
        except Exception as e:
            logging.warning(f"候補の処理中にエラー: {str(e)}")
            continue
    
    if best_candidate and best_similarity >= 0.5:  # 最低限の類似度しきい値
        logging.info(f"最適な候補が見つかりました（類似度: {best_similarity}）: {best_candidate.text.strip()}")
        return best_candidate, best_similarity
    
    return None, best_similarity

def create_driver(headless=False):
    """
    Chromeドライバーを作成する関数
    
    Args:
        headless (bool): ヘッドレスモードで起動するかどうか
        
    Returns:
        WebDriver: 作成されたドライバーインスタンス
    """
    try:
        options = webdriver.ChromeOptions()
        
        if headless:
            options.add_argument('--headless=new')
            # ヘッドレスモード用の追加設定
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-software-rasterizer')
            
        # 共通の最適化設定
        options.add_argument('--window-size=800,600')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-infobars')
        options.add_argument('--disable-notifications')
        options.add_argument('--disable-popup-blocking')
        
        # メモリ使用量の最適化
        options.add_argument('--disable-application-cache')
        options.add_argument('--aggressive-cache-discard')
        options.add_argument('--disable-default-apps')
        
        driver = webdriver.Chrome(options=options)
        logging.info(f"Chromeドライバーを作成しました（ヘッドレスモード: {headless}）")
        
        return driver
    except Exception as e:
        logging.error(f"ドライバーの作成に失敗: {str(e)}")
        raise
    
def search_service_area(postal_code, address, progress_callback=None):
    """
    NTT東日本の提供エリア検索を実行する関数
    
    Args:
        postal_code (str): 郵便番号
        address (str): 住所
        progress_callback (callable): 進捗状況を通知するコールバック関数
        
    Returns:
        dict: 検索結果を含む辞書
    """
    global global_driver
    
    # デバッグログ：入力値の確認
    logging.info(f"=== 検索開始 ===")
    logging.info(f"入力郵便番号（変換前）: {postal_code}")
    logging.info(f"入力住所（変換前）: {address}")
    
    # 郵便番号と住所の正規化
    try:
        postal_code = normalize_address(postal_code)
        address = normalize_address(address)
        logging.info(f"正規化後郵便番号: {postal_code}")
        logging.info(f"正規化後住所: {address}")
    except Exception as e:
        logging.error(f"正規化処理中にエラー: {str(e)}")
        return {"status": "error", "message": f"住所の正規化に失敗しました: {str(e)}"}
    
    # 住所を分割
    if progress_callback:
        progress_callback("住所情報を解析中...")
    
    # キャンセルチェック
    check_cancellation()
    
    address_parts = split_address(address)
    if not address_parts:
        logging.error("住所の分割に失敗しました")
        return {"status": "error", "message": "住所の分割に失敗しました。"}
        
    # 基本住所を構築（番地と号を除く）
    base_address = f"{address_parts['prefecture']}{address_parts['city']}{address_parts['town']}"
    if address_parts['block']:
        base_address += f"{address_parts['block']}丁目"
    
    logging.info(f"住所分割結果 - 基本住所: {base_address}, 番地: {address_parts['number']}")
    
    # 郵便番号のフォーマットチェック
    postal_code_clean = postal_code.replace("-", "")
    if len(postal_code_clean) != 7 or not postal_code_clean.isdigit():
        return {"status": "error", "message": "郵便番号は7桁の数字で入力してください。"}
    
    # 郵便番号を前半3桁と後半4桁に分割
    postal_code_first = postal_code_clean[:3]
    postal_code_second = postal_code_clean[3:]
    
    # ブラウザ設定を読み込む
    browser_settings = {}
    try:
        if os.path.exists("settings.json"):
            with open("settings.json", "r", encoding="utf-8") as f:
                settings = json.load(f)
                browser_settings = settings.get("browser_settings", {})
                logging.info(f"ブラウザ設定を読み込みました: {browser_settings}")
    except Exception as e:
        logging.warning(f"ブラウザ設定の読み込みに失敗しました: {str(e)}")
        browser_settings = {
            "headless": False,
            "show_popup": True,
            "auto_close": True,
            "page_load_timeout": 60,
            "script_timeout": 60
        }
    
    # ヘッドレスモードの設定を取得
    headless_mode = browser_settings.get("headless", False)
    # ポップアップ表示設定を取得
    show_popup = browser_settings.get("show_popup", True)
    # ブラウザ自動終了設定を取得
    auto_close = browser_settings.get("auto_close", True)
    # タイムアウト設定を取得
    page_load_timeout = browser_settings.get("page_load_timeout", 60)
    script_timeout = browser_settings.get("script_timeout", 60)
    
    logging.info(f"ブラウザ設定 - ヘッドレス: {headless_mode}, ポップアップ表示: {show_popup}, 自動終了: {auto_close}")
    
    driver = None
    try:
        # ブラウザを起動
        if progress_callback:
            progress_callback("ブラウザを起動中...")
        
        driver = create_driver(
            headless=headless_mode
        )
        
        # ブラウザ起動後のキャンセルチェック
        check_cancellation()
        
        # グローバル変数に保存
        global_driver = driver
         # タイムアウト設定を適用
        driver.set_page_load_timeout(page_load_timeout)
        driver.set_script_timeout(script_timeout)
        
        # タイムアウト設定後のキャンセルチェック
        check_cancellation()
        
        driver.implicitly_wait(0)  # 暗黙の待機を無効化
        
        # サイトにアクセス
        if progress_callback:
            progress_callback("サイトにアクセス中...")
        
        # サイトアクセス直前のキャンセルチェック
        check_cancellation()
        
        driver.get("https://flets.com/app_new/cao/")
        logging.info("サイトにアクセスしました")
        
        # サイトアクセス完了後のキャンセルチェック
        check_cancellation()
        

        
        # 郵便番号入力ページが表示されるのを待つ
        # 郵便番号入力フィールドが表示されるまで待機
        
        # ページ表示待機前のキャンセルチェック
        check_cancellation()
        
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "id_address_search_zip1"))
        )
        logging.info("郵便番号入力ページが表示されました")
        
        # ページ表示確認後のキャンセルチェック
        check_cancellation()
        

        
        # 郵便番号入力フィールドを探す
        try:
            # キャンセルチェック
            check_cancellation()
            
            # 郵便番号前半3桁を入力
            postal_code_first_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "id_address_search_zip1"))
            )
            postal_code_first_input.clear()
            postal_code_first_input.send_keys(postal_code_first)
            logging.info(f"郵便番号前半3桁を入力: {postal_code_first}")
            
            # 郵便番号後半4桁を入力
            postal_code_second_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "id_address_search_zip2"))
            )
            postal_code_second_input.clear()
            postal_code_second_input.send_keys(postal_code_second)
            logging.info(f"郵便番号後半4桁を入力: {postal_code_second}")
            
            # 再検索ボタンをクリック
            search_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "id_address_search_button"))
            )
            search_button.click()
            logging.info("再検索ボタンをクリックしました")
            
            # エラーメッセージの有無を確認
            try:
                
                # エラーメッセージを探す
                error_message = driver.find_elements(By.XPATH, "//*[contains(text(), '入力に誤りがあります')]")
                
                if error_message:
                    error_text = error_message[0].text
                    logging.error(f"郵便番号エラー: {error_text}")
                    return {"status": "error", "message": f"郵便番号エラー: {error_text}"}
            except Exception as e:
                logging.info(f"エラーメッセージの確認中に例外が発生しましたが、処理を続行します: {str(e)}")
            
        except Exception as e:
            logging.error(f"郵便番号入力処理中にエラー: {str(e)}")
            return {"status": "error", "message": f"郵便番号入力処理中にエラーが発生しました: {str(e)}"}
        
        # 住所候補が表示されるのを待つ
        try:
            if progress_callback:
                progress_callback("住所候補を検索中...")
            
            # キャンセルチェック
            check_cancellation()
            
            # 住所候補リストが表示されるまで待機
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "btn_list"))
            )
            logging.info("住所候補リストが表示されました")
            

            
            # 候補リストを取得
            candidates = driver.find_elements(By.CSS_SELECTOR, ".btn_list li.addressInfo")
            logging.info(f"{len(candidates)} 件の候補が見つかりました")
            
            # 候補の内容をログ出力
            for i, candidate in enumerate(candidates[:5]):  # 最初の5件のみログ出力
                logging.info(f"候補 {i+1}: {candidate.text.strip()}")
            
            # 有効な候補（テキストが空でない）をフィルタリング
            valid_candidates = [c for c in candidates if c.text.strip()]
            
            if not valid_candidates:
                raise NoSuchElementException("有効な住所候補が見つかりませんでした")
            
            logging.info(f"有効な候補数: {len(valid_candidates)}")
            
            # 住所を選択
            best_candidate, similarity = find_best_address_match(address, valid_candidates)
            
            if best_candidate:
                selected_address = best_candidate.text.strip()
                logging.info(f"選択された住所: '{selected_address}' (類似度: {similarity})")
                
                try:
                    # JavaScriptを使用してクリックを実行
                    driver.execute_script("arguments[0].click();", best_candidate)
                    logging.info("JavaScriptを使用して住所を選択しました")
                    
                    # クリック後の待機
                    time.sleep(2)
                    
                    # 番地入力画面への遷移を待機
                    WebDriverWait(driver, 15).until(
                        EC.url_contains("/cao/InputAddressNum")
                    )
                    logging.info("番地入力画面への遷移を確認しました")
                    
                    # ページの読み込み完了を待機
                    WebDriverWait(driver, 20).until(
                        lambda d: d.execute_script('return document.readyState') == 'complete'
                    )
                    logging.info("番地入力ページの読み込みが完了しました")
                    
                    if progress_callback:
                        progress_callback("番地を入力中...")
                    
                    # 番地入力画面の処理に進む
                    result = handle_address_number_input(driver, address_parts, progress_callback)
                    logging.info(f"番地入力処理の結果: {result}")
                    return result
                    
                except Exception as e:
                    logging.error(f"住所選択処理に失敗: {str(e)}")
                    driver.save_screenshot("debug_address_select_error.png")
                    raise
            else:
                logging.error(f"適切な住所候補が見つかりませんでした。入力住所: {address}")
                raise ValueError("適切な住所候補が見つかりませんでした")
            
        except Exception as e:
            logging.error(f"住所選択処理中にエラー: {str(e)}")
            return {"status": "error", "message": f"住所選択処理中にエラーが発生しました: {str(e)}"}
            
    except CancellationError as e:
        logging.info("★★★ 提供エリア検索がキャンセルされました ★★★")
        return {"status": "cancelled", "message": "検索がキャンセルされました"}
    except TimeoutException as e:
        logging.error(f"タイムアウトエラー: {str(e)}")
        return {"status": "failure", "message": "処理がタイムアウトしました"}
    except Exception as e:
        logging.error(f"検索処理中にエラー: {str(e)}")
        return {"status": "error", "message": f"検索処理中にエラーが発生しました: {str(e)}"}
        
    finally:
        # ブラウザを終了
        if driver and auto_close:
            try:
                driver.quit()
                logging.info("ブラウザを終了しました")
            except Exception as e:
                logging.warning(f"ブラウザの終了中にエラー: {str(e)}") 

def find_input_element(driver, attempt_count=0):
    """
    番地入力フォームを見つけるためのヘルパー関数
    
    Args:
        driver: WebDriverインスタンス
        attempt_count: 試行回数
        
    Returns:
        element: 見つかった要素、見つからない場合はNone
    """
    try:
        # iframeの確認
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            for iframe in iframes:
                try:
                    driver.switch_to.frame(iframe)
                    element = driver.find_element(By.NAME, "banchi1to3manualAddressNum1")
                    if element.is_displayed():
                        return element
                except:
                    pass
                finally:
                    driver.switch_to.default_content()
        
        # 複数の方法で要素を探す
        selectors = [
            (By.NAME, "banchi1to3manualAddressNum1"),
            (By.ID, "id_banchi1to3manualAddressNum1"),
            (By.CSS_SELECTOR, "input[name='banchi1to3manualAddressNum1']"),
            (By.CSS_SELECTOR, "input[type='tel'][name='banchi1to3manualAddressNum1']"),
            (By.XPATH, "//input[@name='banchi1to3manualAddressNum1']"),
            (By.XPATH, "//div[contains(@class, '_input')]//input[1]")
        ]
        
        for by, selector in selectors:
            try:
                element = driver.find_element(by, selector)
                if element.is_displayed():
                    return element
            except:
                continue
        
        return None
    except Exception as e:
        logging.warning(f"要素検索中にエラー: {str(e)}")
        return None

def debug_page_state(driver, context=""):
    """
    ページの状態をデバッグ出力する
    
    Args:
        driver: WebDriverインスタンス
        context: デバッグコンテキストの説明
    """
    try:
        logging.info(f"=== デバッグ情報（{context}）===")
        logging.info(f"現在のURL: {driver.current_url}")
        logging.info(f"ページタイトル: {driver.title}")
        
        # ページの読み込み状態
        ready_state = driver.execute_script('return document.readyState')
        logging.info(f"ページの読み込み状態: {ready_state}")
        
        # DOMの準備状態
        is_dom_loaded = driver.execute_script('return document.readyState === "complete" || document.readyState === "interactive"')
        logging.info(f"DOMの準備完了: {is_dom_loaded}")
        
        # iframeの確認
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        logging.info(f"iframe数: {len(iframes)}")
        
        # エラーメッセージの確認
        error_elements = driver.find_elements(By.CLASS_NAME, "error")
        if error_elements:
            logging.info("エラーメッセージが見つかりました:")
            for error in error_elements:
                logging.info(f"エラー: {error.text}")
        
    except Exception as e:
        logging.error(f"デバッグ情報の取得中にエラー: {str(e)}")

def handle_address_number_input(driver, address_parts, progress_callback=None):
    """
    番地入力画面の処理を行う
    
    Args:
        driver: WebDriverインスタンス
        address_parts: 分割された住所情報
        progress_callback: 進捗コールバック関数
    """
    try:
        # ブラウザ設定を読み込む（show_popup用）
        show_popup = True  # デフォルト値
        try:
            if os.path.exists("settings.json"):
                with open("settings.json", "r", encoding="utf-8") as f:
                    settings = json.load(f)
                    browser_settings = settings.get("browser_settings", {})
                    show_popup = browser_settings.get("show_popup", True)
                    logging.info(f"ポップアップ表示設定を読み込みました: {show_popup}")
        except Exception as e:
            logging.warning(f"ブラウザ設定の読み込みに失敗しました: {str(e)}")

        logging.info("=== 番地入力画面の処理開始 ===")
        debug_page_state(driver, "番地入力画面_初期状態")

        # ページ読み込み完了まで待機
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "id_form_main"))
        )
        debug_page_state(driver, "番地入力画面_読み込み完了後")
        


        # 番地がない場合のチェックボックス処理
        if not address_parts.get('number'):
            logging.info("番地が指定されていないため、「番地・号が無い」をチェック")
            try:
                checkbox = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "id_banchi1to3Fixed"))
                )
                if not checkbox.is_selected():
                    driver.execute_script("arguments[0].click();", checkbox)
                logging.info("「番地・号が無い」をチェックしました")
                

            except Exception as e:
                logging.error(f"「番地・号が無い」チェックボックスの操作に失敗: {str(e)}")
                debug_page_state(driver, "チェックボックス操作_失敗")
                raise
        else:
            # 番地入力フィールドの処理
            number_parts = address_parts['number'].split('-')
            logging.info(f"入力する番地: {number_parts}")

            try:
                # 番地1の入力
                number1_input = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "id_banchi1to3manualAddressNum1"))
                )
                # フォーカスを設定してからクリア
                driver.execute_script("arguments[0].focus();", number1_input)
                number1_input.clear()
                number1_input.send_keys(number_parts[0])
                logging.info(f"番地1を入力: {number_parts[0]}")

                # 番地2の入力（存在する場合）
                if len(number_parts) > 1:
                    number2_input = driver.find_element(By.ID, "id_banchi1to3manualAddressNum2")
                    driver.execute_script("arguments[0].focus();", number2_input)
                    number2_input.clear()
                    number2_input.send_keys(number_parts[1])
                    logging.info(f"番地2を入力: {number_parts[1]}")

                # 番地3の入力（存在する場合）
                if len(number_parts) > 2:
                    number3_input = driver.find_element(By.ID, "id_banchi1to3manualAddressNum3")
                    driver.execute_script("arguments[0].focus();", number3_input)
                    number3_input.clear()
                    number3_input.send_keys(number_parts[2])
                    logging.info(f"番地3を入力: {number_parts[2]}")
                
      

            except Exception as e:
                logging.error(f"番地入力に失敗: {str(e)}")
                debug_page_state(driver, "番地入力_失敗")
                raise

        # 住居タイプの選択（デフォルトで戸建てを選択）
        try:
            house_type = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "id_buildType_1"))
            )
            if not house_type.is_selected():
                driver.execute_script("arguments[0].click();", house_type)
            logging.info("住居タイプ: 戸建てを選択")
            


            # 次へボタンをクリック
            next_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "id_nextButton"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
            time.sleep(1)  # スクロール完了を待つ
            driver.execute_script("arguments[0].click();", next_button)
            logging.info("次へボタンをクリックしました")

            # 建物選択画面が表示されたかチェック
            try:
                # 建物選択画面のURLを確認
                WebDriverWait(driver, 10).until(
                    lambda d: "SelectBuild1" in d.current_url
                )
                logging.info("建物選択画面が表示されました")

                # 建物選択画面が表示された時点で集合住宅と判定
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                screenshot_path = f"debug_apartment_confirmation_{timestamp}.png"
                driver.save_screenshot(screenshot_path)
                return {
                    "status": "apartment",
                    "message": "集合住宅",
                    "details": {
                        "判定結果": "NG",
                        "提供エリア": "集合住宅",
                        "備考": "集合住宅のため、判定を終了します"
                    },
                    "screenshot": screenshot_path,
                    "show_popup": show_popup
                }

            except TimeoutException:
                # 建物選択画面が表示されない場合は通常の結果ページへの遷移を待機
                logging.info("建物選択画面はスキップされました")
                WebDriverWait(driver, 10).until(
                    EC.url_contains("ProvideResult")
                )
                logging.info("結果ページへ遷移しました")

        except Exception as e:
            logging.error(f"住居タイプの選択または次へボタンのクリックに失敗: {str(e)}")
            debug_page_state(driver, "住居タイプ選択_次へボタン_失敗")
            raise

        # 結果ページへの遷移を待機
        try:
            WebDriverWait(driver, 10).until(
                EC.url_contains("ProvideResult")
            )
            logging.info("結果ページへ遷移しました")
            
            
            debug_page_state(driver, "結果ページ_表示")

            # 結果テキストの取得を修正
            try:
                # まず、ローディング表示が消えるのを待つ
                WebDriverWait(driver, 10).until_not(
                    EC.presence_of_element_located((By.CLASS_NAME, "loading"))
                )
                
                # 結果テキストを取得（複数の方法で試行）
                result_text = None
                
                # 結果テキストの取得方法を追加
                selectors = [
                    (By.XPATH, "//div[contains(@class, 'main_wrap')]//h1/following-sibling::div"),
                    (By.CLASS_NAME, "resultText"),
                    (By.XPATH, "//div[contains(text(), 'フレッツ光') and contains(text(), 'エリア')]"),
                    (By.XPATH, "//div[contains(@class, 'main_wrap')]//div[contains(text(), 'エリア')]"),
                    (By.XPATH, "//h1[contains(text(), 'エリア')]"),
                    (By.XPATH, "//div[contains(@class, 'result')]"),
                    (By.XPATH, "//div[contains(@class, 'main_wrap')]//div[not(@class)]")
                ]

                for selector_type, selector in selectors:
                    try:
                        element = WebDriverWait(driver, 3).until(
                            EC.presence_of_element_located((selector_type, selector))
                        )
                        if element and element.is_displayed():
                            result_text = element.text.strip()
                            logging.info(f"結果テキストを取得: {result_text} (セレクター: {selector})")
                            if result_text:
                                break
                    except Exception as e:
                        logging.debug(f"セレクター {selector} での検索に失敗: {str(e)}")
                        continue

                # 結果が見つからない場合、ページ全体のテキストから判定
                if not result_text:
                    logging.info("個別の要素での結果テキスト取得に失敗。ページ全体から検索を試みます。")
                    page_text = driver.find_element(By.TAG_NAME, "body").text
                    if "提供エリアです" in page_text:
                        result_text = "提供エリアです"
                        logging.info("ページテキストから「提供エリアです」を検出")
                    elif "提供エリア外です" in page_text:
                        result_text = "提供エリア外です"
                        logging.info("ページテキストから「提供エリア外です」を検出")

                logging.info(f"最終的な結果テキスト: {result_text}")

                # スクリーンショットを保存（結果確認時のみ）
                timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
                
                if result_text:
                    if "提供エリアです" in result_text or "の提供エリアです" in result_text:
                        screenshot_path = f"debug_available_confirmation_{timestamp}.png"
                        driver.save_screenshot(screenshot_path)
                        return {
                            "status": "available",
                            "message": "提供可能",
                            "details": {
                                "判定結果": "OK",
                                "提供エリア": "提供可能エリアです",
                                "備考": "光アクセスのサービスがご利用いただけます"
                            },
                            "screenshot": screenshot_path,
                            "show_popup": show_popup
                        }
                    elif "提供エリア外です" in result_text or "エリア外" in result_text:
                        screenshot_path = f"debug_not_provided_confirmation_{timestamp}.png"
                        driver.save_screenshot(screenshot_path)
                        return {
                            "status": "unavailable",
                            "message": "未提供",
                            "details": {
                                "判定結果": "NG",
                                "提供エリア": "提供対象外エリアです",
                                "備考": "申し訳ございませんが、このエリアではサービスを提供しておりません"
                            },
                            "screenshot": screenshot_path,
                            "show_popup": show_popup
                        }
                    else:
                        screenshot_path = f"debug_investigation_confirmation_{timestamp}.png"
                        driver.save_screenshot(screenshot_path)
                        logging.warning(f"予期しない結果テキスト: {result_text}")
                        return {
                            "status": "failure",
                            "message": "判定失敗",
                            "details": {
                                "判定結果": "判定失敗",
                                "提供エリア": "調査が必要なエリアです",
                                "備考": "住所を特定できないため、担当者がお調べします"
                            },
                            "screenshot": screenshot_path,
                            "show_popup": show_popup
                        }
                else:

                    
                    screenshot_path = f"debug_error_confirmation_{timestamp}.png"
                    driver.save_screenshot(screenshot_path)
                    logging.error("結果テキストが取得できませんでした")
                    return {
                        "status": "failure",
                        "message": "判定失敗",
                        "details": {
                            "判定結果": "判定失敗",
                            "提供エリア": "判定できませんでした",
                            "備考": "結果テキストが取得できませんでした"
                        },
                        "screenshot": screenshot_path,
                        "show_popup": show_popup
                    }

            except Exception as e:
                screenshot_path = f"debug_error_confirmation_{timestamp}.png"
                driver.save_screenshot(screenshot_path)
                logging.error(f"結果テキストの取得中にエラー: {str(e)}")
                return {
                    "status": "failure",
                    "message": "判定失敗",
                    "details": {
                        "判定結果": "判定失敗",
                        "提供エリア": "判定できませんでした",
                        "備考": f"結果の判定に失敗しました: {str(e)}"
                    },
                    "screenshot": screenshot_path,
                    "show_popup": show_popup
                }

        except Exception as e:
            logging.error(f"結果の取得に失敗: {str(e)}")
            debug_page_state(driver, "結果取得_失敗")
            return {"status": "error", "message": "結果の取得に失敗しました"}

    except Exception as e:
        logging.error(f"番地入力画面の処理中にエラー: {str(e)}")
        debug_page_state(driver, "エラー発生時の状態")
        raise 