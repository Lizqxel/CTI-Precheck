"""
住所関連のユーティリティ関数

このモジュールは、住所の正規化や分割などの共通機能を提供します。
"""

import re
import logging

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

def split_address(address):
    """
    住所文字列を都道府県、市区町村、町名、番地に分割する
    
    Args:
        address (str): 分割する住所文字列
        
    Returns:
        dict: 分割された住所情報
    """
    try:
        # 住所を正規化
        address = normalize_address(address)
        
        # 都道府県を抽出
        prefecture_pattern = r'^(東京都|北海道|(?:京都|大阪)府|.+?県)'
        prefecture_match = re.match(prefecture_pattern, address)
        if not prefecture_match:
            raise ValueError("都道府県が見つかりません")
        prefecture = prefecture_match.group(1)
        
        # 都道府県を除去
        remaining = address[len(prefecture):]
        
        # 市区町村を抽出
        city_pattern = r'^(.+?[市区町村])'
        city_match = re.match(city_pattern, remaining)
        if not city_match:
            raise ValueError("市区町村が見つかりません")
        city = city_match.group(1)
        
        # 市区町村を除去
        remaining = remaining[len(city):]
        
        # 町名を抽出（丁目まで含む）
        town_pattern = r'^(.+?(?:丁目)?)'
        town_match = re.match(town_pattern, remaining)
        if not town_match:
            raise ValueError("町名が見つかりません")
        town = town_match.group(1)
        
        # 町名を除去
        remaining = remaining[len(town):]
        
        # 番地と号を抽出
        number_pattern = r'^(\d+(?:-\d+)*)'
        number_match = re.match(number_pattern, remaining)
        number = number_match.group(1) if number_match else ""
        
        # 丁目を抽出
        block_pattern = r'(\d+)丁目'
        block_match = re.search(block_pattern, town)
        block = block_match.group(1) if block_match else ""
        
        # 丁目を除去
        if block:
            town = re.sub(r'\d+丁目', '', town)
        
        return {
            'prefecture': prefecture,
            'city': city,
            'town': town,
            'block': block,
            'number': number,
            'building_id': None
        }
    except Exception as e:
        logging.error(f"住所の分割中にエラー: {str(e)}")
        return None 