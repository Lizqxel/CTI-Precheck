"""
文字列ユーティリティモジュール

このモジュールは、文字列の正規化や類似度計算などの
ユーティリティ関数を提供します。
"""

import re


def normalize_string(text):
    """
    文字列を正規化する
    
    Args:
        text (str): 正規化する文字列
        
    Returns:
        str: 正規化された文字列
    """
    if not text:
        return text
    
    # 全角スペースを半角に変換
    normalized = text.replace('　', ' ')
    
    # 都道府県名を一時的に保存
    prefecture_match = re.match(r'^(.+?[都道府県])', normalized)
    if prefecture_match:
        prefecture = prefecture_match.group(1)
        remaining = normalized[len(prefecture):]
    else:
        prefecture = ""
        remaining = normalized
    
    # 「大字」「字」を削除
    remaining = remaining.replace('大字', '').replace('字', '')
    
    # 全角数字を半角に変換
    zen_to_han = str.maketrans('０１２３４５６７８９', '0123456789')
    remaining = remaining.translate(zen_to_han)
    
    # 漢数字を半角数字に変換（都道府県名は除外）
    kanji_to_number = {
        '一': '1', '二': '2', '三': '3', '四': '4', '五': '5',
        '六': '6', '七': '7', '八': '8', '九': '9', '十': '10',
        '壱': '1', '弐': '2', '参': '3', '肆': '4', '伍': '5',
        '陸': '6', '漆': '7', '捌': '8', '玖': '9', '拾': '10'
    }
    
    # 漢数字を数字に変換（都道府県名は除外）
    for kanji, number in kanji_to_number.items():
        remaining = remaining.replace(kanji, number)
    
    # 全角ハイフンを半角に変換
    remaining = remaining.replace('−', '-').replace('ー', '-').replace('－', '-')
    
    # 結果を結合
    normalized = prefecture + remaining
    
    # 余分な空白を削除
    normalized = normalized.strip()
    
    return normalized


def calculate_similarity(str1, str2):
    """
    2つの文字列の類似度を計算する
    
    Args:
        str1 (str): 比較する文字列1
        str2 (str): 比較する文字列2
        
    Returns:
        float: 類似度（0.0～1.0）
    """
    if not str1 or not str2:
        return 0.0
    
    # 文字列を正規化（全角→半角、大文字→小文字）
    str1 = str1.lower()
    str2 = str2.lower()
    
    # レーベンシュタイン距離を計算
    len1 = len(str1)
    len2 = len(str2)
    
    # 行列を初期化
    matrix = [[0 for _ in range(len2 + 1)] for _ in range(len1 + 1)]
    
    # 行列の初期値を設定
    for i in range(len1 + 1):
        matrix[i][0] = i
    for j in range(len2 + 1):
        matrix[0][j] = j
    
    # レーベンシュタイン距離を計算
    for i in range(1, len1 + 1):
        for j in range(1, len2 + 1):
            cost = 0 if str1[i-1] == str2[j-1] else 1
            matrix[i][j] = min(
                matrix[i-1][j] + 1,      # 削除
                matrix[i][j-1] + 1,      # 挿入
                matrix[i-1][j-1] + cost  # 置換
            )
    
    # 類似度を計算（0.0～1.0）
    max_len = max(len1, len2)
    if max_len == 0:
        return 1.0
    
    distance = matrix[len1][len2]
    similarity = 1.0 - (distance / max_len)
    
    return similarity


def validate_name(text):
    """
    名前が有効かどうかを検証する関数
    
    Args:
        text (str): 検証する名前
        
    Returns:
        bool: 名前が有効な場合はTrue、無効な場合はFalse
    """
    if not text:
        return True  # 空文字列は許可
        
    # 数字を含む場合は無効
    if re.search(r'\d', text):
        return False
        
    return True


def validate_furigana(text):
    """
    フリガナが有効かどうかを検証する関数
    
    Args:
        text (str): 検証するフリガナ
        
    Returns:
        bool: フリガナが有効な場合はTrue、無効な場合はFalse
    """
    if not text:
        return True  # 空文字列は許可
        
    # 数字を含む場合は無効
    if re.search(r'\d', text):
        return False
        
    # カタカナ、ひらがな、スペース以外の文字を含む場合は無効
    if re.search(r'[^\u3040-\u309F\u30A0-\u30FF\s]', text):
        return False
        
    return True


def convert_to_half_width_except_space(text):
    """
    全角文字を半角に変換するが、スペースだけは全角のままにする関数
    
    Args:
        text (str): 変換する文字列
        
    Returns:
        str: 変換された文字列（スペースは全角のまま）
    """
    if not text:
        return ""
        
    # 全角数字を半角に変換
    text = text.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    
    # 全角英字を半角に変換
    text = text.translate(str.maketrans('ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ',
                                      'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'))
    
    # 全角記号を半角に変換（一対一対応するよう注意）
    text = text.translate(str.maketrans('（）［］｛｝「」『』【】〔〕！＂＃＄％＆＇＊＋，．／：；＜＝＞？＠＼＾＿｀｜～',
                                      '()[]{}""\'\'<>[]!"#$%&\'*+,./:;<=>?@\\^_`|~'))
    
    # すべての種類のハイフン、ダッシュを半角ハイフンに変換
    text = text.replace('−', '-').replace('ー', '-').replace('－', '-')
    text = text.replace('―', '-').replace('‐', '-').replace('‑', '-')
    text = text.replace('‒', '-').replace('–', '-').replace('—', '-')
    text = text.replace('﹘', '-').replace('⁃', '-').replace('⎯', '-')
    text = text.replace('⏤', '-').replace('─', '-').replace('━', '-')
    
    # スペースは変換しない（全角スペースはそのまま）
    
    return text 