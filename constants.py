"""
定数定義モジュール
検査水準と計算関連の定数を定義
"""

import re
from datetime import datetime


class InspectionConstants:
    """検査水準と計算関連の定数定義"""
    
    # 検査水準の閾値（不良率 %）
    DEFECT_RATE_THRESHOLD_NORMAL = 0.5  # 普通水準の閾値
    DEFECT_RATE_THRESHOLD_STRICT = 0.5  # きつい水準の閾値
    
    # 検査水準の定義
    INSPECTION_LEVELS = {
        'loose': {
            'threshold': 0, 
            'name': 'ゆるい(I)', 
            'description': '過去の不具合が0件だったため、最もゆるい水準（I）を適用しています。'
        },
        'normal': {
            'threshold': 0.5, 
            'name': '普通(II)', 
            'description': '過去の不具合率が0.5%以下だったため、普通水準（II）を適用しています。'
        },
        'strict': {
            'threshold': float('inf'), 
            'name': 'きつい(III)', 
            'description': '過去の不具合率が0.5%を超えていたため、きつい水準（III）を適用しています。'
        }
    }
    
    # 計算制限値
    MAX_LOT_SIZE = 1000000  # 最大ロットサイズ
    MIN_LOT_SIZE = 1        # 最小ロットサイズ
    MAX_CONFIDENCE = 99.9   # 最大信頼度
    MIN_CONFIDENCE = 50.0   # 最小信頼度
    MAX_C_VALUE = 100       # 最大c値
    MIN_C_VALUE = 0         # 最小c値
    
    # サンプルサイズ計算用定数
    MAX_SAMPLE_SIZE_MULTIPLIER = 0.1  # ロットサイズに対する最大サンプルサイズの倍率
    MAX_CALCULATION_LIMIT = 10000     # 計算の最大制限値
    
    # デフォルト値
    DEFAULT_CONFIDENCE = 95.0  # デフォルト信頼度
    DEFAULT_C_VALUE = 0        # デフォルトc値
    
    # 品番関連
    MAX_PRODUCT_NUMBER_LENGTH = 50  # 品番の最大長
    MIN_RECOMMENDED_LOT_SIZE = 10   # 推奨最小ロットサイズ
    
    # 入力値検証用の正規表現
    PRODUCT_NUMBER_PATTERN = r'^[A-Za-z0-9\-_]+$'  # 品番の形式
    INVALID_CHARS = ['<', '>', ':', '"', '|', '?', '*', '\\', '/']  # 無効な文字
    
    @staticmethod
    def get_inspection_level(defect_rate):
        """不具合率に基づいて検査水準を取得"""
        if defect_rate == 0:
            return InspectionConstants.INSPECTION_LEVELS['loose']
        elif defect_rate <= InspectionConstants.DEFECT_RATE_THRESHOLD_NORMAL:
            return InspectionConstants.INSPECTION_LEVELS['normal']
        else:
            return InspectionConstants.INSPECTION_LEVELS['strict']


# 不具合項目の定義
DEFECT_COLUMNS = [
    "外観キズ", "圧痕", "切粉", "毟れ", "穴大", "穴小", "穴キズ", "バリ", "短寸", "面粗", "サビ", "ボケ", "挽目", "汚れ", "メッキ", "落下",
    "フクレ", "ツブレ", "ボッチ", "段差", "バレル石", "径プラス", "径マイナス", "ゲージ", "異物混入", "形状不良", "こすれ", "変色シミ", "材料キズ", "ゴミ", "その他"
]