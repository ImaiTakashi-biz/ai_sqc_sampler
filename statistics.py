"""
統計計算モジュール
SQC関連の統計計算を管理
"""

import math
from scipy.stats import binom
from constants import InspectionConstants


class SQCStatistics:
    """統計的品質管理の統計計算クラス"""
    
    @staticmethod
    def calculate_defect_rate(defect_data):
        """不具合率の計算"""
        if not defect_data:
            return 0.0
        
        total_defects = 0
        total_quantity = 0
        
        for row in defect_data:
            try:
                # 不具合数（インデックス13）
                defect_count = float(row[13]) if row[13] is not None else 0
                # 総数（インデックス12）
                total_count = float(row[12]) if row[12] is not None else 0
                
                total_defects += defect_count
                total_quantity += total_count
                
            except (ValueError, TypeError, IndexError):
                continue
        
        if total_quantity == 0:
            return 0.0
        
        return (total_defects / total_quantity) * 100
    
    @staticmethod
    def calculate_sample_size(lot_size, confidence_level, c_value, defect_rate=0.0):
        """サンプルサイズを計算（バックアップの計算方法を復元）"""
        try:
            # 型変換を確実に行う
            lot_size = int(lot_size)
            confidence_level = float(confidence_level)
            c_value = int(c_value)
            defect_rate = float(defect_rate)
            
            # 信頼度を小数に変換
            confidence = confidence_level / 100.0
            
            # 不具合率を小数に変換
            p = defect_rate / 100.0
            
            # 抜取検査数の計算（バックアップのロジック）
            n_sample = "計算不可"
            if p > 0 and 0 < confidence < 1:
                try:
                    if c_value == 0:
                        # c=0の場合の計算
                        n_sample = math.ceil(math.log(1 - confidence) / math.log(1 - p))
                    else:
                        # c>0の場合の二分探索
                        low, high = 1, max(
                            lot_size * InspectionConstants.MAX_SAMPLE_SIZE_MULTIPLIER, 
                            InspectionConstants.MAX_CALCULATION_LIMIT
                        )
                        n_sample = f">{high} (計算断念)"
                        while low <= high:
                            mid = (low + high) // 2
                            if mid == 0: 
                                low = 1
                                continue
                            if binom.cdf(c_value, mid, p) >= 1 - confidence:
                                n_sample, high = mid, mid - 1
                            else:
                                low = mid + 1
                except (ValueError, OverflowError): 
                    n_sample = "計算エラー"
            elif p == 0:
                # 不具合率が0%の場合
                n_sample = 1
            
            return n_sample
            
        except Exception as e:
            raise ValueError(f"サンプルサイズの計算中にエラーが発生しました: {e}")
    