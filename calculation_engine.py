"""
統計計算エンジンモジュール
SQC統計計算とデータ処理を管理
"""

import math
import pyodbc
from scipy.stats import binom
from constants import InspectionConstants, DEFECT_COLUMNS


class CalculationEngine:
    """統計計算エンジンクラス"""
    
    def __init__(self, db_manager):
        self.db_manager = db_manager
    
    def build_sql_query(self, base_sql, inputs):
        """SQLクエリの構築"""
        sql_parts = [base_sql]
        params = [inputs['product_number']]
        has_where = ' where ' in base_sql.lower()
        
        if inputs['start_date']:
            sql_parts.append(f"{ 'AND' if has_where else 'WHERE'} [指示日] >= ?")
            params.append(inputs['start_date'])
            has_where = True
            
        if inputs['end_date']:
            sql_parts.append(f"{ 'AND' if has_where else 'WHERE'} [指示日] <= ?")
            params.append(inputs['end_date'])
            
        return " ".join(sql_parts), params

    def fetch_data(self, cursor, inputs):
        """データの取得"""
        data = {'total_qty': 0, 'total_defect': 0, 'defect_rate': 0, 'defect_rates_sorted': [], 'best5': []}
        defect_columns_sum = ", ".join(f"SUM(IIF([{col}] IS NOT NULL AND [{col}]<>0, [{col}], 0))" for col in DEFECT_COLUMNS)
        base_sql = f"SELECT SUM([数量]), SUM([総不具合数]), {defect_columns_sum} FROM t_不具合情報 WHERE [品番] = ?"
        sql, params = self.build_sql_query(base_sql, inputs)
        row = cursor.execute(sql, *params).fetchone()
        
        if not row or row[0] is None: 
            return data
            
        total_qty, total_defect = row[0] or 0, row[1] or 0
        data['total_qty'] = total_qty
        data['total_defect'] = total_defect
        data['defect_rate'] = (total_defect / total_qty * 100) if total_qty > 0 else 0
        
        defect_counts = row[2:]
        defect_rates = []
        if total_qty > 0 and defect_counts:
            for col, count in zip(DEFECT_COLUMNS, defect_counts):
                count = count or 0
                if count > 0:
                    rate = (count / total_qty * 100)
                    defect_rates.append((col, rate, count))
        
        defect_rates.sort(key=lambda x: x[2], reverse=True)
        data['defect_rates_sorted'] = defect_rates
        data['best5'] = [(col, count) for col, rate, count in defect_rates[:5]]
        return data

    def calculate_stats(self, db_data, inputs):
        """統計計算"""
        results = {}
        p = db_data['defect_rate'] / 100
        
        # 検査水準の判定（定数を使用）
        defect_rate = db_data['defect_rate']
        if defect_rate == 0:
            level_info = InspectionConstants.INSPECTION_LEVELS['loose']
        elif defect_rate <= InspectionConstants.DEFECT_RATE_THRESHOLD_NORMAL:
            level_info = InspectionConstants.INSPECTION_LEVELS['normal']
        else:
            level_info = InspectionConstants.INSPECTION_LEVELS['strict']
        
        results['level_text'] = level_info['name']
        results['level_reason'] = level_info['description']
        
        # 抜取検査数の計算
        n_sample = "計算不可"
        warning_message = None
        
        if p > 0 and 0 < inputs['confidence_level']/100 < 1:
            try:
                if inputs['c_value'] == 0:
                    # c=0の場合の計算
                    theoretical_n = math.ceil(math.log(1 - inputs['confidence_level']/100) / math.log(1 - p))
                    
                    # ロットサイズとの比較
                    if theoretical_n > inputs['lot_size']:
                        n_sample = f"全数検査必要（理論値: {theoretical_n:,}個）"
                        warning_message = f"設定条件では理論上{theoretical_n:,}個の抜取が必要ですが、ロットサイズ（{inputs['lot_size']:,}個）を超えています。全数検査を推奨します。"
                    else:
                        n_sample = theoretical_n
                else:
                    # c>0の場合の二分探索
                    low, high = 1, inputs['lot_size']  # ロットサイズを上限に設定
                    n_sample = f"全数検査必要（計算断念）"
                    
                    while low <= high:
                        mid = (low + high) // 2
                        if mid == 0: 
                            low = 1
                            continue
                        if binom.cdf(inputs['c_value'], mid, p) >= 1 - inputs['confidence_level']/100:
                            n_sample, high = mid, mid - 1
                        else:
                            low = mid + 1
                    
                    # c>0でロットサイズを超える場合の警告
                    if n_sample == f"全数検査必要（計算断念）":
                        warning_message = f"c={inputs['c_value']}、信頼度{inputs['confidence_level']:.1f}%の条件では、ロットサイズ（{inputs['lot_size']:,}個）を超える抜取が必要です。全数検査を推奨します。"
                        
            except (ValueError, OverflowError): 
                n_sample = "計算エラー"
        elif p == 0:
            n_sample = 1
        
        # 警告メッセージを結果に追加
        if warning_message:
            results['warning_message'] = warning_message
        
        results['sample_size'] = n_sample
        return results

    def calculate_alternatives(self, db_data, inputs):
        """代替案の計算"""
        p = db_data['defect_rate'] / 100
        lot_size = inputs['lot_size']
        
        alternatives = "【代替案の提案】\n\n"
        
        # 案1: 信頼度を下げる
        alternatives += "1. 信頼度を下げる場合:\n"
        for conf in [95, 90, 85]:
            if p > 0:
                theoretical_n = math.ceil(math.log(1 - conf/100) / math.log(1 - p))
                if theoretical_n <= lot_size:
                    alternatives += f"   信頼度{conf}%: {theoretical_n:,}個\n"
                else:
                    alternatives += f"   信頼度{conf}%: 全数検査必要（理論値: {theoretical_n:,}個）\n"
        alternatives += "\n"
        
        # 案2: c値を上げる
        alternatives += "2. c値を上げる場合:\n"
        for c_val in [1, 2, 3]:
            try:
                low, high = 1, lot_size
                n_sample = "全数検査必要"
                
                while low <= high:
                    mid = (low + high) // 2
                    if mid == 0:
                        low = 1
                        continue
                    if binom.cdf(c_val, mid, p) >= 1 - inputs['confidence_level']/100:
                        n_sample, high = mid, mid - 1
                    else:
                        low = mid + 1
                
                if isinstance(n_sample, int):
                    alternatives += f"   c={c_val}: {n_sample:,}個\n"
                else:
                    alternatives += f"   c={c_val}: {n_sample}\n"
            except:
                alternatives += f"   c={c_val}: 計算エラー\n"
        alternatives += "\n"
        
        # 案3: 組み合わせ
        alternatives += "3. 信頼度とc値を組み合わせる場合:\n"
        for conf in [95, 90]:
            for c_val in [1, 2]:
                try:
                    if p > 0:
                        if c_val == 0:
                            theoretical_n = math.ceil(math.log(1 - conf/100) / math.log(1 - p))
                            if theoretical_n <= lot_size:
                                alternatives += f"   信頼度{conf}%、c={c_val}: {theoretical_n:,}個\n"
                            else:
                                alternatives += f"   信頼度{conf}%、c={c_val}: 全数検査必要\n"
                        else:
                            low, high = 1, lot_size
                            n_sample = "全数検査必要"
                            
                            while low <= high:
                                mid = (low + high) // 2
                                if mid == 0:
                                    low = 1
                                    continue
                                if binom.cdf(c_val, mid, p) >= 1 - conf/100:
                                    n_sample, high = mid, mid - 1
                                else:
                                    low = mid + 1
                            
                            if isinstance(n_sample, int):
                                alternatives += f"   信頼度{conf}%、c={c_val}: {n_sample:,}個\n"
                            else:
                                alternatives += f"   信頼度{conf}%、c={c_val}: {n_sample}\n"
                except:
                    alternatives += f"   信頼度{conf}%、c={c_val}: 計算エラー\n"
        alternatives += "\n"
        
        # 推奨案
        alternatives += "【推奨案】\n"
        alternatives += "現在の条件では統計的に適切な抜取検査が困難です。\n"
        alternatives += "以下のいずれかを検討してください:\n\n"
        alternatives += "• 全数検査の実施\n"
        alternatives += "• 信頼度を95%に下げる\n"
        alternatives += "• c値を1以上に設定する\n"
        alternatives += "• 不良率の仮定を見直す\n\n"
        alternatives += "※ 品質要求に応じて最適な条件を選択してください。"
        
        return alternatives

