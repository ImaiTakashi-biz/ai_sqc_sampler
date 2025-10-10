"""
統計計算エンジンモジュール
SQC統計計算とデータ処理を管理
AQL/LTPD設計に基づく抜取検査計算
"""

import math
from functools import lru_cache

from constants import InspectionConstants, DEFECT_COLUMNS

SAMPLE_SIZE_CACHE_LIMIT = 128
_DEFECT_SUM_SQL = ", ".join(
    f"SUM(IIF([{col}] IS NOT NULL AND [{col}]<>0, [{col}], 0))" for col in DEFECT_COLUMNS
)
_BASE_DEFECT_AGGREGATE_SQL = (
    f"SELECT SUM([数量]), SUM([総不具合数]), {_DEFECT_SUM_SQL} FROM t_不具合情報 WHERE [品番] = ?"
)


_BINOM = None
_HYPERGEOM = None


def _ensure_scipy_distributions():
    global _BINOM, _HYPERGEOM
    if _BINOM is None or _HYPERGEOM is None:
        from scipy.stats import binom as sp_binom, hypergeom as sp_hypergeom
        _BINOM = sp_binom
        _HYPERGEOM = sp_hypergeom
    return _BINOM, _HYPERGEOM


@lru_cache(maxsize=512)
def _cached_binom_cdf(c_value, sample_size, defect_rate):
    binom, _ = _ensure_scipy_distributions()
    return binom.cdf(c_value, sample_size, defect_rate)


@lru_cache(maxsize=512)
def _cached_hypergeom_cdf(c_value, population_size, defect_count, sample_size):
    _, hypergeom = _ensure_scipy_distributions()
    return hypergeom.cdf(c_value, population_size, defect_count, sample_size)


class CalculationEngine:
    """統計計算エンジンクラス"""
    
    def __init__(self, db_manager):
        self.db_manager = db_manager
        self._sample_size_cache = {}
    
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
        sql, params = self.build_sql_query(_BASE_DEFECT_AGGREGATE_SQL, inputs)
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
        """AQL/LTPD設計に基づく統計計算（データベース実績活用版）"""
        results = {}
        
        # AQL/LTPD設計のパラメータ取得
        aql = inputs.get('aql', 0.25)  # デフォルト0.25%
        ltpd = inputs.get('ltpd', 1.0)  # デフォルト1.0%
        alpha = inputs.get('alpha', 5.0)  # デフォルト5%（生産者危険）
        beta = inputs.get('beta', 10.0)  # デフォルト10%（消費者危険）
        c_value = inputs.get('c_value', 0)
        lot_size = inputs.get('lot_size', 1000)
        
        # データベース実績に基づくAQL/LTPD調整
        adjusted_aql, adjusted_ltpd = self._adjust_aql_ltpd_based_on_history(
            aql, ltpd, db_data['defect_rate'], db_data['total_qty']
        )
        
        # 検査水準の簡素化（AQL/LTPD設計ベース）
        if aql <= 0.1:
            level_text = "厳格検査"
            level_reason = f"AQL={aql}%に基づく厳格な統計的設計"
        elif aql <= 0.65:
            level_text = "標準検査"
            level_reason = f"AQL={aql}%に基づく標準的な統計的設計"
        else:
            level_text = "緩和検査"
            level_reason = f"AQL={aql}%に基づく緩和された統計的設計"
        
        results['level_text'] = level_text
        results['level_reason'] = level_reason
        
        # AQL/LTPD設計による抜取数の計算（調整後）
        n_sample, warning_message = self._calculate_aql_ltpd_sample_size(
            adjusted_aql, adjusted_ltpd, alpha, beta, c_value, lot_size
        )
        
        # n>N警告のガイダンス
        guidance_message = self._generate_n_gt_n_guidance(n_sample, lot_size, aql, ltpd, alpha, beta, c_value)
        
        # OCカーブの計算
        oc_curve_data = self._calculate_oc_curve(n_sample, c_value, lot_size)
        
        # 結果の整理
        results['sample_size'] = n_sample
        results['aql'] = adjusted_aql  # 調整後のAQL
        results['ltpd'] = adjusted_ltpd  # 調整後のLTPD
        results['original_aql'] = aql  # 元のAQL
        results['original_ltpd'] = ltpd  # 元のLTPD
        results['alpha'] = alpha
        results['beta'] = beta
        results['oc_curve'] = oc_curve_data
        results['guidance_message'] = guidance_message
        results['adjustment_info'] = self._generate_adjustment_info(
            aql, ltpd, adjusted_aql, adjusted_ltpd, db_data['defect_rate']
        )
        
        if warning_message:
            results['warning_message'] = warning_message
        
        return results
    
    def _calculate_aql_ltpd_sample_size(self, aql, ltpd, alpha, beta, c_value, lot_size):
        """AQL/LTPD設計による抜取数の計算（ロットサイズ考慮版）"""
        cache_key = (
            round(aql, 6),
            round(ltpd, 6),
            round(alpha, 6),
            round(beta, 6),
            int(c_value),
            int(lot_size),
        )
        cached_result = self._sample_size_cache.get(cache_key)
        if cached_result is not None:
            return cached_result

        result = self._calculate_aql_ltpd_sample_size_core(
            aql, ltpd, alpha, beta, c_value, lot_size
        )

        if len(self._sample_size_cache) >= SAMPLE_SIZE_CACHE_LIMIT:
            oldest_key = next(iter(self._sample_size_cache), None)
            if oldest_key is not None:
                self._sample_size_cache.pop(oldest_key, None)
        self._sample_size_cache[cache_key] = result
        return result

    def _calculate_aql_ltpd_sample_size_core(self, aql, ltpd, alpha, beta, c_value, lot_size):
        """AQL/LTPD設計による抜取数の計算（ロットサイズ考慮版）"""
        try:
            # パーセントを小数に変換
            aql_p = aql / 100.0
            ltpd_p = ltpd / 100.0
            alpha_p = alpha / 100.0
            beta_p = beta / 100.0

            # ロットサイズに基づく計算方法の選択
            if lot_size <= 50:
                # 小ロット: 全数検査または高割合抜取
                return self._calculate_small_lot_sample_size(aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size)
            elif lot_size <= 500:
                # 中ロット: 有限母集団補正を重視
                return self._calculate_medium_lot_sample_size(aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size)
            else:
                # 大ロット: 有限母集団補正適用
                return self._calculate_large_lot_sample_size(aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size)

        except (ValueError, OverflowError, ZeroDivisionError):
            return "計算エラー", "AQL/LTPDの値が無効です。"

    def _calculate_small_lot_sample_size(self, aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size):
        """小ロット（50個以下）の抜取数計算"""
        # 小ロットでは全数検査または高割合抜取を推奨
        if lot_size <= 10:
            # 10個以下：全数検査
            return lot_size, f"小ロット（{lot_size}個）のため全数検査を推奨"
        elif lot_size <= 20:
            # 11-20個：80%以上抜取
            n_sample = max(c_value + 1, int(lot_size * 0.8))
            return n_sample, f"小ロット（{lot_size}個）のため高割合抜取（{n_sample}個、{n_sample/lot_size*100:.1f}%）"
        else:
            # 21-50個：60%以上抜取
            n_sample = max(c_value + 1, int(lot_size * 0.6))
            return n_sample, f"小ロット（{lot_size}個）のため高割合抜取（{n_sample}個、{n_sample/lot_size*100:.1f}%）"
    
    def _calculate_medium_lot_sample_size(self, aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size):
        """中ロット（51-500個）の抜取数計算（有限母集団補正重視）"""
        # 有限母集団補正を適用した計算
        if c_value == 0:
            # 超幾何分布ベースの計算
            n_sample = self._calculate_hypergeometric_sample_size(aql_p, ltpd_p, alpha_p, beta_p, lot_size, c_value)
        else:
            # c>0の場合は二分探索（有限母集団補正込み）
            n_sample, warning = self._binary_search_sample_size_with_fpc(aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size)
            if warning:
                return n_sample, warning
        
        # ロットサイズとの比較
        if n_sample > lot_size:
            return lot_size, f"理論値（{n_sample}個）がロットサイズを超えるため全数検査"
        
        return n_sample, None
    
    def _calculate_large_lot_sample_size(self, aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size):
        """大ロット（501個以上）の抜取数計算（有限母集団補正適用）"""
        # 大ロットでも有限母集団補正を適用
        if c_value == 0:
            # 超幾何分布ベースの計算（大ロットでも適用）
            n_sample = self._calculate_hypergeometric_sample_size(aql_p, ltpd_p, alpha_p, beta_p, lot_size, c_value)
        else:
            # c>0の場合は二分探索（有限母集団補正込み）
            n_sample, warning = self._binary_search_sample_size_with_fpc(aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size)
            if warning:
                return n_sample, warning
        
        # ロットサイズとの比較
        if n_sample > lot_size:
            return lot_size, f"理論値（{n_sample}個）がロットサイズを超えるため全数検査"
        
        return n_sample, None
    
    def _calculate_hypergeometric_sample_size(self, aql_p, ltpd_p, alpha_p, beta_p, lot_size, c_value=0):
        """超幾何分布ベースの抜取数計算（有限母集団補正版）"""
        # 二項分布近似で初期値を計算
        n_approx = math.ceil(math.log(beta_p) / math.log(1 - ltpd_p))
        
        # ロットサイズに基づく有限母集団補正
        if lot_size <= 100:
            # 小ロット：高割合抜取
            n_sample = min(lot_size, int(lot_size * 0.8))
        elif lot_size <= 500:
            # 中ロット：強い有限母集団補正
            fpc_factor = 1.0 - (n_approx / lot_size)
            n_sample = int(n_approx * fpc_factor)
        else:
            # 大ロット：軽微な有限母集団補正
            fpc_factor = 1.0 - (n_approx / lot_size) * 0.5
            n_sample = int(n_approx * fpc_factor)
        
        # 最小値と最大値の制限
        n_sample = max(c_value + 1, min(n_sample, lot_size))
        
        return n_sample
    
    def _binary_search_sample_size_with_fpc(self, aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size):
        """有限母集団補正込みの二分探索"""
        low, high = 1, lot_size
        best_n = None
        
        while low <= high:
            mid = (low + high) // 2
            
            # 有限母集団補正の適用判定
            use_fpc = (mid / lot_size > 0.05) or (mid > 50)
            
            if use_fpc:
                # 超幾何分布を使用
                prob_aql = self._hypergeometric_probability(mid, int(lot_size * aql_p), lot_size, c_value)
                prob_ltpd = self._hypergeometric_probability(mid, int(lot_size * ltpd_p), lot_size, c_value)
            else:
                # 二項分布を使用
                prob_aql = self._binomial_probability(mid, aql_p, c_value)
                prob_ltpd = self._binomial_probability(mid, ltpd_p, c_value)
            
            # 条件チェック
            if prob_aql >= (1 - alpha_p) and prob_ltpd <= beta_p:
                best_n = mid
                high = mid - 1
            else:
                low = mid + 1
        
        if best_n is None:
            return lot_size, "条件を満たす抜取数が見つかりません。全数検査を推奨します。"
        
        return best_n, None

    def _binary_search_sample_size(self, aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size):
        """二分探索による抜取数の計算（c>0の場合）"""
        low, high = 1, min(lot_size, 10000)  # 実用的な上限を設定
        best_n = None
        
        while low <= high:
            mid = (low + high) // 2
            
            # 有限母集団補正を考慮した確率計算（改善版）
            # より厳密な判定基準：n/N > 0.05 または n > 50 の場合に超幾何分布を使用
            use_hypergeometric = (mid / lot_size > 0.05) or (mid > 50)
            
            if use_hypergeometric:
                # 超幾何分布での確率計算
                paql = self._hypergeometric_probability(mid, c_value, lot_size, int(lot_size * aql_p))
                pltpd = self._hypergeometric_probability(mid, c_value, lot_size, int(lot_size * ltpd_p))
            else:
                # 二項分布での確率計算
                paql = _cached_binom_cdf(c_value, mid, aql_p)
                pltpd = _cached_binom_cdf(c_value, mid, ltpd_p)
            
            # 条件チェック
            # P(合格|AQL) >= 1 - α かつ P(合格|LTPD) <= β
            if paql >= (1 - alpha_p) and pltpd <= beta_p:
                best_n = mid
                high = mid - 1
            else:
                low = mid + 1
        
        if best_n is None:
            return f"全数検査必要（計算断念）", \
                   f"c={c_value}、AQL={aql_p*100:.2f}%、LTPD={ltpd_p*100:.2f}%の条件では、ロットサイズ（{lot_size:,}個）を超える抜取が必要です。"
        else:
            return best_n, None
    
    def _hypergeometric_probability(self, n, D, N, c):
        """超幾何分布による確率計算"""
        if D == 0:
            return 1.0 if c == 0 else 0.0
        if n > N or c > min(n, D):
            return 0.0
        
        try:
            return _cached_hypergeom_cdf(c, N, D, n)
        except:
            return 0.0
    
    def _calculate_oc_curve(self, n_sample, c_value, lot_size):
        """OCカーブ（Operating Characteristic Curve）の計算"""
        if isinstance(n_sample, str) or n_sample <= 0:
            return []
        
        oc_data = []
        defect_rates = [0.0, 0.1, 0.25, 0.5, 1.0, 1.5, 2.0, 3.0, 5.0, 10.0]  # 不良率（%）
        
        for p_percent in defect_rates:
            p = p_percent / 100.0
            
            # 有限母集団補正の判定（改善版）
            use_hypergeometric = (n_sample / lot_size > 0.05) or (n_sample > 50)
            
            if use_hypergeometric:  # 超幾何分布
                prob = self._hypergeometric_probability(n_sample, int(lot_size * p), lot_size, c_value)
            else:  # 二項分布
                prob = _cached_binom_cdf(c_value, n_sample, p)
            
            oc_data.append({
                'defect_rate': p_percent,
                'acceptance_probability': prob * 100
            })
        
        return oc_data
    
    def _binomial_probability(self, n, p, c):
        """二項分布による合格確率の計算"""
        try:
            if p == 0:
                return 1.0  # 不良率が0%の場合は確実に合格
            
            # 二項分布の累積分布関数
            return _cached_binom_cdf(c, n, p)
        except:
            return 0.0
    
    def _generate_n_gt_n_guidance(self, n_sample, lot_size, aql, ltpd, alpha, beta, c_value):
        """n>N警告のガイダンス生成"""
        if isinstance(n_sample, str) or n_sample > lot_size:
            guidance = "【n>N警告：AQL/LTPDの見直し提案】\n\n"
            guidance += f"現在の設定では抜取数（{n_sample if isinstance(n_sample, int) else '理論値'}個）がロットサイズ（{lot_size:,}個）を超えています。\n\n"
            guidance += "【推奨される見直し案】\n\n"
            
            # 案1: AQLを緩和
            guidance += "1. AQLを緩和する（より良いロットを通しやすく）\n"
            for new_aql in [0.4, 0.65, 1.0, 1.5]:
                test_n, _ = self._calculate_aql_ltpd_sample_size(new_aql, ltpd, alpha, beta, c_value, lot_size)
                if isinstance(test_n, int) and test_n <= lot_size:
                    guidance += f"   • AQL={new_aql}% → 抜取数: {test_n:,}個\n"
            guidance += "\n"
            
            # 案2: LTPDを厳しく
            guidance += "2. LTPDを厳しくする（より悪いロットを止めやすく）\n"
            for new_ltpd in [1.5, 2.0, 2.5, 3.0]:
                test_n, _ = self._calculate_aql_ltpd_sample_size(aql, new_ltpd, alpha, beta, c_value, lot_size)
                if isinstance(test_n, int) and test_n <= lot_size:
                    guidance += f"   • LTPD={new_ltpd}% → 抜取数: {test_n:,}個\n"
            guidance += "\n"
            
            # 案3: リスクを調整
            guidance += "3. リスク（α/β）を調整する\n"
            risk_combinations = [
                (10.0, 10.0, "α=10%, β=10%"),
                (15.0, 10.0, "α=15%, β=10%"),
                (10.0, 15.0, "α=10%, β=15%")
            ]
            
            for new_alpha, new_beta, label in risk_combinations:
                test_n, _ = self._calculate_aql_ltpd_sample_size(aql, ltpd, new_alpha, new_beta, c_value, lot_size)
                if isinstance(test_n, int) and test_n <= lot_size:
                    guidance += f"   • {label} → 抜取数: {test_n:,}個\n"
            guidance += "\n"
            
            # 案4: c値を上げる
            guidance += "4. c値（許容不良数）を上げる\n"
            for new_c in [1, 2, 3]:
                test_n, _ = self._calculate_aql_ltpd_sample_size(aql, ltpd, alpha, beta, new_c, lot_size)
                if isinstance(test_n, int) and test_n <= lot_size:
                    guidance += f"   • c={new_c} → 抜取数: {test_n:,}個\n"
            guidance += "\n"
            
            guidance += "【ISO 2859-1標準の推奨値】\n"
            guidance += "• AQL=0.25%, LTPD=1.0%, α=5%, β=10%, c=0 → 抜取数≈230個\n"
            guidance += "• 小ロット（<1000個）では全数検査も検討してください\n\n"
            guidance += "※ 品質要求に応じて最適な条件を選択してください。"
            
            return guidance
        else:
            return None

    def calculate_alternatives(self, db_data, inputs):
        """AQL/LTPD設計に基づく代替案の計算"""
        lot_size = inputs['lot_size']
        current_aql = inputs.get('aql', 0.25)
        current_ltpd = inputs.get('ltpd', 1.0)
        current_alpha = inputs.get('alpha', 5.0)
        current_beta = inputs.get('beta', 10.0)
        current_c = inputs.get('c_value', 0)
        
        alternatives = "【AQL/LTPD設計による代替案の提案】\n\n"
        
        # 案1: AQLを緩和する
        alternatives += "1. AQLを緩和する場合（より良いロットを通しやすく）:\n"
        for aql in [0.4, 0.65, 1.0, 1.5]:
            n_sample, _ = self._calculate_aql_ltpd_sample_size(
                aql, current_ltpd, current_alpha, current_beta, current_c, lot_size
            )
            if isinstance(n_sample, int):
                alternatives += f"   AQL={aql}%: {n_sample:,}個\n"
            else:
                alternatives += f"   AQL={aql}%: {n_sample}\n"
        alternatives += "\n"
        
        # 案2: LTPDを厳しくする
        alternatives += "2. LTPDを厳しくする場合（より悪いロットを止めやすく）:\n"
        for ltpd in [0.5, 0.8, 1.2, 1.5]:
            n_sample, _ = self._calculate_aql_ltpd_sample_size(
                current_aql, ltpd, current_alpha, current_beta, current_c, lot_size
            )
            if isinstance(n_sample, int):
                alternatives += f"   LTPD={ltpd}%: {n_sample:,}個\n"
            else:
                alternatives += f"   LTPD={ltpd}%: {n_sample}\n"
        alternatives += "\n"
        
        # 案3: リスクを調整する
        alternatives += "3. リスク（α/β）を調整する場合:\n"
        risk_combinations = [
            (10.0, 5.0, "α=10%, β=5%"),
            (5.0, 5.0, "α=5%, β=5%"),
            (10.0, 10.0, "α=10%, β=10%"),
            (15.0, 10.0, "α=15%, β=10%")
        ]
        
        for alpha, beta, label in risk_combinations:
            n_sample, _ = self._calculate_aql_ltpd_sample_size(
                current_aql, current_ltpd, alpha, beta, current_c, lot_size
            )
            if isinstance(n_sample, int):
                alternatives += f"   {label}: {n_sample:,}個\n"
            else:
                alternatives += f"   {label}: {n_sample}\n"
        alternatives += "\n"
        
        # 案4: c値を上げる
        alternatives += "4. c値（許容不良数）を上げる場合:\n"
        for c_val in [1, 2, 3]:
            n_sample, _ = self._calculate_aql_ltpd_sample_size(
                current_aql, current_ltpd, current_alpha, current_beta, c_val, lot_size
            )
            if isinstance(n_sample, int):
                alternatives += f"   c={c_val}: {n_sample:,}個\n"
            else:
                alternatives += f"   c={c_val}: {n_sample}\n"
        alternatives += "\n"
        
        # 推奨案
        alternatives += "【推奨案】\n"
        alternatives += "現在の条件では統計的に適切な抜取検査が困難です。\n"
        alternatives += "以下のいずれかを検討してください:\n\n"
        alternatives += "• AQLを0.4%以上に緩和する\n"
        alternatives += "• LTPDを1.5%以上に設定する\n"
        alternatives += "• α（生産者危険）を10%以上に設定する\n"
        alternatives += "• c値を1以上に設定する\n"
        alternatives += "• 全数検査の実施\n\n"
        alternatives += "※ ISO 2859-1標準に基づく推奨値:\n"
        alternatives += "   AQL=0.25%, LTPD=1.0%, α=5%, β=10%, c=0 → n≈230\n"
        alternatives += "※ 品質要求に応じて最適な条件を選択してください。"
        
        return alternatives
    
    def _adjust_aql_ltpd_based_on_history(self, original_aql, original_ltpd, historical_defect_rate, total_quantity):
        """データベース実績に基づくAQL/LTPD調整"""
        # データが不十分な場合は元の値をそのまま使用
        if total_quantity < 100 or historical_defect_rate is None:
            return original_aql, original_ltpd
        
        # 実績不良率に基づく調整ロジック
        adjusted_aql = original_aql
        adjusted_ltpd = original_ltpd
        
        # 実績不良率が非常に低い場合（0.1%未満）
        if historical_defect_rate < 0.1:
            # AQLをより厳しく設定（品質が良いため）
            adjusted_aql = max(0.1, original_aql * 0.5)
            adjusted_ltpd = max(0.5, original_ltpd * 0.7)
            
        # 実績不良率が低い場合（0.1%～0.5%）
        elif historical_defect_rate < 0.5:
            # AQLをやや厳しく設定
            adjusted_aql = max(0.15, original_aql * 0.7)
            adjusted_ltpd = max(0.7, original_ltpd * 0.8)
            
        # 実績不良率が高い場合（1.0%以上）
        elif historical_defect_rate > 1.0:
            # AQLを緩く設定（品質に課題があるため）
            adjusted_aql = min(2.0, original_aql * 1.5)
            adjusted_ltpd = min(5.0, original_ltpd * 1.3)
            
        # 実績不良率が非常に高い場合（2.0%以上）
        elif historical_defect_rate > 2.0:
            # AQLをさらに緩く設定
            adjusted_aql = min(3.0, original_aql * 2.0)
            adjusted_ltpd = min(8.0, original_ltpd * 1.5)
        
        # データ量に基づく信頼度調整
        confidence_factor = min(1.0, total_quantity / 1000.0)  # 1000個以上で最大信頼度
        
        # 信頼度が低い場合は調整を控えめに
        if confidence_factor < 0.5:
            adjusted_aql = original_aql + (adjusted_aql - original_aql) * confidence_factor
            adjusted_ltpd = original_ltpd + (adjusted_ltpd - original_ltpd) * confidence_factor
        
        return round(adjusted_aql, 3), round(adjusted_ltpd, 3)
    
    def _generate_adjustment_info(self, original_aql, original_ltpd, adjusted_aql, adjusted_ltpd, historical_defect_rate):
        """調整情報の生成"""
        if original_aql == adjusted_aql and original_ltpd == adjusted_ltpd:
            return None
        
        info = "【データベース実績に基づくAQL/LTPD調整】\n\n"
        info += f"実績不良率: {historical_defect_rate:.3f}%\n\n"
        info += f"元の設定:\n"
        info += f"• AQL: {original_aql}% → {adjusted_aql}%\n"
        info += f"• LTPD: {original_ltpd}% → {adjusted_ltpd}%\n\n"
        
        if historical_defect_rate < 0.5:
            info += "調整理由: 実績不良率が低いため、効率的な検査基準に調整\n"
            info += "効果: 検査コストの削減、実務運用の最適化\n"
        elif historical_defect_rate > 1.0:
            info += "調整理由: 実績不良率が高いため、より厳しい検査基準を適用\n"
            info += "効果: 品質維持の強化、不良品流出の防止\n"
        else:
            info += "調整理由: 実績不良率に基づく微調整\n"
            info += "効果: 統計的精度の向上\n"
        
        return info
