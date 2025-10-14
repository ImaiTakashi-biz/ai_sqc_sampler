"""
統計計算エンジンモジュール
SQC統計計算とデータ処理を管理
AQL/LTPD設計に基づく抜取検査計算
"""

import math
from functools import lru_cache

from constants import InspectionConstants, DEFECT_COLUMNS

SAMPLE_SIZE_CACHE_LIMIT = 128

INSPECTION_COMMENT_PRESETS = {
    "tightened": {
        "label": "\u5f37\u5316",
        "aql": 0.10,
        "ltpd": 0.50,
        "alpha": 3.0,
        "beta": 5.0,
        "c_value": 0,
        "description": "\u521d\u671f\u6d41\u52d5\u30fb\u4e0d\u5177\u5408\u518d\u767a\u6642"
    },
    "standard": {
        "label": "\u6a19\u6e96",
        "aql": 0.25,
        "ltpd": 1.00,
        "alpha": 5.0,
        "beta": 10.0,
        "c_value": 0,
        "description": "\u901a\u5e38\u30ed\u30c3\u30c8"
    },
    "reduced": {
        "label": "\u7de9\u548c",
        "aql": 0.40,
        "ltpd": 1.50,
        "alpha": 10.0,
        "beta": 15.0,
        "c_value": 0,
        "description": "\u5b89\u5b9a\u751f\u7523\u30fb\u9867\u5ba2\u4fe1\u983c\u88fd\u54c1"
    }
}

INSPECTION_COMMENT_BY_LABEL = {
    preset["label"]: preset for preset in INSPECTION_COMMENT_PRESETS.values()
}
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
        if total_qty > 0 and defect_counts:
            totals = float(total_qty)
            defect_rates = [
                (col, (count or 0) / totals * 100.0, count or 0)
                for col, count in zip(DEFECT_COLUMNS, defect_counts)
                if (count or 0) > 0
            ]
            defect_rates.sort(key=lambda x: x[2], reverse=True)
            data['defect_rates_sorted'] = defect_rates
            data['best5'] = [(col, count) for col, _, count in defect_rates[:5]]
        else:
            data['defect_rates_sorted'] = []
            data['best5'] = []
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
        
        mode_key = inputs.get('inspection_mode_key') or inputs.get('inspection_mode')
        mode_details = inputs.get('inspection_mode_details') or {}
        mode_label = inputs.get('inspection_mode_label') or mode_details.get('label')

        preset_details = None
        if mode_key and mode_key in INSPECTION_COMMENT_PRESETS:
            preset_details = INSPECTION_COMMENT_PRESETS[mode_key].copy()
        elif mode_label and mode_label in INSPECTION_COMMENT_BY_LABEL:
            preset_details = INSPECTION_COMMENT_BY_LABEL[mode_label].copy()
        elif mode_details:
            preset_details = mode_details.copy()

        if preset_details and mode_label and not preset_details.get('label'):
            preset_details['label'] = mode_label

        active_mode_key = mode_key if mode_key in INSPECTION_COMMENT_PRESETS else None

        if preset_details:
            label = preset_details.get('label') or mode_label or "標準"
            level_text = label if label.endswith("検査") else f"{label}検査"
            level_reason = self._compose_inspection_comment(preset_details)
            if not active_mode_key and preset_details.get('label'):
                lookup_label = preset_details['label']
                for key, preset in INSPECTION_COMMENT_PRESETS.items():
                    if preset['label'] == lookup_label:
                        active_mode_key = key
                        break
        else:
            if aql <= 0.1:
                fallback_key = "tightened"
            elif aql <= 0.4:
                fallback_key = "standard"
            else:
                fallback_key = "reduced"
            preset_details = INSPECTION_COMMENT_PRESETS[fallback_key].copy()
            label = preset_details.get('label', "標準")
            level_text = label if label.endswith("検査") else f"{label}検査"
            level_reason = self._compose_inspection_comment(preset_details)
            active_mode_key = fallback_key

        results['level_text'] = level_text
        results['level_reason'] = level_reason
        results['inspection_mode_label'] = preset_details.get('label', mode_label or "")
        results['inspection_mode_key'] = active_mode_key
        results['inspection_mode_details'] = preset_details.copy() if preset_details else {}
        
        # AQL/LTPD設計による抜取数の計算（調整後）
        n_sample, warning_message = self._calculate_aql_ltpd_sample_size(
            adjusted_aql, adjusted_ltpd, alpha, beta, c_value, lot_size
        )
        if isinstance(warning_message, str) and '全数検査' in warning_message and isinstance(n_sample, int):
            fallback_ratio = {
                'tightened': 0.60,
                'standard': 0.40,
                'reduced': 0.25
            }.get(active_mode_key)
            if fallback_ratio:
                candidate = math.ceil(lot_size * fallback_ratio)
                if candidate < lot_size:
                    n_sample = max(c_value + 1, candidate)
                    warning_message = None

        if isinstance(n_sample, str):
            warning_message = warning_message or n_sample
            n_sample = lot_size

        ratio_baseline = {
            'tightened': 0.60,
            'standard': 0.40,
            'reduced': 0.25
        }
        if isinstance(n_sample, int) and active_mode_key in ratio_baseline:
            minimum = max(1, math.ceil(lot_size * ratio_baseline[active_mode_key]))
            if minimum > lot_size:
                minimum = lot_size
            if n_sample < minimum:
                n_sample = minimum
        max_ratio_map = {
            'tightened': 0.90,
            'standard': 0.70,
            'reduced': 0.35
        }
        if isinstance(n_sample, int) and active_mode_key in max_ratio_map:
            maximum = math.ceil(lot_size * max_ratio_map[active_mode_key])
            if maximum < 1:
                maximum = 1
            if maximum < n_sample:
                n_sample = maximum


        
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

    def _calculate_zero_acceptance_sample_size(self, aql_p, ltpd_p, alpha_p, beta_p):
        """c=0 の場合の理論値（二項近似）を算出"""
        if aql_p <= 0 or ltpd_p <= 0:
            return 0

        # (1 - p)^n >= 1 - alpha -> n >= log(1 - alpha) / log(1 - p)
        n_aql = 0
        if alpha_p < 1:
            n_aql = math.log(1 - alpha_p) / math.log(1 - aql_p)
        n_ltpd = math.log(beta_p) / math.log(1 - ltpd_p)

        n = max(n_aql, n_ltpd)
        if not math.isfinite(n) or n <= 0:
            return 0
        return math.ceil(n)

    def _calculate_small_lot_sample_size(self, aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size):
        """小ロット（50個以下）の抜取数計算"""
        if lot_size <= 10:
            return lot_size, f"小ロット（{lot_size}個）のため全数検査を推奨"

        n_sample, warning = self._binary_search_sample_size_with_fpc(
            aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size
        )
        if isinstance(n_sample, str):
            return lot_size, warning or n_sample
        if n_sample >= lot_size:
            return lot_size, f"小ロット（{lot_size}個）のため全数検査を推奨"

        return n_sample, warning

    def _calculate_medium_lot_sample_size(self, aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size):
        """中ロット（51-500個）の抜取数計算（有限母集団補正重視）"""
        n_sample, warning = self._binary_search_sample_size_with_fpc(
            aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size
        )
        if isinstance(n_sample, str):
            return lot_size, warning or n_sample
        if n_sample >= lot_size:
            return lot_size, f"算出値（{n_sample}個）がロットサイズを超えるため全数検査"

        return n_sample, warning

    def _calculate_large_lot_sample_size(self, aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size):
        """大ロット（501個以上）の抜取数計算（有限母集団補正適用）"""
        n_sample, warning = self._binary_search_sample_size_with_fpc(
            aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size
        )
        if isinstance(n_sample, str):
            return lot_size, warning or n_sample
        if n_sample >= lot_size:
            return lot_size, f"算出値（{n_sample}個）がロットサイズを超えるため全数検査"

        return n_sample, warning

    def _binary_search_sample_size_with_fpc(self, aql_p, ltpd_p, alpha_p, beta_p, c_value, lot_size):
        """二分探索による抜取数の計算（c>0の場合）"""
        low, high = 1, min(lot_size, 10000)  # 実用的な上限を設定
        best_n = None
        
        while low <= high:
            mid = (low + high) // 2
            
            # 有限母集団補正を考慮した確率計算（改善版）
            # より厳密な判定基準：n/N > 0.05 または n > 50 の場合に超幾何分布を使用
            use_hypergeometric = (mid / lot_size > 0.05) or (mid > 50)
            
            if use_hypergeometric:
                # 超幾何分布での確率計算（期待不良数が0にならないよう調整）
                defect_count_aql = max(1, round(lot_size * aql_p))
                defect_count_ltpd = max(1, round(lot_size * ltpd_p))
                paql = self._hypergeometric_probability(mid, c_value, lot_size, defect_count_aql)
                pltpd = self._hypergeometric_probability(mid, c_value, lot_size, defect_count_ltpd)
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
            if c_value == 0:
                approx = self._calculate_zero_acceptance_sample_size(aql_p, ltpd_p, alpha_p, beta_p)
                if approx and approx < lot_size:
                    return max(c_value + 1, approx), None
            return lot_size, f"c={c_value}、AQL={aql_p*100:.2f}%、LTPD={ltpd_p*100:.2f}%の条件では全数検査を推奨します。"
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
    
    @staticmethod
    def _compose_inspection_comment(details):
        """検査水準コメント文字列を生成"""

        def fmt(value, percent=True):
            if value is None:
                return "-"
            try:
                numeric = float(value)
            except (TypeError, ValueError):
                return str(value)

            if percent:
                formatted = f"{numeric:.2f}".rstrip('0').rstrip('.')
                return formatted or "0"

            if numeric.is_integer():
                return str(int(numeric))
            return f"{numeric:.2f}".rstrip('0').rstrip('.') or "0"

        aql_str = fmt(details.get('aql'))
        ltpd_str = fmt(details.get('ltpd'))
        alpha_str = fmt(details.get('alpha'))
        beta_str = fmt(details.get('beta'))
        c_value_str = fmt(details.get('c_value'), percent=False)

        comment = (
            f"\u6761\u4ef6: AQL={aql_str}%, LTPD={ltpd_str}%, "
            f"\u03b1={alpha_str}%, \u03b2={beta_str}%, c={c_value_str}"
        )

        description = details.get('description')
        if description:
            comment += f" | \u7528\u9014: {description}"

        return comment
    
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
        if total_quantity < 100 or historical_defect_rate is None:
            return original_aql, original_ltpd

        rate = historical_defect_rate

        if rate <= 0.1:
            adjustment_factor = 1.10
        elif rate <= 0.5:
            adjustment_factor = 1.05
        elif rate <= 1.5:
            adjustment_factor = 1.0
        elif rate <= 2.5:
            adjustment_factor = 0.9
        else:
            adjustment_factor = 0.8

        target_aql = original_aql * adjustment_factor
        target_ltpd = original_ltpd * adjustment_factor

        confidence_factor = max(0.0, min(1.0, total_quantity / 1500.0))
        adjusted_aql = original_aql + (target_aql - original_aql) * confidence_factor
        adjusted_ltpd = original_ltpd + (target_ltpd - original_ltpd) * confidence_factor

        adjusted_aql = max(0.02, min(5.0, adjusted_aql))
        adjusted_ltpd = max(0.2, min(10.0, adjusted_ltpd))

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

