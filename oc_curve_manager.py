"""
OCカーブ（Operating Characteristic Curve）管理モジュール
AQL/LTPD設計の根拠を視覚化
"""

import tkinter as tk


class OCCurveManager:
    """OCカーブ管理クラス"""
    
    def __init__(self):
        self._plot_modules = None
    
    def _ensure_plot_modules(self):
        if self._plot_modules is None:
            import matplotlib.pyplot as plt
            import matplotlib.font_manager as fm
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from scipy.stats import binom as sp_binom, hypergeom as sp_hypergeom

            self._plot_modules = {
                'plt': plt,
                'fm': fm,
                'FigureCanvasTkAgg': FigureCanvasTkAgg,
                'binom': sp_binom,
                'hypergeom': sp_hypergeom,
            }
            self._setup_japanese_font(self._plot_modules)

        return self._plot_modules

    def _setup_japanese_font(self, modules):
        plt = modules['plt']
        fm = modules['fm']
        try:
            available_fonts = [f.name for f in fm.fontManager.ttflist]
            japanese_fonts = []

            windows_fonts = ['MS Gothic', 'MS Mincho', 'Meiryo', 'Yu Gothic', 'MS PGothic', 'MS PMincho']
            for font in windows_fonts:
                if font in available_fonts:
                    japanese_fonts.append(font)

            other_fonts = ['Hiragino Sans', 'Takao', 'IPAexGothic', 'IPAPGothic', 'VL PGothic', 'Noto Sans CJK JP']
            for font in other_fonts:
                if font in available_fonts:
                    japanese_fonts.append(font)

            if japanese_fonts:
                plt.rcParams['font.family'] = japanese_fonts + ['DejaVu Sans']
            else:
                plt.rcParams['font.family'] = ['DejaVu Sans']
                print('警告: 日本語フォントが見つかりません。グラフの日本語表示に問題がある可能性があります。')

        except Exception as e:
            plt.rcParams['font.family'] = ['DejaVu Sans']
            print(f'フォント設定エラー: {e}')

    def create_oc_curve_dialog(self, parent, oc_data, aql, ltpd, alpha, beta, n_sample, c_value, lot_size):
        """OCカーブ表示ダイアログの作成"""
        dialog = tk.Toplevel(parent)
        dialog.title("OCカーブ（Operating Characteristic Curve）")
        dialog.geometry("800x600")
        dialog.configure(bg="#f8f9fa")
        dialog.resizable(True, True)
        
        # 中央配置
        x = (parent.winfo_screenwidth() // 2) - 400
        y = (parent.winfo_screenheight() // 2) - 300
        dialog.geometry(f"800x600+{x}+{y}")
        
        # モーダル表示
        dialog.transient(parent)
        dialog.grab_set()
        
        # タイトル
        title_label = tk.Label(
            dialog, 
            text="📊 OCカーブ（Operating Characteristic Curve）", 
            font=("Meiryo", 16, "bold"), 
            fg="#2c3e50", 
            bg="#f8f9fa"
        )
        title_label.pack(pady=(20, 10))
        
        # 条件表示
        condition_frame = tk.LabelFrame(
            dialog, 
            text="検査条件", 
            font=("Meiryo", 12, "bold"), 
            fg="#2c3e50", 
            bg="#f8f9fa",
            padx=10,
            pady=10
        )
        condition_frame.pack(fill='x', padx=20, pady=10)
        
        condition_text = f"抜取数: {n_sample:,}個 | c値: {c_value} | ロットサイズ: {lot_size:,}個"
        tk.Label(
            condition_frame, 
            text=condition_text, 
            font=("Meiryo", 10), 
            fg="#495057", 
            bg="#f8f9fa"
        ).pack()
        
        # OCカーブの描画
        self.draw_oc_curve(dialog, oc_data, aql, ltpd, alpha, beta, n_sample, c_value, lot_size)
        
        # 説明テキスト
        explanation_frame = tk.Frame(dialog, bg="#f8f9fa")
        explanation_frame.pack(fill='x', padx=20, pady=10)
        
        explanation_text = (
            "【OCカーブの見方】\n"
            "• 横軸：ロットの真の不良率（%）\n"
            "• 縦軸：そのロットが合格する確率（%）\n"
            f"• 青い点：AQL={aql}%での合格確率（目標：{100-alpha:.0f}%以上）\n"
            f"• 赤い点：LTPD={ltpd}%での合格確率（目標：{beta:.0f}%以下）\n"
            "• 曲線が急峻なほど検査が厳しく、緩やかなほど検査が緩い"
        )
        
        tk.Label(
            explanation_frame, 
            text=explanation_text, 
            font=("Meiryo", 9), 
            fg="#495057", 
            bg="#f8f9fa",
            justify='left',
            wraplength=750
        ).pack(anchor='w')
        
        # 閉じるボタン
        close_button = tk.Button(
            dialog, 
            text="閉じる", 
            command=dialog.destroy, 
            font=("Meiryo", 10, "bold"), 
            bg="#6c757d", 
            fg="#ffffff", 
            relief="flat", 
            padx=20, 
            pady=5
        )
        close_button.pack(pady=20)
    
    def draw_oc_curve(self, parent, oc_data, aql, ltpd, alpha, beta, n_sample, c_value, lot_size):
        """OCカーブの描画"""
        modules = self._ensure_plot_modules()
        plt = modules['plt']
        FigureCanvasTkAgg = modules['FigureCanvasTkAgg']
        # 図の作成
        fig, ax = plt.subplots(figsize=(10, 6))
        fig.patch.set_facecolor('#f8f9fa')
        
        # データの準備
        defect_rates = [point['defect_rate'] for point in oc_data]
        acceptance_probs = [point['acceptance_probability'] for point in oc_data]
        
        # OCカーブの描画
        ax.plot(defect_rates, acceptance_probs, 'b-', linewidth=2, label='OC Curve')
        ax.scatter(defect_rates, acceptance_probs, color='blue', s=30, alpha=0.7)
        
        # AQL点の描画
        aql_prob = self.calculate_acceptance_probability(n_sample, c_value, lot_size, aql)
        ax.scatter([aql], [aql_prob], color='blue', s=100, marker='o', 
                  label=f'AQL={aql}% (Accept={aql_prob:.1f}%)', zorder=5)
        
        # LTPD点の描画
        ltpd_prob = self.calculate_acceptance_probability(n_sample, c_value, lot_size, ltpd)
        ax.scatter([ltpd], [ltpd_prob], color='red', s=100, marker='s', 
                  label=f'LTPD={ltpd}% (Accept={ltpd_prob:.1f}%)', zorder=5)
        
        # 目標線の描画
        ax.axhline(y=100-alpha, color='blue', linestyle='--', alpha=0.5, 
                  label=f'α={alpha}% Target (Accept={100-alpha:.0f}%)')
        ax.axhline(y=beta, color='red', linestyle='--', alpha=0.5, 
                  label=f'β={beta}% Target (Accept={beta:.0f}%)')
        
        # グラフの設定
        ax.set_xlabel('True Defect Rate (%)', fontsize=12)
        ax.set_ylabel('Acceptance Probability (%)', fontsize=12)
        ax.set_title(f'OC Curve (n={n_sample:,}, c={c_value})', fontsize=14, fontweight='bold')
        ax.grid(True, alpha=0.3)
        ax.legend(loc='upper right', fontsize=10)
        ax.set_xlim(0, max(defect_rates) * 1.1)
        ax.set_ylim(0, 100)
        
        # レイアウトの調整
        plt.tight_layout()
        
        # Tkinterに埋め込み
        canvas = FigureCanvasTkAgg(fig, parent)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True, padx=20, pady=10)
    
    def calculate_acceptance_probability(self, n_sample, c_value, lot_size, defect_rate_percent):
        """合格確率の計算"""
        modules = self._ensure_plot_modules()
        binom = modules['binom']
        hypergeom = modules['hypergeom']
        defect_rate = defect_rate_percent / 100.0
        
        # 有限母集団補正の判定
        if n_sample >= lot_size * 0.1:  # 10%以上の場合、超幾何分布
            D = int(lot_size * defect_rate)  # 不良品数
            try:
                prob = hypergeom.cdf(c_value, lot_size, D, n_sample)
                return prob * 100
            except:
                return 0.0
        else:  # 二項分布
            try:
                prob = binom.cdf(c_value, n_sample, defect_rate)
                return prob * 100
            except:
                return 0.0
