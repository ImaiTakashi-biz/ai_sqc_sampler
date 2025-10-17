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
            try:
                # PyInstaller環境でのmatplotlib初期化を安全に行う
                import os
                import sys
                
                # PyInstaller環境でのmatplotlib設定
                if getattr(sys, 'frozen', False):
                    # PyInstallerでビルドされた場合の特別な設定
                    os.environ['MPLBACKEND'] = 'TkAgg'
                
                # matplotlibの初期化
                import matplotlib
                matplotlib.use('TkAgg', force=True)  # 強制的にバックエンドを設定
                
                # matplotlibの内部設定をリセット
                matplotlib.rcdefaults()
                
                # 段階的なインポートでエラーを特定
                try:
                    import matplotlib.pyplot as plt
                    # PyInstaller環境での追加設定
                    if getattr(sys, 'frozen', False):
                        plt.ioff()  # インタラクティブモードを無効化
                except ImportError as e:
                    raise Exception(f"matplotlib.pyplotの読み込みに失敗: {e}")
                
                try:
                    import matplotlib.font_manager as fm
                except ImportError as e:
                    raise Exception(f"matplotlib.font_managerの読み込みに失敗: {e}")
                
                try:
                    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
                except ImportError as e:
                    raise Exception(f"matplotlib.backends.backend_tkaggの読み込みに失敗: {e}")
                
                self._plot_modules = {
                    'plt': plt,
                    'fm': fm,
                    'FigureCanvasTkAgg': FigureCanvasTkAgg,
                }
                self._setup_japanese_font(self._plot_modules)
            except Exception as e:
                print(f"matplotlib import error: {e}")
                # エラーを詳細に記録
                import traceback
                traceback.print_exc()
                raise Exception(f"matplotlibの読み込みに失敗しました: {e}")

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
    
    def _log_plot_error(self, error):
        try:
            import os
            import sys
            import datetime
            base_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
            log_path = os.path.join(base_dir, "oc_curve_errors.log")
            with open(log_path, "a", encoding="utf-8") as log_file:
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                log_file.write(f"[{timestamp}] {repr(error)}\n")
        except Exception:
            pass

    def create_oc_curve_dialog(self, parent, oc_data, aql, ltpd, alpha, beta, n_sample, c_value, lot_size):
        """OCカーブ表示ダイアログの作成"""
        dialog = tk.Toplevel(parent)
        dialog.title("OCカーブ（Operating Characteristic Curve）")
        initial_width, initial_height = 960, 720
        dialog.geometry(f"{initial_width}x{initial_height}")
        dialog.configure(bg="#f8f9fa")
        dialog.resizable(True, True)
        dialog.minsize(900, 700)
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(0, weight=1)
        
        # 中央配置
        x = max((parent.winfo_screenwidth() - initial_width) // 2, 0)
        y = max((parent.winfo_screenheight() - initial_height) // 2, 0)
        dialog.geometry(f"{initial_width}x{initial_height}+{x}+{y}")
        
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
        
        # グラフ表示領域
        content_frame = tk.Frame(dialog, bg="#f8f9fa")
        content_frame.pack(fill='both', expand=True, padx=20, pady=(10, 10))
        content_frame.pack_propagate(False)

        fallback_required = False
        fallback_reason = ""

        if not oc_data:
            fallback_required = True
            fallback_reason = "OCカーブデータがありません。数値表で表示します。"
        else:
            plot_frame = tk.Frame(content_frame, bg="#f8f9fa")
            plot_frame.pack(fill='both', expand=True)
            try:
                self.draw_oc_curve(
                    plot_frame,
                    oc_data,
                    aql,
                    ltpd,
                    alpha,
                    beta,
                    n_sample,
                    c_value,
                    lot_size
                )
            except Exception as e:
                fallback_required = True
                fallback_reason = "グラフ表示に失敗しました。数値表で表示します。"
                print(f"OCカーブ描画エラー: {e}")
                self._log_plot_error(e)
                import traceback
                traceback.print_exc()

        if fallback_required:
            notice_frame = tk.Frame(content_frame, bg="#f8f9fa")
            notice_frame.pack(fill='x', pady=(0, 10))
            tk.Label(
                notice_frame,
                text=fallback_reason,
                font=("Meiryo", 11, "bold"),
                fg="#dc3545",
                bg="#f8f9fa",
                justify='left',
                wraplength=760
            ).pack(anchor='w')
            self.show_alternative_table(content_frame, oc_data, aql, ltpd, alpha, beta)
        
        # 説明テキスト
        footer_frame = tk.Frame(dialog, bg="#f8f9fa")
        footer_frame.pack(fill='x', padx=20, pady=(0, 20))
        
        explanation_text = (
            "【OCカーブの見方】\n"
            "• 横軸：ロットの真の不良率（%）\n"
            "• 縦軸：そのロットが合格する確率（%）\n"
            f"• 青い点：AQL={aql}%での合格確率（目標：{100-alpha:.0f}%以上）\n"
            f"• 赤い点：LTPD={ltpd}%での合格確率（目標：{beta:.0f}%以下）\n"
            "• 曲線が急峻なほど検査が厳しく、緩やかなほど検査が緩い"
        )
        
        explanation_label = tk.Label(
            footer_frame, 
            text=explanation_text, 
            font=("Meiryo", 9), 
            fg="#495057", 
            bg="#f8f9fa",
            justify='left',
            wraplength=760
        )
        explanation_label.pack(side='left', fill='x', expand=True)
        
        # 閉じるボタン
        close_button = tk.Button(
            footer_frame, 
            text="閉じる", 
            command=dialog.destroy, 
            font=("Meiryo", 10, "bold"), 
            bg="#6c757d", 
            fg="#ffffff", 
            relief="flat", 
            padx=20, 
            pady=5
        )
        close_button.pack(side='right', padx=(20, 0))
    
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
        aql_prob = self.calculate_acceptance_probability(oc_data, aql)
        ax.scatter([aql], [aql_prob], color='blue', s=100, marker='o', 
                  label=f'AQL={aql}% (Accept={aql_prob:.1f}%)', zorder=5)
        
        # LTPD点の描画
        ltpd_prob = self.calculate_acceptance_probability(oc_data, ltpd)
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
    
    def calculate_acceptance_probability(self, oc_data, target_defect_rate):
        """OCカーブデータから合格確率を推定（線形補間）"""
        if not oc_data:
            return 0.0
        
        points = []
        for entry in oc_data:
            try:
                if isinstance(entry, dict):
                    rate = entry.get('defect_rate')
                    if rate is None:
                        rate = entry.get('defect_rate_percent')
                    prob = entry.get('acceptance_probability')
                    if prob is None:
                        prob = entry.get('acceptance_prob')
                elif isinstance(entry, (list, tuple)) and len(entry) >= 2:
                    rate, prob = entry[0], entry[1]
                else:
                    continue
                rate = float(rate)
                prob = float(prob)
                points.append((rate, prob))
            except (TypeError, ValueError):
                continue
        
        if not points:
            return 0.0
        
        points.sort(key=lambda x: x[0])
        target = float(target_defect_rate)
        
        for rate, prob in points:
            if abs(rate - target) < 1e-9:
                return prob
        
        if target <= points[0][0]:
            return points[0][1]
        if target >= points[-1][0]:
            return points[-1][1]
        
        for (r1, p1), (r2, p2) in zip(points, points[1:]):
            if r1 <= target <= r2:
                if r2 == r1:
                    return p1
                ratio = (target - r1) / (r2 - r1)
                return p1 + (p2 - p1) * ratio
        
        return points[-1][1]
    
    def show_alternative_table(self, parent, oc_data, aql, ltpd, alpha, beta):
        """OCカーブの代替表示（数値表）"""
        try:
            # 数値表の表示
            table_frame = tk.LabelFrame(
                parent,
                text="OCカーブ数値表（代替表示）",
                font=("Meiryo", 12, "bold"),
                fg="#2c3e50",
                bg="#f8f9fa",
                padx=10,
                pady=10
            )
            table_frame.pack(fill='both', expand=True, padx=20, pady=10)
            
            # テーブルヘッダー
            header_frame = tk.Frame(table_frame, bg="#f8f9fa")
            header_frame.pack(fill='x', pady=(0, 10))
            
            headers = ["不良率(%)", "受入確率", "判定"]
            for i, header in enumerate(headers):
                tk.Label(
                    header_frame,
                    text=header,
                    font=("Meiryo", 10, "bold"),
                    fg="#2c3e50",
                    bg="#e9ecef",
                    width=15,
                    relief="solid",
                    borderwidth=1
                ).grid(row=0, column=i, padx=1, pady=1, sticky="ew")
            
            # データ行
            displayed_rows = 0
            for i, data_point in enumerate(oc_data[:10]):  # 最初の10行のみ表示
                row_frame = tk.Frame(table_frame, bg="#f8f9fa")
                row_frame.pack(fill='x')
                
                # データ形式の確認と変換
                try:
                    if isinstance(data_point, dict):
                        defect_rate_value = data_point.get('defect_rate')
                        if defect_rate_value is None:
                            defect_rate_value = data_point.get('defect_rate_percent')
                        acceptance_value = data_point.get('acceptance_probability')
                        if acceptance_value is None:
                            acceptance_value = data_point.get('acceptance_prob')
                        if defect_rate_value is None or acceptance_value is None:
                            continue
                        defect_rate = float(defect_rate_value)
                        acceptance_prob = float(acceptance_value)
                    elif isinstance(data_point, (list, tuple)) and len(data_point) >= 2:
                        defect_rate = float(data_point[0])
                        acceptance_prob = float(data_point[1])
                    else:
                        # データ形式が異なる場合はスキップ
                        continue
                except (ValueError, TypeError, IndexError):
                    # データ変換に失敗した場合はスキップ
                    continue
                
                displayed_rows += 1
                
                # 不良率
                tk.Label(
                    row_frame,
                    text=f"{defect_rate:.2f}",
                    font=("Meiryo", 9),
                    fg="#495057",
                    bg="#ffffff",
                    width=15,
                    relief="solid",
                    borderwidth=1
                ).grid(row=0, column=0, padx=1, pady=1, sticky="ew")
                
                # 受入確率
                tk.Label(
                    row_frame,
                    text=f"{acceptance_prob:.4f}",
                    font=("Meiryo", 9),
                    fg="#495057",
                    bg="#ffffff",
                    width=15,
                    relief="solid",
                    borderwidth=1
                ).grid(row=0, column=1, padx=1, pady=1, sticky="ew")
                
                # 判定
                if defect_rate <= aql:
                    judgment = "合格"
                    color = "#d4edda"
                elif defect_rate >= ltpd:
                    judgment = "不合格"
                    color = "#f8d7da"
                else:
                    judgment = "境界"
                    color = "#fff3cd"
                
                tk.Label(
                    row_frame,
                    text=judgment,
                    font=("Meiryo", 9),
                    fg="#495057",
                    bg=color,
                    width=15,
                    relief="solid",
                    borderwidth=1
                ).grid(row=0, column=2, padx=1, pady=1, sticky="ew")
            
            # データが表示されなかった場合のメッセージ
            if displayed_rows == 0:
                no_data_label = tk.Label(
                    table_frame,
                    text="OCカーブデータが利用できません",
                    font=("Meiryo", 10),
                    fg="#dc3545",
                    bg="#f8f9fa"
                )
                no_data_label.pack(pady=20)
            
            # 説明
            explanation = tk.Label(
                table_frame,
                text="※ グラフ表示が利用できないため、数値表で代替表示しています",
                font=("Meiryo", 9),
                fg="#6c757d",
                bg="#f8f9fa"
            )
            explanation.pack(pady=(10, 0))
            
        except Exception as e:
            # 代替表示も失敗した場合
            fallback_label = tk.Label(
                parent,
                text=f"代替表示も失敗しました: {str(e)}",
                font=("Meiryo", 10),
                fg="#dc3545",
                bg="#f8f9fa"
            )
            fallback_label.pack(pady=20)
