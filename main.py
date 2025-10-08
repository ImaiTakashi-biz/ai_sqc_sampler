"""
AI SQC Sampler - メインコントローラー
統計的品質管理によるサンプリングサイズ計算アプリケーション
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import pyodbc
from datetime import datetime
from gui import App
from database import DatabaseManager
from statistics import SQCStatistics
from validation import InputValidator
from constants import InspectionConstants
from config_manager import ConfigManager
from settings_dialog import SettingsDialog


class MainController:
    """メインアプリケーションコントローラー"""
    
    def __init__(self):
        self.config_manager = ConfigManager()
        self.db_manager = DatabaseManager(self.config_manager)
        self.app = App(self)
    

    def run(self):
        """アプリケーションの実行"""
        self.app.mainloop()

    def start_calculation_thread(self):
        """計算処理を別スレッドで開始"""
        inputs = self._get_user_inputs()
        if not inputs:
            return
        
        self._setup_progress_window()
        thread = threading.Thread(target=self._calculation_worker, args=(inputs,))
        thread.daemon = True
        thread.start()

    def _get_user_inputs(self):
        """ユーザー入力の取得と検証"""
        inputs = {
            'product_number': self.app.sample_pn_entry.get().strip(),
            'lot_size_str': self.app.sample_qty_entry.get().strip(),
            'start_date': self.app.sample_start_date_entry.get().strip() or None,
            'end_date': self.app.sample_end_date_entry.get().strip() or None,
            'conf_str': self.app.sample_conf_entry.get().strip() or str(InspectionConstants.DEFAULT_CONFIDENCE),
            'c_str': self.app.sample_c_entry.get().strip() or str(InspectionConstants.DEFAULT_C_VALUE)
        }
        
        # 入力値の検証
        is_valid, errors, validated_data = InputValidator.validate_all_inputs(
            inputs['product_number'],
            inputs['lot_size_str'],
            inputs['conf_str'],
            inputs['c_str'],
            inputs['start_date'],
            inputs['end_date']
        )
        
        if not is_valid:
            error_message = "以下の入力エラーがあります：\n" + "\n".join(f"• {error}" for error in errors)
            messagebox.showwarning("入力エラー", error_message)
            return None
        
        return validated_data

    def _setup_progress_window(self):
        """プログレスウィンドウの設定"""
        self.progress_window = tk.Toplevel(self.app)
        self.progress_window.title("計算中...")
        self.progress_window.geometry("400x150")
        self.progress_window.configure(bg="#f0f0f0")
        self.progress_window.resizable(False, False)
        
        # 中央配置
        x = (self.app.winfo_screenwidth() // 2) - 200
        y = (self.app.winfo_screenheight() // 2) - 75
        self.progress_window.geometry(f"400x150+{x}+{y}")
        
        # プログレスバー
        self.progress_bar = ttk.Progressbar(self.progress_window, mode='indeterminate', length=300)
        self.progress_bar.pack(pady=30)
        self.progress_bar.start()
        
        # ステータスラベル
        self.status_label = tk.Label(
            self.progress_window, 
            text="計算処理中...", 
            font=("Meiryo", 12), 
            bg="#f0f0f0"
        )
        self.status_label.pack(pady=10)
        
        # モーダル表示
        self.progress_window.transient(self.app)
        self.progress_window.grab_set()

    def _calculation_worker(self, inputs):
        """計算処理のワーカースレッド"""
        try:
            # ステータス更新
            self.app.after(0, lambda: self.status_label.config(text="データベースに接続中..."))
            
            # データベース接続
            conn = self.db_manager.get_db_connection()
            if not conn:
                raise ConnectionError("データベース接続に失敗しました")
            
            with conn.cursor() as cursor:
                # ステータス更新
                self.app.after(0, lambda: self.status_label.config(text="不具合データを集計中..."))
                
                # データの取得
                db_data = self._fetch_data(cursor, inputs)
                
                # ステータス更新
                self.app.after(0, lambda: self.status_label.config(text="抜取検査数を計算中..."))
                
                # 統計計算
                stats_results = self._calculate_stats(db_data, inputs)
                
            # ステータス更新
            self.app.after(0, lambda: self.status_label.config(text="結果を表示中..."))
            
            # UI更新
            self.app.after(0, self._update_ui, db_data, stats_results, inputs)
            self.app.after(0, self._finish_calculation, True)
            
        except ConnectionError as e:
            self.app.after(0, self._finish_calculation, False)
            
        except pyodbc.Error as e:
            error_msg = f"データベースエラー: {str(e)}"
            self.app.after(0, lambda: messagebox.showerror("データベースエラー", error_msg))
            self.app.after(0, self._finish_calculation, False)
            
        except ValueError as e:
            error_msg = f"計算エラー: {str(e)}"
            self.app.after(0, lambda: messagebox.showerror("計算エラー", error_msg))
            self.app.after(0, self._finish_calculation, False)
            
        except OverflowError as e:
            error_msg = f"計算範囲エラー: {str(e)}"
            self.app.after(0, lambda: messagebox.showerror("計算範囲エラー", error_msg))
            self.app.after(0, self._finish_calculation, False)
            
        except Exception as e:
            error_msg = f"予期しないエラー: {str(e)}"
            self.app.after(0, lambda: messagebox.showerror("システムエラー", error_msg))
            self.app.after(0, self._finish_calculation, False)

    def _finish_calculation(self, success):
        """計算完了処理"""
        if hasattr(self, 'progress_window') and self.progress_window.winfo_exists():
            self.progress_window.destroy()
        if hasattr(self.app, 'calc_button'):
            self.app.calc_button.config(state='normal', text="🚀 計算実行", bg=self.app.PRIMARY_BLUE)
        if success:
            messagebox.showinfo("計算完了", "✅ AIが統計分析を完了しました！")

    def _build_sql_query(self, base_sql, inputs):
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

    def _fetch_data(self, cursor, inputs):
        """データの取得"""
        from constants import DEFECT_COLUMNS
        
        data = {'total_qty': 0, 'total_defect': 0, 'defect_rate': 0, 'defect_rates_sorted': [], 'best5': []}
        defect_columns_sum = ", ".join(f"SUM(IIF([{col}] IS NOT NULL AND [{col}]<>0, [{col}], 0))" for col in DEFECT_COLUMNS)
        base_sql = f"SELECT SUM([数量]), SUM([総不具合数]), {defect_columns_sum} FROM t_不具合情報 WHERE [品番] = ?"
        sql, params = self._build_sql_query(base_sql, inputs)
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

    def _calculate_stats(self, db_data, inputs):
        """統計計算"""
        import math
        from scipy.stats import binom
        from constants import InspectionConstants
        
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

    def _update_ui(self, db_data, stats_results, inputs):
        """UI更新"""
        # プログレスウィンドウを閉じる
        self._close_progress_window()
        
        self._clear_previous_results()
        self.last_db_data, self.last_stats_results, self.last_inputs = db_data, stats_results, inputs
        texts = self._generate_result_texts(db_data, stats_results, inputs)
        self._display_main_results(stats_results, texts['advice'])
        self._display_detailed_results(texts)
        
        # 警告メッセージの表示
        if 'warning_message' in stats_results:
            self._display_warning_message(stats_results['warning_message'])

    def _clear_previous_results(self):
        """以前の結果をクリア"""
        for widget_name in ['main_sample_label', 'level_label', 'reason_label', 'advice_label']:
            if hasattr(self.app, widget_name) and (widget := getattr(self.app, widget_name)):
                widget.destroy()
        self.app.review_frame.pack_forget()
        self.app.best3_frame.pack_forget()
        if hasattr(self.app, 'warning_frame'):
            self.app.warning_frame.destroy()
        if hasattr(self.app, 'hide_export_button'):
            self.app.hide_export_button()

    def _format_int(self, n):
        """整数のフォーマット"""
        try:
            return f"{int(n):,}"
        except (ValueError, TypeError):
            return str(n)

    def _generate_result_texts(self, db_data, stats_results, inputs):
        """結果テキストの生成"""
        sample_size_disp = self._format_int(stats_results['sample_size'])
        period_text = f"（{inputs['start_date'] or '最初'}～{inputs['end_date'] or '最新'}）" if inputs['start_date'] or inputs['end_date'] else "（全期間対象）"
        
        review_text = (
            f"【根拠レビュー】\n・ロットサイズ: {self._format_int(inputs['lot_size'])}\n・対象期間: {period_text}\n"
            f"・数量合計: {self._format_int(db_data['total_qty'])}個\n・不具合数合計: {self._format_int(db_data['total_defect'])}個\n"
            f"・不良率: {db_data['defect_rate']:.2f}%\n・信頼度: {inputs['confidence_level']:.1f}%\n・c値: {inputs['c_value']}\n"
            f"・推奨抜取検査数: {sample_size_disp} 個\n（c={inputs['c_value']}, 信頼度={inputs['confidence_level']:.1f}%の条件で自動計算）"
        )
        
        if db_data['best5']:
            best5_text = "【検査時の注意喚起：過去不具合ベスト5】\n"
            for i, (naiyo, count) in enumerate(db_data['best5'], 1):
                rate = next((r for col, r, c_ in db_data['defect_rates_sorted'] if col == naiyo), 0)
                best5_text += f"{i}. {naiyo}（{self._format_int(count)}個, {rate:.2f}%）\n"
        else: 
            best5_text = "【検査時の注意喚起】\n該当期間に不具合データがありません。"
            
        if db_data['best5'] and db_data['best5'][0][1] > 0:
            advice = f"過去最多の不具合は『{db_data['best5'][0][0]}』です。検査時は特にこの点にご注意ください。"
        elif db_data['total_defect'] > 0: 
            advice = "過去の不具合傾向から特に目立つ項目はありませんが、標準的な検査を心がけましょう。"
        else: 
            advice = "過去の不具合データが少ないため、全般的に注意して検査を行ってください。"
            
        return {'review': review_text, 'best5': best5_text, 'advice': advice}

    def _display_main_results(self, stats_results, advice_text):
        """メイン結果の表示"""
        sample_size_disp = self._format_int(stats_results['sample_size'])
        
        # 抜取検査数の表示
        self.app.main_sample_label = tk.Label(
            self.app.result_frame, 
            text=f"抜取検査数: {sample_size_disp} 個", 
            font=("Meiryo", 32, "bold"), 
            fg="#007bff", 
            bg="#e9ecef", 
            pady=10
        )
        self.app.main_sample_label.pack(pady=(10, 0))
        
        # 検査水準の表示
        self.app.level_label = tk.Label(
            self.app.result_frame, 
            text=f"検査水準: {stats_results['level_text']}", 
            font=("Meiryo", 16, "bold"), 
            fg="#2c3e50", 
            bg="#e9ecef", 
            pady=5
        )
        self.app.level_label.pack()
        
        # 根拠の表示
        self.app.reason_label = tk.Label(
            self.app.result_frame, 
            text=f"根拠: {stats_results['level_reason']}", 
            font=("Meiryo", 12), 
            fg="#6c757d", 
            bg="#e9ecef", 
            pady=5, 
            wraplength=800
        )
        self.app.reason_label.pack()
        
        # アドバイスの表示
        self.app.advice_label = tk.Label(
            self.app.sampling_frame, 
            text=advice_text, 
            font=("Meiryo", 9), 
            fg=self.app.WARNING_RED, 
            bg=self.app.LIGHT_GRAY, 
            wraplength=800, 
            justify='left', 
            padx=15, 
            pady=8, 
            relief="flat", 
            bd=1
        )
        self.app.advice_label.pack(after=self.app.result_label, pady=(0, 5))

    def _display_warning_message(self, warning_message):
        """警告メッセージの表示"""
        # 警告フレームの作成
        warning_frame = tk.Frame(
            self.app.sampling_frame, 
            bg="#fff3cd", 
            relief="solid", 
            bd=2
        )
        warning_frame.pack(fill='x', padx=40, pady=(10, 5))
        
        # 警告アイコンとメッセージ
        warning_label = tk.Label(
            warning_frame, 
            text=f"⚠️ 警告: {warning_message}", 
            font=("Meiryo", 10, "bold"), 
            fg="#856404", 
            bg="#fff3cd", 
            wraplength=800, 
            justify='left', 
            padx=15, 
            pady=10
        )
        warning_label.pack()
        
        # 代替案の提案ボタン
        alternatives_button = tk.Button(
            warning_frame, 
            text="💡 代替案を表示", 
            command=lambda: self._show_alternatives(), 
            font=("Meiryo", 9), 
            bg="#ffc107", 
            fg="#212529", 
            relief="flat", 
            padx=10, 
            pady=5
        )
        alternatives_button.pack(pady=(0, 10))
        
        # 警告フレームを保存（後で削除するため）
        self.app.warning_frame = warning_frame

    def _show_alternatives(self):
        """代替案の表示"""
        if not hasattr(self, 'last_inputs') or not self.last_inputs:
            messagebox.showinfo("情報", "先に計算を実行してください。")
            return
        
        # 代替案ダイアログの作成
        dialog = tk.Toplevel(self.app)
        dialog.title("代替案の提案")
        dialog.geometry("600x500")
        dialog.configure(bg="#f8f9fa")
        dialog.resizable(True, True)
        
        # 中央配置
        x = (self.app.winfo_screenwidth() // 2) - 300
        y = (self.app.winfo_screenheight() // 2) - 250
        dialog.geometry(f"600x500+{x}+{y}")
        
        # モーダル表示
        dialog.transient(self.app)
        dialog.grab_set()
        
        # タイトル
        title_label = tk.Label(
            dialog, 
            text="💡 代替案の提案", 
            font=("Meiryo", 16, "bold"), 
            fg="#2c3e50", 
            bg="#f8f9fa"
        )
        title_label.pack(pady=(20, 10))
        
        # 現在の条件表示
        current_frame = tk.LabelFrame(
            dialog, 
            text="現在の条件", 
            font=("Meiryo", 12, "bold"), 
            fg="#2c3e50", 
            bg="#f8f9fa",
            padx=10,
            pady=10
        )
        current_frame.pack(fill='x', padx=20, pady=10)
        
        current_text = f"ロットサイズ: {self._format_int(self.last_inputs['lot_size'])}個\n"
        current_text += f"不良率: {self.last_db_data['defect_rate']:.3f}%\n"
        current_text += f"信頼度: {self.last_inputs['confidence_level']:.1f}%\n"
        current_text += f"c値: {self.last_inputs['c_value']}"
        
        tk.Label(
            current_frame, 
            text=current_text, 
            font=("Meiryo", 10), 
            fg="#495057", 
            bg="#f8f9fa",
            justify='left'
        ).pack(anchor='w')
        
        # 代替案の計算と表示
        alternatives_frame = tk.LabelFrame(
            dialog, 
            text="代替案", 
            font=("Meiryo", 12, "bold"), 
            fg="#2c3e50", 
            bg="#f8f9fa",
            padx=10,
            pady=10
        )
        alternatives_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # スクロール可能なテキストエリア
        text_frame = tk.Frame(alternatives_frame, bg="#f8f9fa")
        text_frame.pack(fill='both', expand=True)
        
        scrollbar = tk.Scrollbar(text_frame)
        text_widget = tk.Text(
            text_frame, 
            font=("Meiryo", 10), 
            bg="#ffffff", 
            fg="#2c3e50",
            wrap=tk.WORD,
            yscrollcommand=scrollbar.set
        )
        scrollbar.config(command=text_widget.yview)
        
        scrollbar.pack(side='right', fill='y')
        text_widget.pack(side='left', fill='both', expand=True)
        
        # 代替案の計算
        alternatives_text = self._calculate_alternatives()
        text_widget.insert('1.0', alternatives_text)
        text_widget.config(state='disabled')
        
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

    def _calculate_alternatives(self):
        """代替案の計算"""
        import math
        from scipy.stats import binom
        
        p = self.last_db_data['defect_rate'] / 100
        lot_size = self.last_inputs['lot_size']
        
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
                    if binom.cdf(c_val, mid, p) >= 1 - self.last_inputs['confidence_level']/100:
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

    def _display_detailed_results(self, texts):
        """詳細結果の表示"""
        # 根拠レビューの表示
        self.app.review_var.set(texts['review'])
        self.app.review_frame.pack(fill='x', padx=40, pady=10)
        
        # 検査時の注意喚起
        self.app.best3_var.set(texts['best5'])
        self.app.best3_frame.pack(fill='x', padx=40, pady=10)
        
        # エクスポートボタンを表示
        if hasattr(self.app, 'show_export_button'): 
            self.app.show_export_button()

    def _show_error(self, title, message):
        """エラーメッセージの表示"""
        self._close_progress_window()
        messagebox.showerror(title, message)

    def _close_progress_window(self):
        """プログレスウィンドウを閉じる"""
        if hasattr(self, 'progress_window') and self.progress_window.winfo_exists():
            self.progress_window.destroy()

    def show_product_numbers_list(self):
        """品番リストの表示（非同期読み込み）"""
        thread = threading.Thread(target=self._load_product_numbers_async)
        thread.daemon = True
        thread.start()

    def _load_product_numbers_async(self):
        """品番リストの非同期読み込み"""
        progress_window = None
        try:
            # プログレスウィンドウの作成
            progress_window = tk.Toplevel(self.app)
            progress_window.title("品番リスト読み込み中...")
            progress_window.geometry("350x120")
            progress_window.configure(bg="#f0f0f0")
            progress_window.resizable(False, False)
            
            # 中央配置
            x = (self.app.winfo_screenwidth() // 2) - 175
            y = (self.app.winfo_screenheight() // 2) - 60
            progress_window.geometry(f"350x120+{x}+{y}")
            
            # プログレスバー
            progress_bar = ttk.Progressbar(progress_window, mode='indeterminate', length=250)
            progress_bar.pack(pady=20)
            progress_bar.start()
            
            # ステータスラベル
            status_label = tk.Label(
                progress_window, 
                text="データベースから品番リストを読み込み中...", 
                font=("Meiryo", 10), 
                bg="#f0f0f0"
            )
            status_label.pack(pady=5)
            
            # 品番リストの取得
            product_numbers = self.db_manager.fetch_all_product_numbers()
            
            # プログレスウィンドウを閉じる
            if progress_window and progress_window.winfo_exists():
                progress_window.destroy()
            
            # 結果の表示
            self.app.after(0, self._show_product_numbers_result, product_numbers)
            
        except Exception as e:
            if progress_window and progress_window.winfo_exists():
                progress_window.destroy()
            error_message = f"品番リストの読み込み中にエラーが発生しました:\n{str(e)}"
            self.app.after(0, lambda: messagebox.showerror("エラー", error_message))

    def _show_product_numbers_result(self, product_numbers):
        """品番リストの結果表示"""
        if not product_numbers:
            messagebox.showinfo("情報", "表示できる品番がありません。")
            return
        
        # 品番リストウィンドウの作成
        win = tk.Toplevel(self.app)
        win.title(f"品番リスト ({len(product_numbers)}件)")
        win.geometry("400x500")
        win.configure(bg="#f0f0f0")
        
        # 検索フレーム
        search_frame = tk.Frame(win, bg="#f0f0f0")
        search_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(search_frame, text="🔍 検索:", font=("Meiryo", 10), bg="#f0f0f0").pack(side='left', padx=(0, 5))
        
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, font=("Meiryo", 10))
        search_entry.pack(fill='x', expand=True)
        
        # リストフレーム
        list_frame = tk.Frame(win, bg="#f0f0f0")
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # スクロールバー
        scrollbar = tk.Scrollbar(list_frame, orient='vertical')
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=("Meiryo", 10))
        scrollbar.config(command=listbox.yview)
        
        scrollbar.pack(side='right', fill='y')
        listbox.pack(side='left', fill='both', expand=True)
        
        # 検索可能なアイテムの準備
        searchable_items = [(pn, pn.lower()) for pn in product_numbers]
        
        # 初期表示
        for pn, _ in searchable_items:
            listbox.insert('end', pn)
        
        # 検索機能
        def update_listbox(*args):
            search_term = search_var.get().strip().lower()
            listbox.delete(0, 'end')
            filtered_count = 0
            
            for pn, pn_lower in searchable_items:
                if not search_term or search_term in pn_lower:
                    listbox.insert('end', pn)
                    filtered_count += 1
            
            win.title(f"品番リスト ({filtered_count}件)")
        
        search_var.trace("w", update_listbox)
        
        # ダブルクリックで選択
        def on_double_click(event):
            selected_indices = listbox.curselection()
            if not selected_indices:
                return
            
            selected_pn = listbox.get(selected_indices[0])
            self.app.sample_pn_entry.delete(0, 'end')
            self.app.sample_pn_entry.insert(0, selected_pn)
            win.destroy()
        
        listbox.bind("<Double-1>", on_double_click)
        
        # モーダル表示
        win.transient(self.app)
        win.grab_set()
        search_entry.focus_set()
        
        # 中央配置
        win.update_idletasks()
        x = (self.app.winfo_screenwidth() // 2) - 200
        y = (self.app.winfo_screenheight() // 2) - 250
        win.geometry(f"400x500+{x}+{y}")
        
        self.app.wait_window(win)

    def export_results(self):
        """結果のエクスポート"""
        if not hasattr(self, 'last_db_data') or not self.last_db_data: 
            messagebox.showinfo("エクスポート不可", "先に計算を実行してください。")
            return
            
        texts = self._generate_result_texts(self.last_db_data, self.last_stats_results, self.last_inputs)
        sample_size_disp = self._format_int(self.last_stats_results['sample_size'])
        
        content = f"""AI SQC Sampler - 計算結果
計算日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
{'=' * 50}

品番: {self.last_inputs['product_number']}
ロットサイズ: {self._format_int(self.last_inputs['lot_size'])}個
不具合率: {self.last_db_data['defect_rate']:.2f}%
検査水準: {self.last_stats_results['level_text']}
サンプルサイズ: {sample_size_disp} 個
信頼度: {self.last_inputs['confidence_level']:.1f}%
c値: {self.last_inputs['c_value']}

{texts['review']}

{texts['best5']}
"""
        
        try:
            filepath = filedialog.asksaveasfilename(
                title="結果を名前を付けて保存",
                defaultextension=".txt",
                filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")],
                initialfile=f"検査結果_{self.last_inputs['product_number']}.txt"
            )
            if not filepath: 
                return
            with open(filepath, 'w', encoding='utf-8') as f: 
                f.write(content)
            messagebox.showinfo("成功", f"結果を保存しました。\nパス: {filepath}")
        except Exception as e:
            messagebox.showerror("エクスポート失敗", f"ファイルの保存中にエラーが発生しました: {e}")

    def open_config_dialog(self):
        """設定ダイアログの表示"""
        dialog = SettingsDialog(self.app, self.config_manager)
        dialog.show()
        
        # 設定変更後、データベースマネージャーを再初期化
        self.db_manager = DatabaseManager(self.config_manager)

    def show_help(self):
        """ヘルプの表示"""
        messagebox.showinfo("ヘルプ", "AI SQC Sampler - 統計的品質管理によるサンプリングサイズ計算アプリケーション")

    def show_about(self):
        """アプリケーション情報の表示"""
        messagebox.showinfo("アプリケーション情報", 
            "AI SQC Sampler v1.0\n\n"
            "統計的品質管理によるサンプリングサイズ計算アプリケーション\n"
            "Microsoft Accessデータベース対応")

    def test_database_connection(self):
        """データベース接続のテスト"""
        success, message = self.db_manager.test_connection()
        if success:
            messagebox.showinfo("データベース接続テスト", f"✅ {message}")
        else:
            messagebox.showerror("データベース接続テスト", f"❌ {message}")


def main():
    """メイン関数"""
    try:
        controller = MainController()
        controller.run()
    except Exception as e:
        messagebox.showerror("アプリケーションエラー", f"アプリケーションの起動中にエラーが発生しました:\n{str(e)}")


if __name__ == "__main__":
    main()