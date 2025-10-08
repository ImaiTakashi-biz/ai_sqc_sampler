"""
UI管理モジュール
ユーザーインターフェースの表示と更新を管理
"""

import tkinter as tk
from tkinter import messagebox
from datetime import datetime


class UIManager:
    """UI管理クラス"""
    
    def __init__(self, app):
        self.app = app
    
    def update_ui(self, db_data, stats_results, inputs):
        """UI更新"""
        self.clear_previous_results()
        texts = self.generate_result_texts(db_data, stats_results, inputs)
        self.display_main_results(stats_results, texts['advice'])
        self.display_detailed_results(texts)
        
        # 警告メッセージの表示
        if 'warning_message' in stats_results:
            self.display_warning_message(stats_results['warning_message'])

    def clear_previous_results(self):
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

    def format_int(self, n):
        """整数のフォーマット"""
        try:
            return f"{int(n):,}"
        except (ValueError, TypeError):
            return str(n)

    def generate_result_texts(self, db_data, stats_results, inputs):
        """結果テキストの生成"""
        sample_size_disp = self.format_int(stats_results['sample_size'])
        period_text = f"（{inputs['start_date'] or '最初'}～{inputs['end_date'] or '最新'}）" if inputs['start_date'] or inputs['end_date'] else "（全期間対象）"
        
        review_text = (
            f"【根拠レビュー】\n・ロットサイズ: {self.format_int(inputs['lot_size'])}\n・対象期間: {period_text}\n"
            f"・数量合計: {self.format_int(db_data['total_qty'])}個\n・不具合数合計: {self.format_int(db_data['total_defect'])}個\n"
            f"・不良率: {db_data['defect_rate']:.2f}%\n・信頼度: {inputs['confidence_level']:.1f}%\n・c値: {inputs['c_value']}\n"
            f"・推奨抜取検査数: {sample_size_disp} 個\n（c={inputs['c_value']}, 信頼度={inputs['confidence_level']:.1f}%の条件で自動計算）"
        )
        
        if db_data['best5']:
            best5_text = "【検査時の注意喚起：過去不具合ベスト5】\n"
            for i, (naiyo, count) in enumerate(db_data['best5'], 1):
                rate = next((r for col, r, c_ in db_data['defect_rates_sorted'] if col == naiyo), 0)
                best5_text += f"{i}. {naiyo}（{self.format_int(count)}個, {rate:.2f}%）\n"
        else: 
            best5_text = "【検査時の注意喚起】\n該当期間に不具合データがありません。"
            
        if db_data['best5'] and db_data['best5'][0][1] > 0:
            advice = f"過去最多の不具合は『{db_data['best5'][0][0]}』です。検査時は特にこの点にご注意ください。"
        elif db_data['total_defect'] > 0: 
            advice = "過去の不具合傾向から特に目立つ項目はありませんが、標準的な検査を心がけましょう。"
        else: 
            advice = "過去の不具合データが少ないため、全般的に注意して検査を行ってください。"
            
        return {'review': review_text, 'best5': best5_text, 'advice': advice}

    def display_main_results(self, stats_results, advice_text):
        """メイン結果の表示"""
        sample_size_disp = self.format_int(stats_results['sample_size'])
        
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

    def display_warning_message(self, warning_message):
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
            command=lambda: self.show_alternatives(), 
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

    def show_alternatives(self):
        """代替案の表示"""
        if not hasattr(self.app.controller, 'last_inputs') or not self.app.controller.last_inputs:
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
        
        current_text = f"ロットサイズ: {self.format_int(self.app.controller.last_inputs['lot_size'])}個\n"
        current_text += f"不良率: {self.app.controller.last_db_data['defect_rate']:.3f}%\n"
        current_text += f"信頼度: {self.app.controller.last_inputs['confidence_level']:.1f}%\n"
        current_text += f"c値: {self.app.controller.last_inputs['c_value']}"
        
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
        alternatives_text = self.app.controller.calculation_engine.calculate_alternatives(
            self.app.controller.last_db_data, 
            self.app.controller.last_inputs
        )
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

    def display_detailed_results(self, texts):
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

