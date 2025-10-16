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

        # 不具合データがない場合の特別処理
        if stats_results.get('no_defect_data', False):
            self.display_no_defect_data_message(stats_results, inputs)
            if hasattr(self.app, 'show_export_button'):
                self.app.show_export_button()
            return

        texts = self.generate_result_texts(db_data, stats_results, inputs)

        self.display_main_results(stats_results, texts['advice'], texts['best5'], inputs.get('product_number', ''))
        

        # テーブル形式のレビュー情報を表示
        adjustment_info = stats_results.get('adjustment_info')
        if 'review_data' in texts:
            self.display_review_table(texts['review_data'], adjustment_info)
        self.display_detailed_results(texts)
        if 'warning_message' in stats_results:
            self.display_warning_message(stats_results['warning_message'])
        if 'guidance_message' in stats_results and stats_results['guidance_message']:
            self.display_guidance_message(stats_results['guidance_message'])
        if hasattr(self.app, 'show_export_button'):
            self.app.show_export_button()

    def clear_previous_results(self):
        """以前の結果をクリア"""
        for widget_name in ['main_sample_label', 'level_label', 'reason_label', 'advice_label', 'product_label']:
            if hasattr(self.app, widget_name) and (widget := getattr(self.app, widget_name)):
                widget.destroy()
        self.app.review_frame.pack_forget()
        self.app.best3_frame.pack_forget()
        if hasattr(self.app, 'warning_frame'):
            self.app.warning_frame.destroy()
        if hasattr(self.app, 'guidance_frame'):
            self.app.guidance_frame.destroy()
        if hasattr(self.app, 'review_table_frame'):
            self.app.review_table_frame.destroy()
        if hasattr(self.app, 'section_divider'):
            self.app.section_divider.pack_forget()
        if hasattr(self.app, 'section_label'):
            self.app.section_label.pack_forget()
        if hasattr(self.app, 'result_frame'):
            self.app.result_frame.pack_forget()
        if hasattr(self.app, 'hide_export_button'):
            self.app.hide_export_button()

    def format_int(self, n):

        """整数のフォーマット"""

        try:

            return f"{int(n):,}"

        except (ValueError, TypeError):

            return str(n)



    def generate_result_texts(self, db_data, stats_results, inputs):

        """結果テキストの生成（AQL/LTPD設計対応）"""

        sample_size_disp = self.format_int(stats_results['sample_size'])

        period_text = f"（{inputs['start_date'] or '最初'}〜{inputs['end_date'] or '最新'}）" if inputs['start_date'] or inputs['end_date'] else "（全期間対象）"

        

        # AQL/LTPD設計の情報を表示（調整後の値を使用）

        aql = stats_results.get('aql', inputs.get('aql', 0.25))

        ltpd = stats_results.get('ltpd', inputs.get('ltpd', 1.0))

        alpha = inputs.get('alpha', 5.0)

        beta = inputs.get('beta', 10.0)

        c_value = inputs.get('c_value', 0)

        

        # 調整情報があるかチェック

        has_adjustment = 'adjustment_info' in stats_results and stats_results['adjustment_info']

        original_aql = stats_results.get('original_aql', aql)

        original_ltpd = stats_results.get('original_ltpd', ltpd)

        

        # ロットサイズに基づく計算方法の説明

        lot_size = inputs['lot_size']

        if lot_size <= 50:

            calculation_method = "小ロット（高割合抜取・全数検査）"

        elif lot_size <= 500:

            calculation_method = "中ロット（有限母集団補正・超幾何分布）"

        else:

            calculation_method = "大ロット（有限母集団補正・超幾何分布）"

        

        # テーブル形式の結果データを準備

        review_data = {

            'title': '【AQL/LTPD設計による根拠レビュー' + ('（データベース実績活用）' if has_adjustment else '') + '】',

            'basic_info': [

                ('ロットサイズ', f"{self.format_int(inputs['lot_size'])}個（{calculation_method}）"),

                ('対象期間', period_text),

                ('数量合計', f"{self.format_int(db_data['total_qty'])}個"),

                ('不具合数合計', f"{self.format_int(db_data['total_defect'])}個"),

                ('実績不良率', f"{db_data['defect_rate']:.2f}%")

            ],

            'parameters': [

                ('AQL（合格品質水準）', f"{original_aql}% → {aql}%" + ('（実績に基づく調整）' if has_adjustment else '')),

                ('LTPD（不合格品質水準）', f"{original_ltpd}% → {ltpd}%" + ('（実績に基づく調整）' if has_adjustment else '')),

                ('α（生産者危険）', f"{alpha}%"),

                ('β（消費者危険）', f"{beta}%"),

                ('c値（許容不良数）', f"{c_value}"),

                ('推奨抜取検査数', f"{sample_size_disp} 個")

            ],

            'calculation_note': f"（{'調整後' if has_adjustment else ''}AQL={aql}%, LTPD={ltpd}%, α={alpha}%, β={beta}%, c={c_value}の条件で自動計算）"

        }

        if has_adjustment:

            review_text = (

                f"【AQL/LTPD設計による根拠レビュー（データベース実績活用）】\n・ロットサイズ: {self.format_int(inputs['lot_size'])}個（{calculation_method}）\n・対象期間: {period_text}\n"

                f"・数量合計: {self.format_int(db_data['total_qty'])}個\n・不具合数合計: {self.format_int(db_data['total_defect'])}個\n"

                f"・実績不良率: {db_data['defect_rate']:.2f}%\n"

                f"・AQL（合格品質水準）: {original_aql}% → {aql}%（実績に基づく調整）\n"

                f"・LTPD（不合格品質水準）: {original_ltpd}% → {ltpd}%（実績に基づく調整）\n"

                f"・α（生産者危険）: {alpha}%\n・β（消費者危険）: {beta}%\n・c値（許容不良数）: {c_value}\n"

                f"・推奨抜取検査数: {sample_size_disp} 個\n（調整後AQL={aql}%, LTPD={ltpd}%, α={alpha}%, β={beta}%, c={c_value}の条件で自動計算）"

            )

        else:

            review_text = (

                f"【AQL/LTPD設計による根拠レビュー】\n・ロットサイズ: {self.format_int(inputs['lot_size'])}個（{calculation_method}）\n・対象期間: {period_text}\n"

                f"・数量合計: {self.format_int(db_data['total_qty'])}個\n・不具合数合計: {self.format_int(db_data['total_defect'])}個\n"

                f"・不良率: {db_data['defect_rate']:.2f}%\n"

                f"・AQL（合格品質水準）: {aql}%\n・LTPD（不合格品質水準）: {ltpd}%\n"

                f"・α（生産者危険）: {alpha}%\n・β（消費者危険）: {beta}%\n・c値（許容不良数）: {c_value}\n"

                f"・推奨抜取検査数: {sample_size_disp} 個\n（AQL={aql}%, LTPD={ltpd}%, α={alpha}%, β={beta}%, c={c_value}の条件で自動計算）"

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

            

        return {

            'review': review_text,

            'review_data': review_data,  # テーブル形式のデータ

            'best5': best5_text,

            'advice': advice

        }



    def _add_table_section(self, parent, title, rows, review_bg, header_bg, head_fg, body_fg):
        section = tk.LabelFrame(parent, text=title, font=("Meiryo", 10, "bold"), fg=head_fg, bg=review_bg, labelanchor='nw')
        section.pack(fill='x', padx=12, pady=6)
        if not rows:
            tk.Label(section, text="データがありません", font=("Meiryo", 10), fg=body_fg, bg=review_bg, anchor='w').pack(fill='x', padx=10, pady=4)
            return
        table = tk.Frame(section, bg=review_bg)
        table.pack(fill='x', padx=10, pady=4)
        header = tk.Frame(table, bg=header_bg)
        header.pack(fill='x')
        tk.Label(header, text="項目", font=("Meiryo", 10, "bold"), fg=head_fg, bg=header_bg, width=18, anchor='w').pack(side='left', padx=(0, 6))
        tk.Label(header, text="内容", font=("Meiryo", 10, "bold"), fg=head_fg, bg=header_bg, anchor='w').pack(side='left', fill='x', expand=True)
        for item, value in rows:
            row_frame = tk.Frame(table, bg=review_bg)
            row_frame.pack(fill='x', pady=2)
            tk.Label(row_frame, text=item, font=("Meiryo", 10), fg=body_fg, bg=review_bg, width=18, anchor='w').pack(side='left', padx=(0, 6))
            tk.Label(row_frame, text=value, font=("Meiryo", 10), fg=body_fg, bg=review_bg, anchor='w', justify='left', wraplength=620).pack(side='left', fill='x', expand=True)

    def _parse_adjustment_rows(self, adjustment_info):
        if not adjustment_info:
            return None, []
        lines = [line.strip() for line in adjustment_info.splitlines() if line.strip()]
        title = None
        rows = []
        for line in lines:
            if line.startswith("【") and line.endswith("】"):
                title = line
                continue
            cleaned = line.lstrip('• ').replace('：', ':')
            if ':' in cleaned:
                key, value = cleaned.split(':', 1)
            else:
                key, value = "備考", cleaned
            value = value.strip()
            if not value:
                value = '—'
            rows.append((key.strip(), value))
        return title, rows

    def display_review_table(self, review_data, adjustment_info=None):
        review_bg = "#e8f4ff"
        header_bg = "#b5d4ff"
        header_fg = "#2c3e50"
        body_fg = "#34495e"

        frame = tk.Frame(self.app.sampling_frame, bg=review_bg, relief="solid", bd=1)
        frame.pack(fill='x', padx=40, pady=(10, 5))

        tk.Label(frame, text=review_data['title'], font=("Meiryo", 11, "bold"), fg="#2c3e50", bg=review_bg).pack(pady=(10, 5))

        self._add_table_section(frame, "【基本情報】", review_data.get('basic_info', []), review_bg, header_bg, header_fg, body_fg)
        self._add_table_section(frame, "【AQL/LTPD設計パラメータ】", review_data.get('parameters', []), review_bg, header_bg, header_fg, body_fg)

        adj_title, adj_rows = self._parse_adjustment_rows(adjustment_info)
        if adj_rows:
            self._add_table_section(frame, adj_title or "【データベース実績活用】", adj_rows, review_bg, header_bg, header_fg, body_fg)

        tk.Label(frame, text=review_data['calculation_note'], font=("Meiryo", 9), fg="#6c757d", bg=review_bg, anchor='w', justify='left').pack(fill='x', padx=12, pady=(5, 10))

        self.app.review_table_frame = frame

    def display_no_defect_data_message(self, stats_results, inputs):
        """不具合データがない場合のメッセージ表示"""
        
        # 1. セクション区切りとタイトルを表示
        if hasattr(self.app, 'section_divider'):
            self.app.section_divider.pack(fill='x', pady=(20, 8))
        if hasattr(self.app, 'section_label'):
            self.app.section_label.pack(pady=(0, 15))
        
        # 品番の表示
        product_number = inputs.get('product_number', '')
        if hasattr(self.app, 'product_label'):
            self.app.product_label.destroy()
        
        self.app.product_label = tk.Label(
            self.app.sampling_frame,
            text=f"品番: {product_number}",
            font=("Meiryo", 18, "bold"),
            fg="#2c3e50",
            bg=self.app.LIGHT_GRAY,
            pady=5
        )
        self.app.product_label.pack(pady=(0, 10))
        
        # 全数検査の表示
        lot_size = inputs.get('lot_size', 1000)
        sample_size_disp = self.format_int(lot_size)
        
        self.app.main_sample_label = tk.Label(
            self.app.sampling_frame, 
            text=f"全数検査: {sample_size_disp} 個", 
            font=("Meiryo", 32, "bold"), 
            fg="#dc3545",  # 赤色で警告表示
            bg=self.app.LIGHT_GRAY, 
            pady=10
        )
        self.app.main_sample_label.pack(pady=(0, 15))
        
        # 警告メッセージの表示
        self.app.level_label = tk.Label(
            self.app.sampling_frame,
            text="⚠️ 不具合データ（実績）がありません",
            font=("Meiryo", 16, "bold"),
            fg="#dc3545",
            bg=self.app.LIGHT_GRAY,
            pady=5
        )
        self.app.level_label.pack(pady=(0, 10))
        
        # 推奨理由の表示
        self.app.reason_label = tk.Label(
            self.app.sampling_frame,
            text=stats_results.get('comment', ''),
            font=("Meiryo", 12),
            fg="#6c757d",
            bg=self.app.LIGHT_GRAY,
            wraplength=600,
            justify='left'
        )
        self.app.reason_label.pack(pady=(0, 15))
        
        # ガイダンスメッセージの表示
        if 'guidance_message' in stats_results:
            self.app.advice_label = tk.Label(
                self.app.sampling_frame,
                text=stats_results['guidance_message'],
                font=("Meiryo", 11, "bold"),
                fg="#dc3545",
                bg=self.app.LIGHT_GRAY,
                wraplength=600,
                justify='left'
            )
            self.app.advice_label.pack(pady=(0, 15))

    def display_main_results(self, stats_results, advice_text, best5_text, product_number=''):
        """メイン結果の表示"""

        # 1. セクション区切りとタイトルを表示
        if hasattr(self.app, 'section_divider'):
            self.app.section_divider.pack(fill='x', pady=(20, 8))
        if hasattr(self.app, 'section_label'):
            self.app.section_label.pack(pady=(0, 15))
        
        # 1.5 品番の表示
        if not product_number:
            product_number = stats_results.get('product_number', '')
        if not product_number:
            last_inputs = getattr(getattr(self.app, 'controller', None), 'last_inputs', {}) or {}
            product_number = last_inputs.get('product_number', '')
        product_number = product_number or ''
        
        if hasattr(self.app, 'product_label'):
            self.app.product_label.destroy()
        
        self.app.product_label = tk.Label(
            self.app.sampling_frame,
            text=f"品番: {product_number}",
            font=("Meiryo", 18, "bold"),
            fg="#2c3e50",
            bg=self.app.LIGHT_GRAY,
            pady=5
        )
        self.app.product_label.pack(pady=(0, 10))
        

        sample_size_disp = self.format_int(stats_results['sample_size'])

        

        # 2. 抜取検査数の表示

        self.app.main_sample_label = tk.Label(

            self.app.sampling_frame, 

            text=f"抜取検査数: {sample_size_disp} 個", 

            font=("Meiryo", 32, "bold"), 

            fg="#007bff", 

            bg=self.app.LIGHT_GRAY, 

            pady=10

        )

        self.app.main_sample_label.pack(pady=(10, 0))

        

        # 3. アドバイス（過去最多の不具合）の表示（文字サイズを2サイズ大きく）

        self.app.advice_label = tk.Label(

            self.app.sampling_frame, 

            text=advice_text, 

            font=("Meiryo", 11),  # 9 → 11に変更（2サイズ大きく）

            fg=self.app.WARNING_RED, 

            bg=self.app.LIGHT_GRAY, 

            wraplength=800, 

            justify='left', 

            padx=15, 

            pady=8, 

            relief="flat", 

            bd=1

        )

        self.app.advice_label.pack(pady=(0, 5))

        # 4. best5 notice panel beneath advice

        if hasattr(self.app, 'best3_var') and hasattr(self.app, 'best3_frame'):

            self.app.best3_var.set(best5_text)

            padx = getattr(self.app, 'PADDING_X_MEDIUM', 40)

            pady = getattr(self.app, 'PADDING_Y_SMALL', 10)

            self.app.best3_frame.pack(fill='x', padx=padx, pady=pady)

        

        # 5. display inspection level
        self.app.level_label = tk.Label(
            self.app.sampling_frame, 
            text=f"検査水準：{stats_results['level_text']}", 
            font=("Meiryo", 16, "bold"), 
            fg="#2c3e50", 
            bg=self.app.LIGHT_GRAY, 
            pady=5
        )
        self.app.level_label.pack()

        # コメントの表示（条件）
        self.app.reason_label = tk.Label(
            self.app.sampling_frame,
            text=f"コメント：{stats_results['level_reason']}",
            font=("Meiryo", 12),
            fg="#6c757d",
            bg=self.app.LIGHT_GRAY,
            pady=5,
            wraplength=800,
            justify='left'
        )
        self.app.reason_label.pack(pady=(0, 5))

        # 6. display optional action buttons beneath the comment
        if hasattr(self.app, 'oc_curve_button'):
            self.app.oc_curve_button.pack_forget()
            self.app.oc_curve_button.pack(pady=(5, 0))

        if hasattr(self.app, 'inspection_level_button'):
            self.app.inspection_level_button.pack_forget()
            self.app.inspection_level_button.pack(pady=(2, 0))



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

            text=f"⚠ 警告: {warning_message}", 

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

    

    def display_guidance_message(self, guidance_message):

        """ガイダンスメッセージの表示"""

        # ガイダンスフレームの作成

        guidance_frame = tk.Frame(

            self.app.sampling_frame, 

            bg="#e7f3ff", 

            relief="solid", 

            bd=2

        )

        guidance_frame.pack(fill='x', padx=40, pady=(10, 5))

        

        # ガイダンスアイコンとメッセージ

        guidance_label = tk.Label(

            guidance_frame, 

            text=f"ℹ ガイダンス: {guidance_message}", 

            font=("Meiryo", 10, "bold"), 

            fg="#004085", 

            bg="#e7f3ff", 

            wraplength=800, 

            justify='left', 

            padx=15, 

            pady=10

        )

        guidance_label.pack()

        

        # ガイダンスフレームを保存（後で削除するため）

        self.app.guidance_frame = guidance_frame

    

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

        current_text += f"AQL: {self.app.controller.last_inputs.get('aql', 0.25)}%\n"

        current_text += f"LTPD: {self.app.controller.last_inputs.get('ltpd', 1.0)}%\n"

        current_text += f"α（生産者危険）: {self.app.controller.last_inputs.get('alpha', 5.0)}%\n"

        current_text += f"β（消費者危険）: {self.app.controller.last_inputs.get('beta', 10.0)}%\n"

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

        # 根拠レビューの表示は削除（テーブル表示で代替）

        # self.app.review_var.set(texts['review'])

        # self.app.review_frame.pack(fill='x', padx=40, pady=10)

        
        # 検査時の注意喚起（テキストのみ同期）
        if hasattr(self.app, 'best3_var'):
            self.app.best3_var.set(texts['best5'])


