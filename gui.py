import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from datetime import datetime
import platform

class App(tk.Tk):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller

        # --- 変数定義 ---
        self.result_var = tk.StringVar()
        self.review_var = tk.StringVar()
        self.best3_var = tk.StringVar()
        
        # --- カラーパレット ---
        self.PRIMARY_BLUE = "#3498db"
        self.ACCENT_BLUE = "#2980b9"
        self.LIGHT_GRAY = "#ecf0f1"
        self.MEDIUM_GRAY = "#bdc3c7"
        self.DARK_GRAY = "#34495e"
        self.WARNING_RED = "#e74c3c"
        self.INFO_GREEN = "#2ecc71"

        # --- UI定数（最適化版） ---
        self.FONT_FAMILY = "Meiryo"
        self.FONT_SIZE_LARGE = 16
        self.FONT_SIZE_MEDIUM = 12
        self.FONT_SIZE_SMALL = 10
        self.FONT_SIZE_XSMALL = 9
        self.PADDING_X_LARGE = 30  # 40→30に削減
        self.PADDING_X_MEDIUM = 15  # 20→15に削減
        self.PADDING_X_SMALL = 10   # 15→10に削減
        self.PADDING_Y_LARGE = 12   # 15→12に削減
        self.PADDING_Y_MEDIUM = 8   # 10→8に削減
        self.PADDING_Y_SMALL = 4    # 5→4に削減
        self.WRAPLENGTH_DEFAULT = 800

        # --- ウィジェット変数 ---
        self.export_button = None
        self.export_frame = None

        # --- ウィンドウ設定 ---
        self.title("抜取検査数計算ツール - AIアシスト")
        self.geometry("1000x700")
        self.configure(bg=self.LIGHT_GRAY)
        try:
            self.state('zoomed')
        except tk.TclError:
            self._center_window()
        self.bind('<Configure>', self._on_resize)
        
        # --- メニューバーの作成 ---
        self._create_menu_bar()

        # --- ウィジェットの構築 ---
        self._create_widgets()

    def _center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _on_resize(self, event):
        if self.state() == 'zoomed':
            self._center_window()
    
    def _create_menu_bar(self):
        """メニューバーの作成"""
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        
        # ファイルメニュー
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ファイル", menu=file_menu)
        file_menu.add_command(label="設定...", command=self.controller.open_config_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self.quit)
        
        # ツールメニュー
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ツール", menu=tools_menu)
        tools_menu.add_command(label="データベース接続テスト", command=self.controller.test_database_connection)
        tools_menu.add_command(label="品番リスト表示", command=self.controller.show_product_numbers_list)
        
        # ヘルプメニュー
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ヘルプ", menu=help_menu)
        help_menu.add_command(label="ヘルプ", command=self.controller.show_help)
        help_menu.add_command(label="アプリケーション情報", command=self.controller.show_about)

    def _create_widgets(self):
        canvas_frame = tk.Frame(self)
        canvas_frame.pack(fill="both", expand=True)
        yscroll = tk.Scrollbar(canvas_frame, orient='vertical')
        yscroll.grid(row=0, column=1, sticky='ns')
        xscroll = tk.Scrollbar(canvas_frame, orient='horizontal')
        xscroll.grid(row=1, column=0, sticky='ew')
        main_canvas = tk.Canvas(canvas_frame, bg="#ffffff", highlightthickness=0)
        main_canvas.grid(row=0, column=0, sticky='nsew')
        main_canvas.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        yscroll.configure(command=main_canvas.yview)
        xscroll.configure(command=main_canvas.xview)
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)
        main_frame = tk.Frame(main_canvas, bg="#ffffff")
        main_frame_window = main_canvas.create_window((0, 0), window=main_frame, anchor='nw')

        def on_frame_configure(event):
            main_canvas.config(scrollregion=main_canvas.bbox('all'))
            canvas_width = main_canvas.winfo_width()
            frame_width = main_frame.winfo_reqwidth()
            x_pos = max((canvas_width - frame_width) / 2, 0)
            main_canvas.coords(main_frame_window, x_pos, 0)

        main_frame.bind('<Configure>', on_frame_configure)
        main_canvas.bind('<Configure>', lambda e: on_frame_configure(e))

        if platform.system() == 'Windows':
            main_canvas.bind_all('<MouseWheel>', lambda e: main_canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))
        else:
            main_canvas.bind_all('<Button-4>', lambda e: main_canvas.yview_scroll(-1, 'units'))
            main_canvas.bind_all('<Button-5>', lambda e: main_canvas.yview_scroll(1, 'units'))

        header_frame = tk.Frame(main_frame, bg=self.PRIMARY_BLUE, height=60)  # 80→60に削減
        header_frame.pack(fill='x', pady=(self.PADDING_Y_SMALL, self.PADDING_Y_SMALL))  # 上部パディング削減
        header_frame.pack_propagate(False)
        tk.Label(header_frame, text="🤖 AI抜取検査数計算ツール", font=(self.FONT_FAMILY, self.FONT_SIZE_LARGE, "bold"), fg="#ffffff", bg=self.PRIMARY_BLUE).pack(expand=True)

        summary_frame = tk.Frame(main_frame, bg="#e9ecef", relief="flat", bd=1)
        summary_frame.pack(fill='x', pady=(0, self.PADDING_Y_SMALL), padx=self.PADDING_X_LARGE)  # 下部パディング削減
        summary_text = (
            "【このツールの計算方法】\n"
            "本ツールは統計的品質管理（SQC）のAQL/LTPD設計を基盤としたハイブリッド方式を採用しています。\n"
            "ロットサイズに応じて抜取数を動的調整：小ロットは高割合抜取、中・大ロットは有限母集団補正により実務運用を最適化します。\n"
        )
        tk.Label(summary_frame, text=summary_text, fg=self.DARK_GRAY, bg=self.LIGHT_GRAY, font=(self.FONT_FAMILY, self.FONT_SIZE_SMALL), wraplength=950, anchor='w', justify='left', padx=self.PADDING_X_SMALL, pady=self.PADDING_Y_SMALL).pack(fill='x')  # パディング削減

        self.sampling_frame = tk.Frame(main_frame, bg=self.LIGHT_GRAY, relief="flat", bd=2)
        self.sampling_frame.pack(fill='both', expand=True, padx=self.PADDING_X_LARGE, pady=(0, self.PADDING_Y_MEDIUM))  # 下部パディング削減
        tk.Label(self.sampling_frame, text="📊 抜取検査数計算", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM, "bold"), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(pady=(self.PADDING_Y_SMALL, self.PADDING_Y_MEDIUM))  # パディング削減

        input_frame = tk.Frame(self.sampling_frame, bg=self.LIGHT_GRAY)
        input_frame.pack(fill='x', padx=self.PADDING_X_MEDIUM, pady=self.PADDING_Y_MEDIUM)  # パディング削減
        
        row1_frame = tk.Frame(input_frame, bg=self.LIGHT_GRAY)
        row1_frame.pack(fill='x', pady=self.PADDING_Y_SMALL)
        tk.Label(row1_frame, text="品番:", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left', padx=(0, self.PADDING_Y_SMALL))
        self.sample_pn_entry = tk.Entry(row1_frame, width=20, font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), bg="#ffffff", fg=self.DARK_GRAY, relief="flat", bd=1, highlightthickness=1, highlightbackground=self.MEDIUM_GRAY, highlightcolor=self.PRIMARY_BLUE)
        self.sample_pn_entry.pack(side='left', padx=self.PADDING_Y_SMALL)
        tk.Button(row1_frame, text="品番リスト", font=(self.FONT_FAMILY, self.FONT_SIZE_XSMALL), command=self.controller.show_product_numbers_list, bg=self.MEDIUM_GRAY, fg=self.DARK_GRAY, relief="flat").pack(side='left', padx=(5, 0))
        tk.Label(row1_frame, text="数量 (ロットサイズ):", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left', padx=(self.PADDING_Y_MEDIUM, self.PADDING_Y_SMALL))
        self.sample_qty_entry = tk.Entry(row1_frame, width=12, font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), bg="#ffffff", fg=self.DARK_GRAY, relief="flat", bd=1, highlightthickness=1, highlightbackground=self.MEDIUM_GRAY, highlightcolor=self.PRIMARY_BLUE)
        self.sample_qty_entry.pack(side='left', padx=self.PADDING_Y_SMALL)

        # AQL/LTPD設計の入力項目
        row2_frame = tk.Frame(input_frame, bg=self.LIGHT_GRAY)
        row2_frame.pack(fill='x', pady=self.PADDING_Y_SMALL)
        tk.Label(row2_frame, text="AQL(%):", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left', padx=(0, self.PADDING_Y_SMALL))
        self.sample_aql_entry = tk.Entry(row2_frame, width=6, font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), bg="#ffffff", fg=self.DARK_GRAY, relief="flat", bd=1, highlightthickness=1, highlightbackground=self.MEDIUM_GRAY, highlightcolor=self.PRIMARY_BLUE)
        self.sample_aql_entry.pack(side='left', padx=self.PADDING_Y_SMALL)
        self.sample_aql_entry.insert(0, "0.25")
        
        tk.Label(row2_frame, text="LTPD(%):", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left', padx=(self.PADDING_Y_MEDIUM, self.PADDING_Y_SMALL))
        self.sample_ltpd_entry = tk.Entry(row2_frame, width=6, font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), bg="#ffffff", fg=self.DARK_GRAY, relief="flat", bd=1, highlightthickness=1, highlightbackground=self.MEDIUM_GRAY, highlightcolor=self.PRIMARY_BLUE)
        self.sample_ltpd_entry.pack(side='left', padx=self.PADDING_Y_SMALL)
        self.sample_ltpd_entry.insert(0, "1.0")

        row3_frame = tk.Frame(input_frame, bg=self.LIGHT_GRAY)
        row3_frame.pack(fill='x', pady=self.PADDING_Y_SMALL)
        tk.Label(row3_frame, text="α(生産者危険,%):", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left', padx=(0, self.PADDING_Y_SMALL))
        self.sample_alpha_entry = tk.Entry(row3_frame, width=6, font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), bg="#ffffff", fg=self.DARK_GRAY, relief="flat", bd=1, highlightthickness=1, highlightbackground=self.MEDIUM_GRAY, highlightcolor=self.PRIMARY_BLUE)
        self.sample_alpha_entry.pack(side='left', padx=self.PADDING_Y_SMALL)
        self.sample_alpha_entry.insert(0, "5.0")
        
        tk.Label(row3_frame, text="β(消費者危険,%):", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left', padx=(self.PADDING_Y_MEDIUM, self.PADDING_Y_SMALL))
        self.sample_beta_entry = tk.Entry(row3_frame, width=6, font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), bg="#ffffff", fg=self.DARK_GRAY, relief="flat", bd=1, highlightthickness=1, highlightbackground=self.MEDIUM_GRAY, highlightcolor=self.PRIMARY_BLUE)
        self.sample_beta_entry.pack(side='left', padx=self.PADDING_Y_SMALL)
        self.sample_beta_entry.insert(0, "10.0")

        row4_frame = tk.Frame(input_frame, bg=self.LIGHT_GRAY)
        row4_frame.pack(fill='x', pady=self.PADDING_Y_SMALL)
        tk.Label(row4_frame, text="c値(許容不良数):", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left', padx=(0, self.PADDING_Y_SMALL))
        self.sample_c_entry = tk.Entry(row4_frame, width=6, font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), bg="#ffffff", fg=self.DARK_GRAY, relief="flat", bd=1, highlightthickness=1, highlightbackground=self.MEDIUM_GRAY, highlightcolor=self.PRIMARY_BLUE)
        self.sample_c_entry.pack(side='left', padx=self.PADDING_Y_SMALL)
        self.sample_c_entry.insert(0, "0")

        tk.Label(input_frame, text="※ AQL: これ以下なら合格とみなす不良率（例: 0.25%）\n※ LTPD: これ以上なら不合格にしたい不良率（例: 1.0%）\n※ α: 良いロットを誤って不合格にする確率（例: 5%）\n※ β: 悪いロットを誤って合格にする確率（例: 10%）\n※ c値: 抜取検査で許容できる不良品の最大数（例: c=0なら不良品が1つでも見つかれば不合格）", fg=self.DARK_GRAY, bg=self.LIGHT_GRAY, font=(self.FONT_FAMILY, self.FONT_SIZE_SMALL), wraplength=self.WRAPLENGTH_DEFAULT, justify='left').pack(pady=(self.PADDING_Y_SMALL, 0))  # 下部パディング削除

        row5_frame = tk.Frame(input_frame, bg=self.LIGHT_GRAY)
        row5_frame.pack(fill='x', pady=self.PADDING_Y_SMALL)
        tk.Label(row5_frame, text="対象日（開始）:", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left', padx=(0, self.PADDING_Y_SMALL))
        self.sample_start_date_entry = DateEntry(row5_frame, width=12, date_pattern='yyyy-mm-dd', font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), bg=self.LIGHT_GRAY, fg=self.DARK_GRAY, relief="flat", bd=1, highlightthickness=1, highlightbackground=self.MEDIUM_GRAY, highlightcolor=self.PRIMARY_BLUE, showweeknumbers=False, locale='ja_JP')
        self.sample_start_date_entry.pack(side='left', padx=self.PADDING_Y_SMALL)
        self.sample_start_date_entry.delete(0, 'end')
        tk.Button(row5_frame, text="今日", font=(self.FONT_FAMILY, self.FONT_SIZE_XSMALL), command=lambda: self._set_today_date(self.sample_start_date_entry), bg=self.INFO_GREEN, fg="#ffffff", relief="flat").pack(side='left', padx=(2, 2))
        tk.Button(row5_frame, text="クリア", font=(self.FONT_FAMILY, self.FONT_SIZE_XSMALL), command=lambda: self.sample_start_date_entry.delete(0, 'end'), bg=self.MEDIUM_GRAY, fg=self.DARK_GRAY, relief="flat").pack(side='left', padx=(2, self.PADDING_Y_MEDIUM))
        tk.Label(row5_frame, text="～", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left')
        self.sample_end_date_entry = DateEntry(row5_frame, width=12, date_pattern='yyyy-mm-dd', font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), bg=self.LIGHT_GRAY, fg=self.DARK_GRAY, relief="flat", bd=1, highlightthickness=1, highlightbackground=self.MEDIUM_GRAY, highlightcolor=self.PRIMARY_BLUE, showweeknumbers=False, locale='ja_JP')
        self.sample_end_date_entry.pack(side='left', padx=self.PADDING_Y_SMALL)
        self.sample_end_date_entry.delete(0, 'end')
        tk.Button(row5_frame, text="今日", font=(self.FONT_FAMILY, self.FONT_SIZE_XSMALL), command=lambda: self._set_today_date(self.sample_end_date_entry), bg=self.INFO_GREEN, fg="#ffffff", relief="flat").pack(side='left', padx=(2, 2))
        tk.Button(row5_frame, text="クリア", font=(self.FONT_FAMILY, self.FONT_SIZE_XSMALL), command=lambda: self.sample_end_date_entry.delete(0, 'end'), bg=self.MEDIUM_GRAY, fg=self.DARK_GRAY, relief="flat").pack(side='left', padx=(2, self.PADDING_Y_MEDIUM))
        tk.Label(input_frame, text="※ 対象日を未入力の場合は全期間が対象となります。", fg=self.DARK_GRAY, bg=self.LIGHT_GRAY, font=(self.FONT_FAMILY, self.FONT_SIZE_SMALL)).pack(pady=self.PADDING_Y_SMALL)

        button_frame = tk.Frame(input_frame, bg=self.LIGHT_GRAY)
        button_frame.pack(fill='x', pady=self.PADDING_Y_MEDIUM)  # パディング削減
        
        # 計算実行ボタンの上にコメントを表示
        self.input_comment_label = tk.Label(button_frame, text="品番・数量・（任意で対象日）を入力して計算実行ボタンを押してください。", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM, "bold"), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY, wraplength=self.WRAPLENGTH_DEFAULT, justify='center')
        self.input_comment_label.pack(pady=(0, self.PADDING_Y_SMALL))
        
        self.calc_button = tk.Button(button_frame, text="🚀 計算実行", command=self.controller.start_calculation_thread, font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM, "bold"), bg=self.PRIMARY_BLUE, fg="#ffffff", relief="flat", padx=30, pady=self.PADDING_Y_SMALL, cursor="hand2", activebackground=self.ACCENT_BLUE, activeforeground="#ffffff")  # パディング削減
        self.calc_button.pack()
        
        # 追加機能ボタン（sampling_frame内に直接作成）
        self.oc_curve_button = tk.Button(self.sampling_frame, text="📊 OCカーブ表示", command=self.controller.show_oc_curve, font=(self.FONT_FAMILY, self.FONT_SIZE_SMALL), bg=self.INFO_GREEN, fg="#ffffff", relief="flat", padx=15, pady=5, cursor="hand2")
        
        self.inspection_level_button = tk.Button(self.sampling_frame, text="📋 検査水準管理", command=self.controller.show_inspection_level, font=(self.FONT_FAMILY, self.FONT_SIZE_SMALL), bg="#ffc107", fg="#212529", relief="flat", padx=15, pady=5, cursor="hand2")
        
        # 初期状態では非表示
        self.oc_curve_button.pack_forget()
        self.inspection_level_button.pack_forget()

        # セクション区切り（計算実行ボタンの下に配置）
        self.section_divider = tk.Frame(self.sampling_frame, bg="#dee2e6", height=4, relief="flat")
        self.section_divider.pack(fill='x', pady=(20, 8))
        
        self.section_label = tk.Label(self.sampling_frame, text="📈 統計的品質管理 サンプリング結果", 
                                    font=(self.FONT_FAMILY, 12, "bold"), fg="#2c3e50", bg=self.LIGHT_GRAY)
        self.section_label.pack(pady=(0, 15))
        
        # 初期状態では非表示
        self.section_divider.pack_forget()
        self.section_label.pack_forget()

        self.export_frame = tk.Frame(self.sampling_frame, bg=self.LIGHT_GRAY)
        self.export_button = tk.Button(self.export_frame, text="📄 結果をテキストファイルに保存", command=self.controller.export_results, font=(self.FONT_FAMILY, self.FONT_SIZE_SMALL), bg=self.INFO_GREEN, fg="#ffffff", relief="flat", padx=15, pady=5, cursor="hand2", activebackground=self.ACCENT_BLUE)
        self.export_button.pack()
        self.export_frame.pack_forget()

        self.result_frame = tk.Frame(self.sampling_frame, bg=self.LIGHT_GRAY, relief="flat", bd=1)
        self.result_frame.pack(fill='x', padx=self.PADDING_X_MEDIUM, pady=self.PADDING_Y_MEDIUM)  # パディング削減
        self.result_label = tk.Label(self.result_frame, textvariable=self.result_var, font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM, "bold"), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY, padx=self.PADDING_X_SMALL, pady=self.PADDING_Y_MEDIUM, wraplength=self.WRAPLENGTH_DEFAULT, justify='center')  # パディング削減
        self.result_label.pack(fill='x')
        self.result_var.set("")  # 初期状態では空

        self.review_frame = tk.Frame(self.sampling_frame, bg=self.LIGHT_GRAY, relief="flat", bd=1)
        self.review_frame.pack(fill='x', padx=self.PADDING_X_MEDIUM, pady=self.PADDING_Y_SMALL)  # パディング削減
        self.review_frame.pack_forget()
        tk.Label(self.review_frame, textvariable=self.review_var, font=(self.FONT_FAMILY, self.FONT_SIZE_SMALL), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY, padx=self.PADDING_X_SMALL, pady=self.PADDING_Y_SMALL, wraplength=self.WRAPLENGTH_DEFAULT, justify='left').pack(fill='x')

        self.best3_frame = tk.Frame(self.sampling_frame, bg=self.WARNING_RED, relief="flat", bd=1)
        self.best3_frame.pack(fill='x', padx=self.PADDING_X_MEDIUM, pady=self.PADDING_Y_SMALL)  # パディング削減
        self.best3_frame.pack_forget()
        tk.Label(self.best3_frame, textvariable=self.best3_var, font=(self.FONT_FAMILY, self.FONT_SIZE_SMALL, "bold"), fg="#ffffff", bg=self.WARNING_RED, padx=self.PADDING_X_SMALL, pady=self.PADDING_Y_SMALL, wraplength=self.WRAPLENGTH_DEFAULT, justify='left').pack(fill='x')

    def show_export_button(self):
        self.export_frame.pack(pady=self.PADDING_Y_SMALL)  # パディング削減

    def hide_export_button(self):
        self.export_frame.pack_forget()
    
    def _set_today_date(self, date_entry):
        """今日の日付を設定"""
        today = datetime.now().strftime('%Y-%m-%d')
        date_entry.delete(0, 'end')
        date_entry.insert(0, today)
