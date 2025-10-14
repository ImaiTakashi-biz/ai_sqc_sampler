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
        self.inspection_mode_var = tk.StringVar()
        self.inspection_mode_label_to_key = {}
        self.inspection_mode_key_to_label = {}
        self.current_inspection_mode_key = "standard"
        
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
        self._bind_shortcuts()

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
        config_defaults = {
            'aql': 0.25,
            'ltpd': 1.0,
            'alpha': 5.0,
            'beta': 10.0,
            'c_value': 0,
        }
        mode_label = "標準"
        mode_details = {
            'aql': config_defaults['aql'],
            'ltpd': config_defaults['ltpd'],
            'alpha': config_defaults['alpha'],
            'beta': config_defaults['beta'],
            'c_value': config_defaults['c_value'],
            'description': "通常ロット"
        }
        config_manager = getattr(self.controller, "config_manager", None)
        if config_manager:
            defaults_source = getattr(config_manager, "DEFAULT_CONFIG", {})
            config_defaults['aql'] = config_manager.get("default_aql", defaults_source.get("default_aql", 0.25))
            config_defaults['ltpd'] = config_manager.get("default_ltpd", defaults_source.get("default_ltpd", 1.0))
            config_defaults['alpha'] = config_manager.get("default_alpha", defaults_source.get("default_alpha", 5.0))
            config_defaults['beta'] = config_manager.get("default_beta", defaults_source.get("default_beta", 10.0))
            config_defaults['c_value'] = config_manager.get("default_c_value", defaults_source.get("default_c_value", 0))

            if hasattr(config_manager, "get_inspection_mode_choices"):
                mode_choices = config_manager.get_inspection_mode_choices()
                self.inspection_mode_key_to_label = mode_choices
                self.inspection_mode_label_to_key = {label: key for key, label in mode_choices.items()}
                current_mode_key = config_manager.get_inspection_mode()
                self.current_inspection_mode_key = current_mode_key
                mode_label = mode_choices.get(current_mode_key, mode_label)
                if hasattr(config_manager, "get_inspection_mode_details"):
                    mode_details = config_manager.get_inspection_mode_details(current_mode_key)
            else:
                self.inspection_mode_key_to_label = {}
                self.inspection_mode_label_to_key = {}
        else:
            self.inspection_mode_key_to_label = {}
            self.inspection_mode_label_to_key = {}

        if not self.inspection_mode_label_to_key:
            fallback_modes = {
                "tightened": "強化",
                "standard": "標準",
                "reduced": "緩和"
            }
            self.inspection_mode_key_to_label = fallback_modes
            self.inspection_mode_label_to_key = {label: key for key, label in fallback_modes.items()}
            self.current_inspection_mode_key = "standard"
            mode_label = fallback_modes["standard"]
            mode_details = {
                'aql': 0.25,
                'ltpd': 1.0,
                'alpha': 5.0,
                'beta': 10.0,
                'c_value': 0,
                'description': "通常ロット"
            }

        self.inspection_mode_var.set(mode_label)

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

        # マウスホイールスクロール機能
        def handle_mousewheel(event):
            # メインウィンドウ上でのスクロールのみ処理
            if event.widget.winfo_toplevel() is self:
                if platform.system() == 'Windows':
                    main_canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
                else:
                    if event.num == 4:
                        main_canvas.yview_scroll(-1, 'units')
                    elif event.num == 5:
                        main_canvas.yview_scroll(1, 'units')
        
        # マウスホイールイベントのバインド
        if platform.system() == 'Windows':
            main_canvas.bind_all('<MouseWheel>', handle_mousewheel)
        else:
            main_canvas.bind_all('<Button-4>', handle_mousewheel)
            main_canvas.bind_all('<Button-5>', handle_mousewheel)
        
        # キャンバスにフォーカスを設定してマウスホイールイベントを受け取れるようにする
        main_canvas.bind('<Enter>', lambda e: main_canvas.focus_set())
        main_canvas.bind('<Leave>', lambda e: self.focus_set())
        
        # メインフレーム全体でもマウスホイールイベントを受け取れるようにする
        main_frame.bind('<Enter>', lambda e: main_canvas.focus_set())
        
        # ウィンドウ全体でマウスホイールイベントを処理
        self.bind('<MouseWheel>', handle_mousewheel)
        if platform.system() != 'Windows':
            self.bind('<Button-4>', handle_mousewheel)
            self.bind('<Button-5>', handle_mousewheel)

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
            "※有限母集団補正＝ロット全体が限られた数であることを踏まえ、「実績データ＋統計理論」から過不足のない抜取数に調整する仕組み。"
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

        mode_frame = tk.Frame(input_frame, bg=self.LIGHT_GRAY)
        mode_frame.pack(fill='x', pady=self.PADDING_Y_SMALL)
        tk.Label(mode_frame, text="検査区分:", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left', padx=(0, self.PADDING_Y_SMALL))

        mode_values = list(self.inspection_mode_label_to_key.keys())
        self.inspection_mode_selector = ttk.Combobox(
            mode_frame,
            state='readonly',
            values=mode_values,
            textvariable=self.inspection_mode_var,
            width=8,
            font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM)
        )
        self.inspection_mode_selector.pack(side='left', padx=self.PADDING_Y_SMALL)
        self.inspection_mode_selector.bind("<<ComboboxSelected>>", self._handle_inspection_mode_change)
        self.inspection_mode_selector.set(self.inspection_mode_var.get())

        self.inspection_mode_info_label = tk.Label(
            input_frame,
            text=self._format_inspection_mode_summary(mode_details),
            font=(self.FONT_FAMILY, self.FONT_SIZE_SMALL),
            fg=self.DARK_GRAY,
            bg=self.LIGHT_GRAY,
            anchor='w',
            justify='left',
            wraplength=600
        )
        self.inspection_mode_info_label.pack(fill='x', pady=(0, self.PADDING_Y_SMALL))

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
        self.export_button = tk.Button(self.export_frame, text="📄 レポート出力", command=self.controller.export_results, font=(self.FONT_FAMILY, self.FONT_SIZE_SMALL), bg=self.INFO_GREEN, fg="#ffffff", relief="flat", padx=15, pady=5, cursor="hand2", activebackground=self.ACCENT_BLUE)
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

    def _handle_inspection_mode_change(self, event=None):
        """検査区分の変更イベント"""
        selected_label = self.inspection_mode_var.get()
        mode_key = self.inspection_mode_label_to_key.get(selected_label)
        if not mode_key:
            return
        self.current_inspection_mode_key = mode_key
        if hasattr(self.controller, "on_inspection_mode_change"):
            self.controller.on_inspection_mode_change(mode_key)

    def apply_inspection_mode_preset(self, preset, mode_label=None):
        """検査区分プリセットを入力欄へ反映"""
        if mode_label:
            self.inspection_mode_var.set(mode_label)
            if mode_label in self.inspection_mode_label_to_key:
                self.current_inspection_mode_key = self.inspection_mode_label_to_key[mode_label]
        elif preset and preset.get('label'):
            inferred_label = preset.get('label')
            self.inspection_mode_var.set(inferred_label)
            if inferred_label in self.inspection_mode_label_to_key:
                self.current_inspection_mode_key = self.inspection_mode_label_to_key[inferred_label]
        if not preset:
            return

        self.current_mode_details = preset.copy()
        self._update_inspection_mode_info(preset)
        if hasattr(self, 'inspection_mode_selector'):
            self.inspection_mode_selector.set(self.inspection_mode_var.get())

    def refresh_inspection_mode_choices(self, label_to_key_map, current_key):
        """検査区分の選択肢を更新"""
        if not label_to_key_map:
            return

        self.inspection_mode_label_to_key = label_to_key_map
        self.inspection_mode_key_to_label = {key: label for label, key in label_to_key_map.items()}
        values = list(self.inspection_mode_label_to_key.keys())
        if hasattr(self, 'inspection_mode_selector'):
            self.inspection_mode_selector['values'] = values

        label = self.inspection_mode_key_to_label.get(current_key) if hasattr(self, 'inspection_mode_key_to_label') else None
        if label is None and values:
            label = values[0]
            current_key = self.inspection_mode_label_to_key[label]
        if label:
            self.current_inspection_mode_key = current_key
            self.inspection_mode_var.set(label)
            if hasattr(self, 'inspection_mode_selector'):
                self.inspection_mode_selector.set(label)
            if hasattr(self.controller, 'config_manager'):
                details = self.controller.config_manager.get_inspection_mode_details(current_key)
                self.current_mode_details = details
                self._update_inspection_mode_info(details)

    def _format_inspection_mode_summary(self, details):
        """検査区分のサマリー文字列を生成"""
        if not details:
            return ''
        aql = self._format_numeric(details.get('aql'))
        ltpd = self._format_numeric(details.get('ltpd'))
        alpha = self._format_numeric(details.get('alpha'))
        beta = self._format_numeric(details.get('beta'))
        c_value = self._format_numeric(details.get('c_value'))
        description = details.get('description', '')
        return f"AQL: {aql}%  LTPD: {ltpd}%  α: {alpha}%  β: {beta}%  c値: {c_value}\n用途: {description}"

    def _update_inspection_mode_info(self, details):
        """検査区分の説明ラベルを更新"""
        if hasattr(self, 'inspection_mode_info_label'):
            self.inspection_mode_info_label.config(text=self._format_inspection_mode_summary(details))

    def _set_entry_value(self, entry_widget, value):
        """エントリーへ値を反映"""
        if entry_widget is None:
            return
        entry_widget.delete(0, 'end')
        entry_widget.insert(0, self._format_numeric(value))

    @staticmethod
    def _format_numeric(value):
        """数値を表示用に整形"""
        if isinstance(value, float):
            formatted = f"{value:.2f}".rstrip('0').rstrip('.')
            return formatted or '0'
        return str(value) if value is not None else ''

    def show_export_button(self):
        self.export_frame.pack(pady=self.PADDING_Y_SMALL)  # パディング削減

    def hide_export_button(self):
        self.export_frame.pack_forget()
    
    def _set_today_date(self, date_entry):
        """今日の日付を設定"""
        today = datetime.now().strftime('%Y-%m-%d')
        date_entry.delete(0, 'end')
        date_entry.insert(0, today)
    
    def _bind_shortcuts(self):
        """ショートカットキーの設定"""
        def trigger_calculation(event=None):
            self.controller.start_calculation_thread()
            return "break"
        
        def trigger_export(event=None):
            try:
                self.controller.export_results()
            except Exception:
                pass
            return "break"
        
        self.bind_all('<Control-Return>', trigger_calculation)
        self.bind_all('<Control-KP_Enter>', trigger_calculation)
        self.bind_all('<Control-e>', trigger_export)