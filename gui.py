import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
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

        # --- UI定数 ---
        self.FONT_FAMILY = "Meiryo"
        self.FONT_SIZE_LARGE = 16
        self.FONT_SIZE_MEDIUM = 12
        self.FONT_SIZE_SMALL = 10
        self.FONT_SIZE_XSMALL = 9
        self.PADDING_X_LARGE = 40
        self.PADDING_X_MEDIUM = 20
        self.PADDING_X_SMALL = 15
        self.PADDING_Y_LARGE = 15
        self.PADDING_Y_MEDIUM = 10
        self.PADDING_Y_SMALL = 5
        self.WRAPLENGTH_DEFAULT = 800 # This will be dynamic later, but for now

        # --- ウィンドウ設定 ---
        self.title("抜取検査数計算ツール - AIアシスト")
        self.geometry("1000x700")
        self.configure(bg=self.LIGHT_GRAY)
        try:
            self.state('zoomed')
        except tk.TclError:
            self._center_window()
        self.bind('<Configure>', self._on_resize)

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

    def _create_widgets(self):
        # スクロール可能なCanvasで全体をラップ
        # Create a frame to hold the canvas and scrollbars
        canvas_frame = tk.Frame(self)
        canvas_frame.pack(fill="both", expand=True)

        # Vertical Scrollbar
        yscroll = tk.Scrollbar(canvas_frame, orient='vertical')
        yscroll.grid(row=0, column=1, sticky='ns')

        # Horizontal Scrollbar
        xscroll = tk.Scrollbar(canvas_frame, orient='horizontal')
        xscroll.grid(row=1, column=0, sticky='ew')

        # Main Canvas
        main_canvas = tk.Canvas(canvas_frame, bg="#ffffff", highlightthickness=0)
        main_canvas.grid(row=0, column=0, sticky='nsew')

        # Configure scrollbars
        main_canvas.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        yscroll.configure(command=main_canvas.yview)
        xscroll.configure(command=main_canvas.xview)

        # Make the canvas frame expand
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)

        main_frame = tk.Frame(main_canvas, bg="#ffffff")
        main_frame_window = main_canvas.create_window((0, 0), window=main_frame, anchor='nw')

        def on_frame_configure(event):
            # Update scrollregion based on main_frame's size
            main_canvas.config(scrollregion=main_canvas.bbox('all'))

            canvas_width = main_canvas.winfo_width()
            canvas_height = main_canvas.winfo_height()
            frame_width = main_frame.winfo_reqwidth()
            frame_height = main_frame.winfo_reqheight()

            # Horizontal Centering
            x_pos = 0
            if frame_width < canvas_width:
                x_pos = (canvas_width - frame_width) / 2
            
            # Vertical Centering
            y_pos = 0
            if frame_height < canvas_height:
                y_pos = (canvas_height - frame_height) / 2

            main_canvas.coords(main_frame_window, x_pos, y_pos)

        # Bind on_frame_configure to main_frame and main_canvas
        main_frame.bind('<Configure>', on_frame_configure)
        main_canvas.bind('<Configure>', on_frame_configure)

        # マウスホイールで縦スクロール
        if platform.system() == 'Windows':
            main_canvas.bind_all('<MouseWheel>', lambda e: main_canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))
        else:
            main_canvas.bind_all('<Button-4>', lambda e: main_canvas.yview_scroll(-1, 'units'))
            main_canvas.bind_all('<Button-5>', lambda e: main_canvas.yview_scroll(1, 'units'))

        

        # --- Content Wrapper for Centering ---
        content_wrapper = tk.Frame(main_frame, bg="#ffffff")
        content_wrapper.pack(expand=True)

        # ヘッダー部分
        header_frame = tk.Frame(main_frame, bg=self.PRIMARY_BLUE, height=80)
        header_frame.pack(fill='x', pady=(self.PADDING_Y_MEDIUM, self.PADDING_Y_SMALL))
        header_frame.pack_propagate(False)
        tk.Label(header_frame, text="🤖 AI抜取検査数計算ツール", font=(self.FONT_FAMILY, self.FONT_SIZE_LARGE, "bold"), fg="#ffffff", bg=self.PRIMARY_BLUE).pack(expand=True)

        # 計算方法の要約
        summary_frame = tk.Frame(content_wrapper, bg="#e9ecef", relief="flat", bd=1)
        summary_frame.pack(fill='x', pady=(0, self.PADDING_Y_MEDIUM))
        summary_text = (
            "【このツールの計算方法】\n"
            "本ツールは統計的品質管理（SQC: Statistical Quality Control）の考え方に基づき、\n"
            "過去の不具合データから不良率を自動計算し、\n"
            "入力した信頼度・c値（許容不良数）に基づいて、\n"
            "不良品を見逃さないために必要な抜取検査数を統計的手法で算出します。\n"
            "抜取検査数と検査水準（I/II/III）は、その根拠とともに分かりやすく表示されます。"
        )
        tk.Label(summary_frame, text=summary_text, fg=self.DARK_GRAY, bg=self.LIGHT_GRAY, font=(self.FONT_FAMILY, self.FONT_SIZE_SMALL), wraplength=950, anchor='w', justify='left', padx=self.PADDING_X_SMALL, pady=self.PADDING_Y_MEDIUM).pack(fill='x')

        # メイン計算フレーム
        self.sampling_frame = tk.Frame(main_frame, bg=self.LIGHT_GRAY, relief="flat", bd=2)
        self.sampling_frame.pack(fill='both', expand=True, padx=self.PADDING_X_LARGE)
        tk.Label(self.sampling_frame, text="📊 抜取検査数計算", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM, "bold"), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(pady=(self.PADDING_Y_MEDIUM, self.PADDING_Y_LARGE))

        # --- 入力フレーム ---
        input_frame = tk.Frame(self.sampling_frame, bg=self.LIGHT_GRAY)
        input_frame.pack(fill='x', padx=self.PADDING_X_LARGE, pady=self.PADDING_Y_LARGE)
        
        # 1行目：品番と数量
        row1_frame = tk.Frame(input_frame, bg=self.LIGHT_GRAY)
        row1_frame.pack(fill='x', pady=self.PADDING_Y_SMALL)
        tk.Label(row1_frame, text="品番:", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left', padx=(0, self.PADDING_Y_SMALL))
        self.sample_pn_entry = tk.Entry(row1_frame, width=20, font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), bg=self.LIGHT_GRAY, fg=self.DARK_GRAY, relief="flat", bd=1, highlightthickness=1, highlightbackground=self.MEDIUM_GRAY, highlightcolor=self.PRIMARY_BLUE)
        self.sample_pn_entry.pack(side='left', padx=self.PADDING_Y_SMALL)
        tk.Label(row1_frame, text="数量 (ロットサイズ):", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left', padx=(self.PADDING_Y_MEDIUM, self.PADDING_Y_SMALL))
        self.sample_qty_entry = tk.Entry(row1_frame, width=12, font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), bg=self.LIGHT_GRAY, fg=self.DARK_GRAY, relief="flat", bd=1, highlightthickness=1, highlightbackground=self.MEDIUM_GRAY, highlightcolor=self.PRIMARY_BLUE)
        self.sample_qty_entry.pack(side='left', padx=self.PADDING_Y_SMALL)

        # 2行目：信頼度とc値
        row2_frame = tk.Frame(input_frame, bg=self.LIGHT_GRAY)
        row2_frame.pack(fill='x', pady=self.PADDING_Y_SMALL)
        tk.Label(row2_frame, text="信頼度(%):", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left', padx=(0, self.PADDING_Y_SMALL))
        self.sample_conf_entry = tk.Entry(row2_frame, width=6, font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), bg=self.LIGHT_GRAY, fg=self.DARK_GRAY, relief="flat", bd=1, highlightthickness=1, highlightbackground=self.MEDIUM_GRAY, highlightcolor=self.PRIMARY_BLUE)
        self.sample_conf_entry.pack(side='left', padx=self.PADDING_Y_SMALL)
        self.sample_conf_entry.insert(0, "99")
        tk.Label(row2_frame, text="c値(許容不良数):", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left', padx=(self.PADDING_Y_MEDIUM, self.PADDING_Y_SMALL))
        self.sample_c_entry = tk.Entry(row2_frame, width=6, font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), bg=self.LIGHT_GRAY, fg=self.DARK_GRAY, relief="flat", bd=1, highlightthickness=1, highlightbackground=self.MEDIUM_GRAY, highlightcolor=self.PRIMARY_BLUE)
        self.sample_c_entry.pack(side='left', padx=self.PADDING_Y_SMALL)
        self.sample_c_entry.insert(0, "0")

        # 説明文
        explain_frame = tk.Frame(input_frame, bg=self.LIGHT_GRAY)
        explain_frame.pack(fill='x', pady=self.PADDING_Y_SMALL)
        explain_conf = (
            "信頼度とは：抜取検査で『不良品を見逃さない確率』です。例：99%なら99%の確率で不良品を検出できることを意味します。\n"
            "c値とは：検査で『許容できる不良品の最大数』です。c=0なら1つも不良品が見つかってはいけない、c=1なら1個まで許容、という意味です。"
        )
        tk.Label(explain_frame, text=explain_conf, fg=self.DARK_GRAY, bg=self.LIGHT_GRAY, font=(self.FONT_FAMILY, self.FONT_SIZE_XSMALL), wraplength=900).pack()

        # 3行目：日付入力
        row3_frame = tk.Frame(input_frame, bg=self.LIGHT_GRAY)
        row3_frame.pack(fill='x', pady=self.PADDING_Y_SMALL)
        tk.Label(row3_frame, text="対象日（開始）:", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left', padx=(0, self.PADDING_Y_SMALL))
        self.sample_start_date_entry = DateEntry(row3_frame, width=12, date_pattern='yyyy-mm-dd', font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), bg=self.LIGHT_GRAY, fg=self.DARK_GRAY, relief="flat", bd=1, highlightthickness=1, highlightbackground=self.MEDIUM_GRAY, highlightcolor=self.PRIMARY_BLUE)
        self.sample_start_date_entry.pack(side='left', padx=self.PADDING_Y_SMALL)
        self.sample_start_date_entry.delete(0, 'end')
        tk.Button(row3_frame, text="クリア", font=(self.FONT_FAMILY, self.FONT_SIZE_XSMALL), command=lambda: self.sample_start_date_entry.delete(0, 'end'), bg=self.MEDIUM_GRAY, fg=self.DARK_GRAY, relief="flat").pack(side='left', padx=(2, self.PADDING_Y_MEDIUM))
        tk.Label(row3_frame, text="～", font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY).pack(side='left')
        self.sample_end_date_entry = DateEntry(row3_frame, width=12, date_pattern='yyyy-mm-dd', font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM), bg=self.LIGHT_GRAY, fg=self.DARK_GRAY, relief="flat", bd=1, highlightthickness=1, highlightbackground=self.MEDIUM_GRAY, highlightcolor=self.PRIMARY_BLUE)
        self.sample_end_date_entry.pack(side='left', padx=self.PADDING_Y_SMALL)
        self.sample_end_date_entry.delete(0, 'end')
        tk.Button(row3_frame, text="クリア", font=(self.FONT_FAMILY, self.FONT_SIZE_XSMALL), command=lambda: self.sample_end_date_entry.delete(0, 'end'), bg=self.MEDIUM_GRAY, fg=self.DARK_GRAY, relief="flat").pack(side='left', padx=(2, self.PADDING_Y_MEDIUM))
        tk.Label(input_frame, text="※ 対象日を未入力の場合は全期間が対象となります。", fg=self.DARK_GRAY, bg=self.LIGHT_GRAY, font=(self.FONT_FAMILY, self.FONT_SIZE_SMALL)).pack(pady=self.PADDING_Y_SMALL)

        # 計算ボタン
        button_frame = tk.Frame(input_frame, bg=self.LIGHT_GRAY)
        button_frame.pack(fill='x', pady=self.PADDING_Y_LARGE)
        self.calc_button = tk.Button(button_frame, text="🚀 計算実行", command=self.controller.start_calculation_thread, font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM, "bold"), bg=self.PRIMARY_BLUE, fg="#ffffff", relief="flat", padx=30, pady=self.PADDING_Y_MEDIUM, cursor="hand2", activebackground=self.ACCENT_BLUE, activeforeground="#ffffff")
        self.calc_button.pack()

        # --- 結果表示フレーム ---
        self.result_frame = tk.Frame(self.sampling_frame, bg=self.LIGHT_GRAY, relief="flat", bd=1)
        self.result_frame.pack(fill='x', padx=self.PADDING_X_LARGE, pady=self.PADDING_Y_LARGE)
        self.result_label = tk.Label(self.result_frame, textvariable=self.result_var, font=(self.FONT_FAMILY, self.FONT_SIZE_MEDIUM, "bold"), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY, padx=self.PADDING_X_MEDIUM, pady=self.PADDING_Y_LARGE, wraplength=self.WRAPLENGTH_DEFAULT, justify='center')
        self.result_label.pack(fill='x')
        self.result_var.set("品番・数量・（任意で対象日）を入力して計算実行ボタンを押してください。")

        # 根拠レビュー表示
        self.review_frame = tk.Frame(self.sampling_frame, bg=self.LIGHT_GRAY, relief="flat", bd=1)
        self.review_frame.pack(fill='x', padx=self.PADDING_X_LARGE, pady=self.PADDING_Y_MEDIUM)
        self.review_frame.pack_forget() # Hide initially
        tk.Label(self.review_frame, textvariable=self.review_var, font=(self.FONT_FAMILY, self.FONT_SIZE_SMALL), fg=self.DARK_GRAY, bg=self.LIGHT_GRAY, padx=self.PADDING_X_SMALL, pady=self.PADDING_Y_SMALL, wraplength=self.WRAPLENGTH_DEFAULT, justify='left').pack(fill='x')

        # 注意喚起表示
        # 注意喚起表示
        self.best3_frame = tk.Frame(self.sampling_frame, bg=self.WARNING_RED, relief="flat", bd=1)
        self.best3_frame.pack(fill='x', padx=self.PADDING_X_LARGE, pady=self.PADDING_Y_MEDIUM)
        self.best3_frame.pack_forget() # Hide initially
        tk.Label(self.best3_frame, textvariable=self.best3_var, font=(self.FONT_FAMILY, self.FONT_SIZE_SMALL, "bold"), fg="#ffffff", bg=self.WARNING_RED, padx=self.PADDING_X_SMALL, pady=self.PADDING_Y_SMALL, wraplength=self.WRAPLENGTH_DEFAULT, justify='left').pack(fill='x')