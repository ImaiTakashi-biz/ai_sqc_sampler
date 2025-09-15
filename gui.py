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
        
        # --- ウィンドウ設定 ---
        self.title("抜取検査数計算ツール - AIアシスト")
        self.geometry("1000x700")
        self.configure(bg="#ffffff")
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
        header_frame = tk.Frame(content_wrapper, bg="#f8f9fa", height=80)
        header_frame.pack(fill='x', pady=(20, 10))
        header_frame.pack_propagate(False)
        tk.Label(header_frame, text="🤖 AI抜取検査数計算ツール", font=("Meiryo", 16, "bold"), fg="#2c3e50", bg="#f8f9fa").pack(expand=True)

        # 計算方法の要約
        summary_frame = tk.Frame(content_wrapper, bg="#e9ecef", relief="flat", bd=1)
        summary_frame.pack(fill='x', pady=(0, 20))
        summary_text = (
            "【このツールの計算方法】\n"
            "本ツールは統計的品質管理（SQC: Statistical Quality Control）の考え方に基づき、\n"
            "過去の不具合データから不良率を自動計算し、\n"
            "入力した信頼度・c値（許容不良数）に基づいて、\n"
            "不良品を見逃さないために必要な抜取検査数を統計的手法で算出します。\n"
            "抜取検査数と検査水準（I/II/III）は、その根拠とともに分かりやすく表示されます。"
        )
        tk.Label(summary_frame, text=summary_text, fg="#495057", bg="#e9ecef", font=("Meiryo", 10), wraplength=950, anchor='w', justify='left', padx=15, pady=10).pack(fill='x')

        # メイン計算フレーム
        self.sampling_frame = tk.Frame(content_wrapper, bg="#ffffff", relief="flat", bd=2)
        self.sampling_frame.pack(fill='both', expand=True, padx=50)
        tk.Label(self.sampling_frame, text="📊 抜取検査数計算", font=("Meiryo", 14, "bold"), fg="#2c3e50", bg="#ffffff").pack(pady=(20, 15))

        # --- 入力フレーム ---
        input_frame = tk.Frame(self.sampling_frame, bg="#ffffff")
        input_frame.pack(fill='x', padx=40, pady=15)
        
        # 1行目：品番と数量
        row1_frame = tk.Frame(input_frame, bg="#ffffff")
        row1_frame.pack(fill='x', pady=5)
        tk.Label(row1_frame, text="品番:", font=("Meiryo", 11), fg="#2c3e50", bg="#ffffff").pack(side='left', padx=(0, 5))
        self.sample_pn_entry = tk.Entry(row1_frame, width=20, font=("Meiryo", 11), bg="#ffffff", fg="#333333", relief="solid", bd=1)
        self.sample_pn_entry.pack(side='left', padx=5)
        tk.Label(row1_frame, text="数量 (ロットサイズ):", font=("Meiryo", 11), fg="#2c3e50", bg="#ffffff").pack(side='left', padx=(20, 5))
        self.sample_qty_entry = tk.Entry(row1_frame, width=12, font=("Meiryo", 11), bg="#ffffff", fg="#333333", relief="solid", bd=1)
        self.sample_qty_entry.pack(side='left', padx=5)

        # 2行目：信頼度とc値
        row2_frame = tk.Frame(input_frame, bg="#ffffff")
        row2_frame.pack(fill='x', pady=5)
        tk.Label(row2_frame, text="信頼度(%):", font=("Meiryo", 11), fg="#2c3e50", bg="#ffffff").pack(side='left', padx=(0, 5))
        self.sample_conf_entry = tk.Entry(row2_frame, width=6, font=("Meiryo", 11), bg="#ffffff", fg="#333333", relief="solid", bd=1)
        self.sample_conf_entry.pack(side='left', padx=5)
        self.sample_conf_entry.insert(0, "99")
        tk.Label(row2_frame, text="c値(許容不良数):", font=("Meiryo", 11), fg="#2c3e50", bg="#ffffff").pack(side='left', padx=(20, 5))
        self.sample_c_entry = tk.Entry(row2_frame, width=6, font=("Meiryo", 11), bg="#ffffff", fg="#333333", relief="solid", bd=1)
        self.sample_c_entry.pack(side='left', padx=5)
        self.sample_c_entry.insert(0, "0")

        # 説明文
        explain_frame = tk.Frame(input_frame, bg="#ffffff")
        explain_frame.pack(fill='x', pady=5)
        explain_conf = (
            "信頼度とは：抜取検査で『不良品を見逃さない確率』です。例：99%なら99%の確率で不良品を検出できることを意味します。\n"
            "c値とは：検査で『許容できる不良品の最大数』です。c=0なら1つも不良品が見つかってはいけない、c=1なら1個まで許容、という意味です。"
        )
        tk.Label(explain_frame, text=explain_conf, fg="#6c757d", bg="#ffffff", font=("Meiryo", 9), wraplength=900).pack()

        # 3行目：日付入力
        row3_frame = tk.Frame(input_frame, bg="#ffffff")
        row3_frame.pack(fill='x', pady=5)
        tk.Label(row3_frame, text="対象日（開始）:", font=("Meiryo", 11), fg="#2c3e50", bg="#ffffff").pack(side='left', padx=(0, 5))
        self.sample_start_date_entry = DateEntry(row3_frame, width=12, date_pattern='yyyy-mm-dd', font=("Meiryo", 11), bg="#ffffff", fg="#333333")
        self.sample_start_date_entry.pack(side='left', padx=5)
        self.sample_start_date_entry.delete(0, 'end')
        tk.Button(row3_frame, text="クリア", font=("Meiryo", 9), command=lambda: self.sample_start_date_entry.delete(0, 'end'), bg="#f8f9fa", relief="flat").pack(side='left', padx=(2, 10))
        tk.Label(row3_frame, text="～", font=("Meiryo", 11), fg="#2c3e50", bg="#ffffff").pack(side='left')
        self.sample_end_date_entry = DateEntry(row3_frame, width=12, date_pattern='yyyy-mm-dd', font=("Meiryo", 11), bg="#ffffff", fg="#333333")
        self.sample_end_date_entry.pack(side='left', padx=5)
        self.sample_end_date_entry.delete(0, 'end')
        tk.Button(row3_frame, text="クリア", font=("Meiryo", 9), command=lambda: self.sample_end_date_entry.delete(0, 'end'), bg="#f8f9fa", relief="flat").pack(side='left', padx=(2, 10))
        tk.Label(input_frame, text="※ 対象日を未入力の場合は全期間が対象となります。", fg="#6c757d", bg="#ffffff", font=("Meiryo", 10)).pack(pady=2)

        # 計算ボタン
        button_frame = tk.Frame(input_frame, bg="#ffffff")
        button_frame.pack(fill='x', pady=15)
        self.calc_button = tk.Button(button_frame, text="🚀 計算実行", command=self.controller.start_calculation_thread, font=("Meiryo", 12, "bold"), bg="#007bff", fg="#ffffff", relief="flat", padx=30, pady=10, cursor="hand2")
        self.calc_button.pack()

        # --- 結果表示フレーム ---
        self.result_frame = tk.Frame(self.sampling_frame, bg="#e9ecef", relief="flat", bd=1)
        self.result_frame.pack(fill='x', padx=40, pady=15)
        self.result_label = tk.Label(self.result_frame, textvariable=self.result_var, font=("Meiryo", 12, "bold"), fg="#2c3e50", bg="#e9ecef", padx=20, pady=15, wraplength=800, justify='center')
        self.result_label.pack(fill='x')
        self.result_var.set("品番・数量・（任意で対象日）を入力して計算実行ボタンを押してください。")

        # 根拠レビュー表示
        review_frame = tk.Frame(self.sampling_frame, bg="#ffffff")
        review_frame.pack(fill='x', padx=40, pady=10)
        tk.Label(review_frame, textvariable=self.review_var, font=("Meiryo", 10), fg="#6c757d", bg="#ffffff", padx=15, pady=8, wraplength=800, justify='left').pack(fill='x')

        # 注意喚起表示
        best3_frame = tk.Frame(self.sampling_frame, bg="#ffffff")
        best3_frame.pack(fill='x', padx=40, pady=10)
        tk.Label(best3_frame, textvariable=self.best3_var, font=("Meiryo", 10, "bold"), fg="#dc3545", bg="#ffffff", padx=15, pady=8, wraplength=800, justify='left').pack(fill='x')