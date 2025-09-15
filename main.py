import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc
import configparser
import os
from datetime import datetime, timedelta
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from io import BytesIO
from PIL import Image, ImageTk
from tkcalendar import DateEntry
from collections import Counter
import math
import csv
from tkinter import filedialog
from scipy.stats import binom
import threading
import time

# --- 定数 ---
CONFIG_FILE = 'config.ini'
DB_FILE_PATH_KEY = 'path'
DB_SECTION = 'DATABASE'
DEFAULT_DB_PATH = r'C:\Users\SEIZOU-20\PycharmProjects\ai_sqc_sampler\不具合情報記録.accdb'
DEFECT_COLUMNS = [
    "外観キズ", "圧痕", "切粉", "毟れ", "穴大", "穴小", "穴キズ", "バリ", "短寸", "面粗", "サビ", "ボケ", "挽目", "汚れ", "メッキ", "落下",
    "フクレ", "ツブレ", "ボッチ", "段差", "バレル石", "径プラス", "径マイナス", "ゲージ", "異物混入", "形状不良", "こすれ", "変色シミ", "材料キズ", "ゴミ", "その他"
]

# --- 設定管理 ---
def create_default_config():
    """デフォルトの設定ファイルを作成する"""
    config = configparser.ConfigParser()
    config[DB_SECTION] = {DB_FILE_PATH_KEY: DEFAULT_DB_PATH}
    with open(CONFIG_FILE, 'w', encoding='utf-8') as configfile:
        config.write(configfile)

def get_db_path():
    """設定ファイルからDBパスを取得する"""
    if not os.path.exists(CONFIG_FILE):
        create_default_config()
    
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE, encoding='utf-8')
    return config[DB_SECTION][DB_FILE_PATH_KEY]

# --- データベース接続 ---
def get_db_connection():
    """データベース接続を取得する"""
    db_path = get_db_path()
    if not os.path.exists(db_path):
        messagebox.showerror("エラー", f"データベースファイルが見つかりません。\nパス: {db_path}\nconfig.iniを確認してください。")
        return None
    
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={db_path};'
    )
    try:
        return pyodbc.connect(conn_str)
    except pyodbc.Error as ex:
        sqlstate = ex.args[0]
        messagebox.showerror("データベース接続エラー", f"エラーが発生しました: {sqlstate}\n{ex}")
        return None

# --- 抜取検査数計算機能 ---
def _build_sql_query(base_sql, product_number, start_date, end_date):
    """SQLクエリとパラメータを動的に構築する"""
    sql_parts = [base_sql]
    params = [product_number]
    # WHERE句がすでにあるかどうかを簡易的に判定
    has_where = ' where ' in base_sql.lower()
    
    if start_date:
        keyword = "AND" if has_where else "WHERE"
        sql_parts.append(f"{keyword} [指示日] >= ?")
        params.append(start_date)
        has_where = True # 次の条件のために更新
    if end_date:
        keyword = "AND" if has_where else "WHERE"
        sql_parts.append(f"{keyword} [指示日] <= ?")
        params.append(end_date)
        
    return " ".join(sql_parts), params

def _get_user_inputs():
    """GUIからユーザー入力を取得し、辞書として返す"""
    inputs = {
        'product_number': sample_pn_entry.get(),
        'lot_size_str': sample_qty_entry.get(),
        'start_date': sample_start_date_entry.get().strip() or None,
        'end_date': sample_end_date_entry.get().strip() or None,
        'conf_str': sample_conf_entry.get().strip() or "99",
        'c_str': sample_c_entry.get().strip() or "0"
    }
    try:
        inputs['lot_size'] = int(inputs['lot_size_str'])
        inputs['conf'] = float(inputs['conf_str']) / 100
        inputs['c'] = int(inputs['c_str'])
    except ValueError:
        messagebox.showwarning("入力エラー", "数量、信頼度、c値は数値で入力してください。")
        return None
    return inputs

def _setup_progress_window():
    """計算中のプログレスウィンドウをセットアップし、関連ウィジェットを返す"""
    calc_button.config(state='disabled', text="🔄 計算中...", bg="#6c757d")
    
    progress_window = tk.Toplevel(app)
    progress_window.title("AI計算中")
    progress_window.geometry("400x200")
    progress_window.configure(bg="#ffffff")
    progress_window.resizable(False, False)
    
    x = (app.winfo_screenwidth() // 2) - (400 // 2)
    y = (app.winfo_screenheight() // 2) - (200 // 2)
    progress_window.geometry(f"400x200+{x}+{y}")
    
    progress_bar = ttk.Progressbar(progress_window, mode='indeterminate', length=300)
    progress_bar.pack(pady=20)
    progress_bar.start()
    
    progress_label = tk.Label(progress_window, text="🤖 AIが統計計算を開始しました...", font=("Meiryo", 12, "bold"), fg="#2c3e50", bg="#ffffff")
    progress_label.pack(pady=10)
    detail_label = tk.Label(progress_window, text="データベースから過去の不具合データを分析中", font=("Meiryo", 10), fg="#6c757d", bg="#ffffff")
    detail_label.pack(pady=5)
    
    # 既存の結果表示をクリア
    result_var.set("")
    review_var.set("")
    best3_var.set("")
    for widget_name in ['main_sample_label', 'level_label', 'reason_label', 'advice_label']:
        if hasattr(app, widget_name):
            getattr(app, widget_name).destroy()

    app.update_idletasks()
    return progress_window, detail_label

def _fetch_data(cursor, product_number, start_date, end_date):
    """データベースからデータを取得・集計する（UI更新なし）"""
    data = {}
    # 数量・不具合数の集計
    sql, params = _build_sql_query("SELECT SUM([数量]), SUM([総不具合数]) FROM t_不具合情報 WHERE [品番] = ?", product_number, start_date, end_date)
    sum_row = cursor.execute(sql, *params).fetchone()
    data['total_qty'] = sum_row[0] or 0
    data['total_defect'] = sum_row[1] or 0
    data['defect_rate'] = (data['total_defect'] / data['total_qty'] * 100) if data['total_qty'] else 0

    # 不具合内容ごとの集計
    columns_str = ", ".join(f"SUM(IIF([{col}] IS NOT NULL AND [{col}]<>0, [{col}], 0))" for col in DEFECT_COLUMNS)
    sql, params = _build_sql_query(f"SELECT {columns_str} FROM t_不具合情報 WHERE [品番] = ?", product_number, start_date, end_date)
    defect_counts = cursor.execute(sql, *params).fetchone()
    
    defect_rates = []
    if data['total_qty'] > 0 and defect_counts:
        for col, count in zip(DEFECT_COLUMNS, defect_counts):
            count = count or 0
            if count > 0:
                rate = (count / data['total_qty'] * 100)
                defect_rates.append((col, rate, count))
    defect_rates.sort(key=lambda x: x[2], reverse=True)
    data['defect_rates_sorted'] = defect_rates
    data['best5'] = [(col, count) for col, rate, count in defect_rates[:5]]
    
    return data

def _calculate_stats(db_data, lot_size, conf, c):
    """統計計算を実行する（UI更新なし）"""
    results = {}
    defect_rate = db_data['defect_rate']

    # 検査水準決定
    if defect_rate == 0:
        results['level_text'] = "ゆるい(I)"
        results['level_reason'] = "過去の不具合が0件だったため、最もゆるい水準（I）を適用しています。"
    elif 0 < defect_rate <= 0.5:
        results['level_text'] = "普通(II)"
        results['level_reason'] = "過去の不具合率が0.5%以下だったため、普通水準（II）を適用しています。"
    else:
        results['level_text'] = "きつい(III)"
        results['level_reason'] = "過去の不具合率が0.5%を超えていたため、きつい水準（III）を適用しています。"

    # 抜取数計算
    p = defect_rate / 100
    n_sample = "計算不可"
    if p > 0 and 0 < conf < 1:
        try:
            if c == 0:
                n_sample = math.ceil(math.log(1 - conf) / math.log(1 - p))
            else:
                n = 1
                limit = max(lot_size * 2, 10000)
                while n <= limit:
                    if binom.cdf(c, n, p) >= 1 - conf:
                        n_sample = n
                        break
                    n += 1
                else:
                    n_sample = f">{limit} (計算断念)"
        except (ValueError, OverflowError):
            n_sample = "計算エラー"
    results['sample_size'] = n_sample
    return results

def _update_ui(db_data, stats_results, inputs):
    """計算結果をUIに反映する"""
    def format_int(n):
        try:
            return f"{int(n):,}"
        except (ValueError, TypeError):
            return str(n)

    sample_size_disp = format_int(stats_results['sample_size'])
    
    app.main_sample_label = tk.Label(result_frame, text=f"抜取検査数: {sample_size_disp} 個", font=("Meiryo", 32, "bold"), fg="#007bff", bg="#e9ecef", pady=10)
    app.main_sample_label.pack(pady=(10, 0))
    app.level_label = tk.Label(result_frame, text=f"検査水準: {stats_results['level_text']}", font=("Meiryo", 16, "bold"), fg="#2c3e50", bg="#e9ecef", pady=5)
    app.level_label.pack()
    app.reason_label = tk.Label(result_frame, text=f"根拠: {stats_results['level_reason']}", font=("Meiryo", 12), fg="#6c757d", bg="#e9ecef", pady=5, wraplength=800)
    app.reason_label.pack()

    period_text = f"（{inputs['start_date'] or '最初'}～{inputs['end_date'] or '最新'}）" if inputs['start_date'] or inputs['end_date'] else "（全期間対象）"
    review_text = (
        f"【根拠レビュー】\n"
        f"・ロットサイズ: {format_int(inputs['lot_size'])}\n"
        f"・対象期間: {period_text}\n"
        f"・数量合計: {format_int(db_data['total_qty'])}個\n"
        f"・不具合数合計: {format_int(db_data['total_defect'])}個\n"
        f"・不良率: {db_data['defect_rate']:.2f}%\n"
        f"・信頼度: {inputs['conf']*100:.1f}%\n"
        f"・c値: {inputs['c']}\n"
        f"・推奨抜取検査数: {sample_size_disp} 個\n"
        f"（c={inputs['c']}, 信頼度={inputs['conf']*100:.1f}%の条件で自動計算）"
    )
    review_var.set(review_text)

    if db_data['best5']:
        best5_text = "【検査時の注意喚起：過去不具合ベスト5】\n"
        for i, (naiyo, count) in enumerate(db_data['best5'], 1):
            rate = next((r for col, r, c_ in db_data['defect_rates_sorted'] if col == naiyo), 0)
            best5_text += f"{i}. {naiyo}（{format_int(count)}個, {rate:.2f}%）\n"
    else:
        best5_text = "【検査時の注意喚起】\n該当期間に不具合データがありません。"
    best3_var.set(best5_text)

    advice = ""
    if db_data['best5'] and db_data['best5'][0][1] > 0:
        advice = f"過去最多の不具合は『{db_data['best5'][0][0]}』です。検査時は特にこの点にご注意ください。"
    elif db_data['total_defect'] > 0:
        advice = "過去の不具合傾向から特に目立つ項目はありませんが、標準的な検査を心がけましょう。"
    else:
        advice = "過去の不具合データが少ないため、全般的に注意して検査を行ってください。"
    app.advice_label = tk.Label(sampling_frame, text=advice, font=("Meiryo", 9), fg="#dc3545", bg="#ffffff")
    app.advice_label.pack(after=result_label, pady=(0, 5))

def _calculation_worker(inputs, progress_window, detail_label):
    """計算処理の本体（バックグラウンドスレッドで実行）"""
    try:
        app.after(0, lambda: detail_label.config(text="データベースに接続中..."))
        conn = get_db_connection()
        if not conn:
            raise ConnectionError("DB接続に失敗")

        with conn.cursor() as cursor:
            app.after(0, lambda: detail_label.config(text="不具合データを集計中..."))
            db_data = _fetch_data(cursor, inputs['product_number'], inputs['start_date'], inputs['end_date'])
            
            app.after(0, lambda: detail_label.config(text="抜取検査数を計算中..."))
            stats_results = _calculate_stats(db_data, inputs['lot_size'], inputs['conf'], inputs['c'])
        
        app.after(0, lambda: detail_label.config(text="結果を表示中..."))
        app.after(0, _update_ui, db_data, stats_results, inputs)
        app.after(0, _finish_calculation, progress_window, True)

    except Exception as e:
        if "DB接続に失敗" not in str(e):
             app.after(0, lambda: messagebox.showerror("計算エラー", f"バックグラウンド処理中にエラーが発生しました: {e}"))
        app.after(0, _finish_calculation, progress_window, False)

def _finish_calculation(progress_window, success):
    """計算後のUI後処理"""
    if progress_window.winfo_exists():
        progress_window.destroy()
    calc_button.config(state='normal', text="🚀 計算実行", bg="#007bff")
    if success:
        messagebox.showinfo("計算完了", "✅ AIが統計分析を完了しました！")

def calculate_sample_size():
    """計算スレッドを開始するラッパー関数"""
    inputs = _get_user_inputs()
    if not inputs:
        return

    progress_window, detail_label = _setup_progress_window()

    thread = threading.Thread(
        target=_calculation_worker,
        args=(inputs, progress_window, detail_label)
    )
    thread.daemon = True
    thread.start()

# --- GUIの構築 ---
app = tk.Tk()
app.title("抜取検査数計算ツール - AIアシスト")
app.geometry("1000x700")
app.configure(bg="#ffffff")

# 画面中央に配置
app.update_idletasks()
width = app.winfo_width()
height = app.winfo_height()
x = (app.winfo_screenwidth() // 2) - (width // 2)
y = (app.winfo_screenheight() // 2) - (height // 2)
app.geometry(f"{width}x{height}+{x}+{y}")

# 最大化時の動作を設定
def on_maximize():
    # 最大化時に中央に配置
    app.update_idletasks()
    width = app.winfo_width()
    height = app.winfo_height()
    x = (app.winfo_screenwidth() // 2) - (width // 2)
    y = (app.winfo_screenheight() // 2) - (height // 2)
    app.geometry(f"{width}x{height}+{x}+{y}")

# ウィンドウの状態変更を監視
app.bind('<Configure>', lambda e: on_maximize() if app.state() == 'zoomed' else None)

# スクロール可能なCanvasで全体をラップ
main_canvas = tk.Canvas(app, bg="#ffffff", highlightthickness=0)
main_canvas.pack(side='left', fill='both', expand=True)
yscroll = tk.Scrollbar(app, orient='vertical', command=main_canvas.yview)
yscroll.pack(side='right', fill='y')
main_canvas.configure(yscrollcommand=yscroll.set)

main_frame = tk.Frame(main_canvas, bg="#ffffff")
main_canvas.create_window((0, 0), window=main_frame, anchor='nw')

def on_configure(event):
    main_canvas.config(scrollregion=main_canvas.bbox('all'))
main_frame.bind('<Configure>', on_configure)

# マウスホイールで縦スクロール
import platform
if platform.system() == 'Windows':
    main_canvas.bind_all('<MouseWheel>', lambda e: main_canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))
else:
    main_canvas.bind_all('<Button-4>', lambda e: main_canvas.yview_scroll(-1, 'units'))
    main_canvas.bind_all('<Button-5>', lambda e: main_canvas.yview_scroll(1, 'units'))



# ヘッダー部分
header_frame = tk.Frame(main_frame, bg="#f8f9fa", height=80)
header_frame.pack(fill='x', pady=(20, 10))
header_frame.pack_propagate(False)

header_title = tk.Label(header_frame, text="🤖 AI抜取検査数計算ツール", 
                       font=("Meiryo", 16, "bold"), 
                       fg="#2c3e50", bg="#f8f9fa")
header_title.pack(expand=True)

# 計算方法の要約（ヘッダー下に表示）
summary_frame = tk.Frame(main_frame, bg="#e9ecef", relief="flat", bd=1)
summary_frame.pack(fill='x', pady=(0, 20))

summary_text = (
    "【このツールの計算方法】\n"
    "本ツールは統計的品質管理（SQC: Statistical Quality Control）の考え方に基づき、\n"
    "過去の不具合データから不良率を自動計算し、\n"
    "入力した信頼度・c値（許容不良数）に基づいて、\n"
    "不良品を見逃さないために必要な抜取検査数を統計的手法で算出します。\n"
    "抜取検査数と検査水準（I/II/III）は、その根拠とともに分かりやすく表示されます。"
)
summary_label = tk.Label(summary_frame, text=summary_text, 
                        fg="#495057", bg="#e9ecef", 
                        font=("Meiryo", 10), 
                        wraplength=950, anchor='w', justify='left',
                        padx=15, pady=10)
summary_label.pack(fill='x')

# メイン計算フレーム
sampling_frame = tk.Frame(main_frame, bg="#ffffff", relief="flat", bd=2)
sampling_frame.pack(fill='both', expand=True, padx=50)

# タイトル
title_label = tk.Label(sampling_frame, text="📊 抜取検査数計算", 
                      font=("Meiryo", 14, "bold"), 
                      fg="#2c3e50", bg="#ffffff")
title_label.pack(pady=(20, 15))

# 入力フレーム
input_frame = tk.Frame(sampling_frame, bg="#ffffff")
input_frame.pack(fill='x', padx=40, pady=15)

# 1行目：品番と数量
row1_frame = tk.Frame(input_frame, bg="#ffffff")
row1_frame.pack(fill='x', pady=5)

tk.Label(row1_frame, text="品番:", font=("Meiryo", 11), fg="#2c3e50", bg="#ffffff").pack(side='left', padx=(0, 5))
sample_pn_entry = tk.Entry(row1_frame, width=20, font=("Meiryo", 11), bg="#ffffff", fg="#333333", relief="solid", bd=1)
sample_pn_entry.pack(side='left', padx=5)

tk.Label(row1_frame, text="数量 (ロットサイズ):", font=("Meiryo", 11), fg="#2c3e50", bg="#ffffff").pack(side='left', padx=(20, 5))
sample_qty_entry = tk.Entry(row1_frame, width=12, font=("Meiryo", 11), bg="#ffffff", fg="#333333", relief="solid", bd=1)
sample_qty_entry.pack(side='left', padx=5)

# 2行目：信頼度とc値
row2_frame = tk.Frame(input_frame, bg="#ffffff")
row2_frame.pack(fill='x', pady=5)

tk.Label(row2_frame, text="信頼度(%):", font=("Meiryo", 11), fg="#2c3e50", bg="#ffffff").pack(side='left', padx=(0, 5))
sample_conf_entry = tk.Entry(row2_frame, width=6, font=("Meiryo", 11), bg="#ffffff", fg="#333333", relief="solid", bd=1)
sample_conf_entry.pack(side='left', padx=5)
sample_conf_entry.insert(0, "99")

tk.Label(row2_frame, text="c値(許容不良数):", font=("Meiryo", 11), fg="#2c3e50", bg="#ffffff").pack(side='left', padx=(20, 5))
sample_c_entry = tk.Entry(row2_frame, width=6, font=("Meiryo", 11), bg="#ffffff", fg="#333333", relief="solid", bd=1)
sample_c_entry.pack(side='left', padx=5)
sample_c_entry.insert(0, "0")

# 説明文
explain_frame = tk.Frame(input_frame, bg="#ffffff")
explain_frame.pack(fill='x', pady=5)

explain_conf = (
    "信頼度とは：抜取検査で『不良品を見逃さない確率』です。例：99%なら99%の確率で不良品を検出できることを意味します。\n"
    "c値とは：検査で『許容できる不良品の最大数』です。c=0なら1つも不良品が見つかってはいけない、c=1なら1個まで許容、という意味です。"
)
explain_label = tk.Label(explain_frame, text=explain_conf, 
                        fg="#6c757d", bg="#ffffff", 
                        font=("Meiryo", 9), wraplength=900)
explain_label.pack()

# 3行目：日付入力
row3_frame = tk.Frame(input_frame, bg="#ffffff")
row3_frame.pack(fill='x', pady=5)

tk.Label(row3_frame, text="対象日（開始）:", font=("Meiryo", 11), fg="#2c3e50", bg="#ffffff").pack(side='left', padx=(0, 5))
sample_start_date_entry = DateEntry(row3_frame, width=12, date_pattern='yyyy-mm-dd', 
                                   font=("Meiryo", 11), bg="#ffffff", fg="#333333")
sample_start_date_entry.pack(side='left', padx=5)
sample_start_date_entry.delete(0, 'end')

# カレンダーを開いた時に当月を表示
import calendar
from datetime import date

def show_today_month_start(event):
    today = date.today()
    sample_start_date_entry._top_cal.selection_set(today)
    sample_start_date_entry._top_cal._display_calendar(today.year, today.month)
sample_start_date_entry.bind("<Button-1>", show_today_month_start, add='+')

def clear_start_date():
    sample_start_date_entry.delete(0, 'end')
clear_start_btn = tk.Button(row3_frame, text="クリア", font=("Meiryo", 9), command=clear_start_date, bg="#f8f9fa", relief="flat")
clear_start_btn.pack(side='left', padx=(2, 10))

tk.Label(row3_frame, text="～", font=("Meiryo", 11), fg="#2c3e50", bg="#ffffff").pack(side='left')
sample_end_date_entry = DateEntry(row3_frame, width=12, date_pattern='yyyy-mm-dd', 
                                 font=("Meiryo", 11), bg="#ffffff", fg="#333333")
sample_end_date_entry.pack(side='left', padx=5)
sample_end_date_entry.delete(0, 'end')

def show_today_month_end(event):
    today = date.today()
    sample_end_date_entry._top_cal.selection_set(today)
    sample_end_date_entry._top_cal._display_calendar(today.year, today.month)
sample_end_date_entry.bind("<Button-1>", show_today_month_end, add='+')

def clear_end_date():
    sample_end_date_entry.delete(0, 'end')
clear_end_btn = tk.Button(row3_frame, text="クリア", font=("Meiryo", 9), command=clear_end_date, bg="#f8f9fa", relief="flat")
clear_end_btn.pack(side='left', padx=(2, 10))

# 日付説明
date_explain_label = tk.Label(input_frame, text="※ 対象日を未入力の場合は全期間が対象となります。", 
                             fg="#6c757d", bg="#ffffff", font=("Meiryo", 10))
date_explain_label.pack(pady=2)

# 計算ボタン
button_frame = tk.Frame(input_frame, bg="#ffffff")
button_frame.pack(fill='x', pady=15)

calc_button = tk.Button(button_frame, text="🚀 計算実行", 
                       command=calculate_sample_size, 
                       font=("Meiryo", 12, "bold"),
                       bg="#007bff", fg="#ffffff",
                       relief="flat", padx=30, pady=10,
                       cursor="hand2")
calc_button.pack()



# 結果表示フレーム
result_frame = tk.Frame(sampling_frame, bg="#e9ecef", relief="flat", bd=1)
result_frame.pack(fill='x', padx=40, pady=15)

result_var = tk.StringVar()
result_label = tk.Label(result_frame, textvariable=result_var, 
                       font=("Meiryo", 12, "bold"), 
                       fg="#2c3e50", bg="#e9ecef",
                       padx=20, pady=15, wraplength=800, justify='center')
result_label.pack(fill='x')
result_var.set("品番・数量・（任意で対象日）を入力して計算実行ボタンを押してください。")

# 根拠レビュー表示
review_frame = tk.Frame(sampling_frame, bg="#ffffff")
review_frame.pack(fill='x', padx=40, pady=10)

review_var = tk.StringVar()
review_label = tk.Label(review_frame, textvariable=review_var, 
                       font=("Meiryo", 10), 
                       fg="#6c757d", bg="#ffffff",
                       padx=15, pady=8, wraplength=800, justify='left')
review_label.pack(fill='x')

# 注意喚起表示
best3_frame = tk.Frame(sampling_frame, bg="#ffffff")
best3_frame.pack(fill='x', padx=40, pady=10)

best3_var = tk.StringVar()
best3_label = tk.Label(best3_frame, textvariable=best3_var, 
                      font=("Meiryo", 10, "bold"), 
                      fg="#dc3545", bg="#ffffff",
                      padx=15, pady=8, wraplength=800, justify='left')
best3_label.pack(fill='x')

# グラフ表示フレーム（非表示）
# graph_frame = tk.Frame(sampling_frame, bg="#ffffff")
# graph_frame.pack(fill='both', expand=True, padx=20, pady=10)

# graph_title = tk.Label(graph_frame, text="📈 グラフ", 
#                       font=("Meiryo", 12, "bold"), 
#                       fg="#2c3e50", bg="#ffffff")
# graph_title.pack(pady=(0, 10))

if __name__ == "__main__":
    app.mainloop()