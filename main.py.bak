import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pyodbc
import configparser
import os
from scipy.stats import binom
import math
import threading
from gui import App

# --- 定数 ---
CONFIG_FILE = 'config.ini'
DB_FILE_PATH_KEY = 'path'
DB_SECTION = 'DATABASE'
DEFAULT_DB_PATH = r'C:\Users\SEIZOU-20\Desktop\AI関連\access_test\不具合情報記録.accdb'
DEFECT_COLUMNS = [
    "外観キズ", "圧痕", "切粉", "毟れ", "穴大", "穴小", "穴キズ", "バリ", "短寸", "面粗", "サビ", "ボケ", "挽目", "汚れ", "メッキ", "落下",
    "フクレ", "ツブレ", "ボッチ", "段差", "バレル石", "径プラス", "径マイナス", "ゲージ", "異物混入", "形状不良", "こすれ", "変色シミ", "材料キズ", "ゴミ", "その他"
]

class MainController:
    def __init__(self):
        self.app = App(self)
        self.progress_window = None
        self.detail_label = None

    def run(self):
        self.app.mainloop()

    # --- 設定管理 & DB接続 ---
    def _create_default_config(self):
        config = configparser.ConfigParser()
        config[DB_SECTION] = {DB_FILE_PATH_KEY: DEFAULT_DB_PATH}
        with open(CONFIG_FILE, 'w', encoding='utf-8') as configfile:
            config.write(configfile)

    def _get_db_path(self):
        if not os.path.exists(CONFIG_FILE):
            self._create_default_config()
        config = configparser.ConfigParser()
        config.read(CONFIG_FILE, encoding='utf-8')
        return config[DB_SECTION][DB_FILE_PATH_KEY]

    def _get_db_connection(self):
        db_path = self._get_db_path()
        if not os.path.exists(db_path):
            messagebox.showerror("エラー", f"データベースファイルが見つかりません。\nパス: {db_path}\nconfig.iniを確認してください。")
            return None
        conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};' f'DBQ={db_path};')
        try:
            return pyodbc.connect(conn_str)
        except pyodbc.Error as ex:
            sqlstate = ex.args[0]
            messagebox.showerror("データベース接続エラー", f"エラーが発生しました: {sqlstate}\n{ex}")
            return None

    # --- 計算ロジック ---
    def start_calculation_thread(self):
        inputs = self._get_user_inputs()
        if not inputs:
            return
        self._setup_progress_window()
        thread = threading.Thread(target=self._calculation_worker, args=(inputs,))
        thread.daemon = True
        thread.start()

    def _get_user_inputs(self):
        inputs = {
            'product_number': self.app.sample_pn_entry.get(),
            'lot_size_str': self.app.sample_qty_entry.get(),
            'start_date': self.app.sample_start_date_entry.get().strip() or None,
            'end_date': self.app.sample_end_date_entry.get().strip() or None,
            'conf_str': self.app.sample_conf_entry.get().strip() or "99",
            'c_str': self.app.sample_c_entry.get().strip() or "0"
        }
        if not inputs['product_number']:
            messagebox.showwarning("入力エラー", "品番を入力してください。")
            return None
        try:
            inputs['lot_size'] = int(inputs['lot_size_str'])
            inputs['conf'] = float(inputs['conf_str']) / 100
            inputs['c'] = int(inputs['c_str'])
        except ValueError:
            messagebox.showwarning("入力エラー", "数量、信頼度、c値は数値で入力してください。")
            return None
        return inputs

    def _setup_progress_window(self):
        self.app.calc_button.config(state='disabled', text="🔄 計算中...", bg="#6c757d")
        self.progress_window = tk.Toplevel(self.app)
        self.progress_window.title("AI計算中")
        self.progress_window.geometry("400x200")
        self.progress_window.configure(bg="#ffffff")
        self.progress_window.resizable(False, False)
        x = (self.app.winfo_screenwidth() // 2) - (200)
        y = (self.app.winfo_screenheight() // 2) - (100)
        self.progress_window.geometry(f"400x200+{x}+{y}")
        
        progress_bar = ttk.Progressbar(self.progress_window, mode='indeterminate', length=300)
        progress_bar.pack(pady=20)
        progress_bar.start()

        tk.Label(self.progress_window, text="🤖 AIが統計計算を開始しました...", font=("Meiryo", 12, "bold"), fg="#2c3e50", bg="#ffffff").pack(pady=10)
        self.detail_label = tk.Label(self.progress_window, text="データベースから過去の不具合データを分析中", font=("Meiryo", 10), fg="#6c757d", bg="#ffffff")
        self.detail_label.pack(pady=5)
        
        self.app.result_var.set("")
        self.app.review_var.set("")
        self.app.best3_var.set("")
        for widget_name in ['main_sample_label', 'level_label', 'reason_label', 'advice_label']:
            if hasattr(self.app, widget_name):
                widget = getattr(self.app, widget_name)
                if widget:
                    widget.destroy()
        self.app.update_idletasks()

    def _calculation_worker(self, inputs):
        try:
            self.app.after(0, lambda: self.detail_label.config(text="データベースに接続中..."))
            conn = self._get_db_connection()
            if not conn:
                raise ConnectionError("DB接続に失敗")

            with conn.cursor() as cursor:
                self.app.after(0, lambda: self.detail_label.config(text="不具合データを集計中..."))
                db_data = self._fetch_data(cursor, inputs)
                
                self.app.after(0, lambda: self.detail_label.config(text="抜取検査数を計算中..."))
                stats_results = self._calculate_stats(db_data, inputs)
            
            self.app.after(0, lambda: self.detail_label.config(text="結果を表示中..."))
            self.app.after(0, self._update_ui, db_data, stats_results, inputs)
            self.app.after(0, self._finish_calculation, True)

        except Exception as e:
            if "DB接続に失敗" not in str(e):
                 self.app.after(0, lambda: messagebox.showerror("計算エラー", f"バックグラウンド処理中にエラーが発生しました: {e}"))
            self.app.after(0, self._finish_calculation, False)

    def _finish_calculation(self, success):
        if self.progress_window and self.progress_window.winfo_exists():
            self.progress_window.destroy()
        self.app.calc_button.config(state='normal', text="🚀 計算実行", bg="#007bff")
        if success:
            messagebox.showinfo("計算完了", "✅ AIが統計分析を完了しました！")

    def _build_sql_query(self, base_sql, inputs):
        sql_parts = [base_sql]
        params = [inputs['product_number']]
        has_where = ' where ' in base_sql.lower()
        if inputs['start_date']:
            sql_parts.append(f"{'AND' if has_where else 'WHERE'} [指示日] >= ?")
            params.append(inputs['start_date'])
            has_where = True
        if inputs['end_date']:
            sql_parts.append(f"{'AND' if has_where else 'WHERE'} [指示日] <= ?")
            params.append(inputs['end_date'])
        return " ".join(sql_parts), params

    def _fetch_data(self, cursor, inputs):
        data = {}
        sql, params = self._build_sql_query("SELECT SUM([数量]), SUM([総不具合数]) FROM t_不具合情報 WHERE [品番] = ?", inputs)
        sum_row = cursor.execute(sql, *params).fetchone()
        data['total_qty'] = sum_row[0] or 0
        data['total_defect'] = sum_row[1] or 0
        data['defect_rate'] = (data['total_defect'] / data['total_qty'] * 100) if data['total_qty'] else 0

        columns_str = ", ".join(f"SUM(IIF([{col}] IS NOT NULL AND [{col}]<>0, [{col}], 0))" for col in DEFECT_COLUMNS)
        sql, params = self._build_sql_query(f"SELECT {columns_str} FROM t_不具合情報 WHERE [品番] = ?", inputs)
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

    def _calculate_stats(self, db_data, inputs):
        results = {}
        p = db_data['defect_rate'] / 100
        if db_data['defect_rate'] == 0:
            results['level_text'] = "ゆるい(I)"
            results['level_reason'] = "過去の不具合が0件だったため、最もゆるい水準（I）を適用しています。"
        elif 0 < db_data['defect_rate'] <= 0.5:
            results['level_text'] = "普通(II)"
            results['level_reason'] = "過去の不具合率が0.5%以下だったため、普通水準（II）を適用しています。"
        else:
            results['level_text'] = "きつい(III)"
            results['level_reason'] = "過去の不具合率が0.5%を超えていたため、きつい水準（III）を適用しています。"
        
        n_sample = "計算不可"
        if p > 0 and 0 < inputs['conf'] < 1:
            try:
                if inputs['c'] == 0:
                    n_sample = math.ceil(math.log(1 - inputs['conf']) / math.log(1 - p))
                else:
                    n = 1
                    limit = max(inputs['lot_size'] * 2, 10000)
                    while n <= limit:
                        if binom.cdf(inputs['c'], n, p) >= 1 - inputs['conf']:
                            n_sample = n
                            break
                        n += 1
                    else:
                        n_sample = f">{limit} (計算断念)"
            except (ValueError, OverflowError):
                n_sample = "計算エラー"
        results['sample_size'] = n_sample
        return results

    def _update_ui(self, db_data, stats_results, inputs):
        def format_int(n):
            try: return f"{int(n):,}"
            except (ValueError, TypeError): return str(n)

        sample_size_disp = format_int(stats_results['sample_size'])
        
        # 以前のウィジェットを破棄
        for widget_name in ['main_sample_label', 'level_label', 'reason_label', 'advice_label']:
            if hasattr(self.app, widget_name):
                widget = getattr(self.app, widget_name)
                if widget:
                    widget.destroy()

        self.app.main_sample_label = tk.Label(self.app.result_frame, text=f"抜取検査数: {sample_size_disp} 個", font=("Meiryo", 32, "bold"), fg="#007bff", bg="#e9ecef", pady=10)
        self.app.main_sample_label.pack(pady=(10, 0))
        self.app.level_label = tk.Label(self.app.result_frame, text=f"検査水準: {stats_results['level_text']}", font=("Meiryo", 16, "bold"), fg="#2c3e50", bg="#e9ecef", pady=5)
        self.app.level_label.pack()
        self.app.reason_label = tk.Label(self.app.result_frame, text=f"根拠: {stats_results['level_reason']}", font=("Meiryo", 12), fg="#6c757d", bg="#e9ecef", pady=5, wraplength=800)
        self.app.reason_label.pack()

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
        self.app.review_var.set(review_text)

        if db_data['best5']:
            best5_text = "【検査時の注意喚起：過去不具合ベスト5】\n"
            for i, (naiyo, count) in enumerate(db_data['best5'], 1):
                rate = next((r for col, r, c_ in db_data['defect_rates_sorted'] if col == naiyo), 0)
                best5_text += f"{i}. {naiyo}（{format_int(count)}個, {rate:.2f}%）\n"
        else:
            best5_text = "【検査時の注意喚起】\n該当期間に不具合データがありません。"
        self.app.best3_var.set(best5_text)

        advice = ""
        if db_data['best5'] and db_data['best5'][0][1] > 0:
            advice = f"過去最多の不具合は『{db_data['best5'][0][0]}』です。検査時は特にこの点にご注意ください。"
        elif db_data['total_defect'] > 0:
            advice = "過去の不具合傾向から特に目立つ項目はありませんが、標準的な検査を心がけましょう。"
        else:
            advice = "過去の不具合データが少ないため、全般的に注意して検査を行ってください。"
        self.app.advice_label = tk.Label(self.app.sampling_frame, text=advice, font=("Meiryo", 9), fg="#dc3545", bg="#ffffff")
        self.app.advice_label.pack(after=self.app.result_label, pady=(0, 5))

if __name__ == "__main__":
    controller = MainController()
    controller.run()
