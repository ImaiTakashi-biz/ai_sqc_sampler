import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog
import pyodbc
import configparser
import os
from scipy.stats import binom
import math
import threading
from gui import App
from matplotlib.figure import Figure
import numpy as np

CONFIG_FILE = 'config.ini'
DB_FILE_PATH_KEY = 'path'
DB_SECTION = 'DATABASE'
DEFECT_COLUMNS = [
    "外観キズ", "圧痕", "切粉", "毟れ", "穴大", "穴小", "穴キズ", "バリ", "短寸", "面粗", "サビ", "ボケ", "挽目", "汚れ", "メッキ", "落下",
    "フクレ", "ツブレ", "ボッチ", "段差", "バレル石", "径プラス", "径マイナス", "ゲージ", "異物混入", "形状不良", "こすれ", "変色シミ", "材料キズ", "ゴミ", "その他"
]

class MainController:
    def __init__(self):
        self.app = App(self)
        self.progress_window = None
        self.detail_label = None
        self.last_db_data = None
        self.last_stats_results = None
        self.last_inputs = None

    def run(self):
        self.app.mainloop()

    def _get_db_path(self):
        config = configparser.ConfigParser()
        config.read(CONFIG_FILE, encoding='utf-8')

        if not os.path.exists(CONFIG_FILE):
            messagebox.showerror("設定エラー", f"設定ファイル '{CONFIG_FILE}' が見つかりません。アプリケーションを終了します。")
            self.app.quit()
            return None

        if DB_SECTION not in config or DB_FILE_PATH_KEY not in config[DB_SECTION]:
            messagebox.showerror("設定エラー", f"設定ファイル '{CONFIG_FILE}' にデータベースパスが設定されていません。アプリケーションを終了します。")
            self.app.quit()
            return None

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
        self.app.calc_button.config(state='disabled', text="🔄 計算中...", bg=self.app.MEDIUM_GRAY)
        self.progress_window = tk.Toplevel(self.app)
        self.progress_window.title("AI計算中")
        self.progress_window.geometry("400x200")
        self.progress_window.configure(bg=self.app.LIGHT_GRAY)
        self.progress_window.resizable(False, False)
        x = (self.app.winfo_screenwidth() // 2) - 200
        y = (self.app.winfo_screenheight() // 2) - 100
        self.progress_window.geometry(f"400x200+{x}+{y}")
        progress_bar = ttk.Progressbar(self.progress_window, mode='indeterminate', length=300)
        progress_bar.pack(pady=20)
        progress_bar.start()
        tk.Label(self.progress_window, text="🤖 AIが統計計算を開始しました...", font=("Meiryo", 12, "bold"), fg=self.app.DARK_GRAY, bg=self.app.LIGHT_GRAY).pack(pady=10)
        self.detail_label = tk.Label(self.progress_window, text="データベースから過去の不具合データを分析中", font=("Meiryo", 10), fg=self.app.MEDIUM_GRAY, bg=self.app.LIGHT_GRAY)
        self.detail_label.pack(pady=5)
        self.app.result_var.set("")
        self.app.review_var.set("")
        self.app.best3_var.set("")
        self.app.update_idletasks()

    def _calculation_worker(self, inputs):
        try:
            self.app.after(0, lambda: self.detail_label.config(text="データベースに接続中..."))
            conn = self._get_db_connection()
            if not conn: raise ConnectionError("DB接続に失敗")
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
        self.app.calc_button.config(state='normal', text="🚀 計算実行", bg=self.app.PRIMARY_BLUE)
        if success:
            messagebox.showinfo("計算完了", "✅ AIが統計分析を完了しました！")

    def _build_sql_query(self, base_sql, inputs):
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
        data = {'total_qty': 0, 'total_defect': 0, 'defect_rate': 0, 'defect_rates_sorted': [], 'best5': []}
        defect_columns_sum = ", ".join(f"SUM(IIF([{col}] IS NOT NULL AND [{col}]<>0, [{col}], 0))" for col in DEFECT_COLUMNS)
        base_sql = f"SELECT SUM([数量]), SUM([総不具合数]), {defect_columns_sum} FROM t_不具合情報 WHERE [品番] = ?"
        sql, params = self._build_sql_query(base_sql, inputs)
        row = cursor.execute(sql, *params).fetchone()
        if not row or row[0] is None: return data
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
                    low, high = 1, max(inputs['lot_size'] * 2, 10000)
                    n_sample = f">{high} (計算断念)"
                    while low <= high:
                        mid = (low + high) // 2
                        if mid == 0: low = 1; continue
                        if binom.cdf(inputs['c'], mid, p) >= 1 - inputs['conf']:
                            n_sample, high = mid, mid - 1
                        else:
                            low = mid + 1
            except (ValueError, OverflowError): n_sample = "計算エラー"
        results['sample_size'] = n_sample
        return results

    def _update_ui(self, db_data, stats_results, inputs):
        self._clear_previous_results()
        self.last_db_data, self.last_stats_results, self.last_inputs = db_data, stats_results, inputs
        texts = self._generate_result_texts(db_data, stats_results, inputs)
        self._display_main_results(stats_results, texts['advice'])
        self._display_detailed_results(texts)

    def _clear_previous_results(self):
        for widget_name in ['main_sample_label', 'level_label', 'reason_label', 'advice_label']:
            if hasattr(self.app, widget_name) and (widget := getattr(self.app, widget_name)): widget.destroy()
        self.app.review_frame.pack_forget()
        self.app.best3_frame.pack_forget()
        if hasattr(self.app, 'hide_export_button'): self.app.hide_export_button()
        self.last_db_data, self.last_stats_results, self.last_inputs = None, None, None

    def _format_int(self, n):
        try: return f"{int(n):,}"
        except (ValueError, TypeError): return str(n)

    def _generate_result_texts(self, db_data, stats_results, inputs):
        sample_size_disp = self._format_int(stats_results['sample_size'])
        period_text = f"（{inputs['start_date'] or '最初'}～{inputs['end_date'] or '最新'}）" if inputs['start_date'] or inputs['end_date'] else "（全期間対象）"
        review_text = (
            f"【根拠レビュー】\n・ロットサイズ: {self._format_int(inputs['lot_size'])}\n・対象期間: {period_text}\n"
            f"・数量合計: {self._format_int(db_data['total_qty'])}個\n・不具合数合計: {self._format_int(db_data['total_defect'])}個\n"
            f"・不良率: {db_data['defect_rate']:.2f}%\n・信頼度: {inputs['conf']*100:.1f}%\n・c値: {inputs['c']}\n"
            f"・推奨抜取検査数: {sample_size_disp} 個\n（c={inputs['c']}, 信頼度={inputs['conf']*100:.1f}%の条件で自動計算）"
        )
        if db_data['best5']:
            best5_text = "【検査時の注意喚起：過去不具合ベスト5】\n"
            for i, (naiyo, count) in enumerate(db_data['best5'], 1):
                rate = next((r for col, r, c_ in db_data['defect_rates_sorted'] if col == naiyo), 0)
                best5_text += f"{i}. {naiyo}（{self._format_int(count)}個, {rate:.2f}%）\n"
        else: best5_text = "【検査時の注意喚起】\n該当期間に不具合データがありません。"
        if db_data['best5'] and db_data['best5'][0][1] > 0:
            advice = f"過去最多の不具合は『{db_data['best5'][0][0]}』です。検査時は特にこの点にご注意ください。"
        elif db_data['total_defect'] > 0: advice = "過去の不具合傾向から特に目立つ項目はありませんが、標準的な検査を心がけましょう。"
        else: advice = "過去の不具合データが少ないため、全般的に注意して検査を行ってください。"
        return {'review': review_text, 'best5': best5_text, 'advice': advice}

    def _display_main_results(self, stats_results, advice_text):
        sample_size_disp = self._format_int(stats_results['sample_size'])
        self.app.main_sample_label = tk.Label(self.app.result_frame, text=f"抜取検査数: {sample_size_disp} 個", font=("Meiryo", 32, "bold"), fg="#007bff", bg="#e9ecef", pady=10)
        self.app.main_sample_label.pack(pady=(10, 0))
        self.app.level_label = tk.Label(self.app.result_frame, text=f"検査水準: {stats_results['level_text']}", font=("Meiryo", 16, "bold"), fg="#2c3e50", bg="#e9ecef", pady=5)
        self.app.level_label.pack()
        self.app.reason_label = tk.Label(self.app.result_frame, text=f"根拠: {stats_results['level_reason']}", font=("Meiryo", 12), fg="#6c757d", bg="#e9ecef", pady=5, wraplength=800)
        self.app.reason_label.pack()
        self.app.advice_label = tk.Label(self.app.sampling_frame, text=advice_text, font=("Meiryo", 9), fg=self.app.WARNING_RED, bg=self.app.LIGHT_GRAY, wraplength=800, justify='left', padx=15, pady=8, relief="flat", bd=1)
        self.app.advice_label.pack(after=self.app.result_label, pady=(0, 5))

    def _display_detailed_results(self, texts):
        self.app.review_var.set(texts['review'])
        self.app.review_frame.pack(fill='x', padx=40, pady=10)
        self.app.best3_var.set(texts['best5'])
        self.app.best3_frame.pack(fill='x', padx=40, pady=10)
        if hasattr(self.app, 'show_export_button'): self.app.show_export_button()

    def export_results(self):
        if not self.last_db_data: messagebox.showinfo("エクスポート不可", "先に計算を実行してください。"); return
        texts = self._generate_result_texts(self.last_db_data, self.last_stats_results, self.last_inputs)
        sample_size_disp = self._format_int(self.last_stats_results['sample_size'])
        content = (
            f"--- 抜取検査数計算結果 ---\n\n"
            f"品番: {self.last_inputs['product_number']}\nロットサイズ: {self._format_int(self.last_inputs['lot_size'])}\n\n"
            f"【推奨抜取検査数】\n{sample_size_disp} 個\n\n"
            f"【検査水準】\n{self.last_stats_results['level_text']}\n根拠: {self.last_stats_results['level_reason']}\n\n"
            f"{texts['review']}\n\n{texts['best5']}\n\n"
            f"【AIからのアドバイス】\n{texts['advice']}\n"
        )
        try:
            filepath = filedialog.asksaveasfilename(
                title="結果を名前を付けて保存",defaultextension=".txt",
                filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*รา")],
                initialfile=f"検査結果_{self.last_inputs['product_number']}.txt"
            )
            if not filepath: return
            with open(filepath, 'w', encoding='utf-8') as f: f.write(content)
            messagebox.showinfo("成功", f"結果を保存しました。\nパス: {filepath}")
        except Exception as e:
            messagebox.showerror("エクスポート失敗", f"ファイルの保存中にエラーが発生しました: {e}")

    def _fetch_all_product_numbers(self):
        conn = self._get_db_connection()
        if not conn: return []
        try:
            with conn.cursor() as cursor:
                sql = "SELECT DISTINCT [品番] FROM t_不具合情報 ORDER BY [品番]"
                rows = cursor.execute(sql).fetchall()
                return [row[0] for row in rows if row[0]]
        except pyodbc.Error as e:
            messagebox.showerror("データベースエラー", f"品番リストの取得中にエラーが発生しました: {e}")
            return []
        finally:
            if conn: conn.close()

    def show_product_numbers_list(self):
        product_numbers = self._fetch_all_product_numbers()
        if not product_numbers: messagebox.showinfo("情報", "表示できる品番がありません。"); return
        win = tk.Toplevel(self.app)
        win.title("品番リスト")
        win.geometry("300x400")
        search_frame = tk.Frame(win); search_frame.pack(fill='x', padx=5, pady=5)
        tk.Label(search_frame, text="検索:").pack(side='left')
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var); search_entry.pack(fill='x', expand=True)
        list_frame = tk.Frame(win); list_frame.pack(fill='both', expand=True, padx=5, pady=5)
        scrollbar = tk.Scrollbar(list_frame, orient='vertical')
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)
        scrollbar.pack(side='right', fill='y'); listbox.pack(side='left', fill='both', expand=True)
        for pn in product_numbers: listbox.insert('end', pn)
        def update_listbox(*args):
            search_term = search_var.get().lower()
            listbox.delete(0, 'end')
            for pn in product_numbers:
                if search_term in pn.lower(): listbox.insert('end', pn)
        search_var.trace("w", update_listbox)
        def on_double_click(event):
            selected_indices = listbox.curselection()
            if not selected_indices: return
            selected_pn = listbox.get(selected_indices[0])
            self.app.sample_pn_entry.delete(0, 'end'); self.app.sample_pn_entry.insert(0, selected_pn)
            win.destroy()
        listbox.bind("<Double-1>", on_double_click)
        win.transient(self.app)
        win.grab_set()
        self.app.wait_window(win)

if __name__ == "__main__":
    controller = MainController()
    controller.run()