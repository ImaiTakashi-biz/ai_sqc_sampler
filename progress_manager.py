"""
プログレス管理モジュール
プログレス表示とスレッド管理を担当
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import pyodbc
from error_handler import error_handler, ErrorCode
from security_manager import SecurityManager


class ProgressManager:
    """プログレス管理クラス"""
    
    def __init__(self, app, db_manager, calculation_engine, ui_manager):
        self.app = app
        self.db_manager = db_manager
        self.calculation_engine = calculation_engine
        self.ui_manager = ui_manager
        self.security_manager = SecurityManager()
    
    def _set_status(self, text):
        """プログレスウィンドウのステータス表示を更新"""
        if not hasattr(self, 'status_label'):
            return
        def _update():
            self.status_label.config(text=text)
            self.status_label.update_idletasks()
        self.app.after(0, _update)
    
    def setup_progress_window(self):
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
        tk.Label(
            self.progress_window,
            text="処理が完了すると自動的に閉じます。",
            font=("Meiryo", 9),
            fg="#6c757d",
            bg="#f0f0f0"
        ).pack()
        
        # モーダル表示
        self.progress_window.transient(self.app)
        self.progress_window.grab_set()

    def start_calculation_thread(self, inputs):
        """計算処理を別スレッドで開始"""
        if hasattr(self.app, 'calc_button'):
            self.app.calc_button.config(state='disabled', text="計算中...", bg=self.app.MEDIUM_GRAY)
        self.setup_progress_window()
        thread = threading.Thread(target=self.calculation_worker, args=(inputs,))
        thread.daemon = True
        thread.start()

    def calculation_worker(self, inputs):
        """計算処理のワーカースレッド（接続プール対応）"""
        conn = None
        try:
            # ステータス更新
            self._set_status("1/4 データベースに接続中...")
            
            # データベース接続（接続プール使用）
            conn = self.db_manager.get_db_connection()
            if not conn:
                error_handler.handle_error(
                    ErrorCode.DB_CONNECTION_FAILED, 
                    Exception("データベース接続に失敗しました")
                )
                self.app.after(0, self.finish_calculation, False)
                return
            
            with conn.cursor() as cursor:
                # ステータス更新
                self._set_status("2/4 不具合データを集計中...")
                
                # データの取得
                db_data = self.calculation_engine.fetch_data(cursor, inputs)
                
                # ステータス更新
                self._set_status("3/4 抜取検査数を計算中...")
                
                # 統計計算
                stats_results = self.calculation_engine.calculate_stats(db_data, inputs)
                
            # ステータス更新
            self._set_status("4/4 結果を表示中...")
            
            # UI更新
            self.app.after(0, self.ui_manager.update_ui, db_data, stats_results, inputs)
            self.app.after(0, self.finish_calculation, True, db_data, stats_results, inputs)
            
        except pyodbc.Error as e:
            error_handler.handle_error(ErrorCode.DB_QUERY_FAILED, e)
            self.app.after(0, self.finish_calculation, False)
            
        except ValueError as e:
            error_handler.handle_error(ErrorCode.CALCULATION_ERROR, e)
            self.app.after(0, self.finish_calculation, False)
            
        except OverflowError as e:
            error_handler.handle_error(ErrorCode.CALCULATION_OVERFLOW, e)
            self.app.after(0, self.finish_calculation, False)
            
        except Exception as e:
            error_handler.handle_error(ErrorCode.SYSTEM_ERROR, e)
            self.app.after(0, self.finish_calculation, False)
        finally:
            # 接続をプールに返却
            if conn:
                self.db_manager.return_db_connection(conn)

    def finish_calculation(self, success, db_data=None, stats_results=None, inputs=None):
        """計算完了処理"""
        if hasattr(self, 'progress_window') and self.progress_window.winfo_exists():
            self.progress_window.destroy()
        if hasattr(self.app, 'calc_button'):
            self.app.calc_button.config(state='normal', text="🚀 計算実行", bg=self.app.PRIMARY_BLUE)
        if success:
            messagebox.showinfo("計算完了", "✅ AIが統計分析を完了しました！")
            # 結果をコントローラーに保存
            if hasattr(self.app, 'controller'):
                self.app.controller.last_db_data = db_data
                self.app.controller.last_stats_results = stats_results
                self.app.controller.last_inputs = inputs

    def close_progress_window(self):
        """プログレスウィンドウを閉じる"""
        if hasattr(self, 'progress_window') and self.progress_window.winfo_exists():
            self.progress_window.destroy()

