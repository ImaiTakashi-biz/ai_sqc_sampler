"""
品番リスト管理モジュール
品番リストの表示と管理を担当
"""

import tkinter as tk
from tkinter import messagebox
import threading
from security_manager import SecurityManager


class ProductListManager:
    """品番リスト管理クラス"""
    
    def __init__(self, app, db_manager):
        self.app = app
        self.db_manager = db_manager
        self.security_manager = SecurityManager()
    
    def show_product_numbers_list(self):
        """品番リストの表示（非同期読み込み）"""
        thread = threading.Thread(target=self.load_product_numbers_async)
        thread.daemon = True
        thread.start()

    def load_product_numbers_async(self):
        """品番リストの非同期読み込み"""
        progress_window = None
        try:
            # プログレスウィンドウの作成
            progress_window = tk.Toplevel(self.app)
            progress_window.title("品番リスト読み込み中...")
            progress_window.geometry("350x120")
            progress_window.configure(bg="#f0f0f0")
            progress_window.resizable(False, False)
            
            # 中央配置
            x = (self.app.winfo_screenwidth() // 2) - 175
            y = (self.app.winfo_screenheight() // 2) - 60
            progress_window.geometry(f"350x120+{x}+{y}")
            
            # プログレスバー
            progress_bar = tk.ttk.Progressbar(progress_window, mode='indeterminate', length=250)
            progress_bar.pack(pady=20)
            progress_bar.start()
            
            # ステータスラベル
            status_label = tk.Label(
                progress_window, 
                text="データベースから品番リストを読み込み中...", 
                font=("Meiryo", 10), 
                bg="#f0f0f0"
            )
            status_label.pack(pady=5)
            
            # 品番リストの取得
            product_numbers = self.db_manager.fetch_all_product_numbers()
            
            # プログレスウィンドウを閉じる
            if progress_window and progress_window.winfo_exists():
                progress_window.destroy()
            
            # 結果の表示
            self.app.after(0, self.show_product_numbers_result, product_numbers)
            
        except Exception as e:
            if progress_window and progress_window.winfo_exists():
                progress_window.destroy()
            sanitized_error = self.security_manager.sanitize_error_message(str(e))
            error_message = f"品番リストの読み込み中にエラーが発生しました:\n{sanitized_error}"
            self.app.after(0, lambda: messagebox.showerror("エラー", error_message))

    def show_product_numbers_result(self, product_numbers):
        """品番リストの結果表示"""
        if not product_numbers:
            messagebox.showinfo("情報", "表示できる品番がありません。")
            return
        
        # 品番リストウィンドウの作成
        win = tk.Toplevel(self.app)
        win.title(f"品番リスト ({len(product_numbers)}件)")
        win.geometry("400x500")
        win.configure(bg="#f0f0f0")
        
        # 検索フレーム
        search_frame = tk.Frame(win, bg="#f0f0f0")
        search_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(search_frame, text="🔍 検索:", font=("Meiryo", 10), bg="#f0f0f0").pack(side='left', padx=(0, 5))
        
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, font=("Meiryo", 10))
        search_entry.pack(fill='x', expand=True)
        
        # リストフレーム
        list_frame = tk.Frame(win, bg="#f0f0f0")
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # スクロールバー
        scrollbar = tk.Scrollbar(list_frame, orient='vertical')
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=("Meiryo", 10))
        scrollbar.config(command=listbox.yview)
        
        scrollbar.pack(side='right', fill='y')
        listbox.pack(side='left', fill='both', expand=True)
        
        # 検索可能なアイテムの準備
        searchable_items = [(pn, pn.lower()) for pn in product_numbers]
        
        # 初期表示
        for pn, _ in searchable_items:
            listbox.insert('end', pn)
        
        # 検索機能
        def update_listbox(*args):
            search_term = search_var.get().strip().lower()
            listbox.delete(0, 'end')
            filtered_count = 0
            
            for pn, pn_lower in searchable_items:
                if not search_term or search_term in pn_lower:
                    listbox.insert('end', pn)
                    filtered_count += 1
            
            win.title(f"品番リスト ({filtered_count}件)")
        
        search_var.trace("w", update_listbox)
        
        # ダブルクリックで選択
        def on_double_click(event):
            selected_indices = listbox.curselection()
            if not selected_indices:
                return
            
            selected_pn = listbox.get(selected_indices[0])
            self.app.sample_pn_entry.delete(0, 'end')
            self.app.sample_pn_entry.insert(0, selected_pn)
            win.destroy()
        
        listbox.bind("<Double-1>", on_double_click)
        
        # モーダル表示
        win.transient(self.app)
        win.grab_set()
        search_entry.focus_set()
        
        # 中央配置
        win.update_idletasks()
        x = (self.app.winfo_screenwidth() // 2) - 200
        y = (self.app.winfo_screenheight() // 2) - 250
        win.geometry(f"400x500+{x}+{y}")
        
        self.app.wait_window(win)

