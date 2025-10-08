"""
設定ダイアログモジュール
アプリケーションの設定を変更するためのダイアログ
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os


class SettingsDialog:
    """設定ダイアログクラス"""
    
    def __init__(self, parent, config_manager):
        self.parent = parent
        self.config_manager = config_manager
        self.dialog = None
        self.db_path_var = None
        self.confidence_var = None
        self.c_value_var = None
        
    def show(self):
        """設定ダイアログの表示"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("設定")
        self.dialog.geometry("500x600")
        self.dialog.configure(bg="#f0f0f0")
        self.dialog.resizable(True, True)
        self.dialog.minsize(500, 600)
        
        # モーダル表示
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # 中央配置
        self._center_dialog()
        
        # ダイアログの構築
        self._create_widgets()
        
        # フォーカス設定
        self.dialog.focus_set()
        
        # ダイアログが閉じられるまで待機
        self.parent.wait_window(self.dialog)
    
    def _center_dialog(self):
        """ダイアログを中央に配置"""
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - 250
        y = (self.dialog.winfo_screenheight() // 2) - 300
        self.dialog.geometry(f"500x600+{x}+{y}")
    
    def _create_widgets(self):
        """ウィジェットの作成"""
        # メインフレーム
        main_frame = tk.Frame(self.dialog, bg="#f0f0f0")
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # タイトル
        title_label = tk.Label(
            main_frame, 
            text="🔧 アプリケーション設定", 
            font=("Meiryo", 16, "bold"), 
            fg="#2c3e50", 
            bg="#f0f0f0"
        )
        title_label.pack(pady=(0, 20))
        
        # データベース設定セクション
        db_frame = tk.LabelFrame(
            main_frame, 
            text="データベース設定", 
            font=("Meiryo", 12, "bold"), 
            fg="#2c3e50", 
            bg="#f0f0f0",
            padx=10,
            pady=10
        )
        db_frame.pack(fill='x', pady=(0, 15))
        
        # データベースパス
        tk.Label(
            db_frame, 
            text="データベースファイル:", 
            font=("Meiryo", 10), 
            fg="#2c3e50", 
            bg="#f0f0f0"
        ).pack(anchor='w', pady=(0, 5))
        
        path_frame = tk.Frame(db_frame, bg="#f0f0f0")
        path_frame.pack(fill='x', pady=(0, 10))
        
        self.db_path_var = tk.StringVar(value=self.config_manager.get_database_path())
        path_entry = tk.Entry(
            path_frame, 
            textvariable=self.db_path_var, 
            font=("Meiryo", 9), 
            bg="#ffffff", 
            fg="#2c3e50", 
            relief="flat", 
            bd=1, 
            highlightthickness=1, 
            highlightbackground="#bdc3c7", 
            highlightcolor="#3498db"
        )
        path_entry.pack(side='left', fill='x', expand=True, padx=(0, 5))
        
        browse_button = tk.Button(
            path_frame, 
            text="参照...", 
            command=self._browse_database_file, 
            font=("Meiryo", 9), 
            bg="#3498db", 
            fg="#ffffff", 
            relief="flat", 
            padx=10, 
            pady=2
        )
        browse_button.pack(side='right')
        
        # データベース接続テスト
        test_button = tk.Button(
            db_frame, 
            text="🔍 データベース接続テスト", 
            command=self._test_database_connection, 
            font=("Meiryo", 9), 
            bg="#2ecc71", 
            fg="#ffffff", 
            relief="flat", 
            padx=10, 
            pady=5
        )
        test_button.pack(pady=(5, 0))
        
        # デフォルト値設定セクション
        default_frame = tk.LabelFrame(
            main_frame, 
            text="デフォルト値設定", 
            font=("Meiryo", 12, "bold"), 
            fg="#2c3e50", 
            bg="#f0f0f0",
            padx=10,
            pady=10
        )
        default_frame.pack(fill='x', pady=(0, 15))
        
        # 信頼度
        conf_frame = tk.Frame(default_frame, bg="#f0f0f0")
        conf_frame.pack(fill='x', pady=(0, 5))
        
        tk.Label(
            conf_frame, 
            text="デフォルト信頼度(%):", 
            font=("Meiryo", 10), 
            fg="#2c3e50", 
            bg="#f0f0f0"
        ).pack(side='left')
        
        self.confidence_var = tk.StringVar(value=str(self.config_manager.get("default_confidence", 99.0)))
        conf_entry = tk.Entry(
            conf_frame, 
            textvariable=self.confidence_var, 
            width=10, 
            font=("Meiryo", 10), 
            bg="#ffffff", 
            fg="#2c3e50", 
            relief="flat", 
            bd=1, 
            highlightthickness=1, 
            highlightbackground="#bdc3c7", 
            highlightcolor="#3498db"
        )
        conf_entry.pack(side='right')
        
        # c値
        c_frame = tk.Frame(default_frame, bg="#f0f0f0")
        c_frame.pack(fill='x', pady=(0, 5))
        
        tk.Label(
            c_frame, 
            text="デフォルトc値:", 
            font=("Meiryo", 10), 
            fg="#2c3e50", 
            bg="#f0f0f0"
        ).pack(side='left')
        
        self.c_value_var = tk.StringVar(value=str(self.config_manager.get("default_c_value", 0)))
        c_entry = tk.Entry(
            c_frame, 
            textvariable=self.c_value_var, 
            width=10, 
            font=("Meiryo", 10), 
            bg="#ffffff", 
            fg="#2c3e50", 
            relief="flat", 
            bd=1, 
            highlightthickness=1, 
            highlightbackground="#bdc3c7", 
            highlightcolor="#3498db"
        )
        c_entry.pack(side='right')
        
        # ボタンフレーム
        button_frame = tk.Frame(main_frame, bg="#f0f0f0")
        button_frame.pack(fill='x', pady=(30, 20))
        
        # リセットボタン
        reset_button = tk.Button(
            button_frame, 
            text="🔄 デフォルトに戻す", 
            command=self._reset_to_defaults, 
            font=("Meiryo", 10), 
            bg="#e74c3c", 
            fg="#ffffff", 
            relief="flat", 
            padx=15, 
            pady=5
        )
        reset_button.pack(side='left')
        
        # ボタン間のスペース
        tk.Frame(button_frame, bg="#f0f0f0", width=20).pack(side='left')
        
        # キャンセルボタン
        cancel_button = tk.Button(
            button_frame, 
            text="キャンセル", 
            command=self._cancel, 
            font=("Meiryo", 10), 
            bg="#95a5a6", 
            fg="#ffffff", 
            relief="flat", 
            padx=15, 
            pady=5
        )
        cancel_button.pack(side='right')
        
        # OKボタン
        ok_button = tk.Button(
            button_frame, 
            text="OK", 
            command=self._ok, 
            font=("Meiryo", 10, "bold"), 
            bg="#3498db", 
            fg="#ffffff", 
            relief="flat", 
            padx=15, 
            pady=5
        )
        ok_button.pack(side='right', padx=(0, 5))
        
        # デフォルトフォーカス
        ok_button.focus_set()
    
    def _browse_database_file(self):
        """データベースファイルの参照"""
        try:
            current_path = self.db_path_var.get()
            initial_dir = os.path.dirname(current_path) if os.path.dirname(current_path) else os.getcwd()
            
            file_path = filedialog.askopenfilename(
                parent=self.dialog,
                title="データベースファイルを選択",
                initialdir=initial_dir,
                filetypes=[
                    ("Access Database", "*.accdb"),
                    ("Access Database (Legacy)", "*.mdb"),
                    ("すべてのファイル", "*.*")
                ]
            )
            
            if file_path:
                # 相対パスに変換（可能な場合）
                try:
                    rel_path = os.path.relpath(file_path, os.getcwd())
                    if not rel_path.startswith('..'):
                        file_path = rel_path
                except ValueError:
                    pass  # 相対パスに変換できない場合は絶対パスを使用
                
                self.db_path_var.set(file_path)
                
        except Exception as e:
            messagebox.showerror("エラー", f"ファイル選択中にエラーが発生しました:\n{str(e)}")
    
    def _test_database_connection(self):
        """データベース接続のテスト"""
        try:
            # 一時的に設定を更新
            temp_path = self.db_path_var.get()
            if not os.path.exists(temp_path):
                messagebox.showerror("エラー", f"ファイルが存在しません:\n{temp_path}")
                return
            
            # データベース接続テスト
            import pyodbc
            conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                       f'DBQ={temp_path};'
                       r'ReadOnly=False;'
                       r'Exclusive=False;')
            
            conn = pyodbc.connect(conn_str)
            with conn.cursor() as cursor:
                cursor.execute("SELECT COUNT(*) FROM t_不具合情報")
                count = cursor.fetchone()[0]
            conn.close()
            
            messagebox.showinfo("接続テスト成功", f"✅ データベース接続に成功しました！\n\nレコード数: {count}件")
            
        except pyodbc.Error as e:
            if "Microsoft Access Driver" in str(e):
                messagebox.showerror("ドライバーエラー", 
                    "Microsoft Access Driverが見つかりません。\n"
                    "Microsoft Access Database Engine 2016 Redistributableをインストールしてください。")
            else:
                messagebox.showerror("接続エラー", f"データベース接続に失敗しました:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("エラー", f"接続テスト中にエラーが発生しました:\n{str(e)}")
    
    def _reset_to_defaults(self):
        """設定をデフォルトにリセット"""
        if messagebox.askyesno("確認", "設定をデフォルト値に戻しますか？"):
            self.config_manager.reset_to_defaults()
            self.db_path_var.set(self.config_manager.get_database_path())
            self.confidence_var.set(str(self.config_manager.get("default_confidence", 99.0)))
            self.c_value_var.set(str(self.config_manager.get("default_c_value", 0)))
            messagebox.showinfo("完了", "設定をデフォルト値に戻しました。")
    
    def _ok(self):
        """OKボタンの処理"""
        try:
            # 入力値の検証
            confidence = float(self.confidence_var.get())
            c_value = int(self.c_value_var.get())
            
            if not 0 < confidence <= 100:
                messagebox.showerror("エラー", "信頼度は0より大きく100以下の数値で入力してください。")
                return
            
            if c_value < 0:
                messagebox.showerror("エラー", "c値は0以上の整数で入力してください。")
                return
            
            # 設定の保存
            self.config_manager.set_database_path(self.db_path_var.get())
            self.config_manager.set("default_confidence", confidence)
            self.config_manager.set("default_c_value", c_value)
            
            messagebox.showinfo("完了", "設定を保存しました。")
            self.dialog.destroy()
            
        except ValueError:
            messagebox.showerror("エラー", "入力値が正しくありません。数値を確認してください。")
        except Exception as e:
            messagebox.showerror("エラー", f"設定の保存中にエラーが発生しました:\n{str(e)}")
    
    def _cancel(self):
        """キャンセルボタンの処理"""
        self.dialog.destroy()
