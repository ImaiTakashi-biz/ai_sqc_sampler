"""
設定ダイアログモジュール
アプリケーションの設定を変更するためのダイアログ
"""

import tkinter as tk
from tkinter import messagebox, filedialog
import os
from security_manager import SecurityManager


class SettingsDialog:
    """設定ダイアログクラス"""
    
    def __init__(self, parent, config_manager):
        self.parent = parent
        self.config_manager = config_manager
        self.security_manager = getattr(config_manager, "security_manager", SecurityManager())
        self.dialog = None
        self.db_path_var = None
        self.preset_vars = {}
        
    def show(self):
        """設定ダイアログの表示"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("アプリケーション設定")
        self.dialog.geometry("650x600")
        self.dialog.configure(bg="#f0f0f0")
        self.dialog.resizable(True, True)
        self.dialog.minsize(650, 500)
        
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
        x = (self.dialog.winfo_screenwidth() // 2) - 325
        y = (self.dialog.winfo_screenheight() // 2) - 300
        self.dialog.geometry(f"650x600+{x}+{y}")
    
    def _create_widgets(self):
        """ウィジェットの作成"""
        # スクロール可能なフレームの作成
        canvas = tk.Canvas(self.dialog, bg="#f0f0f0", highlightthickness=0)
        scrollbar = tk.Scrollbar(self.dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#f0f0f0")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # レイアウト
        canvas.pack(side="left", fill="both", expand=True, padx=(15, 0), pady=15)
        scrollbar.pack(side="right", fill="y", pady=15)
        
        # スクロール可能フレームの幅を調整
        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # キャンバスの幅に合わせてスクロール可能フレームの幅を調整
            canvas_width = canvas.winfo_width()
            if canvas_width > 1:  # キャンバスが初期化されている場合のみ
                canvas.itemconfig(canvas.find_all()[0], width=canvas_width)
        
        scrollable_frame.bind("<Configure>", configure_scroll_region)
        canvas.bind("<Configure>", configure_scroll_region)
        
        # マウスホイールでのスクロール機能
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_to_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def _unbind_from_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
        
        canvas.bind('<Enter>', _bind_to_mousewheel)
        canvas.bind('<Leave>', _unbind_from_mousewheel)
        
        # メインフレーム（スクロール可能フレーム内）
        main_frame = scrollable_frame
        
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
            padx=15,
            pady=12
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
        
        # 検査区分デフォルト値設定セクション
        presets_frame = tk.LabelFrame(
            main_frame,
            text="検査区分ごとのデフォルト値設定",
            font=("Meiryo", 12, "bold"),
            fg="#2c3e50",
            bg="#f0f0f0",
            padx=15,
            pady=12
        )
        presets_frame.pack(fill='x', pady=(0, 15))
        
        # 説明ラベル
        explanation_label = tk.Label(
            presets_frame,
            text="各検査区分（緩和検査・標準検査・強化検査）のAQL、LTPD、α、β、c値を設定できます。\nこれらの値は検査区分選択時に自動的に適用されます。",
            font=("Meiryo", 9),
            fg="#7f8c8d",
            bg="#f0f0f0",
            wraplength=580,
            justify='left'
        )
        explanation_label.pack(anchor='w', pady=(0, 10))

        self.preset_vars = {}
        param_specs = [
            ("aql", "AQL(%)", float),
            ("ltpd", "LTPD(%)", float),
            ("alpha", "α(%)（生産者危険）", float),
            ("beta", "β(%)（消費者危険）", float),
            ("c_value", "c値", int)
        ]

        for mode_key, label in self.config_manager.get_inspection_mode_choices().items():
            details = self.config_manager.get_inspection_mode_details(mode_key)
            mode_frame = tk.LabelFrame(
                presets_frame,
                text=f"{label}のデフォルト値",
                font=("Meiryo", 11, "bold"),
                fg="#2c3e50",
                bg="#f0f0f0",
                padx=12,
                pady=10
            )
            mode_frame.pack(fill='x', pady=(0, 12))

            self.preset_vars[mode_key] = {}
            for param_key, param_label, caster in param_specs:
                row = tk.Frame(mode_frame, bg="#f0f0f0")
                row.pack(fill='x', pady=(0, 4))

                tk.Label(
                    row,
                    text=param_label,
                    font=("Meiryo", 9),
                    fg="#2c3e50",
                    bg="#f0f0f0"
                ).pack(side='left')

                var = tk.StringVar(value=str(details.get(param_key, "")))
                entry = tk.Entry(
                    row,
                    textvariable=var,
                    width=12,
                    font=("Meiryo", 9),
                    bg="#ffffff",
                    fg="#2c3e50",
                    relief="flat",
                    bd=1,
                    highlightthickness=1,
                    highlightbackground="#bdc3c7",
                    highlightcolor="#3498db"
                )
                entry.pack(side='right', padx=(10, 0))
                self.preset_vars[mode_key][param_key] = (var, caster, param_label)
        
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
            sanitized_error = self.security_manager.sanitize_error_message(str(e))
            messagebox.showerror("エラー", f"ファイル選択中にエラーが発生しました:\n{sanitized_error}")
    
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
                sanitized_error = self.security_manager.sanitize_error_message(str(e))
                messagebox.showerror("接続エラー", f"データベース接続に失敗しました:\n{sanitized_error}")
        except Exception as e:
            sanitized_error = self.security_manager.sanitize_error_message(str(e))
            messagebox.showerror("エラー", f"接続テスト中にエラーが発生しました:\n{sanitized_error}")
    
    def _reset_to_defaults(self):
        """設定をデフォルトにリセット"""
        if messagebox.askyesno("確認", "設定をデフォルト値に戻しますか？"):
            self.config_manager.reset_to_defaults()
            self.db_path_var.set(self.config_manager.get_database_path())
            for mode_key, param_map in self.preset_vars.items():
                details = self.config_manager.get_inspection_mode_details(mode_key)
                for param_key, (var, _, _) in param_map.items():
                    var.set(str(details.get(param_key, "")))
            messagebox.showinfo("完了", "設定をデフォルト値に戻しました。")
    
    def _ok(self):
        """OKボタンの処理"""
        try:
            if not self.config_manager.set_database_path(self.db_path_var.get()):
                return

            for mode_key, param_map in self.preset_vars.items():
                values = {}
                for param_key, (var, caster, label) in param_map.items():
                    raw_value = var.get().strip()
                    if not raw_value:
                        messagebox.showerror("エラー", f"{label} を入力してください。")
                        return
                    try:
                        value = caster(raw_value)
                    except ValueError:
                        messagebox.showerror("エラー", f"{label} には数値を入力してください。")
                        return
                    if caster is float and value < 0:
                        messagebox.showerror("エラー", f"{label} は0以上の数値で入力してください。")
                        return
                    if param_key in ("alpha", "beta") and not (0 <= value <= 100):
                        messagebox.showerror("エラー", f"{label} は0以上100以下の範囲で入力してください。")
                        return
                    if param_key == "c_value" and value < 0:
                        messagebox.showerror("エラー", "c値は0以上の整数で入力してください。")
                        return
                    values[param_key] = value

                self.config_manager.set_inspection_preset(
                    mode_key,
                    aql=values["aql"],
                    ltpd=values["ltpd"],
                    alpha=values["alpha"],
                    beta=values["beta"],
                    c_value=values["c_value"]
                )

            messagebox.showinfo("完了", "設定を保存しました。")
            self.dialog.destroy()

        except ValueError:
            messagebox.showerror("エラー", "入力値が正しくありません。数値を確認してください。")
        except Exception as e:
            sanitized_error = self.security_manager.sanitize_error_message(str(e))
            messagebox.showerror("エラー", f"設定の保存中にエラーが発生しました:\n{sanitized_error}")
    
    def _cancel(self):
        """キャンセルボタンの処理"""
        self.dialog.destroy()
