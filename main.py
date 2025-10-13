"""
AI SQC Sampler - メインコントローラー
統計的品質管理によるサンプリングサイズ計算アプリケーション
"""

from tkinter import messagebox, Toplevel, scrolledtext
import tkinter as tk
import os
import platform
import subprocess
import webbrowser
from pathlib import Path
from gui import App
from database import DatabaseManager
from validation import InputValidator
from config_manager import ConfigManager
from settings_dialog import SettingsDialog
from calculation_engine import CalculationEngine
from ui_manager import UIManager
from progress_manager import ProgressManager
from product_list_manager import ProductListManager
from export_manager import ExportManager
from oc_curve_manager import OCCurveManager
from inspection_level_manager import InspectionLevelManager
from security_manager import SecurityManager


class MainController:
    """メインアプリケーションコントローラー"""
    
    def __init__(self):
        # 基本コンポーネントの初期化
        self.config_manager = ConfigManager()
        self.security_manager = SecurityManager()
        self.db_manager = DatabaseManager(self.config_manager)
        self.app = App(self)
        
        # 各マネージャーの初期化
        self.calculation_engine = CalculationEngine(self.db_manager)
        self.ui_manager = UIManager(self.app)
        self.progress_manager = ProgressManager(self.app, self.db_manager, self.calculation_engine, self.ui_manager)
        self.product_list_manager = ProductListManager(self.app, self.db_manager)
        self.export_manager = ExportManager(self.app)
        self.oc_curve_manager = OCCurveManager()
        self.inspection_level_manager = InspectionLevelManager()
        
        # 結果データの保存用
        self.last_db_data = None
        self.last_stats_results = None
        self.last_inputs = None

    def run(self):
        """アプリケーションの実行"""
        self.app.mainloop()

    def start_calculation_thread(self):
        """計算処理を別スレッドで開始"""
        inputs = self._get_user_inputs()
        if not inputs:
            return
        
        self.progress_manager.start_calculation_thread(inputs)

    def _get_user_inputs(self):
        """ユーザー入力の取得と検証（AQL/LTPD設計対応）"""
        if hasattr(self.app, 'reset_input_highlights'):
            self.app.reset_input_highlights()
        inputs = {
            'product_number': self.app.sample_pn_entry.get().strip(),
            'lot_size_str': self.app.sample_qty_entry.get().strip(),
            'start_date': self.app.sample_start_date_entry.get().strip() or None,
            'end_date': self.app.sample_end_date_entry.get().strip() or None,
            'aql_str': self.app.sample_aql_entry.get().strip() or "0.25",
            'ltpd_str': self.app.sample_ltpd_entry.get().strip() or "1.0",
            'alpha_str': self.app.sample_alpha_entry.get().strip() or "5.0",
            'beta_str': self.app.sample_beta_entry.get().strip() or "10.0",
            'c_str': self.app.sample_c_entry.get().strip() or "0"
        }

        mode_label = None
        mode_key = getattr(self.app, "current_inspection_mode_key", None)
        if hasattr(self.app, "inspection_mode_var"):
            try:
                mode_label = self.app.inspection_mode_var.get()
            except tk.TclError:
                mode_label = None

        if mode_label:
            inputs['inspection_mode_label'] = mode_label
        if mode_key:
            inputs['inspection_mode_key'] = mode_key
            if hasattr(self.config_manager, "get_inspection_mode_details"):
                try:
                    inputs['inspection_mode_details'] = self.config_manager.get_inspection_mode_details(mode_key)
                except Exception:
                    pass
        
        # 入力値の検証
        validator = InputValidator()
        is_valid, errors, validated_data = validator.validate_aql_ltpd_inputs(
            inputs['product_number'],
            inputs['lot_size_str'],
            inputs['aql_str'],
            inputs['ltpd_str'],
            inputs['alpha_str'],
            inputs['beta_str'],
            inputs['c_str'],
            inputs['start_date'],
            inputs['end_date']
        )
        
        if not is_valid:
            error_message = "以下の入力エラーがあります：\n" + "\n".join(f"• {error}" for error in errors)
            self._highlight_invalid_inputs(errors)
            messagebox.showwarning("入力エラー", error_message)
            return None

        if mode_label:
            validated_data['inspection_mode_label'] = mode_label
        if mode_key:
            validated_data['inspection_mode_key'] = mode_key
        if inputs.get('inspection_mode_details'):
            validated_data['inspection_mode_details'] = inputs['inspection_mode_details']

        return validated_data
    
    def _highlight_invalid_inputs(self, errors):
        """入力エラーに応じて対象フィールドを強調表示"""
        if not hasattr(self.app, 'mark_entry_error'):
            return
        
        keyword_map = [
            (("品番",), getattr(self.app, "sample_pn_entry", None)),
            (("数量", "ロット"), getattr(self.app, "sample_qty_entry", None)),
            (("AQL",), getattr(self.app, "sample_aql_entry", None)),
            (("LTPD",), getattr(self.app, "sample_ltpd_entry", None)),
            (("α", "生産者"), getattr(self.app, "sample_alpha_entry", None)),
            (("β", "消費者"), getattr(self.app, "sample_beta_entry", None)),
            (("c値", "許容不良"), getattr(self.app, "sample_c_entry", None)),
            (("開始日", "開始日付", "開始日時"), getattr(self.app, "sample_start_date_entry", None)),
            (("終了日", "終了日付", "終了日時"), getattr(self.app, "sample_end_date_entry", None)),
            (("日付", "期間"), getattr(self.app, "sample_start_date_entry", None))
        ]
        
        for error in errors:
            for keywords, widget in keyword_map:
                if widget is None:
                    continue
                if any(keyword in error for keyword in keywords):
                    self.app.mark_entry_error(widget)
                    break

    def show_product_numbers_list(self):
        """品番リストの表示"""
        self.product_list_manager.show_product_numbers_list()

    def export_results(self):
        """結果のエクスポート"""
        self.export_manager.export_results()

    def on_inspection_mode_change(self, mode_key):
        """検査区分変更時の処理"""
        preset = self.config_manager.apply_inspection_mode(mode_key)
        mode_label = self.config_manager.get_inspection_mode_label(mode_key)
        if hasattr(self.app, "apply_inspection_mode_preset"):
            self.app.apply_inspection_mode_preset(preset, mode_label)

    def open_config_dialog(self):
        """設定ダイアログの表示"""
        dialog = SettingsDialog(self.app, self.config_manager)
        dialog.show()
        
        # 設定変更後、データベースマネージャーを再初期化
        self.db_manager = DatabaseManager(self.config_manager)
        self.calculation_engine = CalculationEngine(self.db_manager)
        self.progress_manager = ProgressManager(self.app, self.db_manager, self.calculation_engine, self.ui_manager)
        self.product_list_manager = ProductListManager(self.app, self.db_manager)

        # 検査区分の反映と入力欄の更新
        if hasattr(self.config_manager, "get_inspection_mode"):
            current_mode_key = self.config_manager.get_inspection_mode()
            if hasattr(self.app, "refresh_inspection_mode_choices"):
                choices = self.config_manager.get_inspection_mode_choices()
                label_to_key = {label: key for key, label in choices.items()}
                self.app.refresh_inspection_mode_choices(label_to_key, current_mode_key)

            if hasattr(self.app, "apply_inspection_mode_preset"):
                defaults_source = getattr(self.config_manager, "DEFAULT_CONFIG", {})
                preset_values = {
                    "aql": self.config_manager.get("default_aql", defaults_source.get("default_aql", 0.25)),
                    "ltpd": self.config_manager.get("default_ltpd", defaults_source.get("default_ltpd", 1.0)),
                    "alpha": self.config_manager.get("default_alpha", defaults_source.get("default_alpha", 5.0)),
                    "beta": self.config_manager.get("default_beta", defaults_source.get("default_beta", 10.0)),
                    "c_value": self.config_manager.get("default_c_value", defaults_source.get("default_c_value", 0)),
                    "description": self.config_manager.get_inspection_mode_details(current_mode_key).get("description", "")
                }
                mode_label = self.config_manager.get_inspection_mode_label(current_mode_key)
                self.app.apply_inspection_mode_preset(preset_values, mode_label)

    def show_help(self):
        """ヘルプの表示（アプリケーション内でREADME内容を表示）"""
        try:
            # READMEファイルのパスを取得
            readme_path = Path(__file__).resolve().parent / "README.md"
            
            if not readme_path.exists():
                messagebox.showerror("エラー", f"READMEファイルが見つかりません。\nファイル: {readme_path.name}")
                return
            
            # READMEファイルの内容を読み込み
            with readme_path.open('r', encoding='utf-8') as f:
                readme_content = f.read()
            
            # ヘルプウィンドウを作成
            self._create_help_window(readme_content, readme_path)
                
        except Exception as e:
            # エラーが発生した場合は代替手段を提供
            sanitized = self.security_manager.sanitize_error_message(str(e)) if hasattr(self, 'security_manager') else str(e)
            error_msg = (
                "READMEファイルを読み込めませんでした。\n\n"
                f"エラー: {sanitized}\n\n"
                "代替手段:\n"
                "1. ファイルエクスプローラーでREADME.mdファイルを手動で開いてください\n"
                "2. テキストエディタでREADME.mdファイルを開いてください"
            )
            messagebox.showerror("ヘルプファイルを読み込めません", error_msg)
    
    def _create_help_window(self, content, readme_path):
        """ヘルプウィンドウの作成"""
        # ヘルプウィンドウを作成
        help_window = Toplevel(self.app)
        help_window.title("AI SQC Sampler - ヘルプ")
        help_window.geometry("1000x750")
        help_window.minsize(720, 480)
        help_window.configure(bg="#f0f0f0")
        help_window.grid_columnconfigure(0, weight=1)
        help_window.grid_rowconfigure(1, weight=1)
        
        # ウィンドウを中央に配置
        help_window.transient(self.app)
        help_window.grab_set()
        
        # 上部フレーム（検索・ナビゲーション）
        top_frame = tk.Frame(help_window, bg="#f0f0f0")
        top_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        top_frame.grid_columnconfigure(0, weight=1)
        
        # 検索機能
        search_frame = tk.Frame(top_frame, bg="#f0f0f0")
        search_frame.pack(side="left", fill="x", expand=True)
        
        tk.Label(search_frame, text="検索:", font=("Meiryo", 9), bg="#f0f0f0").pack(side="left", padx=(0, 5))
        search_entry = tk.Entry(search_frame, font=("Meiryo", 9), width=20)
        search_entry.pack(side="left", padx=(0, 5))
        
        def search_text():
            search_term = search_entry.get().lower()
            if search_term:
                # テキストエリアを一時的に編集可能にする
                text_area.config(state="normal")
                # 既存のハイライトをクリア
                text_area.tag_remove("search", "1.0", "end")
                # 検索実行
                start = "1.0"
                while True:
                    pos = text_area.search(search_term, start, "end", nocase=True)
                    if not pos:
                        break
                    end = f"{pos}+{len(search_term)}c"
                    text_area.tag_add("search", pos, end)
                    start = end
                # ハイライトスタイルを設定
                text_area.tag_config("search", background="yellow", foreground="black")
                # 最初の検索結果にスクロール
                if text_area.tag_ranges("search"):
                    text_area.see("search.first")
                text_area.config(state="disabled")
        
        search_button = tk.Button(
            search_frame,
            text="検索",
            command=search_text,
            font=("Meiryo", 8),
            bg="#3498db",
            fg="white",
            relief="flat",
            padx=10,
            pady=2
        )
        search_button.pack(side="left", padx=(0, 10))
        
        # クリアボタン
        def clear_search():
            text_area.config(state="normal")
            text_area.tag_remove("search", "1.0", "end")
            text_area.config(state="disabled")
            search_entry.delete(0, "end")
        
        clear_button = tk.Button(
            search_frame,
            text="クリア",
            command=clear_search,
            font=("Meiryo", 8),
            bg="#95a5a6",
            fg="white",
            relief="flat",
            padx=10,
            pady=2
        )
        clear_button.pack(side="left")
        
        # 目次ボタン
        def show_toc():
            toc_window = Toplevel(help_window)
            toc_window.title("目次")
            toc_window.geometry("300x400")
            toc_window.configure(bg="#f0f0f0")
            
            toc_text = scrolledtext.ScrolledText(
                toc_window,
                wrap="word",
                font=("Meiryo", 9),
                bg="#ffffff",
                fg="#333333"
            )
            toc_text.pack(fill="both", expand=True, padx=10, pady=10)
            
            # 目次を生成
            toc_content = self._generate_toc(content)
            toc_text.insert("1.0", toc_content)
            toc_text.config(state="disabled")
            
            def jump_to_section(section):
                # セクションにジャンプ
                text_area.config(state="normal")
                pos = text_area.search(section, "1.0", "end")
                if pos:
                    text_area.see(pos)
                text_area.config(state="disabled")
                toc_window.destroy()
            
            # 目次項目をクリック可能にする
            toc_text.config(state="normal")
            for line in toc_content.split('\n'):
                if line.strip() and line.startswith('#'):
                    # セクション名を抽出
                    section_name = line.replace('#', '').strip()
                    # クリックイベントを追加（簡易版）
                    pass
            toc_text.config(state="disabled")
        
        toc_button = tk.Button(
            top_frame,
            text="📋 目次",
            command=show_toc,
            font=("Meiryo", 9),
            bg="#2ecc71",
            fg="white",
            relief="flat",
            padx=15,
            pady=5
        )
        toc_button.pack(side="right")
        
        # メインテキストエリア
        text_area = scrolledtext.ScrolledText(
            help_window,
            wrap="word",
            font=("Meiryo", 10),
            bg="#ffffff",
            fg="#333333",
            padx=15,
            pady=15,
            relief="flat",
            borderwidth=0
        )
        text_area.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        # 内容を挿入
        text_area.insert("1.0", content)
        text_area.config(state="disabled")  # 読み取り専用にする
        
        # 下部ボタンフレーム
        button_frame = tk.Frame(help_window, bg="#f0f0f0")
        button_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
        
        # 外部で開くボタン
        def open_external():
            try:
                resolved_path = readme_path.resolve()
                system_name = platform.system()
                if system_name == 'Windows':
                    os.startfile(str(resolved_path))
                elif system_name == 'Darwin':
                    subprocess.run(['open', str(resolved_path)], check=True)
                elif system_name == 'Linux':
                    subprocess.run(['xdg-open', str(resolved_path)], check=True)
                else:
                    webbrowser.open(resolved_path.as_uri())
            except Exception as e:
                sanitized = self.security_manager.sanitize_error_message(str(e)) if hasattr(self, 'security_manager') else str(e)
                messagebox.showerror("エラー", f"外部アプリケーションで開けませんでした:\n{sanitized}")
        
        external_button = tk.Button(
            button_frame,
            text="📄 外部で開く",
            command=open_external,
            font=("Meiryo", 9),
            bg="#f39c12",
            fg="white",
            relief="flat",
            padx=15,
            pady=5
        )
        external_button.pack(side="left")
        
        # 閉じるボタン
        close_button = tk.Button(
            button_frame,
            text="閉じる",
            command=help_window.destroy,
            font=("Meiryo", 10, "bold"),
            bg="#e74c3c",
            fg="white",
            relief="flat",
            padx=20,
            pady=5,
            cursor="hand2"
        )
        close_button.pack(side="right")
        
        # ウィンドウのフォーカスを設定
        help_window.focus_set()
        search_entry.focus_set()
    
    def _generate_toc(self, content):
        """目次を生成"""
        toc_lines = []
        for line in content.split('\n'):
            if line.startswith('#'):
                level = len(line) - len(line.lstrip('#'))
                title = line.replace('#', '').strip()
                indent = '  ' * (level - 1)
                toc_lines.append(f"{indent}• {title}")
        return '\n'.join(toc_lines)

    def show_about(self):
        """アプリケーション情報の表示"""
        messagebox.showinfo("アプリケーション情報", 
            "AI SQC Sampler v1.0\n\n"
            "統計的品質管理によるサンプリングサイズ計算アプリケーション\n"
            "Microsoft Accessデータベース対応")

    def test_database_connection(self):
        """データベース接続のテスト"""
        success, message = self.db_manager.test_connection()
        if success:
            messagebox.showinfo("データベース接続テスト", f"✅ {message}")
        else:
            messagebox.showerror("データベース接続テスト", f"❌ {message}")
    
    def show_oc_curve(self):
        """OCカーブの表示"""
        if not hasattr(self, 'last_stats_results') or not self.last_stats_results:
            messagebox.showinfo("情報", "先に計算を実行してください。")
            return
        
        if 'oc_curve' not in self.last_stats_results:
            messagebox.showinfo("情報", "OCカーブデータがありません。")
            return
        
        # OCカーブダイアログの表示
        self.oc_curve_manager.create_oc_curve_dialog(
            self.app,
            self.last_stats_results['oc_curve'],
            self.last_inputs.get('aql', 0.25),
            self.last_inputs.get('ltpd', 1.0),
            self.last_inputs.get('alpha', 5.0),
            self.last_inputs.get('beta', 10.0),
            self.last_stats_results['sample_size'],
            self.last_inputs.get('c_value', 0),
            self.last_inputs.get('lot_size', 1000)
        )
    
    def show_inspection_level(self):
        """検査水準管理の表示"""
        # 現在の検査水準を決定
        current_level = self.inspection_level_manager.get_current_inspection_level(
            self.last_inputs.get('aql', 0.25) if hasattr(self, 'last_inputs') and self.last_inputs else 0.25
        )
        
        # 最近の結果（サンプルデータ）
        recent_results = [
            {'date': '2024-01-01', 'passed': True},
            {'date': '2024-01-02', 'passed': True},
            {'date': '2024-01-03', 'passed': False},
            {'date': '2024-01-04', 'passed': True},
            {'date': '2024-01-05', 'passed': True}
        ]
        
        # 検査水準管理ダイアログの表示
        self.inspection_level_manager.create_inspection_level_dialog(
            self.app,
            current_level,
            recent_results,
            self.last_inputs.get('aql', 0.25) if hasattr(self, 'last_inputs') and self.last_inputs else 0.25
        )


def main():
    """メイン関数"""
    try:
        controller = MainController()
        controller.run()
    except Exception as e:
        security_manager = SecurityManager()
        sanitized_error = security_manager.sanitize_error_message(str(e))
        messagebox.showerror("アプリケーションエラー", f"アプリケーションの起動中にエラーが発生しました:\n{sanitized_error}")


if __name__ == "__main__":
    main()
