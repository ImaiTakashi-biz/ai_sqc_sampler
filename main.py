"""
AI SQC Sampler - メインコントローラー
統計的品質管理によるサンプリングサイズ計算アプリケーション
"""

import tkinter as tk
from tkinter import messagebox
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


class MainController:
    """メインアプリケーションコントローラー"""
    
    def __init__(self):
        # 基本コンポーネントの初期化
        self.config_manager = ConfigManager()
        self.db_manager = DatabaseManager(self.config_manager)
        self.app = App(self)
        
        # 各マネージャーの初期化
        self.calculation_engine = CalculationEngine(self.db_manager)
        self.ui_manager = UIManager(self.app)
        self.progress_manager = ProgressManager(self.app, self.db_manager, self.calculation_engine, self.ui_manager)
        self.product_list_manager = ProductListManager(self.app, self.db_manager)
        self.export_manager = ExportManager(self.app)
        
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
        """ユーザー入力の取得と検証"""
        inputs = {
            'product_number': self.app.sample_pn_entry.get().strip(),
            'lot_size_str': self.app.sample_qty_entry.get().strip(),
            'start_date': self.app.sample_start_date_entry.get().strip() or None,
            'end_date': self.app.sample_end_date_entry.get().strip() or None,
            'conf_str': self.app.sample_conf_entry.get().strip() or str(self.config_manager.get('default_confidence', 99.0)),
            'c_str': self.app.sample_c_entry.get().strip() or str(self.config_manager.get('default_c_value', 0))
        }
        
        # 入力値の検証
        validator = InputValidator()
        is_valid, errors, validated_data = validator.validate_all_inputs(
            inputs['product_number'],
            inputs['lot_size_str'],
            inputs['conf_str'],
            inputs['c_str'],
            inputs['start_date'],
            inputs['end_date']
        )
        
        if not is_valid:
            error_message = "以下の入力エラーがあります：\n" + "\n".join(f"• {error}" for error in errors)
            messagebox.showwarning("入力エラー", error_message)
            return None
        
        return validated_data

    def show_product_numbers_list(self):
        """品番リストの表示"""
        self.product_list_manager.show_product_numbers_list()

    def export_results(self):
        """結果のエクスポート"""
        self.export_manager.export_results()

    def open_config_dialog(self):
        """設定ダイアログの表示"""
        dialog = SettingsDialog(self.app, self.config_manager)
        dialog.show()
        
        # 設定変更後、データベースマネージャーを再初期化
        self.db_manager = DatabaseManager(self.config_manager)
        self.calculation_engine = CalculationEngine(self.db_manager)
        self.progress_manager = ProgressManager(self.app, self.db_manager, self.calculation_engine, self.ui_manager)
        self.product_list_manager = ProductListManager(self.app, self.db_manager)

    def show_help(self):
        """ヘルプの表示"""
        messagebox.showinfo("ヘルプ", "AI SQC Sampler - 統計的品質管理によるサンプリングサイズ計算アプリケーション")

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


def main():
    """メイン関数"""
    try:
        controller = MainController()
        controller.run()
    except Exception as e:
        messagebox.showerror("アプリケーションエラー", f"アプリケーションの起動中にエラーが発生しました:\n{str(e)}")


if __name__ == "__main__":
    main()