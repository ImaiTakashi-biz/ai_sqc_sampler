"""
AI SQC Sampler - メインコントローラー
統計的品質管理によるサンプリングサイズ計算アプリケーション
"""

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
from oc_curve_manager import OCCurveManager
from inspection_level_manager import InspectionLevelManager
from security_manager import SecurityManager


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
