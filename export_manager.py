"""
エクスポート管理モジュール
結果のエクスポート機能を管理
"""

from tkinter import filedialog, messagebox
from datetime import datetime


class ExportManager:
    """エクスポート管理クラス"""
    
    def __init__(self, app):
        self.app = app
    
    def export_results(self):
        """結果のエクスポート"""
        if not hasattr(self.app.controller, 'last_db_data') or not self.app.controller.last_db_data: 
            messagebox.showinfo("エクスポート不可", "先に計算を実行してください。")
            return
            
        # 結果テキストの生成
        texts = self.app.controller.ui_manager.generate_result_texts(
            self.app.controller.last_db_data, 
            self.app.controller.last_stats_results, 
            self.app.controller.last_inputs
        )
        sample_size_disp = self.app.controller.ui_manager.format_int(self.app.controller.last_stats_results['sample_size'])
        
        content = f"""AI SQC Sampler - 計算結果
計算日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
{'=' * 50}

品番: {self.app.controller.last_inputs['product_number']}
ロットサイズ: {self.app.controller.ui_manager.format_int(self.app.controller.last_inputs['lot_size'])}個
不具合率: {self.app.controller.last_db_data['defect_rate']:.2f}%
検査水準: {self.app.controller.last_stats_results['level_text']}
サンプルサイズ: {sample_size_disp} 個
信頼度: {self.app.controller.last_inputs['confidence_level']:.1f}%
c値: {self.app.controller.last_inputs['c_value']}

{texts['review']}

{texts['best5']}
"""
        
        try:
            filepath = filedialog.asksaveasfilename(
                title="結果を名前を付けて保存",
                defaultextension=".txt",
                filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")],
                initialfile=f"検査結果_{self.app.controller.last_inputs['product_number']}.txt"
            )
            if not filepath: 
                return
            with open(filepath, 'w', encoding='utf-8') as f: 
                f.write(content)
            messagebox.showinfo("成功", f"結果を保存しました。\nパス: {filepath}")
        except Exception as e:
            messagebox.showerror("エクスポート失敗", f"ファイルの保存中にエラーが発生しました: {e}")

