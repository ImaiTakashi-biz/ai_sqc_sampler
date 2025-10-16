"""
エクスポート管理モジュール
結果のテキストエクスポート機能を管理
"""

from tkinter import filedialog, messagebox
from datetime import datetime
from security_manager import SecurityManager


class ExportManager:
    """エクスポート管理クラス"""
    
    def __init__(self, app):
        self.app = app
        self.security_manager = SecurityManager()
    
    def export_results(self):
        """結果のエクスポート（従来のテキスト出力）"""
        if not hasattr(self.app.controller, 'last_db_data') or not self.app.controller.last_db_data: 
            messagebox.showinfo("エクスポート不可", "先に計算を実行してください。")
            return
        
        # 不具合データがない場合の特別処理
        if self.app.controller.last_stats_results.get('no_defect_data', False):
            self._export_no_defect_data_results()
            return
            
        # 結果テキストの生成
        texts = self.app.controller.ui_manager.generate_result_texts(
            self.app.controller.last_db_data, 
            self.app.controller.last_stats_results, 
            self.app.controller.last_inputs
        )
        sample_size_disp = self.app.controller.ui_manager.format_int(self.app.controller.last_stats_results['sample_size'])
        
        content = f"""AI SQC Sampler - 計算結果（AQL/LTPD設計）
計算日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
{'=' * 60}

品番: {self.app.controller.last_inputs['product_number']}
ロットサイズ: {self.app.controller.ui_manager.format_int(self.app.controller.last_inputs['lot_size'])}個
不具合率: {self.app.controller.last_db_data['defect_rate']:.2f}%
検査水準: {self.app.controller.last_stats_results['level_text']}
サンプルサイズ: {sample_size_disp} 個

【AQL/LTPD設計パラメータ】
AQL（合格品質水準）: {self.app.controller.last_stats_results.get('aql', self.app.controller.last_inputs.get('aql', 0.25)):.3f}%
LTPD（不合格品質水準）: {self.app.controller.last_stats_results.get('ltpd', self.app.controller.last_inputs.get('ltpd', 1.0)):.3f}%
α（生産者危険）: {self.app.controller.last_inputs.get('alpha', 5.0):.1f}%
β（消費者危険）: {self.app.controller.last_inputs.get('beta', 10.0):.1f}%
c値（許容不良数）: {self.app.controller.last_inputs.get('c_value', 0)}

{texts['review']}

{texts['best5']}

{self._get_adjustment_info()}

{self._get_guidance_info()}
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
            sanitized_error = self.security_manager.sanitize_error_message(str(e))
            messagebox.showerror("エクスポート失敗", f"ファイルの保存中にエラーが発生しました: {sanitized_error}")
    
    
    
    def _get_adjustment_info(self):
        """調整情報の取得"""
        if 'adjustment_info' in self.app.controller.last_stats_results and self.app.controller.last_stats_results['adjustment_info']:
            return self.app.controller.last_stats_results['adjustment_info']
        return ""
    
    def _get_guidance_info(self):
        """ガイダンス情報の取得"""
        guidance_parts = []
        
        if 'guidance_message' in self.app.controller.last_stats_results and self.app.controller.last_stats_results['guidance_message']:
            guidance_parts.append(self.app.controller.last_stats_results['guidance_message'])
        
        if 'warning_message' in self.app.controller.last_stats_results and self.app.controller.last_stats_results['warning_message']:
            guidance_parts.append(f"【警告】\n{self.app.controller.last_stats_results['warning_message']}")
        
        return "\n\n".join(guidance_parts) if guidance_parts else ""
    
    def _export_no_defect_data_results(self):
        """不具合データがない場合の結果エクスポート"""
        lot_size = self.app.controller.last_inputs.get('lot_size', 1000)
        sample_size_disp = self.app.controller.ui_manager.format_int(lot_size)
        
        content = f"""AI SQC Sampler - 計算結果（全数検査推奨）
計算日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
{'=' * 60}

品番: {self.app.controller.last_inputs['product_number']}
ロットサイズ: {self.app.controller.ui_manager.format_int(lot_size)}個
不具合率: 0.00%
検査水準: 全数検査推奨
サンプルサイズ: {sample_size_disp} 個

【AQL/LTPD設計パラメータ】
AQL（合格品質水準）: {self.app.controller.last_inputs.get('aql', 0.25):.3f}%
LTPD（不合格品質水準）: {self.app.controller.last_inputs.get('ltpd', 1.0):.3f}%
α（生産者危険）: {self.app.controller.last_inputs.get('alpha', 5.0):.1f}%
β（消費者危険）: {self.app.controller.last_inputs.get('beta', 10.0):.1f}%
c値（許容不良数）: {self.app.controller.last_inputs.get('c_value', 0)}

【検査時の注意喚起】
該当期間に不具合データがありません。

【データベース実績に基づく推奨事項】

実績不良率: 0.000%

元の設定:
• AQL: {self.app.controller.last_inputs.get('aql', 0.25):.3f}%
• LTPD: {self.app.controller.last_inputs.get('ltpd', 1.0):.3f}%

推奨理由: 不具合データ（実績）がありません。統計的抜取検査の根拠となる実績データが不足しているため、全数検査を推奨します。
効果: 品質保証の確実性、不良品流出の完全防止

【ガイダンス】
実績データが不足しているため、統計的抜取検査ではなく全数検査を実施してください。
"""
        
        try:
            filepath = filedialog.asksaveasfilename(
                title="結果を名前を付けて保存",
                defaultextension=".txt",
                filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")],
                initialfile=f"検査結果_{self.app.controller.last_inputs['product_number']}_全数検査推奨.txt"
            )
            if not filepath: 
                return
            with open(filepath, 'w', encoding='utf-8') as f: 
                f.write(content)
            messagebox.showinfo("成功", f"結果を保存しました。\nパス: {filepath}")
        except Exception as e:
            sanitized_error = self.security_manager.sanitize_error_message(str(e))
            messagebox.showerror("エクスポート失敗", f"ファイルの保存中にエラーが発生しました: {sanitized_error}")