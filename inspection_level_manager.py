"""
検査水準管理モジュール
ISO 2859-1標準に基づく通常/強化/緩和の切替ルール
"""

import tkinter as tk
from tkinter import ttk, messagebox
from enum import Enum


class InspectionLevel(Enum):
    """検査水準の定義"""
    NORMAL = "通常検査"
    TIGHTENED = "強化検査"
    REDUCED = "緩和検査"


class InspectionLevelManager:
    """検査水準管理クラス"""
    
    def __init__(self):
        # ISO 2859-1標準の切替ルール
        self.switching_rules = {
            "normal_to_tightened": {
                "condition": "連続5ロット中2ロットが不合格",
                "description": "品質が悪化傾向にあるため強化検査に移行"
            },
            "tightened_to_normal": {
                "condition": "連続5ロットが合格",
                "description": "品質が改善されたため標準検査に復帰"
            },
            "normal_to_reduced": {
                "condition": "連続10ロットが合格 かつ 生産者品質が良好",
                "description": "品質が安定しているため緩和検査に移行"
            },
            "reduced_to_normal": {
                "condition": "1ロットが不合格 または 品質が不安定",
                "description": "品質に問題が生じたため標準検査に復帰"
            }
        }
    
    def get_current_inspection_level(self, aql=None, recent_results=None):
        """現在の検査水準を決定"""
        if recent_results is None:
            recent_results = []
        
        # デフォルトは通常検査
        current_level = InspectionLevel.NORMAL
        
        # 最近の結果に基づく判定
        if len(recent_results) >= 5:
            # 連続5ロット中2ロットが不合格の場合
            recent_5 = recent_results[-5:]
            failed_count = sum(1 for result in recent_5 if not result.get('passed', True))
            
            if failed_count >= 2:
                current_level = InspectionLevel.TIGHTENED
            elif failed_count == 0 and len(recent_results) >= 10:
                # 連続10ロットが合格の場合
                recent_10 = recent_results[-10:]
                if all(result.get('passed', True) for result in recent_10):
                    current_level = InspectionLevel.REDUCED
        
        return current_level
    
    def get_sample_size_adjustment(self, base_sample_size, inspection_level):
        """検査水準に応じた抜取数の調整（非推奨：AQL/LTPD設計を使用）"""
        # 注意：この機能は非推奨です。AQL/LTPD設計による統計計算を使用してください。
        # この機能は後方互換性のために残されています。
        adjustments = {
            InspectionLevel.NORMAL: 1.0,      # 基準値
            InspectionLevel.TIGHTENED: 1.0,   # 調整なし（AQL/LTPD設計に委ねる）
            InspectionLevel.REDUCED: 1.0      # 調整なし（AQL/LTPD設計に委ねる）
        }
        
        return int(base_sample_size * adjustments.get(inspection_level, 1.0))
    
    def create_inspection_level_dialog(self, parent, current_level, recent_results, config_manager=None):
        """検査水準管理ダイアログの作成（運用管理用）"""
        dialog = tk.Toplevel(parent)
        dialog.title("検査水準管理")
        dialog.geometry("750x600")
        dialog.configure(bg="#f8f9fa")
        dialog.resizable(True, True)
        
        # 中央配置
        x = (parent.winfo_screenwidth() // 2) - 375
        y = (parent.winfo_screenheight() // 2) - 300
        dialog.geometry(f"750x600+{x}+{y}")
        
        # モーダル表示
        dialog.transient(parent)
        dialog.grab_set()
        
        # タイトル
        title_label = tk.Label(
            dialog, 
            text="📋 検査水準管理", 
            font=("Meiryo", 16, "bold"), 
            fg="#2c3e50", 
            bg="#f8f9fa"
        )
        title_label.pack(pady=(20, 10))
        
        # 説明
        explanation_label = tk.Label(
            dialog,
            text="検査区分（緩和・標準・強化）の切替ルールと運用管理を行います。\n抜取数は各検査区分のデフォルト値（AQL/LTPD/α/β/c値）により統計的に計算されます。",
            font=("Meiryo", 10),
            fg="#6c757d",
            bg="#f8f9fa",
            wraplength=700,
            justify='center'
        )
        explanation_label.pack(pady=(0, 15))
        
        # 現在の検査水準表示
        current_frame = tk.LabelFrame(
            dialog, 
            text="現在の検査水準", 
            font=("Meiryo", 12, "bold"), 
            fg="#2c3e50", 
            bg="#f8f9fa",
            padx=10,
            pady=10
        )
        current_frame.pack(fill='x', padx=20, pady=10)
        
        level_colors = {
            InspectionLevel.NORMAL: "#28a745",
            InspectionLevel.TIGHTENED: "#dc3545", 
            InspectionLevel.REDUCED: "#ffc107"
        }
        
        level_label = tk.Label(
            current_frame, 
            text=f"検査区分: {current_level.value}", 
            font=("Meiryo", 14, "bold"), 
            fg=level_colors.get(current_level, "#2c3e50"), 
            bg="#f8f9fa"
        )
        level_label.pack()
        
        # 検査区分のデフォルト値表示
        if config_manager:
            try:
                mode_key = self._get_mode_key_from_level(current_level)
                details = config_manager.get_inspection_mode_details(mode_key)
                values_text = f"AQL {details.get('aql', 0):.2f}% | LTPD {details.get('ltpd', 0):.2f}% | α {details.get('alpha', 0):.1f}% | β {details.get('beta', 0):.1f}% | c値 {details.get('c_value', 0)}"
                values_label = tk.Label(
                    current_frame,
                    text=values_text,
                    font=("Meiryo", 10),
                    fg="#495057",
                    bg="#f8f9fa"
                )
                values_label.pack(pady=(5, 0))
            except Exception:
                pass
        
        # 切替ルールの表示
        rules_frame = tk.LabelFrame(
            dialog, 
            text="切替ルール", 
            font=("Meiryo", 12, "bold"), 
            fg="#2c3e50", 
            bg="#f8f9fa",
            padx=10,
            pady=10
        )
        rules_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # スクロール可能なテキストエリア
        text_frame = tk.Frame(rules_frame, bg="#f8f9fa")
        text_frame.pack(fill='both', expand=True)
        
        scrollbar = tk.Scrollbar(text_frame)
        text_widget = tk.Text(
            text_frame, 
            font=("Meiryo", 10), 
            bg="#ffffff", 
            fg="#2c3e50",
            wrap=tk.WORD,
            yscrollcommand=scrollbar.set
        )
        scrollbar.config(command=text_widget.yview)
        
        scrollbar.pack(side='right', fill='y')
        text_widget.pack(side='left', fill='both', expand=True)
        
        # 切替ルールの説明
        rules_text = self.generate_rules_text(current_level, recent_results)
        text_widget.insert('1.0', rules_text)
        text_widget.config(state='disabled')
        
        # 最近の結果表示
        if recent_results:
            results_frame = tk.LabelFrame(
                dialog, 
                text="最近の検査結果（最新5ロット）", 
                font=("Meiryo", 12, "bold"), 
                fg="#2c3e50", 
                bg="#f8f9fa",
                padx=10,
                pady=10
            )
            results_frame.pack(fill='x', padx=20, pady=10)
            
            results_text = self.format_recent_results(recent_results[-5:])
            tk.Label(
                results_frame, 
                text=results_text, 
                font=("Meiryo", 10), 
                fg="#495057", 
                bg="#f8f9fa",
                justify='left'
            ).pack(anchor='w')
        
        # 閉じるボタン
        close_button = tk.Button(
            dialog, 
            text="閉じる", 
            command=dialog.destroy, 
            font=("Meiryo", 10, "bold"), 
            bg="#6c757d", 
            fg="#ffffff", 
            relief="flat", 
            padx=20, 
            pady=5
        )
        close_button.pack(pady=20)
    
    def generate_rules_text(self, current_level, recent_results):
        """切替ルールの説明テキストを生成（運用管理用）"""
        text = f"【現在の検査区分: {current_level.value}】\n\n"
        text += "※ 抜取数は各検査区分のデフォルト値（AQL/LTPD/α/β/c値）により統計的に計算されます\n"
        text += "※ 検査区分は運用管理・品質トレンド監視に使用されます\n"
        text += "※ デフォルト値は設定画面で変更可能です\n\n"
        
        if current_level == InspectionLevel.NORMAL:
            text += "標準検査が適用されています。\n\n"
            text += "【強化検査への移行条件】\n"
            text += f"• 条件: {self.switching_rules['normal_to_tightened']['condition']}\n"
            text += f"• 理由: {self.switching_rules['normal_to_tightened']['description']}\n\n"
            text += "【緩和検査への移行条件】\n"
            text += f"• 条件: {self.switching_rules['normal_to_reduced']['condition']}\n"
            text += f"• 理由: {self.switching_rules['normal_to_reduced']['description']}\n\n"
            
        elif current_level == InspectionLevel.TIGHTENED:
            text += "強化検査が適用されています。\n\n"
            text += "【標準検査への復帰条件】\n"
            text += f"• 条件: {self.switching_rules['tightened_to_normal']['condition']}\n"
            text += f"• 理由: {self.switching_rules['tightened_to_normal']['description']}\n\n"
            
        elif current_level == InspectionLevel.REDUCED:
            text += "緩和検査が適用されています。\n\n"
            text += "【標準検査への復帰条件】\n"
            text += f"• 条件: {self.switching_rules['reduced_to_normal']['condition']}\n"
            text += f"• 理由: {self.switching_rules['reduced_to_normal']['description']}\n\n"
        
        text += "【ISO 2859-1標準について】\n"
        text += "• 国際標準化機構（ISO）が定める抜取検査の国際規格\n"
        text += "• 統計的品質管理の原則に基づく検査区分の切替ルール\n"
        text += "• 品質の変動に応じて検査の厳しさを動的に調整\n"
        text += "• 注意：抜取数は各検査区分のデフォルト値により統計的に計算されます\n"
        text += "• 生産者と消費者の両方のリスクを適切に管理\n"
        text += "• 各検査区分のデフォルト値は設定画面で変更可能です"
        
        return text    
    def _get_mode_key_from_level(self, level):
        """検査水準から検査区分キーを取得"""
        mapping = {
            InspectionLevel.NORMAL: "standard",
            InspectionLevel.TIGHTENED: "tightened", 
            InspectionLevel.REDUCED: "reduced"
        }
        return mapping.get(level, "standard")
    
    def format_recent_results(self, results):
        """最近の結果をフォーマット"""
        if not results:
            return "検査結果データがありません。"
        
        text = ""
        for i, result in enumerate(results, 1):
            status = "合格" if result.get('passed', True) else "不合格"
            date = result.get('date', 'N/A')
            text += f"{i}. {date}: {status}\n"
        
        return text
