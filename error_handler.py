"""
統一エラーハンドリングモジュール
アプリケーション全体のエラー処理を統一管理
"""

import logging
import traceback
from enum import Enum
from typing import Optional, Dict, Any
from tkinter import messagebox
from security_manager import SecurityManager


class ErrorCode(Enum):
    """エラーコード定義"""
    # データベース関連
    DB_CONNECTION_FAILED = "DB001"
    DB_QUERY_FAILED = "DB002"
    DB_ACCESS_DENIED = "DB003"
    DB_TIMEOUT = "DB004"
    
    # 計算関連
    CALCULATION_ERROR = "CALC001"
    CALCULATION_OVERFLOW = "CALC002"
    INVALID_INPUT = "CALC003"
    
    # ファイル関連
    FILE_ACCESS_DENIED = "FILE001"
    FILE_NOT_FOUND = "FILE002"
    FILE_CORRUPTED = "FILE003"
    
    # 設定関連
    CONFIG_INVALID = "CONF001"
    CONFIG_SAVE_FAILED = "CONF002"
    
    # システム関連
    SYSTEM_ERROR = "SYS001"
    MEMORY_ERROR = "SYS002"
    THREAD_ERROR = "SYS003"


class ErrorHandler:
    """統一エラーハンドリングクラス"""
    
    def __init__(self):
        self.security_manager = SecurityManager()
        self.logger = self._setup_logger()
        self._error_counts = {}
    
    def _setup_logger(self):
        """エラーログの設定"""
        logger = logging.getLogger('error_handler')
        logger.setLevel(logging.ERROR)
        
        if not logger.handlers:
            handler = logging.FileHandler('error.log', encoding='utf-8')
            formatter = logging.Formatter(
                '%(asctime)s - %(levelname)s - %(message)s'
            )
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        
        return logger
    
    def handle_error(self, 
                    error_code: ErrorCode, 
                    error: Exception, 
                    context: Optional[Dict[str, Any]] = None,
                    show_user_message: bool = True) -> bool:
        """
        エラーの統一処理
        
        Args:
            error_code: エラーコード
            error: 例外オブジェクト
            context: エラー発生時のコンテキスト情報
            show_user_message: ユーザーにメッセージを表示するか
            
        Returns:
            bool: エラーが処理されたかどうか
        """
        try:
            # エラーカウントの更新
            self._error_counts[error_code] = self._error_counts.get(error_code, 0) + 1
            
            # エラーログの記録
            self._log_error(error_code, error, context)
            
            # セキュリティイベントの記録
            self.security_manager.log_security_event(
                f"ERROR_{error_code.value}", 
                f"エラー発生: {str(error)}"
            )
            
            # ユーザーメッセージの表示
            if show_user_message:
                self._show_user_message(error_code, error, context)
            
            return True
            
        except Exception as e:
            # エラーハンドリング自体でエラーが発生した場合
            self.logger.critical(f"Error handler failed: {str(e)}")
            return False
    
    def _log_error(self, error_code: ErrorCode, error: Exception, context: Optional[Dict[str, Any]]):
        """エラーログの記録"""
        error_info = {
            'code': error_code.value,
            'type': type(error).__name__,
            'message': str(error),
            'context': context or {},
            'traceback': traceback.format_exc()
        }
        
        self.logger.error(f"Error {error_code.value}: {error_info}")
    
    def _show_user_message(self, error_code: ErrorCode, error: Exception, context: Optional[Dict[str, Any]]):
        """ユーザー向けエラーメッセージの表示"""
        sanitized_error = self.security_manager.sanitize_error_message(str(error))
        
        # エラーコード別のメッセージ
        messages = {
            ErrorCode.DB_CONNECTION_FAILED: "データベースに接続できませんでした。\n設定を確認してください。",
            ErrorCode.DB_QUERY_FAILED: "データベースのクエリ実行に失敗しました。",
            ErrorCode.DB_ACCESS_DENIED: "データベースファイルにアクセスできません。\nファイルの権限を確認してください。",
            ErrorCode.DB_TIMEOUT: "データベース接続がタイムアウトしました。\nしばらく待ってから再試行してください。",
            ErrorCode.CALCULATION_ERROR: "計算処理中にエラーが発生しました。",
            ErrorCode.CALCULATION_OVERFLOW: "計算結果が処理可能な範囲を超えました。",
            ErrorCode.INVALID_INPUT: "入力値が正しくありません。",
            ErrorCode.FILE_ACCESS_DENIED: "ファイルにアクセスできません。",
            ErrorCode.FILE_NOT_FOUND: "ファイルが見つかりません。",
            ErrorCode.FILE_CORRUPTED: "ファイルが破損している可能性があります。",
            ErrorCode.CONFIG_INVALID: "設定値が無効です。",
            ErrorCode.CONFIG_SAVE_FAILED: "設定の保存に失敗しました。",
            ErrorCode.SYSTEM_ERROR: "システムエラーが発生しました。",
            ErrorCode.MEMORY_ERROR: "メモリ不足が発生しました。",
            ErrorCode.THREAD_ERROR: "スレッド処理中にエラーが発生しました。"
        }
        
        title = f"エラー ({error_code.value})"
        message = messages.get(error_code, f"予期しないエラーが発生しました:\n{sanitized_error}")
        
        messagebox.showerror(title, message)
    
    def get_error_count(self, error_code: ErrorCode) -> int:
        """エラー発生回数の取得"""
        return self._error_counts.get(error_code, 0)
    
    def reset_error_counts(self):
        """エラーカウントのリセット"""
        self._error_counts.clear()
    
    def is_error_frequent(self, error_code: ErrorCode, threshold: int = 5) -> bool:
        """エラーが頻繁に発生しているかチェック"""
        return self._error_counts.get(error_code, 0) >= threshold


# グローバルエラーハンドラーインスタンス
error_handler = ErrorHandler()
