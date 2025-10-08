"""
データベース操作モジュール
Microsoft Accessデータベースとの接続と操作を管理
"""

import pyodbc
import os
import threading
from tkinter import messagebox
from security_manager import SecurityManager


class DatabaseManager:
    """データベース管理クラス"""
    
    def __init__(self, config_manager=None):
        self._connection_lock = threading.Lock()
        self._product_numbers_cache = None
        self._product_numbers_cache_lock = threading.Lock()
        self.config_manager = config_manager
        self.security_manager = SecurityManager()
    
    def get_db_connection(self):
        """データベース接続の取得"""
        try:
            # 設定からデータベースパスを取得
            if self.config_manager:
                db_path = self.config_manager.get_database_path()
            else:
                # フォールバック: デフォルトパス
                db_path = os.path.join(os.getcwd(), "不具合情報記録.accdb")
            
            # ファイルアクセス権限の検証
            is_valid, message = self.security_manager.validate_file_access(db_path)
            if not is_valid:
                self.security_manager.log_security_event("DB_ACCESS_DENIED", f"DBアクセス拒否: {db_path} - {message}")
                error_msg = f"データベースファイルにアクセスできません:\n{message}\n\n設定からデータベースファイルのパスを変更してください。"
                messagebox.showerror("データベースアクセスエラー", error_msg)
                return None
            
            conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                       f'DBQ={db_path};'
                       r'ReadOnly=False;'
                       r'Exclusive=False;')
            
            conn = pyodbc.connect(conn_str)
            return conn
            
        except pyodbc.Error as e:
            self.security_manager.log_security_event("DB_CONNECTION_ERROR", f"DB接続エラー: {str(e)}")
            if "Microsoft Access Driver" in str(e):
                messagebox.showerror("ドライバーエラー", 
                    "Microsoft Access Driverが見つかりません。\n"
                    "Microsoft Access Database Engine 2016 Redistributableをインストールしてください。")
            else:
                # エラーメッセージをサニタイズ
                sanitized_error = self.security_manager.sanitize_error_message(str(e))
                messagebox.showerror("データベースエラー", f"データベース接続に失敗しました:\n{sanitized_error}")
            return None
        except Exception as e:
            self.security_manager.log_security_event("UNEXPECTED_ERROR", f"予期しないエラー: {str(e)}")
            # エラーメッセージをサニタイズ
            sanitized_error = self.security_manager.sanitize_error_message(str(e))
            messagebox.showerror("エラー", f"予期しないエラーが発生しました:\n{sanitized_error}")
            return None
    
    def get_defect_data(self, product_number, start_date=None, end_date=None):
        """不具合データの取得"""
        conn = self.get_db_connection()
        if not conn:
            return None
        
        try:
            with conn.cursor() as cursor:
                has_start_date = start_date is not None and str(start_date).strip() and str(start_date).strip() != "" and str(start_date).strip() != "None"
                has_end_date = end_date is not None and str(end_date).strip() and str(end_date).strip() != "" and str(end_date).strip() != "None"
                
                if has_start_date and has_end_date:
                    sql = "SELECT * FROM t_不具合情報 WHERE [品番] = ? AND [日付] >= ? AND [日付] <= ?"
                    params = [product_number, str(start_date).strip(), str(end_date).strip()]
                elif has_start_date:
                    sql = "SELECT * FROM t_不具合情報 WHERE [品番] = ? AND [日付] >= ?"
                    params = [product_number, str(start_date).strip()]
                elif has_end_date:
                    sql = "SELECT * FROM t_不具合情報 WHERE [品番] = ? AND [日付] <= ?"
                    params = [product_number, str(end_date).strip()]
                else:
                    sql = "SELECT * FROM t_不具合情報 WHERE [品番] = ?"
                    params = [product_number]
                
                rows = cursor.execute(sql, params).fetchall()
                return rows
        except pyodbc.Error as e:
            messagebox.showerror("データベースエラー", f"不具合データの取得中にエラーが発生しました: {e}")
            return None
        finally:
            if conn:
                conn.close()
    
    def fetch_all_product_numbers(self, force_refresh=False):
        """全品番の取得（キャッシュ機能付き）"""
        if not force_refresh:
            with self._product_numbers_cache_lock:
                if self._product_numbers_cache is not None:
                    return list(self._product_numbers_cache)
        
        product_numbers = self._query_product_numbers_from_db()
        if product_numbers is None:
            return []
        
        with self._product_numbers_cache_lock:
            self._product_numbers_cache = list(product_numbers)
        return product_numbers
    
    def _query_product_numbers_from_db(self):
        """データベースから品番リストを取得"""
        conn = self.get_db_connection()
        if not conn:
            return None
        
        try:
            with conn.cursor() as cursor:
                sql = "SELECT [品番] FROM t_不具合情報 WHERE [品番] IS NOT NULL AND [品番] <> '' ORDER BY [品番]"
                rows = cursor.execute(sql).fetchall()
                raw_numbers = [row[0] for row in rows if row[0]]
                product_numbers = list(dict.fromkeys(raw_numbers))  # 重複削除
                return product_numbers
        except pyodbc.Error as e:
            messagebox.showerror("データベースエラー", f"品番リストの取得中にエラーが発生しました: {e}")
            return None
        finally:
            if conn:
                conn.close()
    
    def test_connection(self):
        """データベース接続のテスト"""
        conn = self.get_db_connection()
        if not conn:
            return False, "データベース接続に失敗しました"
        
        try:
            with conn.cursor() as cursor:
                cursor.execute("SELECT COUNT(*) FROM t_不具合情報")
                count = cursor.fetchone()[0]
                return True, f"接続成功: {count}件のレコードを確認"
        except pyodbc.Error as e:
            return False, f"テーブルアクセスエラー: {e}"
        finally:
            if conn:
                conn.close()