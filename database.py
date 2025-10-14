"""
データベース操作モジュール
Microsoft Accessデータベースとの接続と操作を管理
"""

import pyodbc
import os
import threading
import time
from queue import Queue, Empty
from typing import Optional
from security_manager import SecurityManager
from error_handler import error_handler, ErrorCode
from memory_manager import memory_manager


class ConnectionPool:
    """データベース接続プールクラス"""
    
    def __init__(self, connection_string: str, max_connections: int = 5):
        self.connection_string = connection_string
        self.max_connections = max_connections
        self._pool = Queue(maxsize=max_connections)
        self._active_connections = 0
        self._lock = threading.Lock()
        self._last_health_check = 0
        self._health_check_interval = 300  # 5分間隔
    
    def get_connection(self) -> Optional[pyodbc.Connection]:
        """接続プールから接続を取得"""
        try:
            # ヘルスチェックの実行
            self._perform_health_check()
            
            # プールから接続を取得
            try:
                conn = self._pool.get_nowait()
                if self._is_connection_valid(conn):
                    return conn
                else:
                    conn.close()
            except Empty:
                pass
            
            # 新しい接続を作成
            with self._lock:
                if self._active_connections < self.max_connections:
                    conn = pyodbc.connect(self.connection_string)
                    self._active_connections += 1
                    return conn
            
            # プールが満杯の場合は待機
            conn = self._pool.get(timeout=30)
            if self._is_connection_valid(conn):
                return conn
            else:
                conn.close()
                return self.get_connection()
                
        except Exception as e:
            error_handler.handle_error(ErrorCode.DB_CONNECTION_FAILED, e)
            return None
    
    def return_connection(self, conn: pyodbc.Connection):
        """接続をプールに返却"""
        if conn and self._is_connection_valid(conn):
            try:
                self._pool.put_nowait(conn)
            except:
                conn.close()
                with self._lock:
                    self._active_connections -= 1
    
    def _is_connection_valid(self, conn: pyodbc.Connection) -> bool:
        """接続の有効性チェック"""
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT 1")
            cursor.fetchone()
            cursor.close()
            return True
        except:
            return False
    
    def _perform_health_check(self):
        """ヘルスチェックの実行"""
        current_time = time.time()
        if current_time - self._last_health_check < self._health_check_interval:
            return
        
        self._last_health_check = current_time
        
        # プール内の接続をチェック
        valid_connections = []
        while not self._pool.empty():
            try:
                conn = self._pool.get_nowait()
                if self._is_connection_valid(conn):
                    valid_connections.append(conn)
                else:
                    conn.close()
                    with self._lock:
                        self._active_connections -= 1
            except Empty:
                break
        
        # 有効な接続をプールに戻す
        for conn in valid_connections:
            try:
                self._pool.put_nowait(conn)
            except:
                conn.close()
                with self._lock:
                    self._active_connections -= 1
    
    def close_all(self):
        """すべての接続を閉じる"""
        while not self._pool.empty():
            try:
                conn = self._pool.get_nowait()
                conn.close()
            except Empty:
                break
        
        with self._lock:
            self._active_connections = 0


class DatabaseManager:
    """データベース管理クラス（接続プール対応）"""
    
    def __init__(self, config_manager=None):
        self.config_manager = config_manager
        self.security_manager = SecurityManager()
        self._connection_pool = None
        self._connection_string = None
    
    def _get_connection_string(self) -> Optional[str]:
        """接続文字列の取得"""
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
                error_handler.handle_error(
                    ErrorCode.DB_ACCESS_DENIED, 
                    Exception(f"DBアクセス拒否: {db_path} - {message}"),
                    {"db_path": db_path, "message": message}
                )
                return None
            
            conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                       f'DBQ={db_path};'
                       r'ReadOnly=False;'
                       r'Exclusive=False;')
            
            return conn_str
            
        except Exception as e:
            error_handler.handle_error(ErrorCode.DB_CONNECTION_FAILED, e)
            return None
    
    def _initialize_connection_pool(self):
        """接続プールの初期化"""
        if self._connection_pool is None:
            conn_str = self._get_connection_string()
            if conn_str:
                self._connection_string = conn_str
                self._connection_pool = ConnectionPool(conn_str, max_connections=5)
    
    def get_db_connection(self):
        """データベース接続の取得（接続プール対応）"""
        try:
            # 接続プールの初期化
            self._initialize_connection_pool()
            
            if self._connection_pool is None:
                return None
            
            # 接続プールから接続を取得
            conn = self._connection_pool.get_connection()
            if conn is None:
                error_handler.handle_error(
                    ErrorCode.DB_CONNECTION_FAILED, 
                    Exception("接続プールから接続を取得できませんでした")
                )
            
            return conn
            
        except pyodbc.Error as e:
            raw_message = str(e)
            if "Microsoft Access Driver" in raw_message:
                error_handler.handle_error(
                    ErrorCode.DB_CONNECTION_FAILED, 
                    Exception("Microsoft Access Driverが見つかりません。Microsoft Access Database Engine 2016 Redistributableをインストールしてください。")
                )
            else:
                error_handler.handle_error(ErrorCode.DB_CONNECTION_FAILED, e)
            return None
        except Exception as e:
            error_handler.handle_error(ErrorCode.SYSTEM_ERROR, e)
            return None
    
    def return_db_connection(self, conn):
        """データベース接続の返却"""
        if self._connection_pool and conn:
            self._connection_pool.return_connection(conn)
    
    def close_all_connections(self):
        """すべての接続を閉じる"""
        if self._connection_pool:
            self._connection_pool.close_all()
    
    def fetch_all_product_numbers(self, force_refresh=False):
        """全品番の取得（メモリ管理対応キャッシュ）"""
        cache_key = "product_numbers"
        
        if not force_refresh:
            # メモリマネージャーからキャッシュを取得
            cached_data = memory_manager.get_cached_data(cache_key)
            if cached_data is not None:
                return cached_data
        
        product_numbers = self._query_product_numbers_from_db()
        if product_numbers is None:
            return []
        
        # メモリマネージャーにキャッシュを保存（1時間有効）
        memory_manager.cache_data(cache_key, list(product_numbers), max_age=3600)
        
        return product_numbers
    
    def _query_product_numbers_from_db(self):
        """データベースから品番リストを取得（接続プール対応）"""
        conn = self.get_db_connection()
        if not conn:
            return None
        
        try:
            with conn.cursor() as cursor:
                sql = (
                    "SELECT DISTINCT [品番] "
                    "FROM t_不具合情報 "
                    "WHERE [品番] IS NOT NULL AND [品番] <> '' "
                    "ORDER BY [品番]"
                )
                rows = cursor.execute(sql).fetchall()
                return [row[0] for row in rows if row[0]]
        except pyodbc.Error as e:
            error_handler.handle_error(ErrorCode.DB_QUERY_FAILED, e)
            return None
        finally:
            # 接続をプールに返却
            self.return_db_connection(conn)
    
    def test_connection(self):
        """データベース接続のテスト（接続プール対応）"""
        conn = self.get_db_connection()
        if not conn:
            return False, "データベース接続に失敗しました"
        
        try:
            with conn.cursor() as cursor:
                cursor.execute("SELECT COUNT(*) FROM t_不具合情報")
                count = cursor.fetchone()[0]
                return True, f"接続成功: {count}件のレコードを確認"
        except pyodbc.Error as e:
            error_handler.handle_error(ErrorCode.DB_QUERY_FAILED, e)
            return False, f"テーブルアクセスエラー: {str(e)}"
        finally:
            # 接続をプールに返却
            self.return_db_connection(conn)
