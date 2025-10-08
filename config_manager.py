"""
設定管理モジュール
アプリケーションの設定を管理
"""

import os
import json
from tkinter import filedialog, messagebox
from security_manager import SecurityManager


class ConfigManager:
    """設定管理クラス"""
    
    CONFIG_FILE = "app_config.json"
    
    # デフォルト設定
    DEFAULT_CONFIG = {
        "database_path": "不具合情報記録.accdb",
        "window_geometry": "1000x700",
        "last_product_number": "",
        "default_confidence": 99.0,
        "default_c_value": 0
    }
    
    def __init__(self):
        self.security_manager = SecurityManager()
        self.config = self._load_config()
    
    def _load_config(self):
        """設定ファイルの読み込み"""
        try:
            if os.path.exists(self.CONFIG_FILE):
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # デフォルト設定とマージ（新しい設定項目の追加に対応）
                    merged_config = self.DEFAULT_CONFIG.copy()
                    merged_config.update(config)
                    return merged_config
            else:
                return self.DEFAULT_CONFIG.copy()
        except Exception as e:
            print(f"設定ファイルの読み込みエラー: {e}")
            return self.DEFAULT_CONFIG.copy()
    
    def save_config(self):
        """設定ファイルの保存"""
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"設定ファイルの保存エラー: {e}")
            return False
    
    def get_database_path(self):
        """データベースパスの取得"""
        db_path = self.config.get("database_path", self.DEFAULT_CONFIG["database_path"])
        
        # 暗号化されたパスの場合は復号化
        if db_path.startswith("encrypted:"):
            try:
                encrypted_data = db_path[10:]  # "encrypted:" を除去
                db_path = self.security_manager.decrypt_sensitive_data(encrypted_data)
            except Exception as e:
                self.security_manager.log_security_event("DECRYPT_ERROR", f"パス復号化エラー: {e}")
                return self.DEFAULT_CONFIG["database_path"]
        
        # パスのサニタイズ
        sanitized_path = self.security_manager.sanitize_path(db_path)
        if not sanitized_path:
            self.security_manager.log_security_event("INVALID_PATH", f"無効なパス: {db_path}")
            return self.DEFAULT_CONFIG["database_path"]
        
        # 相対パスの場合は現在のディレクトリからの絶対パスに変換
        if not os.path.isabs(sanitized_path):
            sanitized_path = os.path.join(os.getcwd(), sanitized_path)
        
        return sanitized_path
    
    def set_database_path(self, path):
        """データベースパスの設定"""
        # パスの検証
        is_valid, message = self.security_manager.validate_file_access(path)
        if not is_valid:
            self.security_manager.log_security_event("INVALID_DB_PATH", f"無効なDBパス: {path} - {message}")
            messagebox.showerror("セキュリティエラー", f"データベースパスが無効です: {message}")
            return False
        
        # 機密パスの場合は暗号化して保存
        if "//" in path or "\\\\" in path:  # ネットワークパスの場合
            try:
                encrypted_path = "encrypted:" + self.security_manager.encrypt_sensitive_data(path)
                self.config["database_path"] = encrypted_path
                self.security_manager.log_security_event("PATH_ENCRYPTED", "データベースパスを暗号化して保存")
            except Exception as e:
                self.security_manager.log_security_event("ENCRYPT_ERROR", f"パス暗号化エラー: {e}")
                self.config["database_path"] = path  # 暗号化に失敗した場合は平文で保存
        else:
            self.config["database_path"] = path
        
        return self.save_config()
    
    def select_database_file(self, parent_window=None):
        """データベースファイルの選択ダイアログ"""
        try:
            # 現在のデータベースパスを初期ディレクトリとして使用
            current_path = self.get_database_path()
            initial_dir = os.path.dirname(current_path) if os.path.dirname(current_path) else os.getcwd()
            
            file_path = filedialog.askopenfilename(
                parent=parent_window,
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
                
                self.set_database_path(file_path)
                return file_path
            
            return None
            
        except Exception as e:
            messagebox.showerror("エラー", f"ファイル選択中にエラーが発生しました:\n{str(e)}")
            return None
    
    def get(self, key, default=None):
        """設定値の取得"""
        return self.config.get(key, default)
    
    def set(self, key, value):
        """設定値の設定"""
        self.config[key] = value
        self.save_config()
    
    def reset_to_defaults(self):
        """設定をデフォルトにリセット"""
        self.config = self.DEFAULT_CONFIG.copy()
        self.save_config()
