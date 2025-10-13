"""
設定管理モジュール
アプリケーション全体で共有する設定値の読み書きを担う
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
        "database_path": "不良情報記録.accdb",
        "window_geometry": "1000x700",
        "last_product_number": "",
        "default_aql": 0.25,
        "default_ltpd": 1.0,
        "default_alpha": 5.0,
        "default_beta": 10.0,
        "default_c_value": 0,
        "inspection_mode": "standard"
    }

    # 3段階検査区分プリセット
    INSPECTION_MODES = {
        "tightened": {
            "label": "強化",
            "aql": 0.10,
            "ltpd": 0.50,
            "alpha": 3.0,
            "beta": 5.0,
            "c_value": 0,
            "description": "初期流動・不具合再発時"
        },
        "standard": {
            "label": "標準",
            "aql": 0.25,
            "ltpd": 1.00,
            "alpha": 5.0,
            "beta": 10.0,
            "c_value": 0,
            "description": "通常ロット"
        },
        "reduced": {
            "label": "緩和",
            "aql": 0.40,
            "ltpd": 1.50,
            "alpha": 10.0,
            "beta": 15.0,
            "c_value": 0,
            "description": "安定生産・顧客信頼製品"
        }
    }

    def __init__(self):
        self.security_manager = SecurityManager()
        self.config = self._load_config()

    def _load_config(self):
        """設定ファイルの読み込み"""
        try:
            if os.path.exists(self.CONFIG_FILE):
                with open(self.CONFIG_FILE, "r", encoding="utf-8") as handle:
                    loaded = json.load(handle)
                    merged_config = self.DEFAULT_CONFIG.copy()
                    merged_config.update(loaded)
                    merged_config["inspection_mode"] = self._normalize_mode_key(
                        merged_config.get("inspection_mode", self.DEFAULT_CONFIG["inspection_mode"])
                    )
                    return merged_config

            defaults = self.DEFAULT_CONFIG.copy()
            defaults["inspection_mode"] = self._normalize_mode_key(defaults["inspection_mode"])
            return defaults
        except Exception as exc:
            print(f"設定ファイルの読み込みエラー: {exc}")
            defaults = self.DEFAULT_CONFIG.copy()
            defaults["inspection_mode"] = self._normalize_mode_key(defaults["inspection_mode"])
            return defaults

    def save_config(self):
        """設定ファイルの保存"""
        try:
            with open(self.CONFIG_FILE, "w", encoding="utf-8") as handle:
                json.dump(self.config, handle, ensure_ascii=False, indent=2)
            return True
        except Exception as exc:
            print(f"設定ファイルの保存エラー: {exc}")
            return False

    def get_database_path(self):
        """データベースパスの取得"""
        db_path = self.config.get("database_path", self.DEFAULT_CONFIG["database_path"])

        if db_path.startswith("encrypted:"):
            try:
                encrypted_data = db_path[10:]
                db_path = self.security_manager.decrypt_sensitive_data(encrypted_data)
            except Exception as exc:
                self.security_manager.log_security_event("DECRYPT_ERROR", f"パス復号化エラー: {exc}")
                return self.DEFAULT_CONFIG["database_path"]

        sanitized_path = self.security_manager.sanitize_path(db_path)
        if not sanitized_path:
            self.security_manager.log_security_event("INVALID_PATH", f"無効なパス: {db_path}")
            return self.DEFAULT_CONFIG["database_path"]

        if not os.path.isabs(sanitized_path):
            sanitized_path = os.path.join(os.getcwd(), sanitized_path)

        return sanitized_path

    def set_database_path(self, path):
        """データベースパスの設定"""
        is_valid, message = self.security_manager.validate_file_access(path)
        if not is_valid:
            self.security_manager.log_security_event(
                "INVALID_DB_PATH",
                f"無効なDBパス: {path} - {message}"
            )
            messagebox.showerror("セキュリティエラー", f"データベースパスが無効です: {message}")
            return False

        if "//" in path or "\\\\" in path:
            try:
                encrypted_path = "encrypted:" + self.security_manager.encrypt_sensitive_data(path)
                self.config["database_path"] = encrypted_path
                self.security_manager.log_security_event("PATH_ENCRYPTED", "データベースパスを暗号化して保存しました")
            except Exception as exc:
                self.security_manager.log_security_event("ENCRYPT_ERROR", f"パス暗号化エラー: {exc}")
                self.config["database_path"] = path
        else:
            self.config["database_path"] = path

        return self.save_config()

    def select_database_file(self, parent_window=None):
        """データベースファイルの選択ダイアログ"""
        try:
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
                try:
                    rel_path = os.path.relpath(file_path, os.getcwd())
                    if not rel_path.startswith(".."):
                        file_path = rel_path
                except ValueError:
                    pass

                self.set_database_path(file_path)
                return file_path

            return None

        except Exception as exc:
            sanitized_error = self.security_manager.sanitize_error_message(str(exc))
            messagebox.showerror("エラー", f"ファイル選択中にエラーが発生しました:\n{sanitized_error}")
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
        self.config["inspection_mode"] = self._normalize_mode_key(self.config["inspection_mode"])
        self.save_config()

    def get_inspection_mode(self):
        """現在の検査区分キーを取得"""
        return self._normalize_mode_key(self.config.get("inspection_mode"))

    def get_inspection_mode_label(self, mode_key=None):
        """検査区分キーに対応する表示名を取得"""
        key = self._normalize_mode_key(mode_key or self.get_inspection_mode())
        return self.INSPECTION_MODES[key]["label"]

    def get_inspection_mode_choices(self):
        """検査区分選択肢を取得 (key -> label)"""
        return {key: preset["label"] for key, preset in self.INSPECTION_MODES.items()}

    def get_inspection_mode_details(self, mode_key=None):
        """検査区分の詳細情報を取得"""
        key = self._normalize_mode_key(mode_key or self.get_inspection_mode())
        return self.INSPECTION_MODES[key].copy()

    def apply_inspection_mode(self, mode_key, persist=True):
        """検査区分プリセットを適用し、必要に応じて設定へ保存"""
        normalized_key = self._normalize_mode_key(mode_key)
        preset = self.INSPECTION_MODES[normalized_key]

        if persist:
            self.config["inspection_mode"] = normalized_key
            self.config["default_aql"] = preset["aql"]
            self.config["default_ltpd"] = preset["ltpd"]
            self.config["default_alpha"] = preset["alpha"]
            self.config["default_beta"] = preset["beta"]
            self.config["default_c_value"] = preset["c_value"]
            self.save_config()

        return preset.copy()

    def _normalize_mode_key(self, mode_key):
        """設定ファイル内の検査区分指定を内部キーへ正規化"""
        if mode_key in self.INSPECTION_MODES:
            return mode_key

        if isinstance(mode_key, str):
            for key, preset in self.INSPECTION_MODES.items():
                if preset["label"] == mode_key:
                    return key

        return "standard"
