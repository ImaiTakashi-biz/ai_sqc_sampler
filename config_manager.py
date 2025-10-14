"""
設定管理モジュール
アプリケーション全体で使用する設定値の読み書きを担当する
"""

import os
import json
from copy import deepcopy
from tkinter import filedialog, messagebox

from security_manager import SecurityManager


# 検査区分の定義
# 各検査区分は設定画面でAQL、LTPD、α、β、c値を個別に設定可能
INSPECTION_MODE_META = {
    "tightened": {
        "label": "強化",
        "description": "初期流動・不具合再発時"
    },
    "standard": {
        "label": "標準",
        "description": "通常ロット"
    },
    "reduced": {
        "label": "緩和",
        "description": "安定生産・顧客信頼製品"
    }
}

# 各検査区分のデフォルト値
# ユーザーは設定画面でこれらの値を変更可能
DEFAULT_PRESETS = {
    "tightened": {"aql": 0.10, "ltpd": 0.50, "alpha": 3.0, "beta": 5.0, "c_value": 0},
    "standard": {"aql": 0.25, "ltpd": 1.0, "alpha": 5.0, "beta": 10.0, "c_value": 0},
    "reduced": {"aql": 0.40, "ltpd": 1.5, "alpha": 10.0, "beta": 15.0, "c_value": 0}
}


class ConfigManager:
    """設定管理クラス"""

    CONFIG_FILE = "app_config.json"

    DEFAULT_CONFIG = {
        "database_path": "不良情報記録.accdb",
        "window_geometry": "1000x700",
        "last_product_number": "",
        "default_aql": 0.25,
        "default_ltpd": 1.0,
        "default_alpha": 5.0,
        "default_beta": 10.0,
        "default_c_value": 0,
        "inspection_mode": "standard",
        "inspection_presets": deepcopy(DEFAULT_PRESETS)
    }

    def __init__(self):
        self.security_manager = SecurityManager()
        self.config = self._load_config()
        # アプリ起動時は毎回標準検査に設定（検査区分の設定値は保持）
        self.config["inspection_mode"] = "standard"
        # 検査区分の設定値は上書きせず、default_*項目のみ標準検査の値に更新
        self._sync_legacy_defaults_for_startup()
        # 設定ファイルが存在しない場合は保存
        if not os.path.exists(self.CONFIG_FILE):
            self.save_config()
        else:
            # 起動時の検査区分を標準検査に設定して保存
            self.save_config()

    # ------------------------------------------------------------------
    # 設定ファイルの読み書き
    # ------------------------------------------------------------------
    def _load_config(self):
        """設定ファイルの読み込み"""
        try:
            if os.path.exists(self.CONFIG_FILE):
                with open(self.CONFIG_FILE, "r", encoding="utf-8") as handle:
                    loaded = json.load(handle)

                merged = deepcopy(self.DEFAULT_CONFIG)
                for key, value in loaded.items():
                    if key == "inspection_presets":
                        continue
                    merged[key] = value

                presets = deepcopy(self.DEFAULT_CONFIG["inspection_presets"])
                user_presets = loaded.get("inspection_presets", {})
                for mode_key in INSPECTION_MODE_META.keys():
                    if mode_key in user_presets and isinstance(user_presets[mode_key], dict):
                        for param in ["aql", "ltpd", "alpha", "beta", "c_value"]:
                            if param in user_presets[mode_key]:
                                try:
                                    value = float(user_presets[mode_key][param])
                                except (TypeError, ValueError):
                                    continue
                                presets[mode_key][param] = value if param != "c_value" else int(value)
                legacy_defaults = {
                    "aql": loaded.get("default_aql"),
                    "ltpd": loaded.get("default_ltpd"),
                    "alpha": loaded.get("default_alpha"),
                    "beta": loaded.get("default_beta"),
                    "c_value": loaded.get("default_c_value")
                }
                if any(value is not None for value in legacy_defaults.values()):
                    standard = presets.get("standard", {}).copy()
                    try:
                        if legacy_defaults["aql"] is not None:
                            standard["aql"] = float(legacy_defaults["aql"])
                        if legacy_defaults["ltpd"] is not None:
                            standard["ltpd"] = float(legacy_defaults["ltpd"])
                        if legacy_defaults["alpha"] is not None:
                            standard["alpha"] = float(legacy_defaults["alpha"])
                        if legacy_defaults["beta"] is not None:
                            standard["beta"] = float(legacy_defaults["beta"])
                        if legacy_defaults["c_value"] is not None:
                            standard["c_value"] = int(legacy_defaults["c_value"])
                    except (TypeError, ValueError):
                        pass
                    else:
                        presets["standard"] = standard
                merged["inspection_presets"] = presets
                merged["inspection_mode"] = self._normalize_mode_key(
                    merged.get("inspection_mode", self.DEFAULT_CONFIG["inspection_mode"])
                )
                return merged

            defaults = deepcopy(self.DEFAULT_CONFIG)
            defaults["inspection_mode"] = self._normalize_mode_key(defaults["inspection_mode"])
            return defaults
        except Exception as exc:
            print(f"設定ファイルの読み込みエラー: {exc}")
            defaults = deepcopy(self.DEFAULT_CONFIG)
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

    # ------------------------------------------------------------------
    # データベースパス関連
    # ------------------------------------------------------------------
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
        """データベースファイルの参照"""
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

    # ------------------------------------------------------------------
    # 一般設定値の取得・保存
    # ------------------------------------------------------------------
    def get(self, key, default=None):
        """設定値の取得"""
        return self.config.get(key, default)

    def set(self, key, value):
        """設定値の設定"""
        self.config[key] = value
        self.save_config()

    def reset_to_defaults(self):
        """設定をデフォルトにリセット"""
        self.config = deepcopy(self.DEFAULT_CONFIG)
        self.config["inspection_mode"] = self._normalize_mode_key(self.config["inspection_mode"])
        self.save_config()

    # ------------------------------------------------------------------
    # 検査区分関連
    # ------------------------------------------------------------------
    def get_inspection_mode(self):
        """現在の検査区分キーを取得"""
        return self._normalize_mode_key(self.config.get("inspection_mode"))

    def get_inspection_mode_label(self, mode_key=None):
        """検査区分キーに対応する表示名を取得"""
        key = self._normalize_mode_key(mode_key or self.get_inspection_mode())
        return INSPECTION_MODE_META[key]["label"]

    def get_inspection_mode_choices(self):
        """検査区分選択肢を取得 (key -> label)"""
        return {key: meta["label"] for key, meta in INSPECTION_MODE_META.items()}

    def get_inspection_mode_details(self, mode_key=None):
        """検査区分の詳細情報を取得"""
        key = self._normalize_mode_key(mode_key or self.get_inspection_mode())
        details = deepcopy(INSPECTION_MODE_META[key])
        preset = deepcopy(self.config.get("inspection_presets", {}).get(key, {}))
        defaults = DEFAULT_PRESETS.get(key, {})
        for param in ["aql", "ltpd", "alpha", "beta", "c_value"]:
            details[param] = preset.get(param, defaults.get(param))
        return details

    def apply_inspection_mode(self, mode_key, persist=True):
        """検査区分プリセットを適用し、必要に応じて設定へ保存"""
        normalized_key = self._normalize_mode_key(mode_key)

        if persist:
            self.config["inspection_mode"] = normalized_key
            self._sync_legacy_defaults(normalized_key)
            self.save_config()

        return self.get_inspection_mode_details(normalized_key)

    def set_inspection_preset(self, mode_key, *, aql, ltpd, alpha, beta, c_value):
        """検査区分ごとのプリセット値を更新"""
        normalized_key = self._normalize_mode_key(mode_key)
        presets = self.config.setdefault("inspection_presets", deepcopy(DEFAULT_PRESETS))
        presets[normalized_key] = {
            "aql": float(aql),
            "ltpd": float(ltpd),
            "alpha": float(alpha),
            "beta": float(beta),
            "c_value": int(c_value)
        }
        if normalized_key == self.get_inspection_mode():
            self._sync_legacy_defaults(normalized_key)
        self.save_config()

    # ------------------------------------------------------------------
    # 内部ユーティリティ
    # ------------------------------------------------------------------
    def _normalize_mode_key(self, mode_key):
        """設定ファイル内の検査区分指定を内部キーへ正規化"""
        if mode_key in INSPECTION_MODE_META:
            return mode_key

        if isinstance(mode_key, str):
            for key, meta in INSPECTION_MODE_META.items():
                if meta["label"] == mode_key:
                    return key

        return "standard"

    def _sync_legacy_defaults_for_startup(self):
        """アプリ起動時：各検査区分の設定値を常にデフォルト値に固定"""
        if not isinstance(self.config, dict):
            return

        # 標準検査の値を取得
        standard_details = DEFAULT_PRESETS.get("standard", {})

        # default_*項目を標準検査の値に更新
        for param in ["aql", "ltpd", "alpha", "beta", "c_value"]:
            legacy_key = f"default_{param if param != 'c_value' else 'c_value'}"
            if legacy_key in self.config:
                self.config[legacy_key] = standard_details.get(param)

        # 各検査区分の設定値を常にデフォルト値に固定
        self.config["inspection_presets"] = deepcopy(DEFAULT_PRESETS)

    def _sync_legacy_defaults(self, mode_key=None):
        """旧default_*項目を現在の検査区分パラメータに合わせる（検査区分の設定値は保持）"""
        if not isinstance(self.config, dict):
            return

        active_key = self._normalize_mode_key(mode_key or self.config.get("inspection_mode"))
        try:
            details = self.get_inspection_mode_details(active_key)
        except Exception:
            details = DEFAULT_PRESETS.get(active_key, DEFAULT_PRESETS.get("standard", {}))

        # 検査区分の設定値は上書きせず、default_*項目のみ更新
        for param in ["aql", "ltpd", "alpha", "beta", "c_value"]:
            legacy_key = f"default_{param if param != 'c_value' else 'c_value'}"
            if legacy_key in self.config:
                self.config[legacy_key] = details.get(param)

        # inspection_presetsが存在しない場合のみデフォルト値を設定
        if "inspection_presets" not in self.config:
            self.config["inspection_presets"] = deepcopy(DEFAULT_PRESETS)
        else:
            # 既存のinspection_presetsの各検査区分の設定値は保持し、不足分のみデフォルト値を補完
            for mode, default_values in DEFAULT_PRESETS.items():
                if mode not in self.config["inspection_presets"]:
                    self.config["inspection_presets"][mode] = default_values.copy()
                else:
                    # 既存の設定値は保持し、不足しているパラメータのみデフォルト値を補完
                    for param, default_value in default_values.items():
                        if param not in self.config["inspection_presets"][mode]:
                            self.config["inspection_presets"][mode][param] = default_value
