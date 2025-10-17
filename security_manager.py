"""
セキュリティ管理モジュール
アプリケーションのセキュリティ機能を提供
"""

import os
import sys
import hashlib
import base64
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC


class SecurityManager:
    """セキュリティ管理クラス"""
    
    def __init__(self):
        self._key = None
    
    
    def _get_encryption_key(self):
        """暗号化キーの取得または生成"""
        if self._key is None:
            # マシン固有のキーを生成
            machine_id = self._get_machine_id()
            password = machine_id.encode()
            salt = b'security_salt_2024'
            
            kdf = PBKDF2HMAC(
                algorithm=hashes.SHA256(),
                length=32,
                salt=salt,
                iterations=100000,
            )
            key = base64.urlsafe_b64encode(kdf.derive(password))
            self._key = key
        
        return self._key
    
    def _get_machine_id(self):
        """マシン固有のIDを取得"""
        try:
            import platform
            import uuid
            
            # マシンのハードウェア情報を組み合わせてIDを生成
            machine_info = f"{platform.node()}-{platform.system()}-{platform.machine()}"
            machine_id = hashlib.sha256(machine_info.encode()).hexdigest()[:16]
            return machine_id
        except Exception:
            # フォールバック: 固定値を使用
            return "default_machine_id"
    
    def encrypt_sensitive_data(self, data):
        """機密データの暗号化"""
        try:
            key = self._get_encryption_key()
            fernet = Fernet(key)
            encrypted_data = fernet.encrypt(data.encode())
            return base64.urlsafe_b64encode(encrypted_data).decode()
        except Exception as e:
            # ログ機能を無効化
            return data  # 暗号化に失敗した場合は平文を返す
    
    def decrypt_sensitive_data(self, encrypted_data):
        """機密データの復号化"""
        try:
            key = self._get_encryption_key()
            fernet = Fernet(key)
            decoded_data = base64.urlsafe_b64decode(encrypted_data.encode())
            decrypted_data = fernet.decrypt(decoded_data)
            return decrypted_data.decode()
        except Exception as e:
            # ログ機能を無効化
            return encrypted_data  # 復号化に失敗した場合は元のデータを返す
    
    def sanitize_path(self, path):
        """パスのサニタイズ"""
        if not path:
            return None
        
        # 危険な文字を除去
        dangerous_chars = ['..', '~', '$', '`', '|', '&', ';', '(', ')', '<', '>']
        for char in dangerous_chars:
            if char in path:
                # ログ機能を無効化
                return None
        
        # パスの正規化
        try:
            normalized_path = os.path.normpath(path)
            return normalized_path
        except Exception as e:
            # ログ機能を無効化
            return None
    
    def validate_file_access(self, file_path):
        """ファイルアクセス権限の検証"""
        try:
            # パスのサニタイズ
            sanitized_path = self.sanitize_path(file_path)
            if not sanitized_path:
                return False, "無効なパスです"
            
            # ファイルの存在確認
            if not os.path.exists(sanitized_path):
                return False, "ファイルが存在しません"
            
            # 読み取り権限の確認
            if not os.access(sanitized_path, os.R_OK):
                return False, "ファイルの読み取り権限がありません"
            
            # ファイルサイズの制限（100MB以下）
            file_size = os.path.getsize(sanitized_path)
            if file_size > 100 * 1024 * 1024:  # 100MB
                return False, "ファイルサイズが大きすぎます"
            
            return True, "OK"
            
        except Exception as e:
            # ログ機能を無効化
            return False, "ファイルアクセスエラー"
    
    def sanitize_error_message(self, error_message):
        """エラーメッセージのサニタイズ"""
        # 機密情報を含む可能性のある文字列をマスク
        sensitive_patterns = [
            r'password', r'passwd', r'pwd',
            r'token', r'key', r'secret',
            r'connection', r'connect',
            r'database', r'db',
            r'file://', r'http://', r'https://'
        ]
        
        import re
        sanitized = error_message
        for pattern in sensitive_patterns:
            sanitized = re.sub(pattern, '[MASKED]', sanitized, flags=re.IGNORECASE)
        
        return sanitized
    
    def log_security_event(self, event_type, details):
        """セキュリティイベントのログ記録（無効化）"""
        # ログ機能を無効化
        pass
    
    def validate_input(self, input_data, input_type):
        """入力値の検証"""
        if not input_data:
            return False, "入力値が空です"
        
        # 文字列長の制限
        if len(str(input_data)) > 1000:
            return False, "入力値が長すぎます"
        
        # 危険な文字のチェック
        dangerous_chars = ['<', '>', '"', "'", '&', ';', '|', '`', '$']
        for char in dangerous_chars:
            if char in str(input_data):
                self.log_security_event("DANGEROUS_INPUT", f"危険な文字検出: {char}")
                return False, f"使用できない文字が含まれています: {char}"
        
        return True, "OK"
