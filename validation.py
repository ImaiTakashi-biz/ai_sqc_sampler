"""
入力検証モジュール
ユーザー入力の検証を管理
"""

import re
from datetime import datetime
from constants import InspectionConstants
from security_manager import SecurityManager


class InputValidator:
    """入力値検証クラス"""
    
    def __init__(self):
        self.security_manager = SecurityManager()
    
    @staticmethod
    def validate_all_inputs(product_number, lot_size_str, conf_str, c_str, start_date_str=None, end_date_str=None):
        """全入力値の検証"""
        errors = []
        validated_data = {}
        
        # 品番の検証
        validator = InputValidator()
        product_error = validator.validate_product_number(product_number)
        if product_error:
            errors.append(product_error)
        else:
            validated_data['product_number'] = product_number.strip()
        
        # ロットサイズの検証
        lot_size_error, lot_size = InputValidator.validate_lot_size(lot_size_str)
        if lot_size_error:
            errors.append(lot_size_error)
        else:
            validated_data['lot_size'] = lot_size
        
        # 信頼度の検証
        conf_error, confidence_level = InputValidator.validate_confidence_level(conf_str)
        if conf_error:
            errors.append(conf_error)
        else:
            validated_data['confidence_level'] = confidence_level
        
        # c値の検証
        c_error, c_value = InputValidator.validate_c_value(c_str, lot_size)
        if c_error:
            errors.append(c_error)
        else:
            validated_data['c_value'] = c_value
        
        # 日付範囲の検証
        date_error, dates = InputValidator.validate_date_range(start_date_str, end_date_str)
        if date_error:
            errors.append(date_error)
        else:
            validated_data['start_date'], validated_data['end_date'] = dates
        
        return len(errors) == 0, errors, validated_data
    
    def validate_product_number(self, product_number):
        """品番の検証"""
        if not product_number or not product_number.strip():
            return "品番は必須です"
        
        product_number = product_number.strip()
        
        # セキュリティ検証
        is_valid, message = self.security_manager.validate_input(product_number, "product_number")
        if not is_valid:
            self.security_manager.log_security_event("INVALID_PRODUCT_NUMBER", f"無効な品番: {product_number}")
            return message
        
        if len(product_number) > InspectionConstants.MAX_PRODUCT_NUMBER_LENGTH:
            return f"品番は{InspectionConstants.MAX_PRODUCT_NUMBER_LENGTH}文字以内で入力してください"
        
        # 特殊文字のチェック
        for char in InspectionConstants.INVALID_CHARS:
            if char in product_number:
                self.security_manager.log_security_event("DANGEROUS_CHAR_IN_PRODUCT_NUMBER", f"危険な文字: {char}")
                return f"品番に使用できない文字「{char}」が含まれています"
        
        return None
    
    @staticmethod
    def validate_lot_size(lot_size_str):
        """ロットサイズの検証"""
        if not lot_size_str or not lot_size_str.strip():
            return "数量は必須です", None
        
        try:
            lot_size = int(lot_size_str.strip())
            
            if lot_size < InspectionConstants.MIN_LOT_SIZE:
                return f"数量は{InspectionConstants.MIN_LOT_SIZE}以上の整数で入力してください", None
            
            if lot_size > InspectionConstants.MAX_LOT_SIZE:
                return f"数量は{InspectionConstants.MAX_LOT_SIZE:,}以下で入力してください", None
            
            return None, lot_size
            
        except ValueError:
            return "数量は整数で入力してください", None
    
    @staticmethod
    def validate_confidence_level(conf_str):
        """信頼度の検証"""
        if not conf_str or not conf_str.strip():
            return None, InspectionConstants.DEFAULT_CONFIDENCE
        
        try:
            confidence = float(conf_str.strip())
            
            if not 0 < confidence <= 100:
                return "信頼度は0より大きく100以下の数値で入力してください", None
            
            if confidence < InspectionConstants.MIN_CONFIDENCE:
                return f"信頼度は{InspectionConstants.MIN_CONFIDENCE}%以上で入力してください", None
            
            if confidence > InspectionConstants.MAX_CONFIDENCE:
                return f"信頼度は{InspectionConstants.MAX_CONFIDENCE}%以下で入力してください", None
            
            return None, confidence
            
        except ValueError:
            return "信頼度は数値で入力してください", None
    
    @staticmethod
    def validate_c_value(c_str, lot_size=None):
        """c値の検証"""
        if not c_str or not c_str.strip():
            return None, InspectionConstants.DEFAULT_C_VALUE
        
        try:
            c_value = int(c_str.strip())
            
            if c_value < 0:
                return "c値は0以上の整数で入力してください", None
            
            if c_value > InspectionConstants.MAX_C_VALUE:
                return f"c値は{InspectionConstants.MAX_C_VALUE}以下で入力してください", None
            
            # ロットサイズに対するc値の妥当性チェック
            if lot_size and c_value > lot_size * 0.1:
                return "c値が大きすぎます（ロットサイズの10%以下推奨）", None
            
            return None, c_value
            
        except ValueError:
            return "c値は整数で入力してください", None
    
    @staticmethod
    def validate_date_range(start_date_str, end_date_str):
        """日付範囲の検証"""
        start_date = None
        end_date = None
        
        if start_date_str and start_date_str.strip():
            try:
                start_date = datetime.strptime(start_date_str.strip(), "%Y-%m-%d").date()
            except ValueError:
                return "開始日はYYYY-MM-DD形式で入力してください", (None, None)
        
        if end_date_str and end_date_str.strip():
            try:
                end_date = datetime.strptime(end_date_str.strip(), "%Y-%m-%d").date()
            except ValueError:
                return "終了日はYYYY-MM-DD形式で入力してください", (None, None)
        
        if start_date and end_date and start_date > end_date:
            return "開始日は終了日より前の日付を入力してください", (None, None)
        
        # 未来の日付チェック
        today = datetime.now().date()
        if start_date and start_date > today:
            return "開始日に未来の日付は入力できません", (None, None)
        if end_date and end_date > today:
            return "終了日に未来の日付は入力できません", (None, None)
        
        return None, (start_date, end_date)