from typing import Dict, Any, List, Tuple
import pandas as pd
from .logger import Logger

logger = Logger('validator')

class ConfigValidator:
    """配置验证器"""
    
    @staticmethod
    def validate_field_mapping(mapping: Dict[str, str]) -> Tuple[bool, List[str]]:
        """验证字段映射配置
        
        Args:
            mapping: 字段映射字典
            
        Returns:
            (是否有效, 错误信息列表)
        """
        errors = []
        
        # 检查空值
        empty_fields = [k for k, v in mapping.items() if not v]
        if empty_fields:
            errors.append(f"以下字段未映射: {', '.join(empty_fields)}")
        
        # 检查重复映射
        value_counts = {}
        for k, v in mapping.items():
            if v in value_counts:
                value_counts[v].append(k)
            else:
                value_counts[v] = [k]
        
        duplicates = {v: ks for v, ks in value_counts.items() if len(ks) > 1}
        if duplicates:
            for v, ks in duplicates.items():
                errors.append(f"多个字段映射到同一个Excel列 '{v}': {', '.join(ks)}")
        
        return len(errors) == 0, errors

class DataValidator:
    """数据验证器"""
    
    def __init__(self, config_manager):
        self.config_manager = config_manager
        self.logger = Logger('data_validator')
    
    def validate_excel_data(self, df: pd.DataFrame) -> Tuple[bool, List[str]]:
        """验证Excel数据
        
        Args:
            df: pandas DataFrame对象
            
        Returns:
            (是否有效, 错误信息列表)
        """
        errors = []
        
        # 检查必填字段
        for field in self.config_manager.config['required_fields']:
            mapped_field = self.config_manager.get_field_mapping(field)
            if mapped_field and mapped_field in df.columns:
                # 检查空值
                null_rows = df[df[mapped_field].isnull()].index.tolist()
                if null_rows:
                    errors.append(
                        f"必填字段 '{mapped_field}' 在以下行中为空: {null_rows}"
                    )
        
        # 检查字段类型
        for field, field_type in self.config_manager.config['field_types'].items():
            mapped_field = self.config_manager.get_field_mapping(field)
            if mapped_field and mapped_field in df.columns:
                if field_type == '数字':
                    try:
                        pd.to_numeric(df[mapped_field])
                    except Exception as e:
                        errors.append(
                            f"字段 '{mapped_field}' 包含非数字值"
                        )
                elif field_type == '日期':
                    try:
                        pd.to_datetime(df[mapped_field])
                    except Exception as e:
                        errors.append(
                            f"字段 '{mapped_field}' 包含非日期值"
                        )
        
        return len(errors) == 0, errors
    
    def validate_template_fields(self, template_fields: List[str]) -> Tuple[bool, List[str]]:
        """验证模板字段
        
        Args:
            template_fields: 模板中的字段列表
            
        Returns:
            (是否有效, 错误信息列表)
        """
        errors = []
        
        # 检查所有字段是否都有映射
        for field in template_fields:
            if not self.config_manager.get_field_mapping(field):
                errors.append(f"模板字段 '{field}' 未配置映射")
        
        return len(errors) == 0, errors

def validate_excel_columns(df: pd.DataFrame, required_columns: List[str]) -> Tuple[bool, List[str]]:
    """验证Excel文件的列
    
    Args:
        df: pandas DataFrame对象
        required_columns: 必需的列名列表
        
    Returns:
        (是否有效, 错误信息列表)
    """
    errors = []
    
    # 检查必需的列是否存在
    missing_columns = set(required_columns) - set(df.columns)
    if missing_columns:
        errors.append(f"Excel文件缺少以下列: {', '.join(missing_columns)}")
    
    return len(errors) == 0, errors 