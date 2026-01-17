import unittest
import pandas as pd
from core.config_manager import ConfigManager
from utils.validator import ConfigValidator, DataValidator, validate_excel_columns

class TestValidator(unittest.TestCase):
    def setUp(self):
        """测试前准备"""
        self.config_manager = ConfigManager()
        self.data_validator = DataValidator(self.config_manager)
    
    def test_field_mapping_validation(self):
        """测试字段映射验证"""
        # 测试空映射
        mapping = {'field1': '', 'field2': 'excel2'}
        valid, errors = ConfigValidator.validate_field_mapping(mapping)
        self.assertFalse(valid)
        self.assertTrue(any('field1' in error for error in errors))
        
        # 测试重复映射
        mapping = {'field1': 'excel1', 'field2': 'excel1'}
        valid, errors = ConfigValidator.validate_field_mapping(mapping)
        self.assertFalse(valid)
        self.assertTrue(any('excel1' in error for error in errors))
        
        # 测试有效映射
        mapping = {'field1': 'excel1', 'field2': 'excel2'}
        valid, errors = ConfigValidator.validate_field_mapping(mapping)
        self.assertTrue(valid)
        self.assertEqual(len(errors), 0)
    
    def test_excel_data_validation(self):
        """测试Excel数据验证"""
        # 创建测试数据
        df = pd.DataFrame({
            'name': ['张三', None, '李四'],
            'age': ['25', 'abc', '30'],
            'date': ['2024-01-01', '无效日期', '2024-01-02']
        })
        
        # 配置字段
        self.config_manager.set_field_mapping('姓名', 'name')
        self.config_manager.set_field_mapping('年龄', 'age')
        self.config_manager.set_field_mapping('日期', 'date')
        
        self.config_manager.set_field_type('年龄', '数字')
        self.config_manager.set_field_type('日期', '日期')
        
        self.config_manager.set_required_field('姓名', True)
        
        # 验证数据
        valid, errors = self.data_validator.validate_excel_data(df)
        self.assertFalse(valid)
        self.assertTrue(any('name' in error for error in errors))  # 空值检查
        self.assertTrue(any('age' in error for error in errors))   # 数字检查
        self.assertTrue(any('date' in error for error in errors))  # 日期检查
    
    def test_excel_columns_validation(self):
        """测试Excel列验证"""
        df = pd.DataFrame({
            'name': [],
            'age': []
        })
        
        # 测试缺少列
        required_columns = ['name', 'age', 'date']
        valid, errors = validate_excel_columns(df, required_columns)
        self.assertFalse(valid)
        self.assertTrue(any('date' in error for error in errors))
        
        # 测试所有列都存在
        required_columns = ['name', 'age']
        valid, errors = validate_excel_columns(df, required_columns)
        self.assertTrue(valid)
        self.assertEqual(len(errors), 0) 