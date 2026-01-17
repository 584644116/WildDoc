import unittest
import os
import json
from core.config_manager import ConfigManager

class TestConfigManager(unittest.TestCase):
    def setUp(self):
        """测试前准备"""
        self.test_config_path = 'test_config.json'
        self.config_manager = ConfigManager(self.test_config_path)
    
    def test_field_mapping(self):
        """测试字段映射功能"""
        # 设置映射
        self.config_manager.set_field_mapping('template_field', 'excel_field')
        
        # 验证映射
        self.assertEqual(
            self.config_manager.get_field_mapping('template_field'),
            'excel_field'
        )
    
    def test_field_type(self):
        """测试字段类型功能"""
        # 设置类型
        self.config_manager.set_field_type('field', '数字')
        
        # 验证类型
        self.assertEqual(
            self.config_manager.get_field_type('field'),
            '数字'
        )
    
    def test_required_field(self):
        """测试必填字段功能"""
        # 设置必填
        self.config_manager.set_required_field('field', True)
        
        # 验证必填
        self.assertTrue(self.config_manager.is_field_required('field'))
        
        # 取消必填
        self.config_manager.set_required_field('field', False)
        
        # 验证非必填
        self.assertFalse(self.config_manager.is_field_required('field'))
    
    def test_save_load_config(self):
        """测试配置保存和加载"""
        # 设置测试数据
        self.config_manager.set_field_mapping('field1', 'excel1')
        self.config_manager.set_field_type('field1', '文本')
        self.config_manager.set_required_field('field1', True)
        
        # 保存配置
        self.assertTrue(self.config_manager.save_config())
        
        # 创建新的配置管理器
        new_config = ConfigManager(self.test_config_path)
        
        # 验证加载的配置
        self.assertEqual(new_config.get_field_mapping('field1'), 'excel1')
        self.assertEqual(new_config.get_field_type('field1'), '文本')
        self.assertTrue(new_config.is_field_required('field1'))
    
    def tearDown(self):
        """测试后清理"""
        if os.path.exists(self.test_config_path):
            os.remove(self.test_config_path) 