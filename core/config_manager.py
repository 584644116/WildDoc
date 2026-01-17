import json
import os
from typing import Dict, Any, Optional

class ConfigManager:
    """配置管理器"""
    
    def __init__(self, config_path: str = "config.json"):
        self.config_path = config_path
        self.config: Dict[str, Any] = {
            'field_mappings': {},
            'field_types': {},
            'required_fields': set()
        }
        self.load_config()
    
    def load_config(self) -> None:
        """加载配置文件"""
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    saved_config = json.load(f)
                    self.config.update(saved_config)
                    # 转换required_fields为集合
                    self.config['required_fields'] = set(self.config['required_fields'])
            except Exception as e:
                print(f"加载配置文件失败: {e}")
    
    def save_config(self) -> bool:
        """保存配置到文件"""
        try:
            # 转换required_fields为列表以便JSON序列化
            config_to_save = self.config.copy()
            config_to_save['required_fields'] = list(self.config['required_fields'])
            
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config_to_save, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存配置文件失败: {e}")
            return False
    
    def set_field_mapping(self, template_field: str, custom_field: str) -> None:
        """设置字段映射
        
        Args:
            template_field: 模板中的字段名
            custom_field: 自定义的字段名
        """
        self.config['field_mappings'][template_field] = custom_field
    
    def set_field_type(self, field: str, field_type: str) -> None:
        """设置字段类型
        
        Args:
            field: 字段名
            field_type: 字段类型
        """
        self.config['field_types'][field] = field_type
    
    def set_required_field(self, field: str, required: bool = True) -> None:
        """设置字段是否必填
        
        Args:
            field: 字段名
            required: 是否必填
        """
        if required:
            self.config['required_fields'].add(field)
        else:
            self.config['required_fields'].discard(field)
    
    def get_field_mapping(self, template_field: str) -> Optional[str]:
        """获取字段映射
        
        Args:
            template_field: 模板中的字段名
            
        Returns:
            映射后的字段名，如果不存在返回None
        """
        return self.config['field_mappings'].get(template_field)
    
    def get_field_type(self, field: str) -> str:
        """获取字段类型
        
        Args:
            field: 字段名
            
        Returns:
            字段类型，默认为'text'
        """
        return self.config['field_types'].get(field, 'text')
    
    def is_field_required(self, field: str) -> bool:
        """检查字段是否必填
        
        Args:
            field: 字段名
            
        Returns:
            是否必填
        """
        return field in self.config['required_fields'] 