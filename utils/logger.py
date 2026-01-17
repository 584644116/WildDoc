import logging
import os
from datetime import datetime

class Logger:
    """日志管理器"""
    
    def __init__(self, name: str = 'word_generator'):
        self.logger = logging.getLogger(name)
        self.logger.setLevel(logging.DEBUG)
        
        # 创建日志目录
        log_dir = 'logs'
        os.makedirs(log_dir, exist_ok=True)
        
        # 文件处理器
        log_file = os.path.join(
            log_dir, 
            f'{datetime.now().strftime("%Y%m%d")}.log'
        )
        file_handler = logging.FileHandler(
            log_file, 
            encoding='utf-8'
        )
        file_handler.setLevel(logging.DEBUG)
        
        # 控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # 设置格式
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        # 添加处理器
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
    
    def debug(self, msg):
        self.logger.debug(msg)
    
    def info(self, msg):
        self.logger.info(msg)
    
    def warning(self, msg):
        self.logger.warning(msg)
    
    def error(self, msg):
        self.logger.error(msg)
    
    def critical(self, msg):
        self.logger.critical(msg) 