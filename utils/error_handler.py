from typing import Optional, Callable
from PyQt5.QtWidgets import QMessageBox
import traceback
import sys
from .logger import Logger
import functools

logger = Logger('error_handler')

def show_error_dialog(title: str, message: str, parent=None):
    """显示错误对话框"""
    QMessageBox.critical(parent, title, message)

class ErrorHandler:
    """错误处理器"""
    
    @staticmethod
    def handle(func: Callable):
        """错误处理装饰器"""
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                # 获取错误信息
                error_info = traceback.format_exc()
                
                # 记录日志
                logger.error(f"Function {func.__name__} failed: {str(e)}\n{error_info}")
                
                # 显示错误对话框
                show_error_dialog(
                    "错误",
                    f"操作失败: {str(e)}",
                    parent=kwargs.get('parent')
                )
                
                return None
        return wrapper
    
    @staticmethod
    def handle_gui(func: Callable):
        """GUI错误处理装饰器"""
        @functools.wraps(func)
        def wrapper(self, *args, **kwargs):
            try:
                return func(self, *args, **kwargs)
            except Exception as e:
                # 获取错误信息
                error_info = traceback.format_exc()
                
                # 记录日志
                logger.error(f"GUI operation failed: {str(e)}\n{error_info}")
                
                # 显示错误对话框
                show_error_dialog(
                    "错误",
                    f"操作失败: {str(e)}",
                    parent=self
                )
                
                return None
        return wrapper

def global_exception_handler(exctype, value, tb):
    """全局异常处理器"""
    # 获取错误信息
    error_info = ''.join(traceback.format_exception(exctype, value, tb))
    
    # 记录日志
    logger.critical(f"Uncaught exception:\n{error_info}")
    
    # 显示错误对话框
    show_error_dialog(
        "严重错误",
        f"程序遇到未处理的错误:\n{str(value)}"
    )

# 设置全局异常处理器
sys.excepthook = global_exception_handler 