import os
from docx import Document
import win32com.client
import pythoncom
from typing import Dict, List
import pandas as pd
import re
from .template_analyzer import TemplateAnalyzer
from utils.logger import Logger

class BatchProcessor:
    """批量处理Word文件"""
    
    def __init__(self):
        self.analyzer = TemplateAnalyzer()
        self.logger = Logger('batch_processor')
        # 初始化 Word 应用
        try:
            pythoncom.CoInitialize()
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = False
        except Exception as e:
            self.logger.error(f"初始化 Word 应用失败: {str(e)}")
            self.word_app = None
    
    def __del__(self):
        """清理资源"""
        if hasattr(self, 'word_app') and self.word_app:
            try:
                self.word_app.Quit()
                pythoncom.CoUninitialize()
            except:
                pass

    def clean_path(self, path: str) -> str:
        """清理文件路径"""
        # 替换非法字符
        path = re.sub(r'[<>:"|?*]', '_', path)
        # 移除路径中的空格
        path = path.replace(' ', '_')
        # 统一使用正斜杠
        path = path.replace('\\', '/')
        return path

    def convert_doc_to_docx(self, doc_path: str) -> str:
        """将 .doc 转换为 .docx"""
        if not self.word_app:
            self.logger.error("Word 应用未初始化")
            return None
            
        try:
            # 清理并构造输出路径
            clean_doc_path = self.clean_path(doc_path)
            docx_path = clean_doc_path + 'x'  # 添加 x 后缀
            
            # 创建临时目录
            temp_dir = os.path.join(os.path.dirname(doc_path), 'temp')
            os.makedirs(temp_dir, exist_ok=True)
            
            # 在临时目录中创建文件副本
            temp_doc = os.path.join(temp_dir, os.path.basename(clean_doc_path))
            temp_docx = temp_doc + 'x'
            
            # 复制原文件到临时目录
            import shutil
            shutil.copy2(doc_path, temp_doc)
            
            try:
                # 打开文档
                doc = self.word_app.Documents.Open(temp_doc)
                # 另存为 .docx
                doc.SaveAs2(temp_docx, FileFormat=16)  # 16 表示 docx 格式
                doc.Close()
                
                # 将转换后的文件复制到目标位置
                shutil.copy2(temp_docx, docx_path)
                
                return docx_path
                
            finally:
                # 清理临时文件
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass
            
        except Exception as e:
            self.logger.error(f"转换文档失败 {doc_path}: {str(e)}")
            return None

    def process_directory(self, root_dir: str) -> None:
        """处理目录下的所有Word文件"""
        if not self.word_app:
            self.logger.error("Word 应用未初始化，无法处理文件")
            return
            
        # 存储每个文件夹的数据
        folder_data = {}
        
        # 遍历目录
        for dirpath, dirnames, filenames in os.walk(root_dir):
            word_files = []
            for filename in filenames:
                if (filename.endswith('.doc') or filename.endswith('.docx')) and not filename.startswith('~$'):
                    word_files.append(filename)
            
            if word_files:  # 如果当前文件夹有Word文件
                folder_name = os.path.basename(dirpath)
                folder_data[dirpath] = []
                
                self.logger.info(f"处理文件夹: {dirpath}")
                
                # 处理当前文件夹中的每个Word文件
                for word_file in word_files:
                    try:
                        # 使用绝对路径并清理路径
                        file_path = os.path.abspath(os.path.join(dirpath, word_file))
                        
                        # 检查文件是否存在和可访问
                        if not os.path.exists(file_path):
                            self.logger.error(f"文件不存在: {file_path}")
                            continue
                        
                        # 如果是 .doc 文件，先转换为 .docx
                        if file_path.endswith('.doc'):
                            docx_path = self.convert_doc_to_docx(file_path)
                            if docx_path is None:
                                continue
                            file_path = docx_path
                        
                        # 分析文件
                        result = self.analyzer.analyze(file_path)
                        # 添加文件名
                        result['fields']['文件名'] = word_file
                        folder_data[dirpath].append(result['fields'])
                        self.logger.info(f"成功处理文件: {word_file}")
                        
                        # 如果是转换的文件，删除临时文件
                        if file_path.endswith('.docx') and word_file.endswith('.doc'):
                            try:
                                os.remove(file_path)
                            except:
                                pass
                            
                    except Exception as e:
                        self.logger.error(f"处理文件失败 {word_file}: {str(e)}")
                        continue
                
                # 将当前文件夹的数据转换为DataFrame并保存
                if folder_data[dirpath]:
                    try:
                        df = pd.DataFrame(folder_data[dirpath])
                        # 设置列顺序
                        columns = ['文件名', '姓名', '性别', '出生年月', '入职日期', '文化程度', '专业',
                                  '岗位', '职称', '现从事专业及年限', '人员现已具备的条件', '考核意见']
                        df = df.reindex(columns=columns)
                        
                        # 保存Excel文件
                        excel_name = f"{folder_name}.xlsx"
                        excel_path = os.path.join(dirpath, excel_name)
                        
                        # 确保输出路径存在
                        os.makedirs(os.path.dirname(excel_path), exist_ok=True)
                        
                        # 保存Excel文件
                        df.to_excel(excel_path, index=False)
                        self.logger.info(f"已保存Excel文件: {excel_path}")
                    except Exception as e:
                        self.logger.error(f"保存Excel失败 {dirpath}: {str(e)}") 