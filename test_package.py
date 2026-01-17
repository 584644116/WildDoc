# -*- coding: utf-8 -*-
import os
import sys
import io

# 设置输出编码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 测试打包后的应用是否包含必要的依赖
print("=" * 50)
print("测试打包结果")
print("=" * 50)

dist_dir = r"dist\Word文档生成器"
if os.path.exists(dist_dir):
    print(f"[OK] 找到输出目录: {dist_dir}")
    
    exe_file = os.path.join(dist_dir, "Word文档生成器.exe")
    if os.path.exists(exe_file):
        size_mb = os.path.getsize(exe_file) / (1024 * 1024)
        print(f"[OK] 找到可执行文件: {exe_file}")
        print(f"  文件大小: {size_mb:.2f} MB")
    else:
        print(f"[ERROR] 未找到可执行文件")
    
    internal_dir = os.path.join(dist_dir, "_internal")
    if os.path.exists(internal_dir):
        print(f"[OK] 找到依赖目录: {internal_dir}")
        
        # 检查关键文件
        pyz_file = os.path.join(internal_dir, "PYZ-00.pyz")
        if os.path.exists(pyz_file):
            print(f"[OK] 找到 Python 库压缩包: PYZ-00.pyz")
            
            # 尝试检查是否包含 openpyxl
            try:
                import zipfile
                with zipfile.ZipFile(pyz_file, 'r') as z:
                    files = z.namelist()
                    openpyxl_files = [f for f in files if 'openpyxl' in f.lower()]
                    pandas_files = [f for f in files if 'pandas' in f.lower()]
                    
                    if openpyxl_files:
                        print(f"[OK] 包含 openpyxl 模块 ({len(openpyxl_files)} 个文件)")
                    else:
                        print(f"[ERROR] 未找到 openpyxl 模块")
                    
                    if pandas_files:
                        print(f"[OK] 包含 pandas 模块 ({len(pandas_files)} 个文件)")
                    else:
                        print(f"[ERROR] 未找到 pandas 模块")
            except Exception as e:
                print(f"[WARN] 无法检查 PYZ 内容: {e}")
        
        # 检查 config 和 docs 目录
        config_dir = os.path.join(internal_dir, "config")
        docs_dir = os.path.join(internal_dir, "docs")
        
        if os.path.exists(config_dir):
            print(f"[OK] 包含 config 目录")
        else:
            print(f"[WARN] 未找到 config 目录")
            
        if os.path.exists(docs_dir):
            print(f"[OK] 包含 docs 目录")
        else:
            print(f"[WARN] 未找到 docs 目录")
    else:
        print(f"[ERROR] 未找到依赖目录")
else:
    print(f"[ERROR] 未找到输出目录: {dist_dir}")

print("=" * 50)
print("测试完成")
print("=" * 50)