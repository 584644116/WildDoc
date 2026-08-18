# document_generator.py 逐行详解

> 这是一份面向编程小白的详细代码解释文档。每一个知识点都会用最简单的方式讲解，确保你能够完全理解。

---

## 文件开头：导入必要的工具

```python
import os
import re
from datetime import datetime, date
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import pandas as pd
from utils.logger import Logger
```

| 代码 | 解释 |
|------|------|
| `import os` | 操作系统工具包，可以操作文件、创建文件夹等 |
| `import re` | 正则表达式工具，用于匹配和查找复杂的文本模式 |
| `from datetime import datetime, date` | 日期时间工具，处理日期（如 2026-04-14） |
| `from docx import Document` | 读取和创建 Word 文档（.docx 文件）|
| `from docx.shared import Pt` | 设置字体大小（Pt = Point，点）|
| `from docx.oxml.ns import qn` | 处理 Word 文档中的中文字体设置 |
| `import pandas as pd` | 数据分析工具，处理 Excel 表格数据 |
| `from utils.logger import Logger` | 日志工具，记录程序运行信息 |

---

## 类定义

```python
class DocumentGenerator:
    """Word文档生成器"""
```

**面向小白的解释：**
- `class` 是 Python 的"类"关键字，就像创建一个"工具箱"
- `DocumentGenerator` 是这个工具箱的名字
- `"""Word文档生成器"""` 是这个工具箱的说明标签

---

## 初始化方法（构造函数）

```python
def __init__(self):
    self.logger = Logger('document_generator')
    self.template_path = None
    self.template = None
    self.placeholders = {}
    # 映射: 模板占位符 -> Excel 列名，如果不设置则默认为同名映射
    self.field_mapping = {}
    # 命名规则配置
    self.naming_rule = {
        'columns': [],          # 需要作为文件名的列
        'separator': '',        # 分隔符
        'fixed_text': ''        # 追加固定文本
    }
```

**逐行解释：**

| 代码 | 解释 |
|------|------|
| `def __init__(self):` | 初始化方法，创建这个类的"东西"时自动运行 |
| `self.logger = Logger(...)` | 创建一个日志记录器，用来记录程序运行日志 |
| `self.template_path = None` | 模板文件路径，初始为空 |
| `self.template = None` | 模板文档对象，初始为空 |
| `self.placeholders = {}` | 空字典，用来存放找到的占位符 |
| `self.field_mapping = {}` | 字段映射字典，例如 `[姓名] -> 姓名` |
| `self.naming_rule = {...}` | 文件命名规则配置字典 |

**什么是 self？**
- `self` 就像"这个工具箱本身"
- `self.xxx` 表示这个工具箱里的一个"抽屉"或"组件"

---

## 加载模板方法

```python
def load_template(self, template_path: str) -> bool:
    """加载Word模板"""
    try:
        self.template_path = template_path
        self.template = Document(template_path)
        self.placeholders = self._find_placeholders()
        return True
    except Exception as e:
        self.logger.error(f"加载模板失败: {str(e)}")
        return False
```

| 代码 | 解释 |
|------|------|
| `def load_template(...)` | 定义一个"加载模板"的函数 |
| `template_path: str` | 参数：模板文件路径，类型是字符串 |
| `-> bool` | 返回值类型：布尔值（True/False）|
| `try:` | 尝试执行以下代码，如果出错则跳到 except |
| `Document(template_path)` | 用 docx 库打开 Word 文件 |
| `self._find_placeholders()` | 调用下面的方法查找占位符 |
| `except Exception as e:` | 如果出错，执行这里 |
| `self.logger.error(...)` | 记录错误日志 |
| `return False` | 返回"失败" |

**什么是 try...except？**
- 就像"试试这样做，如果不行就算了"
- 防止程序因为小错误就崩溃

---

## 查找占位符的核心方法

```python
def _find_placeholders(self) -> dict:
    """??????????????????????????????????????????+ ??????????????????"""
    placeholders = {}
    placeholder_re = re.compile(r'\[(.*?)\]')
```

| 代码 | 解释 |
|------|------|
| `def _find_placeholders(self)` | 私有方法（以 `_` 开头），查找 `[xxx]` 这样的占位符 |
| `placeholders = {}` | 创建空字典，存放找到的占位符 |
| `re.compile(r'\[(.*?)\]')` | 创建正则表达式，用于匹配 `[xxx]` 格式的文本 |

**什么是正则表达式 `r'\[(.*?)\]'`？**
- `\[` 匹配左边的方括号 `[`
- `(.*?)` 捕获方括号里的内容（任意字符，非贪婪）
- `\]` 匹配右边的方括号 `]`
- `r'...'` 表示原始字符串，不转义

---

## 辅助函数：提取段落文本

```python
def _extract_text_from_paragraph(paragraph):
    text = paragraph.text or ""
    try:
        elem = paragraph._element
        ns = elem.nsmap or {}
        if 'w' not in ns:
            ns = {**ns, 'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        # Include text from textboxes or other nested runs.
        text_nodes = elem.xpath('.//w:t', namespaces=ns)
        if text_nodes:
            text = ''.join(t.text for t in text_nodes if t.text)
    except Exception:
        # Fallback to paragraph.text if low-level XML access fails.
        pass
    return text
```

| 代码 | 解释 |
|------|------|
| `paragraph.text` | 获取段落的文本内容 |
| `elem = paragraph._element` | 获取段落的底层 XML 元素 |
| `elem.nsmap` | XML 的命名空间映射 |
| `elem.xpath('.//w:t', ...)` | XPath 查询，提取所有文本节点 |
| `''.join(...)` | 把所有文本片段连接成完整字符串 |
| `except: pass` | 如果出错就忽略，继续执行 |

**为什么要这样复杂？**
- 因为 Word 文档内部是 XML 格式
- 直接读取 `.text` 可能漏掉某些文本（如文本框里的内容）

---

## 扫描段落查找占位符

```python
def _scan_paragraphs(paragraphs, table_index=None, row=None, cell=None):
    for paragraph in paragraphs:
        text = _extract_text_from_paragraph(paragraph)
        if '[' in text and ']' in text:
            # ??????????????????????????????????????? "?????????[??????] ?????????[??????]"
            for placeholder in placeholder_re.findall(text):
                if not placeholder:
                    continue
                placeholders[placeholder] = {
                    'table_index': table_index,
                    'row': row,
                    'cell': cell,
                    'paragraph': paragraph,
                    'text': text
                }
```

| 代码 | 解释 |
|------|------|
| `for paragraph in paragraphs:` | 遍历每个段落 |
| `text = _extract_text_from_paragraph(paragraph)` | 提取段落文本 |
| `if '[' in text and ']' in text:` | 如果文本里包含 `[` 和 `]` |
| `placeholder_re.findall(text)` | 用正则找出所有匹配的占位符 |
| `placeholders[placeholder] = {...}` | 把找到的占位符信息存到字典 |

---

## 扫描文档的各个部分

```python
# 1) ???????????????????????????????????????????????????
for table_index, table in enumerate(self.template.tables):
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            _scan_paragraphs(cell.paragraphs, table_index=table_index, row=i, cell=j)

# 2) ????????????????????????????????????????????????
_scan_paragraphs(self.template.paragraphs)

# 3) ????????????/??????????????????????????????
for section in self.template.sections:
    for header in [section.header, section.first_page_header, section.even_page_header]:
        _scan_paragraphs(header.paragraphs)
        for t_index, table in enumerate(header.tables):
            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    _scan_paragraphs(cell.paragraphs, table_index=f"header-{t_index}", row=i, cell=j)
    for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
        _scan_paragraphs(footer.paragraphs)
        for t_index, table in enumerate(footer.tables):
            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    _scan_paragraphs(cell.paragraphs, table_index=f"footer-{t_index}", row=i, cell=j)
```

| 代码 | 解释 |
|------|------|
| `for table_index, table in enumerate(...)` | 遍历文档中的所有表格 |
| `table.rows` | 表格的所有行 |
| `row.cells` | 行的所有单元格 |
| `enumerate(...)` | 遍历同时获取索引（0, 1, 2...）|
| `section.header` | 页眉 |
| `section.footer` | 页脚 |
| `section.first_page_header` | 首页页眉 |
| `section.even_page_header` | 偶数页页眉 |

**为什么要扫描这么多地方？**
- 占位符可能出现在：表格里、正文里、页眉里、页脚里
- 必须全部扫描才能找到所有占位符

---

## 设置映射和命名规则

```python
def set_mapping(self, mapping: dict):
    """设置占位符与Excel列的映射"""
    self.field_mapping = mapping or {}

def set_naming_rule(self, rule: dict):
    """设置文件命名规则"""
    if rule:
        self.naming_rule.update(rule)
```

| 代码 | 解释 |
|------|------|
| `self.field_mapping = mapping or {}` | 设置字段映射，如果 mapping 为空就用空字典 |
| `self.naming_rule.update(rule)` | 更新命名规则，保留默认值 |

---

## 格式化值的方法

```python
def _format_value(self, value):
    """格式化值，特别处理日期类型"""
    if pd.isna(value):
        return ''
    elif isinstance(value, (datetime, pd.Timestamp)):
        # 如果是datetime类型，只返回日期部分
        return value.strftime('%Y-%m-%d')
    elif isinstance(value, date):
        # 如果是date类型，直接格式化
        return value.strftime('%Y-%m-%d')
    else:
        # 其他类型转换为字符串
        return str(value)
```

| 代码 | 解释 |
|------|------|
| `pd.isna(value)` | 检查值是否为空（NaN）|
| `isinstance(value, datetime)` | 检查值是否是日期时间类型 |
| `value.strftime('%Y-%m-%d')` | 把日期格式化为 "2026-04-14" |
| `str(value)` | 其他类型转成字符串 |

**什么是 isinstance？**
- 就像问"这个是什么类型的？"
- `isinstance(5, int)` → True（5 是整数）
- `isinstance("hello", int)` → False（"hello" 不是整数）

---

## 应用字体格式

```python
def _apply_font(self, doc: Document):
    """统一文档字体为宋体 小五"""
    def _set_runs(paragraph):
        for run in paragraph.runs:
            run.font.name = '宋体'
            # East Asia 字体需要额外设置
            try:
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            except Exception:
                pass
            run.font.size = Pt(9)

    for paragraph in doc.paragraphs:
        _set_runs(paragraph)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _set_runs(paragraph)
```

| 代码 | 解释 |
|------|------|
| `run.font.name = '宋体'` | 设置英文字体为宋体 |
| `run._element.rPr.rFonts.set(qn(...), ...)` | 设置中文字体 |
| `run.font.size = Pt(9)` | 设置字号为 9 磅（小五）|
| `for paragraph in doc.paragraphs:` | 遍历文档所有段落 |

**什么是 run？**
- Word 中的 `run` 是一段具有相同格式的文本
- 同一段落里可以有多个不同的 run（不同字体、不同大小）

---

## 生成文件名的方法

```python
def _generate_file_name(self, row: pd.Series) -> str:
    """根据命名规则生成文件名（不含扩展名）"""
    pattern = self.naming_rule.get('pattern')

    if pattern:
        # 查找所有 [column] 占位符
        placeholders = re.findall(r'\[(.*?)\]', pattern)
        
        file_name = pattern
        for col in placeholders:
            raw_value = row.get(col, '')
            formatted_value = self._format_value(raw_value)
            file_name = file_name.replace(f'[{col}]', formatted_value)
        
        base_name = file_name

    else:
        # 回退到旧逻辑或一个默认值
        raw_value = row.get('姓名', 'unknown_file')
        base_name = self._format_value(raw_value)

    # 清理非法路径字符
    return re.sub(r'[\\/:*?"<>|]', '_', base_name)
```

| 代码 | 解释 |
|------|------|
| `self.naming_rule.get('pattern')` | 获取命名规则中的模板 |
| `re.findall(r'\[(.*?)\]', pattern)` | 从模板中找出所有 `[xxx]` |
| `row.get(col, '')` | 从 Excel 行数据中获取指定列的值 |
| `file_name.replace(...)` | 把 `[xxx]` 替换成实际值 |
| `re.sub(r'[\\/:*?"<>|]', '_', ...)` | 把文件名中的非法字符替换成 `_` |

**哪些字符是非法的？**
- `\ / : * ? " < > |` 这些字符不能用在文件名里
- 会被替换成 `_`

---

## 生成单个文档（核心方法）

```python
def generate_single_document(self, row: pd.Series, output_root_dir: str) -> tuple:
    """生成单个文档（线程安全版本）"""
    try:
        # 从模板创建新文档
        doc = Document(self.template_path)
```

| 代码 | 解释 |
|------|------|
| `-> tuple` | 返回一个元组 (成功与否, 结果/错误信息) |
| `Document(self.template_path)` | 复制模板创建新文档 |

---

## 构造有效映射

```python
# 构造占位符->Excel列的有效映射（优先用户配置）
effective_mapping = {p: self.field_mapping.get(p, p) for p in self.placeholders.keys()}
```

| 代码 | 解释 |
|------|------|
| `self.placeholders.keys()` | 所有找到的占位符 |
| `self.field_mapping.get(p, p)` | 如果用户配置了映射就用配置的，否则用同名 |
| `{p: ... for p in ...}` | 字典推导式，生成新字典 |

---

## 第一轮替换：run 内部替换

```python
def replace_in_paragraph(paragraph):
    replaced_any = False
    # 第一轮：run 内部替换
    for run in paragraph.runs:
        original = run.text
        new_text = original
        for placeholder_key, col_name in effective_mapping.items():
            token = f'[{placeholder_key}]'
            if token in new_text:
                raw_value = row.get(col_name, '')
                formatted_value = self._format_value(raw_value)
                new_text = new_text.replace(token, formatted_value)
        if new_text != original:
            run.text = new_text
            replaced_any = True
```

| 代码 | 解释 |
|------|------|
| `for run in paragraph.runs:` | 遍历段落的每个 run |
| `original = run.text` | 记录原始文本 |
| `token = f'[{placeholder_key}]'` | 构造占位符字符串 `[xxx]` |
| `if token in new_text:` | 如果文本包含这个占位符 |
| `row.get(col_name, '')` | 从 Excel 行获取值 |
| `run.text = new_text` | 替换 run 的文本 |

**为什么要分两轮？**
- 有些占位符整个在一个 run 里 → 第一轮就能替换
- 有些占位符被拆分到多个 run → 需要第二轮处理

---

## 第二轮替换：跨 run 的占位符

```python
# 第二轮：处理跨 run 的占位符（合并到首个 run）
if not replaced_any and paragraph.runs:
    combined = ''.join(r.text for r in paragraph.runs)
    original_combined = combined
    for placeholder_key, col_name in effective_mapping.items():
        token = f'[{placeholder_key}]'
        if token in combined:
            raw_value = row.get(col_name, '')
            formatted_value = self._format_value(raw_value)
            combined = combined.replace(token, formatted_value)
    if combined != original_combined:
        paragraph.runs[0].text = combined
        for r in paragraph.runs[1:]:
            r.text = ''
        replaced_any = True
return replaced_any
```

| 代码 | 解释 |
|------|------|
| `''.join(r.text for r in paragraph.runs)` | 把所有 run 的文本合并 |
| `paragraph.runs[0].text = combined` | 替换第一个 run 的文本 |
| `for r in paragraph.runs[1:]: r.text = ''` | 清空其他 run 的文本 |

---

## 处理表格和正文段落

```python
# 先处理表格中的段落
for table in doc.tables:
    for i, table_row in enumerate(table.rows):
        for j, cell in enumerate(table_row.cells):
            for paragraph in cell.paragraphs:
                replace_in_paragraph(paragraph)

# 再处理文档中不在表格内的段落（以防模板里存在）
for paragraph in doc.paragraphs:
    replace_in_paragraph(paragraph)
```

| 代码 | 解释 |
|------|------|
| `doc.tables` | 文档中的所有表格 |
| `cell.paragraphs` | 单元格中的所有段落 |
| `doc.paragraphs` | 文档正文的所有段落 |

---

## 保存生成的文档

```python
# 生成文件名
file_base_name = self._generate_file_name(row)
output_file = os.path.join(output_root_dir, f"{file_base_name}.docx")

# 保存文档
doc.save(output_file)
return True, output_file

except Exception as e:
    return False, str(e)
```

| 代码 | 解释 |
|------|------|
| `self._generate_file_name(row)` | 生成文件名 |
| `os.path.join(output_root_dir, ...)` | 拼接完整路径 |
| `doc.save(output_file)` | 保存 Word 文档 |
| `return True, output_file` | 返回成功和文件路径 |
| `return False, str(e)` | 返回失败和错误信息 |

---

## 批量生成文档方法

```python
def generate_documents(self, df: pd.DataFrame, output_root_dir: str,
                       mapping: dict = None, naming_rule: dict = None,
                       progress_callback=None) -> None:
    """根据Excel的DataFrame生成Word文档（兼容性方法）"""
    if not self.template_path:
        raise ValueError("请先加载模板文件")

    # 应用外部映射/命名规则（如提供）
    if mapping is not None:
        self.set_mapping(mapping)
    if naming_rule is not None:
        self.set_naming_rule(naming_rule)
        
    total_rows = len(df)
    # 为每一行数据生成文档
    for index, (row_idx, row) in enumerate(df.iterrows()):
        success, result = self.generate_single_document(row, output_root_dir)
        if success:
            self.logger.info(f"已生成文档: {result}")
        else:
            self.logger.error(f"生成单个文档失败: {result}")

        if progress_callback:
            progress_callback.emit(int((index + 1) * 100 / total_rows))
```

| 代码 | 解释 |
|------|------|
| `df: pd.DataFrame` | Excel 数据（多行多列的表格）|
| `progress_callback=None` | 进度回调函数（用于显示进度条）|
| `raise ValueError(...)` | 抛出错误，终止程序 |
| `df.iterrows()` | 遍历 Excel 的每一行 |
| `self.logger.info(...)` | 记录信息日志 |
| `progress_callback.emit(...)` | 发送进度更新信号 |

---

## 总结：整体流程图

```
1. 加载模板 (load_template)
   ↓
2. 查找所有占位符 (_find_placeholders)
   ↓
3. 读取 Excel 数据 (pandas)
   ↓
4. 对每一行数据：
   a. 复制模板
   b. 替换占位符为实际值
   c. 生成文件名
   d. 保存文档
```

---

## 常见问题

**Q: 什么是占位符？**
A: 占位符就是 `[xxx]` 这样的标记，代表这里要被替换成真实数据。例如 `[姓名]` 会被替换成 "张三"。

**Q: 为什么要用字典存占位符信息？**
A: 方便后续替换时知道这个占位符在哪个位置（表格、单元格、段落等）。

**Q: 什么是线程安全？**
A: 这个方法不依赖共享的实例变量，可以同时运行多个而不互相干扰，适合批量处理。

---

*文档生成时间: 2026-04-14*
