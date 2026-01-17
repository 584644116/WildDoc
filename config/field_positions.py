FIELD_POSITIONS = {
    "申请考核岗位": {
        "table_index": 0,  # 第一个表格
        "positions": {
            "姓名": {"row": 0, "column": 1},      # 第一行第2个单元格
            "性别": {"row": 0, "column": 4},      # 第一行第5个单元格
            "出生年月": {"row": 0, "column": 8},  # 第一行第9个单元格
            "入职日期": {"row": 1, "column": 1},  # 第二行第2个单元格
            "文化程度": {"row": 1, "column": 4},  # 第二行第5个单元格
            "专业": {"row": 1, "column": 8},      # 第二行第9个单元格
            "岗位": {"row": 2, "column": 1},      # 第三行第2个单元格
            "职称": {"row": 2, "column": 4},      # 第三行第5个单元格
            "现从事专业及年限": {"row": 2, "column": 8},  # 第三行第9个单元格
        }
    }
}

# 对于倒数的行，我们需要在运行时计算具体行号
def get_dynamic_positions(table):
    """获取需要动态计算的行位置"""
    total_rows = len(table.rows)
    return {
        "人员现已具备的条件": {"row": total_rows - 3, "column": 1},  # 倒数第三行第二个单元格
        "考核意见": {"row": total_rows - 2, "column": 1},  # 倒数第二行第二个单元格
    } 