import openpyxl
import re

def clean_column_a(file_path, output_path):
    # 打开 Excel 文件
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active  # 获取第一个工作表

    # 正则表达式用于匹配非字母数字和空格
    pattern = re.compile(r'[^a-zA-Z0-9\u4e00-\u9fa5]')

    # 遍历 A 列的所有单元格，从第2行开始（假设第1行是表头）
    for row_idx, cell in enumerate(sheet['A'], start=1):
        if cell.value and isinstance(cell.value, str):
            # 使用正则表达式删除特殊字符和空格
            cleaned_value = pattern.sub('', cell.value)
            if cell.value != cleaned_value:
                print(f"清理第 {row_idx} 行的值: '{cell.value}' -> '{cleaned_value}'")
            cell.value = cleaned_value

    # 保存修改后的文件
    workbook.save(output_path)
    print(f"清理完成，保存为 {output_path}")

# 使用脚本
file_path = 'input.xlsx'  # 输入文件路径
output_path = 'output.xlsx'  # 输出文件路径

clean_column_a(file_path, output_path)
