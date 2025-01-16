import openpyxl

def find_missing_rows(file_a, file_b, output_file):
    # 打开表格A和表格B
    workbook_a = openpyxl.load_workbook(file_a)
    sheet_a = workbook_a.active  # 获取表格A的第一个工作表

    workbook_b = openpyxl.load_workbook(file_b)
    sheet_b = workbook_b.active  # 获取表格B的第一个工作表

    # 获取表格A的第一列数据
    values_in_a = set(cell.value for cell in sheet_a['A'] if cell.value is not None)

    # 遍历表格B的第A列，找出不在表格A中的值
    missing_rows = []
    for row_idx, cell in enumerate(sheet_b['A'], start=1):
        if cell.value is not None and cell.value not in values_in_a:
            missing_rows.append(row_idx)

    # 输出结果到文件或控制台
    if missing_rows:
        with open(output_file, 'w') as output:
            output.write("以下是表格B中A列的行索引，它们的值在表格A中未找到：\n")
            output.write(", ".join(map(str, missing_rows)))
        print(f"检查完成，未找到的行索引已输出到 {output_file}")
    else:
        print("检查完成，表格B的A列所有值都存在于表格A的第一列中。")

# 使用脚本
file_a = 'a.xlsx'  # 表格A路径
file_b = 'b.xlsx'  # 表格B路径
output_file = 'missing_rows.txt'  # 输出结果路径

find_missing_rows(file_a, file_b, output_file)

#find b'a in a'a