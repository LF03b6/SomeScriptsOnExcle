import pandas as pd

def strip_columns(file_path, output_path):
    # 读取 Excel 表格
    df = pd.read_excel(file_path)

    # 列范围 A 到 K 对应的索引是 0 到 10
    columns_to_strip = df.columns[:11]

    # 去除每列中的开头和结尾空格，保留数字格式
    for col in columns_to_strip:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

    # 将处理后的表格保存到新的文件
    df.to_excel(output_path, index=False)

# 使用示例
input_file = "input.xlsx"  # 输入文件路径
output_file = "output.xlsx"  # 输出文件路径
strip_columns(input_file, output_file)
