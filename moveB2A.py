
import pandas as pd
import os



# 读取 Excel 文件
file_path = 'input.xlsx'  # 替换为你的文件路径
df = pd.read_excel(file_path)

# 打印实际列名，确认第一列和第二列
print("Excel 文件的列名:", df.columns)

# 遍历第二列（B列），检查是否为空
for i in range(len(df)):
    if pd.notna(df.iloc[i, 1]) :  # 如果第二列的单元格不为空
        # 合并第一列和第二列的内容，赋值回第一列
        df.iloc[i, 0] = str(df.iloc[i, 0]) + str(df.iloc[i, 1])

# 指定保存路径
output_dir = 'output'  # 替换为你希望保存的目录路径
output_file = os.path.join(output_dir, 'modified_excel_file.xlsx')

# 如果输出目录不存在，则创建它
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# 将修改后的 DataFrame 保存回 Excel 文件
df.to_excel(output_file, index=False)

print(f"第一列和第二列合并并保存至第一列，文件已保存至 {output_file}")
