import chardet 

def detect_encoding(file_path):
    """
    自动检测文件的编码格式
    :param file_path: 文件路径
    :return: 编码格式（如 'utf-8' 或 'gbk'）
    """
    with open(file_path, "rb") as file:
        raw_data = file.read(1000)  # 读取前1000字节，足够检测编码
        result = chardet.detect(raw_data)
        encoding = result['encoding']
        if not encoding:
            raise ValueError("无法检测文件编码")
        return encoding

def delete_lines_in_file(file_path, start_line, end_line):
    """
    删除文件中指定范围的行
    :param file_path: 文件路径
    :param start_line: 起始行（包含）
    :param end_line: 结束行（包含）
    """
    try:
        # 检测文件编码
        encoding = detect_encoding(file_path)
        print(f"检测到文件编码为: {encoding}")

        # 读取文件内容
        with open(file_path, "r", encoding=encoding) as file:
            lines = file.readlines()

        # 删除指定行
        new_lines = [line for i, line in enumerate(lines, start=1) if i < start_line or i > end_line]

        # 写回文件
        with open(file_path, "w", encoding=encoding) as file:
            file.writelines(new_lines)

        print(f"已成功删除文件 {file_path} 中的第 {start_line} 行到第 {end_line} 行")

    except Exception as e:
        print(f"操作失败: {e}")

# 示例调用
# 示例调用
file_path = "input.xlsx"  # 替换为你的文件路径
start_line = 60
end_line = 4000






