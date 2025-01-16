import openpyxl
from datetime import datetime

def convert_date_format(file_path, output_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    for row in sheet.iter_rows(min_row=1, min_col=10, max_col=10):
        for cell in row:
            if isinstance(cell.value, str):
                
                try:
                    # 将原格式字符串转为datetime对象
                    dt = datetime.strptime(cell.value, '%Y/%m/%d')
                    # 再格式化为新的datetime字符串
                    new_dt = dt.strftime('%Y-%m-%d 00:00:00')
                    cell.value = new_dt
                except ValueError:
                    pass
            elif isinstance(cell.value, datetime):
                # 如果已经是datetime类型，直接格式化
                new_dt = cell.value.strftime('%Y-%m-%d 00:00:00')
                cell.value = new_dt

    workbook.save(output_path)


file_path = 'input.xlsx'
output_path = 'output.xlsx'
convert_date_format(file_path, output_path)