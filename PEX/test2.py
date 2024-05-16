import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# 定义因素和水平
factors = {
    'Type': ['Three layers', 'Two layers'],
    'Bot-Layer': ['NW STI', 'N+AA', 'N+Poly on AA', 'N+Poly on STI', 'M1', 'M2', 'M4', 'M5'],
    'Mid-Layer': ['M1', 'M2', 'M3', 'M4', 'M5', 'N+Poly'],
    'Top-Layer': ['M1', 'M2', 'M3', 'M4', 'M5', 'MT']
}

# 生成三层和两层的数据
three_layers_data = [
    ['Three layers', 'NW STI', 'M1', 'M2'],
    ['Three layers', 'NW STI', 'N+Poly', 'M1'],
    ['Three layers', 'N+AA', 'M1', 'M2'],
    ['Three layers', 'N+Poly on AA', 'M1', 'M2'],
    ['Three layers', 'N+Poly on STI', 'M1', 'M2'],
    ['Three layers', 'M1', 'M2', 'M3'],
    ['Three layers', 'M1', 'M2', 'M4'],
    ['Three layers', 'M1', 'M3', 'M4'],
    ['Three layers', 'M2', 'M3', 'M4'],
    ['Three layers', 'M2', 'M3', 'M5'],
    ['Three layers', 'M4', 'M5', 'MT']
]

two_layers_data = [
    ['Two layers', 'NW STI', 'M1', '/'],
    ['Two layers', 'NW STI', 'N+Poly', '/'],
    ['Two layers', 'N+AA', 'M1', '/'],
    ['Two layers', 'N+Poly on AA', 'M1', '/'],
    ['Two layers', 'N+Poly on STI', 'M1', '/'],
    ['Two layers', 'M1', 'M2', '/'],
    ['Two layers', 'M1', 'M3', '/'],
    ['Two layers', 'M2', 'M3', '/'],
    ['Two layers', 'M2', 'M4', '/'],
    ['Two layers', 'M5', 'MT', '/']
]

# 合并数据
data = three_layers_data + two_layers_data

# 创建DataFrame
df = pd.DataFrame(data, columns=['Type', 'Bot-Layer', 'Mid-Layer', 'Top-Layer'])

# 定义 W 和 S 的水平
W_levels = ['0.13', '0.16', '0.2', '0.42', '10']
S_levels = ['0.18', '0.17', '0.2', '0.42', '0.25', '0.5']

# 生成所有可能的组合
expanded_data = []
for row in data:
    for w_level in W_levels:
        for s_level in S_levels:
            expanded_data.append(row + [w_level, s_level])

# 创建扩展后的DataFrame
expanded_df = pd.DataFrame(expanded_data, columns=['Type', 'Bot-Layer', 'Mid-Layer', 'Top-Layer', 'W', 'S'])

# 定义文件名
file_name = 'expanded_layer_combinations_full_factorial.xlsx'

# 保存到Excel文件
expanded_df.to_excel(file_name, index=False)

# 加载 Excel 文件
wb = load_workbook(file_name)

# 选择第一个工作表
ws = wb.active

# 设置字体和对齐方式
font = Font(name='Times New Roman')
alignment = Alignment(horizontal='center', vertical='center')

# 设置首行字体加粗
for cell in ws[1]:
    cell.font = Font(name='Times New Roman', bold=True)

# 遍历所有单元格，设置字体、对齐方式和行高
for row in ws.iter_rows(min_row=2, min_col=1, max_col=ws.max_column):
    for cell in row:
        cell.font = font
        cell.alignment = alignment
    # 设置行高为 20
    ws.row_dimensions[row[0].row].height = 20

# 自动调整列宽以适应内容
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter  # 获取列名
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2) * 1.2
    ws.column_dimensions[column].width = adjusted_width

# 保存修改后的 Excel 文件
wb.save(file_name)

# 自动打开Excel文件
os.startfile(file_name)

print(f"Excel表格已生成并保存为 '{file_name}' 并已自动打开")
