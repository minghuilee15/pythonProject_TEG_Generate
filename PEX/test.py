import tkinter as tk
from tkinter import ttk
import pandas as pd
from itertools import product
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
import os

variables = {
    'Type': ['Two_layers', 'Three_layers'],
    'Bot-Layer': ['NW STI', 'N+AA', 'N+Poly on AA', 'N+Poly on STI', 'M1', 'M2', 'M4'],
    'Mid-Layer': ['M1', 'M2', 'M3', 'M4', 'M5', 'N+Poly'],
    'Top-Layer': ['M1', 'M2', 'M3', 'M4', 'M5', 'MT', '/'],
    'W': ['0.13', '0.16', '0.2', '0.42', '10', '0.25', '0.5'],
    'S': ['0.5']
}

# 对current_value = 0.13, 0.16, 0.2, 0.42分别生成2倍数递增的值，并添加到"W"键的值中
for start_value in [0.13, 0.16, 0.2, 0.42]:
    current_value = round(start_value * 2, 2)  # 将浮点数四舍五入到两位小数
    variables['W'].append(str(current_value))
    current_value = round(start_value * 3, 2)  # 将浮点数四舍五入到两位小数
    variables['W'].append(str(current_value))

# 添加标题项到组合列表中
all_combinations = [list(variables.keys())]
all_combinations.extend(list(product(*variables.values())))

# 将组合转换为DataFrame
design_df = pd.DataFrame(all_combinations[1:], columns=all_combinations[0])

# 在第一列后增加两列Module和Lable
design_df.insert(0, 'Module', '')
design_df.insert(1, 'Label', '')

# 在最后一列后面增加5列Sens1，Force，COM，Sens2，Description
design_df['NF'] = ''
design_df['L'] = ''
design_df['Row(Y)'] = ''
design_df['Col(X)'] = ''
design_df['Mid1'] = ''
design_df['Bottom'] = ''
design_df['Top'] = ''
design_df['Mid2'] = ''

# #定义Typical_CT尺寸
# Typical_CT_Size = '0.16'
#
# #拉偏CT_Size时候，控制Enclosure1_Size大小不变
# Enclosure1_Size = '0.06'
#
# # 拉偏CT_Size时候，控制Enclosure2_Size大小不变
# Enclosure2_Size_4T_Chain = '0'
# Enclosure2_Size_4T_Chain_End = '0.05'
#
# # 设定Mid-Layer为Poly时，W的尺寸
# Poly_W1 = '0.13'
# Poly_W2 = '0.26'
# Poly_W3 = '0.39'
# Poly_W4 = '6'
# Poly_W5 = '100'
#
# # 设定Mid-Layer为Poly时，S的尺寸
# Poly_S1 = '0.18'
# Poly_S2 = '0.36'
# Poly_S3 = '0.54'
# Poly_S4 = '0.6'
# Poly_S5 = '5'
#
# # 设定Mid-Layer为Poly时，W的尺寸
# M1_W1 = '0.16'
# M1_W2 = '0.32'
# M1_W3 = '0.48'
# M1_W4 = '6'
# M1_W5 = '100'
#
# # 设定Mid-Layer为Poly时，S的尺寸
# M1_S1 = '0.17'
# M1_S2 = '0.34'
# M1_S3 = '0.51'
# M1_S4 = '0.6'
# M1_S5 = '5'

#
# # 拉偏M1 enc CT时候，控制Enclosure1_Size为0.06，Enclosure2_Size为-0.01(4T,Chain)，0和0.03(4T_End,Chain_End)
# M1_E1_1 = '0.06'
# M1_E2_1 = '-0.01' #4T,Chain
# M1_E2_2 = '0'  #4T_End,Chain_End
# M1_E2_3 = '0.03'  #4T_End,Chain_End
#
# # 拉偏Poly enc CT时候，控制Enclosure1_Size为0和0.03，Enclosure2_Size为0(4T,Chain)和0.05(4T_End,Chain_End)
# Poly_E1_1 = '0'
# Poly_E1_2 = '0.03'
# Poly_E2_1 = '0'#4T,Chain
# Poly_E2_2 = '0.05'  #4T_End,Chain_End
#
# 删除不满足条件的行
design_df = design_df[
    ~(
            # Two_Layers时，Top-Layer为/
            ((design_df['Type'] == 'Two_layers') & (design_df['Top-Layer'] != '/')) |
            # Three_Layers时，Top-Layer不为/
            ((design_df['Type'] == 'Three_layers') & (design_df['Top-Layer'] == '/'))

            # # Mid-Layer为N+Poly时，W取Poly的最小宽度及其整数倍
            # (
            #         (design_df['Type'] == 'Two_Layers') &
            #         (design_df['Mid-Layer'] == 'N+Poly') &
            #         (
            #                 (~design_df['W'].isin([Poly_W1, Poly_W2, Poly_W3, Poly_W4, Poly_W5])) |
            #                 (~design_df['S'].isin([Poly_S1, Poly_S2, Poly_S3, Poly_S4, Poly_S5]))
            #         )
            # )
            # |
            # # Mid-Layer为M1时，W取Poly的最小宽度及其整数倍
            # (
            #         (design_df['Type'] == 'Two_Layers') &
            #         (design_df['Mid-Layer'] == 'M1') &
            #         (
            #                 (~design_df['W'].isin([M1_W1, M1_W2, M1_W3, M1_W4, M1_W5])) |
            #                 (~design_df['S'].isin([M1_S1, M1_S2, M1_S3, M1_S4, M1_S5]))
            #         )
            # )

        # # 删除Poly因素中的AA enc CT
        # ((design_df['Device3'] == 'Poly') & (design_df['Split'] == 'AA enc CT')) |
        #
        #
        # # 定义Typical_CT_Size
        # ((design_df['Split'] == 'AA enc CT') & (design_df['Size'] != Typical_CT_Size)) |
        # ((design_df['Split'] == 'M1 enc CT') & (design_df['Size'] != Typical_CT_Size)) |
        # ((design_df['Split'] == 'Poly enc CT') & (design_df['Size'] != Typical_CT_Size)) |
        #
        # # 拉偏CT_Size时候，控制Enclosure1_Size大小不变
        # ((design_df['Split'] == 'Size') & (design_df['Enclosure1'] != Enclosure1_Size)) |
        #
        # # 拉偏CT_Size时候，控制Enclosure2_Size大小不变
        # ((design_df['Type'] == '4T') & (design_df['Split'] == 'Size') & ~(design_df['Enclosure2'] == Enclosure2_Size_4T_Chain)) |
        # ((design_df['Type'] == '4T-End') & (design_df['Split'] == 'Size') & ~(design_df['Enclosure2'] == Enclosure2_Size_4T_Chain_End)) |
        # ((design_df['Type'] == 'Chain') & (design_df['Split'] == 'Size') & ~(design_df['Enclosure2'] == Enclosure2_Size_4T_Chain)) |
        # ((design_df['Type'] == 'Chain-End') & (design_df['Split'] == 'Size') & ~(design_df['Enclosure2'] == Enclosure2_Size_4T_Chain_End)) |
        #
        #
        # # AA enc CT
        #
        # ((design_df['Split'] == 'AA enc CT') & (design_df['Type'] == '4T') & (~(design_df['Enclosure1'].isin([AA_E1_1, AA_E1_2])))) |
        # ((design_df['Split'] == 'AA enc CT') & (design_df['Type'] == '4T') & (design_df['Enclosure2'] != AA_E2_1)) |
        #
        # ((design_df['Split'] == 'AA enc CT') & (design_df['Type'] == '4T-End') & (~(design_df['Enclosure1'].isin([AA_E1_1, AA_E1_2])))) |
        # ((design_df['Split'] == 'AA enc CT') & (design_df['Type'] == '4T-End') & (design_df['Enclosure2'] != AA_E2_2)) |
        #
        # ((design_df['Split'] == 'AA enc CT') & (design_df['Type'] == 'Chain') & (~(design_df['Enclosure1'].isin([AA_E1_1, AA_E1_2])))) |
        # ((design_df['Split'] == 'AA enc CT') & (design_df['Type'] == 'Chain') & (design_df['Enclosure2'] != AA_E2_1)) |
        #
        # ((design_df['Split'] == 'AA enc CT') & (design_df['Type'] == 'Chain-End') & (~(design_df['Enclosure1'].isin([AA_E1_1, AA_E1_2])))) |
        # ((design_df['Split'] == 'AA enc CT') & (design_df['Type'] == 'Chain-End') & (design_df['Enclosure2'] != AA_E2_2)) |
        #
        # # M1 enc CT
        #
        # ((design_df['Split'] == 'M1 enc CT') & (design_df['Type'] == '4T') & (design_df['Enclosure1'] != M1_E1_1)) |
        # ((design_df['Split'] == 'M1 enc CT') & (design_df['Type'] == '4T') & (design_df['Enclosure2'] != M1_E2_1)) |
        #
        # ((design_df['Split'] == 'M1 enc CT') & (design_df['Type'] == '4T-End') & ~(design_df['Enclosure1'] == M1_E1_1)) |
        # ((design_df['Split'] == 'M1 enc CT') & (design_df['Type'] == '4T-End') & (~(design_df['Enclosure2'].isin([M1_E2_2, M1_E2_3])))) |
        #
        # ((design_df['Split'] == 'M1 enc CT') & (design_df['Type'] == 'Chain') & ~(design_df['Enclosure1'] == M1_E1_1)) |
        # ((design_df['Split'] == 'M1 enc CT') & (design_df['Type'] == 'Chain') & (design_df['Enclosure2'] != M1_E2_1)) |
        #
        # ((design_df['Split'] == 'M1 enc CT') & (design_df['Type'] == 'Chain-End') & ~(design_df['Enclosure1'] == M1_E1_1)) |
        # ((design_df['Split'] == 'M1 enc CT') & (design_df['Type'] == 'Chain-End') & (~(design_df['Enclosure2'].isin([M1_E2_2, M1_E2_3])))) |
        #
        # # Poly enc CT
        #
        # ((design_df['Split'] == 'Poly enc CT') & (design_df['Type'] == '4T') & (~(design_df['Enclosure1'].isin([Poly_E1_1, Poly_E1_2])))) |
        # ((design_df['Split'] == 'Poly enc CT') & (design_df['Type'] == '4T') & (design_df['Enclosure2'] != Poly_E2_1)) |
        #
        # ((design_df['Split'] == 'Poly enc CT') & (design_df['Type'] == '4T-End') & (~(design_df['Enclosure1'].isin([Poly_E1_1, Poly_E1_2])))) |
        # ((design_df['Split'] == 'Poly enc CT') & (design_df['Type'] == '4T-End') & (design_df['Enclosure2'] != Poly_E2_2)) |
        #
        # ((design_df['Split'] == 'Poly enc CT') & (design_df['Type'] == 'Chain') & (~(design_df['Enclosure1'].isin([Poly_E1_1, Poly_E1_2])))) |
        # ((design_df['Split'] == 'Poly enc CT') & (design_df['Type'] == 'Chain') & (design_df['Enclosure2'] != Poly_E2_1)) |
        #
        # ((design_df['Split'] == 'Poly enc CT') & (design_df['Type'] == 'Chain-End') & (~(design_df['Enclosure1'].isin([Poly_E1_1, Poly_E1_2])))) |
        # ((design_df['Split'] == 'Poly enc CT') & (design_df['Type'] == 'Chain-End') & (design_df['Enclosure2'] != Poly_E2_2))

    )
]

# 在 DataFrame 中插入一列序号
design_df.insert(0, 'NO.', range(1, len(design_df) + 1))

# 保存设计到Excel文件
excel_file = "orthogonal_experiment_design_CT.xlsx"

# 创建一个新的Excel工作簿
wb = Workbook()
ws = wb.active

# 添加标题行并设置边框
for c_idx, title in enumerate(design_df.columns, 1):
    cell = ws.cell(row=1, column=c_idx, value=title)
    cell.font = Font(name='Times New Roman', bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = Border(bottom=Side(style='medium'))

# 将DataFrame写入工作表
for r_idx, r in enumerate(design_df.values.tolist(), 2):
    for c_idx, value in enumerate(r, 1):
        cell = ws.cell(row=r_idx, column=c_idx, value=value)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.font = Font(name='Times New Roman')

# 设置字体和自动调整列宽
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2) * 1.2
    ws.column_dimensions[column].width = adjusted_width

# 设置行高
for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
    for cell in row:
        ws.row_dimensions[cell.row].height = 25

# 保存Excel文件
wb.save(excel_file)
print(f"正交实验设计已保存到 {excel_file}")

# 自动打开Excel文件
os.system(f'start excel {excel_file}')
