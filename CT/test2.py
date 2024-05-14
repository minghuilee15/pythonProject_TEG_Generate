import tkinter as tk

def get_user_input():
    # 创建一个字典来存储输入的变量及其对应的值
    user_input = {
        "Typical_CT_Size": Typical_CT_Size_entry.get(),
        "Enclosure1_Size": Enclosure1_Size_entry.get(),
        "Enclosure2_Size_4T_Chain": Enclosure2_Size_4T_Chain_entry.get(),
        "Enclosure2_Size_4T_Chain_End": Enclosure2_Size_4T_Chain_End_entry.get(),
        "AA_E1_1": AA_E1_1_entry.get(),
        "AA_E1_2": AA_E1_2_entry.get(),
        "AA_E2_1": AA_E2_1_entry.get(),
        "AA_E2_2": AA_E2_2_entry.get(),
        "M1_E1_1": M1_E1_1_entry.get(),
        "M1_E2_1": M1_E2_1_entry.get(),
        "M1_E2_2": M1_E2_2_entry.get(),
        "M1_E2_3": M1_E2_3_entry.get(),
        "Poly_E1_1": Poly_E1_1_entry.get(),
        "Poly_E1_2": Poly_E1_2_entry.get(),
        "Poly_E2_1": Poly_E2_1_entry.get(),
        "Poly_E2_2": Poly_E2_2_entry.get()
    }

    # # 打印输入的变量及其对应的值
    # for var_name, var_value in user_input.items():
    #     print(var_name + ":", var_value)

# 创建主窗口
root = tk.Tk()

# 添加标签和输入框
Typical_CT_Size_label = tk.Label(root, text="Typical_CT_Size:")
Typical_CT_Size_label.grid(row=0, column=0)
Typical_CT_Size_entry = tk.Entry(root)
Typical_CT_Size_entry.grid(row=0, column=1)

Enclosure1_Size_label = tk.Label(root, text="Enclosure1_Size:")
Enclosure1_Size_label.grid(row=1, column=0)
Enclosure1_Size_entry = tk.Entry(root)
Enclosure1_Size_entry.grid(row=1, column=1)

Enclosure2_Size_4T_Chain_label = tk.Label(root, text="Enclosure2_Size_4T_Chain:")
Enclosure2_Size_4T_Chain_label.grid(row=2, column=0)
Enclosure2_Size_4T_Chain_entry = tk.Entry(root)
Enclosure2_Size_4T_Chain_entry.grid(row=2, column=1)

Enclosure2_Size_4T_Chain_End_label = tk.Label(root, text="Enclosure2_Size_4T_Chain_End:")
Enclosure2_Size_4T_Chain_End_label.grid(row=3, column=0)
Enclosure2_Size_4T_Chain_End_entry = tk.Entry(root)
Enclosure2_Size_4T_Chain_End_entry.grid(row=3, column=1)

AA_E1_1_label = tk.Label(root, text="AA_E1_1:")
AA_E1_1_label.grid(row=4, column=0)
AA_E1_1_entry = tk.Entry(root)
AA_E1_1_entry.grid(row=4, column=1)

AA_E1_2_label = tk.Label(root, text="AA_E1_2:")
AA_E1_2_label.grid(row=5, column=0)
AA_E1_2_entry = tk.Entry(root)
AA_E1_2_entry.grid(row=5, column=1)

AA_E2_1_label = tk.Label(root, text="AA_E2_1:")
AA_E2_1_label.grid(row=6, column=0)
AA_E2_1_entry = tk.Entry(root)
AA_E2_1_entry.grid(row=6, column=1)

AA_E2_2_label = tk.Label(root, text="AA_E2_2:")
AA_E2_2_label.grid(row=7, column=0)
AA_E2_2_entry = tk.Entry(root)
AA_E2_2_entry.grid(row=7, column=1)

M1_E1_1_label = tk.Label(root, text="M1_E1_1:")
M1_E1_1_label.grid(row=8, column=0)
M1_E1_1_entry = tk.Entry(root)
M1_E1_1_entry.grid(row=8, column=1)

M1_E2_1_label = tk.Label(root, text="M1_E2_1:")
M1_E2_1_label.grid(row=9, column=0)
M1_E2_1_entry = tk.Entry(root)
M1_E2_1_entry.grid(row=9, column=1)

M1_E2_2_label = tk.Label(root, text="M1_E2_2:")
M1_E2_2_label.grid(row=10, column=0)
M1_E2_2_entry = tk.Entry(root)
M1_E2_2_entry.grid(row=10, column=1)

M1_E2_3_label = tk.Label(root, text="M1_E2_3:")
M1_E2_3_label.grid(row=11, column=0)
M1_E2_3_entry = tk.Entry(root)
M1_E2_3_entry.grid(row=11, column=1)

Poly_E1_1_label = tk.Label(root, text="Poly_E1_1:")
Poly_E1_1_label.grid(row=12, column=0)
Poly_E1_1_entry = tk.Entry(root)
Poly_E1_1_entry.grid(row=12, column=1)

Poly_E1_2_label = tk.Label(root, text="Poly_E1_2:")
Poly_E1_2_label.grid(row=13, column=0)
Poly_E1_2_entry = tk.Entry(root)
Poly_E1_2_entry.grid(row=13, column=1)

Poly_E2_1_label = tk.Label(root, text="Poly_E2_1:")
Poly_E2_1_label.grid(row=14, column=0)
Poly_E2_1_entry = tk.Entry(root)
Poly_E2_1_entry.grid(row=14, column=1)

Poly_E2_2_label = tk.Label(root, text="Poly_E2_2:")
Poly_E2_2_label.grid(row=15, column=0)
Poly_E2_2_entry = tk.Entry(root)
Poly_E2_2_entry.grid(row=15, column=1)

# 创建按钮用于触发获取用户输入的函数
get_input_button = tk.Button(root, text="获取用户输入", command=get_user_input)
get_input_button.grid(row=16, column=0, columnspan=2)

root.mainloop()
