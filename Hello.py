import openpyxl
import os
import tkinter as tk
from tkinter import messagebox, StringVar, Text

# 定義添加選擇到文本框的函數
def add_to_textbox():
    current_value = selected_option.get()
    textbox.insert(tk.END, current_value + "\n")

# 定義保存到Excel的函數
def save_to_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    lines = textbox.get(1.0, tk.END).split("\n")
    for i, line in enumerate(lines, 1):
        if line:  # 僅當line非空時寫入Excel
            ws[f"A{i}"].value = line
    wb.save("saved_values.xlsx")

window = tk.Tk()
window.title("軟體視窗")
window.geometry("280x400")

# 下拉選項
options_part1 = ["C015", "C016", "C017", "C019", "C021"]
options_part2 = ["C022", "C023", "C040", "C041", "C046"]
options_part3 = ["C1021", "C1022", "C1121", "C1122", "C1301"]
options_part4 = ["H054", "I080", "I1080", "P085", "P086"]
options_part5 = ["P087", "P093", "P094", "W114", "W115"]
options_part6 = ["W117", "W122", "W124", "W127", "W131"]
options_part7 = ["W132", "W140", "W141"]

options = options_part1 + options_part2 + options_part3 + options_part4 + options_part5 + options_part6 + options_part7
 


selected_option = StringVar(window)
selected_option.set(options[0])
dropdown = tk.OptionMenu(window, selected_option, *options)
dropdown.pack(pady=10)

# 按鈕用來添加選擇到文本框
add_button = tk.Button(window, text="添加到文本框", command=add_to_textbox)
add_button.pack(pady=10)

# 文本框用於顯示添加的值
textbox = Text(window, height=10, width=30)
textbox.pack(pady=10)

# 按鈕用來保存文本框的值到Excel
save_button = tk.Button(window, text="保存到Excel", command=save_to_excel)
save_button.pack(pady=10)

window.mainloop()
