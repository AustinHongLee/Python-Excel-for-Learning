import openpyxl

# 創建一個新的工作簿
wb = openpyxl.Workbook()

# 保存工作簿到一個文件
wb.save("new_blank_workbook.xlsx")

# 如果你希望自動打開該工作簿，可以使用os庫
import os
os.startfile("new_blank_workbook.xlsx")
