import openpyxl
import os

print("創建一個新的工作簿...")
# 創建一個新的工作簿
wb = openpyxl.Workbook()
ws = wb.active  # 獲取當前工作表

print("向工作簿中添加數據...")
# 將值寫入 A1 儲存格
ws["A1"].value = "C150x50x6t"

# 將公式寫入 B1 儲存格
ws["B1"].value = "=LEFT(A1,1)"

print("保存工作簿到一個文件...")
# 保存工作簿到一個文件
wb.save("new_blank_workbook.xlsx")
print(f"保存工作簿到 {os.getcwd()}\\new_blank_workbook.xlsx...")

print("打開新創建的工作簿...")
# 如果你希望自動打開該工作簿，可以使用os庫
os.startfile("new_blank_workbook.xlsx")
