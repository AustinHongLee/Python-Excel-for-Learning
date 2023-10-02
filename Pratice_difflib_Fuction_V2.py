import openpyxl
import difflib
import sys
import re

sys.stdout = open('output.txt', 'w')
def exact_match(cell_value, keywords):
    """進行明確匹配"""
    if cell_value is None:
        return None

    for keyword in keywords:
        print(f"正在進行明確匹配，關鍵字為: {keyword}")
        if cell_value.startswith(keyword):
            return keyword
    return None



def fuzzy_match(cell_value, keywords):
    """進行模糊匹配"""
    if cell_value is None:
        return None
    
    max_similarity = 0.8
    matched_keyword = None
    for keyword in keywords:
        similarity = difflib.SequenceMatcher(autojunk=True, a=keyword, b=cell_value).ratio()
        print(f"正在進行模糊匹配，關鍵字為: {keyword}，相似度為: {similarity}")
        if similarity > max_similarity:
            max_similarity = similarity
            matched_keyword = keyword
    if matched_keyword:
        return matched_keyword
    return None

# 載入Excel文件
print("正在嘗試加載Excel文件...")
file_path = (
    r'C:\Users\a0976\OneDrive\文件\GitHub\Python Excel for Learning'
    r'\example.xlsx'
)

wb = openpyxl.load_workbook(file_path)
ws = wb.active
print("Excel文件成功加載。")

# 模糊匹配的關鍵字列表
keywords = ["Pipe", 
            "Bolt", "Nut", "Elbow","Gasket","Nipple","Union","Coupling","Plug","Cap", 
            "Weld Neck Flange","Socket Weld Flange","Slip-On Flange","Lap Joint Flange",
            "Threaded Flange","Blind Flange",
            "Butterfly Valve","Ball Valve","Gate Valve","Dual Plate Check Valve",
            "Globe Valve","Needle Valve","Check Valve","Butterfly Valve",
            "Support","Stub End","Swage Nipple","Spectacle Blind","Orifice Flange",
            "Sockolet", "Weldolet", "Reducer", "Tee"]



print("開始迭代工作表中的3行...")
for row in ws.iter_rows(min_row=1, max_row=30, min_col=1, max_col=1):
    cell = row[0]
    print(f"正在檢查第{cell.row}行的單元格，其值為: {cell.value}")

    matched_keyword = exact_match(cell.value, keywords) or fuzzy_match(cell.value, keywords)

    if matched_keyword:
        print(f"在第{cell.row}行的單元格中找到了匹配項 {matched_keyword}。正在將值寫入B列{cell.row}行。")
        ws.cell(row=cell.row, column=2).value = matched_keyword
    else:
        ws.cell(row=cell.row, column=2).value = "-"
        print(f"在第{cell.row}行的單元格中沒有找到匹配項。")

# 保存Excel文件
print("正在保存Excel文件...")
wb.save(file_path)
print("Excel文件保存成功。")
