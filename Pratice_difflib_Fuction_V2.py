import openpyxl
import difflib

def exact_match(cell_value, keywords):
    """進行明確匹配"""
    for keyword in keywords:
        print(f"正在進行明確匹配，關鍵字為: {keyword}")
        if keyword in cell_value:
            return keyword
    return None


def fuzzy_match(cell_value, keywords, cutoff=0.8):
    """進行模糊匹配"""
    for keyword in keywords:
        print(f"正在進行模糊匹配，關鍵字為: {keyword}，相似度閾值為: {cutoff}")
    matches = difflib.get_close_matches(cell_value, keywords, n=1, cutoff=cutoff)
    if matches:
        return matches[0]
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
keywords = ["Pipe", "Valve", "Bolt", "Nut", "Elbow", "Flange"]

print("開始迭代工作表中的3行...")
for row in ws.iter_rows(min_row=1, max_row=3, min_col=1, max_col=1):
    cell = row[0]
    print(f"正在檢查第{cell.row}行的單元格，其值為: {cell.value}")

    matched_keyword = exact_match(cell.value, keywords) or fuzzy_match(cell.value, keywords)

    if matched_keyword:
        print(f"在第{cell.row}行的單元格中找到了匹配項 {matched_keyword}。正在將值寫入B列{cell.row}行。")
        ws.cell(row=cell.row, column=2).value = matched_keyword
    else:
        print(f"在第{cell.row}行的單元格中沒有找到匹配項。")

# 保存Excel文件
print("正在保存Excel文件...")
wb.save(file_path)
print("Excel文件保存成功。")