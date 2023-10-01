import openpyxl
import difflib

# 載入Excel文件
print("正在嘗試加載Excel文件...")
wb = openpyxl.load_workbook(r'C:\Users\user\OneDrive\文件\GitHub\Python Excel for Learning\example.xlsx')
ws = wb.active
print("Excel文件成功加載。")

# 模糊匹配的關鍵字列表
keywords = ["Pipe", "Valve", "Bolt", "Nut", "Elbow", "Flange"]

print(f"開始迭代工作表中的{ws.max_row}行...")
for row in ws.iter_rows(min_row=1, max_row=3, min_col=1, max_col=1):
    cell = row[0]
    print(f"正在檢查第{cell.row}行的單元格，其值為: {cell.value}")
    
    if cell.value is None:
        print(f"第{cell.row}行的單元格是空的。")
        continue
    
matches = difflib.get_close_matches(cell.value, keywords, n=1, cutoff=0.8) # 將cutoff值降低到0.5
if matches:
    print(f"在第{cell.row}行的單元格中找到了匹配項。正在將值 {matches[0]} 寫入B列{cell.row}行。")
    ws.cell(row=cell.row, column=2).value = matches[0]
else:
    print(f"在第{cell.row}行的單元格中沒有找到匹配項。")
    # 打印用於匹配的單元格值，以及與其相似度最高的keywords中的字符串和其相似度。
    s = difflib.SequenceMatcher(None, cell.value)
    best_match = None
    best_ratio = 0
    for keyword in keywords:
        s.set_seq2(keyword)
        if s.ratio() > best_ratio:
            best_ratio = s.ratio()
            best_match = keyword
    print(f"與第{cell.row}行的單元格值最接近的是 {best_match}，相似度為 {best_ratio:.2f}")


print("正在嘗試保存Excel文件...")
wb.save(r'C:\Users\user\OneDrive\文件\GitHub\Python Excel for Learning\example.xlsx')
print("Excel文件成功保存。")
