import sys, os, glob
sys.path.insert(0, ".")
import openpyxl

D = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "表格示例", "旧的示例，暂时不用")
files = []
for root, dirs, fs in os.walk(D):
    for f in fs:
        if f.lower().endswith(".xlsx") and not f.startswith("~$"):
            files.append(os.path.join(root, f))
print("扫描", len(files), "个文件")
hit_filter = hit_hidden = 0
for p in files:
    wb = openpyxl.load_workbook(p)
    for ws in wb.worksheets:
        af = ws.auto_filter.ref
        hidden = [r for r, dim in ws.row_dimensions.items() if dim.hidden]
        if af:
            hit_filter += 1
            print("  有筛选:", os.path.basename(p), "[%s]" % ws.title, "范围", af,
                  "| 隐藏行数", len(hidden))
        elif hidden:
            hit_hidden += 1
            print("  仅隐藏行:", os.path.basename(p), "[%s]" % ws.title, "隐藏", len(hidden), "行")
    wb.close()
print("有筛选的工作表:", hit_filter, "| 仅隐藏行的:", hit_hidden)
