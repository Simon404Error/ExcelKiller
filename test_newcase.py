import sys
sys.path.insert(0, ".")
from excel_tool import SheetData
D = r"D:\ThePC\下载\另一种表格情况"
for f in ["随县2026年10月份动态调整表.xlsx", "曹门村2026年10月份动态调整表.xlsx"]:
    print("=" * 8, f)
    for name in ["1.汇总表", "2.自然变更", "3.监测对象新识别", "4.消除风险户", "5.风险再标注"]:
        sd = SheetData(D + "\\" + f, name)
        print("  [%s] 表头行%s/共%s行 | 数据%d行 x %d列" %
              (name, sd.header_row, sd.header_rows, len(sd.data_rows), sd.n_cols))
        print("     列名:", [h[:14] for h in sd.headers][:12])
        if sd.data_rows:
            print("     首行:", [str(v)[:10] for v in sd.data_rows[0][:8]])
