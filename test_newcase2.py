import sys, os
sys.path.insert(0, ".")
from excel_tool import SheetData, plan_merge, export_merge
D = r"D:\ThePC\下载\另一种表格情况"
vills = ["曹门村", "车店村", "随县大房湾村"]
for sheet in ["2.自然变更", "3.监测对象新识别", "4.消除风险户", "5.风险再标注"]:
    tabs = []
    for v in vills:
        p = os.path.join(D, "%s2026年10月份动态调整表.xlsx" % v)
        tabs.append(SheetData(p, sheet))
    plan = plan_merge(tabs, None, {}, add_source=True)
    print("[%s] 合并: %s" % (sheet, plan["stats"]))
    # 统计每个村的行数
    print("     各村:", [(t.sheet, len(t.data_rows)) for t in tabs])
print()
# 展示“一户多人留空”现象
sd = SheetData(os.path.join(D, "曹门村2026年10月份动态调整表.xlsx"), "3.监测对象新识别")
for i, r in enumerate(sd.data_rows):
    print("行%d:" % (sd.first_data_row + i), [str(v)[:12] if v is not None else "" for v in r])

