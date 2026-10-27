# -*- coding: utf-8 -*-
"""真实场景测试：村表 补 总表（三种模式）"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from excel_tool import (SheetData, plan_fill, execute_fill, export_report, alias_group)

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CUN = os.path.join(BASE, "表格示例", "旧的示例，暂时不用", "曹门村", "曹门村10月份预警数据.xlsx")
ZONG = os.path.join(BASE, "表格示例", "旧的示例，暂时不用", "（3315条）10月份预警数据.xlsx")
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "测试输出")
os.makedirs(OUT, exist_ok=True)

KOPT = {"nospace": True, "upper": True, "symbol": True}
FILLS = ["核实情况", "是否纳入监测对象", "纳入的类型"]

for sheet in ["131卫健", "2557医保", "128低保", "50特困供养"]:
    src = SheetData(CUN, sheet)
    dst = SheetData(ZONG, sheet)
    si = next(i for i, h in enumerate(src.headers) if alias_group(h) == "id")
    di = next(i for i, h in enumerate(dst.headers) if alias_group(h) == "id")
    maps = []
    for name in FILLS:
        if name in src.headers and name in dst.headers:
            maps.append({"src": src.headers.index(name), "dst": dst.headers.index(name), "mode": "默认"})
    for mode in ("fill", "fill_add"):
        plan = plan_fill([src], dst, [(si, di)], maps, mode, KOPT)
        print("[%s/%s] %s" % (sheet, mode, plan["stats"]))
        if mode == "fill_add":
            out = os.path.join(OUT, "补全_%s.xlsx" % sheet)
            execute_fill(dst, plan, out, dst.n_cols)
            export_report(dst, plan, os.path.join(OUT, "明细_%s.xlsx" % sheet))
            chk = SheetData(out, sheet, dst.header_row, dst.header_rows)
            for name in FILLS:
                if name in dst.headers:
                    j = dst.headers.index(name)
                    filled = sum(1 for r in chk.data_rows if str(r[j] or "").strip())
                    print("   写入后 %s 非空: %d / %d" % (name, filled, len(chk.data_rows)))

# 新增模式：用一个只含少量人的源表，测“目标表没有就追加”
src = SheetData(CUN, "128低保")
dst = SheetData(ZONG, "128低保")
si = next(i for i, h in enumerate(src.headers) if alias_group(h) == "id")
di = next(i for i, h in enumerate(dst.headers) if alias_group(h) == "id")
# 人为制造两条目标表没有的记录
fake_rows = list(src.data_rows[:2])
src.data_rows = fake_rows + [tuple(("X" + str(i) if k == si else v) for k, v in enumerate(fake_rows[0]))
                             for i in range(2)]
plan = plan_fill([src], dst, [(si, di)], [{"src": si, "dst": di, "mode": "默认"}], "add_new", KOPT)
print("新增模式:", plan["stats"])
out = os.path.join(OUT, "新增测试.xlsx")
execute_fill(dst, plan, out, dst.n_cols)
chk = SheetData(out, "128低保", dst.header_row, dst.header_rows)
print("  追加后行数:", len(chk.data_rows), "（原", len(dst.data_rows), "）")
print("测试结束")

