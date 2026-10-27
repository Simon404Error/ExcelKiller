# -*- coding: utf-8 -*-
"""引擎冒烟测试（不开界面）"""
import os, sys, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from excel_tool import (SheetData, list_sheets, plan_fill, execute_fill, export_report,
                        plan_diff, export_diff, plan_merge, export_merge, alias_group)

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CUN = os.path.join(BASE, "表格示例", "旧的示例，暂时不用", "曹门村", "曹门村10月份预警数据.xlsx")
ZONG = os.path.join(BASE, "表格示例", "旧的示例，暂时不用", "（3315条）10月份预警数据.xlsx")
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "测试输出")
os.makedirs(OUT, exist_ok=True)

print("sheets:", list_sheets(ZONG))

for sheet in list_sheets(ZONG):
    for p in (CUN, ZONG):
        sd = SheetData(p, sheet)
        print(p.split(os.sep)[-1], "|", sheet, "| 表头行", sd.header_row, "行数", sd.header_rows,
              "| 数据", len(sd.data_rows), "| 列", sd.headers[:5])

# ---------- 补全：以 2557医保 为例，把总表某列覆盖为空后，用源表补回 ----------
src = SheetData(CUN, "128低保")
dst = SheetData(ZONG, "128低保")
si = [i for i, h in enumerate(src.headers) if alias_group(h) == "id"][0]
di = [i for i, h in enumerate(dst.headers) if alias_group(h) == "id"][0]
checks_src = src.headers.index("核实情况")
checks_dst = dst.headers.index("核实情况")
maps = [{"src": checks_src, "dst": checks_dst, "mode": "默认"}]
plan = plan_fill([src], dst, [(si, di)], maps, "fill", {"nospace": True, "upper": True})
print("补全统计:", plan["stats"])
out1 = os.path.join(OUT, "补全结果_128低保.xlsx")
execute_fill(dst, plan, out1, dst.n_cols)
rep = os.path.join(OUT, "补全明细.xlsx")
export_report(dst, plan, rep)
print("输出:", os.path.exists(out1), os.path.exists(rep))

# 校验写入结果
check = SheetData(out1, "128低保", dst.header_row, dst.header_rows)
print("写入后 核实情况 非空行数:", sum(1 for r in check.data_rows if r[checks_dst] not in (None, "")))

# ---------- 对账 ----------
a = SheetData(CUN, "182人社")
b = SheetData(ZONG, "182人社")
p2 = plan_diff(a, b, [si], [di], {"nospace": True, "upper": True})
print("对账统计:", p2["stats"])
out2 = os.path.join(OUT, "对账报告.xlsx")
export_diff(a, b, p2, out2)
print("输出:", os.path.exists(out2))

# ---------- 汇总 ----------
tabs = [SheetData(CUN, s) for s in list_sheets(CUN)[:4]]
p3 = plan_merge(tabs, dedupe_cols=["身份证"], add_source=True)
print("汇总统计:", p3["stats"])
out3 = os.path.join(OUT, "汇总.xlsx")
export_merge(p3, out3)
print("输出:", os.path.exists(out3))
print("全部测试通过")

