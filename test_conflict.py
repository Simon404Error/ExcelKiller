import sys, os
sys.path.insert(0, ".")
from excel_tool import SheetData, plan_fill, resolve_conflicts, execute_fill_multi, export_report_multi

BASE = os.path.dirname(os.getcwd())
CUN = os.path.join(BASE, "表格示例", "旧的示例，暂时不用", "曹门村", "曹门村10月份预警数据.xlsx"); ZONG = os.path.join(BASE, "表格示例", "旧的示例，暂时不用", "（3315条）10月份预警数据.xlsx")
KOPT = {"nospace": True, "upper": True, "symbol": True}

src = SheetData(CUN, "131卫健")
dst = SheetData(ZONG, "131卫健")
si = src.headers.index("身份证号码"); di = dst.headers.index("身份证号码")
maps = [{"src": src.headers.index("核实情况"), "dst": dst.headers.index("核实情况"), "mode": "默认"}]

# 造一个“另一个源表”，对同一个人给出不同核实情况
src2 = SheetData(CUN, "131卫健")
rows = list(src2.data_rows)
rows[0] = tuple(("另一个人说的：不符合纳入条件" if k == src.headers.index("核实情况") else v)
                for k, v in enumerate(rows[0]))
src2.data_rows = rows
src2.sheet = "131卫健(第二来源)"

plan = plan_fill([src, src2], dst, [(si, di)], maps, "fill", KOPT)
print("冲突数:", len(plan["conflicts"]), "| 统计:", plan["stats"])
for c in plan["conflicts"]:
    print("  冲突:", c["type"], "行", c.get("excel_row"), c["colname"], c["key"])
    for o in c["options"]:
        print("     候选:", o["src"], "=", o["value"])

# 模拟用户选择第 2 个源表
ck = [("cell", c["ri"], c["col"]) if c["type"] == "cell" else ("newrow", c["key"])
      for c in plan["conflicts"]]
resolve_conflicts(plan, {k: 1 for k in ck})
print("解决后修改数:", len(plan["changes"]), "冲突剩:", len(plan["conflicts"]))
print("采用值:", [(c["colname"], c["new"], c["src"]) for c in plan["changes"]])
print("已选记录:", [(r["type"], r["src"], r["value"][:16]) for r in plan["resolved"]])

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "测试输出", "冲突处理测试.xlsx")
os.makedirs(os.path.join(os.path.dirname(os.path.abspath(__file__)), "测试输出"), exist_ok=True)
execute_fill_multi(ZONG, [(dst, plan)], out)
export_report_multi([(dst, plan)], os.path.join(os.path.dirname(os.path.abspath(__file__)), "测试输出", "冲突处理明细.xlsx"))
chk = SheetData(out, "131卫健")
print("落盘核实情况:", chk.data_rows[0][chk.headers.index("核实情况")][:20])
print("测试通过")
