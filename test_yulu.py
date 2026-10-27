# -*- coding: utf-8 -*-
"""雨露计划：11 个村表 → 总和总表，走真实界面逻辑（导入 + 自动对应 + 补全预演）"""
import sys, os, glob, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import excel_tool as et

D = r"D:\Deepseek-Harness\work\ExcelKiller\表格示例\2026秋季雨露计划"
TARGET = [p for p in glob.glob(os.path.join(D, "总和", "*.xls*")) if not os.path.basename(p).startswith("~$")][0]
print("目标:", os.path.basename(TARGET))

app = et.App(); app.update()
t0 = time.time()
app.var_dst_path.set(TARGET)
app.load_dst_workbook(TARGET)
print("目标表加载 %.2fs，工作表: %s" % (time.time() - t0, app.dst_selected_sheets()))

files = sorted(p for p in glob.glob(os.path.join(D, "*.xlsx")) if not os.path.basename(p).startswith("~$"))
t0 = time.time()
app.add_source_paths_all_sheets(files)
print("导入 %d 个村表 %.2fs，源表 %d 个" % (len(files), time.time() - t0, len(app.sources)))

print("对应关系:")
for p in app.pairs:
    print("   ", app.sources[p["src"]].label, "→", p["target"])

app.var_mode.set("fill")
app.key_pairs = []
app.auto_key_pairs()
print("自动匹配键:", app.key_pairs)
app.auto_map()
print("映射:", [(m["src_name"][:12], m["dst_name"][:12]) for m in app.mappings])

t0 = time.time()
plans, warns = app.make_plans()
print("预演 %.2fs，计划 %d 个" % (time.time() - t0, len(plans)))
for dst, plan in plans:
    print("   [%s] 目标%3d行 将修改%3d 冲突%d 异名%d" %
          (dst.sheet, len(dst.data_rows), plan["stats"].get("将修改单元格", 0),
           len(plan["conflicts"]), len(plan.get("name_mismatch", []))))
for w in warns[:6]:
    print("   ⚠", w)
app.after(150, app.destroy); app.mainloop()
print("OK")

