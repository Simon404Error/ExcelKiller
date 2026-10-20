import sys, os
sys.path.insert(0, ".")
import excel_tool as et

D = r"D:\ThePC\下载\另一种表格情况"
TARGET = os.path.join(D, "随县2026年10月份动态调整表.xlsx")
vills = ["曹门村", "车店村", "随县大房湾村"]
files = [os.path.join(D, v + "2026年10月份动态调整表.xlsx") for v in vills]

print("== 智能推断“同户共享列” ==")
for f in files:
    sd = et.open_table(f, "3.监测对象新识别")
    print("  ", os.path.basename(f)[:8], "→", et.guess_inherit_cols(sd))

print("== 后缀统一规则（跨全部源表） ==")
rules = et.auto_suffix_rules_multi([et.open_table(f, "3.监测对象新识别") for f in files])
for r in rules:
    print("   ", r["col"].replace(chr(10), "")[:10], ":", r["old"], "→", r["new"])

app = et.App()
app.update()
app.var_dst_path.set(TARGET)
app.load_dst_workbook(TARGET)
app.lb_dst.selection_clear(0, "end")
for i, n in enumerate(app.dst_sheets):
    if n == "3.监测对象新识别":
        app.lb_dst.selection_set(i)
app.on_dst_select()
for f in files:
    app.sources.append(et.open_table(f, "3.监测对象新识别"))
app.refresh_source_list()
app.var_mode.set("append")
app.key_pairs = []
app.auto_map()
app.var_inherit.set(True)      # 智能识别
app.replace_rules = rules
app.var_std.set(True)

plans, warns = app.make_plans()
dst, plan = plans[0]
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "测试输出", "追加合并_继承+标准化2.xlsx")
et.execute_fill_multi(TARGET, plans, out)
chk = et.SheetData(out, "3.监测对象新识别", dst.header_row, dst.header_rows)
print("落盘 %d 行:" % len(chk.data_rows))
for r in chk.data_rows:
    print("   ", [str(v)[:9] if v is not None else "" for v in r])
app.after(200, app.destroy)
app.mainloop()
print("测试通过")

