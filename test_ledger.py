import sys, os, glob
sys.path.insert(0, ".")
import excel_tool as et

D = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "表格示例", "旧的示例，暂时不用")
T = os.path.join(D, "随县2026年10月份动态调整表.xlsx")

app = et.App(); app.update()
app.var_led_dst.set(T)
app.led_pick_dst.__func__  # noop
# 模拟选择目标表
names = et.list_sheets(T)
app.lb_led_dst.delete(0, "end")
for n in names: app.lb_led_dst.insert("end", n)
for i, n in enumerate(names):
    if "汇总" not in n: app.lb_led_dst.selection_set(i)
print("目标已选表:", [app.lb_led_dst.get(i) for i in app.lb_led_dst.curselection()])

files = sorted(glob.glob(os.path.join(D, "*", "*动态调整*.xlsx")))
app.led_files = files
app.led_refresh()
print("源文件数:", len(app.led_files))

got = app.led_make_plans()
plans, warns = got
for dst, plan in plans:
    print("  [%s] 目标%2d行 + 追加%2d行  来自%d个文件" % (dst.sheet, len(dst.data_rows), len(plan["new_rows"]), len(plan.get("_labels", []))))
print("警告:", warns[:4])

out = os.path.join("测试输出", "台账合并_随县.xlsx")
os.makedirs("测试输出", exist_ok=True)
final = et.execute_fill_multi(T, plans, out)
et.export_report_multi(plans, os.path.join("测试输出", "台账合并_明细.xlsx"))
print("输出:", os.path.exists(final))
for dst, plan in plans:
    chk = et.SheetData(final, dst.sheet)
    print("  [%s] 合并后行数 %d" % (dst.sheet, len(chk.data_rows)))
    for r in chk.data_rows[:3]:
        print("     ", [str(v)[:10] if v is not None else "" for v in r[:8]])
app.after(200, app.destroy); app.mainloop()
print("测试通过")
