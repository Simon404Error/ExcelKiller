import sys, os
sys.path.insert(0, ".")
import excel_tool as et

BASE = os.path.dirname(os.getcwd())
CUN = os.path.join(BASE, "表格示例", "旧的示例，暂时不用", "曹门村", "曹门村10月份预警数据.xlsx"); ZONG = os.path.join(BASE, "表格示例", "旧的示例，暂时不用", "（3315条）10月份预警数据.xlsx")

et.enable_dpi_awareness()
app = et.App()
app.update()
print("初始字体:", app.var_font.get(), "| 窗口:", app.winfo_width(), "x", app.winfo_height())

# 字体调整
for s in (8, 12, 16):
    app.var_font.set(s)
    app.on_font_change()
    app.update()
    print("字体", s, "→", app.sp_font.get(), "| 设置文件:", os.path.exists(et._settings_path()))

# 目标 + 源表
app.var_dst_path.set(ZONG)
app.load_dst_workbook(ZONG)
print("自动匹配键（仅目标无源）:", app.key_pairs)

for name in et.list_sheets(CUN):
    sd = et.open_table(CUN, name)
    if sd.data_rows:
        app.sources.append(sd)
app.refresh_source_list()
print("源表多选支持:", app.lb_src.cget("selectmode"))
print("自动匹配键:", app.key_pairs)
app.lb_src.selection_set(0)
app.lb_src.selection_set(2)
print("多选行:", app.lb_src.curselection())

plans, warns = app.make_plans()
print("计划:", len(plans), "警告:", warns)
tot = sum(len(p["changes"]) for _, p in plans)
print("将修改:", tot)
app.after(200, app.destroy)
app.mainloop()
print("测试通过")
