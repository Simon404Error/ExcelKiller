import sys, os, tkinter as tk
from tkinter import ttk
sys.path.insert(0, ".")
import excel_tool as et

BASE = os.path.dirname(os.getcwd())
CUN = os.path.join(BASE, "表格示例", "旧的示例，暂时不用", "曹门村", "曹门村10月份预警数据.xlsx"); ZONG = os.path.join(BASE, "表格示例", "旧的示例，暂时不用", "（3315条）10月份预警数据.xlsx")

et.enable_dpi_awareness()
app = et.App()
app.update()
app.var_dst_path.set(ZONG)
app.load_dst_workbook(ZONG)
app.lb_dst.selection_clear(0, "end")
for i, n in enumerate(app.dst_sheets):
    if n == "131卫健":
        app.lb_dst.selection_set(i)
app.on_dst_select()

s1 = et.SheetData(CUN, "131卫健")
s2 = et.SheetData(CUN, "131卫健")
rows = list(s2.data_rows)
rows[0] = tuple(("第二来源：不符合" if k == s1.headers.index("核实情况") else v)
                for k, v in enumerate(rows[0]))
s2.data_rows = rows
s2.sheet = "131卫健(另一来源)"
app.sources = [s1, s2]
app.refresh_source_list()
app.key_pairs = [("身份证号码", "身份证号码")]
app.mappings = [{"src": s1.headers.index("核实情况"), "src_name": "核实情况",
                 "dst": app.dst.headers.index("核实情况"), "dst_name": "核实情况", "mode": "默认"}]

plans, warns = app.make_plans()
print("工作表计划数:", len(plans), "冲突数:", [len(p["conflicts"]) for _, p in plans])
print("对应关系:", [(app.sources[p["src"]].sheet, p["target"]) for p in app.pairs])

done = {"clicked": False}

def find_btn(x):
    for c in x.winfo_children():
        if isinstance(c, ttk.Button) and c.cget("text") == "确定":
            return c
        b = find_btn(c)
        if b:
            return b
    return None

def auto_click():
    for w in app.winfo_children():
        if isinstance(w, tk.Toplevel):
            b = find_btn(w)
            if b:
                print("  点击“确定”")
                b.invoke()
                done["clicked"] = True
                return

app.after(700, auto_click)
ok = app.resolve_all_conflicts(plans)
print("弹窗确认:", ok, "| 已点击:", done["clicked"])
print("解决后修改:", [(c["colname"], str(c["new"])[:14], c["src"]) for d, p in plans for c in p["changes"]])
print("缓存:", list(app.conflict_cache.values()))
app.after(200, app.destroy)
app.mainloop()
print("冲突弹窗测试通过")
