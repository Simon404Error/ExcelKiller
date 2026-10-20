# -*- coding: utf-8 -*-
"""界面 + 多工作表流程测试（不开人工交互，自动点击流程）"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import excel_tool as et

base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ZONG = os.path.join(base, "表格示例", "（3315条）10月份预警数据.xlsx")
CUN = os.path.join(base, "表格示例", "曹门村", "曹门村10月份预警数据.xlsx")

et.enable_dpi_awareness()
app = et.App()
app.update()
print("窗口尺寸:", app.winfo_width(), "x", app.winfo_height(),
      "| 屏幕:", app.winfo_screenwidth(), "x", app.winfo_screenheight(),
      "| 缩放因子:", round(app.scale, 2))

# 1) 目标表：加载全部工作表（默认全选）
app.var_dst_path.set(ZONG)
app.load_dst_workbook(ZONG)
app.update()
print("目标已选工作表:", app.dst_selected_sheets())

# 2) 源表：导入村表全部工作表
for name in et.list_sheets(CUN):
    sd = et.open_table(CUN, name)
    if sd.data_rows:
        app.sources.append(sd)
app.refresh_source_list()
app.update()
print("源表数:", len(app.sources))
print("对应关系:")
for p in app.pairs:
    print("   ", app.sources[p["src"]].sheet, "→", p["target"])
print("基准表:", app.lb_dst_info.cget("text"))

# 3) 自动匹配字段映射
app.auto_map()
print("映射条数:", len(app.mappings))

# 4) 设匹配键：身份证（按列名保存，跨工作表自动适配）
sname = next(h for h in app.sources[0].headers if et.alias_group(h) == "id")
dname = next(h for h in app.dst.headers if et.alias_group(h) == "id")
app.key_pairs = [(sname, dname)]
print("匹配键:", app.key_pairs)

# 5) 预演（不写文件）
plans, skipped = app.make_plans()
print("计划数:", len(plans), "跳过:", skipped)
for dst, plan in plans:
    print("   ", dst.sheet, plan["stats"])

# 6) 真正执行到临时文件
out = os.path.join(base, "测试输出", "多工作表补全测试.xlsx")
os.makedirs(os.path.dirname(out), exist_ok=True)
final = et.execute_fill_multi(ZONG, plans, out)
print("写出:", os.path.exists(final))
print("明细:", os.path.exists(et.export_report_multi(
    plans, os.path.join(base, "测试输出", "多工作表明细测试.xlsx"))))
app.after(200, app.destroy)
app.mainloop()
print("界面/多表流程测试通过")
