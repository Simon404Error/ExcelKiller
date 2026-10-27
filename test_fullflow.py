import sys, os, time, glob
sys.path.insert(0, ".")
import excel_tool as et

D = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "表格示例", "旧的示例，暂时不用")
TARGET = os.path.join(D, "（3315条）10月份预警数据.xlsx")
# 只取各村的“预警数据”文件（不含动态调整表）
files = []
for root, dirs, fs in os.walk(D):
    for f in fs:
        if "预警" in f and f.endswith(".xlsx") and not f.startswith("~$"):
            files.append(os.path.join(root, f))

app = et.App(); app.update()
app.var_dst_path.set(TARGET)
app.load_dst_workbook(TARGET)      # 默认全选 9 个工作表
print("目标工作表:", app.dst_selected_sheets())
t0 = time.time()
app.add_source_paths_all_sheets(sorted(files))
print("批量导入 %d 个文件 → 源表 %d 个，耗时 %.1fs" % (len(files), len(app.sources), time.time()-t0))

app.var_mode.set("fill")           # 只填空
app.key_pairs = []
app.auto_key_pairs()               # 自动匹配键
print("自动匹配键:", app.key_pairs)
app.auto_map()
print("映射:", len(app.mappings), "条")
t1 = time.time()
plans, warns = app.make_plans()
print("规划耗时 %.1fs，工作表 %d 个" % (time.time()-t1, len(plans)))
tot_c = tot_n = 0
for dst, plan in plans:
    st = plan["stats"]
    tot_c += st["将修改单元格"]; tot_n += len(plan["conflicts"])
    print("  %-10s 目标%4d行 修改%3d 冲突%d 未匹配源%d" %
          (dst.sheet, st["目标表数据行"], st["将修改单元格"], len(plan["conflicts"]), st["源表未匹配行"]))
print("合计：将修改 %d 处，冲突 %d 处" % (tot_c, tot_n))
for w in warns[:5]:
    print("  ⚠", w)
app.after(200, app.destroy); app.mainloop()
