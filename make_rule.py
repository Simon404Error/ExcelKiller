import sys, os, json
sys.path.insert(0, ".")
from excel_tool import list_sheets, SheetData, alias_group

BASE = os.path.dirname(os.getcwd())
CUN = os.path.join(BASE, "表格示例", "曹门村", "曹门村10月份预警数据.xlsx")
ZONG = os.path.join(BASE, "表格示例", "（3315条）10月份预警数据.xlsx")
sheets = list_sheets(ZONG)

rule = {
    "version": 2,
    "app": "Excel填补助手 for睿 v0.1.0",
    "说明": "示例规则：村表.xlsx 全部工作表 → 总表.xlsx 同名工作表，按身份证号补全“核实情况 / 是否纳入监测对象 / 纳入的类型”。匹配键与字段均按列名保存，换单位、换月份同样适用。",
    "target": {"path": ZONG, "sheets": sheets, "header_row": "自动", "header_rows": "自动"},
    "sources": [],
    "pairs": [],
    "key_pairs": [["公民身份号码", "公民身份号码"]],
    "key_options": {"nospace": True, "upper": True, "symbol": True, "digits": False},
    "mappings": [],
    "mode": "fill",
    "keep_original": True,
}
for i, name in enumerate(sheets):
    sd = SheetData(CUN, name)
    rule["sources"].append({"path": CUN, "sheet": name,
                            "header_row": sd.header_row, "header_rows": sd.header_rows})
    rule["pairs"].append({"src": i, "target": name})

sd0 = SheetData(CUN, sheets[0])
dd0 = SheetData(ZONG, sheets[0])
for fname in ["核实情况", "是否纳入监测对象", "纳入的类型"]:
    if fname in sd0.headers and fname in dd0.headers:
        rule["mappings"].append({"src": sd0.headers.index(fname), "src_name": fname,
                                 "dst": dd0.headers.index(fname), "dst_name": fname,
                                 "mode": "默认"})

open("示例规则_村表补总表_全工作表.json", "w", encoding="utf-8").write(
    json.dumps(rule, ensure_ascii=False, indent=2))
print("规则已生成，映射：", [m["src_name"] for m in rule["mappings"]])
