# -*- coding: utf-8 -*-
"""
Excel 数据助手 v1.0
====================
三大功能：
  ① 信息补全：按源表（可多个工作表）对目标表进行字段补全 / 覆盖 / 新增行
  ② 两表对账：按主键找出“仅A有 / 仅B有 / 字段不一致”
  ③ 多表汇总：多个表合并成一张（自动并集表头 + 来源列 + 可选去重）

设计要点（针对真实表格“格式不固定”）：
  - 表头行自动识别，识别错了可手动指定（支持 2 行合并表头）
  - 列名自动同义匹配（身份证号/公民身份号码/身份证证件号码 …）
  - 匹配键可多列、可设清洗规则（忽略空格、大小写、符号、只比对数字）
  - 三种写入模式 + 逐字段覆盖策略；默认保留原文件，先预演后执行
"""
import os
import re
import sys
import copy
import json
import shutil
import datetime
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment

APP_TITLE = "Excel填补助手 for睿 v0.1.2"

# ============================================================
#                        一、通用清洗
# ============================================================
PUNCT_RE = re.compile(r"[^0-9A-Za-z\u4e00-\u9fff]")


def to_half(s):
    """全角转半角"""
    out = []
    for ch in s:
        o = ord(ch)
        if o == 0x3000:
            out.append(" ")
        elif 0xFF01 <= o <= 0xFF5E:
            out.append(chr(o - 0xFEE0))
        else:
            out.append(ch)
    return "".join(out)


def cell_text(v):
    """单元格 -> 干净文本"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "是" if v else "否"
    if isinstance(v, datetime.datetime):
        if v.hour or v.minute or v.second:
            return v.strftime("%Y-%m-%d %H:%M:%S")
        return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    return str(v).replace("\u3000", " ").replace("\xa0", " ").strip()


def norm_value(v, opts):
    """按清洗规则把值标准化，用于匹配/比较"""
    s = cell_text(v)
    if opts.get("half", True):
        s = to_half(s)
    if opts.get("nospace", True):
        s = re.sub(r"\s+", "", s)
    if opts.get("symbol", False):
        s = PUNCT_RE.sub("", s)
    if opts.get("upper", True):
        s = s.upper()
    if opts.get("digits", False):
        s = re.sub(r"\D", "", s)
    return s


def norm_colname(s):
    s = to_half(cell_text(s))
    s = re.sub(r"\s+", "", s)
    return s.lower()


# 列名同义词分组：用于“自动匹配”表头
ALIAS_GROUPS = [
    ("id", ["身份证号", "身份证号码", "公民身份号码", "身份证证件号码", "身份证件号码",
            "证件号码", "身份证", "公民身份证号"]),
    ("name", ["姓名", "名字", "人员姓名", "当事人姓名", "户主姓名", "被核查人", "人员"]),
    ("phone", ["联系电话", "手机", "手机号码", "手机号", "电话", "联系方式", "移动电话"]),
    ("village", ["村社区名称", "所属村组", "村名", "行政村", "所属村", "村"]),
    ("town", ["乡镇街道", "所属乡镇", "乡镇", "街道", "街道区划"]),
    ("addr", ["家庭住址", "户籍详细地址", "户籍详址", "居住地地址", "地址", "住址", "户籍地址"]),
    ("sex", ["性别"]),
    ("birth", ["出生日期", "出生年月", "出生年月日"]),
    ("check", ["核实情况"]),
    ("monitor", ["是否纳入监测对象"]),
    ("mtype", ["纳入的类型"]),
    ("area", ["所属区县", "区县", "参保地"]),
]
ALIAS_MAP = {}
for _g, _names in ALIAS_GROUPS:
    for _n in _names:
        ALIAS_MAP[norm_colname(_n)] = _g


def alias_group(colname):
    nc = norm_colname(colname)
    if nc in ALIAS_MAP:
        return ALIAS_MAP[nc]
    return None


# ============================================================
#                        二、表结构读取
# ============================================================
def list_sheets(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    names = list(wb.sheetnames)
    wb.close()
    return names


def detect_header_from_ws(ws, merged):
    """返回 (表头行号1-based, 表头行数)

    判据：表头行 = “短文本”单元格最多的一行（列名一般较短、互不重复、
    且其下方紧接着有数据）。避免把含长文本的数据行误判成表头。
    """
    grid = {}
    for r in range(1, min(ws.max_row, 15) + 1):
        grid[r] = [cell_text(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
    rows = {r: [v for v in vals if v] for r, vals in grid.items()}

    best_score, best_row = -1e9, 1
    data_like = re.compile(r"^(\d{11,18}[Xx]?|20\d{6,}|19\d{6,}|\d{4}-\d{2}-\d{2}"
                           r"|\d{4}/\d{1,2}/\d{1,2}|\d+(\.\d+)?)$")
    for r, ne in rows.items():
        if len(ne) < 2:
            continue
        short = [v for v in ne
                 if 0 < len(v) <= 15 and not re.fullmatch(r"-?\d+(\.\d+)?", v)]
        if len(short) < max(2, len(ne) * 0.5):
            continue
        uniq = len(set(ne)) / float(len(ne))
        # 表头里的值几乎都不是“数据样”内容：含身份证/日期/纯数字越多，越不像表头
        like = sum(1 for v in ne if data_like.match(v.replace(" ", "")))
        score = len(short) + uniq * 3 + (1 if r <= 3 else 0) - like * 3
        # 下方需有数据行（数据连续性）
        nxt = rows.get(r + 1) or []
        if len(nxt) >= max(2, len(ne) * 0.4):
            score += 1.5
        if score > best_score:
            best_score, best_row = score, r
    span = 1
    for (r1, r2, c1, c2) in merged:
        if r1 <= best_row <= r2 and c2 > c1:
            span = 2
            break
    # 有些表把部分列名写在上方标题行里（本行为空、上行有值）→ 表头向上扩一行
    if best_row > 1 and span == 1:
        cur, prev = grid.get(best_row, []), grid.get(best_row - 1, [])
        extra = sum(1 for i in range(len(cur))
                    if cur[i] == "" and i < len(prev) and prev[i] != "")
        if extra >= 2 and len([v for v in prev if v]) >= 2:
            best_row -= 1
            span = 2
    if best_row + span - 1 > ws.max_row:
        span = 1
    return best_row, span


class SheetData:
    """一个工作表的只读快照：表头 + 数据行

    传入 wb（已打开的工作簿）可避免同一文件被反复加载，批量导入时快很多。
    """

    def __init__(self, path, sheet=None, header_row=None, header_rows=None, wb=None,
                 visible_only=False):
        self.path = os.path.abspath(path)
        self.visible_only = visible_only
        own = False
        if wb is None:
            wb = openpyxl.load_workbook(self.path, data_only=True)
            own = True
        try:
            self._load(wb, sheet, header_row, header_rows)
        finally:
            if own:
                wb.close()

    def _load(self, wb, sheet, header_row, header_rows):
        self.sheet = sheet if sheet in wb.sheetnames else wb.sheetnames[0]
        ws = wb[self.sheet]
        self.n_rows, self.n_cols = ws.max_row, ws.max_column
        # 筛选/隐藏行：openpyxl 依然能读到数据，这里统计出来供用户选择
        self.filter_ref = (ws.auto_filter.ref or "") if ws.auto_filter else ""
        self.hidden_rows = {r for r, dim in ws.row_dimensions.items() if dim.hidden}
        self.hidden_in_data = 0
        merged = [(m.min_row, m.max_row, m.min_col, m.max_col) for m in ws.merged_cells.ranges]
        if header_row is None:
            header_row, span = detect_header_from_ws(ws, merged)
            if header_rows:
                span = header_rows
        else:
            span = header_rows or 1
        self.header_row = header_row
        self.header_rows = span

        # 合并单元格：让子单元格继承父单元格的值（用于 2 行表头）
        merged_fill = {}
        for (r1, r2, c1, c2) in merged:
            val = cell_text(ws.cell(row=r1, column=c1).value)
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    merged_fill[(r, c)] = val

        headers, used = [], {}
        for c in range(1, self.n_cols + 1):
            parts = []
            for i, r in enumerate(range(header_row, header_row + span)):
                v = merged_fill.get((r, c), cell_text(ws.cell(row=r, column=c).value))
                if i == 0 and span > 1 and len(v) > 12:
                    continue      # 该列上方是大标题（合并单元格），不当列名
                if v and (not parts or parts[-1] != v):
                    parts.append(v)
            name = "_".join(parts) if parts else "列%s" % get_column_letter(c)
            k = name
            n = 2
            while k in used:
                k = "%s(%d)" % (name, n)
                n += 1
            used[k] = 1
            headers.append(k)
        self.headers = headers

        self.first_data_row = header_row + span          # 1-based，数据首行
        hdr_norm = [norm_colname(h) for h in headers]
        rows = []
        for r in range(self.first_data_row, ws.max_row + 1):
            row = tuple(ws.cell(row=r, column=c).value for c in range(1, self.n_cols + 1))
            if all(cell_text(x) == "" for x in row):
                continue
            # 有的表下面又贴了一份“带表头的表”：把这种重复表头行当数据会污染整表，跳过
            hits = 0
            none = 0
            for i, v in enumerate(row):
                t = cell_text(v)
                if not t:
                    continue
                none += 1
                if i < len(hdr_norm) and norm_colname(t) == hdr_norm[i] and t:
                    hits += 1
            if none >= 3 and hits >= max(2, none * 0.5):
                continue
            if r in self.hidden_rows:
                self.hidden_in_data += 1
                if self.visible_only:
                    continue          # 用户选择“只处理可见行”
            rows.append((r, row))
        self.data_rows = [x[1] for x in rows]
        self._excel_rows = [x[0] for x in rows]

    @property
    def visible_count(self):
        return len(self.data_rows) if self.visible_only else len(self.data_rows) - self.hidden_in_data

    @property
    def label(self):
        return "%s / %s" % (os.path.basename(self.path), self.sheet)

    def col_label(self, idx):
        return "%s｜%s" % (get_column_letter(idx + 1), self.headers[idx])

    def col_labels(self):
        return [self.col_label(i) for i in range(len(self.headers))]


def open_table(path, sheet, header_row=None, header_rows=None, wb=None, visible_only=False):
    return SheetData(path, sheet, header_row, header_rows, wb, visible_only)


# ============================================================
#                        三、功能①：信息补全
# ============================================================
MODE_LABELS = [
    ("fill", "仅填空（只补空单元格，不动已有内容）"),
    ("overwrite", "覆盖（用源表值覆盖目标表）"),
    ("add_new", "仅新增（目标表没有的记录追加到末尾）"),
    ("fill_add", "填空并新增（推荐）"),
    ("overwrite_add", "覆盖并新增"),
]


def guess_inherit_cols(table):
    """智能推断“同户共享”的列：在每一户（块）内取值一致、且存在空白的列

    块 = 从“户主姓名”或“序号”非空的行开始到下一户之前
    """
    headers = table.headers
    anchor = None
    for i, h in enumerate(headers):
        if "户主" in h and ("姓名" in h or alias_group(h) == "name"):
            anchor = i
            break
    if anchor is None:
        for i, h in enumerate(headers):
            if norm_colname(h) in ("序号", "编号", "户主姓名"):
                anchor = i
                break
    if anchor is None:
        anchor = 1 if len(headers) > 1 else 0
    blocks = []
    for row in table.data_rows:
        if cell_text(row[anchor]) or not blocks:
            blocks.append([])
        blocks[-1].append(row)
    if not blocks:
        return []
    chosen = []
    for ci, h in enumerate(headers):
        if norm_colname(h) in ("序号", "编号", "序"):
            continue
        has_blank, consistent, has_value = False, True, False
        for b in blocks:
            vals = [cell_text(r[ci]) for r in b if ci < len(r) and cell_text(r[ci])]
            if len(vals) < len(b):
                has_blank = True
            if vals:
                has_value = True
            if len(set(vals)) > 1:
                consistent = False
                break
        if has_blank and consistent and has_value:
            chosen.append(h)
    return chosen


def inherit_blank_rows(table, col_names=None):
    """空白继承：空白单元格取上一行的值（解决“一户多人只填一次户主信息”的台账）

    col_names 为空时自动推断“同户共享列”
    """
    headers = table.headers
    if col_names:
        idxs = [i for i, h in enumerate(headers) if h in col_names]
    else:
        cols = guess_inherit_cols(table)
        if cols:
            idxs = [i for i, h in enumerate(headers) if h in cols]
        else:
            idxs = [i for i, h in enumerate(headers)
                    if norm_colname(h) not in ("序号", "编号", "序")]
    rows = [list(r) for r in table.data_rows]
    prev = [None] * len(headers)
    out = []
    for r in rows:
        cur = list(r)
        for i in idxs:
            if i < len(cur) and cell_text(cur[i]) == "" and prev[i] is not None:
                cur[i] = prev[i]
        for i in idxs:
            if i < len(cur) and cell_text(cur[i]) != "":
                prev[i] = cur[i]
        out.append(tuple(cur))
    new = copy.copy(table)
    new.data_rows = out
    return new


def standardize_table(table, rules):
    """按替换规则标准化字段值：rules=[{"col":列名或空, "old":原值, "new":新值}]"""
    if not rules:
        return table
    idx_map = {}
    for r in rules:
        col = r.get("col") or ""
        key = norm_colname(col)
        idx_map.setdefault(key, []).append((cell_text(r.get("old")), r.get("new")))
    rows = []
    for row in table.data_rows:
        cur = list(row)
        for i, h in enumerate(table.headers):
            hv = cell_text(cur[i]) if i < len(cur) else ""
            if hv == "":
                continue
            for key in (norm_colname(h), ""):      # 先匹配该列规则，再匹配“所有列”规则
                for old, new in idx_map.get(key, []):
                    if old and hv == old:
                        cur[i] = new if isinstance(new, (int, float)) else cell_text(new)
                        hv = cell_text(cur[i])
                        break
        rows.append(tuple(cur))
    new = copy.copy(table)
    new.data_rows = rows
    return new


def auto_suffix_rules(table, suffixes=("镇", "乡", "村", "社区")):
    """自动生成“后缀统一”替换规则：万和 / 万和镇 → 统一为 万和镇"""
    rules, seen = [], set()
    for ci, h in enumerate(table.headers):
        vals = [cell_text(r[ci]) for r in table.data_rows if ci < len(r) and cell_text(r[ci]) != ""]
        s = set(vals)
        for v in s:
            for suf in suffixes:
                if v.endswith(suf) and len(v) > len(suf):
                    short = v[:-len(suf)]
                    if short in s and (norm_colname(h), short) not in seen:
                        seen.add((norm_colname(h), short))
                        rules.append({"col": h, "old": short, "new": v})
    return rules


def auto_suffix_rules_multi(tables, suffixes=("镇", "乡", "村", "社区")):
    """跨多个表统一生成“后缀统一”规则（值集合跨表合并后再比对）"""
    rules, seen = [], set()
    by_col = {}
    for t in tables:
        for ci, h in enumerate(t.headers):
            s = by_col.setdefault(norm_colname(h), {"name": h, "vals": set()})
            for r in t.data_rows:
                if ci < len(r) and cell_text(r[ci]):
                    s["vals"].add(cell_text(r[ci]))
    for key, s in by_col.items():
        for v in s["vals"]:
            for suf in suffixes:
                if v.endswith(suf) and len(v) > len(suf):
                    short = v[:-len(suf)]
                    if short in s["vals"] and (key, short) not in seen:
                        seen.add((key, short))
                        rules.append({"col": s["name"], "old": short, "new": v})
    return rules


def auto_suffix_fill_rules(tables):
    """给“村/镇”类列补齐缺失后缀：曹门 → 曹门村、万和 → 万和镇"""
    want = {}
    for k in ("镇", "乡镇", "所属乡镇", "街道"):
        want[norm_colname(k)] = "镇"
    for k in ("村", "村名", "所属村组", "行政村"):
        want[norm_colname(k)] = "村"
    tails = ("镇", "乡", "村", "社区", "居委会", "农场", "景区", "街道",
             "市", "县", "省", "区")
    stop = {"县", "市", "镇", "村", "区", "乡", "省", "社区", "街道", "无", "空"}
    by_col = {}
    for t in tables:
        for ci, h in enumerate(t.headers):
            s = by_col.setdefault(norm_colname(h), {"name": h, "vals": set()})
            for r in t.data_rows:
                if ci < len(r) and cell_text(r[ci]):
                    s["vals"].add(cell_text(r[ci]))
    rules = []
    for key, s in by_col.items():
        suf = want.get(key)
        if not suf:
            continue
        for v in s["vals"]:
            if v in stop or len(v) < 2:
                continue
            # 重复后缀：万家湾村村 → 万家湾村
            if v.endswith(suf + suf) and len(v) > len(suf) * 2:
                rules.append({"col": s["name"], "old": v, "new": v[:-len(suf)]})
                continue
            if not any(v.endswith(x) for x in tails):
                rules.append({"col": s["name"], "old": v, "new": v + suf})
    return rules


def plan_append(sources, dst, mappings, auto_seq_col=None):
    """追加合并：不判重、不比对，源表所有行按列名映射后追加到目标表末尾"""
    plan = {
        "changes": [], "new_rows": [], "unmatched_dst": [], "unmatched_src": [],
        "dup_dst": [], "dup_src": [], "conflicts": [], "errors": [],
        "_sheet": dst.sheet,
    }
    for src in sources:
        for row in src.data_rows:
            if all(cell_text(v) == "" for v in row):
                continue
            vals = {}
            for m in mappings:
                si = m.get("src")
                sn = m.get("src_name")
                if sn is not None:
                    si = src.headers.index(sn) if sn in src.headers else None
                sv = row[si] if (si is not None and si < len(row)) else None
                if cell_text(sv) != "":
                    vals[m["dst"]] = sv
            plan["new_rows"].append({"src": src.label, "key": "追加", "values": vals})
    # 序号自动重排
    if auto_seq_col is not None:
        start = len(dst.data_rows) + 1
        for i, nr in enumerate(plan["new_rows"]):
            nr["values"][auto_seq_col] = start + i
    plan["_meta"] = {"dst_rows": len(dst.data_rows),
                     "src_rows": sum(len(s.data_rows) for s in sources),
                     "matched": 0}
    plan["stats"] = fill_stats(plan)
    return plan


def build_key_index(table, key_cols, kopt):
    idx = {}
    for ri, row in enumerate(table.data_rows):
        key = tuple(norm_value(row[c], kopt) for c in key_cols)
        if not any(key):
            continue
        idx.setdefault(key, []).append(ri)
    return idx


def key_of(row, key_cols, kopt):
    return tuple(norm_value(row[c], kopt) for c in key_cols)


def plan_fill(sources, dst, key_pairs, mappings, mode, kopt,
              name_pair=None, name_skip=False):
    """生成补全计划（不写文件）

    多源表时，同一目标单元格若被不同源表给出不同值 → 记为“冲突”，
    由界面弹窗让用户指定采用哪个源表的值（resolve_conflicts）。
    """
    modify = mode in ("fill", "overwrite", "fill_add", "overwrite_add")
    add_new = mode in ("add_new", "fill_add", "overwrite_add")
    default_mm = "overwrite" if mode in ("overwrite", "overwrite_add") else "fill"

    dst_keys = [p[1] for p in key_pairs]
    src_keys = [p[0] for p in key_pairs]

    dst_idx = build_key_index(dst, dst_keys, kopt)
    dup_dst = [(k, v) for k, v in dst_idx.items() if len(v) > 1]

    plan = {
        "changes": [], "new_rows": [], "unmatched_dst": [], "unmatched_src": [],
        "dup_dst": dup_dst, "dup_src": [], "conflicts": [], "errors": [],
        "name_mismatch": [], "_sheet": dst.sheet,
    }
    matched_dst_keys = set()
    cands = {}        # (ri, col) -> 候选值
    new_cands = {}    # 键 -> 候选新增行

    for src in sources:
        s_idx = build_key_index(src, src_keys, kopt)
        for k, v in s_idx.items():
            if len(v) > 1:
                plan["dup_src"].append((src.label, k, len(v)))
        for key, hits in s_idx.items():
            srow = src.data_rows[hits[0]]
            key_str = " + ".join(key)
            d_hits = dst_idx.get(key)
            if d_hits:
                matched_dst_keys.add(key)
                if not modify:
                    continue
                d_ri = d_hits[0]
                drow = dst.data_rows[d_ri]
                # 同键异名：键对上了但姓名对不上 → 多半是源表“错行”了
                if name_pair:
                    si_n, di_n = name_pair
                    s_nm = cell_text(srow[si_n]) if si_n < len(srow) else ""
                    d_nm = cell_text(drow[di_n]) if di_n < len(drow) else ""
                    if s_nm and d_nm and norm_value(s_nm, kopt) != norm_value(d_nm, kopt):
                        plan["name_mismatch"].append(
                            (key_str, s_nm, d_nm, src.label))
                        if name_skip:
                            continue
                for m in mappings:
                    sv = srow[m["src"]] if m["src"] < len(srow) else None
                    if cell_text(sv) == "":
                        continue
                    dv = drow[m["dst"]]
                    mm = m.get("mode") or "默认"
                    eff = default_mm if mm == "默认" else ("fill" if mm == "仅填空" else
                                                          "overwrite" if mm == "覆盖" else "skip")
                    if eff == "skip":
                        continue
                    if eff == "fill" and cell_text(dv) != "":
                        continue
                    if cell_text(sv) == cell_text(dv):
                        continue
                    c = cands.setdefault((d_ri, m["dst"]), {
                        "ri": d_ri, "col": m["dst"], "colname": dst.headers[m["dst"]],
                        "key": key_str, "opts": []})
                    c["opts"].append({"src": src.label, "value": sv})
            else:
                plan["unmatched_src"].append((src.label, key_str,
                                              cell_text(srow[src_keys[0]]) if src_keys else ""))
                if add_new:
                    newrow = {}
                    for m in mappings:
                        sv = srow[m["src"]] if m["src"] < len(srow) else None
                        if cell_text(sv) != "":
                            newrow[m["dst"]] = sv
                    nc = new_cands.setdefault(key_str, {"key": key_str, "opts": []})
                    nc["opts"].append({"src": src.label, "values": newrow})

    # ---- 按单元格归并：唯一值直接写入，多值进入冲突清单 ----
    for (ri, col), c in cands.items():
        uniq, seen = [], set()
        for o in c["opts"]:
            nv = norm_value(o["value"], {"symbol": False})
            if nv in seen:
                continue
            seen.add(nv)
            uniq.append(o)
        if len(uniq) == 1:
            plan["changes"].append({
                "sheet": dst.sheet, "ri": ri, "excel_row": dst._excel_rows[ri],
                "col": col, "colname": c["colname"], "old": dst.data_rows[ri][col],
                "new": uniq[0]["value"], "key": c["key"], "src": uniq[0]["src"],
            })
        else:
            c["type"] = "cell"
            c["options"] = uniq
            c["old"] = dst.data_rows[ri][col]
            c["excel_row"] = dst._excel_rows[ri]
            plan["conflicts"].append(c)

    # ---- 新增行归并：同键多源内容不同 → 冲突 ----
    for key_str, nc in new_cands.items():
        opts = nc["opts"]
        if len(opts) == 1:
            plan["new_rows"].append({"src": opts[0]["src"], "key": key_str,
                                     "values": opts[0]["values"]})
            continue
        sig = set()
        for o in opts:
            sig.add(tuple(sorted((k, cell_text(v)) for k, v in o["values"].items())))
        if len(sig) == 1:
            plan["new_rows"].append({"src": opts[0]["src"], "key": key_str,
                                     "values": opts[0]["values"]})
        else:
            nc["type"] = "newrow"
            nc["options"] = opts
            plan["conflicts"].append(nc)

    for key in dst_idx:
        if key not in matched_dst_keys and any(key):
            ri = dst_idx[key][0]
            plan["unmatched_dst"].append((" + ".join(key), dst._excel_rows[ri]))

    plan["_meta"] = {"dst_rows": len(dst.data_rows),
                     "src_rows": sum(len(s.data_rows) for s in sources),
                     "matched": len(matched_dst_keys)}
    plan["stats"] = fill_stats(plan)
    return plan


def fill_stats(plan):
    m = plan.get("_meta", {})
    return {
        "目标表数据行": m.get("dst_rows", 0),
        "源表数据行": m.get("src_rows", 0),
        "匹配成功行": m.get("matched", 0),
        "将修改单元格": len(plan["changes"]),
        "将新增行": len(plan["new_rows"]),
        "待确认冲突": len(plan["conflicts"]),
        "同键异名": len(plan.get("name_mismatch", [])),
        "目标表未匹配行": len(plan["unmatched_dst"]),
        "源表未匹配行": len(plan["unmatched_src"]),
        "目标表重复键": len(plan["dup_dst"]),
        "源表重复键": len(plan["dup_src"]),
    }


def resolve_conflicts(plan, choices):
    """按用户选择解决冲突：choices = {(type, ri/键, col): 候选下标}"""
    for c in plan["conflicts"]:
        if c["type"] == "cell":
            k = ("cell", c["ri"], c["col"])
        else:
            k = ("newrow", c["key"])
        idx = choices.get(k, 0)
        opts = c["options"]
        if not isinstance(idx, int) or idx < 0 or idx >= len(opts):
            idx = 0
        o = opts[idx]
        plan.setdefault("resolved", []).append({
            "sheet": plan.get("_sheet") or "",
            "type": "单元格" if c["type"] == "cell" else "新增行",
            "where": c.get("excel_row", ""),
            "colname": c.get("colname", ""),
            "key": c["key"],
            "src": o["src"],
            "value": cell_text(o.get("value")) if c["type"] == "cell"
                     else ", ".join("%s=%s" % (k, cell_text(v)) for k, v in o["values"].items()),
        })
        if c["type"] == "cell":
            plan["changes"].append({
                "sheet": plan.get("_sheet") or "", "ri": c["ri"],
                "excel_row": c["excel_row"], "col": c["col"], "colname": c["colname"],
                "old": c["old"], "new": o["value"], "key": c["key"], "src": o["src"],
            })
        else:
            plan["new_rows"].append({"src": o["src"], "key": c["key"],
                                     "values": o["values"]})
    plan["conflicts"] = []
    plan["stats"] = fill_stats(plan)
    return plan


def _copy_row_style(ws, from_row, to_row, ncols):
    for c in range(1, ncols + 1):
        s, d = ws.cell(row=from_row, column=c), ws.cell(row=to_row, column=c)
        if s.has_style:
            d._style = s._style


def dir_writable(d):
    """实际试写一个临时文件，判断这个目录能不能写（比看属性可靠）"""
    try:
        d = d or "."
        if not os.path.isdir(d):
            return False          # 目录还不存在，交给写入时创建，不在这里乱建目录
        probe = os.path.join(d, "_wtest_%s.tmp" % datetime.datetime.now().strftime("%H%M%S%f"))
        with open(probe, "wb") as f:
            f.write(b"x")
        os.remove(probe)
        return True
    except Exception:
        return False


def fallback_dirs():
    """按优先级给出几个“大概能写”的备选目录"""
    out = []
    for p in (os.path.join(os.path.expanduser("~"), "Desktop"),
              os.path.join(os.path.expanduser("~"), "Documents"),
              os.path.join(os.path.expanduser("~"), "Downloads"),
              os.environ.get("TEMP", ""),
              os.getcwd()):
        if p and os.path.isdir(p) and p not in out:
            out.append(p)
    return out


def _last_data_row(ws, first_row, ncols):
    """从下往上找真正有数据的最后一行（跳过模板底部预格式化的空白行）"""
    r = ws.max_row
    while r >= first_row:
        for c in range(1, ncols + 1):
            if cell_text(ws.cell(row=r, column=c).value) != "":
                return r
        r -= 1
    return first_row - 1


def execute_fill_multi(target_path, plans, out_path):
    """把多个工作表的补全计划一次写入（复制目标文件后就地改单元格，保留格式/公式）

    plans: [(目标表 SheetData, 补全计划), ...]
    """
    target_path = os.path.abspath(target_path)
    out_path = os.path.abspath(out_path)
    if out_path == target_path:
        bak = target_path + ".bak_%s" % datetime.datetime.now().strftime("%H%M%S")
        shutil.copy2(target_path, bak)
    else:
        os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
        # 目标文件可能正被 Excel 打开（独占/只读），或被安全软件临时锁定：
        # 换一个可写的文件名再写，避免直接报“拒绝访问”
        base, ext = os.path.splitext(out_path)
        for i in range(6):
            cand = out_path if i == 0 else "%s(%d)%s" % (base, i + 1, ext)
            try:
                shutil.copy2(target_path, cand)
                out_path = cand
                break
            except (PermissionError, OSError) as e:
                if i == 5:
                    raise PermissionError(
                        "无法写入：%s\n（可能该文件正在 Excel 中打开，或文件夹只读）\n%s"
                        % (out_path, e))

    wb = openpyxl.load_workbook(out_path)
    for dst, plan in plans:
        ws = wb[dst.sheet]
        for ch in plan["changes"]:
            ws.cell(row=ch["excel_row"], column=ch["col"] + 1).value = ch["new"]
        if plan["new_rows"]:
            # 从“最后一个有数据的行”下面开始写，避免写到模板底部 100 多行之外
            last = _last_data_row(ws, dst.first_data_row, dst.n_cols)
            r = max(last, dst.first_data_row - 1)
            for nr in plan["new_rows"]:
                r += 1
                for c in range(1, dst.n_cols + 1):
                    ws.cell(row=r, column=c).value = nr["values"].get(c - 1)
                # 目标行本身有格式（模板线框）就保留；否则照抄上一行
                try:
                    if not ws.cell(row=r, column=1).has_style and last >= dst.first_data_row:
                        _copy_row_style(ws, last, r, dst.n_cols)
                except Exception:
                    pass
    wb.save(out_path)
    return out_path


def execute_fill(dst, plan, out_path, n_cols=None):
    """单表补全（兼容旧接口）"""
    return execute_fill_multi(dst.path, [(dst, plan)], out_path)


def export_report_multi(plans, out_path):
    """导出明细清单（支持多工作表）"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "修改明细"
    ws.append(["工作表", "目标表行号", "列名", "原值", "新值", "匹配键", "来源"])
    for dst, plan in plans:
        for ch in plan["changes"]:
            ws.append([dst.sheet, ch["excel_row"], ch["colname"], cell_text(ch["old"]),
                       cell_text(ch["new"]), ch["key"], ch["src"]])

    ws2 = wb.create_sheet("新增行")
    ws2.append(["工作表", "匹配键", "来源", "写入内容", "导入结果"])
    for dst, plan in plans:
        for nr in plan["new_rows"]:
            ws2.append([dst.sheet, nr["key"], nr["src"],
                        json.dumps({dst.headers[k]: cell_text(v)
                                    for k, v in nr["values"].items()}, ensure_ascii=False),
                        "已追加"])

    ws3 = wb.create_sheet("目标表未匹配")
    ws3.append(["工作表", "匹配键", "目标表行号"])
    for dst, plan in plans:
        for k, r in plan["unmatched_dst"]:
            ws3.append([dst.sheet, k, r])

    ws4 = wb.create_sheet("源表未匹配")
    ws4.append(["工作表", "来源", "匹配键", "首个键值"])
    for dst, plan in plans:
        for row in plan["unmatched_src"]:
            ws4.append([dst.sheet] + list(row))

    ws5 = wb.create_sheet("重复键")
    ws5.append(["工作表", "位置", "匹配键", "出现次数"])
    for dst, plan in plans:
        for k, v in plan["dup_dst"]:
            ws5.append([dst.sheet, "目标表", " + ".join(k), len(v)])
        for lab, k, n in plan["dup_src"]:
            ws5.append([dst.sheet, lab, " + ".join(k), n])

    nm_rows = []
    for dst, plan in plans:
        for it in plan.get("name_mismatch", []):
            nm_rows.append((dst.sheet,) + tuple(it))
    if nm_rows:
        wsn = wb.create_sheet("同键异名（疑似错行）")
        wsn.append(["工作表", "匹配键", "源表姓名", "目标表姓名", "来源"])
        for r in nm_rows:
            wsn.append(list(r))

    res_rows = []
    for dst, plan in plans:
        res_rows.extend(plan.get("resolved", []))
    if res_rows:
        ws6 = wb.create_sheet("源表冲突（已选）")
        ws6.append(["工作表", "类型", "行号", "列名", "匹配键", "采用的源表", "采用的值"])
        for r in res_rows:
            ws6.append([r["sheet"], r["type"], r["where"], r["colname"],
                        r["key"], r["src"], r["value"]])

    for w in wb.worksheets:
        for cell in w[1]:
            cell.font = Font(bold=True)
        w.freeze_panes = "A2"
        for col in w.columns:
            width = max([len(cell_text(c.value)) for c in col] or [8])
            w.column_dimensions[col[0].column_letter].width = min(60, max(10, width + 2))
    wb.save(out_path)
    return out_path


def export_report(dst, plan, out_path):
    return export_report_multi([(dst, plan)], out_path)


# ============================================================
#                        四、功能②：两表对账
# ============================================================
def plan_diff(a, b, ka, kb, kopt):
    a_idx = build_key_index(a, ka, kopt)
    b_idx = build_key_index(b, kb, kopt)
    only_a = [(k, a_idx[k][0]) for k in a_idx if k not in b_idx]
    only_b = [(k, b_idx[k][0]) for k in b_idx if k not in a_idx]
    pairs = []
    a_norm = {norm_colname(h): i for i, h in enumerate(a.headers)}
    for bi, h in enumerate(b.headers):
        ai = a_norm.get(norm_colname(h))
        if ai is not None:
            pairs.append((ai, bi, h))
    diffs = []
    for k in a_idx:
        if k not in b_idx:
            continue
        arow = a.data_rows[a_idx[k][0]]
        brow = b.data_rows[b_idx[k][0]]
        for ai, bi, h in pairs:
            av, bv = arow[ai], brow[bi]
            if norm_value(av, kopt) != norm_value(bv, kopt):
                diffs.append((" + ".join(k), h, cell_text(av), cell_text(bv)))
    return {"only_a": only_a, "only_b": only_b, "diffs": diffs,
            "cmp_cols": [h for _, _, h in pairs],
            "stats": {"A行数": len(a.data_rows), "B行数": len(b.data_rows),
                      "可比字段": len(pairs), "仅A有": len(only_a),
                      "仅B有": len(only_b), "字段不一致": len(diffs)}}


def export_diff(a, b, plan, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "仅A表有"
    ws.append(["匹配键"] + a.headers)
    for k, ri in plan["only_a"]:
        ws.append([k] + [cell_text(v) for v in a.data_rows[ri]])
    ws2 = wb.create_sheet("仅B表有")
    ws2.append(["匹配键"] + b.headers)
    for k, ri in plan["only_b"]:
        ws2.append([k] + [cell_text(v) for v in b.data_rows[ri]])
    ws3 = wb.create_sheet("字段不一致")
    ws3.append(["匹配键", "列名", "A表值", "B表值"])
    for row in plan["diffs"]:
        ws3.append(list(row))
    for w in wb.worksheets:
        for cell in w[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="DDEBF7")
        w.freeze_panes = "A2"
        for col in w.columns:
            width = max(len(cell_text(c.value)) for c in col) if col else 8
            w.column_dimensions[col[0].column_letter].width = min(60, max(10, width + 2))
    wb.save(out_path)
    return out_path


# ============================================================
#                        五、功能③：多表汇总
# ============================================================
def plan_merge(tables, dedupe_cols=None, kopt=None, add_source=True, empty_key_skip=True):
    kopt = kopt or {}
    dedupe_cols = dedupe_cols or []
    cols, seen = [], {}
    for t in tables:
        for i, h in enumerate(t.headers):
            k = norm_colname(h)
            if k and k not in seen:
                seen[k] = len(cols)
                cols.append(h)
    rows, dup, empty = [], 0, 0
    keyset = set()
    for t in tables:
        for r in t.data_rows:
            vals = [None] * len(cols)
            for i, h in enumerate(t.headers):
                pos = seen.get(norm_colname(h))
                if pos is not None and i < len(r):
                    vals[pos] = r[i]
            if all(cell_text(v) == "" for v in vals):
                empty += 1
                continue
            if dedupe_cols:
                key = []
                for dc in dedupe_cols:
                    pos = seen.get(norm_colname(dc))
                    key.append(norm_value(vals[pos], kopt) if pos is not None else "")
                key = tuple(key)
                if any(key):
                    if key in keyset:
                        dup += 1
                        continue
                    keyset.add(key)
            if add_source:
                vals.append(t.label)
            rows.append(vals)
    if add_source:
        cols.append("来源表")
    return {"cols": cols, "rows": rows,
            "stats": {"表数量": len(tables), "输出列数": len(cols), "输出行数": len(rows),
                      "跳过空行": empty, "去重删除行": dup}}


def export_merge(plan, out_path, sheet_name="汇总"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(plan["cols"])
    for r in plan["rows"]:
        ws.append([cell_text(v) for v in r])
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDEBF7")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    for i, h in enumerate(plan["cols"], start=1):
        width = max([len(cell_text(c.value)) for c in ws[get_column_letter(i)]] or [8])
        ws.column_dimensions[get_column_letter(i)].width = min(60, max(10, width + 2))
    wb.save(out_path)
    return out_path


def now_stamp():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


def _settings_path():
    try:
        base = os.path.dirname(os.path.abspath(
            sys.executable if getattr(sys, "frozen", False) else __file__))
    except Exception:
        base = os.getcwd()
    return os.path.join(base, "设置.json")


def load_settings():
    try:
        with open(_settings_path(), encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def save_settings(d):
    try:
        with open(_settings_path(), "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def enable_dpi_awareness():
    """高分屏下让界面清晰（不做 DPI 感知会被系统拉伸导致发糊）"""
    if sys.platform != "win32":
        return
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)   # PER_MONITOR_DPI_AWARE
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


# ============================================================
#                        六、界面
# ============================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        # ---- 跟随分辨率 + 高分屏清晰显示 ----
        dpi = float(self.winfo_fpixels("1i"))
        self.scale = max(1.0, dpi / 96.0)
        try:
            self.tk.call("tk", "scaling", dpi / 72.0)
        except Exception:
            pass
        import tkinter.font as tkfont
        for fname in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                tkfont.nametofont(fname, root=self).configure(family="Microsoft YaHei UI")
            except Exception:
                pass
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w = int(min(max(sw * 0.86, 1024), sw * 0.96))
        h = int(min(max(sh * 0.84, 680), sh * 0.94))
        self.geometry("%dx%d+%d+%d" % (w, h, max(0, (sw - w) // 2), max(0, (sh - h) // 2 - 10)))
        self.minsize(int(min(1000, sw * 0.7)), int(min(660, sh * 0.6)))

        st = ttk.Style(self)
        try:
            st.theme_use("vista")
        except Exception:
            pass
        base = 10 if self.scale < 1.5 else 11
        st.configure("Title.TLabel", font=("Microsoft YaHei UI", base + 1, "bold"))
        st.configure("Hint.TLabel", foreground="#666666")
        st.configure("Go.TButton", font=("Microsoft YaHei UI", base, "bold"))

        # 功能①状态
        self.sources = []
        self.dst = None            # 基准目标表（用于列名下拉/匹配键）
        self.dst_sheets = []       # 目标工作簿全部工作表名
        self.dst_cache = {}        # (sheet, hdr, span) -> SheetData
        self.pairs = []            # [{"src": 源表下标, "target": 目标工作表名}]
        self.mappings = []         # [{src,dst,mode}]
        self.key_pairs = []        # [(源列名, 目标列名)]
        self.last_plans = []
        self.conflict_cache = {}   # 记住用户对同类冲突的选择，下次自动采用
        self.settings = load_settings()
        self.var_font = tk.IntVar(value=int(self.settings.get("font_size", 10) or 10))

        # ---- 顶部工具条：字体大小 ----
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=10, pady=(8, 0))
        ttk.Label(bar, text="字体大小", style="Title.TLabel").pack(side="left")
        self.sp_font = ttk.Spinbox(bar, from_=8, to=22, width=4, textvariable=self.var_font)
        self.sp_font.pack(side="left", padx=6)
        self._font_job = None
        self._applying_font = False
        self.var_font.trace_add("write", self._font_changed)
        ttk.Button(bar, text="恢复默认", width=9,
                   command=lambda: self.var_font.set(10)).pack(side="left")
        ttk.Label(bar, text="（改数字/点箭头即时生效，设置会自动记住）",
                  style="Hint.TLabel").pack(side="left", padx=6)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.tab_fill = ttk.Frame(nb, padding=6)
        self.tab_ledger = ttk.Frame(nb, padding=6)
        nb.add(self.tab_fill, text="  信息补全  ")
        nb.add(self.tab_ledger, text="  台账合并  ")

        self.build_fill()
        self.build_ledger()
        self.apply_font_size(self.var_font.get(), save=False)

    # ---------- 字体大小（实时生效，无需“应用”按钮） ----------
    def _font_changed(self, *args):
        if self._applying_font:
            return
        try:
            size = int(str(self.var_font.get()).strip())
        except Exception:
            return
        if not (6 <= size <= 24):
            return
        if self._font_job:
            try:
                self.after_cancel(self._font_job)
            except Exception:
                pass
        self._font_job = self.after(120, lambda: self.apply_font_size(size))

    def on_font_change(self):
        try:
            size = int(self.var_font.get())
        except Exception:
            size = 10
        self.apply_font_size(size)

    def apply_font_size(self, size, save=True):
        size = max(6, min(24, int(size)))
        self._applying_font = True
        self.var_font.set(size)
        self._applying_font = False
        import tkinter.font as tkfont
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont",
                     "TkTooltipFont", "TkCaptionFont"):
            try:
                tkfont.nametofont(name, root=self).configure(family="Microsoft YaHei UI", size=size)
            except Exception:
                pass
        st = ttk.Style(self)
        try:
            st.configure("Treeview", rowheight=int(size * 2.2))
            st.configure("Treeview.Heading", font=("Microsoft YaHei UI", size, "bold"))
            st.configure("Title.TLabel", font=("Microsoft YaHei UI", size + 1, "bold"))
            st.configure("Go.TButton", font=("Microsoft YaHei UI", size, "bold"))
        except Exception:
            pass
        try:
            self.txt_fill.configure(font=("Consolas", max(7, size - 1)))
        except Exception:
            pass
        self.settings["font_size"] = size
        if save:
            save_settings(self.settings)

    # ================= 功能②：台账合并（多村台账 → 县级总账） =================
    def build_ledger(self):
        root = self.tab_ledger
        self.led_files = []          # 源文件路径
        self.led_dst = None
        self.led_dst_cache = {}

        f1 = ttk.LabelFrame(root, text=" 1. 目标表 ", padding=6)
        f1.pack(fill="x")
        self.var_led_dst = tk.StringVar()
        self.file_row(f1, "目标文件", self.var_led_dst, self.led_pick_dst)

        r1 = ttk.Frame(f1)
        r1.pack(fill="x", pady=(4, 0))
        ttk.Label(r1, text="要合并到哪些工作表",
                  style="Hint.TLabel").pack(anchor="w")
        lb1 = ttk.Frame(r1)
        lb1.pack(fill="x")
        self.lb_led_dst = tk.Listbox(lb1, height=5, selectmode="extended", exportselection=False)
        sb = ttk.Scrollbar(lb1, orient="vertical", command=self.lb_led_dst.yview)
        self.lb_led_dst.configure(yscrollcommand=sb.set)
        self.lb_led_dst.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.lb_led_dst.bind("<<ListboxSelect>>", lambda e: self.led_sync_sel())

        # 源文件 + 源工作表
        f2 = ttk.LabelFrame(root, text=" 2. 源表 ", padding=6)
        f2.pack(fill="x", pady=(6, 0))
        lb2 = ttk.Frame(f2)
        lb2.pack(fill="x")
        self.lb_led_src = tk.Listbox(lb2, height=4, selectmode="extended", exportselection=False)
        sb2 = ttk.Scrollbar(lb2, orient="vertical", command=self.lb_led_src.yview)
        self.lb_led_src.configure(yscrollcommand=sb2.set)
        self.lb_led_src.pack(side="left", fill="both", expand=True)
        sb2.pack(side="right", fill="y")
        bar2 = ttk.Frame(f2)
        bar2.pack(fill="x", pady=(4, 0))
        ttk.Button(bar2, text="添加文件（可多选）…", command=self.led_add_files).pack(side="left")
        ttk.Button(bar2, text="移除选中", command=self.led_remove).pack(side="left", padx=4)
        ttk.Button(bar2, text="清空", command=self.led_clear).pack(side="left")
        self.lb_led_info = ttk.Label(f2, text="尚未添加源文件", style="Hint.TLabel")
        self.lb_led_info.pack(fill="x", pady=(3, 0))

        r2b = ttk.Frame(f2)
        r2b.pack(fill="x", pady=(4, 0))
        ttk.Label(r2b, text="要参与合并的源表工作表（默认全选，按住 Ctrl 可取消某几个）",
                  style="Hint.TLabel").pack(side="left")
        ttk.Button(r2b, text="全选", width=6, command=self.led_sel_all).pack(side="right")
        ttk.Button(r2b, text="全不选", width=7, command=self.led_sel_none).pack(side="right", padx=4)
        lb3 = ttk.Frame(f2)
        lb3.pack(fill="x")
        self.lb_led_sheets = tk.Listbox(lb3, height=8, selectmode="extended", exportselection=False)
        sb3 = ttk.Scrollbar(lb3, orient="vertical", command=self.lb_led_sheets.yview)
        self.lb_led_sheets.configure(yscrollcommand=sb3.set)
        self.lb_led_sheets.pack(side="left", fill="both", expand=True)
        sb3.pack(side="right", fill="y")
        self.led_items = []      # [(path, sheet, 显示名)]

        # 选项
        f3 = ttk.LabelFrame(root, text=" 3. 合并选项 ", padding=6)
        f3.pack(fill="x", pady=(6, 0))
        self.var_led_inherit = tk.BooleanVar(value=True)
        self.var_led_suffix = tk.BooleanVar(value=True)
        self.var_led_seq = tk.BooleanVar(value=True)
        self.var_led_keep = tk.BooleanVar(value=True)
        ttk.Checkbutton(f3, text="空白继承上一行（同户多人只填一次户主信息）",
                        variable=self.var_led_inherit).pack(anchor="w")
        ttk.Checkbutton(f3, text="自动统一镇/村名称（万和→万和镇、曹门→曹门村）",
                        variable=self.var_led_suffix).pack(anchor="w")
        ttk.Checkbutton(f3, text="序号自动重排（按目标表现有行数续编）",
                        variable=self.var_led_seq).pack(anchor="w")
        ttk.Checkbutton(f3, text="保留原文件（另存为新文件）",
                        variable=self.var_led_keep).pack(anchor="w")
        self.var_led_out = tk.StringVar()
        self.file_row(f3, "输出文件", self.var_led_out, self.led_pick_out)

        f4 = ttk.Frame(root)
        f4.pack(fill="x", pady=(6, 0))
        ttk.Button(f4, text="① 预演（不写文件）", command=self.led_preview,
                   style="Go.TButton").pack(side="left")
        ttk.Button(f4, text="② 执行并保存", command=self.led_execute,
                   style="Go.TButton").pack(side="left", padx=6)
        ttk.Label(f4, text="合并规则：按列名对齐，源表每一行追加到目标表末尾",
                  style="Hint.TLabel").pack(side="left", padx=8)

        self.txt_led = self.log_widget(root, height=10)

    def led_pick_dst(self):
        p = self.ask_open("led_dst", title="选择目标（县级）动态调整表",
                                       filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("全部", "*.*")])
        if not p:
            return
        self.var_led_dst.set(p)
        if not self.settings.get("dir_led_src"):
            self.save_dir("led_src", p)
        self.led_dst_cache = {}
        names = list_sheets(p)
        self.lb_led_dst.delete(0, "end")
        for n in names:
            self.lb_led_dst.insert("end", n)
        self.lb_led_dst.selection_set(0, "end")      # 默认全选所有工作表
        self.led_dst = None
        self.led_sync_sel()
        self.log(self.txt_led, "目标文件：%s（%d 个工作表，已全选）"
                 % (os.path.basename(p), len(names)))

    def led_add_files(self):
        paths = self.ask_open_multi("led_src", title="选择各村的动态调整表（可多选）",
                                            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("全部", "*.*")])
        if not paths:
            return
        for p in paths:
            if p not in self.led_files:
                self.led_files.append(p)
        self.led_refresh()

    def led_remove(self):
        for i in reversed(list(self.lb_led_src.curselection())):
            self.led_files.pop(i)
        self.led_refresh()

    def led_clear(self):
        self.led_files = []
        self.led_refresh()

    def led_refresh(self):
        """刷新源文件列表，并重建“源表工作表”清单（默认全选）"""
        self.lb_led_src.delete(0, "end")
        self.led_items = []
        for p in self.led_files:
            try:
                names = list_sheets(p)
            except Exception:
                names = []
            self.lb_led_src.insert("end", "%s   [%s]" % (
                os.path.basename(p),
                "、".join(n for n in names if "汇总" not in n) or "无工作表"))
            for n in names:
                self.led_items.append((p, n, "%s / %s" % (os.path.basename(p), n)))
        self.lb_led_sheets.delete(0, "end")
        for _, _, lab in self.led_items:
            self.lb_led_sheets.insert("end", lab)
        self.led_sync_sel()
        self.lb_led_info.configure(
            text="已添加 %d 个源文件，共 %d 个工作表" % (len(self.led_files), len(self.led_items))
            if self.led_files else "尚未添加源文件")

    def led_sync_sel(self):
        """让“源表工作表”的选中状态跟随“目标表工作表”的选中状态"""
        if not getattr(self, "led_items", None):
            return
        try:
            targets = [self.lb_led_dst.get(i) for i in self.lb_led_dst.curselection()]
        except Exception:
            return
        self.lb_led_sheets.selection_clear(0, "end")
        if not targets:
            return
        for i, (path, sname, _lab) in enumerate(self.led_items):
            if any(self._sheet_match(sname, t) for t in targets):
                self.lb_led_sheets.selection_set(i)

    def led_sel_all(self):
        if self.led_items:
            self.lb_led_sheets.selection_set(0, "end")

    def led_sel_none(self):
        self.lb_led_sheets.selection_clear(0, "end")

    def led_selected_items(self):
        return [self.led_items[i] for i in self.lb_led_sheets.curselection()]

    def led_pick_out(self):
        p = self.ask_save("out", defaultextension=".xlsx",
                          filetypes=[("Excel", "*.xlsx")])
        if p:
            self.var_led_out.set(p)

    def led_get_dst(self, sheet):
        p = self.var_led_dst.get().strip()
        if sheet not in self.led_dst_cache:
            self.led_dst_cache[sheet] = open_table(p, sheet)
        return self.led_dst_cache[sheet]

    @staticmethod
    def _match_sheet_name(target, names):
        for n in names:
            if norm_colname(n) == norm_colname(target):
                return n
        a = norm_colname(target)
        for n in names:
            b = norm_colname(n)
            if a and b and (a in b or b in a):
                return n
        return None

    @staticmethod
    def _sheet_match(name, target):
        return App._match_sheet_name(target, [name]) is not None

    def led_make_plans(self):
        if not self.var_led_dst.get().strip():
            messagebox.showwarning("提示", "请先选择目标表")
            return None
        sheets = [self.lb_led_dst.get(i) for i in self.lb_led_dst.curselection()]
        if not sheets:
            messagebox.showwarning("提示", "请选择要合并到哪些工作表")
            return None
        sel = self.led_selected_items()
        if not sel:
            messagebox.showwarning("提示", "请选择要参与合并的源表工作表")
            return None
        plans, warns, wbs = [], [], {}
        try:
            for path in sorted({x[0] for x in sel}):
                try:
                    wbs[path] = openpyxl.load_workbook(path, data_only=True)
                except Exception as e:
                    warns.append("%s 打不开：%s" % (os.path.basename(path), e))
            for sheet in sheets:
                dst = self.led_get_dst(sheet)
                srcs, labels = [], []
                for path, sname, _lab in sel:
                    wb = wbs.get(path)
                    if wb is None or not self._sheet_match(sname, sheet):
                        continue
                    try:
                        sd = SheetData(path, sname, wb=wb)
                    except Exception as e:
                        warns.append("%s / %s 读取失败：%s" % (os.path.basename(path), sname, e))
                        continue
                    if not sd.data_rows:
                        continue
                    t = sd
                    if self.var_led_inherit.get():
                        t = inherit_blank_rows(t, None)
                    srcs.append(t)
                    labels.append(os.path.basename(path))
                if not srcs:
                    warns.append("「%s」没有选中的源表有数据，已跳过" % sheet)
                    continue
                if self.var_led_suffix.get():
                    rules = (auto_suffix_rules_multi(srcs) + auto_suffix_fill_rules(srcs))
                    if rules:
                        srcs = [standardize_table(x, rules) for x in srcs]
                        warns.append("「%s」已自动统一名称 %d 条规则" % (sheet, len(rules)))
                base = srcs[0]
                mp = []
                for i, h in enumerate(base.headers):
                    if norm_colname(h) in ("序号", "编号", "序"):
                        continue
                    di = self._resolve(dst, h)
                    if di is not None:
                        mp.append({"src_name": h, "dst": di})
                if not mp:
                    warns.append("「%s」源表与目标表没有可对应的列，已跳过" % sheet)
                    continue
                seq = None
                if self.var_led_seq.get():
                    for i, h in enumerate(dst.headers):
                        if norm_colname(h) in ("序号", "编号", "序"):
                            seq = i
                            break
                plan = plan_append(srcs, dst, mp, seq)
                plan["_labels"] = labels
                plans.append((dst, plan))
        finally:
            for wb in wbs.values():
                try:
                    wb.close()
                except Exception:
                    pass
        return plans, warns

    def led_preview(self):
        got = self.led_make_plans()
        if not got:
            return
        plans, warns = got
        if not plans:
            messagebox.showinfo("提示", "没有可合并的内容")
            return
        self.log(self.txt_led, "———— 预演（%d 个工作表） ————" % len(plans))
        tot = 0
        for dst, plan in plans:
            n = len(plan["new_rows"])
            tot += n
            self.log(self.txt_led, "  「%s」目标现有 %d 行 → 将追加 %d 行（来自 %d 个文件）"
                     % (dst.sheet, len(dst.data_rows), n, len(plan.get("_labels", []))))
        for w in warns:
            self.log(self.txt_led, "  ⚠ " + w)
        self.log(self.txt_led, "  合计追加 %d 行" % tot)
        self.led_preview_window(plans)

    def led_preview_window(self, plans):
        win = tk.Toplevel(self)
        win.title("台账合并预演（未写文件）")
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        win.geometry("%dx%d" % (int(sw * 0.62), int(sh * 0.6)))
        nb = ttk.Notebook(win)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        for dst, plan in plans:
            f = ttk.Frame(nb)
            cols = [h for h in dst.headers]
            tv = ttk.Treeview(f, columns=[str(i) for i in range(len(cols))], show="headings")
            for i, h in enumerate(cols):
                tv.heading(str(i), text=h.replace("\n", ""))
                tv.column(str(i), width=110)
            for nr in plan["new_rows"][:300]:
                vals = [cell_text(nr["values"].get(i))[:26] for i in range(len(cols))]
                tv.insert("", "end", values=vals)
            sb = ttk.Scrollbar(f, orient="vertical", command=tv.yview)
            tv.configure(yscrollcommand=sb.set)
            tv.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")
            nb.add(f, text="%s (+%d)" % (dst.sheet, len(plan["new_rows"])))
        ttk.Button(win, text="关闭", command=win.destroy).pack(pady=6)

    def led_execute(self):
        got = self.led_make_plans()
        if not got:
            return
        plans, warns = got
        if not plans or not any(p["new_rows"] for _, p in plans):
            messagebox.showinfo("提示", "没有可合并的内容")
            return
        src_path = self.var_led_dst.get().strip()
        stem, ext = os.path.splitext(src_path)
        suggest = self.var_led_out.get().strip() or "%s_合并结果_%s%s" % (stem, now_stamp(), ext)
        if not self.var_led_keep.get():
            suggest = src_path
        out = self.ask_output_path(suggest, src_path, key="led_out")
        if not out:
            return
        self.var_led_out.set(out)
        try:
            final = execute_fill_multi(src_path, plans, out)
            rep = os.path.splitext(final)[0] + "_合并明细.xlsx"
            export_report_multi(plans, rep)
        except PermissionError as e:
            messagebox.showerror("写入失败（文件被占用）",
                                 "%s\n\n请这样处理：\n"
                                 "1) 关掉 Excel 里已打开的同名文件，重新点“执行并保存”；\n"
                                 "2) 或者把“输出文件”改到别的目录/别的文件名；\n"
                                 "3) 若文件是只读，右键 → 属性 → 取消“只读”。" % e)
            return
        except Exception:
            messagebox.showerror("写入失败", traceback.format_exc())
            return
        n = sum(len(p["new_rows"]) for _, p in plans)
        for dst, plan in plans:
            self.log(self.txt_led, "  ✔ 「%s」追忆 %d 行" % (dst.sheet, len(plan["new_rows"])))
        self.log(self.txt_led, "✔ 已写入：%s（共追加 %d 行）" % (final, n))
        self.log(self.txt_led, "✔ 明细：%s" % rep)
        messagebox.showinfo("完成", "已保存：\n%s\n\n共追加 %d 行" % (final, n))

    # ---------- 公共小部件 ----------
    def mem_dir(self, key):
        d = self.settings.get("dir_" + key) if isinstance(self.settings, dict) else None
        if d and os.path.isdir(d):
            return d
        return os.path.expanduser("~")

    def save_dir(self, key, path):
        if not path:
            return
        d = path if os.path.isdir(path) else os.path.dirname(path)
        if d and os.path.isdir(d):
            self.settings["dir_" + key] = d
            save_settings(self.settings)

    def ask_open(self, key, **kw):
        kw.setdefault("initialdir", self.mem_dir(key))
        p = filedialog.askopenfilename(**kw)
        if p:
            self.save_dir(key, p)
        return p

    def ask_open_multi(self, key, **kw):
        kw.setdefault("initialdir", self.mem_dir(key))
        ps = filedialog.askopenfilenames(**kw)
        if ps:
            self.save_dir(key, ps[0])
        return ps

    def ask_save(self, key, **kw):
        kw.setdefault("initialdir", self.mem_dir(key))
        p = filedialog.asksaveasfilename(**kw)
        if p:
            self.save_dir(key, p)
        return p

    def file_row(self, parent, label, var, cmd, width=14):
        fr = ttk.Frame(parent)
        fr.pack(fill="x", pady=2)
        ttk.Label(fr, text=label, width=width).pack(side="left")
        ttk.Entry(fr, textvariable=var).pack(side="left", fill="x", expand=True)
        ttk.Button(fr, text="浏览…", width=7, command=cmd).pack(side="left", padx=4)
        return fr

    def log_widget(self, parent, height=10, use_grid=False):
        box = ttk.Frame(parent)
        if use_grid:
            box.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(6, 0))
        else:
            box.pack(fill="both", expand=True, pady=(6, 0))
        t = tk.Text(box, height=height, wrap="none", font=("Consolas", 9))
        sb = ttk.Scrollbar(box, orient="vertical", command=t.yview)
        t.configure(yscrollcommand=sb.set, state="disabled")
        t.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        return t

    def log(self, widget, msg):
        widget.configure(state="normal")
        widget.insert("end", msg + "\n")
        widget.see("end")
        widget.configure(state="disabled")
        self.update_idletasks()

    # ================= 功能①：信息补全 =================
    def build_fill(self):
        root = self.tab_fill
        root.columnconfigure(0, weight=3)
        root.columnconfigure(1, weight=2)
        root.rowconfigure(0, weight=1)

        left = ttk.Frame(root)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        right = ttk.Frame(root)
        right.grid(row=0, column=1, sticky="nsew")

        # --- 1 源表 ---
        f1 = ttk.LabelFrame(left, text=" 1. 源表（数据来源，可添加多个文件 / 多个工作表） ", padding=6)
        f1.pack(fill="x")
        lb_box = ttk.Frame(f1)
        lb_box.pack(fill="x")
        self.lb_src = tk.Listbox(lb_box, height=5, selectmode="extended", exportselection=False,
                                 activestyle="dotbox")
        sb1 = ttk.Scrollbar(lb_box, orient="vertical", command=self.lb_src.yview)
        self.lb_src.configure(yscrollcommand=sb1.set)
        self.lb_src.pack(side="left", fill="both", expand=True)
        sb1.pack(side="right", fill="y")
        self.lb_src.bind("<Double-1>", lambda e: self.edit_source_header())

        bar = ttk.Frame(f1)
        bar.pack(fill="x", pady=(4, 0))
        ttk.Button(bar, text="添加文件…", command=self.add_source_file).pack(side="left")
        ttk.Button(bar, text="添加该文件的全部工作表", command=self.add_source_all_sheets).pack(side="left", padx=4)
        ttk.Button(bar, text="调整表头", command=self.edit_source_header).pack(side="left", padx=4)
        ttk.Button(bar, text="移除选中", command=self.remove_source).pack(side="left")
        ttk.Button(bar, text="清空", command=self.clear_sources).pack(side="left", padx=4)

        r1b = ttk.Frame(f1)
        r1b.pack(fill="x", pady=(3, 0))
        self.var_visible = tk.BooleanVar(value=False)
        ttk.Checkbutton(r1b, text="只处理可见行（表格被筛选时，忽略隐藏的行）",
                        variable=self.var_visible,
                        command=self.reload_sources).pack(side="left")
        ttk.Label(r1b, text="默认读取全部行（含被筛选隐藏的行）", style="Hint.TLabel").pack(side="left", padx=6)

        # --- 2 目标表 ---
        f2 = ttk.LabelFrame(left, text=" 2. 目标表（要被补全的表，工作表可多选） ", padding=6)
        f2.pack(fill="x", pady=(6, 0))
        self.var_dst_path = tk.StringVar()
        self.file_row(f2, "目标文件", self.var_dst_path, self.pick_dst_file)

        r2 = ttk.Frame(f2)
        r2.pack(fill="x", pady=2)
        lb2 = ttk.Frame(r2)
        lb2.pack(side="left", fill="x", expand=True)
        self.lb_dst = tk.Listbox(lb2, height=3, selectmode="extended", exportselection=False,
                                 activestyle="dotbox")
        sb2 = ttk.Scrollbar(lb2, orient="vertical", command=self.lb_dst.yview)
        self.lb_dst.configure(yscrollcommand=sb2.set)
        self.lb_dst.pack(side="left", fill="both", expand=True)
        sb2.pack(side="right", fill="y")
        self.lb_dst.bind("<<ListboxSelect>>", lambda e: self.on_dst_select())
        col2 = ttk.Frame(r2)
        col2.pack(side="left", fill="y", padx=(6, 0))
        ttk.Button(col2, text="全选", width=8, command=self.dst_select_all).pack(fill="x")
        ttk.Button(col2, text="反选", width=8, command=self.dst_invert).pack(fill="x", pady=2)
        ttk.Button(col2, text="清除", width=8, command=self.dst_clear_sel).pack(fill="x")

        r2b = ttk.Frame(f2)
        r2b.pack(fill="x", pady=2)
        ttk.Label(r2b, text="表头行", width=8).pack(side="left")
        self.var_hdr_row = tk.StringVar(value="自动")
        ttk.Spinbox(r2b, from_=1, to=50, width=5, textvariable=self.var_hdr_row).pack(side="left")
        ttk.Label(r2b, text=" 表头行数").pack(side="left")
        self.var_hdr_span = tk.StringVar(value="自动")
        ttk.Spinbox(r2b, from_=1, to=3, width=5, textvariable=self.var_hdr_span).pack(side="left")
        ttk.Button(r2b, text="重新识别", width=9, command=self.reload_dst).pack(side="left", padx=6)
        ttk.Label(r2b, text="自动＝逐表识别", style="Hint.TLabel").pack(side="left")

        r2c = ttk.Frame(f2)
        r2c.pack(fill="x", pady=(2, 0))
        ttk.Label(r2c, text="源表工作表 ↔ 目标表工作表", style="Title.TLabel").pack(side="left")
        ttk.Button(r2c, text="改目标", width=8, command=self.change_pair_target).pack(side="right")
        ttk.Button(r2c, text="自动对应", width=9, command=self.auto_correspond).pack(side="right", padx=4)

        self.tv_pairs = ttk.Treeview(f2, columns=("s", "t"), show="headings", height=4)
        self.tv_pairs.heading("s", text="源表工作表")
        self.tv_pairs.heading("t", text="目标表工作表")
        self.tv_pairs.column("s", width=180)
        self.tv_pairs.column("t", width=180)
        self.tv_pairs.pack(fill="x")
        self.tv_pairs.bind("<Double-1>", lambda e: self.change_pair_target())

        self.lb_dst_info = ttk.Label(f2, text="未加载目标表", style="Hint.TLabel")
        self.lb_dst_info.pack(fill="x", pady=(2, 0))

        # --- 3 匹配键 ---
        f3 = ttk.LabelFrame(left, text=" 3. 匹配键（两边都有的唯一标识，如身份证号） ", padding=6)
        f3.pack(fill="x", pady=(6, 0))
        r3 = ttk.Frame(f3)
        r3.pack(fill="x")
        ttk.Label(r3, text="源表列", width=7).pack(side="left")
        self.cb_keysrc = ttk.Combobox(r3, state="readonly", width=20)
        self.cb_keysrc.pack(side="left")
        ttk.Label(r3, text=" → 目标表列", width=11).pack(side="left")
        self.cb_keydst = ttk.Combobox(r3, state="readonly", width=20)
        self.cb_keydst.pack(side="left")
        ttk.Button(r3, text="自动匹配键", width=11, command=self.auto_key_pairs).pack(side="left", padx=4)
        ttk.Button(r3, text="添加键", width=8, command=self.add_key_pair).pack(side="left")

        tv_box = ttk.Frame(f3)
        tv_box.pack(fill="x", pady=4)
        self.tv_keys = ttk.Treeview(tv_box, columns=("s", "d"), show="headings", height=3)
        self.tv_keys.heading("s", text="源表列")
        self.tv_keys.heading("d", text="目标表列")
        self.tv_keys.column("s", width=200)
        self.tv_keys.column("d", width=200)
        self.tv_keys.pack(side="left", fill="x", expand=True)
        ttk.Button(tv_box, text="删除键", command=self.del_key_pair).pack(side="left", padx=4)

        r3b = ttk.Frame(f3)
        r3b.pack(fill="x")
        self.var_k_nospace = tk.BooleanVar(value=True)
        self.var_k_upper = tk.BooleanVar(value=True)
        self.var_k_symbol = tk.BooleanVar(value=False)
        self.var_k_digits = tk.BooleanVar(value=False)
        ttk.Label(r3b, text="键清洗：", style="Hint.TLabel").pack(side="left")
        ttk.Checkbutton(r3b, text="忽略空格", variable=self.var_k_nospace).pack(side="left")
        ttk.Checkbutton(r3b, text="忽略大小写", variable=self.var_k_upper).pack(side="left")
        ttk.Checkbutton(r3b, text="忽略符号", variable=self.var_k_symbol).pack(side="left")
        ttk.Checkbutton(r3b, text="只比对数字", variable=self.var_k_digits).pack(side="left")

        # --- 右侧 4 映射 ---
        f4 = ttk.LabelFrame(right, text=" 4. 字段映射（源列 → 目标列） ", padding=6)
        f4.pack(fill="both", expand=True)
        r4 = ttk.Frame(f4)
        r4.pack(fill="x")
        ttk.Label(r4, text="源列", width=5).pack(side="left")
        self.cb_map_src = ttk.Combobox(r4, state="readonly", width=18)
        self.cb_map_src.pack(side="left")
        ttk.Label(r4, text=" 目标列", width=7).pack(side="left")
        self.cb_map_dst = ttk.Combobox(r4, state="readonly", width=18)
        self.cb_map_dst.pack(side="left")
        r4b = ttk.Frame(f4)
        r4b.pack(fill="x", pady=3)
        ttk.Button(r4b, text="添加映射", command=self.add_mapping).pack(side="left")
        ttk.Button(r4b, text="自动匹配", command=self.auto_map).pack(side="left", padx=4)
        ttk.Button(r4b, text="删除选中", command=self.del_mapping).pack(side="left")
        ttk.Button(r4b, text="清空", command=self.clear_mapping).pack(side="left", padx=4)

        self.tv_map = ttk.Treeview(f4, columns=("s", "d", "m"), show="headings", height=7)
        for c, t, w in (("s", "源表列", 150), ("d", "目标表列", 150), ("m", "写入方式", 90)):
            self.tv_map.heading(c, text=t)
            self.tv_map.column(c, width=w)
        self.tv_map.pack(fill="both", expand=True)
        self.tv_map.bind("<Double-1>", lambda e: self.toggle_mapping_mode())
        ttk.Label(f4, text="双击一行可切换：默认 / 仅填空 / 覆盖 / 跳过", style="Hint.TLabel").pack(fill="x")

        # --- 5 选项 ---
        f5 = ttk.LabelFrame(right, text=" 5. 写入方式与输出 ", padding=6)
        f5.pack(fill="x", pady=(6, 0))
        self.var_mode = tk.StringVar(value="fill_add")
        for key, text in MODE_LABELS:
            ttk.Radiobutton(f5, text=text, value=key, variable=self.var_mode).pack(anchor="w")

        r5 = ttk.Frame(f5)
        r5.pack(fill="x", pady=(4, 0))
        self.var_keep = tk.BooleanVar(value=True)          # 默认保留原文件
        self.var_export_log = tk.BooleanVar(value=True)
        ttk.Checkbutton(r5, text="保留原文件（另存为新文件）",
                        variable=self.var_keep).pack(side="left")
        ttk.Checkbutton(r5, text="导出明细清单", variable=self.var_export_log).pack(side="left", padx=8)

        # 数据预处理：空白继承 + 值标准化
        self.var_inherit = tk.BooleanVar(value=False)
        self.var_std = tk.BooleanVar(value=False)
        self.inherit_cols = []        # 空 = 智能识别“同户共享列”
        self.replace_rules = []       # [{"col","old","new"}]
        r5b = ttk.Frame(f5)
        r5b.pack(fill="x", pady=(4, 0))
        ttk.Checkbutton(r5b, text="源表空白继承上一行（同户多人）",
                        variable=self.var_inherit).pack(side="left")
        ttk.Button(r5b, text="继承列…", width=8,
                   command=self.set_inherit_cols).pack(side="left", padx=4)
        r5c = ttk.Frame(f5)
        r5c.pack(fill="x")
        ttk.Checkbutton(r5c, text="按规则标准化字段值（如 万和→万和镇）",
                        variable=self.var_std).pack(side="left")
        ttk.Button(r5c, text="替换规则…", width=10,
                   command=self.edit_replace_rules).pack(side="left", padx=4)
        ttk.Button(r5c, text="删除冲突记忆", width=12,
                   command=self.clear_conflict_memory).pack(side="right")

        self.var_name_skip = tk.BooleanVar(value=True)
        ttk.Checkbutton(f5, text="身份/键相同但姓名对不上时不写入（防源表错行）",
                        variable=self.var_name_skip).pack(anchor="w")

        self.var_out = tk.StringVar()
        self.file_row(f5, "输出文件", self.var_out, self.pick_out_file)

        # --- 6 执行 ---
        f6 = ttk.Frame(right)
        f6.pack(fill="x", pady=(6, 0))
        ttk.Button(f6, text="① 预演（不写文件）", command=self.do_preview,
                   style="Go.TButton").pack(side="left")
        ttk.Button(f6, text="② 执行并保存", command=self.do_execute,
                   style="Go.TButton").pack(side="left", padx=6)
        ttk.Button(f6, text="保存规则", command=self.save_rule).pack(side="left")
        ttk.Button(f6, text="加载规则", command=self.load_rule).pack(side="left", padx=4)

        self.txt_fill = self.log_widget(root, height=8, use_grid=True)
        root.rowconfigure(1, weight=1)


    # ---------- 源表操作 ----------
    @staticmethod
    def _ask_header(path, sheet_name=None, extra_title=None):
        """弹窗让用户选工作表与表头行"""
        sheets = list_sheets(path)
        win = tk.Toplevel()
        win.title(extra_title or "选择工作表 / 表头")
        win.transient()
        win.grab_set()
        win.geometry("440x240")
        res = {}

        ttk.Label(win, text=extra_title or "选择工作表 / 表头",
                  style="Title.TLabel").pack(anchor="w", padx=12, pady=(10, 0))
        ttk.Label(win, text="文件：%s" % os.path.basename(path),
                  style="Hint.TLabel").pack(anchor="w", padx=12, pady=(0, 4))
        fr = ttk.Frame(win)
        fr.pack(fill="x", padx=12)
        ttk.Label(fr, text="工作表", width=8).pack(side="left")
        cb = ttk.Combobox(fr, values=sheets, state="readonly", width=28)
        cb.set(sheet_name if sheet_name in sheets else sheets[0])
        cb.pack(side="left")

        fr2 = ttk.Frame(win)
        fr2.pack(fill="x", padx=12, pady=6)
        ttk.Label(fr2, text="表头行", width=8).pack(side="left")
        v_row = tk.StringVar(value="自动")
        ttk.Spinbox(fr2, from_=1, to=50, width=6, textvariable=v_row).pack(side="left")
        ttk.Label(fr2, text="  表头行数").pack(side="left")
        v_span = tk.StringVar(value="自动")
        ttk.Spinbox(fr2, from_=1, to=3, width=4, textvariable=v_span).pack(side="left")

        def ok():
            res["sheet"] = cb.get()
            res["row"] = None if v_row.get().strip() in ("", "自动") else int(v_row.get())
            res["span"] = None if v_span.get().strip() in ("", "自动") else int(v_span.get())
            win.destroy()

        ttk.Button(win, text="确定", command=ok).pack(pady=8)
        win.wait_window()
        return res or None

    def add_source_file(self):
        paths = self.ask_open_multi("src", title="选择源表（可按住 Ctrl / Shift 多选）",
                                            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("全部", "*.*")])
        if not paths:
            return
        self.add_source_paths(list(paths))

    def add_source_paths(self, paths):
        """逐个文件询问工作表/表头后加入源表（支持一次选多个文件）"""
        n = len(paths)
        for i, p in enumerate(paths, start=1):
            info = self._ask_header(p, extra_title=("添加源表 %d/%d" % (i, n)) if n > 1 else None)
            if not info:
                self.log(self.txt_fill, "已跳过：%s" % os.path.basename(p))
                continue
            try:
                sd = open_table(p, info["sheet"], info["row"], info["span"],
                                visible_only=self.var_visible.get())
            except Exception as e:
                messagebox.showerror("读取失败", "%s：%s" % (os.path.basename(p), e))
                continue
            self.sources.append(sd)
            self.log(self.txt_fill, "已添加源表：%s（表头行 %d，数据 %d 行）"
                     % (sd.label, sd.header_row, len(sd.data_rows)))
        self.refresh_source_list()

    def add_source_all_sheets(self):
        paths = self.ask_open_multi("src",
            title="选择源文件（导入其全部工作表，可多选文件）",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("全部", "*.*")])
        if not paths:
            return
        self.add_source_paths_all_sheets(list(paths))

    def add_source_paths_all_sheets(self, paths):
        total = 0
        for p in paths:
            try:
                wb = openpyxl.load_workbook(p, data_only=True)
            except Exception as e:
                self.log(self.txt_fill, "跳过 %s：%s" % (os.path.basename(p), e))
                continue
            try:
                for name in wb.sheetnames:
                    try:
                        sd = SheetData(p, name, wb=wb, visible_only=self.var_visible.get())
                    except Exception as e:
                        self.log(self.txt_fill, "跳过 %s / %s：%s" % (os.path.basename(p), name, e))
                        continue
                    if len(sd.data_rows) == 0:
                        continue
                    self.sources.append(sd)
                    total += 1
            finally:
                wb.close()
            self.log(self.txt_fill, "已导入 %s（%d 个工作表有数据）" % (os.path.basename(p), 1))
        self.log(self.txt_fill, "批量导入完成：文件 %d 个，工作表 %d 个" % (len(paths), total))
        hidden = sum(s.hidden_in_data for s in self.sources)
        if hidden:
            self.log(self.txt_fill,
                     "⚠ 检测到 %d 个工作表带筛选/隐藏行，共隐藏 %d 行；当前按“%s”处理"
                     "（要只算可见行，勾选上方“只处理可见行”）"
                     % (sum(1 for s in self.sources if s.hidden_in_data), hidden,
                        "只处理可见行" if self.var_visible.get() else "读取全部行（含隐藏行）"))
        self.refresh_source_list()
        self.check_structure()

    def refresh_source_list(self):
        self.lb_src.delete(0, "end")
        for sd in self.sources:
            self.lb_src.insert("end", "%s  ｜表头行%d ｜%d行" % (sd.label, sd.header_row, len(sd.data_rows)))
        if self.sources:
            vals = self.sources[0].col_labels()
        else:
            vals = []
        self.cb_map_src["values"] = vals
        if vals and not self.cb_map_src.get():
            self.cb_map_src.current(0)
        self.cb_keysrc["values"] = vals
        if vals and not self.cb_keysrc.get():
            self.cb_keysrc.set(self._guess_id_col(self.sources[0].headers) or vals[0])
        # 源表变化 → 重算工作表对应关系
        try:
            if self.tv_pairs.winfo_exists():
                self.rebuild_pairs()
                self.try_auto_keys()
        except Exception:
            pass

    @staticmethod
    @staticmethod
    def _name_col(table):
        for h in table.headers:
            if alias_group(h) == "name":
                return h
        return None

    @staticmethod
    def _guess_id_col(headers):
        for h in headers:
            if alias_group(h) == "id":
                return h
        for h in headers:
            if "身份" in h or "证件" in h:
                return h
        return None

    def reload_sources(self):
        """切换“只处理可见行”后，按新口径重新读取所有源表"""
        if not self.sources:
            return
        old = list(self.sources)
        self.sources = []
        hidden_total = 0
        for sd in old:
            try:
                nd = open_table(sd.path, sd.sheet, sd.header_row, sd.header_rows,
                                visible_only=self.var_visible.get())
            except Exception as e:
                self.log(self.txt_fill, "重新读取失败 %s：%s" % (sd.label, e))
                continue
            self.sources.append(nd)
            hidden_total += nd.hidden_in_data
        self.refresh_source_list()
        self.log(self.txt_fill, "已按“%s”重新读取 %d 个源表%s"
                 % ("只处理可见行" if self.var_visible.get() else "读取全部行",
                    len(self.sources),
                    "（含筛选隐藏行 %d 行）" % hidden_total if hidden_total else ""))

    def check_structure(self):
        """同一工作表的多份来源，列数/列名不一致时给出提示（避免漏写、错位）"""
        groups = {}
        for s in self.sources:
            groups.setdefault(norm_colname(s.sheet), []).append(s)
        msgs = []
        for key, items in groups.items():
            if len(items) < 2:
                continue
            cols = {}
            for it in items:
                cols.setdefault(it.n_cols, []).append(it)
            if len(cols) > 1:
                main = max(cols, key=lambda k: len(cols[k]))
                odd = [it for k, v in cols.items() if k != main for it in v]
                msgs.append("工作表「%s」列数不一致：多数 %d 列，以下 %d 个不同 → %s"
                            % (items[0].sheet, main, len(odd),
                               "、".join(os.path.basename(o.path) for o in odd[:5])))
            # 列名差异
            base = items[0]
            bset = set(norm_colname(h) for h in base.headers)
            for it in items[1:]:
                sset = set(norm_colname(h) for h in it.headers)
                if bset != sset:
                    extra = sset - bset
                    miss = bset - sset
                    if extra or miss:
                        msgs.append("工作表「%s」列名有差异（%s）：多出 %s ／ 缺少 %s"
                                    % (it.sheet, os.path.basename(it.path),
                                       "、".join(sorted(extra)) or "无",
                                       "、".join(sorted(miss)) or "无"))
                        break
        for m in msgs[:12]:
            self.log(self.txt_fill, "⚠ " + m)
        if not msgs:
            self.log(self.txt_fill, "结构检查：各源表结构一致，无异常")

    def remove_source(self):
        sel = list(self.lb_src.curselection())
        for i in reversed(sel):
            self.sources.pop(i)
        self.refresh_source_list()

    def clear_sources(self):
        self.sources = []
        self.refresh_source_list()

    def edit_source_header(self):
        sel = list(self.lb_src.curselection())
        if not sel:
            messagebox.showinfo("提示", "请先在源表列表中选中一行（可按住 Ctrl/Shift 多选）")
            return
        sd = self.sources[sel[0]]
        title = "设置表头（将应用到选中的 %d 个源表）" % len(sel) if len(sel) > 1 else "设置表头"
        info = self._ask_header(sd.path, sd.sheet, extra_title=title)
        if not info:
            return
        for i in sel:
            old = self.sources[i]
            try:
                self.sources[i] = open_table(old.path, info["sheet"], info["row"], info["span"],
                                             visible_only=self.var_visible.get())
            except Exception as e:
                messagebox.showerror("读取失败", "%s：%s" % (old.sheet, e))
                return
            self.log(self.txt_fill, "已更新源表：%s（表头行 %d，数据 %d 行）"
                     % (self.sources[i].label, self.sources[i].header_row,
                        len(self.sources[i].data_rows)))
        self.refresh_source_list()

    # ---------- 目标表（支持多工作表） ----------
    def pick_dst_file(self):
        p = self.ask_open("dst", title="选择目标表",
                                       filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("全部", "*.*")])
        if not p:
            return
        self.var_dst_path.set(p)
        if not self.settings.get("dir_src"):
            self.save_dir("src", p)
        self.load_dst_workbook(p)

    def load_dst_workbook(self, path=None):
        """列出目标工作簿全部工作表，默认全选，并自动与源表对应"""
        path = path or self.var_dst_path.get().strip()
        if not path or not os.path.exists(path):
            return
        try:
            names = list_sheets(path)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        self.dst_sheets = names
        self.dst_cache = {}
        self.dst = None
        self.lb_dst.delete(0, "end")
        for n in names:
            self.lb_dst.insert("end", n)
        self.lb_dst.selection_set(0, "end")      # 默认全选
        self.pairs = []
        self.rebuild_pairs()
        self.try_auto_keys()
        self.log(self.txt_fill, "已加载目标文件：%s（%d 个工作表，已全选）"
                 % (os.path.basename(path), len(names)))

    def dst_selected_sheets(self):
        return [self.lb_dst.get(i) for i in self.lb_dst.curselection()]

    def dst_select_all(self):
        self.lb_dst.selection_set(0, "end")
        self.on_dst_select()

    def dst_invert(self):
        cur = set(self.lb_dst.curselection())
        for i in range(self.lb_dst.size()):
            if i in cur:
                self.lb_dst.selection_clear(i)
            else:
                self.lb_dst.selection_set(i)
        self.on_dst_select()

    def dst_clear_sel(self):
        self.lb_dst.selection_clear(0, "end")
        self.on_dst_select()

    def on_dst_select(self):
        self.rebuild_pairs()

    @staticmethod
    def _sheet_eq(a, b):
        return norm_colname(a) == norm_colname(b)

    def _match_target(self, sheet_name, targets, used):
        """同名 → 包含关系，依次尝试"""
        for t in targets:
            if t not in used and self._sheet_eq(t, sheet_name):
                return t
        for t in targets:
            if t in used:
                continue
            a, b = norm_colname(sheet_name), norm_colname(t)
            if a and b and (a in b or b in a):
                return t
        return None

    def rebuild_pairs(self):
        """源表工作表 → 目标表工作表 的对应关系（保留用户手改的结果）"""
        try:
            targets = self.dst_selected_sheets()
        except Exception:
            return
        manual = {p["src"]: p["target"] for p in self.pairs}
        pairs, used = [], set()
        for i, s in enumerate(self.sources):
            t = manual.get(i)
            if t not in targets:
                t = None
            elif t in used:
                t = None
            if t is None:
                t = self._match_target(s.sheet, targets, used)
            if t is None:
                # 允许多个源表对应同一个目标工作表（此时会在“冲突”里让你选值）
                t = self._match_target(s.sheet, targets, set())
            if t is None and len(targets) >= len(self.sources) and i < len(targets):
                t = targets[i]
            if t:
                used.add(t)
            pairs.append({"src": i, "target": t})
        self.pairs = pairs
        self.refresh_pairs_tree()
        self.reload_dst()

    def refresh_pairs_tree(self):
        self.tv_pairs.delete(*self.tv_pairs.get_children())
        for p in self.pairs:
            sname = self.sources[p["src"]].sheet if p["src"] < len(self.sources) else "?"
            self.tv_pairs.insert("", "end", values=(sname, p["target"] or "（未指定，双击选择）"))

    def auto_correspond(self):
        if not self.sources:
            messagebox.showinfo("提示", "请先添加源表")
            return
        if not self.dst_selected_sheets():
            messagebox.showinfo("提示", "请先选择目标工作表")
            return
        self.pairs = []
        self.rebuild_pairs()
        n = sum(1 for p in self.pairs if p["target"])
        self.log(self.txt_fill, "自动对应完成：%d/%d 条已匹配（可双击行手动改）" % (n, len(self.pairs)))

    def change_pair_target(self):
        sel = self.tv_pairs.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在对应关系表里选中一行")
            return
        idx = self.tv_pairs.index(sel[0])
        win = tk.Toplevel(self)
        win.title("选择该源表对应的目标工作表")
        win.transient()
        win.grab_set()
        win.geometry("360x340")
        ttk.Label(win, text="源表：%s" % self.sources[self.pairs[idx]["src"]].label,
                  style="Hint.TLabel").pack(anchor="w", padx=12, pady=(10, 4))
        lb = tk.Listbox(win, height=10)
        lb.insert("end", "（跳过这个源表）")
        for n in self.dst_selected_sheets():
            lb.insert("end", n)
        cur = self.pairs[idx]["target"]
        if cur:
            for i, n in enumerate(self.dst_selected_sheets(), start=1):
                if n == cur:
                    lb.selection_set(i)
                    break
        lb.pack(fill="both", expand=True, padx=12)

        def ok():
            s = lb.curselection()
            if s:
                v = lb.get(s[0])
                self.pairs[idx]["target"] = None if v.startswith("（跳过") else v
                self.refresh_pairs_tree()
                self.reload_dst()
            win.destroy()

        ttk.Button(win, text="确定", command=ok).pack(pady=8)
        win.wait_window()

    def get_dst(self, sheet):
        """取目标工作表（带缓存），表头行/行数为“自动”时逐表识别"""
        hr = self.var_hdr_row.get().strip()
        hs = self.var_hdr_span.get().strip()
        key = (sheet, hr, hs)
        if key not in self.dst_cache:
            self.dst_cache[key] = open_table(self.var_dst_path.get().strip(), sheet,
                                             None if hr in ("", "自动") else int(hr),
                                             None if hs in ("", "自动") else int(hs))
        return self.dst_cache[key]

    def reload_dst(self):
        """加载基准目标表：优先用对应关系里第一个已指定的工作表"""
        p = self.var_dst_path.get().strip()
        if not p or not os.path.exists(p):
            return
        name = next((x["target"] for x in self.pairs if x["target"]), None)
        if not name:
            sel = self.dst_selected_sheets()
            name = sel[0] if sel else None
        if not name:
            return
        try:
            self.dst = self.get_dst(name)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        vals = self.dst.col_labels()
        self.cb_map_dst["values"] = vals
        self.cb_keydst["values"] = vals
        g = self._guess_id_col(self.dst.headers)
        if g:
            for lab in vals:
                if lab.endswith("｜" + g):
                    self.cb_keydst.set(lab)
                    break
        n_pair = len(self.pairs)
        n_ok = sum(1 for x in self.pairs if x["target"])
        self.lb_dst_info.configure(
            text="基准表「%s」：表头行 %d（表头 %d 行），%d 行 × %d 列 ｜ 已选工作表 %d 个 ｜ 对应 %d/%d 条"
                 % (self.dst.sheet, self.dst.header_row, self.dst.header_rows,
                    len(self.dst.data_rows), self.dst.n_cols,
                    len(self.dst_selected_sheets()), n_ok, n_pair))

    # ---------- 匹配键 ----------
    @staticmethod
    def _col_by_group(table, group):
        for h in table.headers:
            if alias_group(h) == group:
                return h
        return None

    @staticmethod
    def _unique_ratio(table, col_idx):
        vals = [cell_text(r[col_idx]) for r in table.data_rows
                if col_idx < len(r) and cell_text(r[col_idx]) != ""]
        if not vals:
            return 0.0
        return len(set(vals)) / float(len(vals))

    def auto_key_pairs(self, silent=False):
        """自动识别匹配键：身份证 → 姓名+手机 → 手机 → 姓名 → 村+姓名 → 同名唯一列"""
        if not self.sources or not self.dst:
            if not silent:
                messagebox.showinfo("提示", "请先添加源表和目标表，再自动匹配键")
            return
        src, dst = self.sources[0], self.dst
        combos = [("id",), ("name", "phone"), ("phone",), ("name",), ("village", "name")]
        picked = None
        for combo in combos:
            pairs, okk = [], True
            for g in combo:
                a = self._col_by_group(src, g)
                b = self._col_by_group(dst, g)
                if not a or not b:
                    okk = False
                    break
                pairs.append((a, b))
            if okk:
                picked = pairs
                break
        if not picked:
            # 兜底：同名且取值基本唯一的列
            dnorm = {norm_colname(h): i for i, h in enumerate(dst.headers)}
            best = None
            for h in src.headers:
                di = dnorm.get(norm_colname(h))
                if di is None:
                    continue
                ratio = self._unique_ratio(dst, di)
                if ratio >= 0.9 and (best is None or ratio > best[2]):
                    best = (h, dst.headers[di], ratio)
            if best:
                picked = [(best[0], best[1])]
        if not picked:
            if not silent:
                messagebox.showinfo("提示", "没能自动识别匹配键，请用左右下拉框手动添加")
            return
        self.key_pairs = picked
        self.tv_keys.delete(*self.tv_keys.get_children())
        for a, b in picked:
            self.tv_keys.insert("", "end", values=(a, b))
        di = self._resolve(dst, picked[0][1])
        ratio = self._unique_ratio(dst, di) if di is not None else 0
        self.log(self.txt_fill, "自动匹配键：%s（目标列唯一率 %s）"
                 % (" ↔ ".join(picked[0]), ("%.0f%%" % (ratio * 100)) if ratio else "—"))
        if ratio and ratio < 0.95:
            self.log(self.txt_fill, "  注意：该键在目标表中不唯一，只有第一条会被匹配（重复键在明细清单可见）")

    def try_auto_keys(self):
        """源表/目标表就绪且还没设键时，自动试一次"""
        if self.sources and self.dst and not self.key_pairs:
            self.auto_key_pairs(silent=True)

    def add_key_pair(self):
        if "｜" not in self.cb_keysrc.get() or "｜" not in self.cb_keydst.get():
            messagebox.showinfo("提示", "请先在左右两个下拉框选择列")
            return
        s_lab, d_lab = self.cb_keysrc.get(), self.cb_keydst.get()
        sname = s_lab.split("｜", 1)[1]
        dname = d_lab.split("｜", 1)[1]
        if (sname, dname) in self.key_pairs:
            messagebox.showinfo("提示", "这组匹配键已经添加过了")
            return
        self.key_pairs.append((sname, dname))   # 按“列名”保存，跨工作表自动适配
        self.tv_keys.insert("", "end", values=(sname, dname))
        self.log(self.txt_fill, "匹配键：%s ↔ %s" % (sname, dname))

    def del_key_pair(self):
        for iid in self.tv_keys.selection():
            idx = self.tv_keys.index(iid)
            self.tv_keys.delete(iid)
            self.key_pairs.pop(idx)

    # ---------- 字段映射 ----------
    def _resolve(self, table, name, idx=None):
        """按列名在本表中定位列（同名 → 规范化同名 → 同义词）

        注意：给了列名就严格按名字找，找不到返回 None —— 绝不按下标硬套，
        否则批量处理不同结构的工作表会写错列。
        """
        if name is None or name == "":
            return idx if (isinstance(idx, int) and 0 <= idx < len(table.headers)) else None
        if isinstance(name, int):        # 兼容旧规则里的下标
            return name if 0 <= name < len(table.headers) else None
        if name in table.headers:
            return table.headers.index(name)
        nn = norm_colname(name)
        for i, h in enumerate(table.headers):
            if norm_colname(h) == nn:
                return i
        g = alias_group(name)
        if g:
            for i, h in enumerate(table.headers):
                if alias_group(h) == g:
                    return i
        return None

    def add_mapping(self):
        if not self.sources:
            messagebox.showinfo("提示", "请先添加源表")
            return
        if not self.dst:
            messagebox.showinfo("提示", "请先加载目标表")
            return
        s_lab, d_lab = self.cb_map_src.get(), self.cb_map_dst.get()
        if "｜" not in s_lab or "｜" not in d_lab:
            messagebox.showinfo("提示", "请先在左右两个下拉框选择列")
            return
        sname = s_lab.split("｜", 1)[1]
        dname = d_lab.split("｜", 1)[1]
        si = self.sources[0].col_labels().index(s_lab)
        di = self.dst.col_labels().index(d_lab)
        self.mappings.append({"src": si, "src_name": sname,
                              "dst": di, "dst_name": dname, "mode": "默认"})
        self.refresh_map_tree()

    def refresh_map_tree(self):
        self.tv_map.delete(*self.tv_map.get_children())
        for m in self.mappings:
            s = m.get("src_name") or (self.sources[0].headers[m["src"]] if self.sources else "?")
            d = m.get("dst_name") or (self.dst.headers[m["dst"]] if self.dst else "?")
            self.tv_map.insert("", "end", values=(s, d, m.get("mode", "默认")))

    def toggle_mapping_mode(self):
        sel = self.tv_map.selection()
        if not sel:
            return
        i = self.tv_map.index(sel[0])
        order = ["默认", "仅填空", "覆盖", "跳过"]
        m = self.mappings[i]
        m["mode"] = order[(order.index(m["mode"]) + 1) % len(order)] if m["mode"] in order else "默认"
        self.refresh_map_tree()

    def del_mapping(self):
        for iid in self.tv_map.selection():
            idx = self.tv_map.index(iid)
            self.tv_map.delete(iid)
            self.mappings.pop(idx)

    def clear_mapping(self):
        self.mappings = []
        self.refresh_map_tree()

    def auto_map(self):
        if not self.sources or not self.dst:
            messagebox.showinfo("提示", "请先添加源表和目标表")
            return
        dst = self.dst
        dnorm = {norm_colname(h): i for i, h in enumerate(dst.headers)}
        added = 0
        done_src = set()
        for src in self.sources:
            for h in src.headers:
                if h in done_src:
                    continue
                done_src.add(h)
                di = dnorm.get(norm_colname(h))
                if di is None:
                    g = alias_group(h)
                    if g:
                        for dh, dj in dnorm.items():
                            if alias_group(dst.headers[dj]) == g:
                                di = dj
                                break
                if di is not None and not any(m.get("src_name") == h for m in self.mappings):
                    self.mappings.append({"src": src.headers.index(h), "src_name": h,
                                          "dst": di, "dst_name": dst.headers[di], "mode": "默认"})
                    added += 1
        self.refresh_map_tree()
        self.log(self.txt_fill, "自动匹配完成，新增 %d 条映射（按全部源表的列名合并，请人工核对一遍）" % added)

    # ---------- 执行 ----------
    def _kopt(self):
        return {"nospace": self.var_k_nospace.get(), "upper": self.var_k_upper.get(),
                "symbol": self.var_k_symbol.get(), "digits": self.var_k_digits.get()}

    def _check(self):
        if not self.var_dst_path.get().strip():
            messagebox.showwarning("提示", "请先选择目标表")
            return False
        if not self.dst:
            messagebox.showwarning("提示", "请先选择目标工作表")
            return False
        if not self.sources:
            messagebox.showwarning("提示", "请先添加源表")
            return False
        if not any(p["target"] for p in self.pairs):
            messagebox.showwarning("提示", "请先设置“源表工作表 ↔ 目标表工作表”的对应关系")
            return False
        if not self.key_pairs:
            messagebox.showwarning("提示", "请至少设置一组匹配键")
            return False
        if not self.mappings:
            messagebox.showwarning("提示", "请至少设置一条字段映射")
            return False
        return True

    def build_jobs(self):
        """按目标工作表分组：{目标工作表: [源表, ...]}"""
        jobs, skipped = {}, []
        for p in self.pairs:
            t = p["target"]
            if not t:
                skipped.append(self.sources[p["src"]].sheet)
                continue
            jobs.setdefault(t, []).append(self.sources[p["src"]])
        return jobs, skipped

    def make_plans(self):
        """按目标工作表生成补全计划：匹配键/映射都按“列名”在各表中重新定位"""
        jobs, skipped = self.build_jobs()
        plans, warns = [], list(skipped)
        for sheet, srcs in jobs.items():
            dst = self.get_dst(sheet)
            # 数据预处理：空白继承 / 值标准化
            pre = []
            for s in srcs:
                t = s
                if self.var_inherit.get():
                    t = inherit_blank_rows(t, self.inherit_cols)
                if self.var_std.get() and self.replace_rules:
                    t = standardize_table(t, self.replace_rules)
                pre.append(t)
            srcs = pre
            base = srcs[0]
            kp, ok = [], True
            for sname, dname in self.key_pairs:
                si = self._resolve(base, sname)
                di = self._resolve(dst, dname)
                if si is None or di is None:
                    warns.append("%s：匹配键「%s / %s」在该表找不到对应列，已跳过"
                                 % (sheet, sname, dname))
                    ok = False
                    break
                kp.append((si, di))
            if not ok:
                continue
            mp = []
            for m in self.mappings:
                si = self._resolve(base, m.get("src_name"), m.get("src"))
                di = self._resolve(dst, m.get("dst_name"), m.get("dst"))
                if si is None and di is None:
                    continue          # 该字段与这个工作表无关，静默跳过
                if si is None or di is None:
                    warns.append("%s：字段「%s → %s」在该表找不到，已跳过"
                                 % (sheet, m.get("src_name"), m.get("dst_name")))
                    continue
                mp.append({"src": si, "dst": di, "mode": m.get("mode", "默认")})
            if not mp:
                warns.append("%s：没有可用的字段映射，已跳过" % sheet)
                continue
            # “搬运整行”的模式下，各工作表结构可能不同：
            # 为该表自动补上“同名/同义”的列，避免漏写（补信息模式不动，尊重用户映射）
            if self.var_mode.get() in ("add_new", "fill_add", "overwrite_add"):
                covered = {m["dst"] for m in mp}
                extra = 0
                for di, h in enumerate(dst.headers):
                    if di in covered or norm_colname(h) in ("序号", "编号", "序"):
                        continue
                    si = self._resolve(base, h)
                    if si is None:
                        continue
                    mp.append({"src": si, "dst": di, "mode": "默认"})
                    extra += 1
                if extra:
                    warns.append("%s：按列名自动补充了 %d 个字段映射" % (sheet, extra))
            s_nm = self._name_col(base)
            d_nm = self._name_col(dst)
            npair = (base.headers.index(s_nm), dst.headers.index(d_nm)) \
                if (s_nm and d_nm) else None
            plan = plan_fill(srcs, dst, kp, mp, self.var_mode.get(), self._kopt(),
                             name_pair=npair, name_skip=self.var_name_skip.get())
            plans.append((dst, plan))
        return plans, warns

    # ---------- 数据预处理：空白继承 / 值标准化 ----------
    def set_inherit_cols(self):
        if not self.sources:
            messagebox.showinfo("提示", "请先添加源表")
            return
        src = self.sources[0]
        win = tk.Toplevel(self)
        win.title("选择要“空白继承上一行”的列")
        win.transient(self)
        win.grab_set()
        win.geometry("480x460")
        ttk.Label(win, text="勾选的列：空白时自动取上一行的值（用于“同户多人只填一次户主信息”的台账）",
                  wraplength=440, justify="left").pack(anchor="w", padx=12, pady=(10, 2))
        ttk.Label(win, text="源表：%s" % src.label, style="Hint.TLabel").pack(anchor="w", padx=12)
        lb = tk.Listbox(win, selectmode="extended", height=12)
        lb.pack(fill="both", expand=True, padx=12, pady=6)
        skip = ("序号", "编号", "序")
        guess = set(guess_inherit_cols(src))
        for i, h in enumerate(src.headers):
            ex = ""
            for r in src.data_rows:
                if i < len(r) and cell_text(r[i]):
                    ex = cell_text(r[i])[:12]
                    break
            tag = "★" if h in guess else "  "
            lb.insert("end", "%s%s    （例：%s）" % (tag, h.replace("\n", " "), ex))
            if h in guess:
                lb.selection_set(i)
        for i, h in enumerate(src.headers):
            if norm_colname(h) in skip:
                lb.selection_clear(i)

        def all_on():
            lb.selection_set(0, "end")
        def all_off():
            lb.selection_clear(0, "end")

        def ok():
            chosen = [src.headers[i] for i in lb.curselection()]
            allcols = [h for h in src.headers if norm_colname(h) not in skip]
            self.inherit_cols = [] if len(chosen) >= len(allcols) else chosen
            self.var_inherit.set(True)
            win.destroy()
            self.log(self.txt_fill, "空白继承列：%s"
                     % ("智能识别（同户共享列）" if not self.inherit_cols
                        else "、".join(self.inherit_cols)))

        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=12, pady=8)
        ttk.Button(bar, text="全选", command=all_on).pack(side="left")
        ttk.Button(bar, text="全不选", command=all_off).pack(side="left", padx=4)
        ttk.Button(bar, text="确定", command=ok, style="Go.TButton").pack(side="right")
        ttk.Button(bar, text="取消", command=win.destroy).pack(side="right", padx=6)
        win.wait_window()

    def edit_replace_rules(self):
        win = tk.Toplevel(self)
        win.title("字段值标准化规则")
        win.transient(self)
        win.grab_set()
        win.geometry("700x520")
        ttk.Label(win, text="把源表里的旧值替换成新值；“列”留空表示对所有列生效。"
                           "例：万和 → 万和镇、曹门 → 曹门村",
                  style="Hint.TLabel").pack(anchor="w", padx=12, pady=(10, 2))

        tv = ttk.Treeview(win, columns=("c", "o", "n"), show="headings", height=10)
        for c, t, w in (("c", "列（可留空）", 170), ("o", "原值", 200), ("n", "新值", 200)):
            tv.heading(c, text=t)
            tv.column(c, width=w)
        tv.pack(fill="both", expand=True, padx=12)

        def refresh():
            tv.delete(*tv.get_children())
            for r in self.replace_rules:
                tv.insert("", "end", values=(r.get("col", "") or "（全部列）",
                                             r.get("old", ""), r.get("new", "")))

        fr = ttk.Frame(win)
        fr.pack(fill="x", padx=12, pady=6)
        cols = ["（全部列）"] + (self.sources[0].headers if self.sources else [])
        v_col = tk.StringVar(value="（全部列）")
        cb = ttk.Combobox(fr, textvariable=v_col, values=cols, state="readonly", width=24)
        cb.pack(side="left")
        v_old = tk.StringVar()
        v_new = tk.StringVar()
        ttk.Label(fr, text=" 原值").pack(side="left")
        ttk.Entry(fr, textvariable=v_old, width=14).pack(side="left")
        ttk.Label(fr, text=" 新值").pack(side="left")
        ttk.Entry(fr, textvariable=v_new, width=14).pack(side="left")

        def add():
            o, n = v_old.get().strip(), v_new.get().strip()
            if not o:
                messagebox.showinfo("提示", "请填“原值”", parent=win)
                return
            col = v_col.get()
            self.replace_rules.append({"col": "" if col.startswith("（") else col, "old": o, "new": n})
            v_old.set("")
            v_new.set("")
            refresh()

        ttk.Button(fr, text="添加", command=add).pack(side="left", padx=6)

        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=12, pady=8)

        def del_sel():
            for iid in tv.selection():
                self.replace_rules.pop(tv.index(iid))
            refresh()

        def auto_rules():
            if not self.sources:
                messagebox.showinfo("提示", "请先添加源表", parent=win)
                return
            rules, have = [], {(r["col"], r["old"], r["new"]) for r in self.replace_rules}
            added = 0
            for r in auto_suffix_rules_multi(self.sources):
                if (r["col"], r["old"], r["new"]) not in have:
                    have.add((r["col"], r["old"], r["new"]))
                    self.replace_rules.append(r)
                    added += 1
            refresh()
            messagebox.showinfo("自动规则",
                                "已扫描 %d 个源表，新增 %d 条“后缀统一”规则（如 万和→万和镇），"
                                "可继续手动增改" % (len(self.sources), added), parent=win)

        def ok():
            self.var_std.set(bool(self.replace_rules))
            win.destroy()
            self.log(self.txt_fill, "值标准化规则：%d 条%s"
                     % (len(self.replace_rules), "" if self.replace_rules else "（未设置，已关闭该功能）"))

        def auto_fill_suffix():
            if not self.sources:
                messagebox.showinfo("提示", "请先添加源表", parent=win)
                return
            have = {(r["col"], r["old"], r["new"]) for r in self.replace_rules}
            added = 0
            for r in auto_suffix_fill_rules(self.sources):
                if (r["col"], r["old"], r["new"]) not in have:
                    have.add((r["col"], r["old"], r["new"]))
                    self.replace_rules.append(r)
                    added += 1
            refresh()
            messagebox.showinfo("补齐后缀", "已新增 %d 条规则（如 曹门→曹门村、万和→万和镇）" % added,
                                parent=win)

        ttk.Button(bar, text="删除选中", command=del_sel).pack(side="left")
        ttk.Button(bar, text="自动统一后缀", command=auto_rules).pack(side="left", padx=6)
        ttk.Button(bar, text="补齐缺失后缀", command=auto_fill_suffix).pack(side="left")
        ttk.Button(bar, text="确定", command=ok, style="Go.TButton").pack(side="right")
        ttk.Button(bar, text="取消", command=win.destroy).pack(side="right", padx=6)
        refresh()
        win.wait_window()

    def ask_output_path(self, suggested, src_path, key="out"):
        """执行时让用户选保存位置；位置不可写/被占用就提示并重选，直到有效或取消"""
        cur = suggested
        while True:
            d0 = os.path.dirname(os.path.abspath(cur)) or "."
            p = self.ask_save(key, title="选择结果保存位置",
                              initialfile=os.path.basename(cur),
                              initialdir=d0 if os.path.isdir(d0) else None,
                              defaultextension=".xlsx",
                              filetypes=[("Excel 工作簿", "*.xlsx")])
            if not p:
                return None
            p = os.path.abspath(p)
            if os.path.abspath(src_path) == p:
                if not messagebox.askokcancel(
                        "将覆盖目标文件",
                        "你选择的位置就是目标文件本身：\n%s\n\n"
                        "继续会直接改写它（程序会自动生成 .bak 备份）。确定吗？" % p):
                    cur = p
                    continue
                return p
            d = os.path.dirname(p)
            if not os.path.isdir(d):
                try:
                    os.makedirs(d, exist_ok=True)
                except Exception as e:
                    messagebox.showerror("该位置不可用", "无法创建文件夹：\n%s\n%s" % (d, e))
                    cur = p
                    continue
            if not dir_writable(d):
                messagebox.showerror(
                    "该位置不可用",
                    "这个文件夹写不进去：\n%s\n\n"
                    "常见原因：\n"
                    "· 被安全软件（火绒/电脑管家等）的“文档保护/勒索防护”拦截；\n"
                    "· 文件夹权限受限、或位于只读盘/受保护目录。\n\n"
                    "请换一个位置（例如桌面、D 盘根目录下的新建文件夹）后再试。" % d)
                cur = p
                continue
            if os.path.exists(p):
                try:                      # 已存在：确认能覆盖（未被 Excel 占用）
                    with open(p, "a+b"):
                        pass
                except Exception as e:
                    messagebox.showerror(
                        "该文件正被占用",
                        "无法写入这个文件：\n%s\n\n"
                        "它可能正在 Excel 中打开，或被设为只读。\n"
                        "请先关闭它、或换个文件名/位置。\n\n（%s）" % (p, e))
                    cur = p
                    continue
            return p

    def ensure_output_path(self, out):
        """写入前先实测输出目录是否可写；不可写就让用户换地方"""
        d = os.path.dirname(os.path.abspath(out)) or "."
        if dir_writable(d):
            return out
        cands = [x for x in fallback_dirs() if dir_writable(x)]
        tip = ("输出目录不可写：\n%s\n\n"
               "常见原因：\n"
               "· 该文件夹被安全软件（火绒/电脑管家等）的“文档保护/勒索防护”拦截；\n"
               "· 文件夹权限受限。\n\n"
               % d)
        if cands:
            tip += "是否改为保存到：\n%s\n（文件名不变）" % os.path.join(cands[0], os.path.basename(out))
            if messagebox.askokcancel("该文件夹写不进去", tip):
                newp = os.path.join(cands[0], os.path.basename(out))
                self.log(self.txt_fill, "⚠ 输出目录不可写，已改到：%s" % cands[0])
                return newp
        p = self.ask_save("out", title="请选择可写入的保存位置",
                          initialfile=os.path.basename(out), defaultextension=".xlsx")
        if p:
            return p
        messagebox.showinfo("已取消", "请把输出位置改到其他文件夹后重试，"
                                     "或把本程序加入安全软件的信任区。")
        return None

    def clear_conflict_memory(self):
        """清空“冲突选择记忆”，之后同样的冲突会重新弹窗询问"""
        n = len(self.conflict_cache)
        if n == 0:
            messagebox.showinfo("提示", "当前没有冲突记忆（未做过冲突选择，或已清空）")
            return
        detail = "\n".join("· %s / %s → 采用「%s」"
                           % (k[0], k[3] or k[2], v[0]) for k, v in list(self.conflict_cache.items())[:10])
        if not messagebox.askokcancel("清除冲突记忆",
                                      "当前共记住 %d 条冲突选择：\n%s\n\n"
                                      "清除后，遇到同样的冲突会重新弹窗让你选择。是否清除？"
                                      % (n, detail)):
            return
        self.conflict_cache.clear()
        self.log(self.txt_fill, "✔ 已清除 %d 条冲突记忆，下次遇到冲突将重新弹窗" % n)
        messagebox.showinfo("完成", "已清除 %d 条冲突记忆" % n)

    # ---------- 源表内容冲突 ----------
    @staticmethod
    def _opt_sig(o):
        if "value" in o:
            return cell_text(o["value"])
        return "|".join("%s=%s" % (k, cell_text(v)) for k, v in sorted(o["values"].items()))

    def _conflict_key(self, sheet, c):
        return (sheet, c["type"], c["key"], c.get("colname") or "")

    def resolve_all_conflicts(self, plans):
        """源表之间有冲突时弹窗让用户选择采用哪个源表的值

        返回 True=已全部确定；False=用户取消
        """
        items = []
        for dst, plan in plans:
            for c in plan["conflicts"]:
                items.append((dst.sheet, plan, c))
        if not items:
            return True
        resolved, pending = {}, []
        for sheet, plan, c in items:
            ck = ("cell", c["ri"], c["col"]) if c["type"] == "cell" else ("newrow", c["key"])
            hit = self.conflict_cache.get(self._conflict_key(sheet, c))
            idx = None
            if hit is not None:
                for i, o in enumerate(c["options"]):
                    if (o["src"], self._opt_sig(o)) == hit:
                        idx = i
                        break
            if idx is None:
                pending.append((sheet, plan, c, ck))
            else:
                resolved.setdefault(id(plan), {})[ck] = idx
        if pending:
            if not self.conflict_dialog(pending, resolved):
                return False
        for dst, plan in plans:
            if plan["conflicts"]:
                resolve_conflicts(plan, resolved.get(id(plan), {}))
        n = sum(len(p.get("resolved", [])) for _, p in plans)
        if n:
            self.log(self.txt_fill, "已确认 %d 处源表内容冲突（采用明细见“源表冲突（已选）”页）" % n)
        return True

    def conflict_dialog(self, pending, resolved):
        """冲突选择窗：每条冲突列出各源表的候选值，逐条选择"""
        win = tk.Toplevel(self)
        win.title("源表内容冲突 —— 请选择采用哪个源表的数据")
        win.transient(self)
        win.grab_set()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        win.geometry("%dx%d" % (int(sw * 0.62), int(sh * 0.62)))
        ttk.Label(win, text="源表之间对同一数据给出了不同的值，请选择要导入目标表的那个（可多选行批量设置）",
                  style="Hint.TLabel").pack(anchor="w", padx=12, pady=(10, 2))
        self._dlg_count = ttk.Label(win, text="", style="Title.TLabel")
        self._dlg_count.pack(anchor="w", padx=12)

        body = ttk.Frame(win)
        body.pack(fill="both", expand=True, padx=10, pady=6)

        cols = ("sheet", "key", "where", "cand", "chosen")
        tv = ttk.Treeview(body, columns=cols, show="headings", height=9, selectmode="extended")
        for c, t, w in (("sheet", "工作表", 110), ("key", "匹配键", 160),
                        ("where", "位置", 110), ("cand", "候选数", 60),
                        ("chosen", "已选来源", 200)):
            tv.heading(c, text=t)
            tv.column(c, width=w)
        tv.pack(fill="both", expand=True)

        lb_box = ttk.LabelFrame(win, text=" 候选值（双击即采用该值） ", padding=4)
        lb_box.pack(fill="both", expand=True, padx=10)
        lb = tk.Listbox(lb_box, height=5)
        lb.pack(fill="both", expand=True)

        idx_of = {}     # tree iid -> pending 下标
        chosen = {}     # pending 下标 -> 候选下标

        def refresh_row(order):
            i = order
            sheet, plan, c, ck = pending[i]
            tv.item(str(i), values=(sheet, c["key"],
                                    ("第%s行" % c.get("excel_row")) if c["type"] == "cell" else "新增行",
                                    len(c["options"]),
                                    c["options"][chosen.get(i, 0)]["src"]))

        for i, (sheet, plan, c, ck) in enumerate(pending):
            tv.insert("", "end", iid=str(i), values=("", "", "", "", ""))
            idx_of[str(i)] = i
            chosen[i] = 0
            refresh_row(i)

        def show_cands(_=None):
            lb.delete(0, "end")
            sel = tv.selection()
            if not sel:
                return
            i = idx_of[sel[0]]
            sheet, plan, c, ck = pending[i]
            for j, o in enumerate(c["options"]):
                prefix = "★ " if j == chosen.get(i, 0) else "   "
                lb.insert("end", "%s%d) %s ：%s" % (prefix, j + 1, o["src"], self._opt_sig(o)))
            if chosen.get(i, 0) < lb.size():
                lb.selection_set(chosen[i])

        def apply_choice(_=None):
            s = lb.curselection()
            sel = tv.selection()
            if not s or not sel:
                return
            j = s[0]
            for iid in sel:
                i = idx_of[iid]
                chosen[i] = j
                refresh_row(i)
            show_cands()
            update_count()

        def bulk(which):
            for i, (sheet, plan, c, ck) in enumerate(pending):
                chosen[i] = 0 if which == 0 else len(c["options"]) - 1
                refresh_row(i)
            update_count()

        def update_count():
            self._dlg_count.configure(text="共 %d 处冲突，已选 %d 处（直接“确定”= 每处默认用候选①）"
                                      % (len(pending), len(chosen)))

        tv.bind("<<TreeviewSelect>>", show_cands)
        lb.bind("<Double-1>", apply_choice)

        bar = ttk.Frame(win)
        bar.pack(fill="x", pady=8, padx=10)
        ttk.Button(bar, text="采用选中候选", command=apply_choice).pack(side="left")
        ttk.Button(bar, text="全部用第1个源表", command=lambda: bulk(0)).pack(side="left", padx=6)
        ttk.Button(bar, text="全部用最后一个源表", command=lambda: bulk(1)).pack(side="left")

        result = {"ok": False}

        def ok():
            for i, (sheet, plan, c, ck) in enumerate(pending):
                j = chosen.get(i, 0)
                resolved.setdefault(id(plan), {})[ck] = j
                self.conflict_cache[self._conflict_key(sheet, c)] = (
                    c["options"][j]["src"], self._opt_sig(c["options"][j]))
            result["ok"] = True
            win.destroy()

        ttk.Button(bar, text="确定", command=ok, style="Go.TButton").pack(side="right")
        ttk.Button(bar, text="取消", command=win.destroy).pack(side="right", padx=6)
        update_count()
        win.wait_window()
        return result["ok"]

    def do_preview(self):
        if not self._check():
            return
        try:
            plans, skipped = self.make_plans()
        except Exception:
            messagebox.showerror("预演失败", traceback.format_exc())
            return
        if not self.resolve_all_conflicts(plans):
            return
        self.last_plans = plans
        self.log(self.txt_fill, "———— 预演结果（共 %d 个工作表） ————" % len(plans))
        tot = {"修改单元格": 0, "新增行": 0, "未匹配": 0}
        for dst, plan in plans:
            st = plan["stats"]
            tot["修改单元格"] += st["将修改单元格"]
            tot["新增行"] += st["将新增行"]
            tot["未匹配"] += st["目标表未匹配行"] + st["源表未匹配行"]
            self.log(self.txt_fill, "  「%s」 %s" % (dst.sheet, st))
            if plan["dup_dst"]:
                self.log(self.txt_fill, "     ⚠ 该表有重复键 %d 组（只按第一条匹配）" % len(plan["dup_dst"]))
            nm = plan.get("name_mismatch") or []
            if nm:
                ex = "、".join("%s≠%s" % (a[2], a[1]) for a in nm[:3])
                self.log(self.txt_fill,
                         "     ⚠ 同键异名 %d 条（键相同但姓名不同，疑似源表错行）%s：%s"
                         % (len(nm), "，已跳过不写入" if self.var_name_skip.get() else "，仍会写入",
                            ex))
        if skipped:
            for w in skipped:
                self.log(self.txt_fill, "  ⚠ %s" % w)
        self.log(self.txt_fill, "  合计：修改 %d 处，新增 %d 行，未匹配 %d 行"
                 % (tot["修改单元格"], tot["新增行"], tot["未匹配"]))
        self.preview_window(plans)

    def preview_window(self, plans):
        win = tk.Toplevel(self)
        win.title("预演结果预览（未写任何文件）")
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        win.geometry("%dx%d" % (int(sw * 0.66), int(sh * 0.66)))
        info = "工作表 %d 个 ｜ 修改 %d 处 ｜ 新增 %d 行" % (
            len(plans),
            sum(len(p["changes"]) for _, p in plans),
            sum(len(p["new_rows"]) for _, p in plans))
        ttk.Label(win, text=info, style="Hint.TLabel").pack(anchor="w", padx=10, pady=6)
        nb = ttk.Notebook(win)
        nb.pack(fill="both", expand=True, padx=8, pady=4)

        def sheet(cols, rows, title):
            f = ttk.Frame(nb)
            tv = ttk.Treeview(f, columns=[str(i) for i in range(len(cols))], show="headings")
            for i, c in enumerate(cols):
                tv.heading(str(i), text=c)
                tv.column(str(i), width=max(70, 640 // max(1, len(cols))))
            for r in rows[:1000]:
                tv.insert("", "end", values=r)
            sb = ttk.Scrollbar(f, orient="vertical", command=tv.yview)
            tv.configure(yscrollcommand=sb.set)
            tv.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")
            nb.add(f, text="%s (%d)" % (title, len(rows)))

        ch_rows, nr_rows, ud_rows, us_rows = [], [], [], []
        for dst, plan in plans:
            for c in plan["changes"]:
                ch_rows.append((dst.sheet, c["excel_row"], c["colname"],
                                cell_text(c["old"]), cell_text(c["new"]), c["key"]))
            for n in plan["new_rows"]:
                nr_rows.append((dst.sheet, n["key"], n["src"],
                                ", ".join("%s=%s" % (dst.headers[k], cell_text(v))
                                          for k, v in n["values"].items())))
            for k, r in plan["unmatched_dst"]:
                ud_rows.append((dst.sheet, k, r))
            for x in plan["unmatched_src"]:
                us_rows.append((dst.sheet,) + tuple(x))

        sheet(["工作表", "目标表行号", "列名", "原值", "新值", "匹配键"], ch_rows, "将修改")
        sheet(["工作表", "匹配键", "来源", "写入字段"], nr_rows, "将新增")
        sheet(["工作表", "匹配键", "目标表行号"], ud_rows, "目标表未匹配")
        sheet(["工作表", "来源", "匹配键", "首个键值"], us_rows, "源表未匹配")
        ttk.Button(win, text="关闭", command=win.destroy).pack(pady=6)

    def do_execute(self):
        if not self._check():
            return
        try:
            plans, skipped = self.make_plans()
        except Exception:
            messagebox.showerror("计算失败", traceback.format_exc())
            return
        if not self.resolve_all_conflicts(plans):
            return
        n_ch = sum(len(p["changes"]) for _, p in plans)
        n_new = sum(len(p["new_rows"]) for _, p in plans)
        if not n_ch and not n_new:
            messagebox.showinfo("提示", "没有任何需要写入的内容，请检查匹配键/映射/模式")
            return
        src_path = self.var_dst_path.get().strip()
        stem, ext = os.path.splitext(src_path)
        suggest = self.var_out.get().strip() or "%s_补全结果_%s%s" % (stem, now_stamp(), ext)
        if not self.var_keep.get():
            suggest = src_path
        out = self.ask_output_path(suggest, src_path, key="out")
        if not out:
            return
        self.var_out.set(out)
        try:
            final = execute_fill_multi(src_path, plans, out)
        except PermissionError as e:
            messagebox.showerror("写入失败（文件被占用）",
                                 "%s\n\n请关掉 Excel 中已打开的该文件，或换一个输出文件名后重试。" % e)
            return
        except Exception:
            messagebox.showerror("写入失败", traceback.format_exc())
            return
        self.log(self.txt_fill, "✔ 已写入：%s（%d 个工作表，修改 %d 处，新增 %d 行）"
                 % (final, len(plans), n_ch, n_new))
        if self.var_export_log.get():
            rep = os.path.splitext(final)[0] + "_明细清单.xlsx"
            try:
                export_report_multi(plans, rep)
                self.log(self.txt_fill, "✔ 明细清单：%s" % rep)
            except Exception as e:
                self.log(self.txt_fill, "明细清单导出失败：%s" % e)
        messagebox.showinfo("完成", "已保存：\n%s\n\n工作表 %d 个，修改 %d 处，新增 %d 行"
                            % (final, len(plans), n_ch, n_new))

    def pick_out_file(self):
        p = self.ask_save("out", defaultextension=".xlsx",
                          filetypes=[("Excel", "*.xlsx")])
        if p:
            self.var_out.set(p)

    # ---------- 规则保存 ----------
    def save_rule(self):
        p = self.ask_save("rule", defaultextension=".json",
                          filetypes=[("规则文件", "*.json")])
        if not p:
            return
        rule = {
            "version": 2,
            "app": APP_TITLE,
            "target": {"path": self.var_dst_path.get().strip(),
                       "sheets": self.dst_selected_sheets(),
                       "header_row": self.var_hdr_row.get(),
                       "header_rows": self.var_hdr_span.get()},
            "sources": [{"path": s.path, "sheet": s.sheet,
                         "header_row": s.header_row, "header_rows": s.header_rows}
                        for s in self.sources],
            "pairs": self.pairs,
            "key_pairs": self.key_pairs,
            "key_options": self._kopt(),
            "mappings": self.mappings,
            "mode": self.var_mode.get(),
            "keep_original": self.var_keep.get(),
        }
        with open(p, "w", encoding="utf-8") as f:
            json.dump(rule, f, ensure_ascii=False, indent=2)
        self.log(self.txt_fill, "✔ 规则已保存：%s" % p)

    def load_rule(self):
        p = self.ask_open("rule", filetypes=[("规则文件", "*.json")])
        if not p:
            return
        self.load_rule_file(p)

    def load_rule_file(self, p):
        with open(p, encoding="utf-8") as f:
            rule = json.load(f)
        t = rule.get("target") or {}
        if t.get("path"):
            self.var_dst_path.set(t["path"])
            self.var_hdr_row.set(str(t.get("header_row") or "自动"))
            self.var_hdr_span.set(str(t.get("header_rows") or "自动"))
            try:
                self.dst_sheets = list_sheets(t["path"])
            except Exception as e:
                messagebox.showerror("读取失败", str(e))
                return
            self.dst_cache = {}
            self.lb_dst.delete(0, "end")
            for n in self.dst_sheets:
                self.lb_dst.insert("end", n)
            want = t.get("sheets") or self.dst_sheets
            for i, n in enumerate(self.dst_sheets):
                if n in want:
                    self.lb_dst.selection_set(i)
        self.sources = []
        for s in rule.get("sources", []):
            try:
                self.sources.append(open_table(s["path"], s["sheet"],
                                               s.get("header_row"), s.get("header_rows")))
            except Exception as e:
                self.log(self.txt_fill, "规则中的源表加载失败：%s" % e)
        self.lb_src.delete(0, "end")
        for sd in self.sources:
            self.lb_src.insert("end", "%s  ｜表头行%d ｜%d行" % (sd.label, sd.header_row, len(sd.data_rows)))
        if self.sources:
            self.cb_map_src["values"] = self.sources[0].col_labels()
            self.cb_keysrc["values"] = self.sources[0].col_labels()
        # 恢复对应关系（校验下标/工作表仍然存在）
        pairs = rule.get("pairs")
        if isinstance(pairs, list) and pairs and isinstance(pairs[0], dict):
            self.pairs = [{"src": int(x.get("src", 0)), "target": x.get("target")}
                          for x in pairs if int(x.get("src", 0)) < len(self.sources)]
            self.refresh_pairs_tree()
        else:
            self.pairs = []
        self.rebuild_pairs()
        self.key_pairs = []
        for item in rule.get("key_pairs", []):
            a, b = item[0], item[1]
            if isinstance(a, int):      # 旧规则：下标 → 列名
                a = self.sources[0].headers[a] if self.sources and a < len(self.sources[0].headers) else str(a)
            if isinstance(b, int):
                b = self.dst.headers[b] if self.dst and b < len(self.dst.headers) else str(b)
            self.key_pairs.append((a, b))
        self.tv_keys.delete(*self.tv_keys.get_children())
        for kname, vname in self.key_pairs:
            self.tv_keys.insert("", "end", values=(kname, vname))
        self.mappings = rule.get("mappings", [])
        self.refresh_map_tree()
        ko = rule.get("key_options", {})
        self.var_k_nospace.set(ko.get("nospace", True))
        self.var_k_upper.set(ko.get("upper", True))
        self.var_k_symbol.set(ko.get("symbol", False))
        self.var_k_digits.set(ko.get("digits", False))
        self.var_mode.set(rule.get("mode", "fill_add"))
        self.var_keep.set(rule.get("keep_original", True))
        self.log(self.txt_fill, "✔ 已加载规则：%s（%d 个源表 / %d 条映射）"
                 % (os.path.basename(p), len(self.sources), len(self.mappings)))

def main():
    enable_dpi_awareness()
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
