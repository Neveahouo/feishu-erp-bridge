# -*- coding: utf-8 -*-
"""Excel 模板引擎：从模板 fork 新文档并写入数据。

为什么是"每次 fork 一份新文档"而不是维护一份总账？
  - 总账模式：多人协作时互相覆盖、历史数据越滚越大、出问题难回溯；
  - 一次一单：每次产物独立、可归档、可审计，失败也只影响这一单；
  - 对下游系统更友好：导入接口每次只收到本次变更，不会重复导入历史行。

三个必须处理的坑：
  1. 模板里的"示例行"是活数据，不删会被下游当成真实数据导入；
  2. 模板常带几个空白 Sheet，部分系统的导入器会判定"文件无效"；
  3. 条码被 Excel 当数字存成 6900000000001.0，用 rstrip(".0") 会误删真实的尾零，
     必须只去掉结尾的 ".0" 两个字符。
"""
import datetime
import os
import re
import shutil

import openpyxl
from openpyxl.utils import column_index_from_string

from .config import CFG, path

_INVALID_FN = re.compile(r'[\\/:*?"<>|\r\n\t]+')
_SPACES = re.compile(r"\s+")


# ---------------------------------------------------------------- 工具

def this_monday(d=None):
    d = d or datetime.date.today()
    return d - datetime.timedelta(days=d.weekday())


def _cell_str(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def _drop_dot0(s):
    """6900000000001.0 -> 6900000000001。

    注意不能用 rstrip(".0")：那是字符集语义，会把 "...100" 末尾的 0 也吃掉。
    """
    s = str(s).strip()
    return s[:-2] if s.endswith(".0") else s


def clean_filename(s, limit=40):
    s = _INVALID_FN.sub("_", str(s or "").strip())
    s = _SPACES.sub("", s).strip("._ ")
    return s[:limit] or "未命名"


def build_output_name(items):
    """items: [{"name", "qty"}] -> 加单+品名-份数-日期.xlsx"""
    tpl = getattr(CFG, "OUTPUT_NAME_TEMPLATE", "加单+{name}-{qty}-{date}.xlsx")
    date = datetime.date.today().strftime("%Y-%m-%d")
    total = sum(int(i.get("qty") or 0) for i in items)
    names = [clean_filename(i.get("name") or i.get("barcode")) for i in items]
    head = names[0] if len(items) == 1 else f"{names[0]}等{len(items)}种"
    return tpl.format(name=clean_filename(head, 50), qty=total, date=date)


# ---------------------------------------------------------------- 模板

def strip_empty_sheets(file_path):
    """删除模板里除第一个之外的所有空白 Sheet。"""
    try:
        wb = openpyxl.load_workbook(file_path)
        keep = wb.sheetnames[0]
        changed = False
        for name in list(wb.sheetnames):
            if name == keep:
                continue
            ws = wb[name]
            if ws.max_row <= 1 and ws.max_column <= 1:
                del wb[name]
                changed = True
        if changed:
            wb.save(file_path)
        wb.close()
    except Exception as e:
        print(f"[warn] 清理空白 Sheet 失败: {e}")


def fork_document(out_name):
    """从模板复制出一份新文档，返回绝对路径。"""
    tpl = path(getattr(CFG, "TEMPLATE_FILE", "templates/template.xlsx"))
    if not os.path.exists(tpl):
        raise FileNotFoundError(
            f"模板不存在: {tpl}\n请把你的 Excel 模板放到 templates/ 下，并在 config.py 里配置 TEMPLATE_FILE"
        )

    out_dir = path(getattr(CFG, "OUTPUT_DIR", "data/output"))
    os.makedirs(out_dir, exist_ok=True)
    dst = os.path.join(out_dir, out_name)
    if os.path.exists(dst):
        os.remove(dst)
    shutil.copy2(tpl, dst)
    strip_empty_sheets(dst)
    return dst


def _is_sample_row(ws, row):
    """判断某行是不是模板自带的示例行。

    必须**所有**标记同时命中才算示例行（all 而非 any）。
    原因：固定值列（如"经营体"）在每一行数据里都是同一个值，
    如果用 any 匹配，第二次写入时会把真实数据行误判成示例行删掉。
    建议在 SAMPLE_ROW_MARKERS 里至少配一个唯一标识（如示例条码）。
    """
    markers = getattr(CFG, "SAMPLE_ROW_MARKERS", {}) or {}
    if not markers:
        return False
    return all(_cell_str(ws[f"{col}{row}"].value) == marker
               for col, marker in markers.items())


# ---------------------------------------------------------------- 写入

def fill(items, file_path, date_str=None, dry_run=False, verbose=False):
    """把 items 写入文档。

    items: [{"barcode", "qty", "name"}]
    返回 {"ok", "results", "errors", "date"}
    """
    if not items:
        return {"ok": False, "results": [], "errors": ["没有待写入的数据"]}

    date_str = date_str or this_monday().strftime("%Y-%m-%d")
    rules = CFG.COLUMN_RULES
    qty_col = getattr(CFG, "QTY_COLUMN", "H")
    bc_col = getattr(CFG, "BARCODE_COLUMN", "D")
    date_col = getattr(CFG, "DATE_COLUMN", "A")
    data_start = getattr(CFG, "DATA_START_ROW", 2)

    wb = openpyxl.load_workbook(file_path)
    ws = wb.worksheets[0]

    # 1) 删除模板示例行
    if ws.max_row >= data_start and _is_sample_row(ws, data_start):
        if verbose:
            print(f"删除示例行 row={data_start}")
        if not dry_run:
            ws.delete_rows(data_start, 1)

    results = []
    for item in items:
        barcode = str(item.get("barcode") or "").strip()
        qty = int(item.get("qty") or 0)
        name = item.get("name") or ""

        # 2) 同日期 + 同条码 → 累加（保底幂等：重复消息不会翻倍，因为上层有 message_id 去重）
        target_row, old_qty = None, 0
        for r in range(data_start, ws.max_row + 1):
            if (_cell_str(ws[f"{date_col}{r}"].value)[:10] == date_str
                    and _drop_dot0(_cell_str(ws[f"{bc_col}{r}"].value)) == barcode):
                target_row = r
                try:
                    old_qty = int(float(ws[f"{qty_col}{r}"].value or 0))
                except Exception:
                    old_qty = 0
                break

        if target_row:
            new_qty = old_qty + qty
            if not dry_run:
                ws[f"{qty_col}{target_row}"].value = new_qty
            results.append({"barcode": barcode, "name": name, "action": "merged",
                            "row": target_row, "qty": qty, "qty_after": new_qty})
            continue

        # 3) 找空行，找不到就追加
        target_row = None
        for r in range(data_start, ws.max_row + 2):
            if (_cell_str(ws[f"{bc_col}{r}"].value) == ""
                    and _cell_str(ws[f"{date_col}{r}"].value) == ""):
                target_row = r
                break
        if target_row is None:
            target_row = ws.max_row + 1

        if not dry_run:
            for col, rule in rules.items():
                cell = ws[f"{col}{target_row}"]
                kind = rule.get("type")
                if kind == "date":
                    cell.value = date_str
                elif kind == "fixed":
                    cell.value = rule.get("value")
                elif kind == "barcode":
                    cell.value = barcode
                    cell.number_format = "@"   # 强制文本，防止 Excel 转科学计数法
                elif kind == "product_name":
                    cell.value = name
                elif kind == "quantity":
                    cell.value = qty

        results.append({"barcode": barcode, "name": name, "action": "appended",
                        "row": target_row, "qty": qty, "qty_after": qty})

    if dry_run:
        wb.close()
        return {"ok": True, "results": results, "date": date_str, "dry_run": True}

    wb.save(file_path)
    wb.close()

    # 4) 回读校验：写进去的东西必须能读出来，否则宁可报错也不要发错文件
    errs = _verify(file_path, results, date_str, date_col, bc_col, qty_col)
    if errs:
        return {"ok": False, "results": results, "errors": ["回读校验失败: " + "; ".join(errs)],
                "date": date_str}
    return {"ok": True, "results": results, "date": date_str, "file": file_path}


def _verify(file_path, results, date_str, date_col, bc_col, qty_col):
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    errs = []
    for r in results:
        row = r["row"]
        got_date = _cell_str(ws[f"{date_col}{row}"].value)[:10]
        got_bc = _drop_dot0(_cell_str(ws[f"{bc_col}{row}"].value))
        try:
            got_qty = int(float(ws[f"{qty_col}{row}"].value))
        except Exception:
            got_qty = None
        if got_date != date_str or got_bc != r["barcode"] or got_qty != r["qty_after"]:
            errs.append(f"row{row} 日期={got_date} 条码={got_bc} 数量={got_qty}")
    wb.close()
    return errs
