# -*- coding: utf-8 -*-
"""ERP 适配器：把生成的 Excel 导入业务系统，并推进审批流。

这是整个项目里**最需要按你的系统改写**的文件。默认实现假设：
  - 导入接口接收 JSON 行数组（而不是 multipart 文件上传）
  - 推送接口接收 {"idList": [...], "operationType": 0}
如果你的系统是真正的文件上传（multipart/form-data），把 import_rows 换掉即可，
其余逻辑（认证、错误处理、字段映射）可以直接复用。

安全默认值：
  - ERP_AUTO_PUSH 默认 False，只导入不推送；
  - push() 必须显式传 confirm=True，避免脚本误触真实审批。
"""
import datetime
import json
import os
import urllib.error
import urllib.request

from .config import CFG, path

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")


def _auth_headers():
    mode = getattr(CFG, "ERP_AUTH_MODE", "cookie")
    if mode == "header":
        token = getattr(CFG, "AUTH_HEADER", "")
        return {"Authorization": token} if token else {}
    cookie_file = path(getattr(CFG, "COOKIE_FILE", "secrets/erp_cookie.txt"))
    try:
        with open(cookie_file, encoding="utf-8") as f:
            return {"Cookie": f.read().strip()}
    except FileNotFoundError:
        raise RuntimeError(
            f"找不到凭证文件 {cookie_file}。\n"
            "请把浏览器登录后的 Cookie 整串保存到该文件（该文件已被 .gitignore 忽略）"
        )


def request(path_name, payload=None, method="POST", timeout=60):
    base = getattr(CFG, "ERP_BASE_URL", "").rstrip("/")
    if not base:
        return {"success": False, "errMsg": "未配置 ERP_BASE_URL"}

    url = base + path_name
    headers = {
        "Content-Type": "application/json;charset=UTF-8",
        "User-Agent": UA,
        "Accept": "application/json",
    }
    headers.update(_auth_headers())
    referer = getattr(CFG, "ERP_REFERER", "")
    if referer:
        # 不少系统会校验 Referer 防盗链，缺了直接 403
        headers["Referer"] = referer

    data = json.dumps(payload, ensure_ascii=False).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, headers=headers, method=method)

    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            text = r.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "replace")[:500]
        return {"success": False, "errMsg": f"HTTP {e.code}: {body}"}
    except Exception as e:
        return {"success": False, "errMsg": f"请求异常: {e}"}

    try:
        obj = json.loads(text)
    except json.JSONDecodeError:
        return {"success": False, "errMsg": "返回非 JSON: " + text[:300]}

    ok = obj.get("success", obj.get("code") in (0, 200))
    if not ok:
        msg = obj.get("errMsg") or obj.get("message") or text[:300]
        return {"success": False, "errMsg": str(msg)[:500], "raw": obj}
    return {"success": True, "data": obj.get("data"), "raw": obj}


# ---------------------------------------------------------------- 字段映射

def _fmt_date(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    if len(s) == 8 and s.isdigit():        # 20260907
        return f"{s[:4]}-{s[4:6]}-{s[6:]}"
    return s[:10] if len(s) >= 10 else s


def _fmt_num(v):
    if v is None or v == "":
        return None
    try:
        f = float(v)
        return int(f) if f == int(f) else f
    except Exception:
        return None


def xlsx_to_rows(file_path):
    """按表头名把 Excel 映射成 ERP 导入接口要的 JSON 行数组。"""
    import openpyxl

    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.worksheets[0]

    header = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v is not None:
            header[str(v).strip().lstrip("*")] = c

    field_map = getattr(CFG, "FIELD_MAP", {})
    col_of = {}
    for title, field in field_map.items():
        key = title.lstrip("*").strip()
        if key in header:
            col_of[field] = header[key]
        else:
            # 宽松匹配：表头里包含即可（"产品条码（必填）"也能命中）
            for h, c in header.items():
                if key and key in h:
                    col_of[field] = c
                    break

    bc_field = "productBarCode"
    bc_col = col_of.get(bc_field)
    if not bc_col:
        wb.close()
        raise ValueError("表头里找不到「产品条码」列，请检查 FIELD_MAP 或模板")

    rows = []
    for r in range(2, ws.max_row + 1):
        bc = ws.cell(r, bc_col).value
        if bc is None or str(bc).strip() == "":
            continue
        row = {}
        for field, c in col_of.items():
            v = ws.cell(r, c).value
            if v is None:
                row[field] = None
            elif "date" in field.lower() or "Date" in field:
                row[field] = _fmt_date(v)
            elif field in ("quantity", "qty") or field.endswith("WeekNum"):
                row[field] = _fmt_num(v)
            elif field == bc_field:
                row[field] = str(v).strip()
            else:
                s = str(v).strip()
                row[field] = s or None
        if row.get(bc_field):
            rows.append(row)

    wb.close()
    return rows


# ---------------------------------------------------------------- 业务动作

def query_drafts(page_size=100):
    d = request(getattr(CFG, "ERP_LIST_PATH", "/api/items/page"),
                {"status": 0, "currentPage": 1, "pageSize": page_size})
    if not d["success"]:
        return d
    rows = d.get("data") or []
    return {"success": True, "rows": rows,
            "total": ((d.get("raw") or {}).get("query") or {}).get("total")}


def import_rows(rows):
    """导入数据行，成功后成为"草稿"。"""
    if not rows:
        return {"success": False, "errMsg": "没有可导入的数据行"}
    return request(getattr(CFG, "ERP_IMPORT_PATH", "/api/items/import"), rows)


def push(ids, operation_type=0, confirm=False):
    """推送审批。

    confirm=False 时直接拒绝执行——这是故意的，防止调试脚本误触真实审批流。
    """
    if not confirm:
        return {"success": False, "errMsg": "push() 需要显式 confirm=True"}
    if not ids:
        return {"success": False, "errMsg": "idList 为空，拒绝推送（避免误推全量）"}
    return request(getattr(CFG, "ERP_PUSH_PATH", "/api/items/push"),
                   {"idList": [int(i) for i in ids], "operationType": operation_type})


def import_and_push(file_path, confirm=False):
    """导入 + 推送一条龙。导入后按本次条码限定要推送的草稿，绝不"全选"。"""
    rows = xlsx_to_rows(file_path)
    if not rows:
        return {"ok": False, "errors": ["文件中没有可导入的数据行"]}

    r = import_rows(rows)
    if not r["success"]:
        return {"ok": False, "errors": [f"导入失败: {r['errMsg']}"]}

    if not getattr(CFG, "ERP_AUTO_PUSH", False):
        return {"ok": True, "imported": len(rows), "pushed": 0,
                "msg": "已导入为草稿（ERP_AUTO_PUSH=False，未推送）"}

    # 只推本次涉及条码的草稿，防止把同事的单一起推走
    barcodes = {row.get("productBarCode") for row in rows}
    drafts = query_drafts()
    if not drafts.get("success"):
        return {"ok": True, "imported": len(rows), "pushed": 0,
                "msg": f"导入成功，但查询草稿失败: {drafts.get('errMsg')}"}

    ids = [str(x["id"]) for x in drafts["rows"]
           if str(x.get("productBarCode")) in barcodes]
    if not ids:
        return {"ok": True, "imported": len(rows), "pushed": 0,
                "msg": "导入成功，但未匹配到草稿（可能需要人工确认）"}

    p = push(ids, 0, confirm=confirm)
    if not p["success"]:
        return {"ok": True, "imported": len(rows), "pushed": 0,
                "msg": f"导入成功，推送失败: {p['errMsg']}"}
    return {"ok": True, "imported": len(rows), "pushed": len(ids),
            "msg": f"导入 + 推送成功（{len(ids)} 条）"}
