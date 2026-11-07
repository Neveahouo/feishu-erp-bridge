# -*- coding: utf-8 -*-
"""商品名称查询：在线接口优先，本地缓存兜底。

设计要点：
  - 在线查询保证名称是最新的（业务系统里改名/新建品能立刻反映）；
  - 本地缓存保证在线接口挂掉时仍能出单（可用性优先于完整性）；
  - 两者都没命中就让这一条失败并明确告知，绝不猜一个名字写进表里。
"""
import json
import os
import urllib.parse
import urllib.request

from .config import CFG, path

_cache = None


def _load_cache():
    global _cache
    if _cache is not None:
        return _cache
    _cache = {}
    p = path(getattr(CFG, "PRODUCT_CACHE", "data/product_cache.json"))
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as f:
                _cache = json.load(f)
        except Exception:
            _cache = {}
    return _cache


def _save_cache():
    p = path(getattr(CFG, "PRODUCT_CACHE", "data/product_cache.json"))
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(_cache, f, ensure_ascii=False)


def _query_online(barcode):
    """按配置的 PRODUCT_API 查询。返回名称或 None。"""
    api = getattr(CFG, "PRODUCT_API", "")
    if not api:
        return None
    url = api.replace("{barcode}", urllib.parse.quote(str(barcode)))
    try:
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0",
            "Accept": "application/json",
        })
        with urllib.request.urlopen(req, timeout=10) as r:
            data = json.loads(r.read().decode("utf-8", "replace"))
    except Exception:
        return None

    # 兼容常见的返回结构：{"data": {...}} / {"data": [...]} / 直接对象
    node = data.get("data", data) if isinstance(data, dict) else data
    candidates = []
    if isinstance(node, dict):
        candidates.append(node)
    elif isinstance(node, list):
        candidates.extend([x for x in node if isinstance(x, dict)])

    for item in candidates:
        for key in ("productName", "name", "itemName"):
            val = item.get(key)
            if val:
                # 精确匹配优先：避免模糊搜索结果里带出同款不同规格
                code = str(item.get("productBarCode") or item.get("barcode") or "")
                if not code or code == str(barcode):
                    return str(val).strip()
    return None


def lookup(barcode):
    """返回商品名称，找不到返回 None。"""
    barcode = str(barcode).strip()

    name = _query_online(barcode)
    if name:
        cache = _load_cache()
        if cache.get(barcode) != name:
            cache[barcode] = name
            _save_cache()
        return name

    return _load_cache().get(barcode)


def refresh_from_csv(csv_path):
    """从 CSV 重建本地缓存（列名需含「产品条码」「产品名称」）。"""
    import csv
    global _cache
    _cache = {}
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            bc = (row.get("产品条码") or row.get("条码") or "").strip()
            nm = (row.get("产品名称") or row.get("名称") or "").strip()
            if bc and nm:
                _cache[bc] = nm
    _save_cache()
    return len(_cache)
