# -*- coding: utf-8 -*-
"""离线自测：不连飞书、不连 ERP，只验证核心逻辑。

运行：
    python tests/selftest.py

覆盖：
  1. 消息解析（多种写法、数量缺失、多单）
  2. 条码尾零处理（rstrip(".0") 的经典 bug）
  3. 模板 fork + 示例行删除 + 空白 Sheet 清理
  4. 写入、同条码累加、回读校验
  5. Excel → ERP JSON 行映射
"""
import os
import shutil
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

# 测试需要 config.py；若不存在则临时用示例配置，结束时删除
_created_config = False
_cfg_path = os.path.join(ROOT, "config.py")
if not os.path.exists(_cfg_path):
    shutil.copy2(os.path.join(ROOT, "config.example.py"), _cfg_path)
    _created_config = True

FAILED = []


def check(name, cond, detail=""):
    if cond:
        print(f"  PASS  {name}")
    else:
        print(f"  FAIL  {name}  {detail}")
        FAILED.append(name)


def main():
    tmp = tempfile.mkdtemp(prefix="bridge_test_")
    try:
        from src.config import CFG
        from src import excel_engine, parser, erp_client

        # 测试期间把所有路径指向临时目录
        CFG.TEMPLATE_FILE = os.path.join(tmp, "template.xlsx")
        CFG.OUTPUT_DIR = os.path.join(tmp, "output")
        CFG.PRODUCT_CACHE = os.path.join(tmp, "cache.json")

        print("\n[1] 消息解析")
        r = parser.parse_orders("加单 6900000000001 100")
        check("单条", r["orders"] == [{"barcode": "6900000000001", "qty": 100}], r)

        r = parser.parse_orders("@MyBot 补单 6900000000001 x 100, 6900000000002 200")
        check("多条 + 分隔符", len(r["orders"]) == 2 and r["orders"][1]["qty"] == 200, r)

        r = parser.parse_orders("加单\n6900000000001 100\n6900000000002 200")
        check("多行", len(r["orders"]) == 2, r)

        r = parser.parse_orders("加单 6900000000001")
        check("缺数量应报错", not r["orders"] and r["errors"], r)

        r = parser.parse_orders("加单 6900000000001 100 6900000000002 200")
        check("数量不串单", [o["qty"] for o in r["orders"]] == [100, 200], r)

        print("\n[2] 条码尾零")
        check("去 .0", excel_engine._drop_dot0("6900000000010.0") == "6900000000010")
        check("不吃真尾零", excel_engine._drop_dot0("6900000000010") == "6900000000010")
        check("普通数字", excel_engine._drop_dot0(6900000000010) == "6900000000010")

        print("\n[3] 模板处理")
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(["预算开始日期", "经营体", "产品名称", "产品条码", "预算类型", "备注", "第1周", "数量"])
        ws.append(["2026-01-01", "示例经营体", "示例商品", "6900000000000", "临时计划", "", "", 1])
        wb.create_sheet("Sheet2")     # 空白页，应被清理
        wb.create_sheet("Sheet3")
        wb.save(CFG.TEMPLATE_FILE)
        wb.close()

        doc = excel_engine.fork_document("测试文档.xlsx")
        wb = openpyxl.load_workbook(doc)
        check("空白 Sheet 已清理", wb.sheetnames == ["Sheet1"], wb.sheetnames)
        wb.close()

        print("\n[4] 写入 / 累加 / 校验")
        items = [{"barcode": "6900000000001", "qty": 100, "name": "商品A"},
                 {"barcode": "6900000000002", "qty": 200, "name": "商品B"}]
        res = excel_engine.fill(items, doc)
        check("写入成功", res["ok"], res.get("errors"))
        check("两条记录", len([x for x in res["results"] if x["action"] == "appended"]) == 2)

        # 同条码再次写入 → 累加
        res2 = excel_engine.fill([{"barcode": "6900000000001", "qty": 50, "name": "商品A"}], doc)
        merged = [x for x in res2["results"] if x["action"] == "merged"]
        check("同条码累加", merged and merged[0]["qty_after"] == 150, res2["results"])

        wb = openpyxl.load_workbook(doc)
        ws = wb.worksheets[0]
        check("示例行已删除", ws.max_row == 3, f"max_row={ws.max_row}")
        check("日期为本周一", str(ws["A2"].value) == excel_engine.this_monday().strftime("%Y-%m-%d"))
        check("条码是文本格式", ws["D2"].number_format == "@")
        wb.close()

        print("\n[5] ERP 字段映射")
        rows = erp_client.xlsx_to_rows(doc)
        check("映射出 2 行", len(rows) == 2, rows)
        check("字段名正确", "productBarCode" in rows[0] and "quantity" in rows[0], rows[0])
        check("条码无 .0", rows[0]["productBarCode"] == "6900000000001", rows[0])

        print("\n" + ("=" * 40))
        if FAILED:
            print(f"FAILED: {len(FAILED)} 项 -> {FAILED}")
            return 1
        print("全部通过 ✓")
        return 0

    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        if _created_config:
            try:
                os.remove(_cfg_path)
            except OSError:
                pass


if __name__ == "__main__":
    sys.exit(main())
