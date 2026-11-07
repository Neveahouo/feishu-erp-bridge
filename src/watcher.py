# -*- coding: utf-8 -*-
"""主程序：轮询飞书 → 解析指令 → 生成文档 → 回传 → （可选）推送 ERP。

用法：
    python -m src.watcher --once              # 轮询一次（适合交给任务计划程序）
    python -m src.watcher --loop              # 常驻轮询
    python -m src.watcher --dry-run           # 只检测和解析，不生成文档不回复
    python -m src.watcher --parse "加单 6900000000001 100"   # 只测解析
    python -m src.watcher --seed              # 把当前可见消息标记为已处理（首次部署用）
"""
import argparse
import datetime
import html
import json
import os
import re
import time

from . import erp_client, excel_engine, lark, parser, product_lookup
from .config import CFG, path


# ---------------------------------------------------------------- 日志

def log(*args):
    line = "[{}] {}".format(
        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        " ".join(str(a) for a in args),
    )
    print(line, flush=True)
    try:
        log_file = path(getattr(CFG, "LOG_FILE", "data/bridge.log"))
        os.makedirs(os.path.dirname(log_file), exist_ok=True)
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


# ---------------------------------------------------------------- 幂等状态

def load_state():
    p = path(getattr(CFG, "STATE_FILE", "data/state.json"))
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"processed_ids": []}


def save_state(processed):
    p = path(getattr(CFG, "STATE_FILE", "data/state.json"))
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump({"processed_ids": list(processed)[-1000:]}, f, ensure_ascii=False)


# ---------------------------------------------------------------- 消息筛选

def strip_html(s):
    if not s:
        return ""
    return html.unescape(re.sub(r"<[^>]+>", " ", s))


def _sender_type(m):
    s = m.get("sender") or {}
    return s.get("sender_type") or s.get("type") or ("bot" if s.get("open_bot_id") else "")


def _sender_id(m):
    s = m.get("sender") or {}
    return s.get("id") or s.get("open_id") or ""


def eligible(m, kind):
    """判断一条消息要不要处理。返回 None 或 {"text", "sender_id"}。"""
    if not m.get("message_id") or m.get("msg_type") == "system":
        return None
    if _sender_type(m) not in ("", "user"):
        return None                          # 忽略其他机器人/系统的消息
    sid = _sender_id(m)
    if not sid:
        return None
    if not getattr(CFG, "ALLOW_ANY_SENDER", True) and sid != getattr(CFG, "OWNER_UID", ""):
        return None

    text = strip_html(m.get("content", "") or "")
    if not parser.has_intent(text):
        return None
    if kind == "group" and getattr(CFG, "GROUP_REQUIRE_AT", True) and not parser.is_at_bot(text):
        return None                          # 群里只响应 @机器人
    return {"text": text, "sender_id": sid}


# ---------------------------------------------------------------- 回复文案

def build_reply(result, file_name=None):
    lines = []
    ok_rows = [r for r in result.get("results", []) if r.get("action") in ("appended", "merged")]
    bad_rows = [r for r in result.get("results", []) if r.get("action") not in ("appended", "merged")]

    if ok_rows:
        if file_name:
            lines.append(f"✅ 已生成 **{file_name}**")
        else:
            lines.append("✅ 已写入模板")
        lines.append(f"📅 日期：{result.get('date')}")
        for r in ok_rows:
            act = "累加" if r["action"] == "merged" else "新增"
            lines.append(f"- `{r['barcode']}` {r.get('name', '')} × {r['qty']} → {act}（第 {r['row']} 行）")

    for r in bad_rows:
        lines.append(f"❌ `{r.get('barcode')}`：{r.get('error')}")

    if result.get("errors") and not ok_rows:
        lines = ["❌ 处理失败："] + [f"- {e}" for e in result["errors"]]

    if not lines:
        lines = ["⚠️ 没看懂这条指令。格式：`@<机器人名> 加单 <条码> <数量>`"]
    return "\n".join(lines)


# ---------------------------------------------------------------- 主循环

def poll_once(dry_run=False):
    state = load_state()
    processed = set(state.get("processed_ids", []))

    chats = lark.discover_chats()
    hits = []

    for chat_id, kind in chats:
        try:
            messages = lark.list_messages(chat_id, kind, limit=40)
        except Exception as e:
            log(f"拉取失败 {chat_id} ({kind}): {e}")
            continue
        for m in messages:
            mid = m.get("message_id")
            if not mid or mid in processed:
                continue
            info = eligible(m, kind)
            if info:
                hits.append((chat_id, kind, mid, info["sender_id"], info["text"]))

    if not hits:
        if not dry_run:
            save_state(processed)
        log(f"无新指令（扫描 {len(chats)} 个会话）")
        return {"triggered": 0, "chats": len(chats)}

    if dry_run:
        log(f"DRY-RUN 命中 {len(hits)} 条（不执行）")
        return {"triggered": len(hits), "hits": [
            {"chat": c, "kind": k, "parse": parser.parse_orders(t)} for c, k, _, _, t in hits
        ]}

    for chat_id, kind, mid, sender_id, text in hits:
        parsed = parser.parse_orders(text)
        orders = parsed["orders"]
        log(f"处理 {mid} [{kind}] 指令={orders} 解析错误={parsed['errors']}")

        if not orders:
            lark.send_text(chat_id, build_reply(
                {"results": [], "errors": parsed["errors"] or ["没有识别到条码+数量"]}))
            processed.add(mid)
            continue

        # 1) 查商品名称（在线优先，本地兜底）
        valid, name_map, missing = [], {}, []
        for o in orders:
            name = product_lookup.lookup(o["barcode"])
            if name:
                o["name"] = name
                name_map[o["barcode"]] = name
                valid.append(o)
            else:
                missing.append(f"条码 {o['barcode']} 查不到商品名称，已跳过")

        if not valid:
            lark.send_text(chat_id, build_reply(
                {"results": [], "errors": missing + parsed["errors"]}))
            processed.add(mid)
            continue

        # 2) 从模板 fork 一份新文档并写入
        out_name = excel_engine.build_output_name(valid)
        try:
            doc = excel_engine.fork_document(out_name)
        except Exception as e:
            lark.send_text(chat_id, f"❌ 生成文档失败：{e}")
            processed.add(mid)
            continue

        result = excel_engine.fill(valid, doc)
        for err in (missing + parsed["errors"]):
            result.setdefault("results", []).append(
                {"barcode": "-", "action": "error", "error": err})

        ok_rows = [r for r in result.get("results", [])
                   if r.get("action") in ("appended", "merged")]
        if not ok_rows:
            try:
                os.remove(doc)               # 写失败不留空模板
            except OSError:
                pass

        # 3) 可选：导入 + 推送 ERP
        erp_msg = None
        if ok_rows and getattr(CFG, "ERP_AUTO_PUSH", False):
            try:
                erp = erp_client.import_and_push(doc, confirm=True)
                erp_msg = erp.get("msg")
                log(f"ERP: {erp_msg}")
            except Exception as e:
                erp_msg = f"ERP 推送异常: {e}"
                log(erp_msg)

        # 4) 回执文本 + 文件
        reply = build_reply(result, file_name=out_name if ok_rows else None)
        if erp_msg:
            reply += f"\n🏭 ERP：{erp_msg}"
        sent, _ = lark.send_text(chat_id, reply)
        if not sent:
            lark.send_dm(reply, sender_id)

        if ok_rows:
            ok_file, err = lark.send_file(chat_id, doc, display_name=out_name)
            log(f"文件回传 ok={ok_file}" + (f" {err}" if err else ""))
        processed.add(mid)

    save_state(processed)
    return {"triggered": len(hits)}


def seed():
    """首次部署：把历史消息全部标记为已处理，避免开机瞬间刷屏。"""
    state = load_state()
    processed = set(state.get("processed_ids", []))
    n = 0
    for chat_id, kind in lark.discover_chats():
        try:
            messages = lark.list_messages(chat_id, kind, limit=50)
        except Exception:
            continue
        for m in messages:
            mid = m.get("message_id")
            if mid and mid not in processed:
                processed.add(mid)
                n += 1
    save_state(processed)
    log(f"--seed 标记 {n} 条历史消息")
    return {"seeded": n}


def main():
    ap = argparse.ArgumentParser(description="飞书 → ERP 自动化桥")
    ap.add_argument("--once", action="store_true", help="轮询一次（默认）")
    ap.add_argument("--loop", action="store_true", help="常驻轮询")
    ap.add_argument("--interval", type=int, default=getattr(CFG, "POLL_INTERVAL", 60))
    ap.add_argument("--dry-run", action="store_true", help="只检测不执行")
    ap.add_argument("--seed", action="store_true", help="标记历史消息为已处理")
    ap.add_argument("--parse", help="只测试一条文本的解析结果")
    args = ap.parse_args()

    if args.parse:
        print(json.dumps(parser.parse_orders(args.parse), ensure_ascii=False, indent=2))
        return

    if args.seed:
        print(json.dumps(seed(), ensure_ascii=False))
        return

    log(f"启动 | 群自动发现={getattr(CFG, 'AUTO_DISCOVER_GROUPS', True)} "
        f"强制@={getattr(CFG, 'GROUP_REQUIRE_AT', True)} "
        f"自动推送={getattr(CFG, 'ERP_AUTO_PUSH', False)}")

    if args.loop:
        while True:
            try:
                poll_once(dry_run=args.dry_run)
            except KeyboardInterrupt:
                raise
            except Exception as e:
                log(f"轮询异常（继续）: {e}")
            time.sleep(args.interval)
    else:
        print(json.dumps(poll_once(dry_run=args.dry_run), ensure_ascii=False, default=str))


if __name__ == "__main__":
    main()
