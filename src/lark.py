# -*- coding: utf-8 -*-
"""飞书 IM 封装（基于 lark-cli）。

为什么不用飞书官方 SDK / 事件订阅？
  - 事件订阅需要公网回调地址 + 验签 + 证书，个人或内网场景成本过高；
  - 轮询实现简单、零公网依赖、断线重启后能补齐未处理消息（配合 state 幂等）。
代价是最长 POLL_INTERVAL 秒的延迟，对"加单"这类场景完全够用。

三个踩过的坑，直接决定实现方式：
  1. Windows 上 lark-cli 是 .cmd 脚本。Python 用列表传参时，部分参数会被
     cmd 层吞掉（如 --sheet-id 神秘消失）。必须用 bash -lc + shlex.quote 包一层。
  2. 机器人(bot)身份**不能读群历史消息**（飞书返回 230027 access denied），
     但能读私聊、能发消息。所以：读群用 user 身份，读私聊用 bot 身份，发消息用 bot。
  3. 发文件时 lark-cli 拒绝绝对路径，必须切到工作目录用相对路径。
"""
import json
import os
import shlex
import shutil
import subprocess

from .config import CFG, path

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _resolve_cli():
    for name in ("lark-cli", "lark-cli.cmd"):
        p = shutil.which(name)
        if p:
            return p.replace("\\", "/")
    # 兜底：按实际安装位置改这里，或用环境变量 LARK_CLI
    return os.environ.get("LARK_CLI", "lark-cli")


LARK_CLI = _resolve_cli()


def run(args, timeout=60, cwd=None):
    """调用 lark-cli 并解析 JSON 输出。

    返回 {"ok": True, ...} 或 {"ok": False, "error": "..."}。
    """
    cmd = f"{LARK_CLI} " + " ".join(shlex.quote(str(a)) for a in args)
    try:
        r = subprocess.run(
            ["bash", "-lc", cmd],
            capture_output=True,
            timeout=timeout,
            cwd=cwd or ROOT,
        )
    except subprocess.TimeoutExpired:
        return {"ok": False, "error": f"超时 {timeout}s"}
    except Exception as e:
        return {"ok": False, "error": f"启动失败: {e}"}

    raw = r.stdout
    out = raw.decode("utf-8", "replace")
    if r.returncode != 0:
        # Windows 控制台默认 GBK，utf-8 解不出来时回退
        try:
            out = raw.decode("gbk", "replace")
        except Exception:
            pass
        err = (r.stderr or b"").decode("utf-8", "replace")[:500]
        return {"ok": False, "rc": r.returncode, "error": err or out[:300]}

    out = out.strip()
    if not out:
        return {"ok": False, "error": "空输出"}
    try:
        return json.loads(out)
    except json.JSONDecodeError:
        return {"ok": False, "error": "非 JSON 输出: " + out[:300]}


# ---------------------------------------------------------------- 会话

def discover_chats():
    """返回 [(chat_id, kind)]，kind 为 "p2p" 或 "group"。"""
    chats = []
    for cid in getattr(CFG, "P2P_CHAT_IDS", []):
        if cid:
            chats.append((cid, "p2p"))

    if getattr(CFG, "AUTO_DISCOVER_GROUPS", True):
        d = run(["im", "+chat-list", "--as", "bot", "--types", "group",
                 "--page-size", "100", "--format", "json"])
        if d.get("ok"):
            for c in (d.get("data") or {}).get("chats") or []:
                if c.get("chat_id"):
                    chats.append((c["chat_id"], "group"))
        else:
            # 群发现失败不致命，私聊仍可工作
            print(f"[warn] 群列表获取失败: {d.get('error')}")
    return chats


def list_messages(chat_id, kind, limit=40):
    """拉取会话消息。

    群聊必须用 user 身份（bot 读群历史会 230027），私聊必须用 bot 身份。
    """
    identity = "user" if kind == "group" else "bot"
    d = run(["im", "+chat-messages-list", "--chat-id", chat_id,
             "--as", identity, "--page-limit", str(limit), "--format", "json"])
    if not d.get("ok"):
        raise RuntimeError(f"拉取消息失败({chat_id}, as {identity}): {d.get('error')}")
    return (d.get("data") or {}).get("messages") or []


# ---------------------------------------------------------------- 发送

def send_text(chat_id, text):
    d = run(["im", "+messages-send", "--chat-id", chat_id,
             "--as", "bot", "--markdown", text])
    return d.get("ok"), d.get("error")


def send_dm(text, user_id=None):
    """私信兜底：群消息发不出去时，改用私聊送达。"""
    uid = user_id or getattr(CFG, "OWNER_UID", "")
    if not uid:
        return False, "未配置 OWNER_UID"
    for identity in ("bot", "user"):
        d = run(["im", "+messages-send", "--user-id", uid,
                 "--as", identity, "--markdown", text])
        if d.get("ok"):
            return True, identity
    return False, "私信失败"


def send_file(chat_id, file_path, display_name=None):
    """把文件作为文件消息发到会话。

    lark-cli 拒绝绝对路径，所以先复制到 workdir 下的相对目录再发。
    """
    workdir = path(getattr(CFG, "OUTPUT_DIR", "data/output"))
    os.makedirs(workdir, exist_ok=True)
    name = display_name or os.path.basename(file_path)
    dst = os.path.join(workdir, name)
    shutil.copy2(file_path, dst)

    rel_root = os.path.relpath(dst, ROOT).replace("\\", "/")
    d = run(["im", "+messages-send", "--chat-id", chat_id, "--as", "bot",
             "--file", "./" + rel_root], cwd=ROOT)
    return d.get("ok"), d.get("error")
