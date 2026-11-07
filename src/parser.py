# -*- coding: utf-8 -*-
"""消息解析：从自然语言里抽出「条码 + 数量」。

刻意不接大模型——这类格式相对固定的指令，正则比模型更快、更便宜、更可预测，
而且失败时能给出明确的错误原因（"条码 123 未识别到数量"）而不是一次幻觉。

支持的写法（举例）：
    加单 6900000000001 100
    @MyBot 补单 6900000000001 x 100
    加单 6900000000001 100, 6900000000002 200
    加单
    6900000000001 100
    6900000000002 200
"""
import re

# 条码长度范围（EAN-13 为 13 位，兼容 8~14 位的自编码）
BC_MIN_LEN = 8
BC_MAX_LEN = 14
# 单次加单数量上限，用来把"数量"和"条码"区分开
MAX_QTY = 999_999

_BC_RE = re.compile(rf"(?<!\d)(\d{{{BC_MIN_LEN},{BC_MAX_LEN}}})(?!\d)")
_NUM_RE = re.compile(r"(?<!\d)(\d+)(?!\d)")


def has_intent(text):
    """消息是否包含加单意图词。"""
    from .config import CFG
    return any(w in (text or "") for w in getattr(CFG, "INTENT_WORDS", ["加单"]))


def is_at_bot(text):
    """群消息是否 @ 了机器人。

    飞书 API 返回的 content 里，@ 已被渲染成纯文本（如 "@MyBot 加单 ..."），
    所以按机器人显示名做子串匹配即可。
    """
    from .config import CFG
    text = text or ""
    for name in getattr(CFG, "BOT_NAMES", []):
        if name and f"@{name}" in text:
            return True
    return False


def _pick_qty(segment):
    """从一段文本里挑一个合法的数量。"""
    for raw in _NUM_RE.findall(segment):
        val = int(raw)
        if 1 <= val <= MAX_QTY:
            return val
    return None


def parse_orders(text):
    """返回 {"orders": [{"barcode", "qty"}], "errors": [...]}。"""
    text = text or ""
    matches = list(_BC_RE.finditer(text))
    orders, errors, seen = [], [], set()

    for i, m in enumerate(matches):
        barcode = m.group(1)
        # 只看"当前条码到下一个条码之间"的文本，避免把下一单的数量算到这一单
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        tail = text[m.end():end]

        qty = _pick_qty(tail)
        if qty is None:
            # 兜底：数量可能写在条码前面（"加单 100 6900000000001"）
            head = text[max(0, m.start() - 30):m.start()]
            nums = [int(n) for n in _NUM_RE.findall(head) if 1 <= int(n) <= MAX_QTY]
            qty = nums[-1] if nums else None

        if qty is None:
            errors.append(f"条码 {barcode} 未识别到数量")
            continue
        if barcode in seen:
            continue
        seen.add(barcode)
        orders.append({"barcode": barcode, "qty": qty})

    return {"orders": orders, "errors": errors}
