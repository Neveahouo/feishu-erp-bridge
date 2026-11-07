# -*- coding: utf-8 -*-
"""
配置示例 —— 复制为 config.py 后修改。config.py 已在 .gitignore 中，不会入库。

所有值都可以用环境变量覆盖（见下方 os.environ.get 写法），
这样同一份代码可以在测试机 / 生产机跑不同配置。
"""
import os

# ============================================================
# 一、飞书（IM 侧）
# ============================================================

# 机器人在群里的显示名。群消息必须 @ 到这些名字才会被处理。
# 注意：飞书 OpenAPI 返回的消息 content 里，@ 是**纯文本**（"@MyBot 加单 ..."），
#      不是 <at user_id="..."> 标签，所以这里按名字做子串匹配即可。
BOT_NAMES = [x.strip() for x in os.environ.get("BRIDGE_BOT_NAMES", "MyBot").split(",") if x.strip()]

# 私聊会话 ID（可选）。留空则只处理群消息。
# 获取方式：lark-cli im +chat-list --as bot
P2P_CHAT_IDS = [x.strip() for x in os.environ.get("BRIDGE_P2P", "").split(",") if x.strip()]

# 是否自动监听机器人所在的所有群（拉进新群即生效，无需改配置）
AUTO_DISCOVER_GROUPS = True

# 群消息是否强制要求 @ 机器人。建议 True：避免群里的闲聊误触发。
GROUP_REQUIRE_AT = True

# 是否允许群内任何人触发。False = 只允许 OWNER_UID 触发（更严格）。
ALLOW_ANY_SENDER = True

# 兜底私信接收人（群消息发不出去时降级到这里）。留空则不降级。
OWNER_UID = os.environ.get("BRIDGE_OWNER", "")

# 触发词：消息里出现任意一个才进入解析流程
INTENT_WORDS = ["加单", "补单", "加量", "追加", "增单", "补量"]

# ============================================================
# 二、Excel 模板（文档侧）
# ============================================================

# 模板文件（放在 templates/ 下，不要提交到 git）
TEMPLATE_FILE = "templates/template.xlsx"

# 模板中第 1 行是表头，数据从第 2 行开始。
HEADER_ROW = 1
DATA_START_ROW = 2

# 模板自带的"示例行"识别标记：命中任一条件即视为示例行并删除。
# 不同系统的模板不一样，按自己的模板改。
SAMPLE_ROW_MARKERS = {
    "B": "示例经营体",      # B 列等于该值 → 示例行
    "D": "6900000000000",   # D 列等于该值 → 示例行
}

# 列写入规则 —— 这是整个项目最需要按需修改的地方。
#   date        写入预算开始日期（本周周一，YYYY-MM-DD）
#   fixed       写入固定值
#   barcode     写入商品条码（强制文本，防 Excel 转科学计数法）
#   product_name 写入商品名称（由 lookup 查询得到）
#   quantity    写入数量（同日期+条码的行会累加）
COLUMN_RULES = {
    "A": {"type": "date"},
    "B": {"type": "fixed", "value": "示例经营体"},
    "C": {"type": "product_name"},
    "D": {"type": "barcode"},
    "E": {"type": "fixed", "value": "临时计划"},
    "H": {"type": "quantity"},
}

# 数量列（用于累加判断）
QTY_COLUMN = "H"
# 条码列（用于匹配已有行）
BARCODE_COLUMN = "D"
# 日期列
DATE_COLUMN = "A"

# 生成文档的命名模板。可用变量：{name} 品名 {qty} 合计数量 {date} 今天
OUTPUT_NAME_TEMPLATE = "加单+{name}-{qty}-{date}.xlsx"

# ============================================================
# 三、商品名称查询
# ============================================================

# 在线查询接口（GET，返回 JSON）。留空则只用本地缓存。
# 示例：https://erp.example.com/api/products?barcode={barcode}
PRODUCT_API = os.environ.get("BRIDGE_PRODUCT_API", "")

# 本地缓存文件（条码 → 名称）。可用 --refresh 从 csv/xlsx 重建。
PRODUCT_CACHE = "data/product_cache.json"

# 本地兜底数据源：csv 需含"产品条码""产品名称"两列
PRODUCT_LOCAL_CSV = "data/products.csv"

# ============================================================
# 四、ERP 推送（可选，对接自己的系统）
# ============================================================

# 是否在生成文档后自动导入 + 推送审批。
# 强烈建议先设为 False，手动跑通一次再打开。
ERP_AUTO_PUSH = False

ERP_BASE_URL = os.environ.get("BRIDGE_ERP_BASE", "https://erp.example.com")

# 认证方式二选一：
#   "cookie"  从 COOKIE_FILE 读取整条 Cookie 字符串
#   "header"  直接使用 AUTH_HEADER
ERP_AUTH_MODE = "cookie"
COOKIE_FILE = "secrets/erp_cookie.txt"
AUTH_HEADER = os.environ.get("BRIDGE_ERP_AUTH", "")

# 关键：有些系统要求 Referer 必须来自自家域，否则 403
ERP_REFERER = os.environ.get("BRIDGE_ERP_REFERER", "")

# 接口路径 —— 改成你自己系统的
ERP_IMPORT_PATH = "/api/items/import"     # body = JSON 行数组
ERP_PUSH_PATH = "/api/items/push"         # body = {"idList": [...], "operationType": 0}
ERP_LIST_PATH = "/api/items/page"         # 查询草稿，用于拿 id

# 模板表头 → ERP 字段名。导入接口要求 JSON 行数组时，靠这张表做映射。
FIELD_MAP = {
    "预算开始日期": "startDate",
    "经营体": "businessUnit",
    "产品名称": "productName",
    "产品条码": "productBarCode",
    "预算类型": "planType",
    "数量": "quantity",
    "备注": "memo",
}

# ============================================================
# 五、运行参数
# ============================================================

POLL_INTERVAL = 60          # 轮询间隔（秒）
STATE_FILE = "data/state.json"
LOG_FILE = "data/bridge.log"
OUTPUT_DIR = "data/output"  # 生成文档的暂存目录
