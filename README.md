# 飞书 → ERP 自动化桥

> 在飞书里发一句 `@机器人 加单 6900000000001 100`，30 秒后收到一份填好的 Excel，并自动导入业务系统、推进审批流。
>
> 一个把 **IM 当操作入口**、给没有 API 文档的内部系统做自动化的最小实现。

---

## 这个项目解决什么

很多公司的内部系统（自研 ERP、OA、进销存）都有类似问题：

- 有网页界面，但**没有对外 API 文档**；
- 一次操作要点七八下：打开系统 → 找到菜单 → 下载模板 → 手工填 Excel → 上传 → 筛选 → 勾选 → 提交；
- 每天重复十几次，出错还得重来。

这个项目的思路是：**把这条人工链路压缩成一句聊天指令。**

```
飞书群                     你的电脑 / 服务器                 内部系统
  │                              │                              │
  │  @机器人 加单 条码 数量       │                              │
  ├─────────────────────────────>│                              │
  │                              │ 1. 拉取消息（轮询）           │
  │                              │ 2. 解析条码+数量              │
  │                              │ 3. 查商品名称                 │
  │                              │ 4. 从模板 fork 新 Excel       │
  │                              │ 5. 回传 Excel 到群            │
  │<─────────────────────────────┤                              │
  │                              │ 6. 导入 + 推送审批 ──────────>│
  │                              │                              │
```

---

## 效果

```
👤 @MyBot 加单 6900000000001 100, 6900000000002 200

🤖 ✅ 已生成 加单+示例商品A等2种-300-2026-11-07.xlsx
   📅 日期：2026-11-03
   - 6900000000001 示例商品A × 100 → 新增（第 2 行）
   - 6900000000002 示例商品B × 200 → 新增（第 3 行）
   🏭 ERP：导入 + 推送成功（2 条）
   📎 [加单+示例商品A等2种-300-2026-11-07.xlsx]
```

---

## 四个关键技术点

| # | 问题 | 做法 |
|---|---|---|
| 1 | 内部系统没有 API 文档 | **扒前端 chunk 找接口**：抓包抓不到的动作，去 SPA 的 `app.*.js` / 懒加载 chunk 里搜接口路径字符串，直接还原请求契约。见 [docs/reverse-engineering.md](docs/reverse-engineering.md) |
| 2 | 飞书机器人读不到群消息 | **身份分层**：`bot` 能读私聊、能发消息，但读群历史会返回 `230027 access denied`；读群必须用 `user` 身份（用户本人是群成员）。见 [src/lark.py](src/lark.py) |
| 3 | 消息重复处理 | **`message_id` 幂等**：把处理过的 id 落到 `state.json`，轮询重启后自动跳过。断线不丢单，也不会重复发货 |
| 4 | 多人同时用会串数据 | **一次一单、模板 fork**：每条指令从干净模板复制一份新文档，不维护共享总账。产物可归档、可审计，失败只影响这一单 |

---

## 快速开始

### 1. 前置条件

- Python 3.9+
- 一个飞书自建应用（机器人），已拉进目标群
- [`lark-cli`](https://www.npmjs.com/) 并完成授权（能执行 `lark-cli im +chat-list --as bot`）
- Windows 用户需要 Git Bash（`lark-cli` 通过 `bash -lc` 调用）

### 2. 安装

```bash
git clone <this-repo>
cd feishu-erp-bridge
pip install -r requirements.txt
cp config.example.py config.py
```

### 3. 配置

打开 `config.py`，最少改这四项：

```python
BOT_NAMES = ["你的机器人名字"]        # 群里 @ 它才触发
TEMPLATE_FILE = "templates/template.xlsx"   # 你的 Excel 模板
PRODUCT_API = "https://erp.example.com/api/products?barcode={barcode}"
OUTPUT_NAME_TEMPLATE = "加单+{name}-{qty}-{date}.xlsx"
```

再按你的模板调整 **列写入规则**（`COLUMN_RULES`），这是唯一需要理解业务的地方：

```python
COLUMN_RULES = {
    "A": {"type": "date"},                          # 写入本周周一
    "B": {"type": "fixed", "value": "示例经营体"},   # 固定值
    "C": {"type": "product_name"},                  # 查出来的商品名
    "D": {"type": "barcode"},                       # 条码（强制文本）
    "E": {"type": "fixed", "value": "临时计划"},
    "H": {"type": "quantity"},                      # 数量（同日期+条码累加）
}
```

### 4. 运行

```bash
python -m src.watcher --parse "加单 6900000000001 100"   # 先测解析
python -m src.watcher --seed                            # 标记历史消息（首次部署必做）
python -m src.watcher --dry-run                         # 只检测不执行
python -m src.watcher --loop                            # 正式常驻
```

Windows 可直接双击 `scripts/start_loop.bat`。

### 5. 跑一下自测

```bash
python tests/selftest.py
```

不连飞书、不连 ERP，只验证解析、模板 fork、写入累加、字段映射。全部通过再上线。

---

## 接入你自己的系统

整个项目只有两个文件需要按实际情况改：

### `config.py` —— 列规则与接口路径

```python
ERP_BASE_URL   = "https://your-erp.example.com"
ERP_IMPORT_PATH = "/api/items/import"
ERP_PUSH_PATH   = "/api/items/push"

# 模板表头 → 接口字段
FIELD_MAP = {
    "预算开始日期": "startDate",
    "产品条码": "productBarCode",
    "数量": "quantity",
}
```

### `src/erp_client.py` —— 请求格式

默认假设导入接口收 **JSON 行数组**。如果你的系统要真正的文件上传：

```python
def import_rows(rows):
    # 换成 multipart/form-data 即可，其余逻辑可直接复用
    files = {"file": open(csv_path, "rb")}
    ...
```

> **默认不自动推送**：`ERP_AUTO_PUSH = False`。导入只产生"草稿"，你可以先在系统里核对，确认无误后再打开自动推送。`push()` 还额外要求 `confirm=True`，双重保险。

---

## 部署：怎么让它一直活着

轮询进程必须常驻。三种方式，按可靠性排序：

### ① Windows 服务（最稳，不需要登录）

用 [NSSM](https://nssm.cc/) 注册：

```bat
nssm install FeishuErpBridge "C:\Python311\python.exe" "-m src.watcher --loop"
nssm set FeishuErpBridge AppDirectory "C:\path\to\feishu-erp-bridge"
nssm start FeishuErpBridge
```

### ② 任务计划程序（推荐）

- 触发器：**开机时**
- 操作：启动 `scripts/start_loop.bat`
- 勾选"**如果任务失败，按以下频率重启**"

### ③ 启动文件夹（最简单，但要求用户已登录）

`Win+R` → `shell:startup` → 放入 `start_loop.bat` 的快捷方式。

**别忘了**：运行机器的电源计划设为"从不睡眠"。休眠会中断网络和子进程，这是最常见的"机器人突然不理人"的原因。

---

## 设计取舍

**为什么用轮询而不是飞书事件订阅？**
事件订阅需要公网回调地址、验签、证书，内网/个人场景成本过高。60 秒轮询换来零公网依赖，代价只是延迟——对加单这种场景完全可接受。

**为什么不用大模型解析消息？**
`加单 条码 数量` 是高度结构化的指令。正则更快、更便宜、零幻觉，而且出错时能精确告诉用户"条码 X 没识别到数量"，而不是让模型猜。

**为什么每次生成新文档，而不是维护一份总账？**
总账模式在多人协作下必然出现覆盖、脏数据、难以回溯。一次一单让每次操作都是**不可变的**，天然可审计。

---

## 已知局限

- **单实例**：同一台机器不要跑两个 watcher，会重复发文件。多机部署请用共享 state 或加分布式锁。
- **轮询延迟**：最长 `POLL_INTERVAL` 秒才响应。
- **Cookie 会过期**：对接系统用的是浏览器会话，失效后需要重新抓一次。建议把过期检测做进告警。
- **模板强绑定**：`COLUMN_RULES` 是写死的业务规则，换一套模板就要改配置。

---

## 目录结构

```
feishu-erp-bridge/
├── config.example.py      # 配置模板（复制为 config.py）
├── src/
│   ├── watcher.py         # 主循环：轮询 / 幂等 / 派发
│   ├── lark.py            # 飞书封装（身份分层、文件发送）
│   ├── parser.py          # 消息解析（意图 + 条码 + 数量）
│   ├── excel_engine.py    # 模板 fork、写入、回读校验
│   ├── product_lookup.py  # 商品名查询（在线 + 本地兜底）
│   └── erp_client.py      # ERP 适配器（导入 / 推送）
├── scripts/
│   ├── start_loop.bat
│   └── run_once.bat
├── docs/
│   ├── reverse-engineering.md   # 没有 API 文档时怎么找接口
│   └── lessons.md               # 踩坑记录
└── templates/             # 放你自己的 Excel 模板（不入库）
```

---

## License

MIT。代码可自由使用；请勿将任何公司内部接口、凭证、业务数据提交到公开仓库。
