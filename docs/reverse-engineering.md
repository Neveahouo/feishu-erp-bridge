# 当内部系统没有 API 文档：从前端代码还原接口

> 这是整个项目里最通用的一招。适用于任何 Vue/React SPA 构建的内部系统——只要你能在浏览器里登录，就能还原出它的接口契约。

## 问题

你要自动化一个操作（比如"导入 Excel 并提交审批"），于是按标准流程抓包：

1. F12 → Network → 勾选 Preserve log
2. 在页面上完整操作一遍
3. 导出 HAR

结果打开 HAR 一看——**只有页面加载时的几个列表查询，关键动作一个都没录到**。

可能原因：请求走了 WebSocket、被 service worker 拦截、文件是前端解析后另发请求、或者你操作太快没等请求发出。但更常见的原因是：**你根本没意识到"上传文件"这个动作背后并没有文件上传请求**。

## 思路：代码不会骗人

前端要把请求发出去，接口路径和字段名**必须**以字符串形式存在于 JS 里。抓到不请求，就去代码里找。

### 第 1 步：确认是 SPA，拿到入口 bundle

```bash
curl -s https://your-erp.example.com/ -H "Cookie: $COOKIE" | grep -oE 'src="[^"]+\.js"'
```

典型输出：

```html
<script src="/static/js/manifest.6717b2db.js"></script>
<script src="/static/js/app.edb3abd3.js"></script>
```

`manifest` 是 webpack 的 chunk 映射表，`app` 是主逻辑。

### 第 2 步：在主 bundle 里找目标模块的 chunk id

```python
import re
app = open("app.edb3abd3.js", encoding="utf-8", errors="replace").read()

# 路由通常是懒加载：() => import(/* webpackChunkName: "xxx" */ "./views/xxx")
for m in re.finditer(r"collectionOf(\w+)", app):
    print(app[max(0, m.start()-120):m.start()+200])
```

你会看到类似：

```js
path: "channel-sales-budget",
component: () => n.e(1165).then(() => n(4027))
```

`n.e(1165)` 就是 chunk id = **1165**。

### 第 3 步：从 manifest 查 chunk 的真实文件名

webpack 的 manifest 是个数字 key 的对象：

```js
{1165:"b50db6d4e1c9ab504d86",4027:"a1b2c3d4..."}
```

```python
import re
man = open("manifest.6717b2db.js", encoding="utf-8").read()
m = re.search(r'(?:^|[{,])1165:"([a-f0-9]{16,})"', man)
print(m.group(1))   # b50db6d4e1c9ab504d86
```

拼出 URL：`/static/js/1165.b50db6d4e1c9ab504d86.js`

### 第 4 步：在目标 chunk 里搜接口路径

```bash
curl -s "https://your-erp.example.com/static/js/1165.b50db6d4e1c9ab504d86.js" \
     -H "Cookie: $COOKIE" -o chunk.js

grep -oE '"/api/[a-zA-Z0-9_/-]+"' chunk.js | sort -u
```

通常能直接拿到一组：

```
/api/items/page
/api/items/import
/api/items/push
/api/items/export
```

如果只看到 `xxx_page` 这种符号名（说明路径在公共 API 模块里），去 chunk 0（通常是最大的那个公共模块）里搜这个符号名：

```bash
grep -o '.\{200\}items_import.\{400\}' 0.js
```

### 第 5 步：还原请求体结构

这是最关键的一步——**"上传文件"往往不是文件上传**。

在导入弹窗的 chunk 里找处理逻辑：

```js
// 伪代码，真实代码通常是压缩过的
handleUpload(file) {
  this.parseExcel(file).then(rows => {
    return api.items_import(rows);   // 直接把解析后的行数组 POST 出去
  });
}
```

也就是说：**浏览器在本地把 Excel 解析成 JSON，然后 POST 一个数组**。后端根本不接收文件。

同时还能拿到列名映射：

```js
{
  "预算开始日期": "startDate",
  "产品条码": "productBarCode",
  "第2周": "twoWeekNum"
}
```

### 第 6 步：用 Cookie 复现

```python
import json, urllib.request

req = urllib.request.Request(
    "https://your-erp.example.com/api/items/import",
    data=json.dumps(rows).encode(),
    method="POST",
    headers={
        "Content-Type": "application/json",
        "Cookie": open("cookie.txt").read().strip(),
        "Referer": "https://your-erp.example.com/",   # 很多系统校验这个
    },
)
print(urllib.request.urlopen(req).read().decode())
```

如果返回 `success: true`，链路就通了。

## 几个提高成功率的技巧

| 技巧 | 说明 |
|---|---|
| **搜中文关键词** | 压缩后的代码里，中文 label 往往保留原样（`"推送审批"`），比搜英文变量名更容易定位 |
| **看 operationType** | 同一个 `operation` 接口常用 `operationType` 区分动作：0=提交，4=重新提交。把这个枚举找全，比找十个接口省事 |
| **对比"导出"功能** | 导出接口的字段名通常和导入接口对称，可以互相印证 |
| **注意 `n.e(id)` 的多种写法** | 可能是 `n.e(1165)`、`__webpack_require__.e(1165)`、`import(/* webpackChunkName */)`，正则要放宽 |
| **chunk 0 是宝藏** | 公共 API 封装、请求拦截器、鉴权逻辑通常都在最大的那个 chunk 里 |

## 合规提醒

这套方法只应该用于**你有权限访问的系统**和**你被授权执行的业务流程**。做之前确认：

- 公司安全规定是否允许脚本调用内部接口；
- 是否绕过审批或权限控制（如果绕过了，就是问题）；
- 凭证不要提交到任何代码仓库（本项目 `.gitignore` 已默认忽略 `secrets/`）。

本文只讲方法，不含任何真实系统的域名、路径或凭证。
