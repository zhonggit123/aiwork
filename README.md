# Word 题库自动录入

从 Word 读取题目 → 大模型解析为结构化 JSON → 自动提交到录题页面或后端接口。

## 流程概览

1. **读 Word**：`python-docx` 提取全文，按块切分。
2. **AI 解析**：将每块文本发给大模型（OpenAI 兼容 API），强制输出题目/选项/答案/解析的 JSON。
3. **提交**：二选一  
   - **接口提交（推荐）**：F12 抓包得到「保存题目」的 POST 请求，用 `requests` 直接发 JSON。  
   - **浏览器填表**：用 Playwright 打开录题页，按配置的选择器填输入框并点击提交。

## 安装

```bash
cd "/Users/zhong/Downloads/AI录题"
pip install -r requirements.txt
# 若使用 Playwright 填表，还需安装浏览器
playwright install chromium
```

## 配置

1. 复制示例配置并填写：

```bash
cp config.example.yaml config.yaml
```

2. **必填**：`config.yaml` 中的 `llm`（API 地址、key、模型）。当前默认已配置为**火山引擎豆包**（`base_url: https://ark.cn-beijing.volces.com/api/v3`，模型 `doubao-seed-2-0-pro-260215`）。也可用环境变量 `ARK_API_KEY` 代替配置文件中的 `api_key`。若改用其他 OpenAI 兼容接口（通义、智谱、DeepSeek 等），只需改 `base_url` 和 `model`。

3. **提交方式**：

   - **用接口提交**：`submit.mode` 设为 `api`，然后抓包（见下）填 `submit.api` 的 `url`、`headers`、`body_mapping`。
   - **用浏览器填表**：`submit.mode` 设为 `playwright`，填 `submit.playwright.page_url` 和各输入框的 CSS 选择器（F12 检查元素得到）。

### 如何抓包（接口提交）

1. 浏览器打开录题系统，登录后进入「录入题目」页面。
2. 按 **F12** 打开开发者工具，切到 **Network（网络）**。
3. 在页面上**手动录入一道题并点击保存**。
4. 在 Network 里找到保存时发出的请求（一般是 POST），点开查看：
   - **Request URL** → 填到 `config.yaml` 的 `submit.api.url`
   - **Request Headers** 里的 `Authorization`、`Content-Type` 等 → 填到 `submit.api.headers`
   - **Request Payload**（JSON 体）→ 看字段名（如 `content`、`options`、`answer`），在 `body_mapping` 里把 AI 输出的字段名映射过去。

## 使用

### 方式一：命令行

```bash
# 使用 config.yaml 里的 Word 路径，并提交
python main.py

# 指定 Word 文件
python main.py --word 我的题库.docx

# 只解析、不提交，查看 AI 输出的 JSON
python main.py --dry-run
```

### 方式二：可视化 Web 页面（推荐）

本地起一个网页，上传 Word、表格编辑题目、一键提交或复制 JSON：

```bash
uvicorn app:app --reload --host 0.0.0.0 --port 8765
```

浏览器打开 **http://127.0.0.1:8765**，选择 Word 文件即可解析；可编辑表格后点击「一键提交到系统」或「复制 JSON」供浏览器插件使用。

### 方式三：浏览器插件（在录题页自动填表）

1. 打开 Chrome，地址栏输入 `chrome://extensions/`，开启「开发者模式」。
2. 点击「加载已解压的扩展程序」，选择本项目的 **`extension`** 文件夹。
3. 在扩展的「设置选择器」里，按你们录题页的输入框 id/name 填好选择器（可 F12 检查元素）。
4. 在 **Word 录入页面** 解析并「复制 JSON」后，打开**录题系统页面**，点击插件图标，粘贴 JSON，点击「填充当前页面表单」即可自动逐题填写并点提交。

## Word 格式建议

- 每题之间用**空行**分隔，便于按块切分。
- 题干、选项、答案、解析尽量有固定格式（例如「答案：A」「解析：……」），大模型更容易识别。
- 若含公式/图片，当前脚本只读文字；需要可考虑转 PDF 截图 + 多模态模型。

## 注意事项

- **接口提交**比浏览器填表更快、更稳，建议优先抓包用 API。
- API Key 不要提交到仓库，`config.yaml` 已建议加入 `.gitignore`。
- 若录题接口有频率限制，可在 `submit_api.submit_all` 里加 `time.sleep` 或限流。
