# 接口 JSON → 填充规则汇总（当前试卷）

基于你提供的 `get?topicType=...&topicID=...` 响应样本写死，先兼容这套试卷。

---

## 一、题目正文与媒体

| 接口字段 | 转换后 | 说明 |
|----------|--------|------|
| `topicContent` | `question` | 题干，去 HTML；可空（如听后转述题） |
| `topicOption` | 见下「两种题型」 | 选择题 → options + answer；听后填空/转述 → blanks |
| `analysis` | `explanation` | 解析 |
| `audioOriginalText` | `listening_script` | 听力原文 |
| `topicAttachment` | `audio_url` / `image_url` | **多附件**：`attachmentType` 1→音频，3→图片；按顺序取第一个音频、第一个图片；路径中 `\/` 会规范为 `/` |

---

## 二、topicOption 两种形态（按第一项自动区分）

### 1. 选择题（如 topicType 19 短对话理解）

- 第一项存在 `optionDesc` 或 `isTrue` 或 `option` 为 A/B/C/D。
- 解析：`options` = 每项 `optionDesc` 去 HTML 的数组；`answer` = 接口的 `answer`，若无则取 `isTrue === true` 的 `option`。

### 2. 听后记录/转述（如 topicType 117 听后记录并转述信息）

- 第一项无 `optionDesc`/`isTrue`，而有 `answer` / `topicStem` / `blanks` 等。
- 解析：`type: "listening_fill"`，`blanks` = `[{ question: item.topicStem, answer: item.answer }, ...]`；若 `topicContent` 为空则用各空 `topicStem` 拼成简短 `question`。

---

## 三、题目属性

| 接口字段 | 转换后 | 说明 |
|----------|--------|------|
| `courseTxt` | `course` | 如「牛津译林」 |
| `difficulty` | `difficulty` | "1"→简单，"2"→中等，"3"→困难 |
| `knowledgeTxt` | `knowledge_point` | 知识点长文本 |
| `permissionID` | `question_permission` | "1"→公开，"2"→仅自己可见 |
| `volume.name` | `grade` | 如「七年级上」 |
| `teachingIdTxt[0].teachingTxt` | `unit` | 如「七年级上册/Unit 1 This is me./Comic strip」 |

---

## 四、使用方式

- 控制台/Network 复制 **get?topicID=...** 的**完整响应**（`{ status, info, data: { topic } }`）。
- 插件「高级：JSON 粘贴」粘贴 → 点「仅填充当前页面」。
- 自动识别：选择题填选项+答案；听后填空/转述填多空 `blanks`；多附件时同时填音频+图片（若页面有对应上传框）。

---

## 五、后续扩展

- 新题型或新字段：再贴一份该题的**完整接口 JSON**，按样本补映射即可。
