const delay = (ms) => new Promise((r) => setTimeout(r, ms));

// 驰声等录题页接口返回的 data.topic 会通过注入脚本派发到 document，此处缓存后供「获取当前题 JSON」优先使用
let lastTopicFromApi = null;
document.addEventListener("topicFromApi", (e) => {
  if (e && e.detail && (e.detail.topicContent != null || e.detail.topicID != null)) lastTopicFromApi = e.detail;
});

// 注入到页面上下文，拦截 fetch/XHR 中 status===1 且含 data.topic 的响应，通过 document 事件传给 content script
function injectTopicApiCapture() {
  if (document.getElementById("__ai_luti_topic_capture__")) return;
  const script = document.createElement("script");
  script.id = "__ai_luti_topic_capture__";
  script.textContent = `
(function() {
  function emitTopic(data) {
    if (data && data.status === 1 && data.data && data.data.topic)
      document.dispatchEvent(new CustomEvent('topicFromApi', { detail: data.data.topic }));
  }
  var f = window.fetch;
  if (f) {
    window.fetch = function() {
      return f.apply(this, arguments).then(function(res) {
        var clone = res.clone();
        try {
          clone.json().then(emitTopic).catch(function(){});
        } catch (e) {}
        return res;
      });
    };
  }
  var XHROpen = XMLHttpRequest.prototype.open;
  var XHRSend = XMLHttpRequest.prototype.send;
  if (XHROpen && XHRSend) {
    XMLHttpRequest.prototype.open = function() { this._url = arguments[1]; return XHROpen.apply(this, arguments); };
    XMLHttpRequest.prototype.send = function() {
      var self = this;
      self.addEventListener('load', function() {
        try {
          if (typeof self.responseText === 'string' && self.responseText) {
            var data = JSON.parse(self.responseText);
            emitTopic(data);
          }
        } catch (e) {}
      });
      return XHRSend.apply(this, arguments);
    };
  }
})();
`;
  (document.head || document.documentElement).appendChild(script);
}
injectTopicApiCapture();

// ─── 消息入口 ──────────────────────────────────────────────────────────────
chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.type === "PING") {
    sendResponse({ ok: true });
    return false;
  }
  if (msg.type === "DETECT_FORM") {
    try { sendResponse({ ok: true, ...detectForm() }); }
    catch (e) { sendResponse({ ok: false, error: String(e?.message) }); }
    return true;
  }
  if (msg.type === "DETECT_AND_WALK") {
    (async () => {
      try {
        const result = await runDetectAndWalk(msg.selectors || {});
        sendResponse(result);
      } catch (e) {
        sendResponse({ ok: false, error: String(e?.message || e) });
      }
    })();
    return true;
  }
  if (msg.type === "FILL_FORM") {
    runFill(msg.questions, msg.selectors, msg.defaultAudioUrl, msg.defaultImageUrl, msg.debugSource, msg.ttsSettings)
      .then(sendResponse)
      .catch((e) => sendResponse({ ok: false, error: String(e?.message) }));
    return true;
  }
  if (msg.type === "FILL_FORM_DEBUG") {
    try {
      const report = runFillDebug(msg.questions || [], msg.selectors || {});
      sendResponse({ ok: true, report });
    } catch (e) {
      sendResponse({ ok: false, error: String(e?.message) });
    }
    return false;
  }
  if (msg.type === "GET_CURRENT_QUESTION_HTML") {
    (async () => {
      try {
        const root = await getMeaningfulQuestionRootAsync();
        const html = root ? root.outerHTML : document.body.outerHTML;
        sendResponse({ ok: true, html: (html || "").slice(0, 500000) });
      } catch (e) {
        sendResponse({ ok: false, error: String(e?.message) });
      }
    })();
    return true;
  }
  // 实时获取当前页「本题包含哪些区块」+ 总题数，供后端生成 prompt，不存 storage、不提交 HTML
  if (msg.type === "GET_PAGE_STRUCTURE") {
    try {
      const numbers = findTopicNumbers();
      const total = numbers.length || 1;
      const sectionLabels = getCurrentCardSectionLabels();
      const subCount = getCurrentCardSubCount();
      sendResponse({ ok: true, total, sectionLabels, subCount });
    } catch (e) {
      sendResponse({ ok: false, error: String(e?.message) });
    }
    return false;
  }
  // 从当前页读取表单值，导出为单题或全部题的 JSON（便于用 JSON 形式填充）
  if (msg.type === "GET_FORM_AS_JSON") {
    (async () => {
      try {
        const all = !!(msg.all);
        if (!all && lastTopicFromApi) {
          const fromApi = topicApiToQuestionJson(lastTopicFromApi);
          if (fromApi && (fromApi.question || fromApi.options?.length || fromApi.answer || fromApi.explanation)) {
            sendResponse({ ok: true, json: fromApi, fromApi: true });
            return;
          }
        }
        const result = all ? await getAllFormValuesAsJson() : await getCurrentFormValuesAsJson();
        sendResponse({ ok: true, json: result });
      } catch (e) {
        sendResponse({ ok: false, error: String(e?.message) });
      }
    })();
    return true;
  }
  // 遍历每一题，收集每题的小题数 + 区块标题，供后端在提示词中写明「第 i 题共 N 小题」
  if (msg.type === "GET_FULL_PAGE_STRUCTURE") {
    (async () => {
      try {
        const numbers = findTopicNumbers();
        const total = numbers.length;
        const slots = [];
        if (total === 0) {
          const sectionLabels = getCurrentCardSectionLabels();
          const subCount = getCurrentCardSubCount();
          sendResponse({ ok: true, total: 1, slots: [{ index: 1, sectionLabels, subCount }] });
          return;
        }
        for (let i = 0; i < total; i++) {
          await clickTopicNumber(numbers, i);
          const sectionLabels = getCurrentCardSectionLabels();
          const subCount = getCurrentCardSubCount();
          slots.push({ index: i + 1, sectionLabels, subCount });
        }
        await clickTopicNumber(numbers, 0);
        sendResponse({ ok: true, total, slots });
      } catch (e) {
        sendResponse({ ok: false, error: String(e?.message) });
      }
    })();
    return true;
  }
});

/** 页面是 SPA：题目表单由 JS 挂到 #right-container / #EditPaper，需轮询到有内容再取 */
function getMeaningfulQuestionRootAsync() {
  const hasContent = (el) => {
    if (!el) return false;
    const html = (el.innerHTML || "").trim();
    const text = (el.innerText || el.textContent || "").trim();
    return html.length > 150 || text.length > 80 ||
      !!el.querySelector("textarea, input:not([type=hidden]), [contenteditable], .caption, .question-part, .row");
  };
  const tryOnce = () => {
    const right = document.querySelector("#right-container");
    if (right && hasContent(right)) return right;
    const topic = document.querySelector(".topic-container");
    if (topic && hasContent(topic)) return topic;
    const main = document.querySelector(".main");
    if (main && hasContent(main)) return main;
    const edit = document.querySelector("#EditPaper");
    if (edit && hasContent(edit)) return edit;
    if (edit && edit.children && edit.children.length > 0) return edit;
    if (right && right.children && right.children.length > 0) return right;
    return null;
  };
  return new Promise((resolve) => {
    let attempt = 0;
    const maxAttempts = 10;
    const interval = 400;
    const run = () => {
      const root = tryOnce();
      if (root) {
        resolve(root);
        return;
      }
      attempt++;
      if (attempt >= maxAttempts) {
        resolve(document.querySelector("#right-container") || document.querySelector("#EditPaper") || document.body);
        return;
      }
      setTimeout(run, interval);
    };
    run();
  });
}

// ─── 左侧题目数字方块（驰声页 div.number，点击切题无验证报错）──────────
function findTopicNumbers() {
  // 驰声页：<div data-score="1" class="number"> 1 </div>
  const byClass = [...document.querySelectorAll("div.number[data-score], .list div.number, .topicList-box div.number")];
  if (byClass.length) return byClass;
  // 通用兜底：左侧栏里所有纯数字 div
  return [...document.querySelectorAll(".section-topicList div, .topicList-box div")].filter((el) =>
    /^\s*\d+\s*$/.test(el.textContent || "")
  );
}

/** 获取左侧题号列表中当前选中的题号索引（0-based），用于「仅填充当前页面」从选中题开始填 */
function getCurrentTopicIndex(topicNumbers) {
  if (!topicNumbers || topicNumbers.length === 0) return 0;
  for (let i = 0; i < topicNumbers.length; i++) {
    const el = topicNumbers[i];
    const c = (el.className || "").toLowerCase();
    if (/\b(active|on|current|cur|selected)\b/.test(c)) return i;
    if (el.getAttribute("aria-selected") === "true") return i;
    if (el.getAttribute("data-active") === "true" || el.getAttribute("data-current") === "true") return i;
  }
  return 0;
}

// ─── 点击题号切到第 i 题（0-indexed），等待页面更新 ─────────────────────
async function clickTopicNumber(numbers, index) {
  const el = numbers[index];
  if (!el) return false;
  el.click();
  await delay(900);
  return true;
}

// ─── 检测遍历（通过点左侧数字切题，无验证报错）────────────────────────
function mergeSelectors(a, b) {
  const out = { ...a };
  for (const k of Object.keys(b || {})) {
    if (!b[k]) continue;
    const aList = (out[k] || "").toString().split(",").map((x) => x.trim()).filter(Boolean);
    const bList = (b[k] || "").toString().split(",").map((x) => x.trim()).filter(Boolean);
    // 保留更长的列表，以便一题多空（如表格填空）时能填满所有空
    if (bList.length > aList.length) out[k] = bList.join(",");
    else if (aList.length === 0 && bList.length > 0) out[k] = bList.join(",");
  }
  return out;
}

function mergeFields(existing, newFields) {
  const map = new Map();
  for (const f of existing) map.set(f.role, f);
  for (const f of newFields) {
    if (!map.has(f.role)) map.set(f.role, f);
  }
  return Array.from(map.values());
}

// 采集当前题所在区域内的媒体与题型（图片/音频/题型代码等）
function getCurrentCardMediaAndType() {
  const root = document.querySelector(".topic-container, #topic-section, #right-container") || document.body;
  const within = (el) => root.contains(el);
  const hasImage = root.querySelector("input[type='file'], [data-type='image'], .upload-area, .image-upload") ||
    /图片|上传|image|图片上传/i.test(root.innerText || "");
  const hasAudio = root.querySelector("audio, [data-type='audio'], .audio-upload, source[type*='audio']") ||
    /音频|听力|audio|录音/i.test(root.innerText || "");
  let typeCode = "";
  const typeEl = root.querySelector("[data-topic-type], [data-type], [data-question-type], .topic-type");
  if (typeEl) {
    typeCode = typeEl.getAttribute("data-topic-type") || typeEl.getAttribute("data-type") || typeEl.getAttribute("data-question-type") || (typeEl.textContent || "").trim().slice(0, 20);
  }
  let inputKind = "text";
  if (hasImage && hasAudio) inputKind = "mixed";
  else if (hasImage) inputKind = "image";
  else if (hasAudio) inputKind = "audio";
  let questionJson = null;
  const jsonEl = root.querySelector("[data-question], [data-config], [data-question-config]");
  if (jsonEl) {
    const raw = jsonEl.getAttribute("data-question") || jsonEl.getAttribute("data-config") || jsonEl.getAttribute("data-question-config");
    if (raw) { try { questionJson = JSON.parse(raw); } catch (_) { questionJson = raw.slice(0, 200); } }
  }
  return { hasImage: !!hasImage, hasAudio: !!hasAudio, inputKind, typeCode: (typeCode || "").slice(0, 30), questionJson };
}

// 页面常见区块标题关键词（用于识别「上传音频」「设置题干」等），按出现顺序收集
const SECTION_LABEL_PATTERNS = [
  { pattern: /上传音频/i, label: "上传音频" },
  { pattern: /听力原文|原文\s*[：:]?/i, label: "听力原文" },
  { pattern: /设置题干|题干/i, label: "题干" },
  { pattern: /参考单词|送评单词|送评词|参考词|关键字/i, label: "参考单词" },
  { pattern: /图片选项|选项.*图片|option.*image/i, label: "图片选项" },
  { pattern: /设置选项|选项\s*[A-D]?/i, label: "设置选项" },
  { pattern: /设置答案|答案/i, label: "设置答案" },
  { pattern: /上传图片|图片\s*上传/i, label: "上传图片" },
  { pattern: /^解析\s*[：:]?/i, label: "解析" },
  { pattern: /题目属性\s*信息/i, label: "题目属性信息" },
];

/** 扫描当前题目区域内出现的区块标题（div.caption、.title、带 label 的 col 等），供大模型知道本题包含哪些组件 */
function getCurrentCardSectionLabels() {
  const root = document.querySelector(".topic-container, #topic-section, #topic, #right-container") || document.body;
  const seen = new Set();
  const labels = [];
  const walk = (el) => {
    if (!el || el.nodeType !== 1) return;
    const cls = (el.className || "").toString();
    const text = (el.textContent || "").trim().slice(0, 40);
    const isCaption = /caption|title|label|name/.test(cls) && !/button|btn/.test(cls);
    if (isCaption && text) {
      for (const { pattern, label } of SECTION_LABEL_PATTERNS) {
        if (pattern.test(text) && !seen.has(label)) {
          seen.add(label);
          labels.push(label);
          break;
        }
      }
    }
    if (el.children && el.children.length > 0) {
      for (let i = 0; i < el.children.length; i++) walk(el.children[i]);
    }
  };
  walk(root);

  // 额外：扫描输入框的 placeholder 文字，将有意义的 placeholder 也纳入标签
  // 这样当 caption 标题缺失时，AI 仍能知道该输入框的用途
  const inputEls = root.querySelectorAll(
    "input[placeholder]:not([type='hidden']):not([type='submit']):not([type='button']):not([type='file']):not([type='radio']):not([type='checkbox']), textarea[placeholder]"
  );
  for (const inp of inputEls) {
    const ph = (inp.getAttribute("placeholder") || "").trim().slice(0, 40);
    if (!ph || ph.length < 2) continue;
    for (const { pattern, label } of SECTION_LABEL_PATTERNS) {
      if (pattern.test(ph) && !seen.has(label)) {
        seen.add(label);
        labels.push(label);
        break;
      }
    }
  }

  return labels;
}

/**
 * 获取每个 .question-part 子区域的详情：小节标题文字 + 该区域的区块标签。
 * 返回 null（若无 .question-part），或 [{index, heading, sectionLabels}] 数组。
 */
function getSubQuestionDetails() {
  const root = document.querySelector(".topic-container, #topic-section, #topic, #right-container") || document.body;
  const parts = root.querySelectorAll(".question-part");
  if (!parts || parts.length === 0) return null;

  const details = [];
  for (let pi = 0; pi < parts.length; pi++) {
    const part = parts[pi];

    // 尝试从叶节点文字中找小节标题，如「第一节：信息记录」「第二节：信息转述」
    let heading = "";
    const allEls = part.querySelectorAll("*");
    for (const el of allEls) {
      if (el.children && el.children.length > 0) continue; // 只取叶节点
      const t = (el.textContent || "").trim();
      if (
        /第[一二三四五六七八九十]+节[：:]/.test(t) ||
        /信息记录|信息转述|模仿朗读|交际朗读|听后选择|听后应答/.test(t)
      ) {
        heading = t.slice(0, 60);
        break;
      }
    }

    // 收集本 .question-part 内的区块标签
    const partLabels = [];
    const seen = new Set();
    const walk = (el) => {
      if (!el || el.nodeType !== 1) return;
      const cls = (el.className || "").toString();
      const text = (el.textContent || "").trim().slice(0, 40);
      const isCaption = /caption|title|label|name/.test(cls) && !/button|btn/.test(cls);
      if (isCaption && text) {
        for (const { pattern, label } of SECTION_LABEL_PATTERNS) {
          if (pattern.test(text) && !seen.has(label)) {
            seen.add(label);
            partLabels.push(label);
            break;
          }
        }
      }
      for (let j = 0; j < (el.children ? el.children.length : 0); j++) walk(el.children[j]);
    };
    walk(part);

    // 采集本 part 内所有有意义的 placeholder（直接收集，不依赖 label 判断）
    const partAnswerPhs = [];
    const inputs = part.querySelectorAll("textarea, input[type='text'], input:not([type='hidden']):not([type='submit']):not([type='button']):not([type='file']):not([type='radio']):not([type='checkbox'])");
    for (const inp of inputs) {
      const ph = (inp.getAttribute("placeholder") || "").trim();
      // 只保留对 AI 有意义的提示：含"分隔"、"格式"、"输入"等关键词，且长度足够
      if (ph.length > 8 && /分隔|separator|格式|输入|填写|请输入/i.test(ph)) {
        if (!partAnswerPhs.includes(ph)) partAnswerPhs.push(ph);
      }
    }

    details.push({ index: pi + 1, heading, sectionLabels: partLabels, answerPlaceholders: partAnswerPhs });
  }
  return details;
}

/**
 * 检测当前题目卡片内的选项是否为图片类型。
 * 判断依据：
 * 1. 区块标题含「图片选项」
 * 2. 选项输入框 placeholder 含图片文件名特征（.png/.jpg）
 * 3. 选项输入框当前值是图片文件名
 * 4. 存在图片上传/预览组件
 * 返回 "image" | "text"
 */
function getCurrentCardOptionKind() {
  const root = document.querySelector(".topic-container, #topic-section, #topic, #right-container") || document.body;
  
  // 1. 区块标题含「图片选项」
  const allText = (root.innerText || root.textContent || "").toLowerCase();
  if (/图片选项|option.*image/i.test(allText.slice(0, 2000))) return "image";
  
  // 2. 选项区域内的输入框检测
  const optionRows = root.querySelectorAll(".row");
  for (const row of optionRows) {
    const cap = row.querySelector(".col.caption, .caption, .row-caption");
    const capText = (cap && cap.textContent || "").trim();
    if (!/设置选项|选项/.test(capText)) continue;
    
    const inputs = row.querySelectorAll("input[type='text'], input:not([type='hidden']):not([type='submit']):not([type='file']):not([type='radio']):not([type='checkbox']), textarea");
    for (const inp of inputs) {
      // 检查 placeholder
      const ph = (inp.getAttribute("placeholder") || "").toLowerCase();
      if (/\.png|\.jpg|\.jpeg|\.gif|\.webp|图片|支持.*图片|上传.*图片/.test(ph)) return "image";
      
      // 检查当前值是否是图片文件名
      const val = (inp.value || "").toLowerCase();
      if (/\.(png|jpg|jpeg|gif|webp)$/i.test(val)) return "image";
      
      // 输入框紧邻图片预览元素
      const parent = inp.parentElement;
      if (parent && (parent.querySelector("img, [class*='img'], [class*='image-preview'], [class*='preview']"))) return "image";
    }
    
    // 选项区域内直接有图片预览或 uploadify 图片组件
    if (row.querySelector(".uploadify[id*='option'], [id*='option'][class*='upload'], [id*='opt'][class*='image']")) return "image";
    
    // 检查是否有浏览按钮（通常图片选项会有浏览按钮）
    const browseBtn = row.querySelector("button, a.btn, [role='button']");
    if (browseBtn && /浏览|browse|选择图片|上传/i.test(browseBtn.textContent || "")) return "image";
  }
  
  // 3. option_a/b/c/d 对应的输入框检测（通用兜底）
  const optInputs = root.querySelectorAll(".option input[type=text], .col.option input, [class*='option'] input[type=text]");
  for (const inp of optInputs) {
    const ph = (inp.getAttribute("placeholder") || "").toLowerCase();
    const val = (inp.value || "").toLowerCase();
    if (/\.png|\.jpg|\.jpeg|\.gif|图片/.test(ph)) return "image";
    if (/\.(png|jpg|jpeg|gif|webp)$/i.test(val)) return "image";
  }
  
  // 4. 检查是否有多个浏览/删除按钮对（图片选项的典型特征）
  const browseBtns = root.querySelectorAll("button, a.btn");
  let browseCount = 0;
  for (const btn of browseBtns) {
    if (/浏览|browse/i.test(btn.textContent || "")) browseCount++;
  }
  // 如果有3个以上浏览按钮，很可能是图片选项（A/B/C/D 各一个）
  if (browseCount >= 3) return "image";
  
  return "text";
}

/** 当前题目卡片内的小题数量：统计 .question-part 或答案框数量，供提示词标明「本题共 N 小题」 */
function getCurrentCardSubCount() {
  const root = document.querySelector(".topic-container, #topic-section, #topic") || document.querySelector("#right-container");
  if (!root) return 1;
  // 优先：.question-part 区块数（最明确）
  const parts = root.querySelectorAll(".question-part");
  if (parts && parts.length > 0) return parts.length;
  // 次优：id 严格匹配 answer_N（纯数字后缀）的答案容器，如 div#answer_1、div#answer_2
  // 注意：不能匹配 answer_1_1、answer_1_2 这类嵌套选项元素
  const answerDivs = Array.from(root.querySelectorAll("div[id^='answer_']"))
    .filter(el => /^answer_\d+$/.test(el.id));
  if (answerDivs.length > 0) return answerDivs.length;
  // 兜底：radio 类型，按 name 分组计数（name="answer_1" 等）
  const answerRadios = root.querySelectorAll("input[type='radio'][name^='answer_']");
  if (answerRadios && answerRadios.length > 0) {
    const names = new Set(Array.from(answerRadios).map((el) => el.getAttribute("name")));
    return names.size;
  }
  return 1;
}
function getCurrentQuestionContext(index, numbers, currentSelectors, currentFields) {
  const slot = {
    index: index + 1,
    partName: "",
    typeHint: "",
    typeCode: "",
    subCount: 1,
    inputCount: {},
    optionCount: 0,
    optionKind: "text", // "text" | "image"：选项是文字还是图片类型
    media: { hasImage: false, hasAudio: false, inputKind: "text" },
    labels: {},
    questionJson: null,
    sectionLabels: [], // 当前题区域内出现的区块标题（上传音频、题干、设置选项等），供大模型对应拆分
  };

  // 从题号附近 DOM 读取 partName / typeHint
  const numEl = numbers[index];
  if (numEl) {
    // 优先：往上找到 paperEnter-main 容器，直接读其 .paperEnter-title 子元素
    let ancestor = numEl.parentElement;
    while (ancestor) {
      if (ancestor.classList && ancestor.classList.contains("paperEnter-main")) {
        const titleEl = ancestor.querySelector(".paperEnter-title");
        if (titleEl) {
          const titleText = (titleEl.textContent || "").trim().replace(/\s+/g, " ");
          slot.partName = slot.partName || titleText.slice(0, 30);
          slot.typeHint = slot.typeHint || titleText.slice(0, 30);
        }
        break;
      }
      ancestor = ancestor.parentElement;
    }
    // 兜底：往上 5 层检查 previousElementSibling 的文字（较短文本）
    if (!slot.typeHint) {
      let p = numEl.parentElement;
      for (let up = 0; up < 5 && p; up++, p = p.parentElement) {
        const prev = p.previousElementSibling;
        const text = (prev ? prev.textContent : p.textContent || "").trim();
        if (text && text.length < 80) {
          if (/part\s*[a-d]|第一部分|第二部分|听后选择|听后应答|模仿朗读|交际朗读|信息转述|表格填空|填空/i.test(text)) {
            slot.partName = slot.partName || text.replace(/\s+/g, " ").slice(0, 30);
          }
          if (/听后选择|听后应答|模仿朗读|交际朗读|信息转述|表格填空|单选|多选|判断|朗读|转述/i.test(text)) {
            slot.typeHint = slot.typeHint || text.replace(/\s+/g, " ").slice(0, 24);
          }
        }
      }
    }
  }

  // 各角色输入框数量；答案框数即小题/空数
  const roles = ["question", "keyword", "answer", "explanation", "option_a", "option_b", "option_c", "option_d"];
  roles.forEach((role) => {
    const s = currentSelectors[role];
    const n = !s ? 0 : (typeof s === "string" ? s.split(",").map((x) => x.trim()).filter(Boolean).length : 1);
    if (n > 0) slot.inputCount[role] = n;
    if (role === "answer" && n > 0) slot.subCount = Math.max(slot.subCount, n);
  });
  slot.optionCount = [currentSelectors.option_a, currentSelectors.option_b, currentSelectors.option_c, currentSelectors.option_d]
    .filter(Boolean).length;

  // 当前区域内的图片/音频/题型代码/题目 json
  try {
    const extra = getCurrentCardMediaAndType();
    slot.media = { hasImage: extra.hasImage, hasAudio: extra.hasAudio, inputKind: extra.inputKind };
    if (extra.typeCode) slot.typeCode = extra.typeCode;
    if (extra.questionJson != null) slot.questionJson = extra.questionJson;
  } catch (_) {}

  // 扫描当前题区域内区块标题（div.caption / .title 等），便于大模型按「上传音频、题干、选项、答案」拆分试题
  try {
    slot.sectionLabels = getCurrentCardSectionLabels();
  } catch (_) {}

  // 检测选项是否为图片类型（影响 AI prompt：图片选项用 <<IMG>> 占位）
  try {
    slot.optionKind = getCurrentCardOptionKind();
  } catch (_) {}

  (currentFields || []).forEach((f) => {
    if (f.label) slot.labels[f.role] = f.label;
  });
  return slot;
}

async function runDetectAndWalk(initialSelectors) {
  const numbers = findTopicNumbers();
  let selectors = { ...initialSelectors };
  let fields = [];
  let walked = 0;
  const total = numbers.length;
  const slots = []; // 每题的 subCount（小题数），用于填充时精确合并
  // 遍历过程中只要有一题有顶层音频输入框，就设为 true
  let hasTopLevelAudio = false;

  if (total === 0) {
    const result = detectForm();
    selectors = mergeSelectors(selectors, result.selectors);
    fields = mergeFields(fields, result.fields);
    if (result.hasTopLevelAudio) hasTopLevelAudio = true;
    walked = 1;
    const _sc = getCurrentCardSubCount();
    slots.push({
      index: 1,
      subCount: _sc,
      sectionLabels: getCurrentCardSectionLabels(),
      subQuestions: _sc > 1 ? (getSubQuestionDetails() || []) : [],
      optionKind: getCurrentCardOptionKind(),
    });
  } else {
    for (let i = 0; i < total; i++) {
      await clickTopicNumber(numbers, i);
      const result = detectForm();
      selectors = mergeSelectors(selectors, result.selectors);
      fields = mergeFields(fields, result.fields);
      if (result.hasTopLevelAudio) hasTopLevelAudio = true;
      walked = i + 1;
      // 收集本题小题数 + 区块标签，供 prompt 生成
      const subCount = getCurrentCardSubCount();
      const sectionLabels = getCurrentCardSectionLabels();
      // 若存在多个 .question-part，则额外收集每个小题的标题和标签（供逐小题描述）
      const subQuestions = subCount > 1 ? (getSubQuestionDetails() || []) : [];
      const optionKind = getCurrentCardOptionKind();

      // 读取题型提示：直接读 .paperEnter-title 文字，去掉大节编号和括号注释，保留核心题型名
      let typeHint = "";
      const _readTitleText = (el) => {
        if (!el) return "";
        let raw = (el.textContent || "").trim();
        // 去掉 "第X部分 " 前缀
        raw = raw.replace(/^第[一二三四五六七八九十百\d]+部分\s*/u, "");
        // 去掉 "一、" / "（一）" 等序号前缀
        raw = raw.replace(/^[一二三四五六七八九十百]+[、.．。]\s*/u, "");
        raw = raw.replace(/^[（(][一二三四五六七八九十]+[）)]\s*/u, "");
        // 去掉括号内的注释（如"共15小题，每小题1分，共15分"），避免嵌套括号破坏 prompt 格式
        raw = raw.replace(/[（(][^）)]*[）)]/gu, "").trim();
        return raw.slice(0, 20);
      };
      const numEl_i = numbers[i];
      if (numEl_i) {
        // 方法1：从题号元素向上找 paperEnter-main，读其 .paperEnter-title
        let anc = numEl_i.parentElement;
        while (anc) {
          if (anc.classList && anc.classList.contains("paperEnter-main")) {
            const titleEl = anc.querySelector(".paperEnter-title");
            if (titleEl) typeHint = _readTitleText(titleEl);
            break;
          }
          anc = anc.parentElement;
        }
      }
      // 方法2：直接查当前内容区的大节标题（题号在导航栏外时）
      if (!typeHint) {
        const titleEl = document.querySelector(".paperEnter-main .paperEnter-title, #topic-section .paperEnter-title, .paperEnter-title");
        if (titleEl) typeHint = _readTitleText(titleEl);
      }

      // 采集当前题答案框 / 题干框的 placeholder（用 result.selectors 而非全局累积 selectors）
      const answerPlaceholders = [];
      const qPlaceholders = [];
      const _curSelectors = result.selectors || {};
      for (const [role, phArr] of [["answer", answerPlaceholders], ["question", qPlaceholders]]) {
        const sel = _curSelectors[role] || "";
        if (!sel) continue;
        sel.split(",").map(s => s.trim()).filter(Boolean).forEach(s => {
          try {
            const el = document.querySelector(s);
            const ph = el && (el.getAttribute("placeholder") || "").trim();
            if (ph && ph.length > 3 && !phArr.includes(ph)) phArr.push(ph);
          } catch (_) {}
        });
      }

      // 保存本题当前字段（label 里通常含 placeholder 文字），供后端生成每题专属的字段提示
      const currentSlotFields = (result.fields || []).map(f => ({ role: f.role, label: f.label || "" }));
      slots.push({ index: i + 1, subCount, sectionLabels, subQuestions, optionKind, typeHint, answerPlaceholders, qPlaceholders, currentSlotFields });
      try {
        chrome.runtime.sendMessage({
          type: "DETECT_PROGRESS",
          walked,
          total,
          fields: Object.keys(selectors).filter((k) => selectors[k]),
          currentSectionLabels: sectionLabels,
          typeHint: typeHint,
        });
      } catch (_) {}
    }
    // 检测完回到第 1 题
    await clickTopicNumber(numbers, 0);
  }

  return {
    ok: true,
    selectors,
    fields,
    walked,
    total,
    slots,
    hasTopLevelAudio,  // 页面是否有大题共享音频输入框
    message: `已遍历 ${walked}/${total || walked} 题，了解页面结构，允许上传 Word。`,
  };
}

/** 仅当保存成功后才返回「继续录题」：
 *  - 页面中存在可见的「保存成功」提示（兄弟节点或任意位置），OR 按钮在可见弹窗/dialog 内
 *  - 这样既不误触保存失败时残留的按钮，又不会因父子结构不匹配被误拦
 */
function findNextBtn() {
  // 页面任意可见位置是否有保存成功文案
  const hasSavedOk = (() => {
    const dialogs = document.querySelectorAll(
      ".el-dialog, .el-message-box, .modal, [role='dialog'], .dialog-wrapper, .success-dialog"
    );
    for (const d of dialogs) {
      if (d.offsetParent === null) continue;
      if (/题目保存成功|保存成功/.test(d.innerText || d.textContent || "")) return true;
    }
    // 兜底：搜整个 body（保存成功通常只在弹窗里）
    return /题目保存成功|保存成功/.test(document.body.innerText || "");
  })();

  const goOn = document.querySelector(".go-on, [data-chivox-event*='topicContinue']");
  if (goOn && goOn.offsetParent !== null) return goOn; // .go-on 出现即表示成功

  if (!hasSavedOk) return null; // 没有成功文案，不点

  const candidates = document.querySelectorAll("button, a.btn, [role='button'], .el-button");
  for (const el of candidates) {
    if (el.offsetParent === null) continue;
    if (!/继续录题/.test((el.textContent || "").trim())) continue;
    return el;
  }
  return null;
}

/** 找当前题目区域的「保存/提交」按钮（驰声平台文案是「预览」，class 是 save-topic） */
function findSubmitBtn() {
  // 优先按 class/data 属性精确匹配，避免找到页面其他「预览」按钮
  const precise = document.querySelector(
    ".save-topic, [data-chivox-event*='saveTopic'], .btn-primary.save-topic"
  );
  if (precise && precise.offsetParent !== null) return precise;
  // 兜底：在题目输入区域内找文案含「保存/预览/提交」的按钮
  const scope = document.querySelector("#topic, .topic-input-panel, .paperEnter-main");
  if (scope) {
    for (const btn of scope.querySelectorAll("button, input[type=submit]")) {
      if (/保存|提交|预览|save|submit/i.test((btn.textContent || btn.value || "").trim())) return btn;
    }
  }
  return null;
}

// ─── 表单字段识别 ──────────────────────────────────────────────────────────
const ROLE_KEYWORDS = {
  // 文件上传类（先匹配，避免被其他关键词抢）
  audio_file:  ["上传音频", "音频上传", "录音", "上传听力"],
  image_file:  ["上传图片", "图片上传", "图片"],
  question:    ["设置题干", "题干", "题目", "content", "question"],
  keyword:     ["参考单词", "送评单词", "送评词", "参考词", "topickeyword", "keyword", "关键字"],
  answer:      ["设置答案", "答案", "answer", "正确"],
  listening_script: ["听力原文", "原文"],  // 仅听力材料正文，不要与「解析」混用
  explanation: ["解析", "explanation"],
  option_a:    ["选项a", "选项 a", "optiona", "option_a", "选项1"],
  option_b:    ["选项b", "选项 b", "optionb", "option_b", "选项2"],
  option_c:    ["选项c", "选项 c", "optionc", "option_c", "选项3"],
  option_d:    ["选项d", "选项 d", "optiond", "option_d", "选项4"],
  submit_btn:  ["保存", "预览", "提交", "确定", "save", "submit"],
  next_btn:    ["下一题", "继续录题", "next"],
  // 音频 URL 文本框（非 file 的输入框，用于填音频地址，与 audio_file 区分）
  audio_url:   ["上传音频", "音频上传", "音频链接", "音频地址"],
  // 题目属性（适用年级、课程、难度等）
  grade:       ["适用年级", "年级"],
  course:      ["课程"],
  unit:        ["课程单元", "单元"],
  knowledge_point: ["知识点"],
  difficulty:  ["难度"],
  question_permission: ["题目权限", "权限"],
  recorder:    ["录入人"],
};

/** 题目属性默认值（与录题页题目属性表单一致，填充时若 JSON 未提供则用此默认） */
const DEFAULT_QUESTION_PROPS = {
  grade: "初中 七年级上",
  course: "牛津译林",
  unit: "七年级上册/Unit 1 This is me./Comic strip,",
  knowledge_point: "话题项目/个人情况/个人信息,话题项目/个人",
  difficulty: "中等",
  question_permission: "仅自己可见",
  recorder: "金鸡湖演示老师",
};

/** 无音频/图片数据时使用的占位符（页面有上传框但 JSON 未提供时自动填充） */
const PLACEHOLDER_AUDIO_URL = "https://example.com/audio.mp3";
const PLACEHOLDER_IMAGE_BASE64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function matchRole(text) {
  if (!text) return null;
  const t = text.toLowerCase().replace(/\s+/g, " ").trim();
  for (const [role, kws] of Object.entries(ROLE_KEYWORDS)) {
    if (kws.some((k) => t.includes(k))) return role;
  }
  // 驰声等页面选项可能只标 "A"、"B"、"C"、"D"（无「选项」二字），单独匹配
  if (/^[a-d]\.?$/i.test(t)) {
    const map = { a: "option_a", b: "option_b", c: "option_c", d: "option_d" };
    return map[t.charAt(0).toLowerCase()] || null;
  }
  return null;
}

function getFieldLabel(el) {
  // 先找 <label for="id">
  if (el.id) {
    const lb = document.querySelector(`label[for="${el.id}"]`);
    if (lb) return lb.textContent.trim();
  }
  // 驰声页：<div class="row"><div class="col caption">上传音频：</div><div class="col question">...</div></div>
  let p = el.parentElement;
  for (let up = 0; up < 5 && p; up++, p = p.parentElement) {
    if (p.classList && p.classList.contains("row")) {
      const cap = p.querySelector(".col.caption, .caption");
      if (cap) return (cap.textContent || "").trim().replace(/\s+/g, " ").slice(0, 40);
      break;
    }
  }
  p = el.parentElement;
  for (let i = 0; i < 4 && p; i++, p = p.parentElement) {
    if (p.tagName === "LABEL") return p.textContent.trim();
    const prev = p.previousElementSibling;
    if (prev && /label|title|name|caption/i.test(prev.className || "")) {
      return (prev.textContent || "").trim().replace(/\s+/g, " ").slice(0, 40);
    }
  }
  // placeholder 作为首选兜底：能直接说明输入框用途（如"请输入图片文件名"）
  const placeholder = (el.getAttribute("placeholder") || "").trim();
  let out = (
    placeholder ||
    el.getAttribute("aria-label") ||
    el.getAttribute("name") ||
    el.id || ""
  ).trim();
  // 驰声「设置选项」常为 <span>A</span><input>，题干/选项在紧邻的前一兄弟节点
  if (!out && el.previousElementSibling) {
    const prevText = (el.previousElementSibling.textContent || "").trim();
    if (prevText.length <= 4 && /^[A-Da-d][.、．]?\s*$/.test(prevText)) {
      out = prevText.charAt(0).toUpperCase();
    }
  }
  return out;
}

function getSelector(el) {
  const getIdSelector = (raw) => {
    if (!raw) return null;
    const bogus = /^(undefined|null|NaN|false|true|0)$/.test(raw);
    if (bogus) return null;
    if (/^[a-zA-Z][\w-]*$/.test(raw)) return "#" + raw;
    return `[id="${String(raw).replace(/"/g, '\\"')}"]`;
  };
  // 驰声等页面常用 UUID 作 id（如 32091746-1532-1ed8-1eaa-891a75566192），以数字开头时 #id 在 CSS 中无效，改用 [id="..."]
  if (el.id) {
    const sel = getIdSelector(String(el.id));
    if (sel) return sel;
  }
  const tag = el.tagName.toLowerCase();
  const name = el.getAttribute("name");
  if (name && /^[a-zA-Z][\w.-]*$/.test(name)) return `${tag}[name="${name}"]`;
  // class 兜底
  if (el.className && typeof el.className === "string") {
    const cls = el.className.trim().split(/\s+/).filter((c) => /^[a-zA-Z][\w-]*$/.test(c));
    if (cls.length) return `${tag}.${cls.join(".")}`;
  }
  // 结构路径兜底：当 textarea/input 没有 id/name/class 时，生成相对稳定的 nth-of-type 选择器。
  const seg = (node) => {
    const t = node.tagName.toLowerCase();
    let idx = 1;
    let sib = node.previousElementSibling;
    while (sib) {
      if ((sib.tagName || "").toLowerCase() === t) idx++;
      sib = sib.previousElementSibling;
    }
    return `${t}:nth-of-type(${idx})`;
  };
  const parts = [seg(el)];
  let p = el.parentElement;
  for (let depth = 0; depth < 6 && p; depth++, p = p.parentElement) {
    if (p.hasAttribute && p.hasAttribute("data-fill-part-idx")) {
      const idx = p.getAttribute("data-fill-part-idx");
      return `[data-fill-part-idx="${String(idx).replace(/"/g, '\\"')}"] > ${parts.reverse().join(" > ")}`;
    }
    if (p.id) {
      const idSel = getIdSelector(String(p.id));
      if (idSel) return `${idSel} > ${parts.reverse().join(" > ")}`;
    }
    parts.push(seg(p));
  }
  return null;
}

function detectForm() {
  const selectors = {};
  const fields = [];
  const seen = new Set();
  // 只扫描当前题目区域，避免把左侧「预览保存」等误当提交按钮
  const scope = document.querySelector("#topic-section, .topic-container, #topic") || document.querySelector("#right-container") || document;

  const candidates = scope.querySelectorAll([
    'input:not([type="hidden"]):not([type="submit"]):not([type="button"])',
    "textarea",
    "select",
    "[contenteditable]",
    "script[type='text/plain'][id]",  // 驰声 UEditor 选项/题干（富文本）
    "div.edui-default[id]",
    "button",
    "a.btn, a[role='button']",
    "[role='button']",
  ].join(", "));

  for (const el of candidates) {
    const sel = getSelector(el);
    if (!sel || seen.has(sel)) continue;
    seen.add(sel);

    const tag = el.tagName.toLowerCase();
    const label = getFieldLabel(el);
    const hint = label || el.getAttribute("placeholder") || el.getAttribute("name") || el.id || "";
    let role = matchRole(hint);
    const isRadio = tag === "input" && (el.type || "").toLowerCase() === "radio";
    const name = (el.getAttribute("name") || "").trim();
    // 答案单选框（name="answer_1" 等）的 label 是 A/B/C/D，会误匹配成 option_*；按 name 前缀判为 answer
    if (isRadio && /^answer/i.test(name)) role = "answer";
    const isBtn = tag === "button" || el.type === "submit" || el.getAttribute("role") === "button";
    const isFileInput = tag === "input" && (el.type || "").toLowerCase() === "file";
    let inferredRole = role || (isBtn ? "submit_btn" : null);
    const elId = (el.id || "").toString();
    if (inferredRole === "question" && tag === "div" && /edui\d+_(toolbarbox|iframeholder|bottombar|scalelayer|message_holder)/i.test(elId)) inferredRole = null;
    // 仅对 file 类型的 input 使用 audio_file/image_file；非 file 类型转为 url 角色
    if ((inferredRole === "audio_file" || inferredRole === "image_file") && !isFileInput) {
      if (inferredRole === "audio_file" || /上传音频|音频上传|音频链接|音频地址/i.test(hint)) {
        inferredRole = "audio_url";
      } else if (inferredRole === "image_file" || /上传图片|图片上传|图片链接|图片地址|图片/i.test(hint)) {
        inferredRole = "image_url";
      } else {
        inferredRole = null;
      }
    }
    if (inferredRole === "submit_btn" && /删除|delete/i.test((el.textContent || hint || "").trim())) {
      inferredRole = null;
    }
    if (inferredRole === "audio_url" && (tag === "button" || tag === "a")) inferredRole = null;
    if (inferredRole) {
      if (selectors[inferredRole]) {
        selectors[inferredRole] += "," + sel;
      } else {
        selectors[inferredRole] = sel;
        fields.push({ role: inferredRole, selector: sel, label: hint || sel });
      }
    }
  }
  // 兜底：若未识别到选项框，在「设置选项」/ .option 区域内按顺序把 input 当作 A/B/C/D
  const needOptions = ["option_a", "option_b", "option_c", "option_d"].some((r) => !selectors[r]);
  if (needOptions && scope) {
    const optionInputs = scope.querySelectorAll(".option input[type=text], .option input.txt, .col.option input");
    if (optionInputs.length > 0) {
      const roleByIndex = ["option_a", "option_b", "option_c", "option_d"];
      optionInputs.forEach((el, i) => {
        const r = roleByIndex[i];
        if (!r) return;
        const sel = getSelector(el);
        if (!sel || seen.has(sel)) return;
        seen.add(sel);
        if (selectors[r]) selectors[r] += "," + sel;
        else { selectors[r] = sel; fields.push({ role: r, selector: sel, label: roleByIndex[i] }); }
      });
    }
  }
  // 驰声页兜底：按 .row 的 .caption 文本找「设置题干」「设置选项」「解析」
  const rows = scope.querySelectorAll(".row");
  for (const row of rows) {
    const cap = row.querySelector(".col.caption, .caption, .row-caption");
    const capText = (cap && cap.textContent || "").trim().replace(/\s+/g, " ");
    if (!capText) continue;
    if (!selectors.question && /设置题干|题干|题目/.test(capText)) {
      const ed = row.querySelector("textarea, [contenteditable=true], script[type='text/plain'][id], div.edui-default[id]");
      if (ed) {
        const sel = getSelector(ed);
        if (sel && !seen.has(sel)) {
          seen.add(sel); selectors.question = sel;
          const ph = (ed.getAttribute("placeholder") || "").trim().slice(0, 80);
          fields.push({ role: "question", selector: sel, label: capText, ...(ph ? { placeholder: ph } : {}) });
        }
      }
    }
    if (!selectors.keyword && /参考单词|送评单词|送评词|参考词/.test(capText)) {
      const ed = row.querySelector("textarea, input[type='text'], [contenteditable=true]");
      if (ed) {
        const sel = getSelector(ed);
        if (sel && !seen.has(sel)) {
          seen.add(sel); selectors.keyword = sel;
          const ph = (ed.getAttribute("placeholder") || "").trim().slice(0, 80);
          fields.push({ role: "keyword", selector: sel, label: capText, ...(ph ? { placeholder: ph } : {}) });
        }
      }
    }
    if (needOptions && /设置选项|选项/.test(capText)) {
      const inputs = [...row.querySelectorAll("input[type='text'], input.txt, textarea")];
      const roleByIndex = ["option_a", "option_b", "option_c", "option_d"];
      inputs.slice(0, 4).forEach((el, i) => {
        const r = roleByIndex[i];
        if (!r || selectors[r]) return;
        const sel = getSelector(el);
        if (!sel || seen.has(sel)) return;
        seen.add(sel);
        selectors[r] = sel;
        fields.push({ role: r, selector: sel, label: r });
      });
    }
    if (!selectors.explanation && /解析/.test(capText)) {
      const ed = row.querySelector("textarea, [contenteditable=true], script[type='text/plain'][id], div.edui-default[id]");
      if (ed) {
        const sel = getSelector(ed);
        if (sel && !seen.has(sel)) { seen.add(sel); selectors.explanation = sel; fields.push({ role: "explanation", selector: sel, label: capText }); }
      }
    }
    // 自定义下拉组件通配：标准控件 + class/id 含常见关键词的 div/span
    const DROPDOWN_SEL = "input, select, [class*='select'], [class*='dropdown'], [class*='selector'], [class*='picker'], [class*='cascader'], [class*='comboBox'], [class*='combo-box'], [class*='comboTree'], [class*='combo-tree'], span[id], div[id]:not(.col):not(.row):not(.caption)";
    const findCtrl = (r) => {
      const col = r.querySelector(".col:not(.caption), .col:last-child");
      const pool = col || r;
      const found = pool.querySelector(DROPDOWN_SEL);
      if (found && getSelector(found)) return found;
      // 再找最近的可点击子节点（有 id 或 name 的第一个子节点）
      for (const el of pool.querySelectorAll("[id], [name]")) {
        if (el.tagName === "SPAN" && /error/i.test(el.className || "")) continue;
        const s = getSelector(el);
        if (s && !seen.has(s)) return el;
      }
      return null;
    };
    if (!selectors.grade && /适用年级|年级/.test(capText)) {
      const ctrl = findCtrl(row);
      if (ctrl) { const sel = getSelector(ctrl); if (sel && !seen.has(sel)) { seen.add(sel); selectors.grade = sel; fields.push({ role: "grade", selector: sel, label: capText }); } }
    }
    if (!selectors.course && /课程/.test(capText) && !/课程单元|单元/.test(capText)) {
      const ctrl = findCtrl(row);
      if (ctrl) { const sel = getSelector(ctrl); if (sel && !seen.has(sel)) { seen.add(sel); selectors.course = sel; fields.push({ role: "course", selector: sel, label: capText }); } }
    }
    if (!selectors.unit && /课程单元|^单元/.test(capText)) {
      const ctrl = findCtrl(row);
      if (ctrl) { const sel = getSelector(ctrl); if (sel && !seen.has(sel)) { seen.add(sel); selectors.unit = sel; fields.push({ role: "unit", selector: sel, label: capText }); } }
    }
    if (!selectors.knowledge_point && /知识点/.test(capText)) {
      const ctrl = findCtrl(row);
      if (ctrl) { const sel = getSelector(ctrl); if (sel && !seen.has(sel)) { seen.add(sel); selectors.knowledge_point = sel; fields.push({ role: "knowledge_point", selector: sel, label: capText }); } }
    }
  }
  // 解析：按 placeholder 兜底（驰声常见「请在此输入解析内容」）
  if (!selectors.explanation && scope) {
    const byPlaceholder = scope.querySelector("textarea[placeholder*='解析'], textarea[placeholder*='解析内容'], input[placeholder*='解析']");
    if (byPlaceholder) {
      const sel = getSelector(byPlaceholder);
      if (sel && !seen.has(sel)) { seen.add(sel); selectors.explanation = sel; fields.push({ role: "explanation", selector: sel, label: "解析" }); }
    }
  }
  if (!selectors.keyword && scope) {
    const byPlaceholder = scope.querySelector(
      "textarea[placeholder*='参考单词'], textarea[placeholder*='送评单词'], textarea[placeholder*='送评词'], input[placeholder*='参考单词'], input[placeholder*='送评单词'], input[placeholder*='送评词']"
    );
    if (byPlaceholder) {
      const sel = getSelector(byPlaceholder);
      if (sel && !seen.has(sel)) { seen.add(sel); selectors.keyword = sel; fields.push({ role: "keyword", selector: sel, label: "参考单词" }); }
    }
  }
  // 解析：驰声页 div#analysis contenteditable
  if (!selectors.explanation && scope) {
    const analysisEl = scope.querySelector("#analysis, [id='analysis']");
    if (analysisEl) {
      const sel = getSelector(analysisEl);
      if (sel && !seen.has(sel)) { seen.add(sel); selectors.explanation = sel; fields.push({ role: "explanation", selector: sel, label: "解析" }); }
    }
  }
  // 答案：按「设置答案」行找 div.col.answer / div.ui-radio#answer_*（点 label 选选项）
  if (!selectors.answer && scope) {
    // 多小题时每个子题都有自己的「设置答案」行，全部累加（不 break）
    for (const row of scope.querySelectorAll(".row")) {
      const cap = row.querySelector(".col.caption, .caption, .row-caption");
      const capText = (cap && cap.textContent || "").trim();
      if (!/设置答案|答案/.test(capText)) continue;
      const colAnswer = row.querySelector(".col.answer, .answer");
      const uiRadio = row.querySelector("div.ui-radio[id^='answer_'], div[id^='answer_']");
      const target = colAnswer || uiRadio;
      if (target) {
        const sel = getSelector(target);
        if (sel && !seen.has(sel)) {
          seen.add(sel);
          if (selectors.answer) {
            selectors.answer += "," + sel;
            const f = fields.find(f => f.role === "answer");
            if (f) f.selector = selectors.answer;
          } else {
            selectors.answer = sel;
            fields.push({ role: "answer", selector: sel, label: capText });
          }
        }
        continue; // 继续找下一个小题的答案行
      }
      const firstRadio = row.querySelector("input[type=radio][name^='answer_']");
      if (firstRadio) {
        const sel = getSelector(firstRadio);
        if (sel && !seen.has(sel)) {
          seen.add(sel);
          if (selectors.answer) {
            selectors.answer += "," + sel;
            const f = fields.find(f => f.role === "answer");
            if (f) f.selector = selectors.answer;
          } else {
            selectors.answer = sel;
            fields.push({ role: "answer", selector: sel, label: capText });
          }
        }
      }
    }
  }
  // 选项：按「设置选项」行内所有 input 兜底（不限于 type=text）
  if (needOptions && scope) {
    for (const row of scope.querySelectorAll(".row")) {
      const cap = row.querySelector(".col.caption, .caption, .row-caption");
      const capText = (cap && cap.textContent || "").trim();
      if (!/设置选项|选项/.test(capText)) continue;
      const inputs = [...row.querySelectorAll("input:not([type='hidden']):not([type='submit']):not([type='button']), textarea")];
      const roleByIndex = ["option_a", "option_b", "option_c", "option_d"];
      inputs.slice(0, 4).forEach((el, i) => {
        const r = roleByIndex[i];
        if (!r || selectors[r]) return;
        const sel = getSelector(el);
        if (!sel || seen.has(sel)) return;
        seen.add(sel);
        selectors[r] = sel;
        fields.push({ role: r, selector: sel, label: r });
      });
      break;
    }
  }
  // 题目属性：按常见 id/name 兜底（驰声等页可能用 #grade、#course、#unit 等）
  if (scope) {
    const byIdOrName = (role, ids) => {
      if (selectors[role]) return;
      for (const id of ids) {
        let el = scope.querySelector("[id=\"" + id.replace(/"/g, '\\"') + "\"]");
        if (!el) el = scope.querySelector("[name=\"" + id.replace(/"/g, '\\"') + "\"]");
        if (el) {
          const sel = getSelector(el);
          if (sel && !seen.has(sel)) { seen.add(sel); selectors[role] = sel; fields.push({ role, selector: sel, label: role }); break; }
        }
      }
    };
    byIdOrName("grade", ["grade", "gradeId", "grade_id"]);
    byIdOrName("unit", ["teaching", "unit", "unitId", "unit_id", "teachingId"]);
    byIdOrName("knowledge_point", ["knowledge", "knowledge_point", "knowledgePoint", "knowledgeIDs"]);
    byIdOrName("course", ["course", "courseId", "course_id"]);
    byIdOrName("difficulty", ["difficulty", "difficultyId"]);
    byIdOrName("question_permission", ["permission", "permissionID", "question_permission"]);
    byIdOrName("explanation", ["analysis", "explanation", "explanationContent"]);
  }

  // ── 专项兜底：驰声级联年级（jsCascaderWarp）+ 课程单元树（ui-comboBox）──
  if (!selectors.grade && scope) {
    const el = scope.querySelector(".jsCascaderWarp, [class*='topic-cascader'], div[class*='cascader-input']");
    if (el) {
      const sel = getSelector(el);
      if (sel && !seen.has(sel)) { seen.add(sel); selectors.grade = sel; fields.push({ role: "grade", selector: sel, label: "适用年级" }); }
    }
  }
  if (!selectors.unit && scope) {
    const el = scope.querySelector(".ui-comboTree[id], div[class*='comboTree'][id], .ui-comboBox[id], div[class*='comboBox'][id], div[class*='combo-box'][id]");
    if (el) {
      const sel = getSelector(el);
      if (sel && !seen.has(sel)) { seen.add(sel); selectors.unit = sel; fields.push({ role: "unit", selector: sel, label: "课程单元" }); }
    }
  }
  if (!selectors.knowledge_point && scope) {
    const el = scope.querySelector(".ui-comboTree[id='knowledge'], div[id='knowledge']");
    if (el) {
      const sel = getSelector(el);
      if (sel && !seen.has(sel)) { seen.add(sel); selectors.knowledge_point = sel; fields.push({ role: "knowledge_point", selector: sel, label: "知识点" }); }
    }
  }

  // ── 修正 answer：始终优先用 div.ui-radio 容器，而不是单个 radio input ──
  // 单个 radio 的 value 属性不可预测，而 label[for] 索引顺序是固定的 A/B/C/D
  // 多小题时 querySelectorAll 收集所有答案容器（#answer_1, #answer_2…），避免只取第一个
  if (scope) {
    const answerContainers = scope.querySelectorAll(
      "div.ui-radio[id^='answer_'], .col.answer div[id^='answer_'], div.answer-box"
    );
    if (answerContainers.length > 0) {
      const allSels = Array.from(answerContainers)
        .map(el => getSelector(el))
        .filter(s => s && !seen.has(s));
      allSels.forEach(s => seen.add(s));
      if (allSels.length > 0) {
        const finalSel = allSels.join(",");
        selectors.answer = finalSel;
        const idx = fields.findIndex((f) => f.role === "answer");
        if (idx >= 0) fields[idx].selector = finalSel;
        else fields.push({ role: "answer", selector: finalSel, label: "答案" });
      }
    }
  }

  // ── 优先识别「顶层共享音频 URL」与「顶层共享听力原文」──
  // 关键：必须排除 .question-part 内的小题字段，否则在「大题有共享音频 + 小题也有音频」时
  // 顶层 role 可能误指向第 1 小题，导致共享原文/音频偶发填不进去或串位。
  if (scope) {
    const topAudioInput = Array.from(scope.querySelectorAll(
      "input.audioFileName[id], input[id^='audioFileName'], input[id*='audio'][type='text'], input[class*='audioFile'], input[placeholder*='mp3'], input[placeholder*='音频'], input[placeholder*='MP3']"
    )).find((el) => !el.closest(".question-part"));
    if (topAudioInput) {
      const sel = getSelector(topAudioInput);
      if (sel) {
        selectors.audio_url = sel;
        const idx = fields.findIndex((f) => f.role === "audio_url");
        if (idx >= 0) fields[idx].selector = sel;
        else fields.push({ role: "audio_url", selector: sel, label: "共享音频URL" });
      }
    }
  }

  // ── 识别顶层共享听力原文 textarea（排除 .question-part 内的小题原文）──
  if (scope) {
    const topScriptCandidates = Array.from(scope.querySelectorAll(
      ".audioOriginalText textarea[id], .audioOriginalText textarea, textarea.audioOriginalText, textarea[id*='originalText'], textarea[name*='originalText'], textarea[id*='script'], textarea[name*='script'], textarea[placeholder*='听力'], textarea[placeholder*='原文'], textarea[placeholder*='报告'], textarea[placeholder*='script']"
    )).filter((el) => !el.closest(".question-part"));
    const topScriptEl = topScriptCandidates.find((ta) => {
      const ph = (ta.placeholder || ta.getAttribute("placeholder") || "").trim();
      const row = ta.closest(".row, div");
      const rowText = ((row && row.textContent) || "").slice(0, 100);
      const lbl = row && row.querySelector("label, .label, .caption, [class*='label'], [class*='caption']");
      const lblText = ((lbl && lbl.textContent) || "").trim();
      const clsId = `${ta.className || ""} ${ta.id || ""} ${ta.name || ""}`;
      return /audioOriginalText|originalText|script/i.test(clsId)
        || /听力|原文|报告|script/i.test(ph)
        || /听力原文|原文/.test(lblText)
        || /听力原文|原文/.test(rowText);
    });
    if (topScriptEl) {
      const sel = getSelector(topScriptEl);
      if (sel) {
        selectors.listening_script = sel;
        const idx = fields.findIndex((f) => f.role === "listening_script");
        if (idx >= 0) fields[idx].selector = sel;
        else fields.push({ role: "listening_script", selector: sel, label: "听力原文" });
      }
    }
  }

  // ── 识别题干图片输入框（#imageFileName 或含"图片"placeholder 的 input）──
  if (!selectors.image_url && scope) {
    const el = scope.querySelector(
      "#imageFileName, input.imageFileName, input[id*='imageFile'], input[id*='image'][type='text'], input[placeholder*='jpg'], input[placeholder*='图片']"
    );
    if (el) {
      const sel = getSelector(el);
      if (sel && !seen.has(sel)) { seen.add(sel); selectors.image_url = sel; fields.push({ role: "image_url", selector: sel, label: "题干图片URL" }); }
    }
  }

  // ── 识别图片上传组件（uploadify 内的隐藏 input[type=file]）──
  if (!selectors.image_file && scope) {
    // 查找 uploadify 组件内的隐藏文件上传框（驰声平台常用）
    const uploadifyContainer = scope.querySelector("#image_upload.uploadify, .uploadify[id*='image'], div[id*='image_upload']");
    if (uploadifyContainer) {
      // uploadify 内部的 input[type=file] 通常是隐藏的
      const hiddenFileInput = uploadifyContainer.querySelector("input[type='file']");
      if (hiddenFileInput) {
        const sel = getSelector(hiddenFileInput);
        if (sel && !seen.has(sel)) { seen.add(sel); selectors.image_file = sel; fields.push({ role: "image_file", selector: sel, label: "题干图片上传(uploadify)" }); }
      }
    }
  }

  // ── 识别音频上传组件（uploadify 内的隐藏 input[type=file]）──
  if (!selectors.audio_file && scope) {
    // 查找音频 uploadify 组件（驰声平台常用 #audio_upload1）
    const audioUploadifyContainer = scope.querySelector(
      "#audio_upload1.uploadify, #audio_upload.uploadify, .uploadify[id*='audio'], " +
      "div[id='audio_upload1'], div[id='audio_upload'], div[id*='audio_upload']"
    );
    if (audioUploadifyContainer) {
      const hiddenFileInput = audioUploadifyContainer.querySelector("input[type='file']");
      if (hiddenFileInput) {
        const sel = getSelector(hiddenFileInput);
        if (sel && !seen.has(sel)) { seen.add(sel); selectors.audio_file = sel; fields.push({ role: "audio_file", selector: sel, label: "音频上传(uploadify)" }); }
      }
    }
  }

  // ── 识别音频上传框（普通 input[type=file] 用于上传音频）──
  if (!selectors.audio_file && scope) {
    const fileInputs = scope.querySelectorAll("input[type='file']");
    for (const el of fileInputs) {
      const parent = el.closest(".row, .form-group, .upload-area, .audio-upload, [class*='audio'], [class*='upload']");
      const hint = (parent?.innerText || el.getAttribute("accept") || "").toLowerCase();
      const accept = (el.accept || "").toLowerCase();
      // 判断是音频上传：accept 包含 mp3/audio，或父容器包含"音频"/"录音"
      if (hint.includes("音频") || hint.includes("录音") || accept.includes("mp3") || accept.includes("audio")) {
        const sel = getSelector(el);
        if (sel && !seen.has(sel)) { seen.add(sel); selectors.audio_file = sel; fields.push({ role: "audio_file", selector: sel, label: "音频上传" }); break; }
      }
    }
  }

  // ── 识别图片上传框（普通 input[type=file] 用于上传图片）──
  if (!selectors.image_file && scope) {
    // 查找包含"图片"相关提示的文件上传框
    const fileInputs = scope.querySelectorAll("input[type='file']");
    for (const el of fileInputs) {
      const parent = el.closest(".row, .form-group, .upload-area, .image-upload, [class*='image'], [class*='upload']");
      const hint = (parent?.innerText || el.getAttribute("accept") || "").toLowerCase();
      const accept = (el.accept || "").toLowerCase();
      // 判断是图片上传（而非音频上传）：accept 包含 jpg/jpeg/png/image，或父容器包含"图片"
      if (hint.includes("图片") || accept.includes("jpg") || accept.includes("jpeg") || accept.includes("png") || accept.includes("image")) {
        const sel = getSelector(el);
        if (sel && !seen.has(sel)) { seen.add(sel); selectors.image_file = sel; fields.push({ role: "image_file", selector: sel, label: "题干图片上传" }); break; }
      }
    }
  }

  // ── 识别「第N小题」区块（.question-part）内的子题字段 ──
  // 对应 blank_answer_N / blank_question_N / blank_keyword_N / blank_audio_N / option_a/b/c/d
  if (scope) {
    const parts = scope.querySelectorAll(".question-part"); // 不要求有 id
    const usedEls = new WeakSet(); // 用 DOM 元素引用去重，避免 name 相同跨小题误判
    // 收集各 option role 的 part 专属选择器，循环结束后整体替换全局扫描的通用选择器
    const partOptSels = { option_a: [], option_b: [], option_c: [], option_d: [] };
    parts.forEach((part, partIdx) => {
      // 给每个 part 打上位置标记，供后续生成唯一选择器
      part.setAttribute("data-fill-part-idx", String(partIdx));
      const partSel = `[data-fill-part-idx="${partIdx}"]`;

      // 始终用 partIdx+1 作 role 编号，与 blanks 数组下标对齐（0-based→1-based）。
      // 不再使用 DOM id 尾数：DOM id 可能是 31/32 这类非连续大数，
      // 导致 getValueForRole('blank_script_31') 去取 blanks[30] 越界返回空。
      const n = partIdx + 1;

      // 参考答案 / answer
      // 注意：排除 .audioText / .audioOriginalText 内的 textarea，那是听力原文框
      const answerTa = part.querySelector(".answer textarea[id], .standardAnswer textarea[id]")
        || (() => {
          // 兜底：找 class="textarea" 且有 id 的 textarea，但排除听力原文区域
          const candidates = part.querySelectorAll("textarea.textarea[id]");
          for (const ta of candidates) {
            const parent = ta.closest(".audioText, .audioOriginalText, .col.audioText");
            if (!parent) return ta; // 不在听力原文区域内，可用
          }
          return null;
        })();
      if (answerTa && !usedEls.has(answerTa)) {
        usedEls.add(answerTa);
        const role = `blank_answer_${n}`;
        const sel = getSelector(answerTa);
        if (sel && !seen.has(sel) && !selectors[role]) {
          seen.add(sel); selectors[role] = sel;
          fields.push({ role, selector: sel, label: `第${n}小题参考答案` });
        }
      }

      // 参考音频（audioFileName2 等，以及"支持mp3格式上传"等 placeholder 的音频输入框）
      const audioInput = part.querySelector(
        // 原有精确匹配
        "input.audioFileName[id], input[id^='audioFileName'], " +
        // 扩展：id/class 含 audio（但不含 image/filename 避免误匹配）
        "input[id*='audio'][type='text'], input[class*='audioFile'], " +
        // 扩展：placeholder 含 mp3 或 音频 的 text 输入框
        "input[placeholder*='mp3'], input[placeholder*='音频'], input[placeholder*='MP3']"
      );
      if (audioInput && !usedEls.has(audioInput)) {
        usedEls.add(audioInput);
        const role = `blank_audio_${n}`;
        const basicSel = getSelector(audioInput);
        // 用 partSel 前缀拼接唯一选择器，与 blank_script_N 同理
        const sel = basicSel ? `${partSel} ${basicSel}` : null;
        if (sel && !selectors[role]) {
          selectors[role] = sel;
          fields.push({ role, selector: sel, label: `第${n}小题参考音频URL` });
        }
      }

      // 听力原文（每小题的音频文字转录区域）
      // ⚠️ 必须在 blank_question_N 前检测，优先占用该 textarea，避免被误当题干的兜底逻辑抢走
      // 判断依据：
      //   1. 行标题含「听力原文」/「原文」的 textarea
      //   2. textarea 的 class/id 含 audioOriginalText / originalText / script
      //   3. 父级 div 的 class 含 audioOriginalText
      let scriptTa = null;
      for (const row of Array.from(part.querySelectorAll(".row"))) {
        const cap = row.querySelector(".col.caption, .caption, .row-caption");
        const capText = (cap && cap.textContent || "").trim();
        if (/听力原文|原文/.test(capText)) {
          const ta = row.querySelector("textarea");
          if (ta && !usedEls.has(ta)) { scriptTa = ta; break; }
        }
      }
      // 兜底：按 class/id 特征查找听力原文 textarea
      if (!scriptTa) {
        const candidates = part.querySelectorAll(
          ".audioOriginalText textarea, textarea.audioOriginalText, textarea[id*='originalText'], textarea[id*='script'], textarea[name*='originalText'], textarea[name*='script']"
        );
        for (const ta of candidates) {
          if (!usedEls.has(ta)) { scriptTa = ta; break; }
        }
      }
      // 再兜底：检查行内文字（不依赖 .caption class），扩大 textContent 检查范围
      if (!scriptTa) {
        for (const row of Array.from(part.querySelectorAll(".row, div"))) {
          const rowText = (row.textContent || "").slice(0, 100);
          if (/听力原文|原文[：:]/.test(rowText)) {
            const ta = row.querySelector("textarea");
            if (ta && !usedEls.has(ta)) { scriptTa = ta; break; }
          }
        }
      }
      // 最终兜底：直接检查 <label> 文字 或 textarea 的 placeholder
      // 处理「听力原文：非必填，用于学生报告呈现」这类 placeholder 的情况
      if (!scriptTa) {
        for (const ta of Array.from(part.querySelectorAll("textarea"))) {
          if (usedEls.has(ta)) continue;
          const ph = (ta.placeholder || ta.getAttribute("placeholder") || "").trim();
          // placeholder 含"听力"/"原文"/"报告"的均视为听力原文输入框
          if (/听力|原文|报告|script/i.test(ph)) { scriptTa = ta; break; }
          // 检查相邻的 <label> 文字
          const row = ta.closest(".row, div");
          const lbl = row && row.querySelector("label, .label, [class*='label']");
          const lblText = (lbl && lbl.textContent || "").trim();
          if (/听力原文|原文/.test(lblText)) { scriptTa = ta; break; }
        }
      }
      if (scriptTa) {
        // 无论是否已被宽泛扫描收入 seen，都要先把元素加入 usedEls，
        // 防止下面 blank_question_N 的兜底逻辑把同一个 textarea 抢走
        usedEls.add(scriptTa);
        const role = `blank_script_${n}`;
        const basicSel = getSelector(scriptTa);
        // 用 partSel 前缀拼接唯一选择器，规避宽泛扫描已把 basicSel 收入 seen 的问题
        const sel = basicSel ? `${partSel} ${basicSel}` : null;
        if (sel && !selectors[role]) {
          selectors[role] = sel;
          fields.push({ role, selector: sel, label: `第${n}小题听力原文` });
        }
      }

      // 题干 / question（不要求 textarea 有 id，兼容无 id 的情况）
      // 注意：不用 !seen.has(sel) 过滤，因为 broader scan 可能已把第1小题题干记为 question，
      // 但我们仍需用同一元素作为 blank_question_1 来填 blanks[0].question。
      const stemTa = part.querySelector(
        ".topicStem textarea, textarea.topicStem, .stem textarea, .question-stem textarea"
      ) || (() => {
        // 兜底：找不在 .answer/.standardAnswer/.topicKeyword 内的第一个 textarea
        // 同时排除已被 blank_script_N（听力原文）或 blank_audio_N 等占用的元素
        const all = Array.from(part.querySelectorAll("textarea"));
        const answerTaEl = part.querySelector(".answer textarea, .standardAnswer textarea");
        const kwTaEl = part.querySelector(".topicKeyword textarea");
        return all.find(t => t !== answerTaEl && t !== kwTaEl && !usedEls.has(t));
      })();
      if (stemTa) {
        const role = `blank_question_${n}`;
        const sel = getSelector(stemTa);
        if (sel && !selectors[role]) {
          seen.add(sel); selectors[role] = sel;
          fields.push({ role, selector: sel, label: `第${n}小题题干` });
        }
      }

      // 关键字
      const kwTa = part.querySelector(".topicKeyword textarea[id]");
      if (kwTa) {
        const role = `blank_keyword_${n}`;
        const sel = getSelector(kwTa);
        if (sel && !seen.has(sel) && !selectors[role]) {
          seen.add(sel); selectors[role] = sel;
          fields.push({ role, selector: sel, label: `第${n}小题关键字` });
        }
      }

      // 参考答案（standardAnswer，和 answer 共存时优先用 standardAnswer）
      const stdTa = part.querySelector(".standardAnswer textarea[id]");
      if (stdTa) {
        const role = `blank_answer_${n}`;
        const sel = getSelector(stdTa);
        // 如果已经有 blank_answer_N 了（来自 .answer），用 standardAnswer 覆盖（更精确）
        if (sel && !seen.has(sel)) {
          seen.add(sel);
          const existIdx = fields.findIndex(f => f.role === role);
          if (existIdx >= 0) { fields[existIdx].selector = sel; selectors[role] = sel; }
          else { selectors[role] = sel; fields.push({ role, selector: sel, label: `第${n}小题参考答案` }); }
        }
      }

      // 小题选项 A/B/C/D（多小题题型，选项为普通 input，按子题顺序追加到 option_a/b/c/d）
      // 使用 partSel 前缀生成唯一选择器，避免不同小题同 name 的 input 互相覆盖
      // 注意：audio input 已经在上方通过 usedEls.add(audioInput) 标记，此处只需排除 usedEls 已占用的元素；
      // 不再按 id/class 排除含"image"/"filename"的 input，否则图片选项（class/id 含这些关键字）会被漏掉
      const optionInputs = Array.from(part.querySelectorAll(
        "input[type='text'], input:not([type='radio']):not([type='checkbox']):not([type='file']):not([type='hidden'])"
      )).filter(inp => {
        if (usedEls.has(inp)) return false; // 已被其他 role 使用（audio/script/answer 等）
        const id = (inp.id || "").toLowerCase();
        const cls = (inp.className || "").toLowerCase();
        // 只排除明确属于音频域的 input（含"audio"的 id/class），保留图片选项 input
        return !id.includes("audio") && !cls.includes("audio");
      });
      const optRoles = ["option_a", "option_b", "option_c", "option_d"];
      optionInputs.slice(0, 4).forEach((inp, idx) => {
        const role = optRoles[idx];
        if (!role) return;
        const basicSel = getSelector(inp);
        if (!basicSel) return;
        usedEls.add(inp); // 用元素引用标记，跨小题 name 相同也不会误判
        // 拼接 partSel 前缀，确保不同小题同 name 的 input 有唯一选择器
        partOptSels[role].push(`${partSel} ${basicSel}`);
      });
    });

    // part 专属选择器收集完毕后，整体替换全局扫描的通用选择器
    // 这样 selectorsForRole 的长度 = 小题数，填充时索引一一对应
    ["option_a", "option_b", "option_c", "option_d"].forEach(role => {
      if (partOptSels[role].length === 0) return;
      const finalSel = partOptSels[role].join(",");
      selectors[role] = finalSel;
      const f = fields.find(f => f.role === role);
      if (f) f.selector = finalSel;
      else fields.push({ role, selector: finalSel, label: `选项${role.slice(-1).toUpperCase()}` });
    });
  }

  // 检测是否存在「顶层音频输入框」（即大题共享音频，与各小题 blank_audio_N 区分）
  // 判据：selectors 里有 audio_url，且存在 blank_audio_N（说明是多小题题型）
  // hasTopLevelAudio = true  → 情形 B（共享对话放顶层）
  // hasTopLevelAudio = false → 情形 A（各小题音频独立，应按段拆分）
  const hasBlankAudio = Object.keys(selectors).some(k => /^blank_audio_\d+$/.test(k));
  const hasTopLevelAudio = hasBlankAudio && !!(selectors.audio_url);

  return { selectors, fields, hasTopLevelAudio };
}

// ── 调试桥：在页面控制台输入下方命令可查看 detectForm 检测结果 ──
// document.addEventListener('__extDebugDetectFormResult', e => console.table(e.detail));
// document.dispatchEvent(new Event('__extDebugDetectForm'));
document.addEventListener('__extDebugDetectForm', () => {
  try {
    const r = detectForm();
    const detail = r.fields.map(f => ({ role: f.role, selector: f.selector, label: f.label }));
    console.table(detail);
    document.dispatchEvent(new CustomEvent('__extDebugDetectFormResult', { detail }));
  } catch (e) {
    console.error('[ext debug]', e);
  }
});

/** 当前兼容的试卷接口格式：data.topic 含 topicContent、topicOption 两种形态、topicAttachment 多附件、题目属性等，写死解析 */
function topicApiToQuestionJson(topic) {
  if (!topic || (topic.topicContent == null && topic.topicID == null)) return null;
  const stripHtml = (s) => {
    if (s == null || typeof s !== "string") return "";
    const div = document.createElement("div");
    div.innerHTML = s;
    return (div.textContent || div.innerText || "").trim();
  };
  const normUrl = (u) => (u && typeof u === "string" ? u.replace(/\\\\\//g, "/").trim() : "");

  let options = [];
  let answer = (topic.answer || "").toString().trim();
  let blanks = null;
  let type = null;
  if (topic.topicOption) {
    try {
      const arr = typeof topic.topicOption === "string" ? JSON.parse(topic.topicOption) : topic.topicOption;
      if (Array.isArray(arr) && arr.length > 0) {
        const first = arr[0];
        const isChoice = first && (first.optionDesc != null || first.isTrue != null || first.option === "A" || first.option === "B");
        if (isChoice) {
          options = arr.map((item) => stripHtml(item.optionDesc || ""));
          if (!answer && arr.find((o) => o.isTrue)) answer = (arr.find((o) => o.isTrue).option || "").toString().trim();
        } else {
          type = "listening_fill";
          blanks = arr.map((item) => ({
            question: (item.topicStem != null ? String(item.topicStem) : "").trim(),
            answer: (item.answer != null ? String(item.answer) : "").trim(),
          }));
        }
      }
    } catch (_) {}
  }

  const obj = {
    question: stripHtml(topic.topicContent || ""),
    options: options.filter(Boolean),
    answer: answer || undefined,
    explanation: (topic.analysis || "").toString().trim() || undefined,
  };
  if (type) obj.type = type;
  if (blanks && blanks.length > 0) obj.blanks = blanks;
  if (!obj.question && blanks && blanks.length > 0) obj.question = blanks.map((b) => b.question).filter(Boolean).join(" ").slice(0, 200) || "";

  if ((topic.audioOriginalText || "").toString().trim()) obj.listening_script = topic.audioOriginalText.trim();
  if (topic.topicAttachment) {
    try {
      const att = typeof topic.topicAttachment === "string" ? JSON.parse(topic.topicAttachment) : topic.topicAttachment;
      const list = Array.isArray(att) ? att : [att];
      for (const a of list) {
        if (!a) continue;
        const t = Number(a.attachmentType);
        const url = normUrl(a.attachmentUrl || a.url);
        if (!url) continue;
        if (t === 1 && !obj.audio_url) obj.audio_url = url;
        else if (t === 3 && !obj.image_url) obj.image_url = url;
      }
      if (!obj.audio_url && list[0]) obj.audio_url = normUrl(list[0].attachmentUrl || list[0].url);
    } catch (_) {}
  }
  // 题目属性（按 get?topicID=... 响应写死）
  if ((topic.courseTxt || "").toString().trim()) obj.course = topic.courseTxt.trim();
  if (topic.difficulty != null && topic.difficulty !== "") {
    const d = Number(topic.difficulty);
    if (d === 1) obj.difficulty = "简单";
    else if (d === 2) obj.difficulty = "中等";
    else if (d === 3) obj.difficulty = "困难";
    else obj.difficulty = String(topic.difficulty);
  }
  if ((topic.knowledgeTxt || "").toString().trim()) obj.knowledge_point = topic.knowledgeTxt.trim();
  if (topic.permissionID != null && topic.permissionID !== "") {
    const p = Number(topic.permissionID);
    if (p === 1) obj.question_permission = "公开";
    else if (p === 2) obj.question_permission = "仅自己可见";
    else obj.question_permission = String(topic.permissionID);
  }
  if (topic.volume && typeof topic.volume === "object" && (topic.volume.name || "").toString().trim())
    obj.grade = topic.volume.name.trim();
  if (Array.isArray(topic.teachingIdTxt) && topic.teachingIdTxt.length > 0) {
    const first = topic.teachingIdTxt[0];
    if (first && (first.teachingTxt || "").toString().trim()) obj.unit = first.teachingTxt.trim();
  }
  return obj;
}

/** 从 DOM 读取一个选择器对应的值（input/textarea 的 value，contenteditable 的 textContent，UEditor 的 getContent） */
function readFieldValue(selector) {
  if (!selector || typeof selector !== "string") return "";
  try {
    const el = document.querySelector(selector);
    if (!el) return "";
    // 单选框：只返回当前选中项的值，未选中返回空，避免导出成 answers: ["A","B","C","D"]
    if (el.tagName === "INPUT" && (el.type || "").toLowerCase() === "radio") {
      return el.checked ? (el.value || "").trim() : "";
    }
    if (el.tagName === "INPUT" || el.tagName === "TEXTAREA") return (el.value || "").trim();
    if (el.isContentEditable) return (el.textContent || el.innerText || "").trim();
    if ((el.tagName === "SCRIPT" && (el.type || "").toLowerCase() === "text/plain") || (el.classList && el.classList.contains("edui-default"))) {
      const id = el.id || el.getAttribute("name");
      if (id && !/^edui\d+$/i.test(id) && typeof window.UE !== "undefined") {
        try {
          const ed = window.UE.getEditor(id);
          if (ed && typeof ed.getContent === "function") return (ed.getContent() || "").trim();
        } catch (_) {}
      }
      const container = el.classList && el.classList.contains("edui-default") ? el : el.closest(".edui-default");
      if (container) {
        try {
          const iframe = container.querySelector("iframe");
          if (iframe && iframe.contentDocument && iframe.contentDocument.body) {
            const body = iframe.contentDocument.body;
            return (body.innerText || body.textContent || "").trim();
          }
        } catch (_) {}
      }
      return (el.textContent || "").trim();
    }
  } catch (_) {}
  return "";
}

/** 把当前题目卡片内识别到的表单字段读成一道题的 JSON，格式与填充用的 JSON 一致 */
function getCurrentFormValuesAsJson() {
  const { selectors } = detectForm();
  const qSel = (key) => {
    const s = selectors[key];
    if (!s) return [];
    return typeof s === "string" ? s.split(",").map((x) => x.trim()).filter(Boolean) : [s];
  };
  const opts = [
    readFieldValue(qSel("option_a")[0]),
    readFieldValue(qSel("option_b")[0]),
    readFieldValue(qSel("option_c")[0]),
    readFieldValue(qSel("option_d")[0]),
  ].filter(Boolean);
  const answerSels = qSel("answer");
  const answers = answerSels.length > 1 ? answerSels.map((sel) => readFieldValue(sel)) : [];
  const singleAnswer = answerSels[0] ? readFieldValue(answerSels[0]) : "";
  // 多选时：若实际只有一个被选中（单选框组），只导出 answer 不导出 answers 数组
  const checkedAnswers = (answers || []).filter(Boolean);
  const oneExplanation = readFieldValue(qSel("explanation")[0]) || "";
  const explanationTrim = (oneExplanation === "请在此输入解析内容" || oneExplanation === "请在此输入解析内容。") ? "" : oneExplanation;
  const obj = {
    question: readFieldValue(qSel("question")[0]) || "",
    options: opts.length ? opts : [],
    answer: answerSels.length <= 1 ? singleAnswer : (checkedAnswers.length === 1 ? checkedAnswers[0] : (checkedAnswers.length > 1 ? undefined : singleAnswer)),
    explanation: explanationTrim,
  };
  const keyword = readFieldValue(qSel("keyword")[0]);
  if (keyword) obj.keyword = keyword;
  const listening = readFieldValue(qSel("listening_script")[0]);
  if (listening) obj.listening_script = listening;
  if (answerSels.length > 1 && checkedAnswers.length > 1) obj.answers = answers;
  const audioEl = document.querySelector("#audioFileName, input.audioFileName, [placeholder*='mp3']");
  if (audioEl && (audioEl.value || audioEl.getAttribute("data-url"))) {
    obj.audio_url = (audioEl.getAttribute("data-url") || audioEl.value || "").trim() || undefined;
  }
  const grade = readFieldValue(qSel("grade")[0]);
  if (grade) obj.grade = grade;
  const course = readFieldValue(qSel("course")[0]);
  if (course) obj.course = course;
  const unit = readFieldValue(qSel("unit")[0]);
  if (unit) obj.unit = unit;
  const knowledge_point = readFieldValue(qSel("knowledge_point")[0]);
  if (knowledge_point) obj.knowledge_point = knowledge_point;
  const difficulty = readFieldValue(qSel("difficulty")[0]);
  if (difficulty) obj.difficulty = difficulty;
  const question_permission = readFieldValue(qSel("question_permission")[0]);
  if (question_permission) obj.question_permission = question_permission;
  const recorder = readFieldValue(qSel("recorder")[0]);
  if (recorder) obj.recorder = recorder;
  return obj;
}

/** 遍历每题，分别读取表单并返回题目数组 JSON */
async function getAllFormValuesAsJson() {
  const numbers = findTopicNumbers();
  const list = [];
  if (!numbers.length) {
    list.push(getCurrentFormValuesAsJson());
    return list;
  }
  for (let i = 0; i < numbers.length; i++) {
    await clickTopicNumber(numbers, i);
    list.push(getCurrentFormValuesAsJson());
  }
  await clickTopicNumber(numbers, 0);
  return list;
}

const FILL_DEBUG_ROLES = ["listening_script", "question", "keyword", "explanation", "option_a", "option_b", "option_c", "option_d", "answer"];

/** 调试用：不真正填充，只返回当前页检测到的选择器、每题要填的值、以及每个选择器能否找到元素 */
function runFillDebug(questions, baseSel) {
  const q = questions[0];
  if (!q) return { curSel: {}, roles: [], hint: "没有题目数据" };
  let curSel = { ...baseSel };
  try {
    const { selectors: fresh } = detectForm();
    if (fresh && Object.keys(fresh).length > 0) curSel = { ...baseSel, ...fresh };
  } catch (_) {}
  const qSel = (key) => {
    const s = curSel[key];
    if (!s) return [];
    return typeof s === "string" ? s.split(",").map((x) => x.trim()).filter(Boolean) : [s];
  };
  const getOptionsArray = (q) => {
    let arr = q.options || q.option;
    if (Array.isArray(arr)) return arr;
    if (arr && typeof arr === "object" && !Array.isArray(arr))
      return ["A", "B", "C", "D"].map((k) => (arr[k] != null ? String(arr[k]).trim() : "")).filter(Boolean);
    return [];
  };
  const opts = getOptionsArray(q);
  const getValue = (role) => {
    switch (role) {
      case "listening_script": return (q.listening_script != null ? q.listening_script : "").toString().trim();
      case "explanation": return (q.explanation || "").toString().trim();
      case "question": return (q.question || "").toString().trim();
      case "keyword": return (q.keyword || "").toString().trim();
      case "option_a": return (opts[0] != null ? opts[0] : "").toString().trim();
      case "option_b": return (opts[1] != null ? opts[1] : "").toString().trim();
      case "option_c": return (opts[2] != null ? opts[2] : "").toString().trim();
      case "option_d": return (opts[3] != null ? opts[3] : "").toString().trim();
      case "answer": return (q.answer != null ? q.answer : "").toString().trim();
      default: return "";
    }
  };
  const roles = [];
  for (const role of FILL_DEBUG_ROLES) {
    const selectors = qSel(role);
    const value = getValue(role);
    const checks = selectors.map((sel) => {
      let found = false;
      try { found = !!document.querySelector(sel); } catch (_) {}
      return { selector: sel, found };
    });
    roles.push({
      role,
      valuePreview: value.length > 50 ? value.slice(0, 50) + "…" : value,
      valueLength: value.length,
      selectors: checks,
      summary: selectors.length === 0 ? "未检测到选择器" : (checks.every((c) => c.found) ? "元素均存在" : checks.filter((c) => !c.found).length + " 个选择器找不到元素"),
    });
  }
  return { curSel, roles, hint: "复制下方报告发给我可帮助排查" };
}

// ─── 填写表单 ──────────────────────────────────────────────────────────────

/**
 * 处理 TTS 文本：过滤音标，保留英文单词
 * 规则：
 * 1. 如果文本中既有英文单词又有音标，过滤掉音标，只保留英文单词
 * 2. 如果全是音标（没有普通英文单词），保留音标用于合成
 * 3. 音标格式：/.../ 或 /'.../ 或 /ˈ.../ 等（国际音标）
 */
function prepareTtsText(text) {
  if (!text || typeof text !== 'string') return text;
  
  // 音标模式：匹配 /.../ 格式的音标（包含各种音标符号）
  // 国际音标常见符号：ˈˌːəɪʊæɑɒɔɛɜʌŋθðʃʒtʃdʒ 等
  const phoneticPattern = /\/[^\/]+\//g;
  
  // 检查是否包含音标
  const hasPhonetics = phoneticPattern.test(text);
  if (!hasPhonetics) return text;
  
  // 重置正则的 lastIndex
  phoneticPattern.lastIndex = 0;
  
  // 检查是否有普通英文单词（不在音标内的英文）
  // 先移除所有音标，看剩下的文本是否有英文单词
  const textWithoutPhonetics = text.replace(phoneticPattern, ' ').trim();
  const hasEnglishWords = /[a-zA-Z]{2,}/.test(textWithoutPhonetics);
  
  if (hasEnglishWords) {
    // 有英文单词，过滤掉音标
    // 同时处理可能的序号格式如 "1. truck /trʌk/" => "1. truck"
    let result = text
      .replace(phoneticPattern, '')  // 移除音标
      .replace(/\s{2,}/g, ' ')       // 多个空格合并为一个
      .trim();
    return result;
  } else {
    // 全是音标，保留原文用于合成
    return text;
  }
}

/**
 * 带重试的 TTS 合成请求
 * @param {Object} params TTS 参数
 * @param {string} params.text 要合成的文本
 * @param {boolean} params.dialogue 是否对话格式
 * @param {string} params.provider TTS 服务商
 * @param {string} params.femaleVoice 女声音色
 * @param {string} params.maleVoice 男声音色
 * @param {number} params.femaleSpeed 女声语速
 * @param {number} params.maleSpeed 男声语速
 * @param {number} params.femaleVolume 女声音量
 * @param {number} params.maleVolume 男声音量
 * @param {string} params.contextTexts 上下文文本（豆包专用）
 * @param {number} maxRetries 最大重试次数
 * @param {string} logPrefix 日志前缀
 * @param {Function} logFn 日志函数（可选，默认 console.log）
 * @returns {Promise<string|null>} 成功返回 base64 音频数据，失败返回 null
 */
async function ttsWithRetry(params, maxRetries = 3, logPrefix = "", logFn = null) {
  const _log = logFn || ((...args) => console.log("[AI录题TTS]", ...args));
  const {
    text, dialogue, provider,
    femaleVoice, maleVoice,
    femaleSpeed, maleSpeed,
    femaleVolume, maleVolume,
    contextTexts
  } = params;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      if (attempt > 1) {
        _log(`${logPrefix}TTS 第 ${attempt} 次重试…`);
        await new Promise(r => setTimeout(r, 1000 * attempt)); // 递增延迟
      }
      
      const resp = await fetch("http://127.0.0.1:8766/api/tts", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          text,
          dialogue,
          format: "mp3",
          sample_rate: 24000,
          provider,
          speed_ratio: dialogue ? undefined : femaleSpeed,
          volume_ratio: dialogue ? undefined : femaleVolume,
          female_speaker: femaleVoice,
          male_speaker: maleVoice,
          female_speed: femaleSpeed,
          male_speed: maleSpeed,
          female_volume: femaleVolume,
          male_volume: maleVolume,
          context_texts: provider === "doubao" ? (contextTexts || undefined) : undefined,
        }),
      });
      
      if (resp.ok) {
        const data = await resp.json();
        if (data.audioBase64) {
          _log(`${logPrefix}TTS 合成成功（第 ${attempt} 次），音频大小: ${data.audioBase64.length} 字符`);
          return data.audioBase64;
        } else {
          _log(`${logPrefix}TTS 返回无音频数据（第 ${attempt} 次）`);
        }
      } else {
        const errText = await resp.text().catch(() => "");
        _log(`${logPrefix}TTS 失败 HTTP ${resp.status}（第 ${attempt} 次）: ${errText.slice(0, 80)}`);
      }
    } catch (e) {
      _log(`${logPrefix}TTS 异常（第 ${attempt} 次）: ${e.message}`);
    }
  }
  
  _log(`${logPrefix}TTS 合成失败，已重试 ${maxRetries} 次`);
  return null;
}

async function runFill(questions, selectors, defaultAudioUrl, defaultImageUrl, debugSource, ttsSettings) {
  // TTS 设置默认值
  const tts = ttsSettings || {};
  const ttsProvider = tts.provider || "doubao";  // 服务商：doubao、youdao 或 edge
  // 根据服务商设置默认音色
  let ttsFemaleVoice, ttsMaleVoice;
  if (ttsProvider === "doubao") {
    ttsFemaleVoice = tts.femaleVoice || "S_hWsL9lNS1";
    ttsMaleVoice = tts.maleVoice || "S_iWsL9lNS1";
  } else if (ttsProvider === "youdao") {
    ttsFemaleVoice = tts.femaleVoice || "youxiaodao";
    ttsMaleVoice = tts.maleVoice || "youxiaoguan";
  } else if (ttsProvider === "edge") {
    ttsFemaleVoice = tts.femaleVoice || "en-US-AvaMultilingualNeural";
    ttsMaleVoice = tts.maleVoice || "en-US-GuyNeural";
  } else {
    ttsFemaleVoice = tts.femaleVoice || "S_hWsL9lNS1";
    ttsMaleVoice = tts.maleVoice || "S_iWsL9lNS1";
  }
  const ttsFemaleSpeed = tts.femaleSpeed || 0.85;
  const ttsMaleSpeed = tts.maleSpeed || 0.85;
  const ttsFemaleVolume = tts.femaleVolume || 1.0;
  const ttsMaleVolume = tts.maleVolume || 1.0;
  const ttsContextTexts = tts.contextTexts || "";  // 声音复刻提示词（仅豆包）
  const baseSel = selectors || {};
  const qSel = (sel, key) => {
    const s = sel[key];
    if (!s) return [];
    return typeof s === "string" ? s.split(",").map((x) => x.trim()).filter(Boolean) : [s];
  };

  // 日志：从 AI 识别结果到填充全链路，复制用 copy(localStorage.getItem('__aiLutiLog'))
  const _fillLog = [];
  const log = (...args) => {
    const msg = args.map(a => (typeof a === "object" ? JSON.stringify(a) : String(a))).join(" ");
    const line = `[${new Date().toLocaleTimeString()}] ${msg}`;
    _fillLog.push(line);
    console.log("[AI录题填充]", ...args);
    try { localStorage.setItem("__aiLutiLog", _fillLog.join("\n")); } catch (_) {}
  };
  try { localStorage.removeItem("__aiLutiLog"); } catch (_) {}
  const srcLabel = debugSource === "parse" ? "解析Word" : debugSource === "json" ? "JSON粘贴" : "填充";
  log(`=== 收到填充请求 来源=${srcLabel} 题数=${questions?.length ?? 0} | 复制日志: 录题页控制台执行 copy(localStorage.getItem('__aiLutiLog'))`);
  if (questions?.length) {
    questions.forEach((q, idx) => {
      const type = q?.type ?? "";
      const blanksLen = q?.blanks?.length ?? 0;
      const answer = q?.answer != null ? String(q.answer).slice(0, 40) : (blanksLen ? `blanks[].answer x${blanksLen}` : "");
      const qLen = (q?.question != null ? String(q.question).length : 0) || (blanksLen ? (q.blanks?.map(b => (b?.question && String(b.question).length) || 0)).join(",") : "");
      log(`  [${idx + 1}] type=${type} blanks=${blanksLen} answer=${answer} question_len=${qLen}`);
    });
  }

  /** 题目合并默认属性（JSON 未提供的用默认） */
  const mergedProps = (role) => {
    const val = q[role];
    if (val != null && String(val).trim() !== "") return String(val).trim();
    return DEFAULT_QUESTION_PROPS[role] != null ? String(DEFAULT_QUESTION_PROPS[role]) : "";
  };

  /** 填写一个字段，支持 input / textarea / contenteditable / select / 答案容器(div.ui-radio) */
  function fillField(selector, value, _debugRole) {
    try {
      // 若选择器含 data-fill-part-idx，先确保属性仍在（防止框架重渲染后丢失）
      if (selector && selector.includes("data-fill-part-idx")) {
        document.querySelectorAll(".question-part").forEach((pt, pi) => pt.setAttribute("data-fill-part-idx", String(pi)));
      }
      const el = document.querySelector(selector);
      if (!el) { if (_debugRole) log(`    fillField(${_debugRole}): 元素不存在 sel=${selector.slice(0,50)}`); return false; }
      if (value == null || String(value).trim() === "") { if (_debugRole) log(`    fillField(${_debugRole}): 值为空`); return false; }
      const strVal = String(value);
      // 支持 "C. Four." / "C）Four." / "C Four." 等格式，提取首字母作为答案
      const vRaw = strVal.trim().toUpperCase();
      const v = /^([A-D])[.\uff09\uff1a\s）:]/i.test(vRaw) ? vRaw.charAt(0) : vRaw;
      // 答案容器：div.ui-radio#answer_1 内为 label(for=answer_1_1..4)，点对应 A/B/C/D 的 label
      if (el.tagName === "DIV" && (el.classList.contains("ui-radio") || el.classList.contains("answer") || (el.id && /^answer_/.test(el.id)))) {
        const labels = el.querySelectorAll("label[for]");
        const abcd = ["A", "B", "C", "D"];
        const idx = abcd.indexOf(v) >= 0 ? abcd.indexOf(v) + 1 : (parseInt(v, 10) >= 1 && parseInt(v, 10) <= 4 ? parseInt(v, 10) : -1);
        for (const label of labels) {
          const text = (label.textContent || "").trim().toUpperCase().slice(0, 1);
          const forId = label.getAttribute("for") || "";
          const forNum = (forId.match(/_(\d)$/) || [])[1];
          if (text === v || (forNum && idx >= 1 && parseInt(forNum, 10) === idx) || (forNum && v === abcd[parseInt(forNum, 10) - 1])) {
            if (label.click) { label.click(); return true; }
            const input = el.ownerDocument && el.ownerDocument.getElementById(label.getAttribute("for"));
            if (input && input.click) { input.click(); return true; }
          }
        }
        const radios = el.querySelectorAll("input[type=radio]");
        for (const radio of radios) {
          const rv = (radio.value || "").trim().toUpperCase();
          const labelText = (radio.parentElement && radio.parentElement.textContent || "").trim().toUpperCase().slice(0, 1);
          if (rv === v || labelText === v || (v === "A" && (rv === "1" || labelText === "A")) || (v === "B" && (rv === "2" || labelText === "B")) || (v === "C" && (rv === "3" || labelText === "C")) || (v === "D" && (rv === "4" || labelText === "D"))) {
            radio.checked = true;
            radio.dispatchEvent(new Event("change", { bubbles: true }));
            return true;
          }
        }
      }
      // 多小题选项等：选择器指向外层 div/span 时，若内部是普通 input/textarea（无 UEditor），直接填内部
      if ((el.tagName === "DIV" || el.tagName === "SPAN" || el.tagName === "TD") && !el.querySelector(".ueditor-show") && !el.querySelector("script[type='text/plain']")) {
        const inner = el.querySelector("input:not([type=radio]):not([type=checkbox]), textarea");
        if (inner) {
          inner.focus();
          const proto = inner.tagName === "INPUT" ? HTMLInputElement.prototype : HTMLTextAreaElement.prototype;
          const desc = Object.getOwnPropertyDescriptor(proto, "value");
          const nativeSetter = desc && desc.set;
          if (nativeSetter) nativeSetter.call(inner, strVal);
          else inner.value = strVal;
          inner.dispatchEvent(new Event("input", { bubbles: true }));
          inner.dispatchEvent(new Event("change", { bubbles: true }));
          inner.dispatchEvent(new Event("blur", { bubbles: true }));
          return true;
        }
      }
      el.focus();
      if (el.tagName === "SELECT") {
        const opt = Array.from(el.options).find((o) => (o.value && o.value.trim() === strVal) || (o.textContent || "").trim() === strVal);
        if (opt) {
          el.value = opt.value;
          el.dispatchEvent(new Event("change", { bubbles: true }));
          el.dispatchEvent(new Event("input", { bubbles: true }));
          return true;
        }
        if (el.value !== strVal) {
          el.value = strVal;
          el.dispatchEvent(new Event("change", { bubbles: true }));
          el.dispatchEvent(new Event("input", { bubbles: true }));
        }
        return true;
      }
      if (el.tagName === "INPUT" || el.tagName === "TEXTAREA") {
        if ((el.type || "").toLowerCase() === "radio") {
          const v = strVal.trim().toUpperCase();
          const tryMatch = (radio) => {
            const elVal = (radio.value || "").trim().toUpperCase();
            const dataVal = (radio.getAttribute("data-value") || "").trim().toUpperCase();
            const labelText = (radio.parentElement && radio.parentElement.textContent || radio.nextElementSibling && radio.nextElementSibling.textContent || "").trim().toUpperCase().slice(0, 2);
            return elVal === v || dataVal === v || labelText === v || (v === "A" && (elVal === "1" || labelText.startsWith("A"))) || (v === "B" && (elVal === "2" || labelText.startsWith("B"))) || (v === "C" && (elVal === "3" || labelText.startsWith("C"))) || (v === "D" && (elVal === "4" || labelText.startsWith("D")));
          };
          if (tryMatch(el)) {
            el.checked = true;
            el.dispatchEvent(new Event("change", { bubbles: true }));
            el.dispatchEvent(new Event("input", { bubbles: true }));
            return true;
          }
          const name = el.getAttribute("name");
          if (name) {
            const form = el.form || el.closest("form") || document;
            const group = form.querySelectorAll ? form.querySelectorAll("input[type=radio][name=\"" + name.replace(/"/g, '\\"') + "\"]") : [];
            for (const radio of group) {
              if (tryMatch(radio)) {
                radio.checked = true;
                radio.dispatchEvent(new Event("change", { bubbles: true }));
                radio.dispatchEvent(new Event("input", { bubbles: true }));
                return true;
              }
            }
          }
          return false;
        }
        const proto = el.tagName === "INPUT" ? HTMLInputElement.prototype : HTMLTextAreaElement.prototype;
        const desc = Object.getOwnPropertyDescriptor(proto, "value");
        const nativeSetter = desc && desc.set;
        if (nativeSetter) {
          nativeSetter.call(el, strVal);
        } else {
          el.value = strVal;
        }
        el.dispatchEvent(new Event("input",  { bubbles: true }));
        el.dispatchEvent(new Event("change", { bubbles: true }));
        el.dispatchEvent(new Event("blur",  { bubbles: true }));
        return true;
      }
      if (el.isContentEditable) {
        el.focus();
        // execCommand 是框架（Vue/React）能感知的"真实用户输入"，比直接赋 textContent 更可靠
        try {
          document.execCommand("selectAll", false, null);
          document.execCommand("insertText", false, strVal);
        } catch (_) {
          el.textContent = "";
          el.textContent = strVal;
        }
        el.dispatchEvent(new InputEvent("input", { bubbles: true, data: strVal }));
        el.dispatchEvent(new Event("change", { bubbles: true }));
        return true;
      }
      // 富文本选项：选择器常指向外层 div，可编辑区在内部 iframe body 或 contenteditable
      const iframeInEl = el.querySelector && el.querySelector("iframe");
      if (iframeInEl && iframeInEl.contentDocument && iframeInEl.contentDocument.body) {
        try {
          el.focus();
          if (el.click) el.click();
          const body = iframeInEl.contentDocument.body;
          body.focus();
          body.textContent = strVal || "";
          body.dispatchEvent(new InputEvent("input", { bubbles: true, data: strVal }));
          body.dispatchEvent(new Event("change", { bubbles: true }));
          body.dispatchEvent(new Event("blur", { bubbles: true }));
          el.dispatchEvent(new Event("input", { bubbles: true }));
          return true;
        } catch (_) {}
      }
      const contentEditableDesc = el.querySelector && el.querySelector("[contenteditable=true]");
      if (contentEditableDesc && contentEditableDesc.isContentEditable) {
        contentEditableDesc.focus();
        contentEditableDesc.textContent = strVal || "";
        contentEditableDesc.dispatchEvent(new InputEvent("input", { bubbles: true, data: strVal }));
        return true;
      }
      // 驰声 ui-ueditor：.ui-ueditor > .ueditor-show + script[type=text/plain]，先写 script 和 ueditor-show 再触发展开后的 iframe
      const uiUeditor = el.closest && el.closest(".ui-ueditor") || (el.parentElement && el.parentElement.classList && el.parentElement.classList.contains("ui-ueditor") ? el.parentElement : null);
      if (uiUeditor) {
        const writeUeditor = () => {
          const scriptTag = uiUeditor.querySelector("script[type='text/plain']");
          const showDiv = uiUeditor.querySelector(".ueditor-show");
          if (scriptTag) {
            scriptTag.textContent = strVal || "";
            scriptTag.dispatchEvent(new Event("input", { bubbles: true }));
          }
          if (showDiv) {
            showDiv.textContent = strVal || "";
            showDiv.dispatchEvent(new InputEvent("input", { bubbles: true, data: strVal }));
            showDiv.dispatchEvent(new Event("change", { bubbles: true }));
          }
          const iframe = uiUeditor.querySelector("iframe");
          if (iframe && iframe.contentDocument && iframe.contentDocument.body) {
            const body = iframe.contentDocument.body;
            body.focus();
            body.textContent = strVal || "";
            body.dispatchEvent(new InputEvent("input", { bubbles: true, data: strVal }));
            body.dispatchEvent(new Event("change", { bubbles: true }));
          }
        };
        try {
          writeUeditor();
          // 若 iframe 尚未就绪导致未写入，稍后重试一次（驰声 UEditor 有时延迟加载）
          const iframe = uiUeditor.querySelector("iframe");
          if (iframe && iframe.contentDocument && iframe.contentDocument.body) {
            const body = iframe.contentDocument.body;
            if (strVal && (!body.textContent || body.textContent.trim() === "")) {
              setTimeout(() => { try { writeUeditor(); } catch (_) {} }, 280);
            }
          }
          if (uiUeditor.querySelector("script[type='text/plain']") || uiUeditor.querySelector(".ueditor-show")) return true;
        } catch (_) {}
      }
      // 驰声 UEditor：题干/选项为 script[type=text/plain] 或 div.edui-default，通过 UE.getEditor(id).setContent 写入
      const isUeScript = el.tagName === "SCRIPT" && (el.type || "").toLowerCase() === "text/plain";
      const isUeContainer = el.classList && el.classList.contains("edui-default") && el.id && !/^edui\d+$/i.test(el.id);
      const ueditorId = el.id || el.getAttribute("name");
      if (ueditorId && (isUeScript || isUeContainer) && typeof window.UE !== "undefined") {
        try {
          const ed = window.UE.getEditor(ueditorId);
          if (ed && typeof ed.setContent === "function") {
            ed.setContent(strVal);
            return true;
          }
        } catch (_) {}
      }
      // 兜底：题干/选项容器内的 iframe 编辑区直接写 body（UE 未就绪或 getEditor 取不到时）
      let container = el.classList && el.classList.contains("edui-default") ? el : el.closest(".edui-default");
      if (!container && el.tagName === "SCRIPT") {
        container = el.nextElementSibling || (el.parentElement && el.parentElement.querySelector(".edui-default"));
        if (!container && el.parentElement) {
          const iframeInParent = el.parentElement.querySelector("iframe");
          if (iframeInParent && iframeInParent.contentDocument) container = iframeInParent.parentElement;
        }
      }
      if (!container && el.querySelector && el.querySelector("iframe")) container = el;
      if (container) {
        try {
          container.focus();
          if (container.click) container.click();
          const iframe = container.querySelector("iframe");
          if (iframe && iframe.contentDocument && iframe.contentDocument.body) {
            const body = iframe.contentDocument.body;
            body.focus();
            body.textContent = strVal || "";
            body.dispatchEvent(new InputEvent("input", { bubbles: true, data: strVal }));
            body.dispatchEvent(new Event("change", { bubbles: true }));
            body.dispatchEvent(new Event("blur", { bubbles: true }));
            if (container !== el) container.dispatchEvent(new Event("input", { bubbles: true }));
            return true;
          }
        } catch (_) {}
      }
      // 到这里说明所有路径都未成功
      if (_debugRole) log(`    fillField(${_debugRole}): 无匹配的填充方式 tag=${el.tagName} class=${el.className?.slice?.(0,50) || ""}`);
    } catch (e) { if (_debugRole) log(`    fillField(${_debugRole}): 异常 ${e?.message}`); }
    return false;
  }

  /** 下拉框：点击触发元素展开，再在页面内点击包含 value 文本的选项（适用年级、课程、单元、知识点等） */
  function clickDropdownAndSelect(selector, value) {
    if (!value || !selector) return Promise.resolve(false);
    try {
      const trigger = document.querySelector(selector);
      if (!trigger) return Promise.resolve(false);
      const v = String(value).trim();
      if (trigger.tagName === "INPUT" || trigger.tagName === "TEXTAREA") {
        trigger.focus();
        trigger.value = v;
        trigger.dispatchEvent(new Event("input", { bubbles: true }));
        trigger.dispatchEvent(new Event("change", { bubbles: true }));
        trigger.dispatchEvent(new Event("blur", { bubbles: true }));
      }
      // 如果 trigger 是级联内部的 readonly input，升级到外层容器
      let cascaderRoot = null;
      if (trigger.tagName === "INPUT" && trigger.readOnly) {
        cascaderRoot = trigger.closest(".jsCascaderWarp, [class*='cascader-input'], [class*='cascader-wrap']");
      }
      if (!cascaderRoot && trigger.classList && (trigger.classList.contains("jsCascaderWarp") || trigger.classList.contains("cascader-input") || trigger.classList.contains("topic-cascader"))) {
        cascaderRoot = trigger;
      }
      const isTreeBox = trigger.classList && (trigger.classList.contains("ui-comboBox") || trigger.classList.contains("ui-comboTree"));
      const isUiSelect = trigger.classList && trigger.classList.contains("ui-select");
      if (cascaderRoot) {
        // 级联组件：点 inputWrap 触发展开
        const inputWrap = cascaderRoot.querySelector(".inputWrap") || cascaderRoot.querySelector("input") || cascaderRoot;
        if (inputWrap.click) inputWrap.click();
      } else if (isTreeBox) {
        // ztree 树形组件：点击输入框触发展开（不点展开按钮，避免折叠）
        const inp = trigger.querySelector("input") || trigger;
        if (inp.click) inp.click();
      } else if (isUiSelect) {
        // .ui-select 自定义下拉：点内部 input 展开
        const inp = trigger.querySelector("input") || trigger;
        if (inp.click) inp.click();
      } else {
        trigger.focus();
        trigger.click();
      }

      const tryFindAndClick = (v) => {
        // .ui-select 自定义下拉：搜索 trigger 内的 ul li（仅在其容器内搜索，避免跨组件）
        if (isUiSelect) {
          const lis = trigger.querySelectorAll("ul li");
          for (const li of lis) {
            const text = (li.textContent || "").trim();
            if (text && text.length < 80 && (text === v || text.includes(v) || v.includes(text))) {
              li.click();
              return true;
            }
          }
        }
        // .ui-comboTree 树形选择：在容器内的 ztree 节点里搜索
        if (isTreeBox) {
          const treeNodes = trigger.querySelectorAll("a.treenode_a, a[class*='treenode'], .ztree a, li a");
          for (const a of treeNodes) {
            const text = (a.textContent || a.title || "").trim();
            if (text && text.length < 80 && (text === v || text.includes(v) || v.includes(text))) {
              a.click();
              return true;
            }
          }
        }
        // 驰声级联：在当前级联容器内搜索 li，避免多级联互扰
        const cascaderScope = cascaderRoot || (trigger.tagName === "INPUT" && trigger.readOnly ? trigger.closest(".jsCascaderWarp, [class*='cascader']") : null);
        if (cascaderScope) {
          const contentWrap = cascaderScope.querySelector(".cascaderContentWarp, [class*='cascader-content'], [class*='cascaderContent']");
          if (contentWrap) {
            for (const li of contentWrap.querySelectorAll("li[data-value], li")) {
              const text = (li.textContent || "").trim();
              if (text && text.length < 40 && (text === v || text.includes(v) || v.includes(text))) {
                li.click();
                return true;
              }
            }
          }
        }
        // 兜底：全页搜索级联内容
        const cascaderContent = document.querySelectorAll(".cascaderContentWarp, [class*='cascader-content'], [class*='cascaderContent']");
        for (const panel of cascaderContent) {
          for (const li of panel.querySelectorAll("li[data-value], li")) {
            const text = (li.textContent || "").trim();
            if (text && text.length < 40 && (text === v || text.includes(v) || v.includes(text))) {
              li.click();
              return true;
            }
          }
        }
        // ztree 树节点：a.treenode_a
        const treeNodes = document.querySelectorAll("a.treenode_a, a[class*='treenode'], .ztree a");
        for (const a of treeNodes) {
          const text = (a.textContent || a.title || "").trim();
          if (text && text.length < 80 && (text === v || text.includes(v) || v.includes(text))) {
            a.click();
            return true;
          }
        }
        // 通用候选
        const candidates = document.querySelectorAll(
          ".el-select-dropdown__item, .el-cascader-node__label, .ant-select-item, [role='option'], .dropdown-item, .option-item, li[class*='select'], div[class*='option']"
        );
        for (const el of candidates) {
          if (el.offsetParent === null) continue;
          const text = (el.textContent || el.innerText || el.getAttribute("data-value") || "").trim();
          if (!text || text.length > 80) continue;
          if (text === v || text.includes(v) || v.includes(text)) { el.click(); return true; }
        }
        const panels = document.querySelectorAll(".el-select-dropdown, .ant-select-dropdown, [role='listbox'], [class*='select-dropdown'], [class*='dropdown-menu']");
        for (const panel of panels) {
          if (panel.offsetParent === null) continue;
          const items = panel.querySelectorAll("li, .el-cascader-node, [role='option'], div[class*='item'], div[class*='option']");
          for (const el of items) {
            const text = (el.textContent || el.innerText || "").trim();
            if (text && text.length < 80 && (text === v || text.includes(v) || v.includes(text))) { el.click(); return true; }
          }
        }
        return false;
      };
      return new Promise((resolve) => {
        // 先等 350ms，再试一次；若失败再等 400ms 重试
        setTimeout(() => {
          if (tryFindAndClick(v)) { resolve(true); return; }
          setTimeout(() => {
            if (tryFindAndClick(v)) { resolve(true); return; }
            if (trigger.tagName === "INPUT" || trigger.tagName === "TEXTAREA") resolve(true);
            else resolve(false);
          }, 400);
        }, 350);
      });
    } catch (_) {
      return Promise.resolve(false);
    }
  }

  /** 在多个候选中点击文本与 value 匹配的那一项（用于难度、题目权限等单选） */
  function clickOptionByLabel(selectors, value, scopeSelector) {
    if (!value || !selectors || selectors.length === 0) return false;
    const v = String(value).trim();
    const root = scopeSelector ? document.querySelector(scopeSelector) : null;
    const searchIn = (el) => {
      if (!el) return false;
      const label = getFieldLabel(el);
      const text = (el.textContent || "").trim()
        || (el.nextElementSibling && el.nextElementSibling.textContent || "").trim()
        || (el.parentElement && el.parentElement.textContent || "").trim()
        || label;
      if (text && text.length < 100 && (text.includes(v) || v.includes(text))) {
        const toClick = el.click ? el : (el.querySelector("input") || el);
        if (toClick && toClick.click) { toClick.click(); return true; }
      }
      const scope = el.closest(".row, .form-item, .el-form-item, [class*='form']") || el;
      const allClickables = scope.querySelectorAll("input[type='radio'], [role='radio'], .el-radio, .el-radio-button, .ant-radio-wrapper, label, span, div[class*='radio'], button");
      for (const r of allClickables) {
        const t = (r.textContent || r.getAttribute("aria-label") || r.getAttribute("title") || "").trim();
        if (t && t.length < 50 && (t.includes(v) || v.includes(t))) {
          const toClick = r.tagName === "INPUT" ? r : (r.querySelector("input") || r) || r;
          if (toClick && toClick.click) { toClick.click(); return true; }
        }
      }
      const parent = el.closest(".el-radio, .el-radio-button, [role='radio'], .ant-radio-wrapper, .row");
      if (parent) {
        const all = parent.querySelectorAll("input[type='radio'], [role='radio'], .el-radio, .ant-radio");
        for (const r of all) {
          const t = (r.textContent || r.getAttribute("aria-label") || (r.parentElement && r.parentElement.textContent) || "").trim();
          if (t && (t.includes(v) || v.includes(t))) {
            (r.click ? r : r.querySelector("input") || r).click();
            return true;
          }
        }
      }
      return false;
    };
    for (const sel of selectors) {
      try {
        const el = document.querySelector(sel);
        if (el && searchIn(el)) return true;
      } catch (_) {}
    }
    if (root) {
      const fallbacks = root.querySelectorAll(".el-radio, .el-radio-button, [role='radio'], .ant-radio-wrapper, label, button, span[class*='radio']");
      for (const el of fallbacks) {
        const t = (el.textContent || el.getAttribute("aria-label") || "").trim();
        if (t && t.length < 50 && (t.includes(v) || v.includes(t))) {
          const toClick = el.tagName === "INPUT" ? el : (el.querySelector("input") || el) || el;
          if (toClick && toClick.click) { toClick.click(); return true; }
        }
      }
    }
    return false;
  }

  /** 将 base64 或 data URL 转为 Blob */
  function base64ToBlob(dataUrlOrBase64, defaultMime) {
    let mime = defaultMime || "application/octet-stream";
    let base64 = dataUrlOrBase64.trim();
    if (base64.startsWith("data:")) {
      const i = base64.indexOf(",");
      if (i !== -1) {
        const header = base64.slice(0, i);
        const match = /data:([^;]+);?/.exec(header);
        if (match) mime = match[1].trim();
        base64 = base64.slice(i + 1);
      }
    }
    try {
      const bin = atob(base64.replace(/\s/g, ""));
      const arr = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
      return new Blob([arr], { type: mime });
    } catch (_) {
      return null;
    }
  }

  /** 为文件类型的 input 设置 File（用于上传音频/图片） */
  function fillFileField(selector, file) {
    try {
      const el = document.querySelector(selector);
      if (!el || el.type !== "file" || !(file instanceof File)) return false;
      const dt = new DataTransfer();
      dt.items.add(file);
      el.files = dt.files;
      el.dispatchEvent(new Event("input", { bubbles: true }));
      el.dispatchEvent(new Event("change", { bubbles: true }));
      return true;
    } catch (_) {
      return false;
    }
  }

  /** 从题目对象中解析出音频或图片的 File，供填充上传框使用 */
  async function resolveMediaFile(q, kind, fallbackUrl) {
    const isAudio = kind === "audio";
    const base64Key = isAudio ? "audio_base64" : "image_base64";
    const urlKey = isAudio ? "audio_url" : "image_url";
    const defaultName = isAudio ? "audio.mp3" : "image.png";
    const defaultMime = isAudio ? "audio/mpeg" : "image/png";
    const base64 = q[base64Key];
    const url = (q[urlKey] || "").toString().trim() || (fallbackUrl || "").toString().trim();
    if (base64 && typeof base64 === "string") {
      const blob = base64ToBlob(base64, defaultMime);
      if (blob) return new File([blob], defaultName, { type: blob.type || defaultMime });
    }
    if (url && typeof url === "string" && url.startsWith("http")) {
      try {
        const res = await fetch(url);
        const blob = await res.blob();
        const name = defaultName;
        const mime = blob.type || defaultMime;
        return new File([blob], name, { type: mime });
      } catch (e) { log(`  resolveMediaFile fetch失败: ${e?.message}`); }
    }
    return null;
  }

  // 找左侧所有题号方块
  const topicNumbers = findTopicNumbers();
  let filled = 0;
  let startIndex = 0; // 当前要填的第一题在页面上的题号索引（0-based），用于保存后同步左侧高亮

  const notify = (text) => {
    log("notify:", text);
    try { chrome.runtime.sendMessage({ type: "FILL_PROGRESS", text }); } catch (_) {}
  };

  // 识别到多题时：从左侧当前选中的题号开始填（例如当前是第 3 题则从第 3 题开始填），不切到第一题
  // 识别到只有一题时：不点左侧题号，只填当前页（例如当前是第 3 题就只填第 3 题）
  if (topicNumbers.length > 0 && questions.length > 1) {
    startIndex = getCurrentTopicIndex(topicNumbers);
    const toClick = topicNumbers[startIndex];
    if (toClick) {
      toClick.click();
      await delay(900);
    }
  }

  // ── 页面提示条：显示音频合成状态（悬浮在右上角，精美动效）──
  let ttsStatusBar = document.getElementById("__ai_luti_tts_status");
  if (!ttsStatusBar) {
    ttsStatusBar = document.createElement("div");
    ttsStatusBar.id = "__ai_luti_tts_status";
    ttsStatusBar.style.cssText = `
      position: fixed;
      top: 16px;
      right: 16px;
      z-index: 999999;
      background: linear-gradient(135deg, #c0392b 0%, #e74c3c 100%);
      color: #fff;
      padding: 12px 20px 12px 16px;
      font-size: 14px;
      font-weight: 500;
      border-radius: 8px;
      box-shadow: 0 4px 20px rgba(231, 76, 60, 0.4), 0 2px 8px rgba(0,0,0,0.15);
      display: none;
      align-items: center;
      gap: 10px;
      backdrop-filter: blur(8px);
      transform: translateY(-20px);
      opacity: 0;
      transition: transform 0.3s cubic-bezier(0.34, 1.56, 0.64, 1), opacity 0.3s ease;
    `;
    document.body.appendChild(ttsStatusBar);
    
    // 添加动画样式
    if (!document.getElementById("__ai_tts_styles")) {
      const style = document.createElement("style");
      style.id = "__ai_tts_styles";
      style.textContent = `
        @keyframes __ai_tts_pulse {
          0%, 100% { transform: scale(1); opacity: 1; }
          50% { transform: scale(1.15); opacity: 0.7; }
        }
        @keyframes __ai_tts_wave {
          0%, 100% { height: 8px; }
          50% { height: 16px; }
        }
        .__ai_tts_indicator {
          display: flex;
          align-items: center;
          gap: 3px;
          height: 20px;
        }
        .__ai_tts_bar {
          width: 3px;
          background: #fff;
          border-radius: 2px;
          animation: __ai_tts_wave 0.6s ease-in-out infinite;
        }
        .__ai_tts_bar:nth-child(1) { animation-delay: 0s; }
        .__ai_tts_bar:nth-child(2) { animation-delay: 0.15s; }
        .__ai_tts_bar:nth-child(3) { animation-delay: 0.3s; }
        .__ai_tts_bar:nth-child(4) { animation-delay: 0.45s; }
        .__ai_tts_text {
          font-weight: 500;
          letter-spacing: 0.3px;
        }
      `;
      document.head.appendChild(style);
    }
  }
  const showTtsStatus = (text) => {
    ttsStatusBar.innerHTML = `
      <div class="__ai_tts_indicator">
        <div class="__ai_tts_bar" style="height:12px;"></div>
        <div class="__ai_tts_bar" style="height:8px;"></div>
        <div class="__ai_tts_bar" style="height:16px;"></div>
        <div class="__ai_tts_bar" style="height:10px;"></div>
      </div>
      <span class="__ai_tts_text">${text}</span>
    `;
    ttsStatusBar.style.display = "flex";
    // 触发动画
    requestAnimationFrame(() => {
      ttsStatusBar.style.transform = "translateY(0)";
      ttsStatusBar.style.opacity = "1";
    });
  };
  const hideTtsStatus = () => {
    ttsStatusBar.style.transform = "translateY(-20px)";
    ttsStatusBar.style.opacity = "0";
    setTimeout(() => { ttsStatusBar.style.display = "none"; }, 300);
  };

  for (let i = 0; i < questions.length; i++) {
    const q = questions[i];
    notify(`正在填写第 ${i + 1}/${questions.length} 题…`);
    log(`=== 开始填第 ${i + 1} 题，type=${q.type}，blanks=${q.blanks?.length ?? 0}`);

    // 每题重新检测当前 DOM 选择器（题干/选项等多为动态 ID，切题后会变）
    let curSel = { ...baseSel };
    try {
      const { selectors: fresh } = detectForm();
      if (fresh && Object.keys(fresh).length > 0) {
        curSel = { ...baseSel, ...fresh };
      }
      const allKeys = Object.keys(curSel).filter(k => curSel[k]);
      const audioKeys = allKeys.filter(k => k.includes('audio'));
      log(`detectForm 结果 keys (${allKeys.length}个):`, allKeys.slice(0, 15).join(', '));
      log(`detectForm 音频相关 keys:`, audioKeys.length > 0 ? audioKeys.join(', ') : '无');
    } catch (e) { log("detectForm 报错:", e?.message); }

    // 给当前题表单/UEditor 一点时间就绪再开始填（尤其听后选择等富文本多的题型）
    await delay(350);

    // ── 检测音频上传框（uploadify 容器或普通 file input）──
    let audioFileSels = qSel(curSel, "audio_file");
    let audioUrlSels = qSel(curSel, "audio_url");
    
    if (audioFileSels.length === 0 && audioUrlSels.length === 0) {
      const scope = document.querySelector(".topic-container, #topic-section, #right-container, .question-part") || document.body;
      
      // 检测音频 URL 输入框（扩展选择器）
      const audioUrlInput = scope.querySelector(
        "#audioFileName, #audioFileName1, input.audioFileName, input[id^='audioFileName'], " +
        "input[id*='audio'][type='text'], input[placeholder*='mp3'], input[placeholder*='音频'], " +
        "input[name*='audio'], input[id*='Audio'], .audio-input input[type='text']"
      );
      if (audioUrlInput) {
        const sel = audioUrlInput.id ? `#${audioUrlInput.id}` : (audioUrlInput.name ? `input[name="${audioUrlInput.name}"]` : null);
        if (sel) audioUrlSels = [sel];
        log(`  初始检测到音频URL输入框: ${sel}`);
      }
      
      // 检测音频上传容器（只匹配明确与音频相关的选择器，避免误匹配图片上传框）
      const audioUploadContainer = scope.querySelector(
        "#audio_upload1, #audio_upload, div[id*='audio_upload'], .uploadify[id*='audio'], " +
        "[class*='audio'][class*='upload'], [class*='upload'][class*='audio']"
      );
      if (audioUploadContainer) {
        const fi = audioUploadContainer.querySelector("input[type='file']");
        if (fi) {
          const sel = fi.id ? `#${fi.id}` : "input[type='file']";
          audioFileSels = [sel];
          log(`  初始检测到音频上传框: ${sel}`);
        }
      }
      
      // 如果还没找到，尝试更宽松的匹配
      if (audioFileSels.length === 0) {
        // 查找所有 uploadify 容器
        const allUploadify = scope.querySelectorAll(".uploadify, [class*='uploadify']");
        for (const container of allUploadify) {
          // 判断是否是音频上传（通过周围文本或 ID）
          const parentText = (container.closest(".row, .form-group, label")?.innerText || "").toLowerCase();
          const containerId = (container.id || "").toLowerCase();
          if (parentText.includes("音频") || parentText.includes("录音") || containerId.includes("audio")) {
            const fi = container.querySelector("input[type='file']");
            if (fi) {
              const sel = fi.id ? `#${fi.id}` : "input[type='file']";
              audioFileSels = [sel];
              log(`  宽松匹配到音频上传框: ${sel}`);
              break;
            }
          }
        }
      }
    }

    // ── 异步 TTS 合成：如果有 listening_script 但没有音频，启动异步合成 ──
    let ttsPromise = null;
    let blanksTtsPromises = []; // 子题的 TTS Promise 数组
    
    // 兼容多种字段名：listening_script / listeningScript / script / 听力原文
    let listeningScriptValue = (
      q.listening_script || 
      q.listeningScript || 
      q.script || 
      q["听力原文"] || 
      ""
    ).trim();
    
    // 调试日志：显示题目类型和所有相关字段
    log(`第 ${i + 1} 题详情: type="${q.type || '(无)'}", question="${(q.question || '').slice(0, 50)}", listening_script="${(q.listening_script || '').slice(0, 30)}", blanks=${q.blanks?.length || 0}`);
    
    // 检查是否有子题需要单独合成 TTS
    const blanks = q.blanks || [];
    const blankAudioKeys = Object.keys(curSel).filter(k => /^blank_audio_\d+$/.test(k));
    const hasBlankAudioSelectors = blankAudioKeys.length > 0;
    
    // 详细日志：子题 TTS 条件检查
    log(`  第 ${i + 1} 题 子题TTS检查: blanks数=${blanks.length}, 页面blank_audio选择器=${blankAudioKeys.join(',') || '无'}`);
    
    // 如果有 blanks 且页面有对应的 blank_audio_N 选择器，为每个子题单独合成
    // 注意：如果顶层已有 listening_script（情形B：共享对话），子题不应单独合成，除非子题自己有独立的 listening_script
    const topHasListeningScript = listeningScriptValue.length > 0;
    
    if (blanks.length > 0 && hasBlankAudioSelectors) {
      log(`  第 ${i + 1} 题：检测到 ${blanks.length} 个子题，检查是否需要为子题合成 TTS (顶层有原文=${topHasListeningScript})`);
      
      for (let bi = 0; bi < blanks.length; bi++) {
        const blank = blanks[bi];
        // 子题只用自己的 listening_script，不用 question 替代
        // 如果顶层有共享对话，子题没有独立原文时不应单独合成
        let blankScript = (blank.listening_script || blank.listeningScript || blank.script || "").trim();
        const blankQuestion = (blank.question || "").trim();
        
        // 调试：显示子题的所有相关字段
        log(`    子题 ${bi + 1} 数据: listening_script="${(blank.listening_script || '').slice(0, 30)}", question="${blankQuestion.slice(0, 30)}", audio_url="${blank.audio_url || '无'}"`);
        
        // 只有当顶层没有共享对话时，才用 question 替代（情形A：各小题独立）
        if (!blankScript && blankQuestion.length > 0 && !topHasListeningScript) {
          blankScript = blankQuestion;
          log(`    子题 ${bi + 1}: 顶层无共享原文，使用题干作为 TTS 文本`);
        }
        // 无听力原文时，用答案合成音频（如 listening_fill_and_retell 第2小题的参考答案）
        if (!blankScript && blank.answer) {
          let answerText = "";
          if (Array.isArray(blank.answer)) {
            answerText = blank.answer.filter(x => typeof x === "string").join(" ");
          } else if (typeof blank.answer === "string") {
            answerText = blank.answer.split("#")[0].trim();  // 取第一个参考答案
          }
          if (answerText.length > 30) {
            blankScript = answerText;
            log(`    子题 ${bi + 1}: 无听力原文，使用参考答案作为 TTS 文本 (${answerText.length}字)`);
          }
        }
        
        const blankAudioSel = curSel[`blank_audio_${bi + 1}`];
        
        // 检查子题是否已有音频（避免重复合成）
        const blankHasAudio = (blank.audio_url || "").trim().length > 0 || (blank.audio_base64 || "").length > 0;
        
        if (blankScript && blankAudioSel && !blankHasAudio) {
          log(`    子题 ${bi + 1}: 需要TTS - 有文本(${blankScript.length}字符)，有选择器(${blankAudioSel})，无已有音频`);
          
          // 为这个子题启动 TTS（带重试机制）
          const ttsPromiseForBlank = (async () => {
            // 处理文本：过滤音标，保留英文单词
            const ttsText = prepareTtsText(blankScript);
            if (ttsText !== blankScript) {
              log(`    子题 ${bi + 1}: 过滤音标后文本: "${ttsText.slice(0, 50)}..."`);
            }
            const isDialogue = /^[WwMmQqAa][：:]/m.test(ttsText);
            
            const audioBase64 = await ttsWithRetry({
              text: ttsText,
              dialogue: isDialogue,
              provider: ttsProvider,
              femaleVoice: ttsFemaleVoice,
              maleVoice: ttsMaleVoice,
              femaleSpeed: ttsFemaleSpeed,
              maleSpeed: ttsMaleSpeed,
              femaleVolume: ttsFemaleVolume,
              maleVolume: ttsMaleVolume,
              contextTexts: ttsContextTexts,
            }, 3, `    子题 ${bi + 1}: `, log);
            
            return { index: bi, audioBase64: audioBase64 || null, selector: blankAudioSel };
          })();
          
          blanksTtsPromises.push(ttsPromiseForBlank);
        } else {
          const skipReason = !blankScript ? '无TTS文本' : !blankAudioSel ? '无音频选择器' : blankHasAudio ? '已有音频' : '未知';
          log(`    子题 ${bi + 1}: 跳过TTS (原因: ${skipReason}, 文本长度=${blankScript?.length || 0}, sel=${blankAudioSel || 'null'}, hasAudio=${blankHasAudio})`);
        }
      }
      
    }

    // 通用处理：如果顶层没有 listening_script，也没有子题需要 TTS，但有 question（题干），就用题干来合成音频
    // 这样可以覆盖：模仿朗读、听后应答、以及其他任何有题干但没有听力原文的题型
    if (!listeningScriptValue && q.question && q.question.trim().length > 0 && blanksTtsPromises.length === 0) {
      listeningScriptValue = q.question.trim();
      log(`  第 ${i + 1} 题：无听力原文，使用题干作为 TTS 文本 (type=${q.type}): "${listeningScriptValue.slice(0, 50)}..."`);
    }
    // 无听力原文时，用答案合成音频（如 listening_fill_and_retell 等题型的参考答案）
    if (!listeningScriptValue && blanksTtsPromises.length === 0 && q.answer) {
      let answerText = "";
      if (Array.isArray(q.answer)) {
        answerText = q.answer.filter(x => typeof x === "string").join(" ");
      } else if (typeof q.answer === "string") {
        answerText = q.answer.split("#")[0].trim();
      }
      if (answerText.length > 30) {
        listeningScriptValue = answerText;
        log(`  第 ${i + 1} 题：无听力原文，使用参考答案作为 TTS 文本 (type=${q.type}, ${answerText.length}字)`);
      }
    }

    const hasScript = listeningScriptValue.length > 0;
    const hasAudio = (q.audio_url || "").trim().length > 0 || (q.audio_base64 || "").length > 0;
    const audioUrlVal = (q.audio_url || "").toString().trim() || (defaultAudioUrl || "").toString().trim();
    
    // 检查页面是否有音频输入框（URL输入或文件上传）
    const hasAudioInput = audioUrlSels.length > 0 || audioFileSels.length > 0;
    
    // 详细日志：显示每道题的 TTS 判断条件
    log(`第 ${i + 1} 题 TTS 判断: 顶层hasScript=${hasScript} (${listeningScriptValue.length}字符), hasAudio=${hasAudio}, hasAudioInput=${hasAudioInput}, 子题TTS数=${blanksTtsPromises.length}`);

    // 顶层 TTS：如果顶层有 listening_script 且没有音频，且页面有音频输入框，才合成顶层音频
    // 注意：这与子题 TTS 并行进行，不会阻塞
    if (hasScript && !hasAudio && hasAudioInput) {
      log(`  题目 ${i + 1} 顶层需要 TTS 合成，启动异步合成… 原文前50字: "${listeningScriptValue.slice(0, 50)}..."`);
      
      // 异步调用 TTS API（带重试机制）
      ttsPromise = (async () => {
        // 处理文本：过滤音标，保留英文单词
        const ttsText = prepareTtsText(listeningScriptValue);
        if (ttsText !== listeningScriptValue) {
          log(`  第 ${i + 1} 题：过滤音标后文本: "${ttsText.slice(0, 80)}..."`);
        }
        const isDialogue = /^[WwMmQqAa][：:]/m.test(ttsText);
        log(`  第 ${i + 1} 题：发送 TTS 请求 (${ttsProvider})，文本长度=${ttsText.length}，isDialogue=${isDialogue}`);
        
        return await ttsWithRetry({
          text: ttsText,
          dialogue: isDialogue,
          provider: ttsProvider,
          femaleVoice: ttsFemaleVoice,
          maleVoice: ttsMaleVoice,
          femaleSpeed: ttsFemaleSpeed,
          maleSpeed: ttsMaleSpeed,
          femaleVolume: ttsFemaleVolume,
          maleVolume: ttsMaleVolume,
          contextTexts: ttsContextTexts,
        }, 3, `  第 ${i + 1} 题：`, log);
      })();
    } else {
      // 不需要 TTS 的情况
      if (!hasScript) {
        log(`  第 ${i + 1} 题：无听力原文，跳过 TTS`);
      } else if (hasAudio) {
        log(`  第 ${i + 1} 题：已有音频，跳过 TTS`);
      } else if (!hasAudioInput) {
        log(`  第 ${i + 1} 题：页面无音频输入框，跳过 TTS`);
      }
    }

    // 显示统一的 TTS 状态（顶层 + 子题并行合成）
    const totalTtsCount = (ttsPromise ? 1 : 0) + blanksTtsPromises.length;
    if (totalTtsCount > 0) {
      const desc = totalTtsCount === 1 
        ? (ttsPromise ? '主音频' : '子题音频')
        : (ttsPromise ? `主音频 + ${blanksTtsPromises.length} 个子题` : `${blanksTtsPromises.length} 个子题`);
      showTtsStatus(`第 ${i + 1}/${questions.length} 题：${desc}合成中…`);
      log(`  第 ${i + 1} 题：共 ${totalTtsCount} 个音频并行合成中`);
    }

    // ── 音频填充（如果已有 audio_base64，即 JSON 中已包含音频）──
    let audioFileUploaded = false;
    
    log(`第 ${i + 1} 题音频状态: audioUrlVal=${audioUrlVal ? audioUrlVal.slice(0, 50) + "..." : "(空)"} hasBase64=${!!q.audio_base64} hasScript=${hasScript} audio_file选择器=${audioFileSels.length} audio_url选择器=${audioUrlSels.length}`);
    
    // 有 base64 音频数据时，尝试通过 uploadify 上传（这是 JSON 中已有的音频，不是 TTS 合成的）
    if (q.audio_base64) {
      log(`  第 ${i + 1} 题：检测到已有 audio_base64（${q.audio_base64.length} 字符），尝试 uploadify 上传`);
      
      // 找到音频的 uploadify 上传框
      let audioFileInputSel = audioFileSels[0] || null;
      // 如果没找到，尝试从 audioFileName 输入框推断
      if (!audioFileInputSel && audioUrlSels.length > 0) {
        const textEl = document.querySelector(audioUrlSels[0]);
        if (textEl && textEl.id) {
          // 模式：audioFileName → audio_upload > input[type=file]
          const uploadContainer = document.getElementById("audio_upload1") || document.getElementById("audio_upload");
          if (uploadContainer) {
            const fi = uploadContainer.querySelector("input[type='file']");
            if (fi) audioFileInputSel = fi.id ? `#${fi.id}` : "#audio_upload1 input[type='file']";
          }
        }
      }
      // 再尝试通用查找
      if (!audioFileInputSel) {
        const scope = document.querySelector(".topic-container, #topic-section, #right-container, .question-part") || document.body;
        const uploadContainer = scope.querySelector(
          "[id='audio_upload1'], [id='audio_upload'], [id*='audioUpload'], .audio-upload-wrap, " +
          "div[id*='audio_upload'], .uploadify[id*='audio']"
        );
        if (uploadContainer) {
          const fi = uploadContainer.querySelector("input[type='file']");
          if (fi) audioFileInputSel = fi.id ? `#${fi.id}` : null;
        }
      }
      log(`  第 ${i + 1} 题：音频上传框: ${audioFileInputSel || "未找到"}`);
      
      if (audioFileInputSel) {
        try {
          // 将 base64 转为 File
          const audioBlob = base64ToBlob(q.audio_base64, "audio/mpeg");
          if (audioBlob) {
            const audioFile = new File([audioBlob], `existing_q${i + 1}.mp3`, { type: "audio/mpeg" });
            log(`  第 ${i + 1} 题：音频File创建: name=${audioFile.name} size=${audioFile.size}`);
            
            const fileEl = document.querySelector(audioFileInputSel);
            if (fileEl) {
              let uploaded = false;
              const $ = window.jQuery || window.$;
              if ($ && typeof $.fn.uploadify === "function") {
                try {
                  const settings = $(fileEl).data("uploadify") || $(fileEl).closest("[class*='uploadify'],[id*='upload']").data("uploadify");
                  if (settings) {
                    log(`  音频找到uploadify settings，尝试直接上传`);
                    const xhr = new XMLHttpRequest();
                    const fd = new FormData();
                    const fileFieldName = settings.fileObjName || settings.fileField || "Filedata";
                    fd.append(fileFieldName, audioFile);
                    if (settings.formData) {
                      Object.entries(settings.formData).forEach(([k, v]) => fd.append(k, v));
                    }
                    const uploadUrl = settings.uploader || settings.uploadScript;
                    if (uploadUrl) {
                      await new Promise((resolve) => {
                        xhr.open("POST", uploadUrl);
                        xhr.withCredentials = true;
                        xhr.onload = () => {
                          log(`  音频uploadify直传 => ${xhr.status} ${xhr.responseText.slice(0, 100)}`);
                          try {
                            const resp = JSON.parse(xhr.responseText);
                            const fname = resp.file_path || resp.fileName || resp.filename || resp.name || resp.url || "";
                            if (fname && audioUrlSels.length > 0) {
                              const targetEl = document.querySelector(audioUrlSels[0]);
                              if (targetEl) {
                                const nativeDesc = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value");
                                if (nativeDesc && nativeDesc.set) {
                                  nativeDesc.set.call(targetEl, fname);
                                  targetEl.dispatchEvent(new Event("input", { bubbles: true }));
                                  targetEl.dispatchEvent(new Event("change", { bubbles: true }));
                                  log(`  音频audioFileName已更新: ${fname}`);
                                }
                              }
                            }
                          } catch (_) {}
                          resolve();
                        };
                        xhr.onerror = () => { log(`  音频uploadify直传网络错误`); resolve(); };
                        xhr.send(fd);
                      });
                      uploaded = true;
                      audioFileUploaded = true;
                      await delay(1500); // 音频文件较大，等待更长时间
                    }
                  }
                } catch (e2) {
                  log(`  音频uploadify直传异常: ${e2?.message}`);
                }
              }
              if (!uploaded) {
                const ok = fillFileField(audioFileInputSel, audioFile);
                log(`  音频fillFileField=${ok}`);
                if (ok) {
                  audioFileUploaded = true;
                  await delay(1500);
                }
              }
            }
          }
        } catch (e) {
          log(`  音频上传异常: ${e?.message}`);
        }
      }
    } else if (audioUrlVal && !ttsPromise) {
      // 没有 base64，也没有启动 TTS，但有 URL，尝试下载后上传
      const audioFile = await resolveMediaFile(q, "audio", defaultAudioUrl);
      if (audioFile) {
        for (const s of audioFileSels) {
          if (fillFileField(s, audioFile)) { audioFileUploaded = true; break; }
        }
        if (audioFileUploaded) await delay(1500);
      }
    }
    
    // 无论文件上传是否成功，URL 始终回填到文本框（#audioFileName 等）
    if (audioUrlVal && !audioFileUploaded && !ttsPromise) {
      for (const s of audioUrlSels) {
        const el = document.querySelector(s);
        if (!el) continue;
        if (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.isContentEditable) {
          fillField(s, audioUrlVal);
          el.dispatchEvent(new Event("blur", { bubbles: true }));
          await delay(300);
          break;
        }
      }
    }

    // ── 图片填充 ──────────────────────────────────────────────────────────
    // 若 JSON 未提供 image_url，用 defaultImageUrl 兜底
    const rawImageUrl = (q.image_url || "").toString().trim();
    const imageUrlVal = rawImageUrl || (defaultImageUrl || "").toString().trim();
    const hasImageBase64 = !!q.image_base64;
    const isLocalImageUrl = rawImageUrl.startsWith("http://localhost") && rawImageUrl.includes("/api/images/");
    let imageFileSels = qSel(curSel, "image_file");
    let imageUrlSels = qSel(curSel, "image_url");
    
    // 如果没有检测到图片选择器，尝试在当前题目区域内查找
    if (imageFileSels.length === 0 && imageUrlSels.length === 0) {
      const scope = document.querySelector(".topic-container, #topic-section, #right-container") || document.body;
      // 尝试找图片 URL 输入框
      const imgUrlInput = scope.querySelector("#imageFileName, input.imageFileName, input[id*='imageFile'], input[placeholder*='图片'], input[placeholder*='jpg'], input[placeholder*='png']");
      if (imgUrlInput) {
        const sel = imgUrlInput.id ? `#${imgUrlInput.id}` : (imgUrlInput.name ? `input[name="${imgUrlInput.name}"]` : null);
        if (sel) imageUrlSels = [sel];
      }
      // 尝试找图片文件上传框
      const imgFileInput = scope.querySelector("input[type='file'][accept*='image'], input[type='file'].image-upload");
      if (imgFileInput) {
        const sel = imgFileInput.id ? `#${imgFileInput.id}` : null;
        if (sel) imageFileSels = [sel];
      }
    }
    
    log(`图片填充: imageUrlVal=${imageUrlVal ? imageUrlVal.slice(0, 50) + "..." : "(空)"} isLocalUrl=${isLocalImageUrl} hasBase64=${hasImageBase64} image_file选择器=${imageFileSels.length} image_url选择器=${imageUrlSels.length}`);
    
    // ── 题干图片上传（localhost URL 需要通过 uploadify 上传）──
    if (isLocalImageUrl) {
      log(`  检测到本地题干图片URL，尝试uploadify上传: ${rawImageUrl.slice(0, 60)}`);
      // 找到题干图片的 uploadify 上传框
      let stemFileInputSel = imageFileSels[0] || null;
      // 如果没找到，尝试从 imageFileName 输入框推断
      if (!stemFileInputSel && imageUrlSels.length > 0) {
        const textEl = document.querySelector(imageUrlSels[0]);
        if (textEl && textEl.id) {
          // 模式：imageFileName → image_upload > input[type=file]
          const uploadContainer = document.getElementById("image_upload");
          if (uploadContainer) {
            const fi = uploadContainer.querySelector("input[type='file']");
            if (fi) stemFileInputSel = fi.id ? `#${fi.id}` : "#image_upload input[type='file']";
          }
        }
      }
      // 再尝试通用查找
      if (!stemFileInputSel) {
        const scope = document.querySelector(".topic-container, #topic-section, #right-container") || document.body;
        const uploadContainer = scope.querySelector("[id='image_upload'], [id*='imageUpload'], .image-upload-wrap");
        if (uploadContainer) {
          const fi = uploadContainer.querySelector("input[type='file']");
          if (fi) stemFileInputSel = fi.id ? `#${fi.id}` : null;
        }
      }
      log(`  题干图片上传框: ${stemFileInputSel || "未找到"}`);
      
      if (stemFileInputSel) {
        try {
          const res = await fetch(rawImageUrl);
          if (!res.ok) { log(`  题干图片fetch失败: ${res.status}`); }
          else {
            const blob = await res.blob();
            const ext = (blob.type || "image/jpeg").split("/")[1] || "jpg";
            const file = new File([blob], `stem_image.${ext}`, { type: blob.type || "image/jpeg" });
            log(`  题干图片File创建: name=${file.name} size=${file.size} type=${file.type}`);
            
            const fileEl = document.querySelector(stemFileInputSel);
            if (fileEl) {
              let uploaded = false;
              const $ = window.jQuery || window.$;
              if ($ && typeof $.fn.uploadify === "function") {
                try {
                  const settings = $(fileEl).data("uploadify") || $(fileEl).closest("[class*='uploadify'],[id*='upload']").data("uploadify");
                  if (settings) {
                    log(`  题干图片找到uploadify settings，尝试直接上传`);
                    const xhr = new XMLHttpRequest();
                    const fd = new FormData();
                    const fileFieldName = settings.fileObjName || settings.fileField || "Filedata";
                    fd.append(fileFieldName, file);
                    if (settings.formData) {
                      Object.entries(settings.formData).forEach(([k, v]) => fd.append(k, v));
                    }
                    const uploadUrl = settings.uploader || settings.uploadScript;
                    if (uploadUrl) {
                      await new Promise((resolve) => {
                        xhr.open("POST", uploadUrl);
                        xhr.withCredentials = true;
                        xhr.onload = () => {
                          log(`  题干图片uploadify直传 => ${xhr.status} ${xhr.responseText.slice(0, 100)}`);
                          try {
                            const resp = JSON.parse(xhr.responseText);
                            const fname = resp.file_path || resp.fileName || resp.filename || resp.name || resp.url || "";
                            if (fname && imageUrlSels.length > 0) {
                              const targetEl = document.querySelector(imageUrlSels[0]);
                              if (targetEl) {
                                const nativeDesc = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value");
                                if (nativeDesc && nativeDesc.set) {
                                  nativeDesc.set.call(targetEl, fname);
                                  targetEl.dispatchEvent(new Event("input", { bubbles: true }));
                                  targetEl.dispatchEvent(new Event("change", { bubbles: true }));
                                  log(`  题干图片imageFileName已更新: ${fname}`);
                                }
                              }
                            }
                          } catch (_) {}
                          resolve();
                        };
                        xhr.onerror = () => { log(`  题干图片uploadify直传网络错误`); resolve(); };
                        xhr.send(fd);
                      });
                      uploaded = true;
                      await delay(500);
                    }
                  }
                } catch (e2) {
                  log(`  题干图片uploadify直传异常: ${e2?.message}`);
                }
              }
              if (!uploaded) {
                const ok = fillFileField(stemFileInputSel, file);
                log(`  题干图片fillFileField=${ok}`);
                if (ok) await delay(1500);
              }
            }
          }
        } catch (e) {
          log(`  题干图片上传异常: ${e?.message}`);
        }
      }
    } else if (imageUrlVal || hasImageBase64) {
      // 非 localhost URL，走原有逻辑
      const imageFile = await resolveMediaFile(q, "image", defaultImageUrl);
      log(`  resolveMediaFile => ${imageFile ? "File对象" : "null"}`);
      if (imageFile && imageFileSels.length > 0) {
        for (const s of imageFileSels) {
          const ok = fillFileField(s, imageFile);
          log(`  fill image_file[${s.slice(0, 40)}] => ${ok ? "ok" : "fail"}`);
          if (ok) break;
        }
        await delay(200);
      }
      if (imageUrlVal && imageUrlSels.length > 0) {
        for (const s of imageUrlSels) {
          const el = document.querySelector(s);
          if (!el) { log(`  image_url 元素不存在: ${s}`); continue; }
          if (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.isContentEditable) {
            // 如果是 readonly 输入框（如 uploadify 组件的显示框），先移除 readonly 再填值
            const wasReadonly = el.hasAttribute("readonly");
            if (wasReadonly) {
              el.removeAttribute("readonly");
              log(`  移除 readonly 属性`);
            }
            // 使用原生 setter 设置值
            const proto = el.tagName === "INPUT" ? HTMLInputElement.prototype : HTMLTextAreaElement.prototype;
            const desc = Object.getOwnPropertyDescriptor(proto, "value");
            if (desc && desc.set) {
              desc.set.call(el, imageUrlVal);
            } else {
              el.value = imageUrlVal;
            }
            el.dispatchEvent(new Event("input", { bubbles: true }));
            el.dispatchEvent(new Event("change", { bubbles: true }));
            el.dispatchEvent(new Event("blur", { bubbles: true }));
            // 恢复 readonly（如果原来有的话）
            if (wasReadonly) {
              el.setAttribute("readonly", "readonly");
            }
            log(`  fill image_url => ok value="${imageUrlVal.slice(0, 50)}..."`);
            await delay(150);
            break;
          } else {
            log(`  image_url 元素不是可填充类型: tag=${el.tagName}`);
          }
        }
      }
      // 如果有图片数据但没有找到输入框，记录警告
      if (imageFileSels.length === 0 && imageUrlSels.length === 0) {
        log(`  警告: 有图片数据但未找到图片输入框`);
      }
    } else {
      log(`  跳过图片填充（无数据，defaultImageUrl也为空）`);
    }

    // ── 选项图片上传（Word 图片提取功能）────────────────────────────────────
    // 若 options 里包含来自本地后端 /api/images/ 的 URL，则尝试通过文件上传框上传
    // 同时也会通过下面的文字填充把 URL 写入 imageFileName 输入框（双保险）
    // 收集所有选项URL：顶层 q.options 及各 blank.options（多小题场景）
    const _topOptArr = Array.isArray(q.options || q.option) ? (q.options || q.option) : [];
    const _blanksOptArr = (q.blanks || []).flatMap(b => Array.isArray(b?.options) ? b.options : []);
    const optionLocalUrls = [..._topOptArr, ..._blanksOptArr];
    const hasLocalOptionImages = optionLocalUrls.some(
      (v) => typeof v === "string" && v.includes("/api/images/") && v.startsWith("http://localhost")
    );
    if (hasLocalOptionImages) {
      const _OPT_ROLES = ["option_a", "option_b", "option_c", "option_d"];

      // 从 imageFileName 文本输入框的选择器推断对应的文件上传框选择器
      const deriveFileInputSel = (textSel) => {
        if (!textSel) return null;
        const textEl = document.querySelector(textSel);
        if (!textEl || !textEl.id) return null;
        // 模式：imageFileName4 → image_upload4 > input[type=file]
        const numMatch = textEl.id.match(/(\d+)$/);
        if (numMatch) {
          const n = numMatch[1];
          const uploadContainer = document.getElementById(`image_upload${n}`);
          if (uploadContainer) {
            const fi = uploadContainer.querySelector("input[type='file']");
            if (fi) return fi.id ? `#${fi.id}` : `#image_upload${n} input[type='file']`;
          }
        }
        // 备用：在同一父容器内查找 input[type=file]
        const container = textEl.closest(".row, .upload-wrap, .col, .option-upload") || textEl.parentElement;
        if (container) {
          const fi = container.querySelector("input[type='file']");
          if (fi && fi.id) return `#${fi.id}`;
        }
        return null;
      };

      const fetchAndUpload = async (url, fileInputSel, label) => {
        if (!fileInputSel) return;
        try {
          const res = await fetch(url);
          if (!res.ok) { log(`  选项${label} fetch失败: ${res.status}`); return; }
          const blob = await res.blob();
          const ext = (blob.type || "image/jpeg").split("/")[1] || "jpg";
          const file = new File([blob], `option_${label}.${ext}`, { type: blob.type || "image/jpeg" });
          log(`  选项${label} File创建: name=${file.name} size=${file.size} type=${file.type}`);

          const fileEl = document.querySelector(fileInputSel);
          if (!fileEl) { log(`  选项${label} 找不到file input`); return; }

          // ── 方式1：jQuery uploadify API（最可靠，直接写入内部队列并上传）
          let uploaded = false;
          const $ = window.jQuery || window.$;
          if ($ && typeof $.fn.uploadify === "function") {
            try {
              // 把 File 注入 uploadify 内部队列
              const settings = $(fileEl).data("uploadify") || $(fileEl).closest("[class*='uploadify'],[id*='upload']").data("uploadify");
              if (settings) {
                log(`  选项${label} 找到uploadify settings，尝试直接上传`);
                // 构造 uploadify queue item 并调用其上传逻辑
                const xhr = new XMLHttpRequest();
                const fd = new FormData();
                const fileFieldName = settings.fileObjName || settings.fileField || "Filedata";
                fd.append(fileFieldName, file);
                // 附加 uploadify 的额外 formData
                if (settings.formData) {
                  Object.entries(settings.formData).forEach(([k, v]) => fd.append(k, v));
                }
                const uploadUrl = settings.uploader || settings.uploadScript;
                if (uploadUrl) {
                  await new Promise((resolve) => {
                    xhr.open("POST", uploadUrl);
                    xhr.withCredentials = true;
                    xhr.onload = () => {
                      log(`  选项${label} uploadify直传 => ${xhr.status} ${xhr.responseText.slice(0, 100)}`);
                      // uploadify 回调：把返回的文件名写入 imageFileName input
                      try {
                        const resp = JSON.parse(xhr.responseText);
                        const fname = resp.file_path || resp.fileName || resp.filename || resp.name || resp.url || "";
                        if (fname) {
                          const textSelIdx = _OPT_ROLES.indexOf(
                            _uploadBlanks.length > 0
                              ? _OPT_ROLES.find((_, oi2) => String.fromCharCode(65 + oi2) === label[0])
                              : _OPT_ROLES[label.charCodeAt(0) - 65]
                          );
                          // 尝试更新对应 imageFileName 输入框
                          const allOptSels = _OPT_ROLES.map(r => qSel(curSel, r)[0]);
                          const matchSel = allOptSels.find(s => {
                            const el2 = s && document.querySelector(s);
                            return el2 && el2.closest("[id*='upload']") === fileEl.closest("[id*='upload']");
                          });
                          if (matchSel) {
                            const nativeDesc = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value");
                            const targetEl = document.querySelector(matchSel);
                            if (targetEl && nativeDesc && nativeDesc.set) {
                              nativeDesc.set.call(targetEl, fname);
                              targetEl.dispatchEvent(new Event("input", { bubbles: true }));
                              targetEl.dispatchEvent(new Event("change", { bubbles: true }));
                              log(`  选项${label} imageFileName 已更新: ${fname}`);
                            }
                          }
                        }
                      } catch (_) {}
                      resolve();
                    };
                    xhr.onerror = () => { log(`  选项${label} uploadify直传网络错误`); resolve(); };
                    xhr.send(fd);
                  });
                  uploaded = true;
                  await delay(300);
                }
              }
            } catch (e2) {
              log(`  选项${label} uploadify直传异常: ${e2?.message}`);
            }
          }

          // ── 方式2：fillFileField 兜底（验证 el.files 是否正确设置）
          if (!uploaded) {
            const ok = fillFileField(fileInputSel, file);
            // 验证 el.files 实际上是否被设置（某些 uploadify 会清空）
            const actualSize = fileEl.files?.[0]?.size ?? -1;
            log(`  选项${label} fillFileField=${ok} el.files[0].size=${actualSize}(期望${file.size})`);
            if (ok && actualSize === file.size) {
              await delay(1500);
            } else {
              log(`  选项${label} ⚠ el.files未正确设置，uploadify可能不支持程序化注入`);
            }
          }
        } catch (e) {
          log(`  选项${label} 图片上传异常: ${e?.message}`);
        }
      };

      const _uploadBlanks = q.blanks || [];
      if (_uploadBlanks.length > 0) {
        // 多小题：每个 blank 各自有 A/B/C 选项，对应不同的 DOM 上传框
        for (let oi = 0; oi < _OPT_ROLES.length; oi++) {
          const label = String.fromCharCode(65 + oi);
          const textSelsForRole = qSel(curSel, _OPT_ROLES[oi]);
          for (let j = 0; j < _uploadBlanks.length; j++) {
            const b = _uploadBlanks[j];
            const optUrl = typeof (b?.options?.[oi]) === "string" ? b.options[oi].trim() : "";
            if (!optUrl.startsWith("http://localhost") || !optUrl.includes("/api/images/")) continue;
            const fileInputSel = deriveFileInputSel(textSelsForRole[j]);
            log(`  blank[${j}] 选项${label} 图片URL=${optUrl.slice(0, 60)} fileInputSel=${fileInputSel || "未找到"}`);
            await fetchAndUpload(optUrl, fileInputSel, label + (j + 1));
          }
        }
      } else {
        // 单题顶层 options
        const topOpts = Array.isArray(q.options || q.option) ? (q.options || q.option) : [];
        for (let oi = 0; oi < topOpts.length; oi++) {
          const optUrl = typeof topOpts[oi] === "string" ? topOpts[oi].trim() : "";
          if (!optUrl.startsWith("http://localhost") || !optUrl.includes("/api/images/")) continue;
          const label = String.fromCharCode(65 + oi);
          const fileInputSel = deriveFileInputSel(qSel(curSel, _OPT_ROLES[oi])[0]);
          log(`  顶层选项${label} 图片URL=${optUrl.slice(0, 60)} fileInputSel=${fileInputSel || "未找到"}`);
          await fetchAndUpload(optUrl, fileInputSel, label);
        }
      }
    }

    /** 去除选项文本头部的字母前缀，如 "A. Tim's." → "Tim's."，"(B) Dave's." → "Dave's." */
    const stripOptPrefix = (s) => {
      if (!s) return s;
      return String(s).trim().replace(/^[（(]?[A-Da-d][)）.\s、。，]+/, "").trim();
    };
    /** 从题目 q 归一化出 options 数组（支持 options/option 及 { A: "x", B: "y" }），并去除字母前缀 */
    const getOptionsArray = (q) => {
      let arr = q.options || q.option;
      if (Array.isArray(arr)) return arr.map(stripOptPrefix);
      if (arr && typeof arr === "object" && !Array.isArray(arr))
        return ["A", "B", "C", "D"].map((k) => (arr[k] != null ? stripOptPrefix(String(arr[k]).trim()) : "")).filter(Boolean);
      return [];
    };
    /**
     * 将 AI 返回的值中的 <<IMG>> 占位符替换为高级设置里配置的默认图片 URL/文件名。
     * 用于图片类型选项：AI 对图片选项统一输出 "<<IMG>>"，填充时替换为实际文件名。
     * 同时过滤掉 localhost 图片 URL（这类值由 uploadify 上传回调写入，不应覆写进文本框）。
     */
    const resolveImgPlaceholder = (val) => {
      if (typeof val === "string" && val.trim() === "<<IMG>>") {
        return (defaultImageUrl || "").toString().trim();
      }
      // localhost 图片 URL 不写进 imageFileName 文本框，
      // 让 uploadify 异步上传完成后自己回填平台文件名；
      // 若上传失败，fallback 到 defaultImageUrl（与无图时行为一致）。
      if (typeof val === "string" && val.startsWith("http://localhost") && val.includes("/api/images/")) {
        return (defaultImageUrl || "").toString().trim();
      }
      return val;
    };

    /**
     * 将英文内容中的全角标点替换为半角 ASCII 等价字符。
     * 若字符串含有汉字（CJK），视为中文内容，不做替换（中文标点在中文语境下是正确的）。
     * 适用于 answer、listening_script、keyword 等可能夹带全角标点的字段。
     */
    const sanitizeAnswerChars = (s) => {
      if (!s) return s;
      // 含有汉字则跳过（中文内容保留原格式）
      if (/[\u4e00-\u9fff\u3400-\u4dbf\u{20000}-\u{2a6df}]/u.test(s)) return s;
      return s
        .replace(/\uff0c/g, ",")   // ，→ ,
        .replace(/\u3002/g, ".")   // 。→ .
        .replace(/\uff1f/g, "?")   // ？→ ?
        .replace(/\uff01/g, "!")   // ！→ !
        .replace(/\uff1a/g, ":")   // ：→ :
        .replace(/\uff1b/g, ";")   // ；→ ;
        .replace(/\u2018|\u2019/g, "'")  // '' → '
        .replace(/\u201c|\u201d/g, '"')  // "" → "
        .replace(/\uff08/g, "(")   // （→ (
        .replace(/\uff09/g, ")")   // ）→ )
        .replace(/\u3010/g, "[")   // 【→ [
        .replace(/\u3011/g, "]")   // 】→ ]
        .replace(/\u2026/g, "...") // …→ ...
        .replace(/\u2014/g, "-")   // —→ -
        .replace(/\uff5e/g, "~");  // ～→ ~
    };

    /** 根据「输入框用途」(role) 从题目 q 里取对应字段的值，支持 blank_*_N 动态角色 */
    const getValueForRole = (role) => {
      const opts = getOptionsArray(q);
      // blank_answer_1 / blank_question_2 / blank_keyword_2 / blank_audio_2 等动态角色
      const blankMatch = role.match(/^blank_(answer|question|keyword|audio|script)_(\d+)$/);
      if (blankMatch) {
        const field = blankMatch[1]; // answer / question / keyword / audio / script
        const idx = parseInt(blankMatch[2], 10) - 1;
        const blanks = q.blanks || [];
        const blank = blanks[idx];
        if (!blank) {
          // script 字段兜底到大题顶层听力原文（子题无独立原文时共享大题原文）
          if (field === "script") return sanitizeAnswerChars((q.listening_script || "").toString().trim());
          // audio 字段兜底到默认音频
          if (field === "audio") return (defaultAudioUrl || "").toString().trim();
          return "";
        }
        if (field === "answer") {
          if (Array.isArray(blank.answer)) {
            // 从答案框的 placeholder 动态读取分隔符（如「用'#'分隔」），兜底用 #
            let sep = "#";
            try {
              const answerSel = curSel[role] || curSel[`blank_answer_${idx + 1}`];
              const answerEl = answerSel && document.querySelector(answerSel);
              const ph = answerEl && (answerEl.placeholder || answerEl.getAttribute("placeholder") || "");
              const sepMatch = ph.match(/用['''"](.+?)['''"]分[隔开]/);
              if (sepMatch) sep = sepMatch[1];
            } catch (_) {}
            return sanitizeAnswerChars(blank.answer.map(a => String(a).trim()).join(sep));
          }
          return sanitizeAnswerChars((blank.answer != null ? blank.answer : "").toString().trim());
        }
        if (field === "question") return sanitizeAnswerChars((blank.question != null ? blank.question : "").toString().trim());
        if (field === "keyword") return sanitizeAnswerChars((blank.keyword != null ? blank.keyword : "").toString().trim());
        if (field === "audio") return (blank.audio_url || "").toString().trim() || (defaultAudioUrl || "").toString().trim();
        if (field === "script") return sanitizeAnswerChars(
          (blank.listening_script || blank.script || "").toString().trim()
          || (q.listening_script || "").toString().trim() // 子题没有独立原文时，回退到大题的听力原文
        );
        return "";
      }
      // 听后应答题特殊处理：question→听力原文，candidates→题干
      if (q.type === "listening_response") {
        if (role === "listening_script") {
          // 优先使用 AI 明确提取出的听力原文；仅在缺失时才回退到 question
          return sanitizeAnswerChars((q.listening_script || q.question || "").toString().trim());
        }
        if (role === "question") {
          // 设置题干：填入候选项（换行分隔）
          const candidates = q.candidates || q.options || [];
          if (Array.isArray(candidates) && candidates.length > 0) {
            return candidates.join("\n");
          }
          return "";
        }
      }
      switch (role) {
        case "listening_script": return sanitizeAnswerChars((q.listening_script != null ? q.listening_script : "").toString().trim());
        case "explanation": return sanitizeAnswerChars((q.explanation || "").toString().trim());
        case "question": return sanitizeAnswerChars((q.question || "").toString().trim());
        case "keyword": return sanitizeAnswerChars((q.keyword || "").toString().trim());
        case "option_a": return sanitizeAnswerChars(resolveImgPlaceholder((opts[0] != null ? opts[0] : "").toString().trim()));
        case "option_b": return sanitizeAnswerChars(resolveImgPlaceholder((opts[1] != null ? opts[1] : "").toString().trim()));
        case "option_c": return sanitizeAnswerChars(resolveImgPlaceholder((opts[2] != null ? opts[2] : "").toString().trim()));
        case "option_d": return sanitizeAnswerChars(resolveImgPlaceholder((opts[3] != null ? opts[3] : "").toString().trim()));
        case "answer": return sanitizeAnswerChars((q.answer != null ? q.answer : "").toString().trim());
        // 也支持 JSON 里直接写了任意 key 的情况（AI 返回的动态字段）
        default: return sanitizeAnswerChars(q[role] != null ? String(q[role]).trim() : "");
      }
    };
    // FILLABLE_ROLES 动态生成：取 curSel 中所有检测到的非属性字段
    const SKIP_ROLES = new Set(["grade", "course", "unit", "knowledge_point", "difficulty", "question_permission", "recorder", "audio_file", "image_file", "audio_url", "image_url", "submit_btn", "next_btn"]);
    const FILLABLE_ROLES = Object.keys(curSel).filter(r => curSel[r] && !SKIP_ROLES.has(r));
    // 确保核心字段按合理顺序：一题多小题时 blank_script_1/2 紧跟 listening_script，blank_question_1/2 紧跟 question，blank_answer_1/2 紧跟 answer
    const CORE_ORDER = ["listening_script", "question", "option_a", "option_b", "option_c", "option_d", "answer", "explanation"];
    const getOrder = (r) => {
      const i = CORE_ORDER.indexOf(r);
      if (i >= 0) return i;
      // blank_script_N（小题听力原文）紧跟 listening_script
      const sMatch = r.match(/^blank_script_(\d+)$/);
      if (sMatch) return 0 + parseInt(sMatch[1], 10) / 100;
      const qMatch = r.match(/^blank_question_(\d+)$/);
      if (qMatch) return 1 + parseInt(qMatch[1], 10) / 100;
      const aMatch = r.match(/^blank_answer_(\d+)$/);
      if (aMatch) return 6 + parseInt(aMatch[1], 10) / 100;
      return 10 + (r.localeCompare ? r.localeCompare("") : 0);
    };
    FILLABLE_ROLES.sort((a, b) => {
      const oa = getOrder(a), ob = getOrder(b);
      return oa !== ob ? oa - ob : String(a).localeCompare(String(b));
    });
    // 题目属性（年级/课程/知识点等）用户手动设置一次后平台会保留，不自动填充
    const PROP_ROLES = [];
    log(`FILLABLE_ROLES: ${FILLABLE_ROLES.join(",")}`);
    for (const role of FILLABLE_ROLES) {
      const selectorsForRole = qSel(curSel, role);
      log(`role=${role} sels=${selectorsForRole.length} sel0="${selectorsForRole[0]?.slice(0,50)}"`);
      if (selectorsForRole.length === 0) continue;
      const blanks = q.blanks || [];
      const hasBlanks = blanks.length > 0;
      // 听力原文、题干、第N小题听力原文、第N小题题干、第N小题答案：先点击展开 UEditor 再填（一题多小题的题干在下面按 blanks 单独处理）
      const isQuestionLike = role === "listening_script" || role === "question" || /^blank_(question|script)_\d+$/.test(role);
      const isBlankAnswer = /^blank_answer_\d+$/.test(role);
      // blank_script_N / blank_audio_N 使用了 partSel 前缀选择器，Vue/React 重渲染后需重新打标
      if (/^blank_(script|audio)_\d+$/.test(role)) {
        try { document.querySelectorAll(".question-part").forEach((pt, pi) => pt.setAttribute("data-fill-part-idx", String(pi))); } catch (_) {}
      }
      if ((isQuestionLike || isBlankAnswer) && !(role === "question" && hasBlanks && selectorsForRole.length >= blanks.length)) {
        const firstSel = selectorsForRole[0];
        if (firstSel) {
          try {
            const fieldEl = document.querySelector(firstSel);
            if (fieldEl) {
              const row = fieldEl.closest && fieldEl.closest(".row");
              const ueditorShow = row && row.querySelector(".ueditor-show");
              if (ueditorShow && ueditorShow.click) {
                ueditorShow.click();
                await delay(isBlankAnswer ? 350 : 450);
              } else {
                const toClick = fieldEl.previousElementSibling || fieldEl.parentElement || fieldEl;
                if (toClick && toClick.click) toClick.click();
                fieldEl.focus();
                await delay(isBlankAnswer ? 300 : 400);
              }
            }
          } catch (_) {}
        }
      }
      // 一题多小题（blanks）：听力原文先填顶层，再逐小题填 blank_script_N，与题干/答案同逻辑，避免小题听力原文漏填
      if (role === "listening_script" && hasBlanks && blanks.length > 0) {
        try { document.querySelectorAll(".question-part").forEach((pt, pi) => pt.setAttribute("data-fill-part-idx", String(pi))); } catch (_) {}
        const topVal = (q.listening_script || "").toString().trim();
        if (topVal && selectorsForRole.length > 0) {
          const firstSel = selectorsForRole[0];
          if (firstSel) {
            try {
              const fieldEl = document.querySelector(firstSel);
              if (fieldEl) {
                const row = fieldEl.closest && fieldEl.closest(".row");
                const ueditorShow = row && row.querySelector(".ueditor-show");
                if (ueditorShow && ueditorShow.click) { ueditorShow.click(); await delay(450); }
                else { (fieldEl.previousElementSibling || fieldEl.parentElement || fieldEl).click?.(); fieldEl.focus(); await delay(400); }
              }
            } catch (_) {}
            fillField(firstSel, topVal);
            log(`  fill listening_script (顶层) => ok`);
            await delay(120);
          }
        }
        // 顶层填充后等待页面稳定，再重打 data-fill-part-idx 标记
        await delay(300);
        try { document.querySelectorAll(".question-part").forEach((pt, pi) => pt.setAttribute("data-fill-part-idx", String(pi))); } catch (_) {}
        await delay(100);
        for (let j = 0; j < blanks.length; j++) {
          const val = (blanks[j] && (blanks[j].listening_script != null || blanks[j].script != null) ? String(blanks[j].listening_script || blanks[j].script || "").trim() : "") || "";
          const sel = curSel[`blank_script_${j + 1}`];
          log(`  blank_script_${j + 1}: sel=${sel || "(null)"} val="${val.slice(0,30)}"`);
          if (!sel) { log(`  skip blank_script_${j + 1} (无选择器)`); continue; }
          if (!val) { log(`  skip blank_script_${j + 1} (值为空)`); continue; }
          // 每次填前重新打标
          try { document.querySelectorAll(".question-part").forEach((pt, pi) => pt.setAttribute("data-fill-part-idx", String(pi))); } catch (_) {}
          // 找目标元素：先用存储的选择器，失败时在第 j 个 .question-part 里按 placeholder 特征兜底
          let targetEl = null;
          try { targetEl = document.querySelector(sel); } catch (_) {}
          if (!targetEl) {
            try {
              const parts = document.querySelectorAll(".question-part");
              const part = parts[j];
              if (part) {
                targetEl = Array.from(part.querySelectorAll("textarea")).find(ta => {
                  const ph = (ta.getAttribute("placeholder") || "").toLowerCase();
                  return /听力|原文|报告|script/i.test(ph);
                }) || null;
                if (targetEl) log(`  blank_script_${j + 1} fallback via placeholder OK`);
              }
            } catch (_) {}
          }
          log(`  blank_script_${j + 1} element found: ${!!targetEl}`);
          if (targetEl) {
            try {
              if (targetEl.scrollIntoView) targetEl.scrollIntoView({ block: "nearest", behavior: "auto" });
              await delay(j > 0 ? 200 : 150);
              const row = targetEl.closest && targetEl.closest(".row");
              const ueditorShow = row && row.querySelector(".ueditor-show");
              if (ueditorShow && ueditorShow.click) { ueditorShow.click(); await delay(450); }
              else { (targetEl.previousElementSibling || targetEl.parentElement || targetEl).click?.(); targetEl.focus(); await delay(400); }
            } catch (_) {}
          }
          const fillSel = targetEl ? (targetEl.id ? `#${CSS.escape(targetEl.id)}` : sel) : sel;
          const ok = fillField(fillSel, val);
          log(`  fill blank_script_${j + 1} => ${ok ? "ok" : "fail"} len=${val.length}`);
          if (ok) await delay(120);
        }
        continue;
      }
      // 一题多小题（blanks）：题干按小题索引分别填入多个框
      if (role === "question" && hasBlanks && selectorsForRole.length >= blanks.length) {
        for (let j = 0; j < blanks.length; j++) {
          const val = (blanks[j] && blanks[j].question != null ? String(blanks[j].question).trim() : "") || "";
          if (!val) { log(`  skip question[${j}] (空)`); continue; }
          const sel = selectorsForRole[j];
          if (!sel) continue;
          try {
            const fieldEl = document.querySelector(sel);
            if (fieldEl) {
              const row = fieldEl.closest && fieldEl.closest(".row");
              const ueditorShow = row && row.querySelector(".ueditor-show");
              if (ueditorShow && ueditorShow.click) { ueditorShow.click(); await delay(400); }
              else { (fieldEl.previousElementSibling || fieldEl.parentElement || fieldEl).click?.(); fieldEl.focus(); await delay(350); }
            }
          } catch (_) {}
          const ok = fillField(sel, val);
          log(`  fill question[${j}] => ${ok ? "ok" : "fail"} len=${val.length}`);
          if (ok) await delay(80);
        }
        continue;
      }
      if (role === "option_a") await delay(120);
      if ((role === "option_b" || role === "option_c" || role === "option_d") && !hasBlanks) {
        const value = getValueForRole(role);
        if (!value) continue;
        const firstSel = selectorsForRole[0];
        if (firstSel) {
          try {
            const optionEl = document.querySelector(firstSel);
            if (optionEl) {
              const row = optionEl.closest && optionEl.closest(".row");
              const ueditorShow = row && row.querySelector(".ueditor-show");
              if (ueditorShow && ueditorShow.click) {
                ueditorShow.click();
                await delay(450);
              } else {
                const toClick = optionEl.previousElementSibling || optionEl.parentElement || optionEl;
                if (toClick && toClick.click) toClick.click();
                optionEl.focus();
                await delay(400);
              }
            }
          } catch (_) {}
        }
      }
      // 一题多小题（blanks）：选项 A/B/C/D 按小题索引分别填入。多小题选项多为普通 input，题干/其他题为富文本，需区分
      if (hasBlanks && selectorsForRole.length >= blanks.length && /^option_(a|b|c|d)$/.test(role)) {
        const optIdx = { option_a: 0, option_b: 1, option_c: 2, option_d: 3 }[role];
        log(`填小题选项 role=${role} optIdx=${optIdx} blanks=${blanks.length} sels=${selectorsForRole.length}`);
        for (let j = 0; j < blanks.length; j++) {
          const b = blanks[j];
          // 支持两种结构：blanks[j].options（每小题独立）或 q.options[j]（顶层二维数组）
          const opts = Array.isArray(b?.options) ? b.options
            : (b?.option != null ? (Array.isArray(b.option) ? b.option : [b.option_a, b.option_b, b.option_c, b.option_d].filter(Boolean))
            : (Array.isArray(q.options?.[j]) ? q.options[j] : []));
          const val = resolveImgPlaceholder(stripOptPrefix((opts[optIdx] != null ? String(opts[optIdx]).trim() : "") || ""));
          const sel = selectorsForRole[j];
          log(`  j=${j} val="${val}" sel="${sel}"`);
          if (!val) { log(`  j=${j} 跳过：val为空`); continue; }
          if (!sel) { log(`  j=${j} 跳过：sel为空`); continue; }
          // Vue/React 重渲染会清除注入的 data-fill-part-idx，每次 fill 前重新应用
          try {
            document.querySelectorAll(".question-part").forEach((pt, pi) => pt.setAttribute("data-fill-part-idx", String(pi)));
          } catch (_) {}
          try {
            const optionEl = document.querySelector(sel);
            log(`  j=${j} querySelector结果: ${optionEl ? optionEl.tagName + ' name=' + optionEl.name : 'null'}`);
            if (optionEl) {
              if (optionEl.scrollIntoView) optionEl.scrollIntoView({ block: "nearest", behavior: "auto" });
              await delay(j > 0 ? 120 : 80);
              const isPlainInput = optionEl.tagName === "INPUT" || optionEl.tagName === "TEXTAREA";
              if (isPlainInput) {
                optionEl.focus();
                await delay(80);
              } else {
                const row = optionEl.closest && optionEl.closest(".row");
                const ueditorShow = row && row.querySelector(".ueditor-show");
                if (ueditorShow && ueditorShow.click) {
                  ueditorShow.click();
                  await delay(j > 0 ? 500 : 400);
                } else {
                  const toClick = optionEl.previousElementSibling || optionEl.parentElement || optionEl;
                  if (toClick && toClick.click) toClick.click();
                  optionEl.focus();
                  await delay(j > 0 ? 450 : 350);
                }
              }
            }
          } catch (e) { log(`  j=${j} 异常: ${e?.message}`); }
          const filled = fillField(sel, val);
          log(`  j=${j} fillField结果: ${filled}`);
          if (filled) await delay(100);
        }
        continue;
      }
      if (role === "explanation") {
        const firstSel = selectorsForRole[0];
        if (firstSel) {
          try {
            const explEl = document.querySelector(firstSel);
            if (explEl) {
              const row = explEl.closest && explEl.closest(".row");
              const ueditorShow = row && row.querySelector(".ueditor-show");
              if (ueditorShow && ueditorShow.click) {
                ueditorShow.click();
                await delay(400);
              } else {
                const toClick = explEl.previousElementSibling || explEl.parentElement || explEl;
                if (toClick && toClick.click) toClick.click();
                explEl.focus();
                await delay(300);
              }
            }
          } catch (_) {}
        }
      }
      if (role === "answer" && ( (q.blanks && q.blanks.length > 0) || (q.answers && q.answers.length > 0) )) {
        const blankAnswerToStr = (b, j) => {
          if (!b || b.answer == null) return "";
          if (Array.isArray(b.answer)) {
            // 从对应答案框的 placeholder 动态读取分隔符，兜底用 #
            let sep = "#";
            try {
              const sel = selectorsForRole[j] || (curSel[`blank_answer_${j + 1}`] || "").split(",")[0];
              const el = sel && document.querySelector(sel);
              const ph = el && (el.placeholder || el.getAttribute("placeholder") || "");
              const m = ph.match(/用['''"](.+?)['''"]分[隔开]/);
              if (m) sep = m[1];
            } catch (_) {}
            return b.answer.map(a => String(a).trim()).join(sep);
          }
          return String(b.answer).trim();
        };
        const values = (q.blanks?.length ? q.blanks.map((b, j) => blankAnswerToStr(b, j)) : q.answers.map((v) => (v != null ? String(v).trim() : ""))).map(sanitizeAnswerChars);
        for (let j = 0; j < values.length; j++) {
          const sel = selectorsForRole[j];
          if (!sel || !values[j]) continue;
          const ok = fillField(sel, values[j]);
          log(`  fill answer[${j}] => ${ok ? "ok" : "fail"} value="${String(values[j]).slice(0, 40)}${String(values[j]).length > 40 ? "…" : ""}"`);
          if (ok) await delay(80);
        }
      } else if (role === "answer") {
        const value = sanitizeAnswerChars(getValueForRole(role));
        if (!value) { log(`  skip answer (无值)`); continue; }
        // 优先路径：selector 是容器(div.ui-radio)，fillField 内部处理 label 点击
        const firstEl = (() => { try { return document.querySelector(selectorsForRole[0]); } catch(_) { return null; } })();
        if (firstEl && firstEl.tagName === "DIV") {
          const ok = fillField(selectorsForRole[0], value);
          log(`  fill answer => ${ok ? "ok" : "fail"} value="${value}"`);
          await delay(80);
        } else {
          // selector 是单个 radio：按字母顺序(A=0,B=1,C=2,D=3)找对应索引的 radio，
          // 再通过 label[for=radio.id] 点 label（更可靠，不依赖 radio.value 的具体内容）
          const v = value.trim().toUpperCase();
          const abcd = ["A", "B", "C", "D"];
          const targetIdx = abcd.indexOf(v) >= 0 ? abcd.indexOf(v) : (parseInt(v, 10) >= 1 && parseInt(v, 10) <= 4 ? parseInt(v, 10) - 1 : -1);
          let filled = false;
          if (targetIdx >= 0 && targetIdx < selectorsForRole.length) {
            const sel = selectorsForRole[targetIdx];
            try {
              const radio = document.querySelector(sel);
              if (radio) {
                const lbl = radio.id ? document.querySelector(`label[for="${radio.id}"]`) : null;
                if (lbl) { lbl.click(); filled = true; }
                else { radio.checked = true; radio.dispatchEvent(new Event("change", { bubbles: true })); filled = true; }
              }
            } catch(_) {}
          }
          if (!filled) {
            for (const s of selectorsForRole) { if (fillField(s, value)) { filled = true; break; } }
          }
          log(`  fill answer => ${filled ? "ok" : "fail"} value="${value}"`);
          if (filled) await delay(80);
        }
      } else {
        // 小题听力原文已在 listening_script + hasBlanks 分支里逐空填过，此处不再重复填
        if (hasBlanks && /^blank_script_\d+$/.test(role)) { log(`  skip ${role} (已在 listening_script 分支填过)`); continue; }
        const value = getValueForRole(role);
        if (!value) { log(`  skip ${role} (无值)`); continue; }
        // blank_script_N / blank_audio_N 用 partSel 前缀选择器，Vue/React 可能在 await delay 后清除属性，再次打标
        if (/^blank_(script|audio)_\d+$/.test(role)) {
          try { document.querySelectorAll(".question-part").forEach((pt, pi) => pt.setAttribute("data-fill-part-idx", String(pi))); } catch (_) {}
        }
        // blank_audio_N：填入文本框后派发 blur（与 audio_url 同逻辑）
        if (/^blank_audio_\d+$/.test(role)) {
          for (const s of selectorsForRole) {
            const el = document.querySelector(s);
            if (!el) continue;
            if (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.isContentEditable) {
              const ok = fillField(s, value);
              log(`  fill ${role} => ${ok ? "ok" : "fail"} value="${String(value).slice(0, 50)}${String(value).length > 50 ? "…" : ""}"`);
              el.dispatchEvent(new Event("blur", { bubbles: true }));
              await delay(150);
              break;
            }
          }
        } else {
          let filledAny = false;
          for (const s of selectorsForRole) {
            if (fillField(s, value, role)) { filledAny = true; await delay(role.startsWith("option_") ? 120 : 80); break; }
          }
          log(`  fill ${role} => ${filledAny ? "ok" : "fail"} value="${String(value).slice(0, 50)}${String(value).length > 50 ? "…" : ""}"`);
        }
      }
    }
    for (const role of PROP_ROLES) {
      const selectorsForRole = qSel(curSel, role);
      if (selectorsForRole.length === 0) continue;
      const value = getValueForRole(role);
      if (!value) continue;
      if (role === "difficulty" || role === "question_permission") {
        let done = clickOptionByLabel(selectorsForRole, value, curSel);
        if (!done) {
          for (const s of selectorsForRole) {
            const ok = await clickDropdownAndSelect(s, value);
            if (ok) { done = true; await delay(200); break; }
          }
        }
        if (!done) {
          for (const s of selectorsForRole) { if (fillField(s, value)) { await delay(80); break; } }
        }
        if (done) await delay(80);
      } else if (role === "grade" || role === "course" || role === "unit" || role === "knowledge_point") {
        let done = false;
        for (const s of selectorsForRole) {
          const ok = await clickDropdownAndSelect(s, value);
          if (ok) { done = true; await delay(200); break; }
        }
        if (!done) {
          for (const s of selectorsForRole) {
            if (fillField(s, value)) { await delay(80); break; }
          }
        }
      } else {
        for (const s of selectorsForRole) {
          if (fillField(s, value)) await delay(80);
        }
      }
    }
    await delay(150);
    log(`=== 第 ${i + 1}/${questions.length} 题 字段填充结束`);

    // ── 等待子题 TTS 合成完成并上传音频（一题多小题情况）──
    if (blanksTtsPromises.length > 0) {
      log(`  第 ${i + 1} 题：等待 ${blanksTtsPromises.length} 个子题 TTS 合成完成…`);
      const blanksTtsResults = await Promise.all(blanksTtsPromises);
      
      // TTS 合成期间页面可能重渲染，重新打标 .question-part
      try { document.querySelectorAll(".question-part").forEach((pt, pi) => pt.setAttribute("data-fill-part-idx", String(pi))); } catch (_) {}
      
      for (const result of blanksTtsResults) {
        if (!result.audioBase64) {
          log(`    子题 ${result.index + 1}: TTS 失败，跳过`);
          continue;
        }
        
        showTtsStatus(`第 ${i + 1}/${questions.length} 题：子题 ${result.index + 1} 音频上传中…`);
        log(`    子题 ${result.index + 1}: TTS 成功，音频大小=${result.audioBase64.length}，上传到 ${result.selector}`);
        
        try {
          const audioBlob = base64ToBlob(result.audioBase64, "audio/mpeg");
          if (audioBlob) {
            const audioFile = new File([audioBlob], `tts_q${i + 1}_blank${result.index + 1}.mp3`, { type: "audio/mpeg" });
            
            // 找到对应的音频输入框（先用原始选择器，失败则尝试备选方案）
            let audioUrlEl = document.querySelector(result.selector);
            
            // 如果找不到，尝试从选择器中提取 ID 部分直接查找
            if (!audioUrlEl && result.selector) {
              const idMatch = result.selector.match(/#([\w-]+)/);
              if (idMatch) {
                audioUrlEl = document.getElementById(idMatch[1]);
                if (audioUrlEl) log(`    子题 ${result.index + 1}: 通过 ID 直接找到: #${idMatch[1]}`);
              }
            }
            
            // 还是找不到，尝试按小题索引查找
            if (!audioUrlEl) {
              const parts = document.querySelectorAll(".question-part");
              if (parts[result.index]) {
                audioUrlEl = parts[result.index].querySelector(
                  "input[id*='audioFileName'], input[name*='audio'], input[type='text'][id*='audio']"
                );
                if (audioUrlEl) log(`    子题 ${result.index + 1}: 通过 .question-part[${result.index}] 找到`);
              }
            }
            
            if (!audioUrlEl) {
              log(`    子题 ${result.index + 1}: 未找到音频输入框 ${result.selector}`);
              continue;
            }
            
            // 找到对应的 .question-part 容器
            const part = audioUrlEl.closest(".question-part");
            log(`    子题 ${result.index + 1}: 找到音频输入框，part=${part ? 'Y' : 'N'}`);
            
            let uploaded = false;
            
            // 方案1：在 .question-part 内查找 uploadify 容器
            if (part) {
              const uploadContainer = part.querySelector(
                "[id*='audio_upload'], .uploadify[id*='audio'], [class*='audio-upload'], " +
                "div[id*='upload'][id*='audio'], .uploadify"
              );
              log(`    子题 ${result.index + 1}: part内uploadify容器=${uploadContainer ? uploadContainer.id || uploadContainer.className : 'null'}`);
              
              if (uploadContainer) {
                const fileEl = uploadContainer.querySelector("input[type='file']");
                if (fileEl) {
                  const $ = window.jQuery || window.$;
                  if ($ && typeof $.fn.uploadify === "function") {
                    const settings = $(fileEl).data("uploadify") || $(uploadContainer).data("uploadify");
                    if (settings) {
                      // 检查 fileExt 是否允许音频文件
                      const fileExt = (settings.fileExt || settings.fileTypeExts || "").toLowerCase();
                      const allowsAudio = !fileExt || fileExt.includes("mp3") || fileExt.includes("wav") || fileExt.includes("audio") || fileExt.includes("*.*");
                      if (!allowsAudio) {
                        log(`    子题 ${result.index + 1}: uploadify不支持音频 (fileExt=${fileExt})，跳过`);
                      } else {
                      log(`    子题 ${result.index + 1}: 找到uploadify settings，uploader=${settings.uploader || settings.uploadScript}`);
                      const xhr = new XMLHttpRequest();
                      const fd = new FormData();
                      const fileFieldName = settings.fileObjName || settings.fileField || "Filedata";
                      fd.append(fileFieldName, audioFile);
                      if (settings.formData) {
                        Object.entries(settings.formData).forEach(([k, v]) => fd.append(k, v));
                      }
                      const uploadUrl = settings.uploader || settings.uploadScript;
                      if (uploadUrl) {
                        await new Promise((resolve) => {
                          xhr.open("POST", uploadUrl);
                          xhr.withCredentials = true;
                          xhr.onload = () => {
                            log(`    子题 ${result.index + 1}: uploadify上传完成 => ${xhr.status} resp=${xhr.responseText.slice(0, 80)}`);
                            try {
                              const resp = JSON.parse(xhr.responseText);
                              const fname = resp.file_path || resp.fileName || resp.filename || resp.name || resp.url || "";
                              if (fname) {
                                const nativeDesc = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value");
                                if (nativeDesc && nativeDesc.set) {
                                  nativeDesc.set.call(audioUrlEl, fname);
                                  audioUrlEl.dispatchEvent(new Event("input", { bubbles: true }));
                                  audioUrlEl.dispatchEvent(new Event("change", { bubbles: true }));
                                  audioUrlEl.dispatchEvent(new Event("blur", { bubbles: true }));
                                  log(`    子题 ${result.index + 1}: audioFileName已更新: ${fname}`);
                                }
                              }
                            } catch (pe) { log(`    子题 ${result.index + 1}: 解析响应失败: ${pe.message}`); }
                            resolve();
                          };
                          xhr.onerror = () => { log(`    子题 ${result.index + 1}: uploadify上传网络错误`); resolve(); };
                          xhr.send(fd);
                        });
                        uploaded = true;
                      }
                      }  // allowsAudio
                    } else {
                      log(`    子题 ${result.index + 1}: fileEl存在但无uploadify settings`);
                    }
                  } else {
                    log(`    子题 ${result.index + 1}: jQuery或uploadify不可用`);
                  }
                }
              }
            }
            
            // 方案2：如果方案1失败，尝试用 fillFileField 直接设置
            if (!uploaded) {
              // 尝试在 part 内找 file input
              const partFileInput = part ? part.querySelector("input[type='file']") : null;
              if (partFileInput) {
                // 检查 accept 属性是否支持音频
                const acceptAttr = (partFileInput.accept || "").toLowerCase();
                const acceptsAudio = !acceptAttr || acceptAttr.includes("mp3") || acceptAttr.includes("audio") || acceptAttr.includes("*");
                if (!acceptsAudio) {
                  log(`    子题 ${result.index + 1}: input不支持音频 (accept=${acceptAttr})，跳过`);
                } else {
                  const sel = partFileInput.id ? `#${partFileInput.id}` : null;
                  if (sel) {
                    log(`    子题 ${result.index + 1}: 尝试 fillFileField: ${sel}`);
                    const ok = fillFileField(sel, audioFile);
                    if (ok) {
                      uploaded = true;
                      log(`    子题 ${result.index + 1}: fillFileField 成功`);
                    }
                  }
                }
              }
            }
            
            if (!uploaded) {
              log(`    子题 ${result.index + 1}: 上传失败，所有方案均未成功`);
            }
          }
        } catch (e) {
          log(`    子题 ${result.index + 1}: 音频上传异常: ${e?.message}`);
        }
        
        await delay(800); // 每个子题上传后稍等
      }
      hideTtsStatus();
    }

    // ── 等待 TTS 合成完成并上传音频（顶层音频情况）──
    if (ttsPromise) {
      log(`  第 ${i + 1} 题：等待 TTS 合成完成…`);
      showTtsStatus(`第 ${i + 1}/${questions.length} 题音频合成中…`);
      const ttsAudioBase64 = await ttsPromise;
      
      if (ttsAudioBase64) {
        showTtsStatus(`第 ${i + 1}/${questions.length} 题音频上传中…`);
        log(`  第 ${i + 1} 题：TTS 合成完成，音频大小=${ttsAudioBase64.length}，开始上传`);
        
        // 重新检测当前页面的音频上传框（确保对应当前题目）
        const scope = document.querySelector(".topic-container, #topic-section, #right-container, .question-part") || document.body;
        let currentAudioFileSel = null;
        let currentAudioUrlSel = null;
        
        // 检测音频 URL 输入框（覆盖更多可能的选择器）
        const audioUrlInput = scope.querySelector(
          "#audioFileName, #audioFileName1, input.audioFileName, input[id^='audioFileName'], " +
          "input[id*='audio'][type='text'], input[placeholder*='mp3'], input[placeholder*='音频'], " +
          "input[name*='audio'], input[id*='Audio'], .audio-input input[type='text']"
        );
        if (audioUrlInput) {
          currentAudioUrlSel = audioUrlInput.id ? `#${audioUrlInput.id}` : null;
          log(`  第 ${i + 1} 题：检测到音频URL输入框: ${currentAudioUrlSel}`);
        }
        
        // 检测音频上传容器（只匹配明确与音频相关的选择器，避免误匹配图片上传框）
        const uploadContainer = scope.querySelector(
          "#audio_upload1, #audio_upload, div[id*='audio_upload'], .uploadify[id*='audio'], " +
          "[id='audio_upload1'], [id='audio_upload'], [id*='audioUpload'], .audio-upload-wrap, " +
          "[class*='audio'][class*='upload'], [class*='upload'][class*='audio']"
        );
        if (uploadContainer) {
          const fi = uploadContainer.querySelector("input[type='file']");
          if (fi) {
            currentAudioFileSel = fi.id ? `#${fi.id}` : "input[type='file']";
            log(`  第 ${i + 1} 题：检测到音频上传框: ${currentAudioFileSel}`);
          }
        }
        
        // 如果还没找到，尝试使用初始检测的选择器
        if (!currentAudioFileSel && !currentAudioUrlSel) {
          if (audioFileSels.length > 0) currentAudioFileSel = audioFileSels[0];
          if (audioUrlSels.length > 0) currentAudioUrlSel = audioUrlSels[0];
          log(`  第 ${i + 1} 题：使用初始检测的选择器 - audioFile: ${currentAudioFileSel}, audioUrl: ${currentAudioUrlSel}`);
        }
        
        // 打印更详细的调试信息
        if (!currentAudioFileSel && !currentAudioUrlSel) {
          log(`  第 ${i + 1} 题：未找到任何音频上传框！请检查页面结构`);
          // 尝试列出 scope 内所有可能相关的元素
          const allInputs = scope.querySelectorAll("input[type='file'], .uploadify, [id*='audio'], [class*='audio']");
          log(`  第 ${i + 1} 题：scope内相关元素数量: ${allInputs.length}`);
          allInputs.forEach((el, idx) => {
            if (idx < 5) log(`    - ${el.tagName} id="${el.id}" class="${el.className.slice(0, 50)}"`);
          });
        }
        
        if (currentAudioFileSel || currentAudioUrlSel) {
          try {
            const audioBlob = base64ToBlob(ttsAudioBase64, "audio/mpeg");
            if (audioBlob) {
              const audioFile = new File([audioBlob], `tts_q${i + 1}.mp3`, { type: "audio/mpeg" });
              log(`  第 ${i + 1} 题：音频File创建: name=${audioFile.name} size=${audioFile.size}`);
              
              let uploaded = false;
              
              if (currentAudioFileSel) {
                const fileEl = document.querySelector(currentAudioFileSel);
                if (fileEl) {
                  const $ = window.jQuery || window.$;
                  if ($ && typeof $.fn.uploadify === "function") {
                    try {
                      const settings = $(fileEl).data("uploadify") || $(fileEl).closest("[class*='uploadify'],[id*='upload']").data("uploadify");
                      if (settings) {
                        // 检查 fileExt 是否允许 mp3/音频文件
                        const fileExt = (settings.fileExt || settings.fileTypeExts || "").toLowerCase();
                        const allowsAudio = !fileExt || fileExt.includes("mp3") || fileExt.includes("wav") || fileExt.includes("audio") || fileExt.includes("*.*");
                        if (!allowsAudio) {
                          log(`  第 ${i + 1} 题：uploadify不支持音频 (fileExt=${fileExt})，跳过上传`);
                        } else {
                        log(`  第 ${i + 1} 题：找到uploadify settings，尝试直接上传`);
                        const xhr = new XMLHttpRequest();
                        const fd = new FormData();
                        const fileFieldName = settings.fileObjName || settings.fileField || "Filedata";
                        fd.append(fileFieldName, audioFile);
                        if (settings.formData) {
                          Object.entries(settings.formData).forEach(([k, v]) => fd.append(k, v));
                        }
                        const uploadUrl = settings.uploader || settings.uploadScript;
                        if (uploadUrl) {
                          await new Promise((resolve) => {
                            xhr.open("POST", uploadUrl);
                            xhr.withCredentials = true;
                            xhr.onload = () => {
                              log(`  第 ${i + 1} 题：uploadify上传完成 => ${xhr.status} ${xhr.responseText.slice(0, 100)}`);
                              try {
                                const resp = JSON.parse(xhr.responseText);
                                const fname = resp.file_path || resp.fileName || resp.filename || resp.name || resp.url || "";
                                if (fname && currentAudioUrlSel) {
                                  const targetEl = document.querySelector(currentAudioUrlSel);
                                  if (targetEl) {
                                    const nativeDesc = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value");
                                    if (nativeDesc && nativeDesc.set) {
                                      nativeDesc.set.call(targetEl, fname);
                                      targetEl.dispatchEvent(new Event("input", { bubbles: true }));
                                      targetEl.dispatchEvent(new Event("change", { bubbles: true }));
                                      log(`  第 ${i + 1} 题：audioFileName已更新: ${fname}`);
                                    }
                                  }
                                }
                              } catch (_) {}
                              resolve();
                            };
                            xhr.onerror = () => { log(`  第 ${i + 1} 题：uploadify上传网络错误`); resolve(); };
                            xhr.send(fd);
                          });
                          uploaded = true;
                          await delay(1500);
                        }
                        }  // allowsAudio
                      }
                    } catch (e2) {
                      log(`  第 ${i + 1} 题：uploadify上传异常: ${e2?.message}`);
                    }
                  }
                  // 只有在上传框支持音频的情况下才尝试 fillFileField
                  if (!uploaded) {
                    // 检查 input 的 accept 属性
                    const acceptAttr = (fileEl.accept || "").toLowerCase();
                    const acceptsAudio = !acceptAttr || acceptAttr.includes("mp3") || acceptAttr.includes("audio") || acceptAttr.includes("*");
                    if (acceptsAudio) {
                      const ok = fillFileField(currentAudioFileSel, audioFile);
                      log(`  第 ${i + 1} 题：fillFileField=${ok}`);
                      if (ok) {
                        uploaded = true;
                        await delay(1500);
                      }
                    } else {
                      log(`  第 ${i + 1} 题：input[type=file]不支持音频 (accept=${acceptAttr})，跳过`);
                    }
                  }
                }
              }
              
              if (!uploaded) {
                log(`  第 ${i + 1} 题：音频上传失败或未找到上传框`);
              }
            }
          } catch (e) {
            log(`  第 ${i + 1} 题：音频上传异常: ${e?.message}`);
          }
        } else {
          log(`  第 ${i + 1} 题：未找到音频上传框，跳过音频上传`);
        }
      } else {
        log(`  第 ${i + 1} 题：TTS 合成失败，跳过音频上传`);
      }
      hideTtsStatus();
    }

    const clickKeyCur = (key) => {
      // next_btn 必须始终走 findNextBtn()，该函数要求「继续录题」在保存成功容器内才返回，
      // 避免直接点检测到的「下一题」按钮绕过保存成功判断而误切题
      if (key === "next_btn") {
        const nb = findNextBtn();
        if (nb) {
          nb.focus();
          nb.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, view: window }));
          nb.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, view: window }));
          nb.click();
          return true;
        }
        return false;
      }
      for (const s of qSel(curSel, key)) {
        try {
          const btn = document.querySelector(s);
          if (btn && btn.offsetParent !== null) {
            btn.focus();
            btn.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, view: window }));
            btn.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, view: window }));
            btn.click();
            return true;
          }
        } catch (_) {}
      }
      if (key === "submit_btn") {
        const sb = findSubmitBtn();
        if (sb) {
          sb.focus();
          sb.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, view: window }));
          sb.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, view: window }));
          sb.click();
          return true;
        }
      }
      return false;
    };

    const isLast = i >= questions.length - 1;
    // 驰声平台流程：点「预览/保存」→ 页面自动跳下一题（无需点任何按钮）
    const prevTopicIdx = getCurrentTopicIndex(findTopicNumbers());
    const saved = clickKeyCur("submit_btn");
    log(`点击保存按钮: saved=${saved}, 当前左侧题号索引: ${prevTopicIdx}`);
    if (!isLast) {
      // 以「左侧题号选中索引变化」作为切题完成信号，最多等 25 秒（音频上传需要更长时间）
      let topicChanged = false;
      for (let t = 0; t < 25000; t += 300) {
        await delay(300);
        const curIdx = getCurrentTopicIndex(findTopicNumbers());
        if (curIdx !== prevTopicIdx) {
          log(`左侧题号已切换: ${prevTopicIdx} → ${curIdx}，耗时约 ${t + 300}ms`);
          topicChanged = true;
          break;
        }
      }
      if (!topicChanged) {
        log(`警告：25秒内左侧题号未变化，第 ${i + 1} 题可能保存失败（可能有必填项未填，如音频）`);
        notify(`第 ${i + 1} 题保存未自动跳转，已暂停。请手动补全后点弹窗「继续填充」`);
        hideTtsStatus();
        // 将剩余题目（不含当前失败的题）返回给 popup，由用户手动修复后点「继续填充」
        return {
          ok: "paused",
          filled,
          pausedQuestion: i + 1,       // 1-based，展示给用户
          remaining: questions.slice(i + 1), // 剩余待填题目
          error: `第 ${i + 1} 题保存后未自动跳转，可能有必填项（如音频）未填。`,
        };
      }
      // 题号切换后再等 800ms，让新题表单完成渲染（音频字段初始化需要更长时间）
      await delay(800);
      // 仅通过「保存 + 继续录题」切题，不点左侧题号（点题号切题不会保存，与正式流程一致）
    } else {
      await delay(200);
    }
    filled++;
    notify(`已完成第 ${i + 1}/${questions.length} 题`);
  }

  hideTtsStatus();
  log(`=== 全部填充结束 filled=${filled}，复制日志: copy(localStorage.getItem('__aiLutiLog'))`);
  return { ok: true, filled, message: `填充完成，共 ${filled} 题。` };
}
