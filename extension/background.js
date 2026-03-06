// Background Service Worker：处理豆包识别的长任务
// 弹窗关闭后任务依然继续，结果写入 chrome.storage.local

const BACKEND = "http://127.0.0.1:8766";
const PARSE_TIMEOUT_MS = 3 * 60 * 1000; // 最长等待 3 分钟

// 当前 fetch 的 AbortController，用于取消识别
let currentController = null;
let _timeoutId = null;
// 识别中每秒广播已等待秒数，便于用户判断是请求中还是卡死
let _parseProgressInterval = null;

// 记录当前解析任务对应的录题页 tabId，用于更新该标签上的角标
let parseTabId = null;

// ─── 角标：在扩展图标上显示数字/状态（仅对录题页 tab 生效）────────────────────
function setBadge(tabId, text, color) {
  if (!tabId) return;
  chrome.action.setBadgeText({ tabId, text: text ? String(text).slice(0, 4) : "" });
  chrome.action.setBadgeBackgroundColor({
    tabId,
    color: color || "#000000",
  }).catch(() => {});
}

async function updateBadgeForRecordTab(tabId, reason, extra) {
  if (!tabId) return;
  if (reason === "parse_parsing") {
    setBadge(tabId, "...", "#000000"); // 三个点居中更好，背景统一黑色
  } else if (reason === "parse_done") {
    const n = extra?.count ?? 0;
    setBadge(tabId, n > 0 ? String(n) : "✓", "#000000");
  } else if (reason === "parse_error") {
    setBadge(tabId, "!", "#000000");
  } else if (reason === "parse_idle" || reason === "clear") {
    const { lastDetectedUrl, pageQuestionTotal } = await chrome.storage.sync.get(["lastDetectedUrl", "pageQuestionTotal"]);
    const tab = await chrome.tabs.get(tabId).catch(() => null);
    const url = tab?.url || "";
    if (url && lastDetectedUrl && url === lastDetectedUrl && pageQuestionTotal != null && pageQuestionTotal > 0) {
      setBadge(tabId, String(pageQuestionTotal), "#000000");
    } else {
      setBadge(tabId, "", null);
    }
  }
}

// ─── 监听标签页 URL 变化，自动清理旧状态并更新角标 ───────────────────────────
async function isRecordPage(url) {
  const { recordPagePattern } = await chrome.storage.sync.get("recordPagePattern");
  const custom = (recordPagePattern || "").trim().toLowerCase();
  const u = (url || "").toLowerCase();
  if (custom) return u.includes(custom);
  return u.includes("91tszx") || u.includes("chisheng");
}

chrome.tabs.onUpdated.addListener(async (tabId, changeInfo, tab) => {
  const url = changeInfo.url || tab?.url;
  if (!url) return;

  if (changeInfo.url && (await isRecordPage(changeInfo.url))) {
    const { lastDetectedUrl, pageQuestionTotal } = await chrome.storage.sync.get(["lastDetectedUrl", "pageQuestionTotal"]);
    if (lastDetectedUrl && changeInfo.url !== lastDetectedUrl) {
      if (currentController) {
        currentController._isPageChange = true;
        currentController.abort();
        currentController = null;
      }
      await chrome.storage.sync.remove(["selectors", "lastDetectMessage", "fields", "pageQuestionTotal"]);
      await chrome.storage.local.remove("parseState");
      await chrome.storage.sync.set({ lastDetectedUrl: changeInfo.url });
      setBadge(tabId, "", null);
      broadcastToPopup({ type: "PARSE_CANCELLED" });
    } else if (changeInfo.url === lastDetectedUrl && pageQuestionTotal != null && pageQuestionTotal > 0) {
      setBadge(tabId, String(pageQuestionTotal), "#000000");
    }
  } else if (changeInfo.url) {
    setBadge(tabId, "", null);
  }
});

function broadcastToPopup(msg) {
  chrome.runtime.sendMessage(msg).catch(() => {});
}

/** 由 popup 点击「填入 N 题」触发；在 background 执行，弹窗关闭后仍可完成填入 */
async function handleRequestFill(questions) {
  if (!questions || questions.length === 0) {
    return { ok: false, error: "没有可填入的题目" };
  }
  const { selectors, pageQuestionTotal, pageSlots, defaultAudioUrl, defaultImageUrl } = await chrome.storage.sync.get(["selectors", "pageQuestionTotal", "pageSlots", "defaultAudioUrl", "defaultImageUrl"]);
  if (!selectors || !Object.values(selectors).some(Boolean)) {
    return { ok: false, error: "请先完成「步骤1：检测页面结构」再填充" };
  }
  let tabId = (await chrome.storage.local.get("fillTargetTabId")).fillTargetTabId || null;
  if (!tabId) {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    tabId = tab?.id || null;
  }
  if (!tabId) {
    return { ok: false, error: "无法确定录题页，请确认页面已打开" };
  }
  let toFill = normalizeQuestionsToSlots(questions, null, pageSlots || null);
  const pageTotal = pageQuestionTotal != null && pageQuestionTotal > 0 ? pageQuestionTotal : null;
  // 只裁掉多余的题，不补空题：题数不足时只填有内容的题，不反复点「下一题」填空白
  if (pageTotal != null && toFill.length > pageTotal) {
    toFill = toFill.slice(0, pageTotal);
  }
  try {
    const result = await chrome.tabs.sendMessage(tabId, {
      type: "FILL_FORM",
      questions: toFill,
      selectors,
      defaultAudioUrl: (defaultAudioUrl || "").trim() || undefined,
      defaultImageUrl: (defaultImageUrl || "").trim() || undefined,
      debugSource: "parse",
    });
    if (result?.ok === false) {
      return { ok: false, error: result.error || "填充出错" };
    }
    broadcastToPopup({ type: "FILL_DONE", message: result?.message || `填充完成，共 ${result?.filled ?? toFill.length} 题。` });
    await chrome.storage.local.remove("fillTargetTabId");
    setState({ status: "idle" });
    updateBadgeForRecordTab(tabId, "parse_idle");
    return { ok: true, filled: result?.filled ?? toFill.length, message: result?.message };
  } catch (e) {
    if (String(e).includes("Could not establish connection") || String(e).includes("receiving end does not exist")) {
      return { ok: false, error: "录题页未响应，请刷新该页面后再试" };
    }
    return { ok: false, error: e?.message || String(e) };
  }
}

function setState(state) {
  return chrome.storage.local.set({ parseState: state });
}

/** 将小题列表规范为大题列表（与 popup 中逻辑一致），供 FILL_FORM 使用 */
function normalizeQuestionsToSlots(questions, pageTotal, pageSlots) {
  if (!questions || questions.length === 0) return questions;
  const noBlanks = (x) => !x.blanks || !Array.isArray(x.blanks) || x.blanks.length === 0;
  const hasBlanks = (x) => x.blanks && Array.isArray(x.blanks) && x.blanks.length > 0;

  // ── 优先路径：pageSlots 已知每题 subCount ──
  if (pageSlots && pageSlots.length > 0) {
    const out = [];
    let qi = 0;
    for (let si = 0; si < pageSlots.length && qi < questions.length; si++) {
      const slot = pageSlots[si];
      const need = (slot.subCount && slot.subCount > 1) ? slot.subCount : 1;
      if (hasBlanks(questions[qi])) { out.push(questions[qi]); qi++; continue; }
      if (need === 1) {
        out.push(questions[qi]); qi++;
      } else {
        const group = questions.slice(qi, qi + need);
        const fullScript = group.reduce((best, cur) =>
          (cur.listening_script || "").length > best.length ? (cur.listening_script || "") : best, "");
        out.push({
          type: group[0].type,
          listening_script: fullScript || group[0].listening_script || "",
          question: "",
          options: [],
          answer: "",
          explanation: group.map(b => b.explanation).filter(Boolean).join("；"),
          blanks: group.map(b => ({
            question: b.question || "",
            answer: b.answer != null ? String(b.answer).trim() : "",
            options: Array.isArray(b.options) ? b.options : [],
            candidates: Array.isArray(b.candidates) ? b.candidates : [],
          })),
        });
        qi += need;
      }
    }
    while (qi < questions.length) { out.push(questions[qi]); qi++; }
    return out.map(q => normalizeAnswers(q));
  }

  // ── 兜底路径：按 listening_script / 题型推测合并 ──
  const out = [];
  let i = 0;
  while (i < questions.length) {
    const q = questions[i];
    if (hasBlanks(q)) { out.push(q); i++; continue; }
    if (q.type === "listening_fill" && noBlanks(q)) {
      const blanks = [];
      let j = i;
      while (j < questions.length && questions[j].type === "listening_fill" && noBlanks(questions[j])) {
        blanks.push({
          question: questions[j].question || "",
          answer: questions[j].answer != null ? String(questions[j].answer).trim() : "",
        });
        j++;
      }
      if (blanks.length > 0) {
        out.push({
          type: "listening_fill",
          listening_script: q.listening_script || "",
          question: q.question || blanks.map((b) => b.question).join(" ").slice(0, 200),
          options: q.options || [],
          answer: blanks[0]?.answer ?? "",
          explanation: q.explanation || "",
          blanks,
        });
        i = j;
        continue;
      }
    }
    const normScript = (s) => (s || "").trim().replace(/\r\n/g, "\n");
    const sameConversation = (a, b) => {
      if (!a || !b) return false;
      const na = normScript(a), nb = normScript(b);
      return na === nb || na.includes(nb) || nb.includes(na);
    };
    if (noBlanks(q) && q.listening_script && q.listening_script.trim() !== "") {
      let j = i + 1;
      let groupScript = normScript(q.listening_script);
      while (
        j < questions.length && questions[j].type === q.type && noBlanks(questions[j]) &&
        questions[j].listening_script && sameConversation(groupScript, normScript(questions[j].listening_script))
      ) {
        const ns = normScript(questions[j].listening_script);
        if (ns.length > groupScript.length) groupScript = ns;
        j++;
      }
      if (j > i + 1) {
        const group = questions.slice(i, j);
        const fullScript = group.reduce((best, cur) =>
          (cur.listening_script || "").length > best.length ? (cur.listening_script || "") : best, "");
        out.push({
          type: q.type, listening_script: fullScript, question: "", options: [], answer: "",
          explanation: group.map(b => b.explanation).filter(Boolean).join("；"),
          blanks: group.map(b => ({
            question: b.question || "", answer: b.answer != null ? String(b.answer).trim() : "",
            options: Array.isArray(b.options) ? b.options : [],
            candidates: Array.isArray(b.candidates) ? b.candidates : [],
          })),
        });
        i = j; continue;
      }
    }
    out.push(q);
    i++;
  }
  // 统一归一化所有答案字段：无论 AI 输出 "C. Four." / "(C)" / "3" 都转为 "C"
  return out.map(q => normalizeAnswers(q));
}

/** 把单道题的 answer / blanks[*].answer 归一化为纯字母（A/B/C/D）或原始值（非选择题） */
function normalizeAnswers(q) {
  const norm = (v) => {
    if (v == null) return v;
    const s = String(v).trim();
    // "C. Four." / "C）答案" / "(C)" / "C、" → "C"
    const m = s.match(/^[（(]?([A-Da-d])[)）.\s、。，,]/);
    if (m) return m[1].toUpperCase();
    // 纯字母 "C" 或 "c"
    if (/^[A-Da-d]$/.test(s)) return s.toUpperCase();
    // 数字 "1"~"4" → "A"~"D"
    if (/^[1-4]$/.test(s)) return String.fromCharCode(64 + parseInt(s, 10));
    return s; // 其他（填空题答案等）原样保留
  };
  const q2 = { ...q };
  // 数组答案（表格填空等）不做字母归一化，只对字符串类型处理
  if (q2.answer != null && !Array.isArray(q2.answer)) q2.answer = norm(q2.answer);
  if (Array.isArray(q2.blanks)) {
    q2.blanks = q2.blanks.map(b => {
      if (!b || b.answer == null || Array.isArray(b.answer)) return b;
      return { ...b, answer: norm(b.answer) };
    });
  }
  return q2;
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.type === "START_PARSE") {
    parseTabId = msg.tabId || sender.tab?.id || null;
    if (parseTabId) chrome.storage.local.set({ fillTargetTabId: parseTabId });
    handleParse(msg.filesData);
    if (parseTabId) updateBadgeForRecordTab(parseTabId, "parse_parsing");
    sendResponse({ ok: true });
    return false;
  }

  if (msg.type === "REQUEST_FILL") {
    handleRequestFill(msg.questions).then(sendResponse).catch((e) => sendResponse({ ok: false, error: String(e?.message) }));
    return true;
  }

  if (msg.type === "CANCEL_PARSE") {
    if (currentController) {
      currentController.abort();
      currentController = null;
    }
    setState({ status: "idle" });
    if (parseTabId) updateBadgeForRecordTab(parseTabId, "parse_idle");
    parseTabId = null;
    sendResponse({ ok: true });
    return false;
  }

  if (msg.type === "GET_PARSE_STATE") {
    chrome.storage.local.get("parseState").then(({ parseState }) => {
      sendResponse({ state: parseState || { status: "idle" } });
    });
    return true;
  }

  if (msg.type === "CLEAR_PARSE_STATE") {
    setState({ status: "idle" });
    if (parseTabId) updateBadgeForRecordTab(parseTabId, "parse_idle");
    sendResponse({ ok: true });
    return false;
  }

  if (msg.type === "REFRESH_BADGE") {
    if (msg.tabId) updateBadgeForRecordTab(msg.tabId, "parse_idle");
    sendResponse({ ok: true });
    return false;
  }
});

async function handleParse(filesData) {
  // 中止上一个未完成的请求
  if (currentController) currentController.abort();
  clearTimeout(_timeoutId);

  currentController = new AbortController();
  const signal = currentController.signal;

  // 超时自动中止
  _timeoutId = setTimeout(() => {
    if (currentController) {
      currentController._isTimeout = true;
      currentController.abort();
    }
  }, PARSE_TIMEOUT_MS);

  await setState({ status: "parsing", text: "正在豆包识别…" });
  const parseStartTime = Date.now();
  const sendProgress = () => {
    const elapsed = Math.floor((Date.now() - parseStartTime) / 1000);
    broadcastToPopup({ type: "PARSE_PROGRESS", text: "正在豆包识别…", elapsed });
  };
  sendProgress(); // 立即发一次
  _parseProgressInterval = setInterval(sendProgress, 1000);

  try {
    const formData = new FormData();
    for (const f of filesData) {
      const byteString = atob(f.base64);
      const bytes = new Uint8Array(byteString.length);
      for (let i = 0; i < byteString.length; i++) {
        bytes[i] = byteString.charCodeAt(i);
      }
      const blob = new Blob([bytes], { type: f.mimeType });
      formData.append("files", blob, f.name);
    }

    const { fields, pageQuestionTotal, pageSlots, parseDebugMode, selectedModel } = await chrome.storage.sync.get([
      "fields", "pageQuestionTotal", "pageSlots", "parseDebugMode", "selectedModel",
    ]);
    if (fields && fields.length > 0) {
      formData.append("field_structure", JSON.stringify(fields));
    }
    if (pageQuestionTotal != null && pageQuestionTotal > 0) {
      formData.append("expected_total", String(pageQuestionTotal));
    }
    // 优先使用「分析页面」时存储的 pageSlots（含每题 subCount + sectionLabels）
    // 兜底：实时向录题页取当前题的区块信息
    if (pageSlots && pageSlots.length > 0) {
      formData.append("page_structure", JSON.stringify(pageSlots));
    } else if (parseTabId) {
      try {
        const struct = await chrome.tabs.sendMessage(parseTabId, { type: "GET_PAGE_STRUCTURE" });
        if (struct && struct.ok && struct.total > 0) {
          const sectionLabels = struct.sectionLabels || [];
          const subCount = struct.subCount || 1;
          const slots = Array.from({ length: struct.total }, (_, i) => ({ index: i + 1, sectionLabels, subCount }));
          formData.append("page_structure", JSON.stringify(slots));
        }
      } catch (_) {}
    }
    if (parseDebugMode) {
      formData.append("debug", "1");
    }
    if (selectedModel && selectedModel !== "default") {
      formData.append("model_override", selectedModel);
    }

    const r = await fetch(BACKEND + "/api/parse-multiple", {
      method: "POST",
      body: formData,
      signal,
    });

    if (!r.ok) {
      // 尝试解析后端返回的错误详情
      let errDetail = `HTTP ${r.status}`;
      try {
        const body = await r.json();
        errDetail = body?.detail?.error || body?.detail || body?.message || errDetail;
      } catch {
        errDetail = (await r.text().catch(() => errDetail)) || errDetail;
      }
      throw new Error(errDetail);
    }

    const data = await r.json();
    const questions = data.questions || [];
    const debug_info = data.debug_info || null;

    if (questions.length === 0) {
      const errText = "未解析到题目，请检查 Word 内容格式是否正确";
      await setState({ status: "error", text: errText });
      if (parseTabId) updateBadgeForRecordTab(parseTabId, "parse_error");
      broadcastToPopup({ type: "PARSE_ERROR", text: errText });
      return;
    }

    const doneText = `豆包识别完成，共 ${questions.length} 题。`;
    await setState({ status: "done", questions, debug_info, text: doneText });
    if (parseTabId) updateBadgeForRecordTab(parseTabId, "parse_done", { count: questions.length });
    broadcastToPopup({ type: "PARSE_DONE", questions, debug_info, text: doneText });

  } catch (e) {
    if (e.name === "AbortError") {
      if (currentController?._isTimeout) {
        const errText = "识别超时（超过3分钟），请检查：①后端服务是否运行 ②豆包 API 是否可用";
        await setState({ status: "error", text: errText });
        if (parseTabId) updateBadgeForRecordTab(parseTabId, "parse_error");
        broadcastToPopup({ type: "PARSE_ERROR", text: errText });
      } else if (currentController?._isPageChange) {
        await setState({ status: "idle" });
        if (parseTabId) updateBadgeForRecordTab(parseTabId, "parse_idle");
        broadcastToPopup({ type: "PARSE_CANCELLED" });
      } else {
        await setState({ status: "idle" });
        if (parseTabId) updateBadgeForRecordTab(parseTabId, "parse_idle");
        broadcastToPopup({ type: "PARSE_CANCELLED" });
      }
      return;
    }
    const errText = `解析失败：${e.message || e}`;
    await setState({ status: "error", text: errText });
    if (parseTabId) updateBadgeForRecordTab(parseTabId, "parse_error");
    broadcastToPopup({ type: "PARSE_ERROR", text: errText });
  } finally {
    if (_parseProgressInterval) {
      clearInterval(_parseProgressInterval);
      _parseProgressInterval = null;
    }
    clearTimeout(_timeoutId);
    currentController = null;
    parseTabId = null;
  }
}
