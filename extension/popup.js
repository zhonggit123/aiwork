const BACKEND = "http://127.0.0.1:8766";
const OUR_PAGE_URL = "http://127.0.0.1:8766";
const onlyWhenDetected = document.getElementById("onlyWhenDetected");
const whenNotDetected = document.getElementById("whenNotDetected");
const statusBadge = document.getElementById("statusBadge");

async function isRecordPage(tab) {
  const { recordPagePattern } = await chrome.storage.sync.get("recordPagePattern");
  const custom = (recordPagePattern || "").trim().toLowerCase();
  const url = (tab?.url || "").toLowerCase();
  if (custom) return url.includes(custom);
  return url.includes("91tszx") || url.includes("chisheng");
}

async function refreshMonitorState() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  const detected = tab?.id && (await isRecordPage(tab));
  const resultEl = document.getElementById("detect-result");
  if (detected) {
    // 检测页面 URL 是否变化，若变化则清除旧状态
    const { lastDetectedUrl } = await chrome.storage.sync.get("lastDetectedUrl");
    const currentUrl = tab.url || "";
    if (lastDetectedUrl && currentUrl !== lastDetectedUrl) {
      // URL 变了，清除旧的检测结果和解析状态
      await chrome.storage.sync.remove(["selectors", "lastDetectMessage", "fields", "pageQuestionTotal"]);
      await chrome.storage.local.remove("parseState");
      // 通知 background 取消正在进行的解析
      chrome.runtime.sendMessage({ type: "CANCEL_PARSE" }).catch(() => {});
      setDropZoneState("idle");
    }
    // 更新记录的 URL
    await chrome.storage.sync.set({ lastDetectedUrl: currentUrl });

    statusBadge.title = "已连接录题页";
    statusBadge.classList.remove("inactive");
    onlyWhenDetected.classList.remove("hidden");
    whenNotDetected.classList.add("hidden");
    const { selectors, lastDetectMessage } = await chrome.storage.sync.get(["selectors", "lastDetectMessage"]);
    const hintEl = document.getElementById("uploadStepHint");
    const resultEl = document.getElementById("detect-result");
    if (selectors && Object.keys(selectors).filter(k => selectors[k]).length > 0) {
      const fields = Object.keys(selectors).filter(k => selectors[k]).join("、");
      resultEl.textContent = (lastDetectMessage || "已了解各题结构，允许上传 Word。")
        + `（已识别字段：${fields}）`;
      resultEl.className = "message-area text-success";
      resultEl.style.display = "block";
      
      const detectBtnIcon = document.querySelector("#detect .card-icon");
      const detectBtnTitle = document.querySelector("#detect .card-title");
      const detectBtnDesc = document.querySelector("#detect .card-desc");
      
      // 更新按钮状态为已识别
      if (detectBtnIcon) {
        detectBtnIcon.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7"></path></svg>`;
        detectBtnIcon.style.color = "var(--brand-green)";
        detectBtnIcon.style.background = "rgba(16, 185, 129, 0.1)";
        detectBtnIcon.classList.remove("icon-spin");
      }
      if (detectBtnTitle) {
        detectBtnTitle.textContent = "已分析";
        detectBtnTitle.style.color = "var(--brand-green)";
      }
      if (detectBtnDesc) {
        detectBtnDesc.textContent = "页面结构已提取";
      }
      
      if (hintEl) {
        hintEl.textContent = "已提取当前页面录题结构，可直接上传。";
        hintEl.style.color = "var(--brand-green)";
        hintEl.style.fontWeight = "500";
      }
      // 已分析状态下清除「请先分析页面」类错误横条
      setMsg("", false);
      
      // Clear result auto hiding logic for persistent info
      setTimeout(() => {
        if (resultEl && resultEl.style.display !== "none") {
          resultEl.style.display = "none";
        }
      }, 5000);

    } else {
      resultEl.textContent = "尚未分析当前页面，请点击上方「分析页面」按钮。";
      resultEl.className = "message-area text-muted";
      resultEl.style.display = "block";
      
      const detectBtnIcon = document.querySelector("#detect .card-icon");
      const detectBtnTitle = document.querySelector("#detect .card-title");
      const detectBtnDesc = document.querySelector("#detect .card-desc");
      
      // 恢复按钮初始状态
      if (detectBtnIcon) {
        detectBtnIcon.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg>`;
        detectBtnIcon.style.color = "var(--gray-600)";
        detectBtnIcon.style.background = "var(--gray-100)";
        detectBtnIcon.classList.remove("icon-spin");
      }
      if (detectBtnTitle) {
        detectBtnTitle.textContent = "分析页面";
        detectBtnTitle.style.color = "var(--gray-800)";
      }
      if (detectBtnDesc) {
        detectBtnDesc.textContent = "识别当前表单";
      }

      if (hintEl) {
        hintEl.textContent = "⚠️ 注意：请先点击上方「分析页面」以获取结构！";
        hintEl.style.color = "var(--warning)";
        hintEl.style.fontWeight = "500";
      }
    }
  } else {
    statusBadge.title = "未检测到录题页";
    statusBadge.classList.add("inactive");
    onlyWhenDetected.classList.add("hidden");
    whenNotDetected.classList.remove("hidden");
    if (tab?.id) chrome.runtime.sendMessage({ type: "REFRESH_BADGE", tabId: tab.id }).catch(() => {});
  }
}

refreshMonitorState();
restoreParseState();

// ─── 设置面板 ─────────────────────────────────────────────────────────────────
(async () => {
  const toggleBtn = document.getElementById("settingsToggle");
  const dropdown  = document.getElementById("settingsDropdown");
  if (!toggleBtn || !dropdown) return;

  // 打开/关闭下拉
  const openDropdown = () => {
    dropdown.classList.remove("hidden");
    toggleBtn.classList.add("settings-open");
  };
  const closeDropdown = () => {
    dropdown.classList.add("hidden");
    toggleBtn.classList.remove("settings-open");
  };

  toggleBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    dropdown.classList.contains("hidden") ? openDropdown() : closeDropdown();
  });
  // 点击面板外关闭
  document.addEventListener("click", (e) => {
    if (!dropdown.classList.contains("hidden") && !dropdown.contains(e.target)) {
      closeDropdown();
    }
  });
  dropdown.addEventListener("click", (e) => e.stopPropagation());

  // ── 模型选择 ──────────────────────────────────────────────
  const { selectedModel, parseDebugMode: savedDebug, defaultAudioUrl: savedAudio, defaultImageUrl: savedImage }
    = await chrome.storage.sync.get(["selectedModel", "parseDebugMode", "defaultAudioUrl", "defaultImageUrl"]);

  const radios = document.querySelectorAll('input[name="modelChoice"]');
  const savedM = selectedModel || "default";
  radios.forEach(r => { r.checked = (r.value === savedM); });
  radios.forEach(r => {
    r.addEventListener("change", () => {
      if (r.checked) chrome.storage.sync.set({ selectedModel: r.value });
    });
  });

  // 从后端读取默认模型名显示
  try {
    const cfg = await (await fetch("http://localhost:18080/api/config")).json();
    const defModel = (cfg?.llm?.model || "").split("/").pop(); // 只显示短名
    if (defModel) {
      const labelEl = document.getElementById("modelDefaultLabel");
      if (labelEl) labelEl.textContent = `默认 (${defModel})`;
    }
  } catch (_) {}

  // ── 解析调试模式切换开关 ──────────────────────────────────
  const cbDebug      = document.getElementById("parseDebugMode");
  const debugSwitch  = document.getElementById("debugToggleSwitch");
  const syncDebugUI  = (on) => {
    if (cbDebug)     cbDebug.checked = on;
    if (debugSwitch) debugSwitch.classList.toggle("on", on);
  };
  syncDebugUI(!!savedDebug);
  const debugToggleItem = debugSwitch?.closest(".settings-toggle-item");
  if (debugToggleItem) {
    debugToggleItem.addEventListener("click", () => {
      const next = !cbDebug?.checked;
      syncDebugUI(next);
      chrome.storage.sync.set({ parseDebugMode: next });
    });
  }

  // ── 高级 JSON → 全屏页面导航 ─────────────────────────────
  const advPage       = document.getElementById("advancedPage");
  const advPageBack   = document.getElementById("advancedPageBack");
  const advToggleItem = document.getElementById("settingsAdvancedToggle");

  const openAdvPage = () => {
    if (advPage) advPage.classList.remove("hidden");
    closeDropdown();
  };
  const closeAdvPage = () => {
    if (advPage) advPage.classList.add("hidden");
  };

  if (advToggleItem) advToggleItem.addEventListener("click", openAdvPage);
  if (advPageBack)   advPageBack.addEventListener("click", closeAdvPage);
  // ── 默认媒体 URL ──────────────────────────────────────────
  const audioEl = document.getElementById("defaultAudioUrl");
  const imageEl = document.getElementById("defaultImageUrl");
  if (audioEl) {
    if (savedAudio) audioEl.value = savedAudio;
    const saveAudio = () => chrome.storage.sync.set({ defaultAudioUrl: audioEl.value.trim() });
    audioEl.addEventListener("change", saveAudio);
    audioEl.addEventListener("blur",   saveAudio);
  }
  if (imageEl) {
    if (savedImage) imageEl.value = savedImage;
    const saveImage = () => chrome.storage.sync.set({ defaultImageUrl: imageEl.value.trim() });
    imageEl.addEventListener("change", saveImage);
    imageEl.addEventListener("blur",   saveImage);
  }
})();

// ─── 接收来自 background.js 和 content.js 的实时消息 ───────────────────────
chrome.runtime.onMessage.addListener((msg) => {
  const resultEl = document.getElementById("detect-result");
  if (resultEl) resultEl.style.display = "block";
  if (msg.type === "DETECT_PROGRESS") {
    const resultEl = document.getElementById("detect-result");
    if (!resultEl) return;
    const fields = (msg.fields || []).join("、");
    const total = msg.total ? `/${msg.total}` : "";
    resultEl.textContent = `正在遍历第 ${msg.walked}${total} 题，已识别字段：${fields || "检测中…"}`;
    resultEl.className = "message-area text-muted";
    
    const detectBtnIcon = document.querySelector("#detect .card-icon");
    if (detectBtnIcon && !detectBtnIcon.classList.contains("icon-spin")) {
      detectBtnIcon.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>`;
      detectBtnIcon.classList.add("icon-spin");
      detectBtnIcon.style.color = "var(--brand-green)";
      detectBtnIcon.style.background = "rgba(16, 185, 129, 0.1)";
      
      const detectBtnTitle = document.querySelector("#detect .card-title");
      if (detectBtnTitle) {
        detectBtnTitle.textContent = "分析中...";
        detectBtnTitle.style.color = "var(--brand-green)";
      }
      const detectBtnDesc = document.querySelector("#detect .card-desc");
      if (detectBtnDesc) {
        detectBtnDesc.textContent = "正在遍历各题";
      }
    }
  }
  if (msg.type === "FILL_PROGRESS") {
    setMsg(msg.text || "", false);
  }
  if (msg.type === "PARSE_PROGRESS") {
    const subEl = document.getElementById("dzParsingSubtext");
    if (subEl) {
      const elapsed = msg.elapsed;
      if (typeof elapsed === "number") {
        subEl.textContent = `已等待 ${elapsed} 秒 · 关闭弹窗不会中断`;
      } else {
        subEl.textContent = "关闭弹窗不会中断";
      }
    }
    // 同时在消息区显示当前文件名
    if (msg.text) setMsg(msg.text, false);
  }
  if (msg.type === "PARSE_DONE") {
    setDropZoneState("done", { questions: msg.questions, debug_info: msg.debug_info });
    setMsg("", false);
    updateFillCacheDot();
    pushJsonHistory(JSON.stringify(msg.questions)).catch(() => {});
    doFill(msg.questions);
  }
  if (msg.type === "PARSE_ERROR") {
    setDropZoneState("idle");
    setErrorMsg(msg.text);
    updateFillCacheDot();
  }
  if (msg.type === "PARSE_CANCELLED") {
    setDropZoneState("idle");
    setMsg("识别已取消，可重新上传文件。", false);
    updateFillCacheDot();
  }
  if (msg.type === "FILL_DONE") {
    setMsg(msg.message || "填入完成", false);
  }
});

// ─── 统一提示与错误消息区 ──────────────────────────────────────────────────
function setMsg(text, isError) {
  const el = document.getElementById("msg");
  if (!el) return;
  el.textContent = text;
  
  if (text) {
    el.className = "message-area" + (isError ? " text-error" : " text-success");
    el.style.display = "block";
  } else {
    el.style.display = "none";
  }

  // 仅在有错误/提示文案时显示复制按钮，并确保点击复制的是当前这段文案
  const copyBtn = document.getElementById("copyErrBtn");
  if (copyBtn) {
    if (isError && text) {
      copyBtn.classList.remove("hidden");
      copyBtn.onclick = () => {
        navigator.clipboard.writeText(text).then(() => {
          copyBtn.textContent = "✅ 已复制";
          setTimeout(() => { copyBtn.textContent = "复制错误信息"; }, 1500);
        });
      };
    } else {
      copyBtn.classList.add("hidden");
    }
  }
}

// 显示错误 + 复制按钮
function setErrorMsg(text) {
  setMsg(text, true);
  const copyBtn = document.getElementById("copyErrBtn");
  if (!copyBtn) return;
  copyBtn.classList.remove("hidden");
  copyBtn.onclick = () => {
    navigator.clipboard.writeText(text).then(() => {
      copyBtn.textContent = "✅ 已复制";
      setTimeout(() => { copyBtn.textContent = "复制错误信息"; }, 1500);
    });
  };
}

document.getElementById("openOurUrl").addEventListener("click", () => {
  chrome.tabs.create({ url: OUR_PAGE_URL });
});


document.getElementById("detect").addEventListener("click", async () => {
  const resultEl = document.getElementById("detect-result");
  resultEl.textContent = "正在检测页面结构，将自动切换下一题遍历各题…";
  resultEl.className = "message-area text-muted";
  resultEl.style.display = "block";

  const detectBtnIcon = document.querySelector("#detect .card-icon");
  const detectBtnTitle = document.querySelector("#detect .card-title");
  const detectBtnDesc = document.querySelector("#detect .card-desc");

  // 状态更新为分析中
  if (detectBtnIcon) {
    detectBtnIcon.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>`;
    detectBtnIcon.classList.add("icon-spin");
    detectBtnIcon.style.color = "var(--brand-green)";
    detectBtnIcon.style.background = "rgba(16, 185, 129, 0.1)";
  }
  if (detectBtnTitle) {
    detectBtnTitle.textContent = "分析中...";
    detectBtnTitle.style.color = "var(--brand-green)";
  }
  if (detectBtnDesc) detectBtnDesc.textContent = "正在遍历各题";

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) {
    resultEl.textContent = "无法获取当前标签页";
    resultEl.className = "message-area text-error";
    if (detectBtnIcon) {
      detectBtnIcon.classList.remove("icon-spin");
      detectBtnIcon.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg>`;
      detectBtnIcon.style.color = "var(--gray-600)";
      detectBtnIcon.style.background = "var(--gray-100)";
    }
    if (detectBtnTitle) {
      detectBtnTitle.textContent = "分析页面";
      detectBtnTitle.style.color = "var(--gray-800)";
    }
    if (detectBtnDesc) detectBtnDesc.textContent = "识别当前表单";
    return;
  }
  // 先发 PING 确认 content 已注入，避免误报「注入失败」
  try {
    await chrome.tabs.sendMessage(tab.id, { type: "PING" });
  } catch (pingErr) {
    resultEl.textContent = "无法连接录题页，请确保当前标签页是录题页并刷新该页后再点「开始检测」。";
    resultEl.classList.add("text-error");
    return;
  }
  resultEl.textContent = "正在遍历各题（约 10～30 秒），请勿关闭本弹窗…";
  const { selectors: stored } = await chrome.storage.sync.get("selectors");
  try {
    const res = await chrome.tabs.sendMessage(tab.id, { type: "DETECT_AND_WALK", selectors: stored || {} });
    
    if (detectBtnIcon) detectBtnIcon.classList.remove("icon-spin");

    if (!res || !res.ok) {
      resultEl.textContent = res?.error || "分析失败，请确认在录题页面";
      resultEl.className = "message-area text-error";
      if (detectBtnIcon) {
        detectBtnIcon.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg>`;
        detectBtnIcon.style.color = "var(--gray-600)";
        detectBtnIcon.style.background = "var(--gray-100)";
      }
      if (detectBtnTitle) {
        detectBtnTitle.textContent = "分析页面";
        detectBtnTitle.style.color = "var(--gray-800)";
      }
      if (detectBtnDesc) detectBtnDesc.textContent = "识别当前表单";
      return;
    }
    const { selectors, fields, message, walked, total, slots } = res;
    const walked_n = walked || "?";
    const total_n = total ? `/${total}` : "";
    const msg = message || `已遍历 ${walked_n}${total_n} 题，了解页面结构，允许上传 Word。`;
    await chrome.storage.sync.set({
      selectors,
      fields: fields || [],
      lastDetectMessage: msg,
      pageQuestionTotal: total != null && total > 0 ? total : null,
      pageSlots: slots && slots.length > 0 ? slots : null,
    });
    if (tab?.id) {
      chrome.storage.local.set({ fillTargetTabId: tab.id });
      if (total != null && total > 0) {
        chrome.runtime.sendMessage({ type: "REFRESH_BADGE", tabId: tab.id }).catch(() => {});
      }
    }
    resultEl.textContent = msg + `（已识别字段：${(fields || []).map(f => f.label || f.role).join("、")}）`;
    resultEl.className = "message-area text-success";
    resultEl.style.display = "block";
    
    // 更新按钮状态为已识别
    if (detectBtnIcon) {
      detectBtnIcon.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7"></path></svg>`;
      detectBtnIcon.style.color = "var(--brand-green)";
      detectBtnIcon.style.background = "rgba(16, 185, 129, 0.1)";
    }
    if (detectBtnTitle) {
      detectBtnTitle.textContent = "已分析";
      detectBtnTitle.style.color = "var(--brand-green)";
    }
    if (detectBtnDesc) {
      detectBtnDesc.textContent = `包含 ${walked_n} 道题`;
    }
    
    const hintEl = document.getElementById("uploadStepHint");
    if (hintEl) {
      hintEl.textContent = "已提取当前页面录题结构，可直接上传。";
      hintEl.style.color = "var(--brand-green)";
      hintEl.style.fontWeight = "500";
    }

    setTimeout(() => { resultEl.style.display = "none"; }, 5000);
  } catch (e) {
    const errMsg = String(e?.message || e);
    if (detectBtnIcon) {
      detectBtnIcon.classList.remove("icon-spin");
      detectBtnIcon.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg>`;
      detectBtnIcon.style.color = "var(--gray-600)";
      detectBtnIcon.style.background = "var(--gray-100)";
    }
    if (detectBtnTitle) {
      detectBtnTitle.textContent = "分析页面";
      detectBtnTitle.style.color = "var(--gray-800)";
    }
    if (detectBtnDesc) detectBtnDesc.textContent = "识别当前表单";

    if (/receiving end|establish connection|target closed/i.test(errMsg)) {
      resultEl.textContent = "连接中断（检测时间较长时请勿关闭弹窗），请刷新录题页后重新点「开始检测」。";
    } else {
      resultEl.textContent = "注入失败：" + (errMsg || "请刷新录题页面后再试");
    }
    resultEl.className = "message-area text-error";
  }
});

function getWordFiles(files) {
  if (!files || !files.length) return [];
  return Array.from(files).filter((f) => {
    const n = (f.name || "").toLowerCase();
    return n.endsWith(".docx") || n.endsWith(".doc");
  });
}

// ─── 记录发起上传时的录题页 tabId（避免解析期间切标签导致填充到错误页）
let targetTabId = null;

// ─── 消息区 ────────────────────────────────────────────────────────────────
// function setMsg() was moved up

// 题型 → 中文（含听说题型），复制/预览时使用
const QUESTION_TYPE_MAP = {
  single: "单选", multiple: "多选", judge: "判断", blank: "填空",
  listening_choice: "听后选择", listening_response: "听后应答", reading_aloud: "模仿朗读",
  listening_fill: "听后填空", listening_retell: "信息转述",
};

// ─── Drop Zone 状态机 ───────────────────────────────────────────────────────
// state: "idle" | "selected" | "parsing" | "done"
function copyQuestionListToClipboard(questions, typeMap) {
  const map = typeMap || QUESTION_TYPE_MAP;
  const lines = (questions || []).map((q, i) => {
    const t = map[q.type] || q.type || "?";
    const txt = (q.question || "").trim();
    return `${i + 1}. [${t}] ${txt}`;
  });
  const text = lines.join("\n");
  if (!text) return;
  navigator.clipboard.writeText(text).then(() => {
    setMsg("已复制 " + questions.length + " 题到剪贴板", false);
    setTimeout(() => setMsg("", false), 2000);
  }).catch(() => {
    setMsg("复制失败，请手动选择上方文字复制", true);
  });
}

/** 复制完整题目（题干、选项、答案、解析），便于发给他人调试 prompt */
function copyFullQuestionsToClipboard(questions, typeMap) {
  const map = typeMap || QUESTION_TYPE_MAP;
  const blocks = (questions || []).map((q, i) => {
    const t = map[q.type] || q.type || (q.type || "?");
    const lines = [
      `第 ${i + 1} 题 [${t}]`,
      "题干：" + (q.question || "").trim(),
    ];
    const opts = q.options;
    if (opts && Array.isArray(opts) && opts.length) {
      lines.push("选项：" + opts.join("  "));
    }
    if ((q.answer || "").toString().trim()) {
      lines.push("答案：" + (q.answer || "").toString().trim());
    }
    if ((q.explanation || "").toString().trim()) {
      lines.push("解析：" + (q.explanation || "").toString().trim());
    }
    return lines.join("\n");
  });
  const text = blocks.join("\n\n---\n\n");
  if (!text) return;
  navigator.clipboard.writeText(text).then(() => {
    setMsg("已复制完整题目（" + questions.length + " 题）到剪贴板，可粘贴到 Word 发给我调试", false);
    setTimeout(() => setMsg("", false), 3000);
  }).catch(() => {
    setMsg("复制失败，请手动选择上方文字复制", true);
  });
}

/** 复制题目为 JSON，便于调试或导入 */
function copyQuestionsJsonToClipboard(questions) {
  const text = JSON.stringify(questions || [], null, 2);
  if (!text) return;
  navigator.clipboard.writeText(text).then(() => {
    setMsg("已复制 JSON 到剪贴板", false);
    setTimeout(() => setMsg("", false), 2000);
  }).catch(() => {
    setMsg("复制失败", true);
  });
}

// ─── JSON 历史（最近 5 条，用于右下角历史图标下拉）────────────────────────
const JSON_HISTORY_MAX = 5;
async function getJsonHistory() {
  const { jsonHistory } = await chrome.storage.local.get("jsonHistory");
  return Array.isArray(jsonHistory) ? jsonHistory : [];
}
async function pushJsonHistory(jsonText) {
  if (!jsonText || typeof jsonText !== "string" || !jsonText.trim()) return;
  const list = await getJsonHistory();
  const trimmed = jsonText.trim();
  const next = { text: trimmed, time: Date.now() };
  const filtered = list.filter((item) => item.text !== trimmed);
  const nextList = [next, ...filtered].slice(0, JSON_HISTORY_MAX);
  await chrome.storage.local.set({ jsonHistory: nextList });
}

function setDropZoneState(state, data = {}) {
  const dz       = document.getElementById("dropZone");
  const dzN      = document.getElementById("dzNormal");
  const dzP      = document.getElementById("dzParsing");
  const dzD      = document.getElementById("dzDone");
  const upBtn    = document.getElementById("uploadAndFill");

  // 如果关键元素不在 DOM（弹窗在非录题页打开），静默跳过
  if (!dz || !dzN || !dzP || !dzD || !upBtn) return;

  // 重置所有子层
  dzN.classList.add("hidden");
  dzP.classList.add("hidden");
  dzD.classList.add("hidden");
  dz.className = "drop-zone";

  if (state === "idle") {
    dzN.classList.remove("hidden");
    dzN.querySelector(".drop-icon").innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>`;
    dzN.querySelector(".drop-text").textContent = "点击或拖拽 Word 文件";
    dzN.querySelector(".drop-subtext").textContent = "支持 .docx 格式，可多选";
    upBtn.style.opacity = "1";
    upBtn.style.pointerEvents = "auto";
    const upTitle = upBtn.querySelector(".card-title");
    if (upTitle) upTitle.textContent = "解析并填入";

  } else if (state === "selected") {
    dzN.classList.remove("hidden");
    dz.classList.add("has-files");
    const names = (data.files || []).map(f => f.name);
    dzN.querySelector(".drop-icon").innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>`;
    dzN.querySelector(".drop-text").textContent =
      names.length === 1 ? names[0] : `${names.length} 个文件`;
    dzN.querySelector(".drop-subtext").textContent =
      names.length > 1 ? names.join("、").slice(0, 50) : "点击下方「解析并填入」开始";
    upBtn.style.opacity = "1";
    upBtn.style.pointerEvents = "auto";
    const upTitle = upBtn.querySelector(".card-title");
    if (upTitle) upTitle.textContent = "解析并填入";

  } else if (state === "parsing") {
    dzP.classList.remove("hidden");
    dz.classList.add("is-parsing");
    upBtn.style.opacity = "0.6";
    upBtn.style.pointerEvents = "none";
    const upTitle = upBtn.querySelector(".card-title");
    if (upTitle) upTitle.textContent = "识别中...";
    const subEl = document.getElementById("dzParsingSubtext");
    if (subEl) subEl.textContent = "已等待 0 秒 · 关闭弹窗不会中断";
  } else if (state === "done") {
    dzD.classList.remove("hidden");
    dz.classList.add("is-done");
    // 填充结果摘要
    const qs = data.questions || [];
    const debugInfo = data.debug_info || null;
    document.getElementById("resultsTitle").textContent = `✅ 识别完成，共 ${qs.length} 题`;
    const typeMap = QUESTION_TYPE_MAP;
    const preview = qs.slice(0, 3).map((q) => {
      const t = typeMap[q.type] || q.type || "?";
      const txt = (q.question || "").slice(0, 30) + ((q.question || "").length > 30 ? "…" : "");
      return `<div class="result-item">[${t}] ${txt}</div>`;
    }).join("") + (qs.length > 3 ? `<div class="result-more">…还有 ${qs.length - 3} 题</div>` : "");
    document.getElementById("resultsPreview").innerHTML = preview;
    // 页面题数与识别题数不一致时提示
    const mismatchEl = document.getElementById("resultsMismatch");
    if (mismatchEl) {
      chrome.storage.sync.get("pageQuestionTotal", ({ pageQuestionTotal }) => {
        if (pageQuestionTotal != null && pageQuestionTotal > 0 && pageQuestionTotal !== qs.length) {
          mismatchEl.textContent = `⚠️ 页面仅 ${pageQuestionTotal} 题，识别出 ${qs.length} 题，请核对题目数量或重新上传。`;
          mismatchEl.classList.remove("hidden");
        } else {
          mismatchEl.classList.add("hidden");
        }
      });
    }
    // 「复制请求 Prompt」按钮
    const copyDebugBtn = document.getElementById("copyDebugInfo");
    if (copyDebugBtn) {
      if (debugInfo && debugInfo.system_prompt) {
        copyDebugBtn.classList.remove("hidden");
        copyDebugBtn.onclick = (e) => {
          e.stopPropagation();
          e.preventDefault();
          const text = "=== System Prompt ===\n\n" + debugInfo.system_prompt + "\n\n=== Word 原文 ===\n\n" + (debugInfo.user_content || "").slice(0, 50000);
          navigator.clipboard.writeText(text).then(() => {
            setMsg("已复制请求 Prompt 到剪贴板", false);
            setTimeout(() => setMsg("", false), 2000);
          }).catch(() => setMsg("复制失败", true));
        };
      } else {
        copyDebugBtn.classList.add("hidden");
      }
    }
    // 「复制返回 JSON」按钮 — 复制完整的所有题目 JSON，可用于高级设置调试
    const copyJsonBtn = document.getElementById("copyJson");
    if (copyJsonBtn) {
      copyJsonBtn.classList.remove("hidden");
      copyJsonBtn.onclick = (e) => {
        e.stopPropagation();
        e.preventDefault();
        const jsonStr = JSON.stringify(qs, null, 2);
        navigator.clipboard.writeText(jsonStr).then(() => {
          setMsg(`已复制 ${qs.length} 题完整 JSON 到剪贴板`, false);
          setTimeout(() => setMsg("", false), 2000);
        }).catch(() => setMsg("复制失败", true));
      };
    }

    // 恢复「解析并填入」卡片为「重新解析」，避免还显示「识别中...」
    upBtn.style.opacity = "1";
    upBtn.style.pointerEvents = "auto";
    const upTitleDone = upBtn.querySelector(".card-title");
    const upDescDone = upBtn.querySelector(".card-desc");
    if (upTitleDone) upTitleDone.textContent = "重新解析";
    if (upDescDone) upDescDone.textContent = "上传新文件重新识别";
  }
}

// ─── 将 File 读取为 base64 ─────────────────────────────────────────────────
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload  = (e) => resolve({
      name: file.name,
      mimeType: file.type || "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      base64: e.target.result.split(",")[1],
    });
    reader.onerror = () => reject(new Error(`读取文件失败：${file.name}`));
    reader.readAsDataURL(file);
  });
}

/**
 * 将「小题」列表规范为「大题」列表：连续的同题型 listening_fill 合并为一条，用 blanks 表示多空，
 * 这样录题页上一题多空时能填满所有输入框。
 * @param {Array} questions - 解析结果（可能多条独立小题）
 * @param {number|null} pageTotal - 录题页题数（可选）
 * @param {Array|null} pageSlots - 页面分析时每题的 subCount 信息（优先用于精确合并）
 * @returns {Array} 按大题组织的题目，适合 FILL_FORM
 */
function normalizeQuestionsToSlots(questions, pageTotal, pageSlots) {
  if (!questions || questions.length === 0) return questions;
  const noBlanks = (x) => !x.blanks || !Array.isArray(x.blanks) || x.blanks.length === 0;
  const hasBlanks = (x) => x.blanks && Array.isArray(x.blanks) && x.blanks.length > 0;

  // ── 优先路径：pageSlots 已知每题 subCount，直接按 subCount 取对应数量的 AI 题合并 ──
  if (pageSlots && pageSlots.length > 0) {
    const out = [];
    let qi = 0; // 当前消耗到第几个 AI 题
    for (let si = 0; si < pageSlots.length && qi < questions.length; si++) {
      const slot = pageSlots[si];
      const need = (slot.subCount && slot.subCount > 1) ? slot.subCount : 1;
      // 已有 blanks 的题直接占 1 槽
      if (hasBlanks(questions[qi])) {
        out.push(questions[qi]);
        qi++;
        continue;
      }
      if (need === 1) {
        out.push(questions[qi]);
        qi++;
      } else {
        // 取 need 道 AI 题合并成一道带 blanks 的大题
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
    // 剩余未消耗的 AI 题追加（AI 比页面多的情况）
    while (qi < questions.length) { out.push(questions[qi]); qi++; }
    return out;
  }

  // ── 兜底路径：无 pageSlots，按 listening_script / 题型推测合并 ──
  const out = [];
  let i = 0;
  while (i < questions.length) {
    const q = questions[i];
    // 已有 blanks 的直接保留
    if (hasBlanks(q)) {
      out.push(q);
      i++;
      continue;
    }
    // 连续 listening_fill 小题合并为一大题
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
    // 连续同类型、同 listening_script（或同一段被拆分的对话）的小题合并
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
        j < questions.length &&
        questions[j].type === q.type &&
        noBlanks(questions[j]) &&
        questions[j].listening_script &&
        sameConversation(groupScript, normScript(questions[j].listening_script))
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
          type: q.type,
          listening_script: fullScript,
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
        i = j;
        continue;
      }
    }
    out.push(q);
    i++;
  }
  return out;
}

// ─── 触发填充 ──────────────────────────────────────────────────────────────
async function doFill(questions) {
  if (!questions || questions.length === 0) {
    setMsg("没有可填入的题目", true);
    return;
  }

  // 优先用记录的 tabId，重开弹窗时用当前活跃 tab
  let tabId = targetTabId;
  if (!tabId) {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    tabId = tab?.id;
  }
  if (!tabId) {
    setMsg("无法确定录题页，请确认页面已打开", true);
    return;
  }

  const { selectors, pageQuestionTotal, pageSlots, defaultAudioUrl, defaultImageUrl } = await chrome.storage.sync.get(["selectors", "pageQuestionTotal", "pageSlots", "defaultAudioUrl", "defaultImageUrl"]);
  // 检查是否已做过页面检测
  if (!selectors || !Object.values(selectors).some(Boolean)) {
    setMsg("请先点击上方「分析页面」提取题库结构", true);
    return;
  }

  // 先将小题合并为大题（优先用 pageSlots 的 subCount 精确合并，无则按 listening_script 推测）
  const pageTotal = pageQuestionTotal != null && pageQuestionTotal > 0 ? pageQuestionTotal : null;
  let toFill = normalizeQuestionsToSlots(questions, pageTotal != null ? pageTotal : null, pageSlots || null);
  let trimMsg = "";

  if (pageTotal != null) {
    if (toFill.length > pageTotal) {
      const mergedCount = toFill.length;
      toFill = toFill.slice(0, pageTotal);
      trimMsg = `解析出 ${questions.length} 小题，合并后 ${mergedCount} 大题，页面共 ${pageTotal} 题，已按模板填入前 ${pageTotal} 题。`;
    } else if (toFill.length < pageTotal) {
      const mergedCount = toFill.length;
      const empty = { type: "", question: "", options: [], answer: "", explanation: "", listening_script: "" };
      toFill = [...toFill, ...Array.from({ length: pageTotal - mergedCount }, () => ({ ...empty }))];
      trimMsg = `解析出 ${questions.length} 小题，合并后 ${mergedCount} 大题，页面共 ${pageTotal} 题，已填入并补齐 ${pageTotal - mergedCount} 道空题。`;
    }
  }

  setMsg(trimMsg || `共 ${toFill.length} 题，正在填入录题页…`, false);

  const historyBtn = document.getElementById("jsonHistoryBtn");
  if (historyBtn) {
    historyBtn.disabled = true;
    historyBtn.style.opacity = "0.6";
    historyBtn.style.pointerEvents = "none";
  }

  const restoreFillBtn = () => {
    if (historyBtn) {
      historyBtn.disabled = false;
      historyBtn.style.opacity = "";
      historyBtn.style.pointerEvents = "";
    }
  };

  try {
    const result = await chrome.tabs.sendMessage(tabId, {
      type: "FILL_FORM",
      questions: toFill,
      selectors,
      defaultAudioUrl: (defaultAudioUrl || "").trim() || undefined,
      defaultImageUrl: (defaultImageUrl || "").trim() || undefined,
      debugSource: "parse",
    });
    restoreFillBtn();
    if (result?.ok === false) {
      setMsg((result.error || "填充出错") + "。可先点上方「复制题目列表」或「复制完整题目」把识别结果发给我排查。", true);
      return;
    }
    const done = trimMsg || result?.message || `填充完成，共 ${result?.filled ?? toFill.length} 题。`;
    setMsg(done, false);
  } catch (e) {
    restoreFillBtn();
    if (String(e).includes("Could not establish connection") || String(e).includes("receiving end does not exist")) {
      setMsg("录题页未响应，请刷新该页面后再试。可先点上方「复制题目列表」或「复制完整题目」把识别结果发给我排查。", true);
    } else {
      setMsg(`填充失败：${e.message || e}。可先点上方「复制题目列表」或「复制完整题目」把识别结果发给我排查。`, true);
    }
  }
}

// ─── 上传 Word → 发给 background worker（弹窗关闭不中断）───────────────────
async function doUploadAndFill(filesToUse) {
  if (!filesToUse || filesToUse.length === 0) {
    setMsg("请选择或拖入至少一个 .docx 文件", true);
    return;
  }

  const { selectors } = await chrome.storage.sync.get("selectors");
  const hasStructure = selectors && Object.values(selectors).some(Boolean);
  if (!hasStructure) {
    setMsg("请先点击「分析页面」提取题库结构，再解析 Word", true);
    return;
  }

  setMsg("", false);

  // 记录当前 tabId 作为填充目标
  const [curTab] = await chrome.tabs.query({ active: true, currentWindow: true });
  targetTabId = curTab?.id || null;

  let filesData;
  try {
    filesData = await Promise.all(Array.from(filesToUse).map(fileToBase64));
  } catch (e) {
    setMsg(e.message || "文件读取失败", true);
    return;
  }

  setDropZoneState("parsing");

  chrome.runtime.sendMessage({ type: "START_PARSE", filesData, tabId: targetTabId || curTab?.id }).catch(() => {
    setDropZoneState("idle");
    setMsg("无法连接后台，请在 chrome://extensions 重新加载插件", true);
  });
}

// 更新 footer「仅填入」按钮旁的绿点：有缓存则显示，否则隐藏
async function updateFillCacheDot() {
  const dot = document.getElementById("fillCacheDot");
  if (!dot) return;
  const { parseState } = await chrome.storage.local.get("parseState");
  const hasCache = parseState && parseState.status === "done" && parseState.questions && parseState.questions.length > 0;
  dot.classList.toggle("hidden", !hasCache);
}

// ─── 弹窗打开时从 storage 恢复上次任务状态 ────────────────────────────────
async function restoreParseState() {
  const { parseState } = await chrome.storage.local.get("parseState");
  await updateFillCacheDot();
  if (!parseState || parseState.status === "idle") return;

  if (parseState.status === "parsing") {
    setDropZoneState("parsing");
  } else if (parseState.status === "done") {
    setDropZoneState("done", { questions: parseState.questions, debug_info: parseState.debug_info });
    setMsg("上次识别结果仍可填入，或重新上传新文件。", false);
  } else if (parseState.status === "error") {
    setMsg(parseState.text, true);
  }
}

document.getElementById("uploadAndFill").addEventListener("click", async () => {
  // 检查是否已做过页面检测
  const { selectors } = await chrome.storage.sync.get("selectors");
  const hasStructure = selectors && Object.values(selectors).some(Boolean);
  if (!hasStructure) {
    setMsg("请先点击上方「分析页面」提取题库结构，再解析 Word", true);
    const btnDetect = document.getElementById("detect");
    if (btnDetect) {
      btnDetect.style.transform = "scale(1.02)";
      btnDetect.style.boxShadow = "0 0 0 2px var(--brand-green)";
      setTimeout(() => {
        btnDetect.style.transform = "none";
        btnDetect.style.boxShadow = "none";
      }, 300);
    }
    return;
  }

  // done 状态下按钮文字是"重新上传"，点击先重置再打开文件选择
  if (document.getElementById("dropZone").classList.contains("is-done")) {
    setDropZoneState("idle");
    document.getElementById("wordFiles").value = "";
    document.getElementById("wordFiles").click();
    return;
  }
  const list = getWordFiles(document.getElementById("wordFiles").files);
  
  if (!list.length) {
    document.getElementById("wordFiles").click();
    return;
  }
  
  doUploadAndFill(list);
});

// ─── Drop Zone 事件 ────────────────────────────────────────────────────────
const dropZone = document.getElementById("dropZone");
if (dropZone) {
  // 点击 drop zone：仅在非识别中状态下打开文件选择；点击按钮（如复制、填入）不触发
  dropZone.addEventListener("click", async (e) => {
    if (dropZone.classList.contains("is-parsing")) return;
    if (e.target.closest("button") || e.target.closest(".results-actions") || e.target.closest(".btn-row")) return;

    // 先检查是否已经分析过页面结构
    const { selectors } = await chrome.storage.sync.get("selectors");
    const hasStructure = selectors && Object.values(selectors).some(Boolean);
    if (!hasStructure) {
      setMsg("请先点击上方「分析页面」提取题库结构", true);
      const btnDetect = document.getElementById("detect");
      if (btnDetect) {
        btnDetect.style.transform = "scale(1.02)";
        btnDetect.style.boxShadow = "0 0 0 2px var(--brand-green)";
        setTimeout(() => {
          btnDetect.style.transform = "none";
          btnDetect.style.boxShadow = "none";
        }, 300);
      }
      return;
    }

    document.getElementById("wordFiles").click();
  });

  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (!dropZone.classList.contains("is-parsing")) {
      dropZone.classList.add("drag-over");
    }
  });
  dropZone.addEventListener("dragleave", (e) => {
    e.preventDefault();
    dropZone.classList.remove("drag-over");
  });
  dropZone.addEventListener("drop", async (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.remove("drag-over");
    if (dropZone.classList.contains("is-parsing")) return; // 识别中禁止拖入

    // 先检查是否已经分析过页面结构
    const { selectors } = await chrome.storage.sync.get("selectors");
    const hasStructure = selectors && Object.values(selectors).some(Boolean);
    if (!hasStructure) {
      setMsg("请先点击上方「分析页面」提取题库结构", true);
      const btnDetect = document.getElementById("detect");
      if (btnDetect) {
        btnDetect.style.transform = "scale(1.02)";
        btnDetect.style.boxShadow = "0 0 0 2px var(--brand-green)";
        setTimeout(() => {
          btnDetect.style.transform = "none";
          btnDetect.style.boxShadow = "none";
        }, 300);
      }
      return;
    }

    const list = getWordFiles(e.dataTransfer.files);
    if (list.length) {
      setDropZoneState("selected", { files: list });
      doUploadAndFill(list);
    }
  });
}

document.getElementById("wordFiles").addEventListener("change", function () {
  const list = getWordFiles(this.files);
  if (list.length) {
    setDropZoneState("selected", { files: list });
    setMsg("", false);
  }
});

// ─── 取消识别按钮 ──────────────────────────────────────────────────────────
document.getElementById("cancelParse").addEventListener("click", () => {
  chrome.runtime.sendMessage({ type: "CANCEL_PARSE" });
  setDropZoneState("idle");
  setMsg("识别已取消，可重新上传文件。", false);
});

// ─── 重置按钮：清除识别结果 ─────────────────────────────────────────────────
document.getElementById("resetParseState").addEventListener("click", () => {
  chrome.runtime.sendMessage({ type: "CLEAR_PARSE_STATE" });
  setDropZoneState("idle");
  document.getElementById("wordFiles").value = "";
  setMsg("已重置，可重新上传文件。", false);
});

// ─── 右下角历史图标：点击展开最近 5 条 JSON，点击某条即填充 ─────────────────
const jsonHistoryBtn = document.getElementById("jsonHistoryBtn");
const jsonHistoryDropdown = document.getElementById("jsonHistoryDropdown");
if (jsonHistoryBtn && jsonHistoryDropdown) {
  jsonHistoryBtn.addEventListener("click", async (e) => {
    e.stopPropagation();
    const isOpen = !jsonHistoryDropdown.classList.contains("hidden");
    if (isOpen) {
      jsonHistoryDropdown.classList.add("hidden");
      return;
    }
    const list = await getJsonHistory();
    jsonHistoryDropdown.innerHTML = "";
    if (list.length === 0) {
      jsonHistoryDropdown.classList.remove("hidden");
      return;
    }
    for (const item of list) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "json-history-item";
      let preview = item.text.slice(0, 60);
      if (item.text.length > 60) preview += "…";
      const timeStr = item.time ? new Date(item.time).toLocaleString("zh-CN", { month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" }) : "";
      btn.innerHTML = `<span class="item-preview">${escapeHtml(preview)}</span><span class="item-time">${escapeHtml(timeStr)}</span>`;
      btn.addEventListener("click", async (ev) => {
        ev.preventDefault();
        jsonHistoryDropdown.classList.add("hidden");
        let questions;
        try {
          const parsed = JSON.parse(item.text);
          questions = parseApiResponseToQuestions(parsed);
          if (!questions && Array.isArray(parsed) && parsed.length > 0) questions = parsed;
        } catch (_) {
          setMsg("该条历史 JSON 格式错误，无法填充", true);
          return;
        }
        if (!Array.isArray(questions) || questions.length === 0) {
          setMsg("该条历史无有效题目数据", true);
          return;
        }
        const ta = document.getElementById("json");
        if (ta) ta.value = item.text;
        doFill(questions);
      });
      jsonHistoryDropdown.appendChild(btn);
    }
    jsonHistoryDropdown.classList.remove("hidden");
  });
  document.addEventListener("click", () => {
    if (!jsonHistoryDropdown.classList.contains("hidden")) {
      jsonHistoryDropdown.classList.add("hidden");
    }
  });
  jsonHistoryDropdown.addEventListener("click", (e) => e.stopPropagation());
}
function escapeHtml(s) {
  const div = document.createElement("div");
  div.textContent = s;
  return div.innerHTML;
}

document.getElementById("fill").addEventListener("click", async () => {
  const ta = document.getElementById("json");
  const msg = document.getElementById("json-msg");

  msg.textContent = "";
  msg.classList.remove("text-error");
  let questions;
  try {
    const parsed = JSON.parse(ta.value.trim());
    questions = parseApiResponseToQuestions(parsed);
    if (!questions && Array.isArray(parsed) && parsed.length > 0) questions = parsed;
  } catch (e) {
    msg.textContent = "JSON 格式错误（请检查是否完整、未截断，末尾勿多逗号）";
    msg.classList.add("text-error");
    return;
  }
  if (!Array.isArray(questions) || questions.length === 0) {
    msg.textContent = "请粘贴题目数组，或控制台接口返回的完整 JSON（含 data.topic）";
    msg.classList.add("text-error");
    return;
  }
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) {
    msg.textContent = "无法获取当前标签页";
    msg.classList.add("text-error");
    return;
  }
  const sel = await getSelectorsForJsonPanel(tab.id);
  const { defaultAudioUrl: dau, defaultImageUrl: diu } = await chrome.storage.sync.get(["defaultAudioUrl", "defaultImageUrl"]);
  try {
    await chrome.tabs.sendMessage(tab.id, {
      type: "FILL_FORM",
      questions,
      selectors: sel,
      defaultAudioUrl: (dau || "").trim() || undefined,
      defaultImageUrl: (diu || "").trim() || undefined,
      debugSource: "json",
    });
    pushJsonHistory(ta.value.trim()).catch(() => {});
    msg.textContent = "已发送 " + questions.length + " 题";
    msg.classList.add("text-success");
  } catch (e) {
    msg.textContent = "注入失败，请刷新录题页面后重试";
    msg.classList.add("text-error");
  }
});

document.getElementById("fillDebug").addEventListener("click", async () => {
  const ta = document.getElementById("json");
  const msg = document.getElementById("json-msg");
  msg.textContent = "";
  msg.classList.remove("text-error", "text-success");
  let questions;
  try {
    const parsed = JSON.parse(ta.value.trim());
    questions = parseApiResponseToQuestions(parsed);
    if (!questions && Array.isArray(parsed) && parsed.length > 0) questions = parsed;
  } catch (e) {
    msg.textContent = "JSON 格式错误（请检查是否完整、未截断，末尾勿多逗号）";
    msg.classList.add("text-error");
    return;
  }
  if (!Array.isArray(questions) || questions.length === 0) {
    msg.textContent = "请先粘贴题目数组或接口完整 JSON（含 data.topic）";
    msg.classList.add("text-error");
    return;
  }
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) {
    msg.textContent = "无法获取当前标签页";
    msg.classList.add("text-error");
    return;
  }
  const sel = await getSelectorsForJsonPanel(tab.id);
  try {
    const res = await chrome.tabs.sendMessage(tab.id, { type: "FILL_FORM_DEBUG", questions, selectors: sel });
    if (!res?.ok || !res.report) {
      msg.textContent = res?.error || "调试返回异常";
      msg.classList.add("text-error");
      return;
    }
    const r = res.report;
    const lines = [r.hint || "", "当前页合并后的选择器 curSel:", JSON.stringify(r.curSel, null, 2)];
    (r.roles || []).forEach((item) => {
      lines.push(`\n--- ${item.role} ---`);
      lines.push(`取值预览: ${item.valuePreview || "(空)"} (长度 ${item.valueLength})`);
      lines.push(`结论: ${item.summary}`);
      (item.selectors || []).forEach((c) => {
        lines.push(`  ${c.found ? "✓" : "✗"} ${c.selector}`);
      });
    });
    const text = lines.join("\n");
    await navigator.clipboard.writeText(text);
    msg.classList.add("text-success");
    msg.style.whiteSpace = "pre-wrap";
    msg.style.fontSize = "11px";
    msg.textContent = "已复制调试报告到剪贴板（未找到元素会标 ✗）。预览：\n\n" + text.slice(0, 1800) + (text.length > 1800 ? "\n\n…(完整在剪贴板，可粘贴发给我)" : "");
  } catch (e) {
    msg.textContent = "调试失败: " + (e?.message || e) + "。请确保在录题页并刷新后重试。";
    msg.classList.add("text-error");
  }
});

/** 把控制台粘贴的「完整接口 JSON」转成填充用的题目数组。当前只兼容你给的格式：{ status: 1, data: { topic: {...} } }，选项用 optionDesc/isTrue/option。 */
function parseApiResponseToQuestions(parsed) {
  if (!parsed || typeof parsed !== "object") return null;
  if (Array.isArray(parsed) && parsed.length > 0) {
    const first = parsed[0];
    if (first && (first.topicContent != null || first.topicOption != null)) {
      return parsed.map((t) => topicFromApiTopic(t)).filter(Boolean);
    }
    if (first && (first.question != null || first.options != null)) return parsed;
  }
  const data = parsed.data;
  if (!data) return null;
  const topic = data.topic;
  if (topic) {
    const one = topicFromApiTopic(topic);
    return one ? [one] : null;
  }
  if (Array.isArray(data)) {
    const list = data.map((t) => topicFromApiTopic(t)).filter(Boolean);
    return list.length ? list : null;
  }
  return null;
}

function topicFromApiTopic(topic) {
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

// ─── 生成 AI 填写模板 ──────────────────────────────────────────────────────
const ROLE_LABELS = {
  question: "题干内容",
  option_a: "选项A内容",
  option_b: "选项B内容",
  option_c: "选项C内容",
  option_d: "选项D内容",
  answer: "正确答案（单选填 A/B/C/D，判断填 对/错）",
  explanation: "解析内容（可为空）",
  listening_script: "听力原文（听力题才需要，其他题留空字符串）",
  audio_url: "音频 URL（听力题，一般留空）",
  image_url: "图片 URL（图片题，一般留空）",
};

async function generateAiTemplate() {
  const templateMsgEl = document.getElementById("template-msg");
  if (templateMsgEl) { templateMsgEl.textContent = "正在检测页面字段…"; templateMsgEl.className = "message-area text-muted"; }

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) {
    if (templateMsgEl) { templateMsgEl.textContent = "无法获取当前标签页"; templateMsgEl.className = "message-area text-error"; }
    return;
  }

  let fields = [];
  try {
    const res = await chrome.tabs.sendMessage(tab.id, { type: "DETECT_FORM" });
    if (res?.ok && res.fields && res.fields.length > 0) {
      // 过滤掉属性类字段，保留填题内容字段（包括动态 blank_* 字段）
      const SKIP = new Set(["grade", "course", "unit", "knowledge_point", "difficulty", "question_permission", "recorder", "audio_file", "image_file", "submit_btn", "next_btn"]);
      fields = res.fields.filter(f => !SKIP.has(f.role));
    }
  } catch (_) {}

  // 若未检测到字段，用默认常见字段集
  if (fields.length === 0) {
    fields = ["question", "option_a", "option_b", "option_c", "option_d", "answer", "explanation"].map(r => ({ role: r }));
  }

  const roleList = fields.map(f => `  "${f.role}": "${ROLE_LABELS[f.role] || f.label || f.role}"`);
  const jsonTemplate = fields.reduce((obj, f) => { obj[f.role] = ""; return obj; }, {});

  const prompt = [
    "你是一个题目录入助手。请根据我提供的题目内容，提取对应字段，严格输出 JSON 数组格式（每道题为一个对象），不要有任何额外解释。",
    "",
    "字段说明：",
    "{",
    ...roleList,
    "}",
    "",
    "输出格式示例（可有多题）：",
    JSON.stringify([jsonTemplate], null, 2),
    "",
    "注意事项：",
    "- 题干、选项保留原文，不要添加「A.」「选项A：」等前缀",
    "- answer 字段只填字母（A/B/C/D），不含选项内容",
    "- 没有的字段填空字符串 \"\"，不要省略字段名",
    "- 如果有多道题，输出多个对象",
    "",
    "题目内容如下（请在此行之后粘贴你的题目）：",
    "---",
  ].join("\n");

  try {
    await navigator.clipboard.writeText(prompt);
    if (templateMsgEl) {
      templateMsgEl.textContent = `已复制 AI 提示模板（含 ${fields.length} 个字段）。请粘贴到 AI 对话框末尾，再接着粘贴你的题目内容，发给 AI 后将返回的 JSON 粘贴到下方文本框。`;
      templateMsgEl.className = "message-area text-success";
      setTimeout(() => { if (templateMsgEl) { templateMsgEl.textContent = ""; templateMsgEl.className = "message-area"; } }, 12000);
    }
  } catch (_) {
    if (templateMsgEl) { templateMsgEl.textContent = "复制失败，请检查浏览器剪贴板权限"; templateMsgEl.className = "message-area text-error"; }
  }
}

document.getElementById("generateAiTemplate").addEventListener("click", generateAiTemplate);

/** 高级 JSON 面板用：优先用当前页 DETECT_FORM 得到的选择器，否则用 storage 或默认，保证「获取→粘贴→填充」用同一套选择器 */
async function getSelectorsForJsonPanel(tabId) {
  try {
    const res = await chrome.tabs.sendMessage(tabId, { type: "DETECT_FORM" });
    if (res?.ok && res.selectors && Object.keys(res.selectors).length > 0) return res.selectors;
  } catch (_) {}
  const { selectors } = await chrome.storage.sync.get("selectors");
  return selectors || getDefaultSelectors();
}

function getDefaultSelectors() {
  return {
    question: "input[name=question], #question, [data-field=question], textarea[name=content]",
    option_a: "input[name=optionA], #optionA, [data-field=optionA]",
    option_b: "input[name=optionB], #optionB, [data-field=optionB]",
    option_c: "input[name=optionC], #optionC, [data-field=optionC]",
    option_d: "input[name=optionD], #optionD, [data-field=optionD]",
    answer: "input[name=answer], #answer, [data-field=answer]",
    explanation: "textarea[name=explanation], #explanation, [data-field=explanation]",
    submit_btn: "button[type=submit], input[type=submit], .btn-submit, [data-action=submit]",
    next_btn: "",
  };
}

// ─── JSON 粘贴框持久化（popup 关闭后内容不丢失）─────────────────────────────
(async () => {
  const ta = document.getElementById("json");
  if (!ta) return;
  // 恢复上次内容
  try {
    const { savedJsonText } = await chrome.storage.local.get("savedJsonText");
    if (savedJsonText) ta.value = savedJsonText;
  } catch (_) {}
  // 实时保存（debounce 800ms）
  let _saveTimer;
  ta.addEventListener("input", () => {
    clearTimeout(_saveTimer);
    _saveTimer = setTimeout(() => {
      chrome.storage.local.set({ savedJsonText: ta.value });
    }, 800);
  });
})();

// UI Toggles
const detectBtn = document.getElementById("detect");

if (detectBtn) {
  detectBtn.addEventListener("click", async () => {
    const resultEl = document.getElementById("detect-result");
    resultEl.style.display = "block";
    resultEl.textContent = "正在分析页面结构，请稍候...";
    resultEl.className = "message-area text-muted";

    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id) {
      resultEl.textContent = "无法获取当前标签页";
      resultEl.className = "message-area text-error";
      return;
    }

    const detectBtnIcon = document.querySelector("#detect .card-icon");
    const detectBtnTitle = document.querySelector("#detect .card-title");
    const detectBtnDesc = document.querySelector("#detect .card-desc");

    // 状态更新为分析中
    if (detectBtnIcon) {
      detectBtnIcon.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>`;
      detectBtnIcon.classList.add("icon-spin");
      detectBtnIcon.style.color = "var(--brand-green)";
      detectBtnIcon.style.background = "rgba(16, 185, 129, 0.1)";
    }
    if (detectBtnTitle) {
      detectBtnTitle.textContent = "分析中...";
      detectBtnTitle.style.color = "var(--brand-green)";
    }
    if (detectBtnDesc) detectBtnDesc.textContent = "正在遍历各题";
    
    try {
      await chrome.tabs.sendMessage(tab.id, { type: "PING" });
    } catch (e) {
      resultEl.textContent = "未连接页面，请刷新录题页后重试";
      resultEl.className = "message-area text-error";
      if (detectBtnIcon) {
        detectBtnIcon.classList.remove("icon-spin");
        detectBtnIcon.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg>`;
        detectBtnIcon.style.color = "var(--gray-600)";
        detectBtnIcon.style.background = "var(--gray-100)";
      }
      if (detectBtnTitle) {
        detectBtnTitle.textContent = "分析页面";
        detectBtnTitle.style.color = "var(--gray-800)";
      }
      if (detectBtnDesc) detectBtnDesc.textContent = "识别当前表单";
      return;
    }

    const { selectors: stored } = await chrome.storage.sync.get("selectors");
    try {
      const res = await chrome.tabs.sendMessage(tab.id, { type: "DETECT_AND_WALK", selectors: stored || {} });

      if (!res || !res.ok) {
        resultEl.textContent = res?.error || "分析失败，请确认在录题页面";
        resultEl.className = "message-area text-error";
        if (detectBtnIcon) {
          detectBtnIcon.classList.remove("icon-spin");
          detectBtnIcon.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg>`;
          detectBtnIcon.style.color = "var(--gray-600)";
          detectBtnIcon.style.background = "var(--gray-100)";
        }
        if (detectBtnTitle) {
          detectBtnTitle.textContent = "分析页面";
          detectBtnTitle.style.color = "var(--gray-800)";
        }
        if (detectBtnDesc) detectBtnDesc.textContent = "识别当前表单";
        return;
      }

      if (detectBtnIcon) detectBtnIcon.classList.remove("icon-spin");

    const { selectors, fields, message, walked, total, slots } = res;
    await chrome.storage.sync.set({
      selectors,
      fields: fields || [],
      lastDetectMessage: message,
      pageQuestionTotal: total != null && total > 0 ? total : null,
      pageSlots: slots && slots.length > 0 ? slots : null,
    });

      resultEl.textContent = `成功分析 ${walked} 题，已识别字段：${(fields || []).map(f => f.label || f.role).join("、")}`;
      resultEl.className = "message-area text-success";
      resultEl.style.display = "block";
      
      // 更新按钮状态为已识别
      if (detectBtnIcon) {
        detectBtnIcon.classList.remove("icon-spin");
        detectBtnIcon.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7"></path></svg>`;
        detectBtnIcon.style.color = "var(--brand-green)";
        detectBtnIcon.style.background = "rgba(16, 185, 129, 0.1)";
      }
      if (detectBtnTitle) {
        detectBtnTitle.textContent = "已分析";
        detectBtnTitle.style.color = "var(--brand-green)";
      }
      if (detectBtnDesc) {
        detectBtnDesc.textContent = `包含 ${walked} 道题`;
      }
      
      const hintEl = document.getElementById("uploadStepHint");
      if (hintEl) {
        hintEl.textContent = "已提取当前页面录题结构，可直接上传。";
        hintEl.style.color = "var(--brand-green)";
        hintEl.style.fontWeight = "500";
      }

      setTimeout(() => { resultEl.style.display = "none"; }, 5000);
    } catch (e) {
      if (detectBtnIcon) {
        detectBtnIcon.classList.remove("icon-spin");
        detectBtnIcon.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg>`;
        detectBtnIcon.style.color = "var(--gray-600)";
        detectBtnIcon.style.background = "var(--gray-100)";
      }
      if (detectBtnTitle) {
        detectBtnTitle.textContent = "分析页面";
        detectBtnTitle.style.color = "var(--gray-800)";
      }
      if (detectBtnDesc) detectBtnDesc.textContent = "识别当前表单";
      
      resultEl.textContent = "分析中断：" + e.message;
      resultEl.className = "message-area text-error";
    }
  });
}

