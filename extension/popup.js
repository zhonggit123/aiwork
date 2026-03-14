const BACKEND = "http://127.0.0.1:8766";
const OUR_PAGE_URL = "http://127.0.0.1:8766";
const onlyWhenDetected = document.getElementById("onlyWhenDetected");
const whenNotDetected = document.getElementById("whenNotDetected");
const statusBadge = document.getElementById("statusBadge");

// 检测是否为侧边栏模式（通过窗口高度判断，侧边栏通常比 popup 高很多）
(function detectSidePanelMode() {
  const isSidePanel = window.innerHeight > 600 || window.location.search.includes("sidepanel");
  if (isSidePanel) {
    document.documentElement.classList.add("sidepanel-mode");
    document.body.classList.add("sidepanel-mode");
  }
  // 监听窗口大小变化
  window.addEventListener("resize", () => {
    if (window.innerHeight > 600) {
      document.documentElement.classList.add("sidepanel-mode");
      document.body.classList.add("sidepanel-mode");
    }
  });
})();

/**
 * 存入 storage.sync 前精简 slots，只保留填充必需的字段，避免超过 8KB/item 配额。
 * currentSlotFields / sectionLabels / subQuestions 等仅在构建 prompt 时有用，无需持久化。
 */
function slimSlotsForStorage(slots) {
  if (!slots) return null;
  return slots.map(s => ({
    subCount:  s.subCount,
    typeHint:  s.typeHint  || undefined,
    typeCode:  s.typeCode  || undefined,
    hasAudio:  s.hasAudio  || undefined,
    hasImage:  s.hasImage  || undefined,
  }));
}

/** 数字转中文序号（1→一，2→二...15→十五） */
function toChineseNumeral(n) {
  const digits = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"];
  if (n <= 10) return digits[n];
  if (n < 20) return "十" + (n % 10 === 0 ? "" : digits[n % 10]);
  if (n < 100) return digits[Math.floor(n / 10)] + "十" + (n % 10 === 0 ? "" : digits[n % 10]);
  return String(n);
}

/** 分析页面卡片图标：放大镜（默认与已分析态） */
const DETECT_ICON_MAGNIFIER = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>`;

/**
 * 根据已识别字段列表生成「题型与结构」的友好描述，用于遍历进度展示（识别中）。
 * @param {string[]} fields - 字段名数组（如 audio_url, option_a, blank_audio_1 ...）
 * @returns {string} 简短描述，如「听力填空 · 5 个空 · 四选一 · 含答案与解析」
 */
function describeStructureFromFields(fields) {
  if (!Array.isArray(fields) || fields.length === 0) return "正在识别题型与结构…";
  const set = new Set(fields);

  const has = (name) => set.has(name);
  const countBlanks = (key) => {
    let n = 0;
    for (const f of fields) if (f && typeof f === "string" && f.startsWith(key)) n = Math.max(n, parseInt(f.replace(/\D/g, ""), 10) || 1);
    return n;
  };

  const parts = [];
  const hasListening = has("audio_url") || has("audio_file") || has("listening_script");
  const blankCount = Math.max(countBlanks("blank_audio_"), countBlanks("blank_script_"), countBlanks("blank_question_"));
  if (blankCount > 0) {
    parts.push(hasListening ? "听力填空" : "填空");
    parts.push(`${blankCount} 个空`);
  } else if (hasListening) parts.push("听力题");
  if (has("question") || has("keyword")) { if (parts.length === 0) parts.push("题干与关键词"); }

  const opts = [has("option_a"), has("option_b"), has("option_c"), has("option_d")].filter(Boolean).length;
  if (opts === 4) parts.push("四选一"); else if (opts === 3) parts.push("三选一"); else if (opts >= 1) parts.push(`${opts} 个选项`);
  if (has("answer")) parts.push("含答案");
  if (has("explanation")) parts.push("含解析");
  if (["grade", "course", "unit", "difficulty", "knowledge_point"].some(has)) parts.push("含元信息");

  return parts.length > 0 ? parts.join(" · ") : "正在识别题型与结构…";
}

/** 设置检测结果为简单文案（进度/错误等） */
function setDetectResultSimple(text, className) {
  const wrap = document.getElementById("detect-result");
  if (!wrap) return wrap;
  const simple = wrap.querySelector(".detect-result-simple");
  const structured = wrap.querySelector(".detect-result-structured");
  if (simple) {
    simple.textContent = text;
    simple.className = "detect-result-simple message-area " + (className || "").trim();
    simple.style.display = text ? "block" : "none";
  }
  if (structured) structured.style.display = "none";
  wrap.style.display = text ? "block" : "none";
  return wrap;
}

/**
 * 设置检测结果为成功。优先展示「题型与结构」（每题的小题数、含哪些组件），无 slots 时退回展示已识别字段列表。
 * @param {string} msg - 顶部摘要（如「已遍历 16/16 题…」）
 * @param {Array} [fields] - 原始字段列表（role/label），slots 为空时用于生成详情
 * @param {Array} [slots] - 每题结构：[{ index, subCount, sectionLabels, optionKind }]
 * @param {boolean} [hasTopLevelAudio] - 是否有大题共享音频框，可选展示在摘要中
 */
function setDetectResultSuccess(msg, fields, slots, hasTopLevelAudio) {
  const wrap = document.getElementById("detect-result");
  if (!wrap) return wrap;
  const simple = wrap.querySelector(".detect-result-simple");
  const structured = wrap.querySelector(".detect-result-structured");
  const summary = wrap.querySelector(".detect-result-summary");
  const detail = wrap.querySelector(".detect-result-detail");
  const toggle = wrap.querySelector(".detect-result-toggle");
  const toggleText = toggle && toggle.querySelector(".detect-result-toggle-text");
  if (simple) simple.style.display = "none";
  if (structured) structured.style.display = "block";

  let summaryLine = msg || "";
  if (hasTopLevelAudio === true && summaryLine) summaryLine += "；本页含大题共享音频。";
  if (hasTopLevelAudio === false && summaryLine) summaryLine += "；本页仅小题独立音频。";
  if (summary) {
    summary.textContent = summaryLine;
    summary.className = "detect-result-summary message-area text-success";
  }

  if (slots && slots.length > 0) {
    if (toggleText) toggleText.textContent = "题型与结构";
    if (detail) {
      const lines = slots.map((s) => {
        const n = s.subCount != null ? s.subCount : 1;
        const parts = s.sectionLabels && s.sectionLabels.length > 0
          ? s.sectionLabels.join("、")
          : "—";
        const opt = (s.optionKind === "image") ? "图片" : "文字";
        return `第${s.index}题：共${n}小题 · 含：${parts} · 选项：${opt}`;
      });
      detail.textContent = lines.join("\n");
      detail.className = "detect-result-detail message-area text-success collapsed";
    }
  } else {
    if (toggleText) toggleText.textContent = "已识别字段";
    const fieldList = (fields || []).map(f => (typeof f === "string" ? f : (f.label || f.role || ""))).filter(Boolean);
    if (detail) {
      detail.textContent = "已识别字段：" + (fieldList.length ? fieldList.join("、") : "—");
      detail.className = "detect-result-detail message-area text-success collapsed";
    }
  }

  if (toggle) {
    toggle.setAttribute("aria-expanded", "false");
    toggle.onclick = () => {
      const expanded = toggle.getAttribute("aria-expanded") === "true";
      toggle.setAttribute("aria-expanded", !expanded);
      if (detail) detail.classList.toggle("collapsed", expanded);
    };
  }
  wrap.style.display = "block";
  return wrap;
}

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
    const { selectors, lastDetectMessage, pageSlots, hasTopLevelAudio } = await chrome.storage.sync.get(["selectors", "lastDetectMessage", "pageSlots", "hasTopLevelAudio"]);
    const hintEl = document.getElementById("uploadStepHint");
    const resultEl = document.getElementById("detect-result");
    if (selectors && Object.keys(selectors).filter(k => selectors[k]).length > 0) {
      const fieldKeys = Object.keys(selectors).filter(k => selectors[k]);
      setDetectResultSuccess(lastDetectMessage || "已了解各题结构，允许上传 Word。", fieldKeys, pageSlots || null, hasTopLevelAudio);
      
      const detectBtnIcon = document.querySelector("#detect .card-icon");
      const detectBtnTitle = document.querySelector("#detect .card-title");
      const detectBtnDesc = document.querySelector("#detect .card-desc");
      
      // 更新按钮状态为已识别（三弧 logo 绿色）
      const detectEl = document.getElementById("detect");
      if (detectEl) detectEl.classList.add("is-analyzed");
      if (detectBtnIcon) {
        detectBtnIcon.innerHTML = DETECT_ICON_MAGNIFIER;
        detectBtnIcon.style.color = "var(--brand-green)";
        detectBtnIcon.style.background = "transparent";
        detectBtnIcon.classList.remove("icon-spin", "card-icon--searching");
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
      // 不再自动隐藏检测结果，保留「已遍历…」「已识别字段」一直可见

    } else {
      setDetectResultSimple("尚未分析当前页面，请点击上方「分析页面」按钮。", "text-muted");
      
      const detectBtnIcon = document.querySelector("#detect .card-icon");
      const detectBtnTitle = document.querySelector("#detect .card-title");
      const detectBtnDesc = document.querySelector("#detect .card-desc");
      
      // 恢复按钮初始状态（三弧 logo 白色）
      const detectElReset = document.getElementById("detect");
      if (detectElReset) detectElReset.classList.remove("is-analyzed");
      if (detectBtnIcon) {
        detectBtnIcon.innerHTML = DETECT_ICON_MAGNIFIER;
        detectBtnIcon.style.color = "";
        detectBtnIcon.style.background = "";
        detectBtnIcon.classList.remove("icon-spin", "card-icon--searching");
      }
      if (detectBtnTitle) {
        detectBtnTitle.textContent = "分析页面";
        detectBtnTitle.style.color = "";
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
  const { 
    selectedModel, 
    reasoningEffort: savedEffort, 
    parseDebugMode: savedDebug, 
    defaultAudioUrl: savedAudio, 
    defaultImageUrl: savedImage,
    ttsFemaleVoice: savedFemaleVoice,
    ttsMaleVoice: savedMaleVoice,
    ttsFemaleSpeed: savedFemaleSpeed,
    ttsMaleSpeed: savedMaleSpeed,
  } = await chrome.storage.sync.get([
    "selectedModel", "reasoningEffort", "parseDebugMode", "defaultAudioUrl", "defaultImageUrl",
    "ttsFemaleVoice", "ttsMaleVoice", "ttsFemaleSpeed", "ttsMaleSpeed"
  ]);

  const radios = document.querySelectorAll('input[name="modelChoice"]');
  const savedM = selectedModel || "doubao-seed-2-0-pro-260215";
  radios.forEach(r => { r.checked = (r.value === savedM); });
  radios.forEach(r => {
    r.addEventListener("change", () => {
      if (r.checked) chrome.storage.sync.set({ selectedModel: r.value });
    });
  });

  // ── 思考程度（reasoning.effort）──────────────────────────────────────────
  const effortEl = document.getElementById("reasoningEffort");
  if (effortEl) {
    const validEffort = ["minimal", "low", "medium", "high"];
    const effort = validEffort.includes(savedEffort) ? savedEffort : "medium";
    effortEl.value = effort;
    effortEl.addEventListener("change", () => {
      chrome.storage.sync.set({ reasoningEffort: effortEl.value });
    });
  }

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

  // ── TTS 语音合成设置 ──────────────────────────────────────────
  const ttsFemaleVoiceEl = document.getElementById("ttsFemaleVoice");
  const ttsMaleVoiceEl = document.getElementById("ttsMaleVoice");
  const ttsFemaleSpeedEl = document.getElementById("ttsFemaleSpeed");
  const ttsMaleSpeedEl = document.getElementById("ttsMaleSpeed");
  const ttsFemaleVolumeEl = document.getElementById("ttsFemaleVolume");
  const ttsMaleVolumeEl = document.getElementById("ttsMaleVolume");

  // 从 storage 获取音量设置
  const { ttsFemaleVolume: savedFemaleVolume, ttsMaleVolume: savedMaleVolume } = 
    await chrome.storage.sync.get(["ttsFemaleVolume", "ttsMaleVolume"]);

  // TTS 默认值
  const TTS_DEFAULTS = {
    femaleVoice: "en_female_amanda_mars_bigtts",
    maleVoice: "zh_male_jieshuonansheng_mars_bigtts",  // Morgan 解说男声
    femaleSpeed: "0.85",
    maleSpeed: "0.85",
    femaleVolume: "1.0",
    maleVolume: "1.0",
  };

  // 女声设置
  if (ttsFemaleVoiceEl) {
    ttsFemaleVoiceEl.value = savedFemaleVoice || TTS_DEFAULTS.femaleVoice;
    ttsFemaleVoiceEl.addEventListener("change", () => {
      chrome.storage.sync.set({ ttsFemaleVoice: ttsFemaleVoiceEl.value });
    });
  }
  if (ttsFemaleSpeedEl) {
    ttsFemaleSpeedEl.value = savedFemaleSpeed || TTS_DEFAULTS.femaleSpeed;
    ttsFemaleSpeedEl.addEventListener("change", () => {
      chrome.storage.sync.set({ ttsFemaleSpeed: ttsFemaleSpeedEl.value });
    });
  }
  if (ttsFemaleVolumeEl) {
    ttsFemaleVolumeEl.value = savedFemaleVolume || TTS_DEFAULTS.femaleVolume;
    ttsFemaleVolumeEl.addEventListener("change", () => {
      chrome.storage.sync.set({ ttsFemaleVolume: ttsFemaleVolumeEl.value });
    });
  }

  // 男声设置
  if (ttsMaleVoiceEl) {
    ttsMaleVoiceEl.value = savedMaleVoice || TTS_DEFAULTS.maleVoice;
    ttsMaleVoiceEl.addEventListener("change", () => {
      chrome.storage.sync.set({ ttsMaleVoice: ttsMaleVoiceEl.value });
    });
  }
  if (ttsMaleSpeedEl) {
    ttsMaleSpeedEl.value = savedMaleSpeed || TTS_DEFAULTS.maleSpeed;
    ttsMaleSpeedEl.addEventListener("change", () => {
      chrome.storage.sync.set({ ttsMaleSpeed: ttsMaleSpeedEl.value });
    });
  }
  if (ttsMaleVolumeEl) {
    ttsMaleVolumeEl.value = savedMaleVolume || TTS_DEFAULTS.maleVolume;
    ttsMaleVolumeEl.addEventListener("change", () => {
      chrome.storage.sync.set({ ttsMaleVolume: ttsMaleVolumeEl.value });
    });
  }

  // TTS 重置默认值按钮
  const ttsResetBtn = document.getElementById("ttsResetDefaults");
  if (ttsResetBtn) {
    ttsResetBtn.addEventListener("click", () => {
      // 清除存储中的 TTS 设置
      chrome.storage.sync.remove([
        "ttsFemaleVoice", "ttsMaleVoice",
        "ttsFemaleSpeed", "ttsMaleSpeed",
        "ttsFemaleVolume", "ttsMaleVolume"
      ], () => {
        // 重置 UI 为默认值
        if (ttsFemaleVoiceEl) ttsFemaleVoiceEl.value = TTS_DEFAULTS.femaleVoice;
        if (ttsMaleVoiceEl) ttsMaleVoiceEl.value = TTS_DEFAULTS.maleVoice;
        if (ttsFemaleSpeedEl) ttsFemaleSpeedEl.value = TTS_DEFAULTS.femaleSpeed;
        if (ttsMaleSpeedEl) ttsMaleSpeedEl.value = TTS_DEFAULTS.maleSpeed;
        if (ttsFemaleVolumeEl) ttsFemaleVolumeEl.value = TTS_DEFAULTS.femaleVolume;
        if (ttsMaleVolumeEl) ttsMaleVolumeEl.value = TTS_DEFAULTS.maleVolume;
        // 显示提示
        ttsResetBtn.textContent = "已重置 ✓";
        setTimeout(() => { ttsResetBtn.textContent = "重置"; }, 1500);
      });
    });
  }

  // ── TTS 试听弹窗 ──────────────────────────────────────────
  const ttsTestModal = document.getElementById("ttsTestModal");
  const ttsTestModalTitle = document.getElementById("ttsTestModalTitle");
  const ttsTestModalClose = document.getElementById("ttsTestModalClose");
  const ttsTestModalOk = document.getElementById("ttsTestModalOk");
  const ttsTestText = document.getElementById("ttsTestText");
  const ttsTestInfo = document.getElementById("ttsTestInfo");
  const ttsTestPlay = document.getElementById("ttsTestPlay");
  const ttsTestDownload = document.getElementById("ttsTestDownload");
  const ttsTestFemaleBtn = document.getElementById("ttsTestFemale");
  const ttsTestMaleBtn = document.getElementById("ttsTestMale");

  let currentTestGender = "female";
  let currentAudioBase64 = null;
  let currentAudioPlayer = null;

  const openTtsTestModal = (gender) => {
    currentTestGender = gender;
    currentAudioBase64 = null;
    ttsTestModalTitle.textContent = gender === "female" ? "👩 试听女声" : "👨 试听男声";
    ttsTestInfo.textContent = "";
    ttsTestInfo.className = "modal-info";
    // 设置默认英文试听文本
    if (ttsTestText && !ttsTestText.value?.trim()) {
      ttsTestText.value = "Hello, this is a test. The quick brown fox jumps over the lazy dog.";
    }
    ttsTestModal.classList.remove("hidden");
  };

  const closeTtsTestModal = () => {
    ttsTestModal.classList.add("hidden");
    if (currentAudioPlayer) {
      currentAudioPlayer.pause();
      currentAudioPlayer = null;
    }
  };

  if (ttsTestFemaleBtn) ttsTestFemaleBtn.addEventListener("click", () => openTtsTestModal("female"));
  if (ttsTestMaleBtn) ttsTestMaleBtn.addEventListener("click", () => openTtsTestModal("male"));
  if (ttsTestModalClose) ttsTestModalClose.addEventListener("click", closeTtsTestModal);
  if (ttsTestModalOk) ttsTestModalOk.addEventListener("click", closeTtsTestModal);
  if (ttsTestModal) {
    ttsTestModal.querySelector(".modal-backdrop")?.addEventListener("click", closeTtsTestModal);
  }

  // 试听按钮
  if (ttsTestPlay) {
    ttsTestPlay.addEventListener("click", async () => {
      const text = ttsTestText?.value?.trim();
      if (!text) {
        ttsTestInfo.textContent = "请输入要合成的文本";
        ttsTestInfo.className = "modal-info error";
        return;
      }

      // 获取当前设置
      const voice = currentTestGender === "female" 
        ? (ttsFemaleVoiceEl?.value || "zh_female_wanwanxiaohe_moon_bigtts")
        : (ttsMaleVoiceEl?.value || "zh_male_wennuanahu_moon_bigtts");
      const speed = currentTestGender === "female"
        ? parseFloat(ttsFemaleSpeedEl?.value || "0.85")
        : parseFloat(ttsMaleSpeedEl?.value || "0.85");
      const volume = currentTestGender === "female"
        ? parseFloat(ttsFemaleVolumeEl?.value || "1.0")
        : parseFloat(ttsMaleVolumeEl?.value || "1.0");

      // 检测英文音色是否包含中文
      const isEnglishVoice = voice.startsWith("en_");
      const hasChinese = /[\u4e00-\u9fa5]/.test(text);
      if (isEnglishVoice && hasChinese) {
        ttsTestInfo.textContent = "当前选择的是英文音色，不支持中文内容。请使用纯英文文本或切换为中文音色。";
        ttsTestInfo.className = "modal-info error";
        return;
      }

      ttsTestInfo.textContent = "正在合成音频...";
      ttsTestInfo.className = "modal-info loading";

      try {
        const resp = await fetch("http://127.0.0.1:8766/api/tts", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            text,
            speaker: voice,
            speed_ratio: speed,
            volume_ratio: volume,
            format: "mp3",
            sample_rate: 24000,
          }),
        });

        if (!resp.ok) {
          const errText = await resp.text().catch(() => "");
          throw new Error(errText || `HTTP ${resp.status}`);
        }

        const data = await resp.json();
        if (!data.audioBase64) {
          throw new Error("返回数据无音频");
        }

        currentAudioBase64 = data.audioBase64;
        ttsTestInfo.textContent = `合成成功！语速: ${speed}x, 音量: ${volume}x`;
        ttsTestInfo.className = "modal-info";

        // 自动播放
        const audioUrl = `data:audio/mp3;base64,${currentAudioBase64}`;
        if (currentAudioPlayer) currentAudioPlayer.pause();
        currentAudioPlayer = new Audio(audioUrl);
        currentAudioPlayer.volume = Math.min(volume, 1.0);
        currentAudioPlayer.play();

      } catch (e) {
        ttsTestInfo.textContent = `合成失败: ${e.message}`;
        ttsTestInfo.className = "modal-info error";
      }
    });
  }

  // 下载按钮
  if (ttsTestDownload) {
    ttsTestDownload.addEventListener("click", () => {
      if (!currentAudioBase64) {
        ttsTestInfo.textContent = "请先点击试听生成音频";
        ttsTestInfo.className = "modal-info error";
        return;
      }

      const blob = base64ToBlob(currentAudioBase64, "audio/mpeg");
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `tts_${currentTestGender}_${Date.now()}.mp3`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      ttsTestInfo.textContent = "下载已开始";
      ttsTestInfo.className = "modal-info";
    });
  }

  // base64 转 Blob
  function base64ToBlob(base64, mimeType) {
    const byteChars = atob(base64);
    const byteNumbers = new Array(byteChars.length);
    for (let i = 0; i < byteChars.length; i++) {
      byteNumbers[i] = byteChars.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    return new Blob([byteArray], { type: mimeType });
  }

  // ── TTS 自定义合成弹窗（支持对话模式）──────────────────────────
  const ttsSynthModal = document.getElementById("ttsSynthModal");
  const ttsSynthModalClose = document.getElementById("ttsSynthModalClose");
  const ttsSynthModalCancel = document.getElementById("ttsSynthModalCancel");
  const ttsSynthText = document.getElementById("ttsSynthText");
  const ttsSynthInfo = document.getElementById("ttsSynthInfo");
  const ttsSynthPlay = document.getElementById("ttsSynthPlay");
  const ttsSynthDownload = document.getElementById("ttsSynthDownload");
  const ttsSynthAudioWrap = document.getElementById("ttsSynthAudioWrap");
  const ttsSynthAudio = document.getElementById("ttsSynthAudio");
  const ttsSynthesizeBtn = document.getElementById("ttsSynthesizeBtn");

  let synthAudioBase64 = null;

  const openTtsSynthModal = () => {
    synthAudioBase64 = null;
    ttsSynthInfo.textContent = "";
    ttsSynthInfo.className = "modal-info";
    ttsSynthAudioWrap.classList.add("hidden");
    ttsSynthDownload.disabled = true;
    ttsSynthModal.classList.remove("hidden");
  };

  const closeTtsSynthModal = () => {
    ttsSynthModal.classList.add("hidden");
    if (ttsSynthAudio) {
      ttsSynthAudio.pause();
      ttsSynthAudio.src = "";
    }
  };

  if (ttsSynthesizeBtn) ttsSynthesizeBtn.addEventListener("click", openTtsSynthModal);
  if (ttsSynthModalClose) ttsSynthModalClose.addEventListener("click", closeTtsSynthModal);
  if (ttsSynthModalCancel) ttsSynthModalCancel.addEventListener("click", closeTtsSynthModal);
  if (ttsSynthModal) {
    ttsSynthModal.querySelector(".modal-backdrop")?.addEventListener("click", closeTtsSynthModal);
  }

  // 合成按钮
  if (ttsSynthPlay) {
    ttsSynthPlay.addEventListener("click", async () => {
      const text = ttsSynthText?.value?.trim();
      if (!text) {
        ttsSynthInfo.textContent = "请输入要合成的文本";
        ttsSynthInfo.className = "modal-info error";
        return;
      }

      // 获取当前 TTS 设置
      const femaleVoice = ttsFemaleVoiceEl?.value || TTS_DEFAULTS.femaleVoice;
      const maleVoice = ttsMaleVoiceEl?.value || TTS_DEFAULTS.maleVoice;
      const femaleSpeed = parseFloat(ttsFemaleSpeedEl?.value || TTS_DEFAULTS.femaleSpeed);
      const maleSpeed = parseFloat(ttsMaleSpeedEl?.value || TTS_DEFAULTS.maleSpeed);
      const femaleVolume = parseFloat(ttsFemaleVolumeEl?.value || TTS_DEFAULTS.femaleVolume);
      const maleVolume = parseFloat(ttsMaleVolumeEl?.value || TTS_DEFAULTS.maleVolume);

      // 检测是否为对话模式
      const isDialogue = /^[WwMmQqAa][：:]/m.test(text);

      ttsSynthInfo.textContent = isDialogue ? "正在合成对话音频..." : "正在合成音频...";
      ttsSynthInfo.className = "modal-info loading";
      ttsSynthPlay.disabled = true;
      ttsSynthAudioWrap.classList.add("hidden");

      try {
        const resp = await fetch("http://127.0.0.1:8766/api/tts", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            text,
            dialogue: isDialogue,
            format: "mp3",
            sample_rate: 24000,
            // 非对话模式使用女声设置
            speaker: isDialogue ? undefined : femaleVoice,
            speed_ratio: isDialogue ? undefined : femaleSpeed,
            volume_ratio: isDialogue ? undefined : femaleVolume,
            // 对话模式参数
            female_speaker: femaleVoice,
            male_speaker: maleVoice,
            female_speed: femaleSpeed,
            male_speed: maleSpeed,
            female_volume: femaleVolume,
            male_volume: maleVolume,
          }),
        });

        if (!resp.ok) {
          const errText = await resp.text().catch(() => "");
          throw new Error(errText || `HTTP ${resp.status}`);
        }

        const data = await resp.json();
        if (!data.audioBase64) {
          throw new Error("返回数据无音频");
        }

        synthAudioBase64 = data.audioBase64;
        const modeDesc = isDialogue ? "对话模式" : "单音色模式";
        ttsSynthInfo.textContent = `合成成功！(${modeDesc})`;
        ttsSynthInfo.className = "modal-info";

        // 显示播放器
        const audioUrl = `data:audio/mp3;base64,${synthAudioBase64}`;
        ttsSynthAudio.src = audioUrl;
        ttsSynthAudioWrap.classList.remove("hidden");
        ttsSynthDownload.disabled = false;

        // 自动播放
        ttsSynthAudio.play();

      } catch (e) {
        ttsSynthInfo.textContent = `合成失败: ${e.message}`;
        ttsSynthInfo.className = "modal-info error";
      } finally {
        ttsSynthPlay.disabled = false;
      }
    });
  }

  // 下载按钮
  if (ttsSynthDownload) {
    ttsSynthDownload.addEventListener("click", () => {
      if (!synthAudioBase64) {
        ttsSynthInfo.textContent = "请先点击合成生成音频";
        ttsSynthInfo.className = "modal-info error";
        return;
      }

      const blob = base64ToBlob(synthAudioBase64, "audio/mpeg");
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `tts_custom_${Date.now()}.mp3`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      ttsSynthInfo.textContent = "下载已开始";
      ttsSynthInfo.className = "modal-info";
    });
  }

  // ── 侧边栏模式 ──────────────────────────────────────────────────
  const openSidePanelBtn = document.getElementById("openSidePanel");
  if (openSidePanelBtn) {
    openSidePanelBtn.addEventListener("click", async () => {
      try {
        // Chrome 114+ 支持 side panel API
        if (chrome.sidePanel) {
          const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
          if (tab?.id) {
            await chrome.sidePanel.open({ tabId: tab.id });
            window.close(); // 关闭弹窗
          }
        } else {
          alert("当前浏览器版本不支持侧边栏功能，请升级到 Chrome 114 或更高版本。");
        }
      } catch (e) {
        console.error("打开侧边栏失败:", e);
        alert("打开侧边栏失败: " + e.message);
      }
    });
  }
})();

// ─── 接收来自 background.js 和 content.js 的实时消息 ───────────────────────
chrome.runtime.onMessage.addListener((msg) => {
  const resultEl = document.getElementById("detect-result");
  if (resultEl) resultEl.style.display = "block";
  if (msg.type === "DETECT_PROGRESS") {
    const total = msg.total ? `/${msg.total}` : "";
    // 优先展示当前题自身的区块标签（实时反映本题结构），无则退回累计字段描述
    const currentLabels = msg.currentSectionLabels;
    const desc = (currentLabels && currentLabels.length > 0)
      ? currentLabels.join("、")
      : describeStructureFromFields(msg.fields || []);
    // 在进度前显示题型名称（简洁格式：一、听后选择）
    const typeHint = msg.typeHint || "";
    const typePrefix = typeHint ? `${toChineseNumeral(msg.walked)}、${typeHint}，` : "";
    setDetectResultSimple(`正在遍历第 ${msg.walked}${total} 题 · ${typePrefix}含：${desc}`, "text-muted");
    if (!resultEl) return;
    
    const detectBtnIcon = document.querySelector("#detect .card-icon");
    if (detectBtnIcon && !detectBtnIcon.classList.contains("card-icon--searching")) {
      detectBtnIcon.innerHTML = DETECT_ICON_MAGNIFIER;
      detectBtnIcon.classList.add("card-icon--searching");
      detectBtnIcon.classList.remove("icon-spin");
      detectBtnIcon.style.color = "var(--brand-green)";
      detectBtnIcon.style.background = "transparent";
      
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
  if (msg.type === "TTS_PROGRESS") {
    // TTS 进度现在在录题页面上显示，这里仅作日志记录
    console.log("[TTS] 进度:", msg);
  }
  if (msg.type === "PARSE_DONE") {
    setDropZoneState("done", { questions: msg.questions, debug_info: msg.debug_info });
    setMsg("识别完成，正在自动填入录题页…", false);
    updateFillCacheDot();
    pushJsonHistory(JSON.stringify(msg.questions), msg.debug_info || null).catch(() => {});
    // TTS 音频合成在 content.js 填充时异步进行
    doFill(msg.questions, true);
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
    if (!isError && (text.includes("豆包识别") || text.includes("识别中"))) {
      el.classList.add("msg-parsing");
    } else {
      el.classList.remove("msg-parsing");
    }
  } else {
    el.style.display = "none";
    el.classList.remove("msg-parsing");
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

/**
 * 填充暂停时，在 msg 区域显示提示，并展示「继续填充」按钮。
 * 用户手动补全当前题并让页面跳到下一题后，点按钮即可继续。
 */
function showPausedFillUI(pausedQuestion, remaining) {
  const resumeBtn = document.getElementById("resumeFillBtn");
  if (!resumeBtn) return;

  const count = remaining ? remaining.length : 0;
  // 转换为在整个 _lastToFill 中的绝对题号（1-based）
  const absQuestion = _lastFillOffset + pausedQuestion;
  setMsg(
    `第 ${absQuestion} 题保存后未自动跳转（可能有必填项未填，如音频）。` +
    `请手动补全该题并保存，待页面切换到下一题后，点击「继续填充」。`,
    true
  );

  if (count > 0) {
    resumeBtn.textContent = `继续填充（剩余 ${count} 题）`;
    resumeBtn.classList.remove("hidden");
    resumeBtn.onclick = async () => {
      resumeBtn.classList.add("hidden");
      resumeBtn.onclick = null;
      await doFill(remaining);
    };
  } else {
    resumeBtn.classList.add("hidden");
    resumeBtn.onclick = null;
  }

  // 同步「从第X题起填充」：把输入框定位到暂停的那道题，方便用户核对后手动调整
  if (_lastToFill && _lastToFill.length > 0) {
    const input = document.getElementById("fillFromIndexInput");
    if (input) {
      input.value = absQuestion;
      updateFillFromIndexPreview();
    }
    const wrap = document.getElementById("fillFromIndexWrap");
    if (wrap) wrap.classList.remove("hidden");
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
  const detectEl = document.getElementById("detect");
  setDetectResultSimple("正在检测页面结构，将自动切换下一题遍历各题…", "text-muted");

  const detectBtnIcon = document.querySelector("#detect .card-icon");
  const detectBtnTitle = document.querySelector("#detect .card-title");
  const detectBtnDesc = document.querySelector("#detect .card-desc");

  // 状态更新为分析中（放大镜 + 查找动效）
  if (detectEl) detectEl.classList.add("is-analyzing");
  if (detectBtnIcon) {
    detectBtnIcon.innerHTML = DETECT_ICON_MAGNIFIER;
    detectBtnIcon.classList.add("card-icon--searching");
    detectBtnIcon.classList.remove("icon-spin");
    detectBtnIcon.style.color = "var(--brand-green)";
    detectBtnIcon.style.background = "transparent";
  }
  if (detectBtnTitle) {
    detectBtnTitle.textContent = "分析中...";
    detectBtnTitle.style.color = "var(--brand-green)";
  }
  if (detectBtnDesc) detectBtnDesc.textContent = "正在遍历各题";

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) {
    setDetectResultSimple("无法获取当前标签页", "text-error");
    if (detectBtnIcon) {
      detectBtnIcon.classList.remove("icon-spin", "card-icon--searching");
      detectBtnIcon.innerHTML = DETECT_ICON_MAGNIFIER;
      detectBtnIcon.style.color = "";
      detectBtnIcon.style.background = "";
    }
    const detectEl = document.getElementById("detect");
    if (detectEl) detectEl.classList.remove("is-analyzed");
    if (detectBtnTitle) {
      detectBtnTitle.textContent = "分析页面";
      detectBtnTitle.style.color = "";
    }
    if (detectBtnDesc) detectBtnDesc.textContent = "识别当前表单";
    if (detectEl) detectEl.classList.remove("is-analyzing");
    return;
  }
  // 先发 PING 确认 content 已注入，避免误报「注入失败」
  try {
    await chrome.tabs.sendMessage(tab.id, { type: "PING" });
  } catch (pingErr) {
    setDetectResultSimple("无法连接录题页，请确保当前标签页是录题页并刷新该页后再点「开始检测」。", "text-error");
    if (detectEl) detectEl.classList.remove("is-analyzing");
    return;
  }
  setDetectResultSimple("正在遍历各题（约 10～30 秒），请勿关闭本弹窗…", "text-muted");
  const { selectors: stored } = await chrome.storage.sync.get("selectors");
  try {
    const res = await chrome.tabs.sendMessage(tab.id, { type: "DETECT_AND_WALK", selectors: stored || {} });
    
    if (detectBtnIcon) detectBtnIcon.classList.remove("icon-spin", "card-icon--searching");

    if (!res || !res.ok) {
      setDetectResultSimple(res?.error || "分析失败，请确认在录题页面", "text-error");
      if (detectBtnIcon) {
        detectBtnIcon.innerHTML = DETECT_ICON_MAGNIFIER;
        detectBtnIcon.style.color = "";
        detectBtnIcon.style.background = "";
      }
      if (detectBtnTitle) {
        detectBtnTitle.textContent = "分析页面";
        detectBtnTitle.style.color = "";
      }
      if (detectBtnDesc) detectBtnDesc.textContent = "识别当前表单";
      if (detectEl) detectEl.classList.remove("is-analyzing", "is-analyzed");
      return;
    }
    const { selectors, fields, message, walked, total, slots, hasTopLevelAudio } = res;
    const walked_n = walked || "?";
    const total_n = total ? `/${total}` : "";
    const msg = message || `已遍历 ${walked_n}${total_n} 题，了解页面结构，允许上传 Word。`;
    await chrome.storage.sync.set({
      selectors,
      fields: fields || [],
      lastDetectMessage: msg,
      pageQuestionTotal: total != null && total > 0 ? total : null,
      pageSlots: slots && slots.length > 0 ? slimSlotsForStorage(slots) : null,
      hasTopLevelAudio: !!hasTopLevelAudio,
    });
    // 完整 slots（含 currentSlotFields/subQuestions/sectionLabels）存 local，供后端构建 prompt 用
    chrome.storage.local.set({ pageSlotsFull: slots && slots.length > 0 ? slots : null });
    if (tab?.id) {
      chrome.storage.local.set({ fillTargetTabId: tab.id });
      if (total != null && total > 0) {
        chrome.runtime.sendMessage({ type: "REFRESH_BADGE", tabId: tab.id }).catch(() => {});
      }
    }
    setDetectResultSuccess(msg, fields || [], slots || [], hasTopLevelAudio);
    
    // 更新按钮状态为已识别（三弧 logo 绿色）
    if (detectEl) detectEl.classList.remove("is-analyzing"), detectEl.classList.add("is-analyzed");
    if (detectBtnIcon) {
      detectBtnIcon.innerHTML = DETECT_ICON_MAGNIFIER;
      detectBtnIcon.style.color = "var(--brand-green)";
      detectBtnIcon.style.background = "transparent";
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
    // 分析完成后保留检测结果可见，不自动隐藏
  } catch (e) {
    const errMsg = String(e?.message || e);
    if (detectEl) detectEl.classList.remove("is-analyzing", "is-analyzed");
    if (detectBtnIcon) {
      detectBtnIcon.classList.remove("icon-spin", "card-icon--searching");
      detectBtnIcon.innerHTML = DETECT_ICON_MAGNIFIER;
      detectBtnIcon.style.color = "";
      detectBtnIcon.style.background = "";
    }
    if (detectBtnTitle) {
      detectBtnTitle.textContent = "分析页面";
      detectBtnTitle.style.color = "";
    }
    if (detectBtnDesc) detectBtnDesc.textContent = "识别当前表单";

    if (/receiving end|establish connection|target closed/i.test(errMsg)) {
      setDetectResultSimple("连接中断（检测时间较长时请勿关闭弹窗），请刷新录题页后重新点「开始检测」。", "text-error");
    } else {
      setDetectResultSimple("注入失败：" + (errMsg || "请刷新录题页面后再试"), "text-error");
    }
  }
});

function getWordFiles(files) {
  if (!files || !files.length) return [];
  return Array.from(files).filter((f) => {
    const n = (f.name || "").toLowerCase();
    return n.endsWith(".docx") || n.endsWith(".doc") || n.endsWith(".pdf");
  });
}

/** 从 drop 事件的 dataTransfer 取文件（弹窗内 dataTransfer.files 常为空，用 items 兜底） */
function getFilesFromDrop(dataTransfer) {
  if (!dataTransfer) return [];
  const fromFiles = dataTransfer.files && dataTransfer.files.length > 0
    ? Array.from(dataTransfer.files) : [];
  if (fromFiles.length > 0) return fromFiles;
  if (!dataTransfer.items || !dataTransfer.items.length) return [];
  const fromItems = [];
  for (let i = 0; i < dataTransfer.items.length; i++) {
    const item = dataTransfer.items[i];
    if (item.kind === "file") {
      const f = item.getAsFile();
      if (f) fromItems.push(f);
    }
  }
  return fromItems;
}

// ─── 记录发起上传时的录题页 tabId（避免解析期间切标签导致填充到错误页）
let targetTabId = null;

// ─── 从第X题起填充：记录最近一次 doFill 得到的规范化题目列表和起始偏移 ──────
let _lastToFill    = null;     // 规范化后的完整题目数组
let _lastFillOffset = 0;       // 本次 fill 在 _lastToFill 中的 0-based 起始索引
let _isFilling     = false;    // 全局填充锁，防止并发触发多次 FILL_FORM

/**
 * 统一开关「填充中」状态：禁用/恢复所有填充按钮，设置全局锁。
 * 需要禁用的按钮：立即填入、从第X题填入、继续填充、高级面板填入、历史记录。
 */
function setFillingState(on) {
  _isFilling = on;
  const FILL_BTN_IDS = ["fillFromResult", "fillFromIndexGo", "resumeFillBtn", "fill"];
  for (const id of FILL_BTN_IDS) {
    const el = document.getElementById(id);
    if (!el) continue;
    el.disabled        = on;
    el.style.opacity   = on ? "0.55" : "";
    el.style.cursor    = on ? "not-allowed" : "";
    el.style.pointerEvents = on ? "none" : "";
  }
  const historyBtn = document.getElementById("jsonHistoryBtn");
  if (historyBtn) {
    historyBtn.disabled        = on;
    historyBtn.style.opacity   = on ? "0.55" : "";
    historyBtn.style.pointerEvents = on ? "none" : "";
  }
}

// ─── 消息区 ────────────────────────────────────────────────────────────────
// function setMsg() was moved up

// 题型 → 中文（含听说题型），复制/预览时使用
const QUESTION_TYPE_MAP = {
  single: "单选", multiple: "多选", judge: "判断", blank: "填空",
  listening_choice: "听后选择", listening_response: "听后应答", reading_aloud: "交际朗读",
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
async function pushJsonHistory(jsonText, debugInfo) {
  console.log("[pushJsonHistory] 开始保存历史, jsonText长度:", jsonText?.length);
  if (!jsonText || typeof jsonText !== "string" || !jsonText.trim()) {
    console.log("[pushJsonHistory] jsonText 无效，跳过");
    return;
  }
  const list = await getJsonHistory();
  console.log("[pushJsonHistory] 当前历史条数:", list.length);
  const trimmed = jsonText.trim();
  const next = {
    text: trimmed,
    time: Date.now(),
    debug_info: debugInfo && (debugInfo.system_prompt || debugInfo.user_content)
      ? { system_prompt: debugInfo.system_prompt || "", user_content: debugInfo.user_content || "" }
      : undefined,
  };
  const filtered = list.filter((item) => item.text !== trimmed);
  const nextList = [next, ...filtered].slice(0, JSON_HISTORY_MAX);
  await chrome.storage.local.set({ jsonHistory: nextList });
  console.log("[pushJsonHistory] 保存完成，新历史条数:", nextList.length);
}

let _parsingDotsTimer = null;

function setDropZoneState(state, data = {}) {
  const dz       = document.getElementById("dropZone");
  const dzN      = document.getElementById("dzNormal");
  const dzP      = document.getElementById("dzParsing");
  const dzD      = document.getElementById("dzDone");
  const upBtn    = document.getElementById("uploadAndFill");

  // 如果关键元素不在 DOM（弹窗在非录题页打开），静默跳过
  if (!dz || !dzN || !dzP || !dzD || !upBtn) return;

  // 停止「识别中. / .. / ...」循环
  if (_parsingDotsTimer) {
    clearInterval(_parsingDotsTimer);
    _parsingDotsTimer = null;
  }
  const upIcon = upBtn.querySelector(".card-icon");
  if (upIcon) upIcon.classList.remove("card-icon--parsing");

  // 重置所有子层
  dzN.classList.add("hidden");
  dzP.classList.add("hidden");
  dzD.classList.add("hidden");
  dz.className = "drop-zone";

  if (state === "idle") {
    dzN.classList.remove("hidden");
    dzN.querySelector(".drop-icon").innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>`;
    dzN.querySelector(".drop-text").textContent = "点击选择文件";
    dzN.querySelector(".drop-subtext").textContent = "支持 Word (.docx) 和 PDF 格式，可多选";
    upBtn.classList.remove("is-parsing");
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
    upBtn.classList.remove("is-parsing");
    upBtn.style.pointerEvents = "auto";
    const upTitle = upBtn.querySelector(".card-title");
    if (upTitle) upTitle.textContent = "解析并填入";

  } else if (state === "parsing") {
    dzP.classList.remove("hidden");
    dz.classList.add("is-parsing");
    upBtn.classList.add("is-parsing");
    upBtn.style.pointerEvents = "none";
    if (upIcon) upIcon.classList.add("card-icon--parsing");
    const upTitle = upBtn.querySelector(".card-title");
    const dzParsingText = dzP.querySelector(".drop-text");
    const dots = [".", "..", "..."];
    let dotCount = 2; // 先显示 "."，再 ".."，再 "..."
    const tick = () => {
      dotCount = (dotCount + 1) % 3;
      const d = dots[dotCount];
      if (upTitle) upTitle.textContent = "识别中" + d;
      if (dzParsingText) dzParsingText.textContent = "正在 AI 识别中" + d;
    };
    tick();
    _parsingDotsTimer = setInterval(tick, 420);
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
    upBtn.classList.remove("is-parsing");
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
    const fillFromResultBtn = document.getElementById("fillFromResult");
    if (fillFromResultBtn) {
      if (qs.length > 0) {
        fillFromResultBtn.classList.remove("hidden");
        fillFromResultBtn.classList.add("btn-fill-result");
        fillFromResultBtn.onclick = async (e) => {
          e.stopPropagation();
          e.preventDefault();
          await doFill(qs);
        };
      } else {
        fillFromResultBtn.classList.add("hidden");
        fillFromResultBtn.onclick = null;
      }
    }

    // 恢复「解析并填入」卡片为「重新解析」，避免还显示「识别中...」
    upBtn.style.opacity = "1";
    upBtn.style.pointerEvents = "auto";
    const upTitleDone = upBtn.querySelector(".card-title");
    const upDescDone = upBtn.querySelector(".card-desc");
    if (upTitleDone) upTitleDone.textContent = "重新解析";
    if (upDescDone) upDescDone.textContent = "上传新文件重新识别";

    // 识别完成后立刻预规范化题目，让「从第X题起填充」面板即时可用
    if (qs.length > 0) {
      (async () => {
        const { pageQuestionTotal, pageSlots, hasTopLevelAudio } =
          await chrome.storage.sync.get(["pageQuestionTotal", "pageSlots", "hasTopLevelAudio"]);
        const pageTotal = pageQuestionTotal != null && pageQuestionTotal > 0 ? pageQuestionTotal : null;
        const normalized = normalizeQuestionsToSlots(qs, pageTotal, pageSlots || null, hasTopLevelAudio);
        // 仅展示真实题数（不含 doFill 里补齐的空题占位）
        _lastToFill    = pageTotal && normalized.length > pageTotal ? normalized.slice(0, pageTotal) : normalized;
        _lastFillOffset = 0;
        showFillFromIndexUI(_lastToFill);
      })();
    }
  }
}

// ─── 将 File 读取为 base64 ─────────────────────────────────────────────────
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload  = (e) => {
      let mimeType = file.type;
      if (!mimeType) {
        const ext = (file.name || "").toLowerCase().split(".").pop();
        if (ext === "pdf") mimeType = "application/pdf";
        else if (ext === "docx") mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        else if (ext === "doc") mimeType = "application/msword";
        else mimeType = "application/octet-stream";
      }
      resolve({
        name: file.name,
        mimeType,
        base64: e.target.result.split(",")[1],
      });
    };
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
 * @param {boolean} [hasTopLevelAudio] - 页面是否检测到大题共享音频输入框
 * @returns {Array} 按大题组织的题目，适合 FILL_FORM
 */
function normalizeQuestionsToSlots(questions, pageTotal, pageSlots, hasTopLevelAudio) {
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
            keyword: b.keyword || "",
            answer: b.answer != null ? String(b.answer).trim() : "",
            options: Array.isArray(b.options) ? b.options : [],
            candidates: Array.isArray(b.candidates) ? b.candidates : [],
            listening_script: b.listening_script || "",
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
            keyword: b.keyword || "",
            answer: b.answer != null ? String(b.answer).trim() : "",
            options: Array.isArray(b.options) ? b.options : [],
            candidates: Array.isArray(b.candidates) ? b.candidates : [],
            listening_script: b.listening_script || "",
          })),
        });
        i = j;
        continue;
      }
    }
    out.push(q);
    i++;
  }

  // ── 后处理：根据页面是否有「顶层共享音频框」修正 listening_script 的分布 ──
  // hasTopLevelAudio===true  → 情形 B：blanks[n].listening_script 应留小题各自内容，
  //                            顶层 listening_script 存共享对话（LLM 已做，保持原样）
  // hasTopLevelAudio===false → 情形 A：页面只有各小题音频框，共享对话必须拆到每个 blank，
  //                            若顶层 listening_script 有内容而各 blank 为空 → 下放
  if (hasTopLevelAudio === false) {
    for (const q of out) {
      if (!q.blanks || q.blanks.length === 0) continue;
      const topScript = (q.listening_script || "").trim();
      if (!topScript) continue;
      const allBlanksEmpty = q.blanks.every(b => !(b.listening_script || "").trim());
      if (allBlanksEmpty) {
        // 顶层有内容，各 blank 全空 → 把共享内容下发到每个 blank
        q.blanks.forEach(b => { b.listening_script = topScript; });
        q.listening_script = "";  // 顶层清空（页面没有顶层音频框）
      }
    }
  }

  return out;
}

// ─── 从第X题起填充：工具函数 ────────────────────────────────────────────────

/** 根据当前输入框值更新预览文字 */
function updateFillFromIndexPreview() {
  const input   = document.getElementById("fillFromIndexInput");
  const preview = document.getElementById("fillFromIndexPreview");
  if (!input || !preview || !_lastToFill || _lastToFill.length === 0) {
    if (preview) preview.textContent = "";
    return;
  }
  const maxIdx = _lastToFill.length;
  const idx    = Math.min(Math.max(parseInt(input.value, 10) || 1, 1), maxIdx) - 1;
  const q      = _lastToFill[idx];
  if (!q) { preview.textContent = ""; return; }
  const typeLabel = QUESTION_TYPE_MAP[q.type] || q.type || "?";
  const raw       = (q.question || (q.blanks && q.blanks[0]?.question) || "(无题干)").trim();
  const excerpt   = raw.slice(0, 38) + (raw.length > 38 ? "…" : "");
  preview.textContent = `[${typeLabel}] ${excerpt}`;
}

/** 展示「从第X题起填充」面板，并把输入框范围设为 1～total */
function showFillFromIndexUI(questions) {
  const wrap       = document.getElementById("fillFromIndexWrap");
  const input      = document.getElementById("fillFromIndexInput");
  const totalLabel = document.getElementById("fillFromIndexTotalLabel");
  if (!wrap || !input || !totalLabel) return;
  const total = questions.length;
  input.max = total;
  // 不重置已有值（用户可能已手动调整），仅将越界值夹回合法范围
  const cur = parseInt(input.value, 10) || 1;
  if (cur < 1 || cur > total) input.value = 1;
  totalLabel.textContent = `/ ${total} 题起`;
  updateFillFromIndexPreview();
  wrap.classList.remove("hidden");
}

/** 出错时将面板滚入视野并短暂绿色边框提示，避免用户看不到 */
function bringFillFromIndexIntoView() {
  const wrap = document.getElementById("fillFromIndexWrap");
  if (!wrap || wrap.classList.contains("hidden")) return;
  // 稍微延迟等错误消息渲染完再滚动
  setTimeout(() => {
    wrap.scrollIntoView({ behavior: "smooth", block: "nearest" });
    wrap.style.transition = "box-shadow 0.25s ease";
    wrap.style.boxShadow  = "0 0 0 2px var(--brand-green)";
    setTimeout(() => { wrap.style.boxShadow = ""; }, 1400);
  }, 80);
}

/** 重置时隐藏面板并清空缓存题目，同时解除填充锁（避免重置后锁残留） */
function hideFillFromIndexUI() {
  const wrap = document.getElementById("fillFromIndexWrap");
  if (wrap) wrap.classList.add("hidden");
  _lastToFill    = null;
  _lastFillOffset = 0;
  if (_isFilling) setFillingState(false);
}

/**
 * 直接填充已规范化的题目切片（跳过 normalizeQuestionsToSlots 和 pageTotal 补齐），
 * 适合从第 N 题恢复填充的场景。
 */
async function doFillFromNormalized(normalizedSlice, startIdx1Based) {
  if (!normalizedSlice || normalizedSlice.length === 0) {
    setMsg("没有可填入的题目", true);
    return;
  }
  if (_isFilling) {
    setMsg("正在填充中，请等当前任务完成后再操作。", false);
    return;
  }
  let tabId = targetTabId;
  if (!tabId) {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    tabId = tab?.id;
  }
  if (!tabId) { setMsg("无法确定录题页，请确认页面已打开", true); return; }

  const { selectors, defaultAudioUrl, defaultImageUrl, ttsFemaleVoice, ttsMaleVoice, ttsFemaleSpeed, ttsMaleSpeed, ttsFemaleVolume, ttsMaleVolume } =
    await chrome.storage.sync.get(["selectors", "defaultAudioUrl", "defaultImageUrl", "ttsFemaleVoice", "ttsMaleVoice", "ttsFemaleSpeed", "ttsMaleSpeed", "ttsFemaleVolume", "ttsMaleVolume"]);

  setMsg(`从第 ${startIdx1Based} 题起填充，共 ${normalizedSlice.length} 题…`, false);
  const _resumeBtn = document.getElementById("resumeFillBtn");
  if (_resumeBtn) { _resumeBtn.classList.add("hidden"); _resumeBtn.onclick = null; }

  _lastFillOffset = startIdx1Based - 1;
  setFillingState(true);

  try {
    const result = await chrome.tabs.sendMessage(tabId, {
      type: "FILL_FORM",
      questions: normalizedSlice,
      selectors,
      defaultAudioUrl: (defaultAudioUrl || "").trim() || undefined,
      defaultImageUrl: (defaultImageUrl || "").trim() || undefined,
      ttsSettings: {
        femaleVoice: ttsFemaleVoice || "en_female_amanda_mars_bigtts",
        maleVoice: ttsMaleVoice || "zh_male_jieshuonansheng_mars_bigtts",
        femaleSpeed: parseFloat(ttsFemaleSpeed) || 0.85,
        maleSpeed: parseFloat(ttsMaleSpeed) || 0.85,
        femaleVolume: parseFloat(ttsFemaleVolume) || 1.0,
        maleVolume: parseFloat(ttsMaleVolume) || 1.0,
      },
      debugSource: "parse",
    });
    setFillingState(false);
    if (result?.ok === "paused") {
      showPausedFillUI(result.pausedQuestion, result.remaining || []);
      return;
    }
    if (result?.ok === false) {
      setMsg((result.error || "填充出错") + "。", true);
      bringFillFromIndexIntoView();
      return;
    }
    setMsg(result?.message || `填充完成，共 ${result?.filled ?? normalizedSlice.length} 题。`, false);
  } catch (e) {
    setFillingState(false);
    if (String(e).includes("Could not establish connection") || String(e).includes("receiving end does not exist")) {
      setMsg("录题页未响应，请刷新该页面后再试。", true);
    } else {
      setMsg(`填充失败：${e.message || e}。`, true);
    }
    bringFillFromIndexIntoView();
  }
}

// 事件绑定：从第X题起填充控件
(function () {
  const decBtn = document.getElementById("fillFromIndexDec");
  const incBtn = document.getElementById("fillFromIndexInc");
  const input  = document.getElementById("fillFromIndexInput");
  const goBtn  = document.getElementById("fillFromIndexGo");
  if (!decBtn || !incBtn || !input || !goBtn) return;

  decBtn.addEventListener("click", () => {
    const v = Math.max(1, (parseInt(input.value, 10) || 1) - 1);
    input.value = v;
    updateFillFromIndexPreview();
  });
  incBtn.addEventListener("click", () => {
    const max = parseInt(input.max, 10) || 999;
    const v   = Math.min(max, (parseInt(input.value, 10) || 1) + 1);
    input.value = v;
    updateFillFromIndexPreview();
  });
  input.addEventListener("input",  () => updateFillFromIndexPreview());
  input.addEventListener("change", () => {
    const max = parseInt(input.max, 10) || 999;
    const v   = Math.min(max, Math.max(1, parseInt(input.value, 10) || 1));
    input.value = v;
    updateFillFromIndexPreview();
  });
  goBtn.addEventListener("click", async () => {
    if (!_lastToFill || _lastToFill.length === 0) {
      setMsg("暂无已解析的题目，请先上传 Word 文件", true);
      return;
    }
    const max   = _lastToFill.length;
    const idx   = Math.min(max, Math.max(1, parseInt(input.value, 10) || 1));
    input.value = idx;
    const slice = _lastToFill.slice(idx - 1);
    await doFillFromNormalized(slice, idx);
  });
})();

// ─── 触发填充 ──────────────────────────────────────────────────────────────
async function doFill(questions, skipTts = false) {
  if (!questions || questions.length === 0) {
    setMsg("没有可填入的题目", true);
    return;
  }
  if (_isFilling) {
    setMsg("正在填充中，请等当前任务完成后再操作。", false);
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

  // TTS 音频合成现在在 content.js 中异步进行，不再在 popup 中阻塞等待
  // 统计需要合成的题目数，仅用于显示提示
  const needsTtsCount = questions.filter(q => {
    const hasScript = (q.listening_script || "").trim().length > 0;
    const hasAudio = (q.audio_url || "").trim().length > 0 || (q.audio_base64 || "").length > 0;
    return hasScript && !hasAudio;
  }).length;
  
  if (needsTtsCount > 0) {
    console.log(`[TTS] 有 ${needsTtsCount} 题需要合成音频，将在填充时异步进行`);
  }

  const { 
    selectors, pageQuestionTotal, pageSlots, defaultAudioUrl, defaultImageUrl, hasTopLevelAudio,
    ttsFemaleVoice, ttsMaleVoice, ttsFemaleSpeed, ttsMaleSpeed, ttsFemaleVolume, ttsMaleVolume
  } = await chrome.storage.sync.get([
    "selectors", "pageQuestionTotal", "pageSlots", "defaultAudioUrl", "defaultImageUrl", "hasTopLevelAudio",
    "ttsFemaleVoice", "ttsMaleVoice", "ttsFemaleSpeed", "ttsMaleSpeed", "ttsFemaleVolume", "ttsMaleVolume"
  ]);
  // 未检测过页面结构时给出提示，但不阻断填充（content.js 内 runFill 会再次 detectForm 新检测）
  const hasStructure = selectors && Object.values(selectors).some(Boolean);
  if (!hasStructure) {
    setMsg("⚠️ 尚未分析页面结构，建议先点「检测页面结构」。尝试直接填入…", false);
  }

  // 先将小题合并为大题（优先用 pageSlots 的 subCount 精确合并，无则按 listening_script 推测）
  const pageTotal = pageQuestionTotal != null && pageQuestionTotal > 0 ? pageQuestionTotal : null;
  let toFill = normalizeQuestionsToSlots(questions, pageTotal != null ? pageTotal : null, pageSlots || null, hasTopLevelAudio);
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

  // 保存规范化后的题目列表，供「从第X题起填充」使用
  _lastToFill    = toFill;
  _lastFillOffset = 0;
  showFillFromIndexUI(toFill);

  setMsg(trimMsg || `共 ${toFill.length} 题，正在填入录题页…`, false);
  // 开始新的填充时隐藏上次遗留的「继续填充」按钮
  const _resumeBtn = document.getElementById("resumeFillBtn");
  if (_resumeBtn) { _resumeBtn.classList.add("hidden"); _resumeBtn.onclick = null; }

  setFillingState(true);

  try {
    const result = await chrome.tabs.sendMessage(tabId, {
      type: "FILL_FORM",
      questions: toFill,
      selectors,
      defaultAudioUrl: (defaultAudioUrl || "").trim() || undefined,
      defaultImageUrl: (defaultImageUrl || "").trim() || undefined,
      ttsSettings: {
        femaleVoice: ttsFemaleVoice || "en_female_amanda_mars_bigtts",
        maleVoice: ttsMaleVoice || "zh_male_jieshuonansheng_mars_bigtts",
        femaleSpeed: parseFloat(ttsFemaleSpeed) || 0.85,
        maleSpeed: parseFloat(ttsMaleSpeed) || 0.85,
        femaleVolume: parseFloat(ttsFemaleVolume) || 1.0,
        maleVolume: parseFloat(ttsMaleVolume) || 1.0,
      },
      debugSource: "parse",
    });
    setFillingState(false);
    if (result?.ok === "paused") {
      showPausedFillUI(result.pausedQuestion, result.remaining || []);
      return;
    }
    if (result?.ok === false) {
      setMsg((result.error || "填充出错") + "。可先点上方「复制题目列表」或「复制完整题目」把识别结果发给我排查。", true);
      bringFillFromIndexIntoView();
      return;
    }
    const done = trimMsg || result?.message || `填充完成，共 ${result?.filled ?? toFill.length} 题。`;
    setMsg(done, false);
  } catch (e) {
    setFillingState(false);
    if (String(e).includes("Could not establish connection") || String(e).includes("receiving end does not exist")) {
      setMsg("录题页未响应，请刷新该页面后再试。可先点上方「复制题目列表」或「复制完整题目」把识别结果发给我排查。", true);
    } else {
      setMsg(`填充失败：${e.message || e}。可先点上方「复制题目列表」或「复制完整题目」把识别结果发给我排查。`, true);
    }
    bringFillFromIndexIntoView();
  }
}

// ─── 上传 Word → 发给 background worker（弹窗关闭不中断）───────────────────
async function doUploadAndFill(filesToUse) {
  if (!filesToUse || filesToUse.length === 0) {
    setMsg("请选择或拖入至少一个 Word 或 PDF 文件", true);
    return;
  }

  const { selectors } = await chrome.storage.sync.get("selectors");
  const hasStructure = selectors && Object.values(selectors).some(Boolean);
  if (!hasStructure) {
    setMsg("", false);
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
    // 用后台持久化的 startedAt 计算真实耗时，避免弹窗关闭/后台节流导致 UI 计时不准
    try {
      const startedAt = parseState.startedAt;
      const subEl = document.getElementById("dzParsingSubtext");
      if (subEl && typeof startedAt === "number" && startedAt > 0) {
        const elapsed = Math.max(0, Math.floor((Date.now() - startedAt) / 1000));
        subEl.textContent = `已等待 ${elapsed} 秒 · 关闭弹窗不会中断`;
      }
    } catch (_) {}
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
    setMsg("", false);
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

    // 先检查是否已经分析过页面结构（上方已有「尚未分析…」提示，仅高亮按钮）
    const { selectors } = await chrome.storage.sync.get("selectors");
    const hasStructure = selectors && Object.values(selectors).some(Boolean);
    if (!hasStructure) {
      setMsg("", false);
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

  // 必须 dragenter + dragover 都 preventDefault 并设置 dropEffect，弹窗内拖拽文件才能生效
  dropZone.addEventListener("dragenter", (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer) e.dataTransfer.dropEffect = "copy";
    if (!dropZone.classList.contains("is-parsing")) dropZone.classList.add("drag-over");
  }, { capture: true });
  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer) e.dataTransfer.dropEffect = "copy";
    if (!dropZone.classList.contains("is-parsing")) dropZone.classList.add("drag-over");
  }, { capture: true });
  dropZone.addEventListener("dragleave", (e) => {
    e.preventDefault();
    // 若进入子元素会误触发 dragleave，仅当真正离开 dropZone 时移除样式
    const rect = dropZone.getBoundingClientRect();
    const x = e.clientX, y = e.clientY;
    if (x <= rect.left || x >= rect.right || y <= rect.top || y >= rect.bottom) {
      dropZone.classList.remove("drag-over");
    }
  });
  dropZone.addEventListener("drop", async (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.remove("drag-over");
    if (dropZone.classList.contains("is-parsing")) return; // 识别中禁止拖入

    // 先检查是否已经分析过页面结构（上方已有「尚未分析…」提示，仅高亮按钮）
    const { selectors } = await chrome.storage.sync.get("selectors");
    const hasStructure = selectors && Object.values(selectors).some(Boolean);
    if (!hasStructure) {
      setMsg("", false);
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

    const rawFiles = getFilesFromDrop(e.dataTransfer);
    const list = getWordFiles(rawFiles);
    if (list.length) {
      setDropZoneState("selected", { files: list });
      doUploadAndFill(list);
    } else if (rawFiles.length > 0) {
      setMsg("请选择 Word (.docx) 或 PDF 文件", true);
    } else {
      // 弹窗不支持拖拽，帮用户打开文件选择
      setMsg("请在下方的文件选择窗口中选择文件", false);
      document.getElementById("wordFiles").click();
    }
  }, { capture: true });
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
  hideFillFromIndexUI();
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
      const wrap = document.createElement("div");
      wrap.className = "json-history-item";
      let preview = item.text.slice(0, 60);
      if (item.text.length > 60) preview += "…";
      const timeStr = item.time ? new Date(item.time).toLocaleString("zh-CN", { month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" }) : "";
      const mainBtn = document.createElement("button");
      mainBtn.type = "button";
      mainBtn.className = "json-history-item-main";
      mainBtn.innerHTML = `<span class="item-preview">${escapeHtml(preview)}</span><span class="item-time">${escapeHtml(timeStr)}</span>`;
      const doFillThis = async () => {
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
      };
      mainBtn.addEventListener("click", async (ev) => {
        ev.preventDefault();
        await doFillThis();
      });
      const fillBtn = document.createElement("button");
      fillBtn.type = "button";
      fillBtn.className = "json-history-item-copy json-history-item-fill";
      fillBtn.title = "一键填入当前页面";
      fillBtn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="5 3 19 12 5 21 5 3"/></svg>`;
      fillBtn.addEventListener("click", async (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        await doFillThis();
      });
      const copyBtn = document.createElement("button");
      copyBtn.type = "button";
      copyBtn.className = "json-history-item-copy";
      copyBtn.title = "复制试题 JSON";
      copyBtn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>`;
      copyBtn.addEventListener("click", async (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        try {
          await navigator.clipboard.writeText(item.text);
          setMsg("已复制试题 JSON 到剪贴板", false);
          setTimeout(() => setMsg("", false), 1500);
        } catch (_) {
          setMsg("复制失败", true);
        }
      });
      wrap.appendChild(mainBtn);
      wrap.appendChild(fillBtn);
      wrap.appendChild(copyBtn);
      if (item.debug_info && item.debug_info.system_prompt) {
        const promptBtn = document.createElement("button");
        promptBtn.type = "button";
        promptBtn.className = "json-history-item-copy";
        promptBtn.title = "复制请求 Prompt";
        promptBtn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>`;
        promptBtn.addEventListener("click", async (ev) => {
          ev.preventDefault();
          ev.stopPropagation();
          const text = "=== System Prompt ===\n\n" + (item.debug_info.system_prompt || "") + "\n\n=== Word 原文 ===\n\n" + (item.debug_info.user_content || "").slice(0, 50000);
          try {
            await navigator.clipboard.writeText(text);
            setMsg("已复制请求 Prompt 到剪贴板", false);
            setTimeout(() => setMsg("", false), 1500);
          } catch (_) {
            setMsg("复制失败", true);
          }
        });
        wrap.appendChild(promptBtn);
      }
      jsonHistoryDropdown.appendChild(wrap);
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
  
  // TTS 音频合成现在在 content.js 中异步进行，填充时会自动合成
  const needsTtsCount = questions.filter(q => {
    const hasScript = (q.listening_script || "").trim().length > 0;
    const hasAudio = (q.audio_url || "").trim().length > 0 || (q.audio_base64 || "").length > 0;
    return hasScript && !hasAudio;
  }).length;
  
  if (needsTtsCount > 0) {
    console.log(`[TTS] 高级面板：有 ${needsTtsCount} 题需要合成音频，将在填充时异步进行`);
  }
  
  const sel = await getSelectorsForJsonPanel(tab.id);
  const { defaultAudioUrl: dau, defaultImageUrl: diu, ttsFemaleVoice, ttsMaleVoice, ttsFemaleSpeed, ttsMaleSpeed, ttsFemaleVolume, ttsMaleVolume } = 
    await chrome.storage.sync.get(["defaultAudioUrl", "defaultImageUrl", "ttsFemaleVoice", "ttsMaleVoice", "ttsFemaleSpeed", "ttsMaleSpeed", "ttsFemaleVolume", "ttsMaleVolume"]);
  try {
    await chrome.tabs.sendMessage(tab.id, {
      type: "FILL_FORM",
      questions,
      selectors: sel,
      defaultAudioUrl: (dau || "").trim() || undefined,
      defaultImageUrl: (diu || "").trim() || undefined,
      ttsSettings: {
        femaleVoice: ttsFemaleVoice || "en_female_amanda_mars_bigtts",
        maleVoice: ttsMaleVoice || "zh_male_jieshuonansheng_mars_bigtts",
        femaleSpeed: parseFloat(ttsFemaleSpeed) || 0.85,
        maleSpeed: parseFloat(ttsMaleSpeed) || 0.85,
        femaleVolume: parseFloat(ttsFemaleVolume) || 1.0,
        maleVolume: parseFloat(ttsMaleVolume) || 1.0,
      },
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
    "- 题干、选项保留原文，不要添加「A.」「选项A：」等前缀；题干内容若开头是题号（如 1. 2. 一、(1) 等）请去掉题号只保留正文",
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
    const detectEl = document.getElementById("detect");
    setDetectResultSimple("正在分析页面结构，请稍候...", "text-muted");

    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id) {
      setDetectResultSimple("无法获取当前标签页", "text-error");
      return;
    }

    const detectBtnIcon = document.querySelector("#detect .card-icon");
    const detectBtnTitle = document.querySelector("#detect .card-title");
    const detectBtnDesc = document.querySelector("#detect .card-desc");

    // 状态更新为分析中（放大镜 + 查找动效）
    if (detectEl) detectEl.classList.add("is-analyzing");
    if (detectBtnIcon) {
      detectBtnIcon.innerHTML = DETECT_ICON_MAGNIFIER;
      detectBtnIcon.classList.add("card-icon--searching");
      detectBtnIcon.classList.remove("icon-spin");
      detectBtnIcon.style.color = "var(--brand-green)";
      detectBtnIcon.style.background = "transparent";
    }
    if (detectBtnTitle) {
      detectBtnTitle.textContent = "分析中...";
      detectBtnTitle.style.color = "var(--brand-green)";
    }
    if (detectBtnDesc) detectBtnDesc.textContent = "正在遍历各题";
    
    try {
      await chrome.tabs.sendMessage(tab.id, { type: "PING" });
    } catch (e) {
      setDetectResultSimple("未连接页面，请刷新录题页后重试", "text-error");
      if (detectBtnIcon) {
        detectBtnIcon.classList.remove("icon-spin", "card-icon--searching");
        detectBtnIcon.innerHTML = DETECT_ICON_MAGNIFIER;
        detectBtnIcon.style.color = "";
        detectBtnIcon.style.background = "";
      }
      if (detectBtnTitle) {
        detectBtnTitle.textContent = "分析页面";
        detectBtnTitle.style.color = "";
      }
      if (detectBtnDesc) detectBtnDesc.textContent = "识别当前表单";
      if (detectEl) detectEl.classList.remove("is-analyzing", "is-analyzed");
      return;
    }

    const { selectors: stored } = await chrome.storage.sync.get("selectors");
    try {
      const res = await chrome.tabs.sendMessage(tab.id, { type: "DETECT_AND_WALK", selectors: stored || {} });

      if (!res || !res.ok) {
        setDetectResultSimple(res?.error || "分析失败，请确认在录题页面", "text-error");
        if (detectBtnIcon) {
          detectBtnIcon.classList.remove("icon-spin", "card-icon--searching");
          detectBtnIcon.innerHTML = DETECT_ICON_MAGNIFIER;
          detectBtnIcon.style.color = "";
          detectBtnIcon.style.background = "";
        }
        if (detectBtnTitle) {
          detectBtnTitle.textContent = "分析页面";
          detectBtnTitle.style.color = "";
        }
        if (detectBtnDesc) detectBtnDesc.textContent = "识别当前表单";
        if (detectEl) detectEl.classList.remove("is-analyzing", "is-analyzed");
        return;
      }

      if (detectBtnIcon) detectBtnIcon.classList.remove("icon-spin", "card-icon--searching");

    const { selectors, fields, message, walked, total, slots, hasTopLevelAudio } = res;
    await chrome.storage.sync.set({
      selectors,
      fields: fields || [],
      lastDetectMessage: message,
      pageQuestionTotal: total != null && total > 0 ? total : null,
      pageSlots: slots && slots.length > 0 ? slimSlotsForStorage(slots) : null,
      hasTopLevelAudio: !!hasTopLevelAudio,
    });
    // 完整 slots 存 local，供后端构建 prompt 用
    chrome.storage.local.set({ pageSlotsFull: slots && slots.length > 0 ? slots : null });

      const totalN = total != null && total > 0 ? `/${total}` : "";
      const msg = message || `成功分析 ${walked}${totalN} 题，了解页面结构，允许上传 Word。`;
      setDetectResultSuccess(msg, fields || [], slots || [], hasTopLevelAudio);
      
      // 更新按钮状态为已识别（三弧 logo 绿色）
      if (detectEl) detectEl.classList.remove("is-analyzing"), detectEl.classList.add("is-analyzed");
      if (detectBtnIcon) {
        detectBtnIcon.classList.remove("icon-spin", "card-icon--searching");
        detectBtnIcon.innerHTML = DETECT_ICON_MAGNIFIER;
        detectBtnIcon.style.color = "var(--brand-green)";
        detectBtnIcon.style.background = "transparent";
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

      // 分析完成后保留检测结果可见，不自动隐藏
    } catch (e) {
      if (detectBtnIcon) {
        detectBtnIcon.classList.remove("icon-spin", "card-icon--searching");
        detectBtnIcon.innerHTML = DETECT_ICON_MAGNIFIER;
        detectBtnIcon.style.color = "";
        detectBtnIcon.style.background = "";
      }
      if (detectBtnTitle) {
        detectBtnTitle.textContent = "分析页面";
        detectBtnTitle.style.color = "";
      }
      if (detectBtnDesc) detectBtnDesc.textContent = "识别当前表单";
      if (detectEl) detectEl.classList.remove("is-analyzing", "is-analyzed");
      
      setDetectResultSimple("分析中断：" + e.message, "text-error");
    }
  });
}

