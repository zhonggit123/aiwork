const tbody = document.getElementById("tbody");
const fileInput = document.getElementById("file");
const uploadSection = document.getElementById("uploadSection");
const parseStatus = document.getElementById("parse-status");
const actions = document.getElementById("actions");
const tableWrap = document.getElementById("table-wrap");
const copyBtn = document.getElementById("copy-json");
const submitBtn = document.getElementById("submit-api");
const submitStatus = document.getElementById("submit-status");

let questions = [];

function setStatus(el, text, isError = false) {
  el.textContent = text;
  el.style.color = isError ? "#f87171" : "#94a3b8";
}

function renderTable() {
  tbody.innerHTML = "";
  questions.forEach((q, i) => {
    const tr = document.createElement("tr");
    const opts = Array.isArray(q.options) ? q.options.join("\n") : (q.options || "");
    tr.innerHTML = `
      <td>${i + 1}</td>
      <td><input type="text" data-i="${i}" data-k="type" value="${escapeAttr(q.type || "")}" /></td>
      <td><textarea data-i="${i}" data-k="question" rows="2">${escapeText(q.question || "")}</textarea></td>
      <td><textarea data-i="${i}" data-k="options" rows="3">${escapeText(opts)}</textarea></td>
      <td><input type="text" data-i="${i}" data-k="answer" value="${escapeAttr(q.answer || "")}" /></td>
      <td><textarea data-i="${i}" data-k="explanation" rows="2">${escapeText(q.explanation || "")}</textarea></td>
    `;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll("input, textarea").forEach((el) => {
    el.addEventListener("change", syncFromCell);
    el.addEventListener("blur", syncFromCell);
  });
}

function escapeAttr(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
function escapeText(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function syncFromCell(e) {
  const el = e.target;
  const i = parseInt(el.dataset.i, 10);
  const k = el.dataset.k;
  if (isNaN(i) || !questions[i]) return;
  let val = el.value;
  if (k === "options") val = val.split("\n").map((s) => s.trim()).filter(Boolean);
  questions[i][k] = val;
}

function isWordFile(file) {
  const n = (file.name || "").toLowerCase();
  return n.endsWith(".docx") || n.endsWith(".doc");
}

function getFilesFromInputOrList(files) {
  return files ? Array.from(files).filter(isWordFile) : [];
}

async function parseAndShow(filesToUse) {
  if (!filesToUse || filesToUse.length === 0) return;
  setStatus(parseStatus, "正在解析，请稍候…");
  const form = new FormData();
  if (filesToUse.length === 1) {
    form.append("file", filesToUse[0]);
  } else {
    filesToUse.forEach((f) => form.append("files", f));
  }
  const endpoint = filesToUse.length === 1 ? "/api/parse" : "/api/parse-multiple";
  try {
    const r = await fetch(endpoint, { method: "POST", body: form });
    const data = await r.json();
    if (!r.ok) throw new Error(data.detail || "解析失败");
    questions = data.questions || [];
    setStatus(parseStatus, `解析完成，共 ${questions.length} 道题`);
    if (questions.length > 0) {
      actions.hidden = false;
      tableWrap.hidden = false;
      renderTable();
    } else {
      actions.hidden = true;
      tableWrap.hidden = true;
      setStatus(parseStatus, "未解析到题目，请检查 Word 内容或格式", true);
    }
  } catch (err) {
    setStatus(parseStatus, err.message || "解析失败", true);
  }
}

fileInput.addEventListener("change", () => {
  const list = getFilesFromInputOrList(fileInput.files);
  if (list.length) parseAndShow(list);
});

if (uploadSection) {
  uploadSection.addEventListener("dragover", (e) => {
    e.preventDefault();
    e.stopPropagation();
    uploadSection.classList.add("drag-over");
  });
  uploadSection.addEventListener("dragleave", (e) => {
    e.preventDefault();
    e.stopPropagation();
    uploadSection.classList.remove("drag-over");
  });
  uploadSection.addEventListener("drop", (e) => {
    e.preventDefault();
    e.stopPropagation();
    uploadSection.classList.remove("drag-over");
    const list = getFilesFromInputOrList(e.dataTransfer.files);
    if (list.length) parseAndShow(list);
    else setStatus(parseStatus, "请拖入 .docx 或 .doc 文件", true);
  });
}

copyBtn.addEventListener("click", () => {
  syncAllFromTable();
  const json = JSON.stringify(questions, null, 2);
  navigator.clipboard.writeText(json).then(
    () => setStatus(parseStatus, "已复制到剪贴板"),
    () => setStatus(parseStatus, "复制失败", true)
  );
});

function syncAllFromTable() {
  tbody.querySelectorAll("input, textarea").forEach((el) => {
    const i = parseInt(el.dataset.i, 10);
    const k = el.dataset.k;
    if (isNaN(i) || !questions[i]) return;
    let val = el.value;
    if (k === "options") val = val.split("\n").map((s) => s.trim()).filter(Boolean);
    questions[i][k] = val;
  });
}

submitBtn.addEventListener("click", async () => {
  syncAllFromTable();
  setStatus(submitStatus, "提交中…");
  try {
    const r = await fetch("/api/submit", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ questions }),
    });
    const data = await r.json();
    if (!r.ok) throw new Error(data.detail || "提交失败");
    setStatus(submitStatus, `成功 ${data.success}/${data.total} 题`);
  } catch (err) {
    setStatus(submitStatus, err.message || "提交失败", true);
  }
});
