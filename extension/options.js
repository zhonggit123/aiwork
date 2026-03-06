const keys = ["question", "option_a", "option_b", "option_c", "option_d", "answer", "explanation", "submit_btn", "next_btn"];

chrome.storage.sync.get(["selectors", "recordPagePattern"], (data) => {
  const s = data.selectors || {};
  keys.forEach((k) => {
    const el = document.getElementById(k);
    if (el && s[k]) el.value = s[k];
  });
  const urlEl = document.getElementById("urlPattern");
  if (urlEl) urlEl.value = data.recordPagePattern || "";
});

document.getElementById("save").addEventListener("click", () => {
  const selectors = {};
  keys.forEach((k) => {
    const el = document.getElementById(k);
    if (el && el.value.trim()) selectors[k] = el.value.trim();
  });
  const urlPattern = (document.getElementById("urlPattern") || {}).value;
  chrome.storage.sync.set({ selectors, recordPagePattern: (urlPattern || "").trim() }, () => {
    document.getElementById("saved").style.display = "block";
    setTimeout(() => (document.getElementById("saved").style.display = "none"), 2000);
  });
});
