/* ──────────────────────────────────────────
   PhiloArgument — taskpane.js
   Office JS Add-in for Analytic Philosophy
   ────────────────────────────────────────── */

"use strict";

/* ─── State ──────────────────────────────── */
const state = {
  selectedType: null,
  premiseCount: 2,
  subPremiseCount: 1,
  recentSymbols: JSON.parse(localStorage.getItem("philo_recent_syms") || "[]"),
};

/* ─── Block metadata ─────────────────────── */
const BLOCK_META = {
  claim:     { label: "주장 (Claim)",     prefix: "C",  color: "2E7DD6", bg: "E8F0FA" },
  premise:   { label: "전제 (Premise)",   prefix: "P",  color: "3A9464", bg: "E6F4ED" },
  support:   { label: "근거 (Support)",   prefix: "S",  color: "7C52B8", bg: "F0EAF9" },
  objection: { label: "반론 (Objection)", prefix: "O",  color: "C85040", bg: "FDECEA" },
  reply:     { label: "재반론 (Reply)",   prefix: "R",  color: "C49030", bg: "FDF4E3" },
  remark:    { label: "논평 (Remark)",    prefix: "N",  color: "888888", bg: "F5F5F5" },
};

/* ─── Office.js init ─────────────────────── */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initTabs();
    initCompose();
    initFormalize();
    initSymbols();
    initReview();
    renderRecentSymbols();
  }
});

/* ─── Tabs ───────────────────────────────── */
function initTabs() {
  document.querySelectorAll(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach(b => {
        b.classList.remove("active");
        b.setAttribute("aria-selected", "false");
      });
      document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
      btn.classList.add("active");
      btn.setAttribute("aria-selected", "true");
      document.getElementById(`tab-${btn.dataset.tab}`).classList.add("active");
    });
  });
}

/* ─── Tab 1: Compose ─────────────────────── */
function initCompose() {
  document.querySelectorAll(".block-type-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".block-type-btn").forEach(b => b.classList.remove("selected"));
      btn.classList.add("selected");
      state.selectedType = btn.dataset.type;

      const showCounterex = ["claim", "premise"].includes(state.selectedType);
      document.getElementById("counterexample-row").style.display = showCounterex ? "block" : "none";
    });
  });

  document.getElementById("btn-preview-block").addEventListener("click", previewBlock);
  document.getElementById("btn-insert-block").addEventListener("click", insertBlock);
}

function previewBlock() {
  const type = state.selectedType;
  if (!type) { showStatus("블록 유형을 먼저 선택하세요."); return; }

  const label   = document.getElementById("block-label-input").value.trim();
  const content = document.getElementById("block-content-input").value.trim();
  const counterex = document.getElementById("counterexample-input").value.trim();

  if (!content) { showStatus("내용을 입력하세요."); return; }

  const meta = BLOCK_META[type];
  const displayLabel = label || meta.prefix + "?";

  const preview = document.getElementById("block-preview");
  const previewContent = document.getElementById("preview-content");

  let html = `<div class="doc-block doc-block-${type}">
    <div class="doc-block-label">[${displayLabel}] ${meta.label}</div>
    <div>${escapeHtml(content)}</div>`;

  if (counterex) {
    html += `<div class="doc-block-counterex">
      <div class="doc-block-counterex-label">⚠ 반례 메모</div>
      ${escapeHtml(counterex)}
    </div>`;
  }
  html += `</div>`;

  previewContent.innerHTML = html;
  preview.style.display = "block";
}

async function insertBlock() {
  const type = state.selectedType;
  if (!type) { showStatus("블록 유형을 먼저 선택하세요."); return; }

  const label     = document.getElementById("block-label-input").value.trim();
  const content   = document.getElementById("block-content-input").value.trim();
  const counterex = document.getElementById("counterexample-input").value.trim();

  if (!content) { showStatus("내용을 입력하세요."); return; }

  const meta = BLOCK_META[type];
  const displayLabel = label || meta.prefix + "?";

  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("isEmpty");
    await context.sync();

    // Insert a paragraph after selection
    const para = range.insertParagraph("", Word.InsertLocation.after);

    // Style the paragraph
    para.style = "Normal";
    para.leftIndent = 18;
    para.rightIndent = 0;
    para.spaceAfter = 6;
    para.spaceBefore = 6;

    // Build the content runs
    // Type label run (bold, colored)
    const labelRun = para.insertText(`[${displayLabel}] ${meta.label}\n`, Word.InsertLocation.start);
    labelRun.bold = true;
    labelRun.font.size = 9;
    labelRun.font.color = meta.color;

    // Body run
    const bodyRun = para.insertText(content, Word.InsertLocation.end);
    bodyRun.font.color = "000000";
    bodyRun.font.size = 11;
    bodyRun.bold = false;

    // Shading / highlight approximation using a bookmarked content control
    const cc = para.insertContentControl();
    cc.tag = `philo-${type}`;
    cc.title = `${meta.label} ${displayLabel}`;
    cc.appearance = Word.ContentControlAppearance.boundingBox;
    cc.cannotDelete = false;
    cc.cannotEdit = false;

    // If counterexample, add a second paragraph
    if (counterex) {
      const cePara = para.insertParagraph("", Word.InsertLocation.after);
      cePara.leftIndent = 36;
      cePara.spaceBefore = 0;
      cePara.spaceAfter = 8;

      const ceLabel = cePara.insertText("⚠ 반례 메모: ", Word.InsertLocation.start);
      ceLabel.bold = true;
      ceLabel.font.size = 9;
      ceLabel.font.color = "C49030";

      const ceBody = cePara.insertText(counterex, Word.InsertLocation.end);
      ceBody.font.size = 10;
      ceBody.font.color = "7A5C1E";
      ceBody.italic = true;
    }

    await context.sync();
    showStatus(`"${displayLabel}" 블록이 삽입되었습니다.`);
  }).catch(err => showStatus("오류: " + err.message));
}

/* ─── Tab 2: Formalize ───────────────────── */
function initFormalize() {
  document.getElementById("btn-add-premise").addEventListener("click", () => {
    state.premiseCount++;
    addPremiseRow("premises-container", state.premiseCount, false);
  });
  document.getElementById("btn-add-subpremise").addEventListener("click", () => {
    state.subPremiseCount++;
    addPremiseRow("subpremises-container", state.subPremiseCount, true);
  });
  document.getElementById("premises-container").addEventListener("click", (e) => {
    if (e.target.classList.contains("btn-remove-premise")) {
      e.target.closest(".premise-row").remove();
      reindexPremises("premises-container", false);
    }
  });
  document.getElementById("subpremises-container").addEventListener("click", (e) => {
    if (e.target.classList.contains("btn-remove-subpremise")) {
      e.target.closest(".premise-row").remove();
      reindexPremises("subpremises-container", true);
    }
  });

  // Live validity check on input
  document.addEventListener("input", (e) => {
    if (e.target.classList.contains("premise-input") || e.target.id === "conclusion-input") {
      runValidityCheck();
    }
  });

  document.getElementById("btn-insert-formal").addEventListener("click", insertFormalArgument);
}

function addPremiseRow(containerId, index, isSub) {
  const container = document.getElementById(containerId);
  const row = document.createElement("div");
  row.className = "premise-row";
  row.dataset.index = index;
  const prefix = isSub ? "P2." : "P";
  const removeCls = isSub ? "btn-remove-subpremise" : "btn-remove-premise";
  row.innerHTML = `
    <span class="premise-num">${prefix}${index}</span>
    <input type="text" class="form-input premise-input" placeholder="${index}번째 전제"/>
    <button class="btn-icon ${removeCls}" title="삭제" aria-label="이 전제 삭제">✕</button>`;
  container.appendChild(row);
}

function reindexPremises(containerId, isSub) {
  const rows = document.querySelectorAll(`#${containerId} .premise-row`);
  const prefix = isSub ? "P2." : "P";
  rows.forEach((row, i) => {
    row.dataset.index = i + 1;
    row.querySelector(".premise-num").textContent = `${prefix}${i + 1}`;
  });
}

function runValidityCheck() {
  const premises = [...document.querySelectorAll("#premises-container .premise-input")]
    .map(i => i.value.trim()).filter(Boolean);
  const conclusion = document.getElementById("conclusion-input").value.trim();

  const icon = document.getElementById("validity-icon");
  const text = document.getElementById("validity-text");

  if (!premises.length && !conclusion) {
    icon.textContent = "○";
    text.textContent = "전제와 결론을 입력하면 형식 점검이 실행됩니다.";
    text.style.color = "";
    return;
  }
  if (!premises.length) {
    icon.textContent = "⚠";
    text.textContent = "전제가 없습니다. 최소 하나의 전제가 필요합니다.";
    text.style.color = "#923a2a";
    return;
  }
  if (!conclusion) {
    icon.textContent = "⚠";
    text.textContent = "결론이 비어 있습니다.";
    text.style.color = "#923a2a";
    return;
  }
  // Heuristic: check if conclusion shares key terms with premises
  const premiseWords = new Set(
    premises.join(" ").toLowerCase().replace(/[^a-z가-힣\s]/g, "").split(/\s+/).filter(w => w.length > 1)
  );
  const conclusionWords = conclusion.toLowerCase().replace(/[^a-z가-힣\s]/g, "").split(/\s+/).filter(w => w.length > 1);
  const overlap = conclusionWords.filter(w => premiseWords.has(w));

  if (overlap.length === 0) {
    icon.textContent = "⚠";
    text.textContent = "결론과 전제 사이에 공유된 용어가 없습니다. 논증의 연결을 확인하세요.";
    text.style.color = "#7a5c1e";
  } else {
    icon.textContent = "✓";
    text.textContent = `전제 ${premises.length}개, 결론 1개. 형식상 연결이 확인됩니다.`;
    text.style.color = "#2d6e4e";
  }
}

async function insertFormalArgument() {
  const title      = document.getElementById("arg-title").value.trim();
  const premises   = [...document.querySelectorAll("#premises-container .premise-input")].map(i => i.value.trim());
  const conclusion = document.getElementById("conclusion-input").value.trim();

  if (premises.every(p => !p) || !conclusion) {
    showStatus("전제와 결론을 모두 입력하세요.");
    return;
  }

  await Word.run(async (context) => {
    const sel = context.document.getSelection();

    // Title paragraph (optional)
    if (title) {
      const titlePara = sel.insertParagraph(title, Word.InsertLocation.after);
      titlePara.bold = true;
      titlePara.font.size = 11;
      titlePara.spaceAfter = 4;
      titlePara.spaceBefore = 10;
    }

    // Premises
    premises.forEach((p, i) => {
      if (!p) return;
      const para = sel.insertParagraph(`P${i + 1}.\t${p}`, Word.InsertLocation.after);
      para.font.size = 11;
      para.leftIndent = 24;
      para.spaceBefore = 2;
      para.spaceAfter = 2;
    });

    // Separator line (via a borderBottom paragraph hack)
    const sepPara = sel.insertParagraph("", Word.InsertLocation.after);
    sepPara.leftIndent = 24;
    sepPara.spaceBefore = 2;
    sepPara.spaceAfter = 2;
    sepPara.bottomBorderVisible = true;  // Only works if borderType API available

    // Conclusion
    const conPara = sel.insertParagraph(`∴\t${conclusion}`, Word.InsertLocation.after);
    conPara.font.size = 11;
    conPara.font.color = "2d6e4e";
    conPara.bold = true;
    conPara.leftIndent = 24;
    conPara.spaceBefore = 4;
    conPara.spaceAfter = 10;

    await context.sync();
    showStatus("정식 논증이 삽입되었습니다.");
  }).catch(err => showStatus("오류: " + err.message));
}

/* ─── Tab 3: Symbols ─────────────────────── */
function initSymbols() {
  // Click to insert
  document.querySelectorAll(".sym-btn").forEach(btn => {
    btn.addEventListener("click", () => insertSymbol(btn.dataset.sym, btn));
  });

  // Search filter
  document.getElementById("symbol-search").addEventListener("input", (e) => {
    const q = e.target.value.toLowerCase().trim();
    document.querySelectorAll(".symbol-group").forEach(group => {
      let groupVisible = false;
      group.querySelectorAll(".sym-btn").forEach(btn => {
        const match = !q || btn.dataset.name.toLowerCase().includes(q) || btn.dataset.sym.includes(q);
        btn.style.display = match ? "" : "none";
        if (match) groupVisible = true;
      });
      group.style.display = groupVisible ? "" : "none";
    });
  });
}

async function insertSymbol(sym, btn) {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertText(sym, Word.InsertLocation.replace);
    await context.sync();

    // Visual feedback
    btn.classList.add("just-used");
    setTimeout(() => btn.classList.remove("just-used"), 600);

    // Toast
    showSymbolToast(`"${sym}" 삽입됨`);

    // Recent symbols
    addRecentSymbol(sym, btn.dataset.name);
  }).catch(err => {
    showSymbolToast("삽입 실패: " + err.message);
  });
}

function addRecentSymbol(sym, name) {
  state.recentSymbols = [{ sym, name }, ...state.recentSymbols.filter(s => s.sym !== sym)].slice(0, 12);
  localStorage.setItem("philo_recent_syms", JSON.stringify(state.recentSymbols));
  renderRecentSymbols();
}

function renderRecentSymbols() {
  const row  = document.getElementById("recent-symbols-row");
  const grid = document.getElementById("recent-symbols-grid");
  if (!state.recentSymbols.length) { row.style.display = "none"; return; }
  row.style.display = "block";
  grid.innerHTML = state.recentSymbols.map(s =>
    `<button class="sym-btn" data-sym="${s.sym}" data-name="${s.name}">${s.sym}</button>`
  ).join("");
  grid.querySelectorAll(".sym-btn").forEach(btn => {
    btn.addEventListener("click", () => insertSymbol(btn.dataset.sym, btn));
  });
}

function showSymbolToast(msg) {
  const t = document.getElementById("symbol-toast");
  t.textContent = msg;
  t.style.display = "block";
  setTimeout(() => { t.style.display = "none"; }, 1500);
}

/* ─── Tab 4: Review ──────────────────────── */
function initReview() {
  document.getElementById("btn-scan").addEventListener("click", scanDocument);
}

async function scanDocument() {
  const btn = document.getElementById("btn-scan");
  btn.textContent = "스캔 중…";
  btn.disabled = true;

  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const contentControls = context.document.contentControls;
      contentControls.load("items/tag,items/title,items/text");
      await context.sync();

      const blocks = contentControls.items.map(cc => ({
        tag: cc.tag || "",
        title: cc.title || "",
        text: cc.text || "",
      })).filter(b => b.tag.startsWith("philo-"));

      // Also scan paragraphs for bracket pattern [P1] [C1] etc.
      const paragraphs = body.paragraphs;
      paragraphs.load("items/text");
      await context.sync();

      const bracketBlocks = [];
      paragraphs.items.forEach(p => {
        const m = p.text.match(/^\[([A-Za-z0-9.]+)\]\s+(주장|전제|근거|반론|재반론|논평)/);
        if (m) bracketBlocks.push({ label: m[1], typeName: m[2], text: p.text });
      });

      displayScanResults(blocks, bracketBlocks, paragraphs.items.length);
    });
  } catch (err) {
    showStatus("스캔 오류: " + err.message);
  } finally {
    btn.textContent = "문서 스캔";
    btn.disabled = false;
  }
}

function displayScanResults(ccBlocks, bracketBlocks, totalParas) {
  const resultsEl = document.getElementById("scan-results");
  resultsEl.style.display = "block";

  const allBlocks = [...ccBlocks, ...bracketBlocks];
  const claims     = ccBlocks.filter(b => b.tag === "philo-claim").length    + bracketBlocks.filter(b => b.typeName === "주장").length;
  const premises   = ccBlocks.filter(b => b.tag === "philo-premise").length  + bracketBlocks.filter(b => b.typeName === "전제").length;
  const objections = ccBlocks.filter(b => b.tag === "philo-objection").length + bracketBlocks.filter(b => b.typeName === "반론").length;
  const remarks    = ccBlocks.filter(b => b.tag === "philo-remark").length   + bracketBlocks.filter(b => b.typeName === "논평").length;

  document.getElementById("scan-summary").innerHTML = `
    <strong>스캔 완료</strong> — 총 ${totalParas}개 단락 분석<br>
    주장 <strong>${claims}</strong>개 &nbsp;|&nbsp;
    전제 <strong>${premises}</strong>개 &nbsp;|&nbsp;
    반론 <strong>${objections}</strong>개 &nbsp;|&nbsp;
    논평 <strong>${remarks}</strong>개
  `;

  // Build issues
  const issues = [];
  if (claims === 0) issues.push({ level: "error", msg: "주장(Claim)이 없습니다. 논증의 핵심 주장을 명시하세요." });
  if (claims > 0 && premises === 0) issues.push({ level: "warn", msg: `주장이 ${claims}개 있지만 전제(Premise)가 없습니다. 근거를 추가하세요.` });
  if (objections > 0) {
    const replies = ccBlocks.filter(b => b.tag === "philo-reply").length;
    if (replies < objections) {
      issues.push({ level: "warn", msg: `반론이 ${objections}개 있으나 재반론은 ${replies}개입니다. 미응답 반론을 검토하세요.` });
    }
  }
  if (allBlocks.length === 0) issues.push({ level: "warn", msg: "PhiloArgument 블록이 발견되지 않았습니다. '작성' 탭에서 블록을 삽입하세요." });
  if (issues.length === 0) issues.push({ level: "ok", msg: "발견된 구조적 문제가 없습니다." });

  const ul = document.getElementById("issues-ul");
  ul.innerHTML = issues.map(iss => `
    <li class="issue-item ${iss.level}">
      <span class="issue-icon">${iss.level === "error" ? "✕" : iss.level === "warn" ? "⚠" : "✓"}</span>
      <span>${iss.msg}</span>
    </li>`).join("");

  // Structure tree
  const treeEl = document.getElementById("tree-content");
  if (allBlocks.length === 0) {
    treeEl.textContent = "블록 없음";
    return;
  }
  const typeColors = {
    "philo-claim": "#2e7dd6", "philo-premise": "#3a9464", "philo-support": "#7c52b8",
    "philo-objection": "#c85040", "philo-reply": "#c49030", "philo-remark": "#888",
    "주장": "#2e7dd6", "전제": "#3a9464", "반론": "#c85040", "재반론": "#c49030",
    "근거": "#7c52b8", "논평": "#888"
  };
  treeEl.innerHTML = allBlocks.map((b, i) => {
    const key = b.tag || b.typeName;
    const color = typeColors[key] || "#888";
    const name = b.title || b.label || `블록 ${i + 1}`;
    const snippet = (b.text || "").slice(0, 40) + ((b.text || "").length > 40 ? "…" : "");
    return `<div class="tree-node">
      <span class="node-type-badge" style="background:${color}20;color:${color}">${b.tag ? b.tag.replace("philo-","") : b.typeName}</span>
      <span style="font-weight:600">${name}</span>
      ${snippet ? `<span style="color:#888"> — ${escapeHtml(snippet)}</span>` : ""}
    </div>`;
  }).join("");
}

/* ─── Utilities ──────────────────────────── */
function showStatus(msg) {
  const bar  = document.getElementById("status-bar");
  const text = document.getElementById("status-text");
  text.textContent = msg;
  bar.style.display = "block";
  setTimeout(() => { bar.style.display = "none"; }, 3000);
}

function escapeHtml(str) {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
