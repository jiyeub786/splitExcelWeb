const sheetTaskTableBody = document.querySelector("#sheetTaskTable tbody");
const addSheetTaskRowBtn = document.getElementById("addSheetTaskRow");
const sheetNamesList = document.getElementById("sheetNamesList");
const splitListArea = document.getElementById("splitListArea");
const splitListLineNumbers = document.getElementById("splitListLineNumbers");
const splitListCount = document.getElementById("splitListCount");
const resultFileNmInput = document.getElementById("resultFileNm");
const resultFileDateInput = document.getElementById("resultFileDate");
const logPanel = document.getElementById("logPanel");
const statusLine = document.getElementById("statusLine");
const startBtn = document.getElementById("startBtn");
const testBtn = document.getElementById("testBtn");
const cancelBtn = document.getElementById("cancelBtn");
const forceStopBtn = document.getElementById("forceStopBtn");
const downloadLink = document.getElementById("downloadLink");
const saveConfigBtn = document.getElementById("saveConfigBtn");
const loadConfigInput = document.getElementById("loadConfigInput");

const sourceFileInput = document.getElementById("sourceFile");
const templateFileInput = document.getElementById("templateFile");
const templateSameAsSource = document.getElementById("templateSameAsSource");
const templateFileField = document.getElementById("templateFileField");

const extractSheetName = document.getElementById("extractSheetName");
const extractColumn = document.getElementById("extractColumn");
const extractHeaderRows = document.getElementById("extractHeaderRows");
const extractColumnBtn = document.getElementById("extractColumnBtn");
const extractStatus = document.getElementById("extractStatus");

const progressTrack = document.getElementById("progressTrack");
const progressFill = document.getElementById("progressFill");
const progressLabel = document.getElementById("progressLabel");

const rangePickerModal = document.getElementById("rangePickerModal");
const rangePickerTitle = document.getElementById("rangePickerTitle");
const rangePickerHint = document.getElementById("rangePickerHint");
const rangePickerBody = document.getElementById("rangePickerBody");
const rangePickerClose = document.getElementById("rangePickerClose");
const rangePickerApply = document.getElementById("rangePickerApply");
const rangePickerSelection = document.getElementById("rangePickerSelection");
const rangePickerExtendLastRow = document.getElementById("rangePickerExtendLastRow");

let currentJobId = null;
let currentEventSource = null;
let currentJobTotal = 0;
let currentJobCompleted = 0;

/* ---------- 시트 작업 정의 테이블 ---------- */

function addSheetTaskRow(data) {
  const row = document.createElement("tr");
  row.innerHTML = `
    <td><input type="text" class="sheet-name" list="sheetNamesList" value="${data?.sheet_name ?? ""}" title="${data?.sheet_name ?? ""}" /></td>
    <td><input type="text" class="filter-column" value="${colLetter(data?.filter_index ?? 1)}" placeholder="예: A" /></td>
    <td>
      <div class="range-cell">
        <input type="text" class="copy-range" value="${data?.copy_range ?? ""}" placeholder="A1:D500" />
        <button type="button" class="pick-range-btn secondary" data-target="copy" title="미리보기에서 복사 범위 선택">선택</button>
      </div>
    </td>
    <td>
      <div class="range-cell">
        <input type="text" class="paste-range" value="${data?.paste_range ?? ""}" placeholder="A2" />
        <button type="button" class="pick-range-btn secondary" data-target="paste" title="미리보기에서 붙여넣기 위치 선택">선택</button>
      </div>
    </td>
    <td><button type="button" class="remove-row-btn icon-btn" title="행 삭제">−</button></td>
  `;
  row.querySelector(".remove-row-btn").addEventListener("click", () => row.remove());
  row.querySelectorAll(".pick-range-btn").forEach((btn) => {
    btn.addEventListener("click", () => openRangePicker(btn.dataset.target, row));
  });
  sheetTaskTableBody.appendChild(row);
}

addSheetTaskRowBtn.addEventListener("click", () => addSheetTaskRow());
addSheetTaskRow(); // 최초 1행 기본 제공

// 시트명이 좁은 칸에서 잘려도 마우스오버로 전체 값을 볼 수 있도록 title 동기화
sheetTaskTableBody.addEventListener("input", (e) => {
  if (e.target.classList.contains("sheet-name")) {
    e.target.title = e.target.value;
  }
});

function columnIndexFromLetters(letters) {
  let idx = 0;
  for (const ch of letters.toUpperCase()) {
    if (ch < "A" || ch > "Z") return NaN;
    idx = idx * 26 + (ch.charCodeAt(0) - 64);
  }
  return idx;
}

function parseFilterColumn(raw) {
  const v = raw.trim();
  if (/^[A-Za-z]+$/.test(v)) return columnIndexFromLetters(v);
  return Number(v);
}

function collectSheetTasks() {
  return Array.from(sheetTaskTableBody.querySelectorAll("tr")).map((row) => ({
    sheet_name: row.querySelector(".sheet-name").value.trim(),
    filter_index: parseFilterColumn(row.querySelector(".filter-column").value),
    copy_range: row.querySelector(".copy-range").value.trim(),
    paste_range: row.querySelector(".paste-range").value.trim(),
  }));
}

function collectSplitList() {
  return splitListArea.value
    .split("\n")
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
}

/* 분리 대상 목록 줄번호 + 개수 표시 + 내용에 맞춘 높이 자동 조절 */
const SPLIT_LIST_MAX_HEIGHT = 220; // px, 이보다 길어지면 내부 스크롤

function updateSplitListGutter() {
  const lineCount = splitListArea.value.split("\n").length;
  const numbers = [];
  for (let i = 1; i <= lineCount; i++) numbers.push(i);
  splitListLineNumbers.textContent = numbers.join("\n");
  splitListCount.textContent = `${collectSplitList().length}개`;

  // 컨테이너가 flex(align-items: flex-start)라 서로 높이를 밀어주지 않으므로, 두 요소를 모두
  // auto로 되돌린 뒤 textarea의 실제 콘텐츠 높이(scrollHeight)를 측정해 둘 다 그 값으로 맞춘다.
  splitListArea.style.height = "auto";
  splitListLineNumbers.style.height = "auto";
  const fitHeight = Math.min(splitListArea.scrollHeight, SPLIT_LIST_MAX_HEIGHT);
  splitListArea.style.height = `${fitHeight}px`;
  splitListLineNumbers.style.height = `${fitHeight}px`;
}

splitListArea.addEventListener("input", updateSplitListGutter);
splitListArea.addEventListener("scroll", () => {
  splitListLineNumbers.scrollTop = splitListArea.scrollTop;
});
updateSplitListGutter();

function collectOptions() {
  return {
    zoom_level1: Number(document.getElementById("zoomLevel1").value),
    zoom_level2: Number(document.getElementById("zoomLevel2").value),
    hide_guideline_yn: document.getElementById("hideGuideline").value,
    remove_formula_path_yn: document.getElementById("removeFormulaPath").value,
  };
}

function collectConfig(testMode) {
  return {
    result_file_nm: document.getElementById("resultFileNm").value.trim(),
    result_file_date: document.getElementById("resultFileDate").value.trim(),
    sheet_tasks: collectSheetTasks(),
    split_list: collectSplitList(),
    options: collectOptions(),
    test_mode: testMode,
  };
}

function applyConfig(cfg) {
  document.getElementById("resultFileNm").value = cfg.result_file_nm ?? "";
  document.getElementById("resultFileDate").value = cfg.result_file_date ?? "";
  document.getElementById("zoomLevel1").value = cfg.options?.zoom_level1 ?? 120;
  document.getElementById("zoomLevel2").value = cfg.options?.zoom_level2 ?? 100;
  document.getElementById("hideGuideline").value = cfg.options?.hide_guideline_yn ?? "Y";
  document.getElementById("removeFormulaPath").value = cfg.options?.remove_formula_path_yn ?? "Y";
  splitListArea.value = (cfg.split_list ?? []).join("\n");

  sheetTaskTableBody.innerHTML = "";
  (cfg.sheet_tasks ?? []).forEach((t) => addSheetTaskRow(t));
  if ((cfg.sheet_tasks ?? []).length === 0) addSheetTaskRow();

  updateSplitListGutter();
}

saveConfigBtn.addEventListener("click", () => {
  const cfg = collectConfig(false);
  const blob = new Blob([JSON.stringify(cfg, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "splitexcel_config.json";
  a.click();
  URL.revokeObjectURL(url);
});

loadConfigInput.addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  try {
    const text = await file.text();
    applyConfig(JSON.parse(text));
  } catch (err) {
    alert("설정 파일을 읽을 수 없습니다: " + err);
  }
  loadConfigInput.value = "";
});

/* ---------- 드래그 앤 드롭 업로드 ---------- */

function setupDropzone(dropzoneEl, inputEl, fileNameEl, onFileChanged) {
  const showFileName = () => {
    const file = inputEl.files[0];
    fileNameEl.textContent = file ? file.name : "";
    dropzoneEl.classList.toggle("has-file", !!file);
    if (onFileChanged) onFileChanged(file);
  };

  dropzoneEl.addEventListener("click", () => inputEl.click());
  dropzoneEl.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      inputEl.click();
    }
  });
  inputEl.addEventListener("change", showFileName);

  ["dragenter", "dragover"].forEach((evt) =>
    dropzoneEl.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzoneEl.classList.add("dragover");
    })
  );
  ["dragleave", "dragend"].forEach((evt) =>
    dropzoneEl.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzoneEl.classList.remove("dragover");
    })
  );
  dropzoneEl.addEventListener("drop", (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropzoneEl.classList.remove("dragover");
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      inputEl.files = files;
      showFileName();
    }
  });
}

function todayYYYYMMDD() {
  const d = new Date();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${d.getFullYear()}${mm}${dd}`;
}

if (!resultFileDateInput.value.trim()) resultFileDateInput.value = todayYYYYMMDD();

setupDropzone(
  document.getElementById("sourceDropzone"),
  sourceFileInput,
  document.getElementById("sourceFileName"),
  (file) => {
    if (!file) return;
    fetchSheetNames(file);
    if (!resultFileNmInput.value.trim()) {
      resultFileNmInput.value = file.name.replace(/\.[^./]+$/, "");
    }
    if (!resultFileDateInput.value.trim()) {
      resultFileDateInput.value = todayYYYYMMDD();
    }
  }
);
setupDropzone(document.getElementById("templateDropzone"), templateFileInput, document.getElementById("templateFileName"));

templateSameAsSource.addEventListener("change", () => {
  templateFileField.style.display = templateSameAsSource.checked ? "none" : "";
});

function getTemplateFile() {
  return templateSameAsSource.checked ? sourceFileInput.files[0] : templateFileInput.files[0];
}

/* ---------- 시트명 자동완성 ---------- */

async function fetchSheetNames(file) {
  try {
    const formData = new FormData();
    formData.append("file", file);
    const res = await fetch("/api/preview/sheets", { method: "POST", body: formData });
    if (!res.ok) return;
    const data = await res.json();
    sheetNamesList.innerHTML = "";
    (data.sheets ?? []).forEach((name) => {
      const opt = document.createElement("option");
      opt.value = name;
      sheetNamesList.appendChild(opt);
    });
  } catch (err) {
    // 미리보기 실패는 치명적이지 않으므로 조용히 무시(자동완성만 못 받음)
  }
}

/* ---------- 분리 대상 목록 자동 추출 ---------- */

extractColumnBtn.addEventListener("click", async () => {
  const file = sourceFileInput.files[0];
  if (!file) return alert("원본파일을 먼저 선택해주세요");
  const sheetName = extractSheetName.value.trim();
  const column = extractColumn.value.trim();
  if (!sheetName) return alert("시트명을 입력해주세요");
  if (!column) return alert("열을 입력해주세요 (예: C 또는 3)");

  extractStatus.textContent = "추출 중...";
  const formData = new FormData();
  formData.append("file", file);
  formData.append("sheet_name", sheetName);
  formData.append("column", column);
  formData.append("header_rows", extractHeaderRows.value.trim() || "1");

  try {
    const res = await fetch("/api/preview/column-values", { method: "POST", body: formData });
    const data = await res.json();
    if (!res.ok) {
      extractStatus.textContent = "";
      return alert("추출 실패: " + (data.detail ?? res.status));
    }
    if (splitListArea.value.trim() && !confirm(`분리 대상 목록을 ${data.count}개 값으로 덮어쓸까요?`)) {
      extractStatus.textContent = "";
      return;
    }
    splitListArea.value = data.values.join("\n");
    updateSplitListGutter();
    extractStatus.textContent = `${data.count}개 값을 채웠습니다`;
  } catch (err) {
    extractStatus.textContent = "";
    alert("추출 실패: " + err);
  }
});

/* ---------- 범위 선택 뷰어 ---------- */

const rangePickerState = {
  mode: null, // 'copy' | 'paste'
  rowEl: null,
  anchor: null, // {r, c} (1-based)
  focus: null,
  dragging: false,
  sheetMaxRow: null, // 시트의 실제 마지막 데이터 행("마지막 행까지 확장" 버튼에 사용)
};

// 드래그 중 그리드 가장자리에 마우스가 닿으면 자동으로 스크롤(긴 시트를 끝까지 드래그하기 위함).
// 셀 하나를 클릭한 뒤 Shift+클릭으로 먼 셀을 바로 지정하는 것도 같은 문제의 대안 경로.
const AUTO_SCROLL_EDGE = 36;
const AUTO_SCROLL_STEP = 22;
let lastPointer = { x: 0, y: 0 };
let autoScrollTimer = null;

document.addEventListener("mousemove", (e) => {
  lastPointer = { x: e.clientX, y: e.clientY };
});

function startAutoScroll(container, table) {
  if (autoScrollTimer) return;
  autoScrollTimer = setInterval(() => {
    if (!rangePickerState.dragging) {
      stopAutoScroll();
      return;
    }
    const rect = container.getBoundingClientRect();
    let dx = 0;
    let dy = 0;
    if (lastPointer.y < rect.top + AUTO_SCROLL_EDGE) dy = -AUTO_SCROLL_STEP;
    else if (lastPointer.y > rect.bottom - AUTO_SCROLL_EDGE) dy = AUTO_SCROLL_STEP;
    if (lastPointer.x < rect.left + AUTO_SCROLL_EDGE) dx = -AUTO_SCROLL_STEP;
    else if (lastPointer.x > rect.right - AUTO_SCROLL_EDGE) dx = AUTO_SCROLL_STEP;
    if (dx === 0 && dy === 0) return;

    container.scrollBy(dx, dy);
    const el = document.elementFromPoint(lastPointer.x, lastPointer.y);
    const td = el && el.closest ? el.closest("td") : null;
    if (td && td.dataset.row) {
      rangePickerState.focus = { r: Number(td.dataset.row), c: Number(td.dataset.col) };
      updateGridSelectionHighlight(table);
    }
  }, 40);
}

function stopAutoScroll() {
  if (autoScrollTimer) {
    clearInterval(autoScrollTimer);
    autoScrollTimer = null;
  }
}

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function openRangePicker(mode, rowEl) {
  const sheetName = rowEl.querySelector(".sheet-name").value.trim();
  if (!sheetName) return alert("먼저 이 행의 시트명을 입력해주세요");

  const file = mode === "copy" ? sourceFileInput.files[0] : getTemplateFile();
  if (!file) return alert(mode === "copy" ? "원본파일을 먼저 선택해주세요" : "양식파일을 먼저 선택해주세요(또는 '원본파일과 동일' 체크)");

  rangePickerState.mode = mode;
  rangePickerState.rowEl = rowEl;
  rangePickerState.anchor = null;
  rangePickerState.focus = null;
  rangePickerState.sheetMaxRow = null;

  rangePickerTitle.textContent = mode === "copy" ? "복사 범위 선택 (원본파일)" : "붙여넣기 위치 선택 (양식파일)";
  rangePickerHint.textContent =
    mode === "copy"
      ? "셀을 클릭한 채로 드래그하면 범위가 지정됩니다. 화면 가장자리에 닿으면 자동으로 스크롤되고, 첫 셀을 클릭한 뒤 끝 셀을 Shift+클릭해도 됩니다. 데이터가 아주 길면 시작 셀만 선택한 뒤 \"마지막 행까지 확장\"을 눌러도 됩니다."
      : "붙여넣기를 시작할 셀 하나를 클릭하세요.";
  rangePickerSelection.textContent = "";
  rangePickerExtendLastRow.hidden = mode !== "copy";
  rangePickerBody.innerHTML = '<p class="hint">불러오는 중...</p>';
  rangePickerModal.hidden = false;

  const formData = new FormData();
  formData.append("file", file);
  formData.append("sheet_name", sheetName);
  formData.append("max_rows", "200");
  formData.append("max_cols", "40");

  fetch("/api/preview/grid", { method: "POST", body: formData })
    .then(async (res) => {
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail ?? res.status);
      renderGridPreview(data);
    })
    .catch((err) => {
      rangePickerBody.innerHTML = `<p class="hint">미리보기를 불러오지 못했습니다: ${err}</p>`;
    });
}

function renderGridPreview(data) {
  rangePickerState.sheetMaxRow = data.max_row || null;
  const rows = data.rows ?? [];
  const colCount = rows.reduce((max, r) => Math.max(max, r.length), 0);

  const table = document.createElement("table");
  table.className = "grid-preview-table";

  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  headRow.appendChild(document.createElement("th"));
  for (let c = 1; c <= colCount; c++) {
    const th = document.createElement("th");
    th.textContent = colLetter(c);
    headRow.appendChild(th);
  }
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  rows.forEach((rowValues, rIdx) => {
    const tr = document.createElement("tr");
    const rowHeadTh = document.createElement("th");
    rowHeadTh.textContent = String(rIdx + 1);
    tr.appendChild(rowHeadTh);

    for (let c = 1; c <= colCount; c++) {
      const td = document.createElement("td");
      td.textContent = rowValues[c - 1] ?? "";
      td.dataset.row = rIdx + 1;
      td.dataset.col = c;
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  rangePickerBody.innerHTML = "";
  rangePickerBody.appendChild(table);

  if (data.truncated_rows || data.truncated_cols) {
    const note = document.createElement("p");
    note.className = "hint";
    note.textContent = `실제 시트는 ${data.max_row}행 × ${data.max_column}열입니다. 미리보기는 앞부분 일부만 표시합니다.`;
    rangePickerBody.appendChild(note);
  }

  table.addEventListener("mousedown", (e) => {
    const td = e.target.closest("td");
    if (!td) return;
    e.preventDefault();
    const cell = { r: Number(td.dataset.row), c: Number(td.dataset.col) };
    if (e.shiftKey && rangePickerState.anchor && rangePickerState.mode === "copy") {
      // 첫 셀을 클릭한 뒤 Shift+클릭으로 먼 셀까지 바로 범위를 지정(드래그 없이도 가능하게).
      rangePickerState.focus = cell;
      updateGridSelectionHighlight(table);
      return;
    }
    rangePickerState.anchor = cell;
    rangePickerState.focus = cell;
    rangePickerState.dragging = rangePickerState.mode === "copy";
    updateGridSelectionHighlight(table);
    if (rangePickerState.dragging) startAutoScroll(rangePickerBody, table);
  });
  table.addEventListener("mouseover", (e) => {
    if (!rangePickerState.dragging) return;
    const td = e.target.closest("td");
    if (!td) return;
    rangePickerState.focus = { r: Number(td.dataset.row), c: Number(td.dataset.col) };
    updateGridSelectionHighlight(table);
  });
}

document.addEventListener("mouseup", () => {
  rangePickerState.dragging = false;
  stopAutoScroll();
});

function updateGridSelectionHighlight(table) {
  const { anchor, focus } = rangePickerState;
  if (!anchor || !focus) return;
  const minR = Math.min(anchor.r, focus.r);
  const maxR = Math.max(anchor.r, focus.r);
  const minC = Math.min(anchor.c, focus.c);
  const maxC = Math.max(anchor.c, focus.c);

  table.querySelectorAll("td").forEach((td) => {
    const r = Number(td.dataset.row);
    const c = Number(td.dataset.col);
    td.classList.toggle("selected", r >= minR && r <= maxR && c >= minC && c <= maxC);
  });

  rangePickerSelection.textContent =
    rangePickerState.mode === "copy"
      ? `선택: ${colLetter(minC)}${minR}:${colLetter(maxC)}${maxR}`
      : `선택: ${colLetter(anchor.c)}${anchor.r}`;
}

rangePickerExtendLastRow.addEventListener("click", () => {
  const { anchor, focus, mode, sheetMaxRow } = rangePickerState;
  if (mode !== "copy") return;
  if (!anchor || !focus) return alert("먼저 시작 셀(열 범위)을 클릭이나 드래그로 선택해주세요");
  if (!sheetMaxRow) return alert("시트의 마지막 행 정보를 아직 불러오지 못했습니다");

  const minR = Math.min(anchor.r, focus.r);
  const minC = Math.min(anchor.c, focus.c);
  const maxC = Math.max(anchor.c, focus.c);
  rangePickerState.anchor = { r: minR, c: minC };
  rangePickerState.focus = { r: sheetMaxRow, c: maxC };

  const table = rangePickerBody.querySelector("table");
  if (table) updateGridSelectionHighlight(table);
});

rangePickerApply.addEventListener("click", () => {
  const { mode, rowEl, anchor, focus } = rangePickerState;
  if (!anchor || !focus || !rowEl) return alert("먼저 셀을 선택해주세요");

  if (mode === "copy") {
    const minR = Math.min(anchor.r, focus.r);
    const maxR = Math.max(anchor.r, focus.r);
    const minC = Math.min(anchor.c, focus.c);
    const maxC = Math.max(anchor.c, focus.c);
    rowEl.querySelector(".copy-range").value = `${colLetter(minC)}${minR}:${colLetter(maxC)}${maxR}`;
  } else {
    rowEl.querySelector(".paste-range").value = `${colLetter(anchor.c)}${anchor.r}`;
  }
  closeRangePicker();
});

rangePickerClose.addEventListener("click", closeRangePicker);
rangePickerModal.addEventListener("click", (e) => {
  if (e.target === rangePickerModal) closeRangePicker();
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && !rangePickerModal.hidden) closeRangePicker();
});

function closeRangePicker() {
  stopAutoScroll();
  rangePickerModal.hidden = true;
  rangePickerState.mode = null;
  rangePickerState.rowEl = null;
  rangePickerState.anchor = null;
  rangePickerState.focus = null;
  rangePickerState.dragging = false;
  rangePickerState.sheetMaxRow = null;
}

/* ---------- 로그 / 진행률 ---------- */

function appendLog(line) {
  logPanel.textContent += line + "\n";
  logPanel.scrollTop = logPanel.scrollHeight;

  if (/ 완료\. 처리 시간: \d+초/.test(line) && !line.includes("모든작업")) {
    currentJobCompleted += 1;
    updateProgress();
  }
}

function resetProgress(total) {
  currentJobTotal = total;
  currentJobCompleted = 0;
  progressTrack.style.display = total > 0 ? "" : "none";
  progressLabel.style.display = total > 0 ? "" : "none";
  updateProgress();
}

function updateProgress() {
  if (currentJobTotal <= 0) return;
  const pct = Math.min(100, Math.round((currentJobCompleted / currentJobTotal) * 100));
  progressFill.style.width = pct + "%";
  progressLabel.textContent = `진행: ${currentJobCompleted}/${currentJobTotal} (${pct}%)`;
}

function setRunningUI(running) {
  startBtn.disabled = running;
  testBtn.disabled = running;
  cancelBtn.disabled = !running;
  forceStopBtn.disabled = !running;
}

async function submitJob(testMode) {
  const sourceFile = sourceFileInput.files[0];
  const templateFile = getTemplateFile();

  if (!sourceFile) return alert("원본파일을 선택해주세요");
  if (!templateFile) return alert("양식파일을 선택해주세요 (또는 '원본파일과 동일' 체크)");

  const cfg = collectConfig(testMode);
  if (!cfg.result_file_nm) return alert("저장 명칭을 입력해주세요");
  if (cfg.sheet_tasks.length === 0) return alert("시트 작업 정의를 1개 이상 입력해주세요");
  if (cfg.sheet_tasks.some((t) => !Number.isFinite(t.filter_index) || t.filter_index < 1)) {
    return alert("필터 열을 올바르게 입력해주세요 (예: A, B, C ... 또는 숫자)");
  }
  if (cfg.split_list.length === 0) return alert("분리 대상 목록을 1개 이상 입력해주세요");

  const formData = new FormData();
  formData.append("source_file", sourceFile);
  formData.append("template_file", templateFile);
  formData.append("config", JSON.stringify(cfg));

  logPanel.textContent = "";
  downloadLink.style.display = "none";
  statusLine.textContent = "작업 등록 중...";
  resetProgress(testMode ? 1 : cfg.split_list.length);

  const res = await fetch("/api/jobs", { method: "POST", body: formData });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    statusLine.textContent = "등록 실패";
    return alert("작업 등록 실패: " + (err.detail ?? res.status));
  }

  const { job_id } = await res.json();
  currentJobId = job_id;
  statusLine.textContent = `실행 중 (job_id: ${job_id})`;
  setRunningUI(true);
  connectLogStream(job_id);
  pollJobStatus(job_id);
}

function connectLogStream(jobId) {
  if (currentEventSource) currentEventSource.close();
  const es = new EventSource(`/api/jobs/${jobId}/logs`);
  currentEventSource = es;

  es.onmessage = (event) => appendLog(event.data);
  es.addEventListener("done", (event) => {
    appendLog(`[스트림 종료] 최종 상태: ${event.data}`);
    es.close();
  });
  es.onerror = () => {
    es.close();
  };
}

async function pollJobStatus(jobId) {
  const poll = async () => {
    const res = await fetch(`/api/jobs/${jobId}`);
    if (!res.ok) return;
    const job = await res.json();

    if (["done", "error", "cancelled"].includes(job.status)) {
      statusLine.textContent = `완료: ${job.status}` + (job.error ? ` (${job.error})` : "");
      setRunningUI(false);
      if (job.result_ready) {
        downloadLink.href = `/api/jobs/${jobId}/result`;
        downloadLink.style.display = "inline-block";
      }
      return;
    }
    statusLine.textContent = `실행 중... (${job.status})`;
    setTimeout(poll, 1000);
  };
  poll();
}

startBtn.addEventListener("click", () => submitJob(false));
testBtn.addEventListener("click", () => submitJob(true));

cancelBtn.addEventListener("click", async () => {
  if (!currentJobId) return;
  await fetch(`/api/jobs/${currentJobId}/cancel`, { method: "POST" });
  statusLine.textContent = "취소 요청됨...";
});

/* ---------- 다크/라이트 테마 전환 ---------- */

const THEME_STORAGE_KEY = "splitexcelweb-theme";
const themeToggle = document.getElementById("themeToggle");

function getCurrentTheme() {
  return document.documentElement.getAttribute("data-theme") === "light" ? "light" : "dark";
}

function applyThemeButtonLabel(theme) {
  themeToggle.textContent = theme === "light" ? "☀️ 라이트" : "🌙 다크";
}

function setTheme(theme) {
  document.documentElement.setAttribute("data-theme", theme);
  try {
    localStorage.setItem(THEME_STORAGE_KEY, theme);
  } catch (e) {
    // localStorage를 못 쓰는 환경(프라이빗 모드 등)이어도 화면 전환 자체는 계속 동작해야 함
  }
  applyThemeButtonLabel(theme);
}

applyThemeButtonLabel(getCurrentTheme());
themeToggle.addEventListener("click", () => {
  setTheme(getCurrentTheme() === "light" ? "dark" : "light");
});

/* ---------- 설명서 ---------- */

const helpBtn = document.getElementById("helpBtn");
const helpModal = document.getElementById("helpModal");
const helpClose = document.getElementById("helpClose");

helpBtn.addEventListener("click", () => {
  helpModal.hidden = false;
});
helpClose.addEventListener("click", () => {
  helpModal.hidden = true;
});
helpModal.addEventListener("click", (e) => {
  if (e.target === helpModal) helpModal.hidden = true;
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && !helpModal.hidden) helpModal.hidden = true;
});

forceStopBtn.addEventListener("click", async () => {
  if (!currentJobId) return;
  if (!confirm("Excel 프로세스를 강제 종료합니다. 계속할까요?")) return;
  const res = await fetch(`/api/jobs/${currentJobId}/force-stop-excel`, { method: "POST" });
  const data = await res.json();
  statusLine.textContent = `강제 종료 요청됨 (종료된 프로세스: ${data.killed_processes ?? 0}개)`;
});
