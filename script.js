/* =========================================================
   號段分配工具集 — script.js（重構版）
   ---------------------------------------------------------
   架構總覽：
     Utils      共用零依賴工具（型別轉換、非同步輔助、下載）
     UI         共用 DOM 操作（顯示/隱藏、錯誤提示、格式化文字）
     Router     頁面切換
     StackTool  功能① 欄位堆疊 — parse / render / events 三層分離
     AutoTool   功能② MAC 自動整合 — 同上，並用 async/await
                取代原本五層巢狀 setTimeout

   設計原則：
     1. 「資料處理」與「畫面渲染」分開寫，方便日後測試或替換 UI
     2. 全部用 async/await 表達流程順序，不再靠 callback 巢狀縮排
     3. 對外只保留 HTML 既有呼叫的全域函式名稱（goPage、onDrop…），
        內部邏輯都收進對應模組，不再散落成一堆獨立全域變數/函式
     4. 錯誤一律走同一個通知函式，不再讓 alert() 散落各處
   ========================================================= */
'use strict';

/* ══════════════════════════════════════════════════════════
   Utils — 共用工具，不碰 DOM
   ══════════════════════════════════════════════════════════ */
const Utils = {
  /** 任意儲存格值 → 乾淨字串，避免數字被轉成科學記號 */
  toCleanString(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') return Number.isInteger(v) ? v.toString() : String(v);
    return String(v).trim();
  },

  /** 千分位數字 + 單位，例如 1234 → "1,234 筆" */
  formatCount(n, unit) {
    return `${Number(n || 0).toLocaleString()}${unit ? ' ' + unit : ''}`;
  },

  /** 讓瀏覽器有機會重繪一次（取代原本一堆 setTimeout(fn, 100)） */
  tick(ms = 120) {
    return new Promise(resolve => setTimeout(resolve, ms));
  },

  /** 以 Promise 包裝 FileReader，讀成 ArrayBuffer */
  readAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = e => resolve(e.target.result);
      reader.onerror = () => reject(new Error('讀取檔案時發生錯誤，請再試一次。'));
      reader.readAsArrayBuffer(file);
    });
  },

  /** 檔名副檔名檢查 */
  hasExtension(filename, exts) {
    const pattern = new RegExp(`\\.(${exts.join('|')})$`, 'i');
    return pattern.test(filename);
  },

  /** 去除副檔名 */
  stripExtension(filename) {
    return filename.replace(/\.[^.]+$/, '');
  },

  /** 把工作表所有儲存格強制轉成文字格式，避免 Excel 顯示科學記號 */
  forceTextFormat(ws) {
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (cell && cell.v !== '') {
          cell.t = 's';
          cell.v = String(cell.v);
          cell.z = '@';
        }
      }
    }
    return ws;
  },

  /** 觸發瀏覽器下載活頁簿 */
  downloadWorkbook(workbook, filename) {
    XLSX.writeFile(workbook, filename);
  },
};

/** 讓「處理失敗」有一個明確、可辨識的錯誤型別，方便統一攔截 */
class ToolError extends Error {}

/* ══════════════════════════════════════════════════════════
   UI — 共用 DOM 小工具
   ══════════════════════════════════════════════════════════ */
const UI = {
  $(id) { return document.getElementById(id); },
  /**
   * 顯示元素。務必給明確的 display 值（預設 'block'），
   * 不能寫成 el.style.display = ''，那只是清空 inline style，
   * 一旦 CSS 有 display:none 的預設規則，元素會直接退回隱藏狀態。
   */
  show(el, display = 'block') { el.style.display = display; },
  hide(el) { el.style.display = 'none'; },

  /**
   * 在指定容器「之後」插入（或重用）一個錯誤提示框。
   * 直接沿用功能①既有的 .error-box 樣式，兩個功能視覺語言一致，
   * 不需要另外改 HTML／CSS。
   */
  ensureErrorBox(afterEl, id) {
    let box = this.$(id);
    if (!box) {
      box = document.createElement('div');
      box.id = id;
      box.className = 'error-box';
      afterEl.insertAdjacentElement('afterend', box);
    }
    return box;
  },

  setError(box, message) {
    box.textContent = message;
    box.classList.toggle('show', Boolean(message));
  },

  clearError(box) {
    this.setError(box, '');
  },
};

/* ══════════════════════════════════════════════════════════
   Router — 頁面切換
   ══════════════════════════════════════════════════════════ */
const Router = {
  go(id) {
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    UI.$('page-' + id).classList.add('active');
    UI.$('nav-' + id).classList.add('active');
    window.scrollTo(0, 0);
  },
};

/* ══════════════════════════════════════════════════════════
   StackTool — 功能① 欄位堆疊
   把「讀檔」「資料處理」「畫面渲染」「事件綁定」拆成各自的函式，
   process() 只負責串流程，不摻雜運算細節。
   ══════════════════════════════════════════════════════════ */
const StackTool = (() => {
  const ACCEPTED_EXT = ['xlsx', 'xlsm', 'csv'];

  let dom = null;      // 快取常用 DOM 節點
  let errorBox = null;
  let outputWorkbook = null;
  let outputFilename = 'merged.xlsx';

  function cacheDom() {
    dom = {
      drop: UI.$('stDrop'),
      fileInput: UI.$('stFile'),
      result: UI.$('stResult'),
      fname: UI.$('stFname'),
      imeiTotal: UI.$('stImeiTotal'),
      macTotal: UI.$('stMacTotal'),
      imeiBreak: UI.$('stImeiBreak'),
      macBreak: UI.$('stMacBreak'),
      previewBody: UI.$('stPreviewBody'),
      downloadBtn: UI.$('stDownload'),
      resetBtn: UI.$('stReset'),
    };
    errorBox = UI.$('stError');
  }

  /** 讀取原始檔案 → 回傳兩欄（A/B）＋兩欄（C/D）的純資料陣列，不碰 DOM */
  function readSheetColumns(buffer) {
    const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
    if (!workbook.SheetNames.length) {
      throw new ToolError('檔案裡找不到任何工作表。');
    }
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, blankrows: false });
    if (rows.length < 1) {
      throw new ToolError('工作表是空的，沒有資料可以處理。');
    }

    const header = rows[0] || [];
    const body = rows.slice(1);
    const colA = [], colB = [], colC = [], colD = [];
    for (const row of body) {
      const a = Utils.toCleanString(row && row[0]);
      const b = Utils.toCleanString(row && row[1]);
      const c = Utils.toCleanString(row && row[2]);
      const d = Utils.toCleanString(row && row[3]);
      if (a !== '') colA.push(a);
      if (b !== '') colB.push(b);
      if (c !== '') colC.push(c);
      if (d !== '') colD.push(d);
    }
    if (colC.length === 0 && colD.length === 0) {
      throw new ToolError('C 欄和 D 欄都沒有資料，沒有東西可以接到底下。請確認檔案是 A／B 與 C／D 兩組並排的格式。');
    }

    return {
      sheetName,
      headA: Utils.toCleanString(header[0]) || 'IMEI',
      headB: Utils.toCleanString(header[1]) || 'MAC',
      colA, colB, colC, colD,
    };
  }

  /** 把兩組欄位堆疊成單一兩欄，並組出可下載的活頁簿 */
  function buildMergedWorkbook({ sheetName, headA, headB, colA, colB, colC, colD }) {
    const imei = colA.concat(colC);
    const mac = colB.concat(colD);
    const n = Math.max(imei.length, mac.length);

    const aoa = [[headA, headB]];
    for (let i = 0; i < n; i++) aoa.push([imei[i] || '', mac[i] || '']);

    const ws = Utils.forceTextFormat(XLSX.utils.aoa_to_sheet(aoa));
    ws['!cols'] = [{ wch: 20 }, { wch: 18 }];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, ws, sheetName);

    return { workbook, imei, mac, joinAt: colA.length };
  }

  /** 挑出接合處前後幾筆＋頭尾幾筆，組成預覽用的列索引 */
  function pickPreviewIndexes(total, joinAt) {
    const head = [0, 1, 2];
    const around = [joinAt - 1, joinAt, joinAt + 1].filter(i => i >= 0 && i < total);
    const tail = [total - 2, total - 1].filter(i => i >= 0);
    return [...new Set([...head, ...around, ...tail])]
      .filter(i => i >= 0 && i < total)
      .sort((a, b) => a - b);
  }

  function renderStats(counts) {
    dom.imeiTotal.innerHTML = `${counts.imei.length.toLocaleString()} <small>筆</small>`;
    dom.macTotal.innerHTML = `${counts.mac.length.toLocaleString()} <small>筆</small>`;
    dom.imeiBreak.innerHTML = `原 A 欄 <b>${counts.colA.toLocaleString()}</b> ＋ C 欄移入 <b>${counts.colC.toLocaleString()}</b>`;
    dom.macBreak.innerHTML = `原 B 欄 <b>${counts.colB.toLocaleString()}</b> ＋ D 欄移入 <b>${counts.colD.toLocaleString()}</b>`;
  }

  function renderPreview({ imei, mac, joinAt }) {
    const total = Math.max(imei.length, mac.length);
    const indexes = pickPreviewIndexes(total, joinAt);

    dom.previewBody.innerHTML = '';
    let prevIndex = -1;
    for (const i of indexes) {
      if (prevIndex !== -1 && i > prevIndex + 1) {
        const gapRow = document.createElement('tr');
        gapRow.className = 'gap';
        gapRow.innerHTML = '<td colspan="3">· · ·</td>';
        dom.previewBody.appendChild(gapRow);
      }
      const row = document.createElement('tr');
      if (i >= joinAt) row.className = 'from-move';
      row.innerHTML = `<td class="idx">${i + 2}</td><td>${imei[i] || ''}</td><td>${mac[i] || ''}</td>`;
      dom.previewBody.appendChild(row);
      prevIndex = i;
    }
  }

  function renderResult(originalName, parsed, merged) {
    UI.clearError(errorBox);
    dom.fname.textContent = originalName;
    renderStats({
      imei: merged.imei, mac: merged.mac,
      colA: parsed.colA.length, colB: parsed.colB.length,
      colC: parsed.colC.length, colD: parsed.colD.length,
    });
    renderPreview(merged);
    dom.result.classList.add('show');
    dom.result.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  async function handleFile(file) {
    UI.clearError(errorBox);
    if (!file) return;

    if (!Utils.hasExtension(file.name, ACCEPTED_EXT)) {
      UI.setError(errorBox, '請選擇 .xlsx、.xlsm 或 .csv 檔案。');
      return;
    }

    try {
      const buffer = await Utils.readAsArrayBuffer(file);
      const parsed = readSheetColumns(buffer);
      const merged = buildMergedWorkbook(parsed);

      outputWorkbook = merged.workbook;
      outputFilename = Utils.stripExtension(file.name) + '_合併完成.xlsx';

      renderResult(file.name, parsed, merged);
    } catch (err) {
      console.error(err);
      dom.result.classList.remove('show');
      UI.setError(errorBox, err instanceof ToolError ? err.message : '讀取失敗：' + (err.message || '檔案格式無法解析'));
    }
  }

  function reset() {
    outputWorkbook = null;
    dom.fileInput.value = '';
    dom.result.classList.remove('show');
    UI.clearError(errorBox);
    dom.drop.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  function bindEvents() {
    dom.downloadBtn.addEventListener('click', () => {
      if (outputWorkbook) Utils.downloadWorkbook(outputWorkbook, outputFilename);
    });
    dom.resetBtn.addEventListener('click', reset);

    dom.drop.addEventListener('click', () => dom.fileInput.click());
    dom.drop.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); dom.fileInput.click(); }
    });
    dom.fileInput.addEventListener('change', e => handleFile(e.target.files[0]));

    ['dragenter', 'dragover'].forEach(ev =>
      dom.drop.addEventListener(ev, e => { e.preventDefault(); dom.drop.classList.add('dragging'); }));
    ['dragleave', 'drop'].forEach(ev =>
      dom.drop.addEventListener(ev, e => { e.preventDefault(); dom.drop.classList.remove('dragging'); }));
    dom.drop.addEventListener('drop', e => handleFile(e.dataTransfer.files[0]));
  }

  function init() {
    cacheDom();
    bindEvents();
  }

  return { init };
})();

/* ══════════════════════════════════════════════════════════
   AutoTool — 功能② MAC 自動整合
   原本是 5 層巢狀 setTimeout；改成 async/await 依序執行，
   每個步驟仍保留 Utils.tick() 讓畫面上的膠囊狀態看得到變化，
   但邏輯攤平成一直線，不再往右縮排。
   ══════════════════════════════════════════════════════════ */
const AutoTool = (() => {
  const ACCEPTED_EXT = ['xlsx', 'xls'];
  const PREVIEW_LIMIT = 200;
  const SOURCE_SHEET_KEYWORD = '号段分配';
  const OUTPUT_SHEET_NAME = '工廠使用';
  const MAC_GROUP_SIZE = 4;

  let dom = null;
  let errorBox = null;
  let busy = false;

  // 處理完成後的狀態，供下載／預覽使用
  const state = { workbook: null, baseName: '', previewRows: [] };

  function cacheDom() {
    dom = {
      uploadZone: UI.$('uploadZone'),
      fileInput: UI.$('fileInput'),
      processingArea: UI.$('processingArea'),
      resultArea: UI.$('resultArea'),
      fName: UI.$('fName'),
      fInfo: UI.$('fInfo'),
      sumImei: UI.$('sumImei'),
      sumMac: UI.$('sumMac'),
      previewCount: UI.$('previewCount'),
      previewBody: UI.$('previewBody'),
      previewFoot: UI.$('previewFoot'),
      dlMsg: UI.$('dlMsg'),
      steps: [1, 2, 3, 4, 5].map(n => UI.$('ps' + n)),
    };
    errorBox = UI.ensureErrorBox(dom.uploadZone, 'autoError');
  }

  function setStep(n) {
    dom.steps.forEach((el, idx) => {
      const stepNumber = idx + 1;
      el.className = stepNumber < n ? 'p-step done' : stepNumber === n ? 'p-step active' : 'p-step';
    });
  }

  /* ---------- 純資料處理（不碰 DOM） ---------- */

  function findSourceSheet(workbook) {
    const name = workbook.SheetNames.find(n => n.includes(SOURCE_SHEET_KEYWORD));
    if (!name) {
      throw new ToolError(`找不到「${SOURCE_SHEET_KEYWORD}」工作表，請確認檔案格式。`);
    }
    return name;
  }

  function readImeiAndMac(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
    const imeiData = rows.map(r => (r[0] != null ? String(r[0]) : ''));           // 含標題列
    const macRaw = rows.slice(1)
      .map(r => (r[1] != null ? String(r[1]) : ''))
      .filter(v => v);                                                            // 去標題、去空值
    return { imeiData, macRaw };
  }

  function groupMacBySet(macRaw, groupSize) {
    const totalGroups = Math.ceil(macRaw.length / groupSize);
    const groups = [];
    for (let i = 0; i < totalGroups; i++) {
      const offset = i * groupSize;
      groups.push(Array.from({ length: groupSize }, (_, k) => macRaw[offset + k] || ''));
    }
    return groups;
  }

  function removeSheetIfExists(workbook, sheetName) {
    if (!workbook.SheetNames.includes(sheetName)) return;
    const idx = workbook.SheetNames.indexOf(sheetName);
    workbook.SheetNames.splice(idx, 1);
    delete workbook.Sheets[sheetName];
  }

  function buildFactoryRows(imeiData, macGroups) {
    return imeiData.map((imei, i) => {
      if (i === 0) {
        return ['IMEI', 'mac1', 'mac2', 'mac3', 'mac4',
          'pcba sn(按照Thinkstart規則生成)', '組裝sn(按照Thinkstart規則生成)'];
      }
      const group = macGroups[i - 1] || ['', '', '', ''];
      return [imei, group[0], group[1], group[2], group[3], '', ''];
    });
  }

  function buildFactorySheet(rows) {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 18 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 28 }, { wch: 28 }];
    return ws;
  }

  /* ---------- 畫面渲染（不碰運算邏輯） ---------- */

  function renderSummary(file, workbook, imeiCount, macGroupCount, rowCount) {
    dom.fName.textContent = file.name;
    dom.fInfo.textContent = `${workbook.SheetNames.length} 個工作表`;
    dom.sumImei.textContent = Utils.formatCount(imeiCount, '筆');
    dom.sumMac.textContent = Utils.formatCount(macGroupCount, '組');
    dom.previewCount.textContent = Utils.formatCount(rowCount, '列');
  }

  function cellClass(value, emptyClass, filledClass) {
    return value ? filledClass : emptyClass;
  }

  function renderPreview(rows) {
    const limit = Math.min(rows.length, PREVIEW_LIMIT);
    const html = [];
    for (let i = 0; i < limit; i++) {
      const r = rows[i];
      const isHeader = i === 0;
      html.push(
        '<tr>',
        `<td class="cy ${isHeader ? '' : 'cv'}">${r[0] || ''}</td>`,
        `<td class="${cellClass(r[1], 'ce', 'cv')}">${r[1] || ''}</td>`,
        `<td class="${cellClass(r[2], 'ce', 'cv')}">${r[2] || ''}</td>`,
        `<td class="${cellClass(r[3], 'ce', 'cv')}">${r[3] || ''}</td>`,
        `<td class="${cellClass(r[4], 'ce', 'cv')}">${r[4] || ''}</td>`,
        `<td class="ce">${r[5] || ''}</td>`,
        `<td class="ce">${r[6] || ''}</td>`,
        '</tr>'
      );
    }
    dom.previewBody.innerHTML = html.join('');

    const remaining = rows.length - limit;
    dom.previewFoot.textContent = remaining > 0
      ? `顯示前 ${limit.toLocaleString()} / ${rows.length.toLocaleString()} 列`
      : `共 ${rows.length.toLocaleString()} 列`;
  }

  /* ---------- 流程控制 ---------- */

  async function processFile(file) {
    if (busy) return;                 // 防止處理中又觸發一次
    if (!Utils.hasExtension(file.name, ACCEPTED_EXT)) {
      UI.setError(errorBox, '請選擇 .xlsx 或 .xls 檔案。');
      return;
    }

    busy = true;
    UI.clearError(errorBox);
    UI.hide(dom.uploadZone);
    UI.show(dom.processingArea, 'flex');
    UI.hide(dom.resultArea);

    try {
      state.baseName = Utils.stripExtension(file.name);

      setStep(1);
      const buffer = await Utils.readAsArrayBuffer(file);
      await Utils.tick();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sourceSheetName = findSourceSheet(workbook);
      const { imeiData, macRaw } = readImeiAndMac(workbook, sourceSheetName);

      setStep(2);
      await Utils.tick();
      const macGroups = groupMacBySet(macRaw, MAC_GROUP_SIZE);

      setStep(3);
      await Utils.tick();
      removeSheetIfExists(workbook, OUTPUT_SHEET_NAME);

      setStep(4);
      await Utils.tick();
      const rows = buildFactoryRows(imeiData, macGroups);
      const sheet = buildFactorySheet(rows);

      setStep(5);
      await Utils.tick();
      XLSX.utils.book_append_sheet(workbook, sheet, OUTPUT_SHEET_NAME);

      state.workbook = workbook;
      state.previewRows = rows;

      UI.hide(dom.processingArea);
      renderSummary(file, workbook, imeiData.length - 1, macGroups.length, rows.length);
      renderPreview(rows);
      UI.show(dom.resultArea);
    } catch (err) {
      console.error(err);
      UI.setError(errorBox, err instanceof ToolError ? err.message : '處理失敗：' + err.message);
      reset();
    } finally {
      busy = false;
    }
  }

  function download() {
    if (!state.workbook) return;
    dom.dlMsg.className = 'smsg';
    dom.dlMsg.textContent = '生成中...';
    setTimeout(() => {
      try {
        Utils.downloadWorkbook(state.workbook, `${state.baseName}_明細表.xlsx`);
        dom.dlMsg.textContent = '✓ 已下載完成';
      } catch (err) {
        dom.dlMsg.className = 'smsg err';
        dom.dlMsg.textContent = '錯誤：' + err.message;
      }
    }, 50);
  }

  function reset() {
    state.workbook = null;
    state.previewRows = [];
    UI.hide(dom.resultArea);
    UI.hide(dom.processingArea);
    UI.show(dom.uploadZone);
    dom.fileInput.value = '';
    dom.dlMsg.textContent = '';
    UI.clearError(errorBox);
  }

  /* ---------- 事件處理（薄薄一層，轉呼叫上面的流程） ---------- */

  function onDragOver(e) { e.preventDefault(); dom.uploadZone.classList.add('drag'); }
  function onDragLeave() { dom.uploadZone.classList.remove('drag'); }
  function onDrop(e) {
    e.preventDefault();
    onDragLeave();
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  }
  function onFileChange(e) {
    const file = e.target.files[0];
    if (file) processFile(file);
  }

  function init() {
    cacheDom();
  }

  return { init, onDragOver, onDragLeave, onDrop, onFileChange, download, reset };
})();

/* ══════════════════════════════════════════════════════════
   App — 進入點
   ══════════════════════════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', () => {
  StackTool.init();
  AutoTool.init();
});

/* ══════════════════════════════════════════════════════════
   向後相容層 — HTML 目前用 inline onclick/onchange 呼叫這些
   全域函式名稱，這裡只是薄薄轉呼叫對應模組，不含任何邏輯。
   若之後想改成 addEventListener 統一綁定、拿掉 inline handler，
   刪掉這一段、改在 App 裡 querySelector 綁定即可。
   ══════════════════════════════════════════════════════════ */
function goPage(id) { Router.go(id); }
function onDragOver(e) { AutoTool.onDragOver(e); }
function onDragLeave() { AutoTool.onDragLeave(); }
function onDrop(e) { AutoTool.onDrop(e); }
function onFileChange(e) { AutoTool.onFileChange(e); }
function downloadResult() { AutoTool.download(); }
function resetAll() { AutoTool.reset(); }
