// ── Zoom ──────────────────────────────────────────────────────────────────────

let currentZoom = 100;

function adjustZoom(delta) {
  currentZoom = Math.max(70, Math.min(400, currentZoom + delta));
  document.getElementById('zoom-label').textContent = currentZoom + '%';
  pywebview.api.apply_zoom(currentZoom);
}

window.addEventListener('pywebviewready', async () => {
  const saved = await pywebview.api.get_zoom();
  if (saved && saved !== 100) {
    currentZoom = saved;
    document.getElementById('zoom-label').textContent = currentZoom + '%';
    pywebview.api.apply_zoom(currentZoom);
  }
});

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ── Status & Toast ────────────────────────────────────────────────────────────

function setStatus(msg, ok) {
  const bar = document.getElementById('statusbar');
  bar.textContent = msg;
  bar.className = 'statusbar ' + (ok ? 'ok' : 'err');
}

let _toastTimer = null;
function showToast(msg, ok, duration = 3500) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = `toast show ${ok ? 'ok' : 'err'}`;
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => { t.className = 'toast'; }, duration);
}

// ── Unsaved changes tracking ──────────────────────────────────────────────────

let isDirty = false;

function markDirty() {
  isDirty = true;
  try { pywebview.api.set_dirty(true); } catch(e) {}
}
function markClean() {
  isDirty = false;
  try { pywebview.api.set_dirty(false); } catch(e) {}
}

// Warn before closing the app (works on Windows Edge)
window.onbeforeunload = e => {
  if (isDirty) { e.preventDefault(); e.returnValue = ''; return ''; }
};

// ── Confirm modal ─────────────────────────────────────────────────────────────

function showConfirm(msg, onYes, yesLabel = '確定放棄') {
  document.getElementById('confirm-msg').textContent = msg;
  document.getElementById('confirm-yes').textContent = yesLabel;
  document.getElementById('confirm-modal').classList.add('show');
  document.getElementById('confirm-yes').onclick = () => { hideConfirm(); onYes(); };
}
function hideConfirm() {
  document.getElementById('confirm-modal').classList.remove('show');
}

// Close modal on backdrop click or Escape key
document.getElementById('confirm-modal').addEventListener('click', e => {
  if (e.target === document.getElementById('confirm-modal')) hideConfirm();
});
document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && document.getElementById('confirm-modal').classList.contains('show')) {
    hideConfirm();
    e.stopImmediatePropagation();  // don't also clear the page selection
  }
});

// ── Password prompt modal ─────────────────────────────────────────────────────

// Resolves to the entered password, or null if cancelled
function askPassword(msg) {
  const modal  = document.getElementById('pw-modal');
  const input  = document.getElementById('pw-modal-input');
  document.getElementById('pw-modal-msg').textContent = msg;
  input.value = '';
  modal.classList.add('show');
  input.focus();
  return new Promise(resolve => {
    const done = value => {
      modal.classList.remove('show');
      input.onkeydown = null;
      resolve(value);
    };
    document.getElementById('pw-modal-ok').onclick     = () => done(input.value);
    document.getElementById('pw-modal-cancel').onclick = () => done(null);
    input.onkeydown = e => {
      if (e.key === 'Enter')  done(input.value);
      if (e.key === 'Escape') done(null);
    };
  });
}

// ── Editor state ──────────────────────────────────────────────────────────────

let editorPages = []; // [{srcFile, origIndex, b64, rotation, selected, deleted}]
let dragSrcIdx  = null;
let selectAnchor = null;  // last clicked page, start of a Shift+click range

// ── Undo / Redo ───────────────────────────────────────────────────────────────

const UNDO_LIMIT = 100;
const undoStack = [];
const redoStack = [];

function snapshot() {
  return editorPages.map(p => ({ page: p, rotation: p.rotation, deleted: p.deleted }));
}

function restoreSnapshot(snap) {
  editorPages = snap.map(s => {
    s.page.rotation = s.rotation;
    s.page.deleted  = s.deleted;
    return s.page;
  });
  markDirty();
  renderEditorGrid();
}

// Call before every change to page order / rotation / deletion
function pushUndo() {
  undoStack.push(snapshot());
  if (undoStack.length > UNDO_LIMIT) undoStack.shift();
  redoStack.length = 0;
}

// Record an undo step, apply the change, and refresh
function applyEdit(change) {
  pushUndo();
  change();
  markDirty();
  renderEditorGrid();
}

function undo() {
  if (!undoStack.length) return;
  redoStack.push(snapshot());
  restoreSnapshot(undoStack.pop());
}

function redo() {
  if (!redoStack.length) return;
  undoStack.push(snapshot());
  restoreSnapshot(redoStack.pop());
}

function updateUndoButtons() {
  document.getElementById('btn-undo').disabled = !undoStack.length;
  document.getElementById('btn-redo').disabled = !redoStack.length;
}

// ── File operations ───────────────────────────────────────────────────────────

async function openNewPdf() {
  if (isDirty && editorPages.filter(p => !p.deleted).length > 0) {
    showConfirm('有未儲存的變更，開啟新 PDF 將會遺失目前的修改，確定繼續嗎？', async () => {
      const path = await pywebview.api.browse_open();
      if (!path) return;
      await loadPdf(path, false);
    });
    return;
  }
  const path = await pywebview.api.browse_open();
  if (!path) return;
  await loadPdf(path, false);
}

async function appendPdf() {
  const path = await pywebview.api.browse_open();
  if (!path) return;
  await loadPdf(path, true);
}

async function loadPdf(path, append) {
  const btnOpen   = document.querySelector('.btn-primary');
  const btnAppend = document.getElementById('btn-append');
  btnOpen.disabled = btnAppend.disabled = true;

  setStatus('載入中…', true);
  const grid = document.getElementById('editor-grid');
  if (!append) {
    grid.innerHTML = '<div class="editor-loading"><div class="spinner"></div><span>載入頁面中，請稍候…</span></div>';
  }

  const fileName = fileNameOf(path);
  let res = await pywebview.api.open_pdf_for_editor(path, '');
  while (res && res.need_password) {
    const pw = await askPassword(`「${fileName}」${res.msg}`);
    if (pw === null) { res = { ok: false, msg: '已取消開啟加密檔案' }; break; }
    res = await pywebview.api.open_pdf_for_editor(path, pw);
  }
  btnOpen.disabled = false;
  btnAppend.disabled = editorPages.length === 0;

  if (!res || !res.ok) {
    const msg = res ? res.msg : '未知錯誤';
    if (append || editorPages.length) renderEditorGrid();
    else grid.innerHTML = `<div class="editor-placeholder"><p>載入失敗：${escapeHtml(msg)}</p></div>`;
    return setStatus(msg, false);
  }

  const newPages = res.pages.map(p => ({
    srcFile:   path,
    origIndex: p.index,
    b64:       p.b64,
    aspect:    p.w / p.h,
    rotation:  0,
    selected:  false,
    deleted:   false,
  }));

  if (append) {
    pushUndo();
    editorPages = [...editorPages, ...newPages];
    markDirty();
  } else {
    editorPages = newPages;
    undoStack.length = redoStack.length = 0;
    selectAnchor = null;
    markClean();
  }

  // Show toolbar elements
  document.getElementById('btn-append').disabled = false;
  document.getElementById('page-tools-sep').style.display = '';
  document.getElementById('page-tools').style.display = 'flex';
  document.getElementById('save-tools').style.display = 'flex';

  renderEditorGrid();
  const action = append ? `已附加 ${newPages.length} 頁` : `已載入 ${newPages.length} 頁`;
  setStatus(`${action}，共 ${editorPages.filter(p=>!p.deleted).length} 頁`, true);
}

// True while the output path came from the save dialog (not typed by hand)
let outputFromDialog = false;
document.getElementById('output-path').addEventListener('input', () => { outputFromDialog = false; });

async function browseOutput() {
  const path = await pywebview.api.browse_save('edited_output');
  if (path) {
    document.getElementById('output-path').value = path;
    outputFromDialog = true;
  }
}

function fileNameOf(path) {
  return path.split(/[/\\]/).pop();
}

// Validates the output path; resolves to the normalised path, or '' if the
// user backed out (invalid path, or declined to overwrite an existing file)
async function resolveOutput(path, fromDialog, allowSource = true) {
  const chk = await pywebview.api.check_output(path);
  if (!chk.ok) { showToast(chk.msg, false); return ''; }
  if (!allowSource && editorPages.some(p => p.srcFile === chk.path)) {
    showToast('不能覆寫正在編輯的來源檔案，請另存新檔', false);
    return '';
  }
  if (!chk.exists || (fromDialog && chk.dialog_confirms_overwrite)) return chk.path;
  return new Promise(resolve => {
    showConfirm(`「${fileNameOf(chk.path)}」已存在，確定要覆寫嗎？`, () => resolve(chk.path), '覆寫');
    // Cancel / backdrop / Escape all go through hideConfirm
    const modal = document.getElementById('confirm-modal');
    const obs = new MutationObserver(() => {
      if (!modal.classList.contains('show')) { obs.disconnect(); resolve(''); }
    });
    obs.observe(modal, { attributes: true, attributeFilter: ['class'] });
  });
}

// ── Encrypt toggle ────────────────────────────────────────────────────────────

function toggleEncrypt() {
  const checked = document.getElementById('encrypt-check').checked;
  document.getElementById('encrypt-pw').style.display = checked ? 'block' : 'none';
}

// ── Render ────────────────────────────────────────────────────────────────────

const THUMB_W = 120;

// Box the thumbnail occupies after rotation (width fixed, height follows aspect)
function thumbWrapStyle(page) {
  const a = page.aspect || 0.75;
  const h = page.rotation % 180 === 0 ? THUMB_W / a : THUMB_W * a;
  return `width:${THUMB_W}px;height:${h.toFixed(1)}px`;
}

// Image keeps its unrotated size; rotating it around the centre fills the box
function thumbImgStyle(page) {
  const a = page.aspect || 0.75;
  const [w, h] = page.rotation % 180 === 0 ? [THUMB_W, THUMB_W / a] : [THUMB_W * a, THUMB_W];
  return `width:${w.toFixed(1)}px;height:${h.toFixed(1)}px;` +
         `transform:translate(-50%,-50%) rotate(${page.rotation}deg)`;
}

function renderEditorGrid() {
  const grid = document.getElementById('editor-grid');
  grid.innerHTML = '';

  const visible = editorPages.filter(p => !p.deleted);
  if (!visible.length) {
    grid.innerHTML = '<div class="editor-placeholder"><p>所有頁面已刪除</p></div>';
    updateInfo();
    return;
  }

  // Check if pages come from more than one source file
  const srcFiles = [...new Set(visible.map(p => p.srcFile))];
  const multiSrc = srcFiles.length > 1;

  visible.forEach((page, displayIdx) => {
    const realIdx = editorPages.indexOf(page);
    const srcName = multiSrc ? fileNameOf(page.srcFile) : '';

    const div = document.createElement('div');
    div.className = 'page-thumb' + (page.selected ? ' selected' : '');
    div.draggable = true;

    div.innerHTML = `
      <div class="thumb-img-wrap" style="${thumbWrapStyle(page)}">
        <img src="${page.b64}" draggable="false" style="${thumbImgStyle(page)}">
        ${page.rotation ? `<div class="rotation-badge">↻ ${page.rotation}°</div>` : ''}
        <div class="thumb-overlay">
          <button class="thumb-btn" title="旋轉 90°" onclick="rotateOnePage(${realIdx},event)">↻</button>
          <button class="thumb-btn delete" title="刪除" onclick="deleteOnePage(${realIdx},event)">✕</button>
        </div>
      </div>
      <div class="thumb-footer">
        <span class="thumb-num">${displayIdx + 1}</span>
        ${srcName ? `<span class="thumb-src" title="${escapeHtml(page.srcFile)}">${escapeHtml(srcName)}</span>` : ''}
      </div>`;

    // Select on click
    div.addEventListener('click', e => {
      if (e.shiftKey && selectAnchor && visible.includes(selectAnchor)) {
        const [a, b] = [visible.indexOf(selectAnchor), displayIdx].sort((x, y) => x - y);
        visible.slice(a, b + 1).forEach(p => { p.selected = true; });
        renderEditorGrid();
        return;
      }
      page.selected = !page.selected;
      selectAnchor = page;
      div.classList.toggle('selected', page.selected);
      updateInfo();
    });

    // Drag-and-drop
    div.addEventListener('dragstart', e => {
      dragSrcIdx = realIdx;
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', String(realIdx));
      setTimeout(() => {
        div.classList.add('dragging');
        document.getElementById('editor-grid').classList.add('is-dragging');
      }, 0);
    });

    div.addEventListener('dragend', () => {
      dragSrcIdx = null;
      div.classList.remove('dragging');
      document.getElementById('editor-grid').classList.remove('is-dragging');
      document.querySelectorAll('.page-thumb').forEach(el => el.classList.remove('drag-over'));
    });

    div.addEventListener('dragover', e => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      if (dragSrcIdx === realIdx) return;
      document.querySelectorAll('.page-thumb').forEach(el => el.classList.remove('drag-over'));
      div.classList.add('drag-over');
    });

    div.addEventListener('drop', e => {
      e.preventDefault();
      document.querySelectorAll('.page-thumb').forEach(el => el.classList.remove('drag-over'));
      if (dragSrcIdx === null || dragSrcIdx === realIdx) return;
      const rect = div.getBoundingClientRect();
      const insertBefore = e.clientX < rect.left + rect.width / 2;
      const from = dragSrcIdx;
      dragSrcIdx = null;
      applyEdit(() => {
        const moved = editorPages.splice(from, 1)[0];
        let insertAt = from < realIdx ? realIdx - 1 : realIdx;
        if (!insertBefore) insertAt = Math.min(insertAt + 1, editorPages.length);
        editorPages.splice(insertAt, 0, moved);
      });
    });

    grid.appendChild(div);
  });

  // End drop zone — lets pages be dragged to the very last position
  const endZone = document.createElement('div');
  endZone.className = 'end-drop-zone';
  endZone.textContent = '放置於此';
  endZone.addEventListener('dragover', e => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    endZone.classList.add('drag-active');
  });
  endZone.addEventListener('dragleave', () => {
    endZone.classList.remove('drag-active');
  });
  endZone.addEventListener('drop', e => {
    e.preventDefault();
    endZone.classList.remove('drag-active');
    if (dragSrcIdx === null) return;
    const from = dragSrcIdx;
    dragSrcIdx = null;
    applyEdit(() => editorPages.push(editorPages.splice(from, 1)[0]));
  });
  grid.appendChild(endZone);

  updateInfo();
}

function updateInfo() {
  updateUndoButtons();
  const active   = editorPages.filter(p => !p.deleted);
  const selected = active.filter(p => p.selected);
  let text = `共 ${active.length} 頁`;
  if (selected.length) text += `　已選取 ${selected.length} 頁`;
  setStatus(text, true);
}

// ── Page actions ──────────────────────────────────────────────────────────────

function rotateOnePage(idx, e) {
  e.stopPropagation();
  applyEdit(() => { editorPages[idx].rotation = (editorPages[idx].rotation + 90) % 360; });
}

function deleteOnePage(idx, e) {
  e.stopPropagation();
  applyEdit(() => { editorPages[idx].deleted = true; });
}

function editorSelectAll() {
  const active = editorPages.filter(p => !p.deleted);
  const allSel = active.every(p => p.selected);
  active.forEach(p => p.selected = !allSel);
  renderEditorGrid();
}

function editorClearSelection() {
  editorPages.forEach(p => p.selected = false);
  renderEditorGrid();
}

function editorDeleteSelected() {
  const targets = editorPages.filter(p => p.selected && !p.deleted);
  if (!targets.length) { showToast('請先選取要刪除的頁面！', false); return; }
  applyEdit(() => targets.forEach(p => { p.deleted = true; p.selected = false; }));
  showToast(`已刪除 ${targets.length} 頁（Ctrl+Z 可復原）`, true);
}

function editorRotateSelected() {
  const targets = editorPages.filter(p => p.selected && !p.deleted);
  if (!targets.length) { showToast('請先選取要旋轉的頁面！', false); return; }
  applyEdit(() => targets.forEach(p => p.rotation = (p.rotation + 90) % 360));
}

// ── Keyboard shortcuts ────────────────────────────────────────────────────────

document.addEventListener('keydown', e => {
  if (document.querySelector('.modal-overlay.show')) return;
  const inField = ['INPUT', 'TEXTAREA'].includes(e.target.tagName);
  const ctrl    = e.ctrlKey || e.metaKey;
  const key     = e.key.toLowerCase();
  const loaded  = editorPages.length > 0;

  if (ctrl && key === 'o') { e.preventDefault(); openNewPdf(); return; }
  if (ctrl && key === 's') { e.preventDefault(); if (loaded) saveEditor(); return; }
  if (ctrl && key === 'r') { e.preventDefault(); return; }  // don't reload the page and lose edits
  if (inField || !loaded) return;

  if (ctrl && key === 'z' && !e.shiftKey)              { e.preventDefault(); undo(); }
  else if (ctrl && (key === 'y' || (key === 'z' && e.shiftKey))) { e.preventDefault(); redo(); }
  else if (ctrl && key === 'a') {
    e.preventDefault();
    editorPages.forEach(p => { if (!p.deleted) p.selected = true; });
    renderEditorGrid();
  }
  else if (!ctrl && (e.key === 'Delete' || e.key === 'Backspace')) { e.preventDefault(); editorDeleteSelected(); }
  else if (!ctrl && key === 'r')      editorRotateSelected();
  else if (e.key === 'Escape')        editorClearSelection();
});

// ── Save / Extract ────────────────────────────────────────────────────────────

function buildPageList(pagesArr) {
  return pagesArr.map(p => ({
    src:      p.srcFile,
    orig_idx: p.origIndex,
    rotation: p.rotation,
  }));
}

async function saveEditor() {
  const active = editorPages.filter(p => !p.deleted);
  if (!active.length) { showToast('沒有頁面可儲存！', false); return; }

  const encrypted = document.getElementById('encrypt-check').checked;
  const password  = encrypted ? document.getElementById('encrypt-pw').value : '';
  if (encrypted && !password) { showToast('請輸入加密密碼！', false); return; }

  const outInput = document.getElementById('output-path');
  let out = outInput.value.trim();
  let fromDialog = outputFromDialog;
  if (!out) {
    out = await pywebview.api.browse_save('edited_output');
    if (!out) return;
    fromDialog = true;
  }
  out = await resolveOutput(out, fromDialog);
  if (!out) return;
  outInput.value = out;
  outputFromDialog = fromDialog;

  setStatus('儲存中…', true);
  const res = await pywebview.api.save_edited_pdf(buildPageList(active), out, password)
    || { ok: false, msg: '發生未知錯誤' };
  setStatus(res.msg, res.ok);
  showToast(res.msg, res.ok);
  if (!res.ok) return;
  markClean();

  // Overwrote one of the sources: its page indices no longer match, so
  // continue editing from the freshly saved file instead
  if (editorPages.some(p => p.srcFile === res.path)) {
    await loadPdf(res.path, false);
    setStatus(res.msg, true);
  }
}

async function extractSelected() {
  const selected = editorPages.filter(p => !p.deleted && p.selected);
  if (!selected.length) { showToast('請先選取要擷取的頁面！', false); return; }

  let out = await pywebview.api.browse_save('extract_output');
  if (!out) return;
  out = await resolveOutput(out, true, false);
  if (!out) return;

  setStatus('擷取中…', true);
  const res = await pywebview.api.save_edited_pdf(buildPageList(selected), out, '')
    || { ok: false, msg: '發生未知錯誤' };
  setStatus(res.msg, res.ok);
  showToast(res.msg, res.ok);
}
