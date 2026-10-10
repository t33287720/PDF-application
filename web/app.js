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

// Escape closes the open dialog; each dialog registers how it closes
const modalClosers = {};
document.addEventListener('keydown', e => {
  if (e.key !== 'Escape') return;
  const open = [...document.querySelectorAll('.modal-overlay.show')].pop();
  if (open && modalClosers[open.id]) {
    modalClosers[open.id]();
    e.stopImmediatePropagation();  // don't also clear the page selection
  }
});

// Close on backdrop click too
function setupModal(id, close) {
  modalClosers[id] = close;
  const modal = document.getElementById(id);
  modal.addEventListener('click', e => { if (e.target === modal) close(); });
}
setupModal('confirm-modal', hideConfirm);

// ── Dropdown menus ────────────────────────────────────────────────────────────

function toggleMenu(id, e) {
  e.stopPropagation();
  const menu = document.getElementById(id);
  const willOpen = !menu.classList.contains('open');
  closeMenus();
  menu.classList.toggle('open', willOpen);
}
function closeMenus() {
  document.querySelectorAll('.dropdown.open').forEach(m => m.classList.remove('open'));
}
document.addEventListener('click', closeMenus);   // also after picking an item

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
      delete modalClosers['pw-modal'];
      resolve(value);
    };
    modalClosers['pw-modal'] = () => done(null);
    document.getElementById('pw-modal-ok').onclick     = () => done(input.value);
    document.getElementById('pw-modal-cancel').onclick = () => done(null);
    input.onkeydown = e => {
      if (e.key === 'Enter')  done(input.value);
      if (e.key === 'Escape') done(null);
    };
  });
}

// ── Editor state ──────────────────────────────────────────────────────────────

// [{srcFile, srcName, srcOrig, origIndex, size, aspect, b64, rotation, crop, selected, deleted, blank?}]
// crop: null or [left, top, right, bottom] in pt, relative to the page before rotation
let editorPages = [];
let dragSrcPage = null;
let selectAnchor = null;  // last clicked page, start of a Shift+click range

// Bookmarks: [{level, title, page}] (page = an editorPages entry). Until the user
// edits them, export keeps the source bookmarks as before (bookmarksEdited = false).
let bookmarks = [];
let bookmarksEdited = false;
let nextPageId = 1;

// ── Undo / Redo ───────────────────────────────────────────────────────────────

const UNDO_LIMIT = 100;
const undoStack = [];
const redoStack = [];

function snapshot() {
  return {
    pages: editorPages.map(p => ({ page: p, rotation: p.rotation, deleted: p.deleted, crop: p.crop })),
    bookmarks: bookmarks.map(b => ({ ...b })),
    bookmarksEdited,
  };
}

function restoreSnapshot(snap) {
  editorPages = snap.pages.map(s => {
    s.page.rotation = s.rotation;
    s.page.deleted  = s.deleted;
    s.page.crop     = s.crop;
    syncAspect(s.page);
    return s.page;
  });
  bookmarks = snap.bookmarks;
  bookmarksEdited = snap.bookmarksEdited;
  markDirty();
  renderEditorGrid();
  if (document.getElementById('bookmarks-modal').classList.contains('show')) renderBookmarks();
}

// Call before every change to page order / rotation / deletion / cropping
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

// The first file replaces the current pages (unless appending); the rest are appended
async function loadFiles(paths, append) {
  for (const [i, path] of paths.entries()) await loadPdf(path, append || i > 0);
}

async function openNewPdf() {
  const pick = async () => loadFiles(await pywebview.api.browse_open(), false);
  if (isDirty && editorPages.filter(p => !p.deleted).length > 0) {
    showConfirm('有未儲存的變更，開啟新檔案將會遺失目前的修改，確定繼續嗎？', pick);
    return;
  }
  await pick();
}

async function appendPdf() {
  await loadFiles(await pywebview.api.browse_open(), true);
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

  const newPages = res.sizes.map(([w, h], i) => ({
    srcFile:   res.path,  // for images: the PDF they were converted to
    srcName:   res.name,
    srcOrig:   path,
    origIndex: i,
    b64:       null,      // filled in by loadThumbnails()
    size:      [w, h],    // points, as displayed before any rotation here
    aspect:    w / h,
    rotation:  0,
    crop:      null,
    selected:  false,
    deleted:   false,
    id:        nextPageId++,
  }));
  const newBookmarks = (res.toc || [])
    .filter(([, , idx]) => newPages[idx])
    .map(([level, title, idx]) => ({ level, title, page: newPages[idx] }));

  if (!append) thumbGeneration++;   // stop filling thumbnails of the old document
  if (append) {
    pushUndo();
    editorPages = [...editorPages, ...newPages];
    bookmarks = [...bookmarks, ...newBookmarks];
    markDirty();
  } else {
    editorPages = newPages;
    bookmarks = newBookmarks;
    bookmarksEdited = false;
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
  loadThumbnails(res.path, newPages);
}

// ── Thumbnails (fetched in batches after the page grid is shown) ─────────────

const THUMB_BATCH = 8;
let thumbGeneration = 0;

async function loadThumbnails(path, pages) {
  const gen = thumbGeneration;
  for (let start = 0; start < pages.length; start += THUMB_BATCH) {
    const res = await pywebview.api.get_thumbnails(path, start, THUMB_BATCH);
    if (gen !== thumbGeneration) return;
    if (!res || !res.ok) { setStatus(`縮圖載入失敗：${res ? res.msg : '未知錯誤'}`, false); return; }
    res.thumbs.forEach((b64, i) => {
      const page = pages[start + i];
      page.b64 = b64;
      if (page.el) {
        if (page.crop) refreshCropThumb(page);
        else page.el.querySelector('img').src = b64;
        page.el.classList.remove('loading');
      }
    });
  }
}

// ── Files dropped from the OS (paths are delivered by the Python side) ──────

async function handleDroppedFiles(paths) {
  document.body.classList.remove('file-drag');
  // With nothing open the first file opens and the rest are appended;
  // otherwise every dropped file is appended to the current pages
  await loadFiles(paths, editorPages.length > 0);
}

// Highlight the window while files (not pages) are dragged over it
let fileDragDepth = 0;
const isFileDrag = e => e.dataTransfer && [...e.dataTransfer.types].includes('Files');
document.addEventListener('dragenter', e => {
  if (!isFileDrag(e)) return;
  fileDragDepth++;
  document.body.classList.add('file-drag');
});
document.addEventListener('dragleave', e => {
  if (!isFileDrag(e)) return;
  if (--fileDragDepth <= 0) { fileDragDepth = 0; document.body.classList.remove('file-drag'); }
});
document.addEventListener('drop', () => {
  fileDragDepth = 0;
  document.body.classList.remove('file-drag');
});

// True while the output path came from the save dialog (not typed by hand)
let outputFromDialog = false;
const outputPathInput = document.getElementById('output-path');
outputPathInput.addEventListener('input', () => {
  outputFromDialog = false;
  outputPathInput.title = outputPathInput.value;
});

// The field is narrow in small windows, so the full path also goes in the tooltip
function setOutputPath(path) {
  outputPathInput.value = path;
  outputPathInput.title = path;
}

async function browseOutput() {
  const path = await pywebview.api.browse_save('edited_output');
  if (path) {
    setOutputPath(path);
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

// ── Output options ────────────────────────────────────────────────────────────

const optionsModal = document.getElementById('options-modal');

// A checkbox reveals the settings block tied to it via data-for
optionsModal.addEventListener('change', e => {
  if (e.target.type !== 'checkbox') return;
  const body = optionsModal.querySelector(`.opt-body[data-for="${e.target.id}"]`);
  if (body) body.classList.toggle('show', e.target.checked);
});
optionsModal.addEventListener('change', e => {
  if (e.target.id !== 'opt-meta') return;
  document.getElementById('opt-meta-body').classList.toggle('show', e.target.value === 'custom');
});
setupModal('options-modal', () => closeOptions());
optionsModal.addEventListener('keydown', e => {
  if (e.key === 'Enter' && e.target.tagName === 'INPUT') closeOptions();
});

function openOptions() {
  optionsModal.classList.add('show');
}

function closeOptions() {
  if (!getOutputOptions()) return;   // keep open until the settings are valid
  optionsModal.classList.remove('show');
  updateOptionsBadge();
}

// Returns the options for the backend, or null (with a toast) if incomplete
function getOutputOptions() {
  const opts = { compress: document.getElementById('opt-compress').value };
  if (document.getElementById('opt-encrypt').checked) {
    opts.password = document.getElementById('opt-password').value;
    if (!opts.password) { showToast('請輸入加密密碼！', false); return null; }
  }
  if (document.getElementById('opt-watermark').checked) {
    const text = document.getElementById('opt-wm-text').value.trim();
    if (!text) { showToast('請輸入浮水印文字！', false); return null; }
    opts.watermark = {
      text,
      size:     document.getElementById('opt-wm-size').value,
      opacity:  Number(document.getElementById('opt-wm-opacity').value),
      diagonal: document.getElementById('opt-wm-diagonal').checked,
    };
  }
  if (document.getElementById('opt-pagenum').checked) {
    const start = parseInt(document.getElementById('opt-pn-start').value, 10);
    if (!(start >= 0)) { showToast('頁碼起始值必須是 0 以上的整數！', false); return null; }
    opts.page_numbers = {
      format:     document.getElementById('opt-pn-format').value,
      position:   document.getElementById('opt-pn-position').value,
      start,
      skip_first: document.getElementById('opt-pn-skip').checked,
    };
  }
  const metaMode = document.getElementById('opt-meta').value;
  if (metaMode !== 'keep') {
    opts.meta_mode = metaMode;
    if (metaMode === 'custom') {
      opts.title  = document.getElementById('opt-meta-title').value.trim();
      opts.author = document.getElementById('opt-meta-author').value.trim();
    }
  }
  return opts;
}

function updateOptionsBadge() {
  const opts = getOutputOptions() || {};
  const count = [opts.password, opts.compress !== 'none', opts.watermark, opts.page_numbers, opts.meta_mode].filter(Boolean).length;
  const badge = document.getElementById('options-badge');
  badge.textContent = count;
  badge.classList.toggle('show', count > 0);
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

// Each page's DOM element is built once and reused; rendering only
// re-orders elements and refreshes what changed (no image re-decoding)
function createThumb(page) {
  const div = document.createElement('div');
  div.className = page.b64 || page.blank ? 'page-thumb' : 'page-thumb loading';
  div.dataset.pageId = page.id;
  div.draggable = true;
  div.innerHTML = `
    <div class="thumb-img-wrap">
      <img draggable="false" alt="">
      <div class="rotation-badge"></div>
      <div class="thumb-overlay">
        <button class="thumb-btn" title="旋轉 90°" aria-label="旋轉 90°">↻</button>
        <button class="thumb-btn delete" title="刪除" aria-label="刪除此頁">✕</button>
      </div>
    </div>
    <div class="thumb-footer">
      <span class="thumb-num"></span>
      <span class="thumb-src"></span>
    </div>`;
  const [btnRotate, btnDelete] = div.querySelectorAll('.thumb-btn');
  btnRotate.addEventListener('click', e => rotateOnePage(editorPages.indexOf(page), e));
  btnDelete.addEventListener('click', e => deleteOnePage(editorPages.indexOf(page), e));

  // Select on click (Shift+click selects a range)
  div.addEventListener('click', e => {
    const visible = editorPages.filter(p => !p.deleted);
    if (e.shiftKey && selectAnchor && visible.includes(selectAnchor)) {
      const [a, b] = [visible.indexOf(selectAnchor), visible.indexOf(page)].sort((x, y) => x - y);
      visible.slice(a, b + 1).forEach(p => { p.selected = true; });
      renderEditorGrid();
      return;
    }
    page.selected = !page.selected;
    selectAnchor = page;
    div.classList.toggle('selected', page.selected);
    updateInfo();
  });

  div.addEventListener('dblclick', () => openPreview(page));

  // Drag-and-drop
  div.addEventListener('dragstart', e => {
    dragSrcPage = page;
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', 'page');
    setTimeout(() => {
      div.classList.add('dragging');
      document.getElementById('editor-grid').classList.add('is-dragging');
    }, 0);
  });

  div.addEventListener('dragend', () => {
    dragSrcPage = null;
    div.classList.remove('dragging');
    document.getElementById('editor-grid').classList.remove('is-dragging');
    clearDragOver();
  });

  div.addEventListener('dragover', e => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    if (!dragSrcPage || dragSrcPage === page) return;
    clearDragOver();
    div.classList.add('drag-over');
  });

  div.addEventListener('drop', e => {
    e.preventDefault();
    clearDragOver();
    if (!dragSrcPage || dragSrcPage === page) return;
    const rect = div.getBoundingClientRect();
    const insertBefore = e.clientX < rect.left + rect.width / 2;
    const moved = dragSrcPage;
    dragSrcPage = null;
    applyEdit(() => {
      editorPages.splice(editorPages.indexOf(moved), 1);
      const target = editorPages.indexOf(page);
      editorPages.splice(insertBefore ? target : target + 1, 0, moved);
    });
  });

  page.el = div;
  return div;
}

function clearDragOver() {
  document.querySelectorAll('.page-thumb.drag-over').forEach(el => el.classList.remove('drag-over'));
}

function updateThumb(page, displayIdx, srcName) {
  const div = page.el || createThumb(page);
  div.classList.toggle('selected', page.selected);
  div.querySelector('.thumb-img-wrap').style.cssText = thumbWrapStyle(page);
  const img = div.querySelector('img');
  img.style.cssText = thumbImgStyle(page);
  if (page.crop) refreshCropThumb(page);
  const shown = page.crop ? page.cropSrc : page.b64;
  if (shown && img.getAttribute('src') !== shown) img.src = shown;
  const badge = div.querySelector('.rotation-badge');
  badge.textContent = page.rotation ? `↻ ${page.rotation}°` : '';
  badge.style.display = page.rotation ? '' : 'none';
  img.alt = `第 ${displayIdx + 1} 頁`;
  div.querySelector('.thumb-num').textContent = displayIdx + 1;
  const src = div.querySelector('.thumb-src');
  src.textContent = srcName;
  src.title = srcName && !page.blank ? page.srcOrig : '';
  src.style.display = srcName ? '' : 'none';
  return div;
}

// End drop zone — lets pages be dragged to the very last position
const endZone = document.createElement('div');
endZone.className = 'end-drop-zone';
endZone.textContent = '放置於此';
endZone.addEventListener('dragover', e => {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
  endZone.classList.add('drag-active');
});
endZone.addEventListener('dragleave', () => endZone.classList.remove('drag-active'));
endZone.addEventListener('drop', e => {
  e.preventDefault();
  endZone.classList.remove('drag-active');
  if (!dragSrcPage) return;
  const moved = dragSrcPage;
  dragSrcPage = null;
  applyEdit(() => editorPages.push(editorPages.splice(editorPages.indexOf(moved), 1)[0]));
});

function renderEditorGrid() {
  const grid = document.getElementById('editor-grid');
  const visible = editorPages.filter(p => !p.deleted);
  if (!visible.length) {
    grid.innerHTML = '<div class="editor-placeholder"><p>所有頁面已刪除</p></div>';
    updateInfo();
    return;
  }

  // Show the source file name only when pages come from more than one PDF
  const multiSrc = new Set(visible.filter(p => !p.blank).map(p => p.srcFile)).size > 1;
  const label = page => page.blank ? '空白頁' : multiSrc ? page.srcName : '';
  grid.replaceChildren(
    ...visible.map((page, i) => updateThumb(page, i, label(page))),
    endZone,
  );
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

// Inserts after the last selected page (or at the end), sized like its neighbour
function insertBlankPage() {
  const visible  = editorPages.filter(p => !p.deleted);
  const selected = visible.filter(p => p.selected);
  const ref = selected.length ? selected[selected.length - 1] : visible[visible.length - 1];
  let [w, h] = ref ? croppedSize(ref) : [595, 842];   // A4
  if (ref && ref.rotation % 180) [w, h] = [h, w];
  const page = {
    blank: true, srcFile: null, origIndex: null, size: [w, h], aspect: w / h,
    b64: null, rotation: 0, crop: null, selected: false, deleted: false, id: nextPageId++,
  };
  applyEdit(() => {
    editorPages.splice(ref ? editorPages.indexOf(ref) + 1 : editorPages.length, 0, page);
  });
}

function editorRotateSelected() {
  const targets = editorPages.filter(p => p.selected && !p.deleted);
  if (!targets.length) { showToast('請先選取要旋轉的頁面！', false); return; }
  applyEdit(() => targets.forEach(p => p.rotation = (p.rotation + 90) % 360));
}

// ── Crop ──────────────────────────────────────────────────────────────────────

const MM_TO_PT = 72 / 25.4;
const MIN_CROPPED_PT = 10;   // never leave less than this of a page

setupModal('crop-modal', closeCrop);

// Size in pt of what remains of the page after cropping (before rotation)
function croppedSize(page) {
  const [l, t, r, b] = page.crop || [0, 0, 0, 0];
  return [page.size[0] - l - r, page.size[1] - t - b];
}

function syncAspect(page) {
  const [w, h] = croppedSize(page);
  page.aspect = w / h;
}

// Margins [l, t, r, b] as seen after turning the page `deg` clockwise, and back
function toDisplayed(m, deg) {
  for (let i = 0; i < deg / 90; i++) m = [m[3], m[0], m[1], m[2]];
  return m;
}
function fromDisplayed(m, deg) {
  for (let i = 0; i < deg / 90; i++) m = [m[1], m[2], m[3], m[0]];
  return m;
}

// Resolves to a data URL of `src` (a render of the whole page) trimmed by `crop`, or null
function cropImage(src, size, crop) {
  return new Promise(resolve => {
    const img = new Image();
    img.onload = () => {
      const k = img.naturalWidth / size[0];
      const [l, t, r, b] = crop;
      const w = Math.max(1, Math.round((size[0] - l - r) * k));
      const h = Math.max(1, Math.round((size[1] - t - b) * k));
      const canvas = document.createElement('canvas');
      canvas.width = w;
      canvas.height = h;
      canvas.getContext('2d').drawImage(img, l * k, t * k, w, h, 0, 0, w, h);
      resolve(canvas.toDataURL('image/jpeg', 0.85));
    };
    img.onerror = () => resolve(null);
    img.src = src;
  });
}

// Builds the cropped thumbnail in the background, then redraws
async function refreshCropThumb(page) {
  if (!page.crop || !page.b64) return;
  const key = page.crop.join(',');
  if (page.cropKey === key) return;
  page.cropKey = key;
  const src = await cropImage(page.b64, page.size, page.crop);
  if (!src || page.cropKey !== key) return;
  page.cropSrc = src;
  renderEditorGrid();
}

const cropInputs = { top: 'crop-top', bottom: 'crop-bottom', left: 'crop-left', right: 'crop-right' };

function cropTargets() {
  const all = document.querySelector('input[name="crop-target"]:checked').value === 'all';
  return editorPages.filter(p => !p.deleted && !p.blank && (all || p.selected));
}

function updateCropSummary() {
  const n = cropTargets().length;
  document.getElementById('crop-summary').textContent =
    n ? `將套用到 ${n} 頁（空白頁不會裁切）` : '沒有可裁切的頁面，請先選取頁面';
}

function openCrop() {
  const hasSelection = editorPages.some(p => p.selected && !p.deleted && !p.blank);
  document.querySelector(`input[name="crop-target"][value="${hasSelection ? 'selected' : 'all'}"]`).checked = true;
  // Start from the first target's current margins, as seen on screen
  const first = cropTargets()[0];
  const [l, t, r, b] = first && first.crop ? toDisplayed(first.crop, first.rotation) : [0, 0, 0, 0];
  const values = { left: l, top: t, right: r, bottom: b };
  for (const [side, id] of Object.entries(cropInputs)) {
    document.getElementById(id).value = +(values[side] / MM_TO_PT).toFixed(1);
  }
  updateCropSummary();
  document.getElementById('crop-modal').classList.add('show');
  document.getElementById(cropInputs.top).focus();
}

function closeCrop() {
  document.getElementById('crop-modal').classList.remove('show');
}

document.querySelectorAll('input[name="crop-target"]').forEach(el => el.addEventListener('change', updateCropSummary));

// Sets (replaces) the crop of the target pages; all zeros removes it
function applyCrop() {
  const mm = {};
  for (const [side, id] of Object.entries(cropInputs)) {
    const v = parseFloat(document.getElementById(id).value);
    if (!(v >= 0)) { showToast('裁切寬度必須是 0 以上的數字', false); return; }
    mm[side] = v * MM_TO_PT;
  }
  const targets = cropTargets();
  if (!targets.length) { showToast('請先選取要裁切的頁面！', false); return; }

  const displayed = [mm.left, mm.top, mm.right, mm.bottom];
  const clearing = displayed.every(v => v === 0);
  const plan = targets.map(p => {
    const crop = clearing ? null : fromDisplayed(displayed, p.rotation);
    const ok = !crop || (p.size[0] - crop[0] - crop[2] >= MIN_CROPPED_PT &&
                         p.size[1] - crop[1] - crop[3] >= MIN_CROPPED_PT);
    return { page: p, crop, ok };
  });
  const skipped = plan.filter(x => !x.ok).length;
  if (skipped === plan.length) { showToast('裁切範圍超過頁面大小，請縮小數值', false); return; }

  closeCrop();
  applyEdit(() => plan.forEach(({ page, crop, ok }) => {
    if (!ok) return;
    page.crop = crop;
    page.cropKey = page.cropSrc = null;
    syncAspect(page);
  }));
  showToast(clearing ? '已清除裁切'
    : `已裁切 ${plan.length - skipped} 頁` + (skipped ? `（${skipped} 頁因超出頁面大小而略過）` : '') + '（Ctrl+Z 可復原）', true);
}

// ── Page preview ──────────────────────────────────────────────────────────────

let previewPage = null;
let previewToken = 0;

async function openPreview(page) {
  previewPage = page;
  const token = ++previewToken;
  const visible = editorPages.filter(p => !p.deleted);
  const pos = visible.indexOf(page);
  const fig = document.querySelector('.preview-figure');
  const img = document.getElementById('preview-img');
  document.getElementById('preview-modal').classList.add('show');
  document.querySelector('.preview-nav.prev').disabled = pos <= 0;
  document.querySelector('.preview-nav.next').disabled = pos >= visible.length - 1;
  document.getElementById('preview-caption').textContent = page.blank
    ? `第 ${pos + 1} / ${visible.length} 頁　空白頁`
    : `第 ${pos + 1} / ${visible.length} 頁　${page.srcName} 原第 ${page.origIndex + 1} 頁`;
  img.alt = `第 ${pos + 1} 頁預覽`;
  fig.classList.remove('ready');

  const res = page.blank
    ? { ok: true, b64: blankImage(page) }
    : await pywebview.api.get_preview(page.srcFile, page.origIndex);
  if (token !== previewToken) return;  // user moved on to another page
  if (!res || !res.ok) { closePreview(); showToast(res ? res.msg : '預覽失敗', false); return; }
  img.src = page.crop ? await cropImage(res.b64, page.size, page.crop) || res.b64 : res.b64;
  if (token !== previewToken) return;
  img.style.transform = page.rotation ? `rotate(${page.rotation}deg)` : '';
  img.classList.toggle('quarter', page.rotation % 180 !== 0);
  fig.classList.add('ready');
}

function blankImage(page) {
  const canvas = document.createElement('canvas');
  canvas.width  = 600;
  canvas.height = Math.round(600 / page.aspect);
  const g = canvas.getContext('2d');
  g.fillStyle = '#fff';
  g.fillRect(0, 0, canvas.width, canvas.height);
  return canvas.toDataURL();
}

function stepPreview(delta) {
  const visible = editorPages.filter(p => !p.deleted);
  const next = visible[visible.indexOf(previewPage) + delta];
  if (next) openPreview(next);
}

function closePreview() {
  previewToken++;
  previewPage = null;
  document.getElementById('preview-modal').classList.remove('show');
}

document.getElementById('preview-modal').addEventListener('click', e => {
  if (e.target.id === 'preview-modal' || e.target.classList.contains('preview-figure')) closePreview();
});

// ── Keyboard shortcuts ────────────────────────────────────────────────────────

document.addEventListener('keydown', e => {
  if (document.querySelector('.modal-overlay.show')) return;
  if (previewPage) {
    if (e.key === 'Escape' || e.key === ' ') { e.preventDefault(); closePreview(); }
    else if (e.key === 'ArrowLeft')  stepPreview(-1);
    else if (e.key === 'ArrowRight') stepPreview(1);
    return;
  }
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
  return pagesArr.map(p => p.blank
    ? { blank: true, width: p.size[0], height: p.size[1], rotation: p.rotation, bid: p.id }
    : { src: p.srcFile, orig_idx: p.origIndex, rotation: p.rotation, crop: p.crop, bid: p.id });
}

// Output options plus the edited bookmark list (omitted until edited)
function withBookmarks(options) {
  if (!bookmarksEdited) return options;
  return {
    ...options,
    bookmarks: bookmarks
      .filter(b => !b.page.deleted)
      .map(b => ({ level: b.level, title: b.title, bid: b.page.id })),
  };
}

async function saveEditor() {
  const active = editorPages.filter(p => !p.deleted);
  if (!active.length) { showToast('沒有頁面可儲存！', false); return; }

  const options = getOutputOptions();
  if (!options) { openOptions(); return; }

  let out = outputPathInput.value.trim();
  let fromDialog = outputFromDialog;
  if (!out) {
    out = await pywebview.api.browse_save('edited_output');
    if (!out) return;
    fromDialog = true;
  }
  out = await resolveOutput(out, fromDialog);
  if (!out) return;
  setOutputPath(out);
  outputFromDialog = fromDialog;

  setStatus('儲存中…', true);
  const res = await pywebview.api.save_edited_pdf(buildPageList(active), out, withBookmarks(options))
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

// ── Split ─────────────────────────────────────────────────────────────────────

setupModal('split-modal', closeSplit);

function openSplit() {
  document.getElementById('split-modal').classList.add('show');
  updateSplitSummary();
  document.getElementById('split-ranges').focus();
}

function closeSplit() {
  document.getElementById('split-modal').classList.remove('show');
}

// "1-3, 5, 9-7" -> [[0,1,2], [4], [8,7,6]] (0-based), or an error message
function parseRanges(text, total) {
  const parts = text.split(/[,，;；]/).map(s => s.trim()).filter(Boolean);
  if (!parts.length) return { error: '請輸入頁碼範圍' };
  const groups = [];
  for (const part of parts) {
    const m = part.match(/^(\d+)(?:\s*[-–~]\s*(\d+))?$/);
    if (!m) return { error: `無法辨識「${part}」` };
    const a = Number(m[1]), b = m[2] ? Number(m[2]) : a;
    if (a < 1 || b < 1 || a > total || b > total) return { error: `「${part}」超出範圍（1 – ${total}）` };
    const step = a <= b ? 1 : -1;
    const group = [];
    for (let n = a; n !== b + step; n += step) group.push(n - 1);
    groups.push(group);
  }
  return { groups };
}

// Page index groups for the chosen mode, or { error }
function splitGroups() {
  const total = editorPages.filter(p => !p.deleted).length;
  const mode  = document.querySelector('input[name="split-mode"]:checked').value;
  if (mode === 'ranges') return parseRanges(document.getElementById('split-ranges').value, total);
  const size = mode === 'single' ? 1 : parseInt(document.getElementById('split-every').value, 10);
  if (!(size >= 1)) return { error: '頁數必須是 1 以上的整數' };
  const groups = [];
  for (let i = 0; i < total; i += size) {
    groups.push(Array.from({ length: Math.min(size, total - i) }, (_, k) => i + k));
  }
  return { groups };
}

function updateSplitSummary() {
  const res = splitGroups();
  const summary = document.getElementById('split-summary');
  summary.textContent = res.error || `將產生 ${res.groups.length} 個檔案`;
  summary.classList.toggle('err', !!res.error);
  document.getElementById('split-go').disabled = !!res.error;
}

document.getElementById('split-modal').addEventListener('input', updateSplitSummary);
document.getElementById('split-modal').addEventListener('keydown', e => {
  if (e.key === 'Enter' && e.target.tagName === 'INPUT') runSplit();
});
document.getElementById('split-modal').addEventListener('change', updateSplitSummary);
// Typing in a mode's field selects that mode
document.getElementById('split-ranges').addEventListener('focus', () => {
  document.querySelector('input[name="split-mode"][value="ranges"]').checked = true;
});
document.getElementById('split-every').addEventListener('focus', () => {
  document.querySelector('input[name="split-mode"][value="every"]').checked = true;
});

async function runSplit() {
  const res = splitGroups();
  if (res.error) return;
  const options = getOutputOptions();
  if (!options) { closeSplit(); openOptions(); return; }
  const folder = await pywebview.api.browse_folder();
  if (!folder) return;
  closeSplit();

  const visible = editorPages.filter(p => !p.deleted);
  const prefix  = exportPrefix(visible);
  const groups  = res.groups.map(idx => {
    const first = idx[0] + 1, last = idx[idx.length - 1] + 1;
    return {
      name:  `${prefix}_${first === last ? first : `${first}-${last}`}`,
      pages: buildPageList(idx.map(i => visible[i])),
    };
  });
  setStatus(`分割成 ${groups.length} 個檔案中…`, true);
  const out = await pywebview.api.split_pdf(groups, folder, withBookmarks(options))
    || { ok: false, msg: '發生未知錯誤' };
  setStatus(out.msg, out.ok);
  showToast(out.msg, out.ok);
}

// ── Export as images ──────────────────────────────────────────────────────────

setupModal('images-modal', closeImageExport);

document.getElementById('img-format').addEventListener('change', e => {
  document.getElementById('img-quality-row').style.display = e.target.value === 'jpg' ? '' : 'none';
});
document.getElementById('img-quality').addEventListener('input', e => {
  document.getElementById('img-quality-label').textContent = e.target.value;
});

function openImageExport() {
  const hasSel = editorPages.some(p => p.selected && !p.deleted);
  const scope = document.getElementById('img-scope');
  scope.querySelector('[value="selected"]').disabled = !hasSel;
  scope.value = hasSel ? 'selected' : 'all';
  document.getElementById('images-modal').classList.add('show');
}

function closeImageExport() {
  document.getElementById('images-modal').classList.remove('show');
}

// File name prefix for exports: the first real page's source name, sans extension
function exportPrefix(pages) {
  const first = pages.find(p => !p.blank);
  return first ? first.srcName.replace(/\.[^.]+$/, '') : 'page';
}

async function exportImages() {
  const all   = document.getElementById('img-scope').value === 'all';
  const pages = editorPages.filter(p => !p.deleted && (all || p.selected));
  if (!pages.length) { showToast('沒有頁面可匯出！', false); return; }
  const options = getOutputOptions();
  if (!options) { closeImageExport(); openOptions(); return; }

  const folder = await pywebview.api.browse_folder();
  if (!folder) return;
  closeImageExport();

  setStatus(`匯出 ${pages.length} 張圖片中…`, true);
  const res = await pywebview.api.export_images(
    buildPageList(pages), folder,
    document.getElementById('img-format').value,
    Number(document.getElementById('img-dpi').value),
    options, exportPrefix(pages),
    Number(document.getElementById('img-quality').value),
  ) || { ok: false, msg: '發生未知錯誤' };
  setStatus(res.msg, res.ok);
  showToast(res.msg, res.ok);
}

async function exportText() {
  const live = editorPages.filter(p => !p.deleted);
  const selected = live.filter(p => p.selected);
  const pages = selected.length ? selected : live;
  if (!pages.length) { showToast('沒有頁面可匯出！', false); return; }

  const out = await pywebview.api.browse_save(exportPrefix(pages), 'txt');
  if (!out) return;

  setStatus(`匯出 ${pages.length} 頁文字中…`, true);
  const res = await pywebview.api.export_text(buildPageList(pages), out)
    || { ok: false, msg: '發生未知錯誤' };
  setStatus(res.msg, res.ok);
  showToast(res.msg, res.ok);
}

async function extractSelected() {
  const selected = editorPages.filter(p => !p.deleted && p.selected);
  if (!selected.length) { showToast('請先選取要擷取的頁面！', false); return; }
  const options = getOutputOptions();
  if (!options) { openOptions(); return; }

  let out = await pywebview.api.browse_save('extract_output');
  if (!out) return;
  out = await resolveOutput(out, true, false);
  if (!out) return;

  setStatus('擷取中…', true);
  const res = await pywebview.api.save_edited_pdf(buildPageList(selected), out, withBookmarks(options))
    || { ok: false, msg: '發生未知錯誤' };
  setStatus(res.msg, res.ok);
  showToast(res.msg, res.ok);
}

// ── Bookmarks ─────────────────────────────────────────────────────────────────

setupModal('bookmarks-modal', closeBookmarks);

// Bookmarks of pages still present, in page order (stable within a page)
function visibleBookmarks() {
  const order = new Map(editorPages.filter(p => !p.deleted).map((p, i) => [p, i]));
  return bookmarks
    .filter(b => order.has(b.page))
    .map((b, i) => ({ b, i }))
    .sort((x, y) => order.get(x.b.page) - order.get(y.b.page) || x.i - y.i)
    .map(x => x.b);
}

function pageNumberOf(page) {
  return editorPages.filter(p => !p.deleted).indexOf(page) + 1;
}

function openBookmarks() {
  const sel = editorPages.filter(p => p.selected && !p.deleted);
  document.getElementById('bm-page').value = sel.length ? pageNumberOf(sel[0]) : 1;
  document.getElementById('bm-title').value = '';
  document.getElementById('bookmarks-modal').classList.add('show');
  renderBookmarks();
  document.getElementById('bm-title').focus();
}

function closeBookmarks() {
  document.getElementById('bookmarks-modal').classList.remove('show');
}

function renderBookmarks() {
  const list = document.getElementById('bm-list');
  const items = visibleBookmarks();
  if (!items.length) {
    list.innerHTML = '<p class="bm-empty">目前沒有書籤</p>';
    return;
  }
  list.replaceChildren(...items.map(b => {
    const row = document.createElement('div');
    row.className = 'bm-row';
    row.style.paddingLeft = `${(b.level - 1) * 18}px`;
    const mk = (text, label, fn, disabled) => {
      const btn = document.createElement('button');
      btn.className = 'btn-tool bm-btn';
      btn.textContent = text;
      btn.title = btn.ariaLabel = label;
      btn.disabled = !!disabled;
      btn.onclick = fn;
      return btn;
    };
    const name = document.createElement('button');
    name.className = 'bm-name';
    name.textContent = b.title;
    name.title = `跳到第 ${pageNumberOf(b.page)} 頁`;
    name.onclick = () => jumpToPage(b.page);
    const pg = document.createElement('span');
    pg.className = 'bm-page';
    pg.textContent = `第 ${pageNumberOf(b.page)} 頁`;
    row.append(
      name, pg,
      mk('◀', '提高層級', () => editBookmarks(() => b.level--), b.level <= 1),
      mk('▶', '降低層級', () => editBookmarks(() => b.level++), b.level >= 6),
      mk('✎', '改名', () => {
        const input = document.createElement('input');
        input.className = 'opt-input bm-rename';
        input.value = b.title;
        input.setAttribute('aria-label', '書籤名稱');
        let done = false;
        const finish = save => {
          if (done) return;
          done = true;
          const t = input.value.trim();
          if (save && t && t !== b.title) editBookmarks(() => { b.title = t; });
          else renderBookmarks();
        };
        input.onkeydown = e => {
          if (e.key === 'Enter')  finish(true);
          if (e.key === 'Escape') { e.stopPropagation(); finish(false); }
        };
        input.onblur = () => finish(true);
        name.replaceWith(input);
        input.focus();
        input.select();
      }),
      mk('✕', '刪除書籤', () => editBookmarks(() => { bookmarks = bookmarks.filter(x => x !== b); })),
    );
    return row;
  }));
}

function editBookmarks(change) {
  pushUndo();
  change();
  bookmarksEdited = true;
  markDirty();
  renderBookmarks();
  updateUndoButtons();
}

function addBookmark() {
  const titleEl = document.getElementById('bm-title');
  const title = titleEl.value.trim();
  const active = editorPages.filter(p => !p.deleted);
  const n = Math.round(Number(document.getElementById('bm-page').value));
  if (!title) { showToast('請輸入書籤名稱！', false); titleEl.focus(); return; }
  if (!(n >= 1 && n <= active.length)) { showToast(`頁碼須介於 1 – ${active.length}`, false); return; }
  editBookmarks(() => bookmarks.push({ level: 1, title, page: active[n - 1] }));
  titleEl.value = '';
  titleEl.focus();
}

document.getElementById('bm-title').addEventListener('keydown', e => {
  if (e.key === 'Enter') addBookmark();
});

function jumpToPage(page) {
  closeBookmarks();
  const el = document.querySelector(`#editor-grid [data-page-id="${page.id}"]`);
  if (el) el.scrollIntoView({ block: 'center', behavior: 'smooth' });
}
