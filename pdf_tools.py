#!/usr/bin/env python
# coding: utf-8

import os
import sys
import base64
import json
import secrets
import threading
import webview
from webview.dom import DOMEventHandler
import fitz  # PyMuPDF


THUMB_PX    = 240   # thumbnail render width (2x the 120px display size)
PREVIEW_PX  = 1600  # preview render size of the page's longer side

# Image downsampling for "compress": (images above this DPI, target DPI, JPEG quality)
COMPRESS_LEVELS = {
    "medium": (180, 150, 75),
    "high":   (110, 96, 60),
}

# Encrypted output: anything but copying text out is allowed
ENCRYPTED_PERMS = int(
    fitz.PDF_PERM_PRINT | fitz.PDF_PERM_PRINT_HQ | fitz.PDF_PERM_MODIFY |
    fitz.PDF_PERM_ANNOTATE | fitz.PDF_PERM_FORM | fitz.PDF_PERM_ASSEMBLE |
    fitz.PDF_PERM_ACCESSIBILITY
)

_is_dirty   = False


def _settings_path():
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "PDF_Tools", "settings.json")


def _load_settings():
    try:
        with open(_settings_path(), encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _save_settings(settings):
    try:
        path = _settings_path()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except OSError:
        pass   # settings are a convenience; never fail an action over them


def _auto_name(directory, prefix):
    directory = directory or "."
    i = 1
    while True:
        path = os.path.join(directory, f"{prefix}_{i}.pdf")
        if not os.path.exists(path):
            return path
        i += 1


def _page_runs(pages):
    """Group pages into (src, first, last) runs of ascending consecutive indices."""
    runs = []
    for p in pages:
        src, idx = p["src"], p["orig_idx"]
        if runs and runs[-1][0] == src and runs[-1][2] == idx - 1:
            runs[-1][2] = idx
        else:
            runs.append([src, idx, idx])
    return runs


def _remap_toc(pages, open_docs):
    """Carry source bookmarks over to the pages' new positions; drop removed ones."""
    new_pos = {}
    for n, p in enumerate(pages, start=1):
        new_pos.setdefault((p["src"], p["orig_idx"]), n)
    toc, seen = [], set()
    for p in pages:
        src = p["src"]
        if src in seen:
            continue
        seen.add(src)
        for level, title, page_no, *_ in open_docs[src].get_toc(simple=False):
            n = new_pos.get((src, page_no - 1))
            if n is None:
                continue
            # Parents may have been dropped: keep levels contiguous from 1
            level = min(level, toc[-1][0] + 1 if toc else 1)
            toc.append([level, title, n])
    return toc


class API:

    def __init__(self):
        self._passwords = {}   # src path -> password for encrypted source PDFs
        # pywebview runs each JS call on its own thread and MuPDF isn't
        # thread-safe, so every fitz operation goes through this lock
        self._lock = threading.RLock()
        self._thumb_doc = None  # (path, doc) kept open between thumbnail batches
        self._settings = _load_settings()

    def _remember_dir(self, path):
        folder = os.path.dirname(path)
        if folder and self._settings.get("last_dir") != folder:
            self._settings["last_dir"] = folder
            _save_settings(self._settings)

    def _last_dir(self):
        folder = self._settings.get("last_dir", "")
        return folder if os.path.isdir(folder) else ""

    def _close_thumb_doc(self):
        if self._thumb_doc:
            self._thumb_doc[1].close()
            self._thumb_doc = None

    def _open_src(self, path):
        doc = fitz.open(path)
        if doc.needs_pass and not doc.authenticate(self._passwords.get(path, "")):
            doc.close()
            raise ValueError(f"無法解鎖加密檔案：{os.path.basename(path)}")
        return doc

    # ── Zoom ──────────────────────────────────────────────────────────────────

    def set_dirty(self, value):
        global _is_dirty
        _is_dirty = bool(value)

    def get_zoom(self):
        return self._settings.get("zoom", 100)

    def apply_zoom(self, level):
        if self._settings.get("zoom") != level:
            self._settings["zoom"] = level
            _save_settings(self._settings)
        # GTK (Linux)
        try:
            import webview.platforms.gtk as gtk_platform
            for win in gtk_platform.BrowserView.instances.values():
                win.webview.set_zoom_level(level / 100.0)
            return
        except Exception:
            pass
        # Windows Edge / fallback: CSS transform on root-wrapper
        # Scale down the wrapper then scale back up so it fills the viewport exactly
        try:
            scale = level / 100.0
            comp  = 100.0 / scale
            if scale == 1.0:
                js = ("var w=document.getElementById('root-wrapper');"
                      "w.style.transform='';"
                      "w.style.width='100vw';"
                      "w.style.height='100vh';")
            else:
                js = (f"var w=document.getElementById('root-wrapper');"
                      f"w.style.transformOrigin='0 0';"
                      f"w.style.transform='scale({scale})';"
                      f"w.style.width='{comp:.4f}vw';"
                      f"w.style.height='{comp:.4f}vh';")
            webview.windows[0].evaluate_js(js)
        except Exception:
            pass

    # ── File dialogs ──────────────────────────────────────────────────────────

    def browse_open(self):
        result = webview.windows[0].create_file_dialog(
            webview.FileDialog.OPEN,
            directory=self._last_dir(),
            file_types=("PDF Files (*.pdf)",)
        )
        if not result:
            return ""
        self._remember_dir(result[0])
        return result[0]

    def browse_save(self, default_name="output"):
        result = webview.windows[0].create_file_dialog(
            webview.FileDialog.SAVE,
            directory=self._last_dir(),
            save_filename=f"{default_name}.pdf",
            file_types=("PDF Files (*.pdf)",)
        )
        if not result:
            return ""
        path = result[0] if isinstance(result, (list, tuple)) else result
        self._remember_dir(path)
        return path

    def check_output(self, path):
        """Normalise an output path and report whether it already exists."""
        path = os.path.abspath(os.path.expanduser(path.strip()))
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        if not os.path.isdir(os.path.dirname(path)):
            return {"ok": False, "msg": f"資料夾不存在：{os.path.dirname(path)}"}
        return {
            "ok": True,
            "path": path,
            "exists": os.path.exists(path),
            # WinForms' SaveFileDialog already asks before overwriting; GTK's doesn't
            "dialog_confirms_overwrite": sys.platform == "win32",
        }

    # ── Editor ────────────────────────────────────────────────────────────────

    def open_pdf_for_editor(self, path, password=""):
        """Check the file (and password) and return page sizes; thumbnails
        are fetched afterwards in batches through get_thumbnails()."""
        try:
            if not os.path.isfile(path):
                return {"ok": False, "msg": "PDF 不存在！"}
            with self._lock:
                doc = fitz.open(path)
                try:
                    if doc.needs_pass:
                        password = password or self._passwords.get(path, "")
                        if not password:
                            return {"ok": False, "need_password": True, "msg": "此 PDF 需要密碼"}
                        if not doc.authenticate(password):
                            return {"ok": False, "need_password": True, "msg": "密碼錯誤，請重新輸入"}
                        self._passwords[path] = password
                    if len(doc) == 0:
                        return {"ok": False, "msg": "此 PDF 沒有任何頁面"}
                    sizes = [[pg.rect.width, pg.rect.height] for pg in doc]
                finally:
                    doc.close()
            return {"ok": True, "sizes": sizes}
        except Exception as e:
            return {"ok": False, "msg": str(e)}

    def get_thumbnails(self, path, start, count):
        try:
            with self._lock:
                if not self._thumb_doc or self._thumb_doc[0] != path:
                    self._close_thumb_doc()
                    self._thumb_doc = (path, self._open_src(path))
                doc = self._thumb_doc[1]
                thumbs = []
                for i in range(start, min(start + count, len(doc))):
                    page = doc[i]
                    zoom = THUMB_PX / max(page.rect.width, 1)
                    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                    b64 = base64.b64encode(pix.tobytes("jpeg", jpg_quality=75)).decode()
                    thumbs.append(f"data:image/jpeg;base64,{b64}")
            return {"ok": True, "thumbs": thumbs}
        except Exception as e:
            return {"ok": False, "msg": str(e)}

    def _build_and_save(self, pages, out_path, options):
        # Build new document (pages may come from different source files)
        open_docs = {}
        new_doc = fitz.open()
        try:
            try:
                for p in pages:
                    if p["src"] not in open_docs:
                        open_docs[p["src"]] = self._open_src(p["src"])
                # Insert consecutive pages of the same source in one call,
                # which also keeps internal links between them
                for src, start, end in _page_runs(pages):
                    new_doc.insert_pdf(open_docs[src], from_page=start, to_page=end)
                toc = _remap_toc(pages, open_docs)   # needs the sources still open
            finally:
                for doc in open_docs.values():
                    doc.close()
            for pg, p in zip(new_doc, pages):
                if p["rotation"]:
                    pg.set_rotation((pg.rotation + p["rotation"]) % 360)
            if toc:
                new_doc.set_toc(toc)

            save_opts = dict(garbage=4, deflate=True)
            level = COMPRESS_LEVELS.get(options.get("compress"))
            if level:
                threshold, target, quality = level
                new_doc.rewrite_images(dpi_threshold=threshold, dpi_target=target, quality=quality)
                save_opts.update(use_objstms=True)
            password = options.get("password")
            if password:
                # A separate random owner password keeps the permission
                # restrictions enforceable (with owner == user they'd be moot)
                save_opts.update(encryption=fitz.PDF_ENCRYPT_AES_256,
                               user_pw=password,
                               owner_pw=secrets.token_urlsafe(24),
                               permissions=ENCRYPTED_PERMS)
            new_doc.save(out_path, **save_opts)
        finally:
            new_doc.close()

    def get_preview(self, path, index):
        try:
            with self._lock:
                doc = self._open_src(path)
                try:
                    page = doc[index]
                    zoom = PREVIEW_PX / max(page.rect.width, page.rect.height, 1)
                    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                    b64 = base64.b64encode(pix.tobytes("jpeg", jpg_quality=85)).decode()
                finally:
                    doc.close()
            return {"ok": True, "b64": f"data:image/jpeg;base64,{b64}"}
        except Exception as e:
            return {"ok": False, "msg": str(e)}

    def save_edited_pdf(self, pages, out, options=None):
        """
        pages   : list of {src, orig_idx, rotation}
        out     : output path (empty = auto)
        options : {password, compress: "none" | "medium" | "high"}
        """
        options = options or {}
        password = options.get("password", "")
        try:
            if not pages:
                return {"ok": False, "msg": "沒有頁面可儲存！"}

            first_dir = os.path.dirname(pages[0]["src"]) or "."
            out_path = out or _auto_name(first_dir, "edited")

            with self._lock:
                self._close_thumb_doc()   # output may overwrite that file
                self._build_and_save(pages, out_path, options)

            if password:
                self._passwords[out_path] = password
            else:
                self._passwords.pop(out_path, None)
            return {"ok": True, "path": out_path, "msg": f"儲存成功！已儲存：{out_path}"}
        except Exception as e:
            return {"ok": False, "msg": str(e)}


def _bind_file_drop(window):
    """Files dropped from the OS: only the Python side gets their full paths."""
    def on_drop(e):
        files = e.get("dataTransfer", {}).get("files", [])
        paths = [f["pywebviewFullPath"] for f in files
                 if f.get("pywebviewFullPath", "").lower().endswith(".pdf")]
        if paths:
            window.evaluate_js(f"handleDroppedFiles({json.dumps(paths)})")
        elif files:
            window.evaluate_js("showToast('只能拖入 PDF 檔案！', false)")

    doc_events = window.dom.document.events
    doc_events.dragenter += DOMEventHandler(lambda e: None, True, False)
    doc_events.dragover  += DOMEventHandler(lambda e: None, True, False, debounce=500)
    doc_events.drop      += DOMEventHandler(on_drop, True, False)


def _on_closing():
    if not _is_dirty:
        return True
    # Linux — GTK dialog
    try:
        import gi
        gi.require_version('Gtk', '3.0')
        from gi.repository import Gtk
        dialog = Gtk.MessageDialog(
            message_type=Gtk.MessageType.QUESTION,
            buttons=Gtk.ButtonsType.YES_NO,
            text="有未儲存的變更",
            secondary_text="確定要關閉嗎？未儲存的變更將會遺失。",
        )
        response = dialog.run()
        dialog.destroy()
        return response == Gtk.ResponseType.YES
    except Exception:
        pass
    # Windows — native MessageBox
    try:
        import ctypes
        result = ctypes.windll.user32.MessageBoxW(
            0,
            "確定要關閉嗎？未儲存的變更將會遺失。",
            "有未儲存的變更",
            0x0024,  # MB_YESNO | MB_ICONQUESTION
        )
        return result == 6  # IDYES
    except Exception:
        return True


if __name__ == "__main__":
    api = API()
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    web_dir = os.path.join(base_dir, "web")
    window = webview.create_window(
        "PDF Tools",
        url=os.path.join(web_dir, "index.html"),
        js_api=api,
        width=960,
        height=660,
        min_size=(720, 500),
    )
    window.events.closing += _on_closing
    webview.start(_bind_file_drop, window)
