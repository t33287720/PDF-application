#!/usr/bin/env python
# coding: utf-8

import os
import sys
import atexit
import base64
import json
import math
import secrets
import shutil
import tempfile
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

# Images are converted to a temporary PDF (one page per image / TIFF frame)
IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff")
OPEN_FILE_TYPES = (
    "PDF 與圖片 (*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff)",
    "PDF (*.pdf)",
    "圖片 (*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff)",
)


def _is_image(path):
    return path.lower().endswith(IMAGE_EXTS)


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


def _unique_path(path):
    """`path`, or `name (2).ext`, `name (3).ext`… if it is taken."""
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(path)
    i = 2
    while os.path.exists(f"{stem} ({i}){ext}"):
        i += 1
    return f"{stem} ({i}){ext}"


def _safe_prefix(name, default="page"):
    name = "".join(c for c in (name or "") if c not in '\\/:*?"<>|').strip()
    return name or default


def _page_runs(pages):
    """Group pages into [src, first, last] runs of ascending consecutive indices.
    A blank page is a run of its own: [None, width, height]."""
    runs = []
    for p in pages:
        if p.get("blank"):
            runs.append([None, p["width"], p["height"]])
            continue
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
        if not p.get("blank"):
            new_pos.setdefault((p["src"], p["orig_idx"]), n)
    toc, seen = [], set()
    for p in pages:
        src = p.get("src")
        if p.get("blank") or src in seen:
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


_fonts = {}


def _font_for(text):
    """Base-14 Helvetica for plain ASCII, the built-in CJK font otherwise."""
    name = "helv" if text.isascii() else "cjk"
    if name not in _fonts:
        _fonts[name] = fitz.Font(name)
    return _fonts[name]


def _stamp_text(page, text, fontsize, pos, angle=0.0, align="center",
                color=(0, 0, 0), opacity=1.0):
    """Write text centred vertically on `pos`, both given as the page is
    displayed (after its /Rotate), turned `angle` degrees counter-clockwise."""
    font = _font_for(text)
    width = font.text_length(text, fontsize)
    pivot = fitz.Point(pos) * page.derotation_matrix
    dx = {"left": 0, "center": -width / 2, "right": -width}[align]
    rad = math.radians(angle)   # y points down, so this turns counter-clockwise
    turn = fitz.Matrix(math.cos(rad), math.sin(rad), -math.sin(rad), math.cos(rad), 0, 0)
    r = page.rotation_matrix
    upright = fitz.Matrix(r.a, r.b, r.c, r.d, 0, 0)   # undo /Rotate for the glyphs
    writer = fitz.TextWriter(page.rect)
    writer.append(fitz.Point(pivot.x + dx, pivot.y + fontsize * 0.35), text,
                  font=font, fontsize=fontsize)
    writer.write_text(page, color=color, opacity=opacity, morph=(pivot, turn * upright))


WATERMARK_SPAN = {"small": 0.35, "medium": 0.55, "large": 0.75}


def _add_watermark(page, wm):
    """wm: {text, size: small|medium|large, opacity: 0-1, diagonal: bool}"""
    text = wm["text"]
    w, h = page.rect.width, page.rect.height
    diagonal = wm.get("diagonal", True)
    angle = math.degrees(math.atan2(h, w)) if diagonal else 0.0
    span = WATERMARK_SPAN.get(wm.get("size"), 0.55) * (math.hypot(w, h) if diagonal else w)
    fontsize = span / max(_font_for(text).text_length(text, 1), 1e-6)
    _stamp_text(page, text, fontsize, (w / 2, h / 2), angle=angle,
                color=(0.5, 0.5, 0.5), opacity=float(wm.get("opacity", 0.25)))


PAGE_NUMBER_MARGIN = 28   # distance from the page edge, in points


def _add_page_number(page, text, position="bottom-center", fontsize=11):
    """position: {top|bottom}-{left|center|right}, relative to the page as displayed"""
    w, h = page.rect.width, page.rect.height
    vertical, horizontal = position.split("-")
    y = PAGE_NUMBER_MARGIN if vertical == "top" else h - PAGE_NUMBER_MARGIN
    x = {"left": PAGE_NUMBER_MARGIN, "center": w / 2, "right": w - PAGE_NUMBER_MARGIN}[horizontal]
    _stamp_text(page, text, fontsize, (x, y), align=horizontal)


def _apply_metadata(doc, options):
    """meta_mode "clear" wipes all document info (and XMP); "custom" keeps only
    the given title / author; anything else leaves the document untouched.
    Runs before saving, so encryption covers it too."""
    mode = options.get("meta_mode")
    if mode not in ("clear", "custom"):
        return
    meta = {key: "" for key in ("title", "author", "subject", "keywords", "creator",
                                "producer", "creationDate", "modDate", "trapped")}
    if mode == "custom":
        meta["title"] = str(options.get("title") or "")
        meta["author"] = str(options.get("author") or "")
    doc.set_metadata(meta)
    doc.del_xml_metadata()


def _add_page_numbers(doc, spec):
    """spec: {format ('{n}', '{total}' placeholders), position, start, skip_first, size}"""
    fmt = spec.get("format") or "{n}"
    start = int(spec.get("start", 1))
    skip = 1 if spec.get("skip_first") else 0
    last = start + len(doc) - skip - 1
    for i, page in enumerate(doc):
        if i < skip:
            continue
        text = fmt.replace("{n}", str(start + i - skip)).replace("{total}", str(last))
        _add_page_number(page, text, spec.get("position", "bottom-center"),
                         float(spec.get("size", 11)))


class API:

    def __init__(self):
        self._passwords = {}   # src path -> password for encrypted source PDFs
        # pywebview runs each JS call on its own thread and MuPDF isn't
        # thread-safe, so every fitz operation goes through this lock
        self._lock = threading.RLock()
        self._thumb_doc = None  # (path, doc) kept open between thumbnail batches
        self._settings = _load_settings()
        self._tmp_dir = None    # holds PDFs converted from images

    def _image_to_pdf(self, path):
        if not self._tmp_dir:
            self._tmp_dir = tempfile.mkdtemp(prefix="pdf_tools_")
            atexit.register(shutil.rmtree, self._tmp_dir, ignore_errors=True)
        stem = os.path.splitext(os.path.basename(path))[0]
        fd, pdf_path = tempfile.mkstemp(prefix=f"{stem}_", suffix=".pdf", dir=self._tmp_dir)
        os.close(fd)
        img = fitz.open(path)
        try:
            pdf = fitz.open("pdf", img.convert_to_pdf())   # honours EXIF orientation
        finally:
            img.close()
        try:
            pdf.save(pdf_path, garbage=4, deflate=True)
        finally:
            pdf.close()
        return pdf_path

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
            allow_multiple=True,
            file_types=OPEN_FILE_TYPES,
        )
        if not result:
            return []
        self._remember_dir(result[0])
        return list(result)

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

    def browse_folder(self):
        result = webview.windows[0].create_file_dialog(
            webview.FileDialog.FOLDER,
            directory=self._last_dir(),
        )
        if not result:
            return ""
        folder = result[0] if isinstance(result, (list, tuple)) else result
        self._remember_dir(os.path.join(folder, ""))
        return folder

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
        are fetched afterwards in batches through get_thumbnails().
        Images are converted first; `path` in the result is the PDF to use."""
        try:
            if not os.path.isfile(path):
                return {"ok": False, "msg": "檔案不存在！"}
            name = os.path.basename(path)
            with self._lock:
                if _is_image(path):
                    path = self._image_to_pdf(path)
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
            return {"ok": True, "sizes": sizes, "path": path, "name": name}
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

    def _build_doc(self, pages, options):
        """Assemble the pages (which may come from different source files)
        into a new in-memory document, with watermark / page numbers applied.
        The caller closes it."""
        open_docs = {}
        new_doc = fitz.open()
        try:
            try:
                for p in pages:
                    if not p.get("blank") and p["src"] not in open_docs:
                        open_docs[p["src"]] = self._open_src(p["src"])
                # Insert consecutive pages of the same source in one call,
                # which also keeps internal links between them
                for src, a, b in _page_runs(pages):
                    if src is None:
                        new_doc.new_page(width=a, height=b)
                    else:
                        new_doc.insert_pdf(open_docs[src], from_page=a, to_page=b)
                toc = _remap_toc(pages, open_docs)   # needs the sources still open
            finally:
                for doc in open_docs.values():
                    doc.close()
            for pg, p in zip(new_doc, pages):
                if p["rotation"]:
                    pg.set_rotation((pg.rotation + p["rotation"]) % 360)
            if toc:
                new_doc.set_toc(toc)

            watermark = options.get("watermark")
            if watermark and watermark.get("text"):
                for pg in new_doc:
                    _add_watermark(pg, watermark)
            page_numbers = options.get("page_numbers")
            if page_numbers:
                _add_page_numbers(new_doc, page_numbers)
            if watermark or page_numbers:
                new_doc.subset_fonts()   # embed only the glyphs actually used
            _apply_metadata(new_doc, options)
            return new_doc
        except Exception:
            new_doc.close()
            raise

    def _build_and_save(self, pages, out_path, options):
        new_doc = self._build_doc(pages, options)
        try:
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

    def split_pdf(self, groups, folder, options=None):
        """groups: [{name, pages}]; each is saved as <folder>/<name>.pdf, never
        overwriting an existing file."""
        try:
            if not groups:
                return {"ok": False, "msg": "沒有要輸出的檔案！"}
            if not os.path.isdir(folder):
                return {"ok": False, "msg": f"資料夾不存在：{folder}"}
            count = 0
            with self._lock:
                self._close_thumb_doc()
                for group in groups:
                    name = _safe_prefix(group.get("name"), "part")
                    path = _unique_path(os.path.join(folder, f"{name}.pdf"))
                    self._build_and_save(group["pages"], path, options or {})
                    count += 1
            return {"ok": True, "count": count, "msg": f"已分割成 {count} 個檔案：{folder}"}
        except Exception as e:
            return {"ok": False, "msg": str(e)}

    def export_images(self, pages, folder, fmt="png", dpi=150, options=None, prefix="",
                      quality=90):
        """Render each page to <folder>/<prefix>_001.<fmt>… (never overwriting);
        watermark / page number options apply, the PDF-only ones don't."""
        try:
            if not pages:
                return {"ok": False, "msg": "沒有頁面可匯出！"}
            if not os.path.isdir(folder):
                return {"ok": False, "msg": f"資料夾不存在：{folder}"}
            fmt = "jpg" if fmt == "jpg" else "png"
            zoom = float(dpi) / 72
            prefix = _safe_prefix(prefix)
            count = 0
            with self._lock:
                doc = self._build_doc(pages, options or {})
                try:
                    for i, page in enumerate(doc, start=1):
                        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                        path = _unique_path(os.path.join(folder, f"{prefix}_{i:03d}.{fmt}"))
                        if fmt == "jpg":
                            pix.save(path, jpg_quality=int(quality))
                        else:
                            pix.save(path)
                        count += 1
                finally:
                    doc.close()
            return {"ok": True, "count": count, "msg": f"已匯出 {count} 張圖片到：{folder}"}
        except Exception as e:
            return {"ok": False, "msg": str(e)}

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
        pages   : list of {src, orig_idx, rotation} or {blank: true, width, height, rotation}
        out     : output path (empty = auto)
        options : {password, compress: "none" | "medium" | "high",
                   watermark: {text, size, opacity, diagonal},
                   page_numbers: {format, position, start, skip_first, size},
                   meta_mode: "keep" | "clear" | "custom", title, author}
        """
        options = options or {}
        password = options.get("password", "")
        try:
            if not pages:
                return {"ok": False, "msg": "沒有頁面可儲存！"}

            first_src = next((p["src"] for p in pages if not p.get("blank")), "")
            first_dir = os.path.dirname(first_src) or "."
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
                 if f.get("pywebviewFullPath", "").lower().endswith((".pdf",) + IMAGE_EXTS)]
        if paths:
            window.evaluate_js(f"handleDroppedFiles({json.dumps(paths)})")
        elif files:
            window.evaluate_js("showToast('只能拖入 PDF 或圖片檔案！', false)")

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
