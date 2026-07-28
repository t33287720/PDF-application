#!/usr/bin/env python
# coding: utf-8

import os
import sys
import base64
import tempfile
import webview
import fitz  # PyMuPDF
from pikepdf import Pdf, Permissions, Encryption


_zoom_level = 100
_is_dirty   = False


def _auto_name(directory, prefix):
    directory = directory or "."
    i = 1
    while True:
        path = os.path.join(directory, f"{prefix}_{i}.pdf")
        if not os.path.exists(path):
            return path
        i += 1


def _auto_name_image(path):
    if not os.path.exists(path):
        return path
    root, ext = os.path.splitext(path)
    i = 2
    while True:
        candidate = f"{root}_{i}{ext}"
        if not os.path.exists(candidate):
            return candidate
        i += 1


class API:

    # ── Zoom ──────────────────────────────────────────────────────────────────

    def set_dirty(self, value):
        global _is_dirty
        _is_dirty = bool(value)

    def get_zoom(self):
        return _zoom_level

    def apply_zoom(self, level):
        global _zoom_level
        _zoom_level = level
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
            file_types=("PDF Files (*.pdf)",)
        )
        return result[0] if result else ""

    def browse_save(self, default_name="output"):
        result = webview.windows[0].create_file_dialog(
            webview.FileDialog.SAVE,
            save_filename=f"{default_name}.pdf",
            file_types=("PDF Files (*.pdf)",)
        )
        if not result:
            return ""
        return result[0] if isinstance(result, (list, tuple)) else result

    def browse_folder(self):
        result = webview.windows[0].create_file_dialog(webview.FileDialog.FOLDER)
        if not result:
            return ""
        return result[0] if isinstance(result, (list, tuple)) else result

    # ── Editor ────────────────────────────────────────────────────────────────

    def open_pdf_for_editor(self, path):
        try:
            if not os.path.isfile(path):
                return {"ok": False, "msg": "PDF 不存在！"}
            doc = fitz.open(path)
            try:
                pages = []
                mat = fitz.Matrix(0.22, 0.22)
                for i in range(len(doc)):
                    pix = doc[i].get_pixmap(matrix=mat)
                    b64 = base64.b64encode(pix.tobytes("jpeg", jpg_quality=70)).decode()
                    pages.append({"index": i, "b64": f"data:image/jpeg;base64,{b64}"})
            finally:
                doc.close()
            return {"ok": True, "pages": pages}
        except Exception as e:
            return {"ok": False, "msg": str(e)}

    def save_edited_pdf(self, pages, out, password):
        """
        pages    : list of {src, orig_idx, rotation}
        out      : output path (empty = auto)
        password : encryption password (empty = none)
        """
        try:
            if not pages:
                return {"ok": False, "msg": "沒有頁面可儲存！"}

            first_dir = os.path.dirname(pages[0]["src"]) or "."
            out_path = out or _auto_name(first_dir, "edited")

            # Build new document (pages may come from different source files)
            open_docs = {}
            new_doc = fitz.open()
            try:
                try:
                    for p in pages:
                        src = p["src"]
                        if src not in open_docs:
                            open_docs[src] = fitz.open(src)
                        new_doc.insert_pdf(open_docs[src],
                                           from_page=p["orig_idx"],
                                           to_page=p["orig_idx"])
                        if p["rotation"]:
                            pg = new_doc[-1]
                            pg.set_rotation((pg.rotation + p["rotation"]) % 360)
                finally:
                    for doc in open_docs.values():
                        doc.close()

                if password:
                    # fitz doesn't encrypt; save temp then encrypt with pikepdf
                    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                        tmp_path = tmp.name
                    try:
                        new_doc.save(tmp_path)
                        new_doc.close()
                        new_doc = None        # mark closed so outer finally skips it
                        pdf = Pdf.open(tmp_path)
                        try:
                            no_extract = Permissions(extract=False)
                            pdf.save(out_path, encryption=Encryption(
                                user=password, owner=password, allow=no_extract))
                        finally:
                            pdf.close()
                    finally:
                        if os.path.exists(tmp_path):
                            os.unlink(tmp_path)
                else:
                    new_doc.save(out_path)
                    new_doc.close()
                    new_doc = None            # mark closed
            finally:
                if new_doc is not None:
                    new_doc.close()

            return {"ok": True, "msg": f"儲存成功！已儲存：{out_path}"}
        except Exception as e:
            return {"ok": False, "msg": str(e)}

    def export_pages_as_images(self, pages, out_dir, fmt, dpi, quality):
        """
        pages   : list of {src, orig_idx, rotation}
        out_dir : output directory
        fmt     : 'jpg' or 'png'
        dpi     : target resolution (72 = 100%)
        quality : JPEG quality 1-100 (ignored for png)
        """
        try:
            if not pages:
                return {"ok": False, "msg": "沒有頁面可匯出！"}
            if not out_dir or not os.path.isdir(out_dir):
                return {"ok": False, "msg": "輸出資料夾不存在！"}

            ext = "jpg" if fmt == "jpg" else "png"
            zoom = max(dpi, 36) / 72.0
            mat = fitz.Matrix(zoom, zoom)

            open_docs = {}
            count = 0
            try:
                for p in pages:
                    src = p["src"]
                    if src not in open_docs:
                        open_docs[src] = fitz.open(src)
                    doc = open_docs[src]
                    page = doc[p["orig_idx"]]
                    if p.get("rotation"):
                        page.set_rotation((page.rotation + p["rotation"]) % 360)
                    pix = page.get_pixmap(matrix=mat)

                    base = os.path.splitext(os.path.basename(src))[0]
                    out_path = os.path.join(out_dir, f"{base}_p{p['orig_idx'] + 1:03d}.{ext}")
                    out_path = _auto_name_image(out_path)
                    if fmt == "jpg":
                        pix.save(out_path, output="jpg", jpg_quality=quality)
                    else:
                        pix.save(out_path, output="png")
                    count += 1
            finally:
                for doc in open_docs.values():
                    doc.close()

            return {"ok": True, "msg": f"已匯出 {count} 張圖片至：{out_dir}"}
        except Exception as e:
            return {"ok": False, "msg": str(e)}


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
    webview.start()
