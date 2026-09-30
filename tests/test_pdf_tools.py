import os
import random
import sys

import fitz
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pdf_tools  # noqa: E402


def make_pdf(path, n, landscape=(), password=None, toc=None, image=False):
    doc = fitz.open()
    img = None
    if image:
        pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 400, 400), False)
        pix.set_rect(pix.irect, (200, 100, 50))
        img = pix.tobytes("png")
    for i in range(n):
        w, h = (842, 595) if i in landscape else (595, 842)
        page = doc.new_page(width=w, height=h)
        page.insert_text((72, 100), f"Page {i + 1}", fontsize=30)
        if img:
            page.insert_image(fitz.Rect(50, 150, 300, 400), stream=img)
    if toc:
        doc.set_toc(toc)
    opts = dict(garbage=4, deflate=True)
    if password:
        opts.update(encryption=fitz.PDF_ENCRYPT_AES_256, user_pw=password, owner_pw=password)
    doc.save(str(path), **opts)
    doc.close()
    return str(path)


def pages_of(src, *indices, rotation=0):
    return [{"src": src, "orig_idx": i, "rotation": rotation} for i in indices]


def texts(path, password=None):
    doc = fitz.open(path)
    if password:
        doc.authenticate(password)
    result = [(pg.get_text().strip(), pg.rotation) for pg in doc]
    doc.close()
    return result


@pytest.fixture
def api(tmp_path, monkeypatch):
    # Keep settings out of the real user profile
    monkeypatch.setenv("APPDATA", str(tmp_path / "appdata"))
    return pdf_tools.API()


# ── helpers ──────────────────────────────────────────────────────────────────

def test_page_runs_groups_consecutive_pages():
    pages = pages_of("a", 0, 1, 2, 5) + pages_of("b", 0) + pages_of("a", 6, 3)
    assert pdf_tools._page_runs(pages) == [
        ["a", 0, 2], ["a", 5, 5], ["b", 0, 0], ["a", 6, 6], ["a", 3, 3],
    ]


# ── opening ──────────────────────────────────────────────────────────────────

def test_open_returns_sizes_and_thumbnails(api, tmp_path):
    src = make_pdf(tmp_path / "a.pdf", 3, landscape={1})
    res = api.open_pdf_for_editor(src)
    assert res["ok"]
    assert res["sizes"][0] == [595, 842]
    assert res["sizes"][1] == [842, 595]
    thumbs = api.get_thumbnails(src, 1, 10)
    assert thumbs["ok"] and len(thumbs["thumbs"]) == 2
    assert thumbs["thumbs"][0].startswith("data:image/jpeg;base64,")


def test_open_missing_file(api, tmp_path):
    assert not api.open_pdf_for_editor(str(tmp_path / "nope.pdf"))["ok"]


def test_open_encrypted_source(api, tmp_path):
    src = make_pdf(tmp_path / "locked.pdf", 2, password="abc")
    assert api.open_pdf_for_editor(src)["need_password"]
    wrong = api.open_pdf_for_editor(src, "x")
    assert wrong["need_password"] and not wrong["ok"]
    assert api.open_pdf_for_editor(src, "abc")["ok"]
    # password is remembered for thumbnails, preview and saving
    assert api.get_thumbnails(src, 0, 2)["ok"]
    assert api.get_preview(src, 1)["ok"]
    out = str(tmp_path / "out.pdf")
    assert api.save_edited_pdf(pages_of(src, 1, 0), out)["ok"]
    assert texts(out) == [("Page 2", 0), ("Page 1", 0)]


# ── saving ───────────────────────────────────────────────────────────────────

def test_save_reorders_rotates_and_merges(api, tmp_path):
    a = make_pdf(tmp_path / "a.pdf", 3)
    b = make_pdf(tmp_path / "b.pdf", 2)
    pages = pages_of(a, 2) + pages_of(b, 1, rotation=90) + pages_of(a, 0, rotation=270)
    out = str(tmp_path / "out.pdf")
    res = api.save_edited_pdf(pages, out)
    assert res["ok"] and res["path"] == out
    assert texts(out) == [("Page 3", 0), ("Page 2", 90), ("Page 1", 270)]


def test_save_does_not_bloat_shared_resources(api, tmp_path):
    src = make_pdf(tmp_path / "img.pdf", 20, image=True)
    out = str(tmp_path / "out.pdf")
    api.save_edited_pdf(pages_of(src, *range(20)), out)
    assert os.path.getsize(out) < os.path.getsize(src) * 1.2


def test_save_remaps_bookmarks(api, tmp_path):
    toc = [[1, "Intro", 1], [1, "Part A", 2], [2, "A.1", 3], [1, "Part B", 4]]
    src = make_pdf(tmp_path / "toc.pdf", 4, toc=toc)
    out = str(tmp_path / "out.pdf")
    # drop page 2 (Part A) and move page 4 to the front
    api.save_edited_pdf(pages_of(src, 3, 0, 2), out)
    doc = fitz.open(out)
    # "A.1" lost its parent, so it is promoted to keep the outline valid
    assert doc.get_toc() == [[1, "Intro", 2], [2, "A.1", 3], [1, "Part B", 1]]
    doc.close()


def test_compress_downsamples_images(api, tmp_path):
    # a noisy (poorly compressible) 1500px image shown at ~3 inches -> ~500 DPI
    doc = fitz.open()
    page = doc.new_page()
    pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 1500, 1500), False)
    pix.set_rect(pix.irect, (255, 255, 255))
    rnd = random.Random(1)
    for _ in range(40000):
        pix.set_pixel(rnd.randrange(1500), rnd.randrange(1500), (rnd.randrange(256), 0, 0))
    page.insert_image(fitz.Rect(50, 50, 266, 266), pixmap=pix)
    src = str(tmp_path / "photo.pdf")
    doc.save(src)
    plain, small = str(tmp_path / "plain.pdf"), str(tmp_path / "small.pdf")
    api.save_edited_pdf(pages_of(src, 0), plain)
    api.save_edited_pdf(pages_of(src, 0), small, {"compress": "high"})
    assert os.path.getsize(small) < os.path.getsize(plain) / 3


def test_save_encrypted_restricts_copying(api, tmp_path):
    src = make_pdf(tmp_path / "a.pdf", 2)
    out = str(tmp_path / "enc.pdf")
    assert api.save_edited_pdf(pages_of(src, 0, 1), out, {"password": "pw"})["ok"]
    doc = fitz.open(out)
    assert doc.needs_pass
    assert doc.authenticate("pw") == 2          # user level, not owner
    assert not doc.permissions & fitz.PDF_PERM_COPY
    assert doc.permissions & fitz.PDF_PERM_PRINT
    doc.close()
    # the app can reopen its own encrypted output without asking again
    assert api.open_pdf_for_editor(out)["ok"]


def test_save_over_source(api, tmp_path):
    src = make_pdf(tmp_path / "a.pdf", 3)
    api.open_pdf_for_editor(src)
    api.get_thumbnails(src, 0, 3)               # keeps the source open
    assert api.save_edited_pdf(pages_of(src, 2, 1), src)["ok"]
    assert texts(src) == [("Page 3", 0), ("Page 2", 0)]


def test_save_nothing(api):
    assert not api.save_edited_pdf([], "x.pdf")["ok"]


# ── output path ──────────────────────────────────────────────────────────────

def test_check_output(api, tmp_path):
    res = api.check_output(str(tmp_path / "result"))
    assert res["ok"] and res["path"] == str(tmp_path / "result.pdf") and not res["exists"]
    make_pdf(tmp_path / "result.pdf", 1)
    assert api.check_output(str(tmp_path / "result.pdf"))["exists"]
    assert not api.check_output(str(tmp_path / "missing" / "x.pdf"))["ok"]


# ── settings ─────────────────────────────────────────────────────────────────

def test_settings_persist(api, tmp_path):
    assert api.get_zoom() == 100
    api.apply_zoom(130)
    api._remember_dir(str(tmp_path / "a.pdf"))
    again = pdf_tools.API()
    assert again.get_zoom() == 130
    assert again._last_dir() == str(tmp_path)


def test_corrupt_settings_fall_back(api):
    os.makedirs(os.path.dirname(pdf_tools._settings_path()), exist_ok=True)
    with open(pdf_tools._settings_path(), "w") as f:
        f.write("{not json")
    assert pdf_tools.API().get_zoom() == 100


# ── overlays (watermark / page numbers) ──────────────────────────────────────

def _render_gray(page):
    return page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5), colorspace=fitz.csGRAY)


def _same_as_upright(stamp, rotation):
    """Stamp a rotated page and an unrotated page of the displayed size the
    same way; rendered, they must look identical."""
    rotated = fitz.open().new_page(width=595, height=842)
    rotated.set_rotation(rotation)
    upright = fitz.open().new_page(width=rotated.rect.width, height=rotated.rect.height)
    stamp(rotated)
    stamp(upright)
    a, b = _render_gray(rotated), _render_gray(upright)
    assert (a.width, a.height) == (b.width, b.height)
    return sum(abs(x - y) for x, y in zip(a.samples, b.samples)) / len(a.samples) < 0.05


@pytest.mark.parametrize("rotation", [0, 90, 180, 270])
def test_stamp_text_follows_page_rotation(rotation):
    def stamp(page):
        w, h = page.rect.width, page.rect.height
        pdf_tools._stamp_text(page, "Bottom 12", 30, (w / 2, h - 40))
        pdf_tools._stamp_text(page, "R", 30, (w - 40, 40), align="right")
        pdf_tools._stamp_text(page, "DRAFT", 80, (w / 2, h / 2), angle=40)
    assert _same_as_upright(stamp, rotation)


def test_watermark(api, tmp_path):
    src = make_pdf(tmp_path / "a.pdf", 3)
    out = str(tmp_path / "wm.pdf")
    wm = {"text": "機密 DRAFT", "size": "large", "opacity": 0.3, "diagonal": True}
    assert api.save_edited_pdf(pages_of(src, 0, 1) + pages_of(src, 2, rotation=90), out,
                               {"watermark": wm})["ok"]
    doc = fitz.open(out)
    assert all("機密 DRAFT" in pg.get_text() for pg in doc)
    doc.close()
    # the CJK font is subset, not embedded whole (~3.5 MB)
    assert os.path.getsize(out) < 200_000
    assert _same_as_upright(lambda pg: pdf_tools._add_watermark(pg, wm), 90)
