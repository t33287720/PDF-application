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

def blank(width=595, height=842, rotation=0):
    return {"blank": True, "width": width, "height": height, "rotation": rotation}


def test_page_runs_groups_consecutive_pages():
    pages = (pages_of("a", 0, 1, 2, 5) + pages_of("b", 0) + [blank(100, 200)]
             + pages_of("a", 6, 3))
    assert pdf_tools._page_runs(pages) == [
        ["a", 0, 2], ["a", 5, 5], ["b", 0, 0], [None, 100, 200], ["a", 6, 6], ["a", 3, 3],
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


def test_save_with_blank_pages(api, tmp_path):
    toc = [[1, "One", 1], [1, "Two", 2]]
    src = make_pdf(tmp_path / "a.pdf", 2, toc=toc)
    out = str(tmp_path / "out.pdf")
    pages = [blank(842, 595)] + pages_of(src, 0) + [blank(rotation=90)] + pages_of(src, 1)
    assert api.save_edited_pdf(pages, out)["ok"]
    doc = fitz.open(out)
    assert [(pg.rect.width, pg.rect.height, pg.rotation) for pg in doc] == [
        (842, 595, 0), (595, 842, 0), (842, 595, 90), (595, 842, 0)]
    assert doc[0].get_text() == "" and doc[1].get_text().strip() == "Page 1"
    assert doc.get_toc() == [[1, "One", 2], [1, "Two", 4]]
    doc.close()
    # a document made of blank pages only
    assert api.save_edited_pdf([blank()], str(tmp_path / "b.pdf"))["ok"]


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


def test_page_numbers(api, tmp_path):
    src = make_pdf(tmp_path / "a.pdf", 4)
    out = str(tmp_path / "pn.pdf")
    spec = {"format": "第 {n} 頁，共 {total} 頁", "position": "bottom-right",
            "start": 1, "skip_first": True}
    pages = pages_of(src, 0, 1) + pages_of(src, 2, rotation=90) + pages_of(src, 3)
    assert api.save_edited_pdf(pages, out, {"page_numbers": spec})["ok"]
    doc = fitz.open(out)
    assert "第" not in doc[0].get_text()            # cover skipped
    assert "第 1 頁，共 3 頁" in doc[1].get_text()
    assert "第 3 頁，共 3 頁" in doc[3].get_text()
    # on the rotated page the number still sits at the displayed bottom-right
    words = doc[2].get_text("words")
    num = fitz.Rect(words[-1][:4]) * doc[2].rotation_matrix
    assert num.x1 > doc[2].rect.width * 0.8 and num.y1 > doc[2].rect.height * 0.9
    doc.close()


@pytest.mark.parametrize("rotation", [90, 270])
def test_page_number_follows_rotation(rotation):
    assert _same_as_upright(lambda pg: pdf_tools._add_page_number(pg, "12 / 30", "bottom-right"), rotation)


# ── images ───────────────────────────────────────────────────────────────────

def test_open_image_converts_to_pdf(api, tmp_path):
    pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 400, 300), False)
    pix.set_rect(pix.irect, (30, 120, 200))
    img = str(tmp_path / "photo.png")
    pix.save(img)
    res = api.open_pdf_for_editor(img)
    assert res["ok"] and res["name"] == "photo.png"
    assert res["path"].endswith(".pdf") and res["path"] != img
    assert res["sizes"] == [[300, 225]]          # 400x300 px at 96 DPI
    assert api.get_thumbnails(res["path"], 0, 1)["ok"]
    out = str(tmp_path / "out.pdf")
    assert api.save_edited_pdf(pages_of(res["path"], 0, rotation=90), out)["ok"]
    doc = fitz.open(out)
    assert len(doc) == 1 and doc[0].get_images()
    doc.close()


def test_open_broken_image(api, tmp_path):
    bad = tmp_path / "broken.jpg"
    bad.write_bytes(b"not an image")
    assert not api.open_pdf_for_editor(str(bad))["ok"]


def test_export_images(api, tmp_path):
    src = make_pdf(tmp_path / "a.pdf", 3, landscape={2})
    out_dir = tmp_path / "imgs"
    out_dir.mkdir()
    pages = pages_of(src, 0) + [blank(200, 100)] + pages_of(src, 2, rotation=90)
    res = api.export_images(pages, str(out_dir), "png", 72, {"page_numbers": {}}, "my/doc")
    assert res["ok"] and res["count"] == 3
    names = sorted(os.listdir(out_dir))
    assert names == ["mydoc_001.png", "mydoc_002.png", "mydoc_003.png"]
    sizes = [(pix.width, pix.height) for pix in (fitz.Pixmap(str(out_dir / n)) for n in names)]
    assert sizes == [(595, 842), (200, 100), (595, 842)]   # landscape page turned upright
    # a second export never overwrites
    api.export_images(pages_of(src, 0), str(out_dir), "jpg", 50, None, "mydoc")
    api.export_images(pages_of(src, 0), str(out_dir), "jpg", 50, None, "mydoc")
    assert {"mydoc_001.jpg", "mydoc_001 (2).jpg"} <= set(os.listdir(out_dir))
    # lower JPG quality -> smaller file
    api.export_images(pages_of(src, 0), str(out_dir), "jpg", 150, None, "q", quality=95)
    api.export_images(pages_of(src, 0), str(out_dir), "jpg", 150, None, "q", quality=20)
    assert os.path.getsize(out_dir / "q_001 (2).jpg") < os.path.getsize(out_dir / "q_001.jpg")


def test_split_pdf(api, tmp_path):
    src = make_pdf(tmp_path / "a.pdf", 5)
    out_dir = tmp_path / "parts"
    out_dir.mkdir()
    (out_dir / "a_1-2.pdf").write_bytes(b"existing")          # must not be overwritten
    groups = [{"name": "a_1-2", "pages": pages_of(src, 0, 1)},
              {"name": "a_5-4", "pages": pages_of(src, 4, 3)},
              {"name": "a_3", "pages": pages_of(src, 2) + [blank()]}]
    res = api.split_pdf(groups, str(out_dir), {"password": "pw"})
    assert res["ok"] and res["count"] == 3
    assert (out_dir / "a_1-2.pdf").read_bytes() == b"existing"
    assert texts(str(out_dir / "a_1-2 (2).pdf"), "pw") == [("Page 1", 0), ("Page 2", 0)]
    assert texts(str(out_dir / "a_5-4.pdf"), "pw") == [("Page 5", 0), ("Page 4", 0)]
    assert texts(str(out_dir / "a_3.pdf"), "pw") == [("Page 3", 0), ("", 0)]
    assert not api.split_pdf(groups, str(tmp_path / "missing"))["ok"]


def test_metadata_clear_removes_document_info(api, tmp_path):
    src = str(tmp_path / "meta.pdf")
    doc = fitz.open()
    doc.new_page()
    doc.set_metadata({"title": "T", "author": "Someone", "creator": "App"})
    doc.save(src)
    doc.close()
    out = str(tmp_path / "clean.pdf")
    assert api.save_edited_pdf(pages_of(src, 0), out, {"meta_mode": "clear"})["ok"]
    doc = fitz.open(out)
    assert not any(doc.metadata.get(k) for k in ("title", "author", "creator", "creationDate", "modDate"))
    doc.close()


def test_metadata_custom_and_with_encryption(api, tmp_path):
    src = make_pdf(tmp_path / "a.pdf", 1)
    out = str(tmp_path / "custom.pdf")
    opts = {"meta_mode": "custom", "title": "合約", "author": "我", "password": "pw"}
    assert api.save_edited_pdf(pages_of(src, 0), out, opts)["ok"]
    doc = fitz.open(out)
    assert doc.authenticate("pw")
    assert doc.metadata["title"] == "合約" and doc.metadata["author"] == "我"
    assert not doc.metadata["creator"]
    doc.close()


def test_open_returns_source_toc(api, tmp_path):
    src = make_pdf(tmp_path / "a.pdf", 3, toc=[[1, "One", 1], [2, "Sub", 3]])
    res = api.open_pdf_for_editor(src)
    assert res["ok"] and res["toc"] == [[1, "One", 0], [2, "Sub", 2]]


def test_edited_bookmarks_follow_page_order_and_deletion(api, tmp_path):
    src = make_pdf(tmp_path / "a.pdf", 3, toc=[[1, "Old", 1]])
    pages = [dict(p, bid=i) for i, p in enumerate(pages_of(src, 2, 0, 1), start=1)]
    marks = [
        {"level": 3, "title": "第一章", "bid": 2},     # level too deep: clamped to previous + 1
        {"level": 2, "title": "小節", "bid": 3},
        {"level": 1, "title": "附錄", "bid": 1},
        {"level": 1, "title": "已刪除頁", "bid": 99},   # page not in output: dropped
        {"level": 1, "title": "  ", "bid": 1},          # empty title: dropped
    ]
    out = str(tmp_path / "out.pdf")
    assert api.save_edited_pdf(pages, out, {"bookmarks": marks})["ok"]
    doc = fitz.open(out)
    assert doc.get_toc() == [[1, "附錄", 1], [2, "第一章", 2], [2, "小節", 3]]
    doc.close()


def test_edited_bookmarks_empty_list_removes_all_and_none_keeps_source(api, tmp_path):
    src = make_pdf(tmp_path / "a.pdf", 2, toc=[[1, "One", 1]])
    pages = [dict(p, bid=i) for i, p in enumerate(pages_of(src, 0, 1), start=1)]
    cleared, kept = str(tmp_path / "cleared.pdf"), str(tmp_path / "kept.pdf")
    assert api.save_edited_pdf(pages, cleared, {"bookmarks": []})["ok"]
    assert api.save_edited_pdf(pages, kept, {})["ok"]
    for path, expected in ((cleared, []), (kept, [[1, "One", 1]])):
        doc = fitz.open(path)
        assert doc.get_toc() == expected
        doc.close()


@pytest.mark.parametrize("src_rot", [0, 90, 180, 270])
@pytest.mark.parametrize("edit_rot", [0, 90])
def test_crop_trims_margins_for_every_rotation(api, tmp_path, src_rot, edit_rot):
    src = make_pdf(tmp_path / "a.pdf", 1)
    doc = fitz.open(src)
    doc[0].set_rotation(src_rot)
    before = doc[0].search_for("Page 1")[0] * doc[0].rotation_matrix   # search_for ignores /Rotate
    before.normalize()
    view = doc[0].rect
    doc.save(str(tmp_path / "rot.pdf"))
    doc.close()

    left, top, right, bottom = 50, 60, 20, 10
    out = str(tmp_path / "out.pdf")
    page = {"src": str(tmp_path / "rot.pdf"), "orig_idx": 0, "rotation": edit_rot,
            "crop": [left, top, right, bottom]}
    assert api.save_edited_pdf([page], out)["ok"]

    res = fitz.open(out)
    pg = res[0]
    w, h = view.width - left - right, view.height - top - bottom
    expected = (w, h) if edit_rot == 0 else (h, w)
    assert (round(pg.rect.width), round(pg.rect.height)) == tuple(round(v) for v in expected)
    if edit_rot == 0:   # the text moved by exactly the cropped margins
        after = pg.search_for("Page 1")[0] * pg.rotation_matrix
        after.normalize()
        assert after.x0 == pytest.approx(before.x0 - left, abs=0.5)
        assert after.y0 == pytest.approx(before.y0 - top, abs=0.5)
    res.close()


def test_crop_skips_blank_pages_and_oversized_margins(api, tmp_path):
    src = make_pdf(tmp_path / "a.pdf", 1)
    out = str(tmp_path / "out.pdf")
    pages = [dict(blank(300, 400), crop=[10, 10, 10, 10]),
             {"src": src, "orig_idx": 0, "rotation": 0, "crop": [400, 0, 400, 0]}]
    assert api.save_edited_pdf(pages, out)["ok"]
    res = fitz.open(out)
    assert (res[0].rect.width, res[0].rect.height) == (300, 400)
    assert (res[1].rect.width, res[1].rect.height) == (595, 842)
    res.close()


def test_crop_respects_existing_cropbox(api, tmp_path):
    src = make_pdf(tmp_path / "a.pdf", 1)
    doc = fitz.open(src)
    doc[0].set_cropbox(fitz.Rect(30, 40, 500, 700))
    doc.save(str(tmp_path / "cb.pdf"))
    doc.close()
    out = str(tmp_path / "out.pdf")
    page = {"src": str(tmp_path / "cb.pdf"), "orig_idx": 0, "rotation": 0, "crop": [10, 20, 0, 0]}
    assert api.save_edited_pdf([page], out)["ok"]
    res = fitz.open(out)
    assert (round(res[0].rect.width), round(res[0].rect.height)) == (460, 640)
    res.close()
