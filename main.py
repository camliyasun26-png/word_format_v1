# main.py
import asyncio
import io
import os
import re
import shutil
import tempfile
import time
import uuid
import zipfile
from xml.etree import ElementTree as ET

import json

import mammoth
from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from converter.doc_converter import is_doc, to_doc, to_docx
from formatter.applier import apply_format_to_paragraph, apply_page_margins, center_tables_and_images
from formatter.preview import render_paragraph_html
from formatter.template_loader import list_templates, load_template
from parser.context import ChapterContext
from parser.rules import classify_paragraph

EMU_PER_CM = 360000
PT_PER_CM = 72 / 2.54
_WP_DRAWING_NS = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
_VML_NS = "{urn:schemas-microsoft-com:vml}"
_MC_NS = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
_VML_STYLE_WIDTH_RE = re.compile(r"(?:^|;)\s*width\s*:\s*([\d.]+)\s*([a-z%]*)", re.I)


def _vml_shape_width_cm(shape) -> float | None:
    if shape.find(f"{_VML_NS}imagedata") is None:
        return None
    style = shape.get("style") or ""
    m = _VML_STYLE_WIDTH_RE.search(style)
    if not m:
        return None
    try:
        val = float(m.group(1))
    except ValueError:
        return None
    unit = (m.group(2) or "pt").lower()
    if unit in ("pt", ""):
        return val / PT_PER_CM
    if unit == "in":
        return val * 2.54
    if unit == "cm":
        return val
    if unit == "mm":
        return val / 10
    if unit == "px":
        return val / 96 * 2.54
    return None


def _extract_image_widths_cm(docx_bytes: bytes) -> list[float]:
    widths = []
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
            with zf.open("word/document.xml") as f:
                tree = ET.parse(f)
        skip = set()
        for fb in tree.iter(f"{_MC_NS}Fallback"):
            for d in fb.iter():
                skip.add(id(d))
        for el in tree.iter():
            if id(el) in skip:
                continue
            if el.tag == f"{_WP_DRAWING_NS}extent":
                cx = el.get("cx")
                if cx and cx.isdigit():
                    widths.append(int(cx) / EMU_PER_CM)
            elif el.tag == f"{_VML_NS}shape":
                w = _vml_shape_width_cm(el)
                if w is not None:
                    widths.append(w)
    except (KeyError, ET.ParseError, zipfile.BadZipFile):
        pass
    return widths


def _inject_image_widths(html: str, widths_cm: list[float]) -> str:
    it = iter(widths_cm)

    def repl(m):
        w = next(it, None)
        if w is None:
            return m.group(0)
        return f'<img style="width:{w:.2f}cm;height:auto;max-width:100%" '

    return re.sub(r'<img(?=\s)', repl, html)


_W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _extract_media_anchors(docx_bytes: bytes) -> list[dict]:
    """Walk top-level body children, return ordered media anchors.

    Each anchor: {id, kind: 'image'|'table', after_para_index}.
    after_para_index = -1 means before the first paragraph.
    Image inside a paragraph anchors to that paragraph's index.
    """
    anchors: list[dict] = []
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
            with zf.open("word/document.xml") as f:
                tree = ET.parse(f)
    except (KeyError, ET.ParseError, zipfile.BadZipFile):
        return anchors

    body = tree.getroot().find(f"{_W_NS}body")
    if body is None:
        return anchors

    skip_ids = set()
    for fb in body.iter(f"{_MC_NS}Fallback"):
        for d in fb.iter():
            skip_ids.add(id(d))

    para_idx = -1
    media_counter = 0
    for child in list(body):
        if child.tag == f"{_W_NS}p":
            para_idx += 1
            for el in child.iter():
                if id(el) in skip_ids:
                    continue
                if el.tag in (f"{_WP_DRAWING_NS}extent", f"{_VML_NS}shape"):
                    if el.tag == f"{_VML_NS}shape" and el.find(f"{_VML_NS}imagedata") is None:
                        continue
                    anchors.append({
                        "id": f"media-{media_counter}",
                        "kind": "image",
                        "after_para_index": para_idx,
                    })
                    media_counter += 1
        elif child.tag == f"{_W_NS}tbl":
            anchors.append({
                "id": f"media-{media_counter}",
                "kind": "table",
                "after_para_index": para_idx,
            })
            media_counter += 1
    return anchors


def _inject_media_ids(html: str, anchors: list[dict]) -> str:
    """Assign sequential id="media-N" to each <img> and <table> in HTML order."""
    counter = [0]
    kinds = [a["kind"] for a in anchors]

    def repl(m):
        tag = m.group(1).lower()
        i = counter[0]
        if i >= len(kinds):
            return m.group(0)
        expected = "image" if tag == "img" else "table"
        if kinds[i] != expected:
            return m.group(0)
        counter[0] += 1
        return f'<{m.group(1)} id="media-{i}"'

    return re.sub(r'<(img|table)(?=[\s>])', repl, html, flags=re.I)


TEMPLATES_DIR = "templates"
UPLOADS_DIR = "uploads"
SESSION_TTL_SECONDS = 2 * 3600
CLEANUP_INTERVAL_SECONDS = 30 * 60

os.makedirs(UPLOADS_DIR, exist_ok=True)

app = FastAPI()
app.mount("/static", StaticFiles(directory="frontend"), name="static")


def _session_dir(doc_id: str) -> str:
    safe = re.fullmatch(r"[0-9a-fA-F-]{8,64}", doc_id or "")
    if not safe:
        raise HTTPException(status_code=400, detail="invalid doc_id")
    return os.path.join(UPLOADS_DIR, doc_id)


def _touch(path: str) -> None:
    try:
        os.utime(path, None)
    except OSError:
        pass


def _cleanup_expired_sessions() -> None:
    now = time.time()
    if not os.path.isdir(UPLOADS_DIR):
        return
    for name in os.listdir(UPLOADS_DIR):
        d = os.path.join(UPLOADS_DIR, name)
        if not os.path.isdir(d):
            continue
        try:
            age = now - os.path.getmtime(d)
        except OSError:
            continue
        if age > SESSION_TTL_SECONDS:
            shutil.rmtree(d, ignore_errors=True)


async def _cleanup_loop() -> None:
    while True:
        try:
            _cleanup_expired_sessions()
        except Exception:
            pass
        await asyncio.sleep(CLEANUP_INTERVAL_SECONDS)


@app.on_event("startup")
async def _start_cleanup() -> None:
    _cleanup_expired_sessions()
    asyncio.create_task(_cleanup_loop())


def _build_original_html(docx_bytes: bytes) -> str:
    style_map = """
p[style-name='Heading 1'] => h1:fresh
p[style-name='Heading 2'] => h2:fresh
"""
    result = mammoth.convert_to_html(io.BytesIO(docx_bytes), style_map=style_map)
    widths = _extract_image_widths_cm(docx_bytes)
    html = _inject_image_widths(result.value, widths)
    anchors = _extract_media_anchors(docx_bytes)
    return _inject_media_ids(html, anchors)


def _parse_docx(docx_path: str) -> list[dict]:
    doc = Document(docx_path)
    ctx = ChapterContext()
    paragraphs = []
    for i, para in enumerate(doc.paragraphs):
        result = classify_paragraph(para, ctx)
        result["index"] = i
        if result["detected_level"] != "Body":
            ctx.push(result["detected_level"], para.text[:20])
        paragraphs.append(result)
    return paragraphs


@app.get("/")
def index():
    return FileResponse("frontend/index.html")


@app.get("/templates")
def get_templates():
    return list_templates(TEMPLATES_DIR)


@app.get("/template")
def get_template(name: str = "report_cn.yaml"):
    return load_template(os.path.join(TEMPLATES_DIR, name))


@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    # 验证文件扩展名
    filename = file.filename.lower()
    if not (filename.endswith(".doc") or filename.endswith(".docx")):
        raise HTTPException(status_code=400, detail="仅支持 .doc 和 .docx 格式的文件")
    
    doc_id = uuid.uuid4().hex
    sess = _session_dir(doc_id)
    os.makedirs(sess, exist_ok=True)

    raw_path = os.path.join(sess, "original_" + file.filename)
    with open(raw_path, "wb") as f:
        f.write(await file.read())

    was_doc = is_doc(raw_path)
    try:
        if was_doc:
            docx_path = to_docx(raw_path, sess)
        else:
            docx_path = raw_path

        paragraphs = _parse_docx(docx_path)
    except Exception as e:
        # 清理已创建的会话目录
        shutil.rmtree(sess, ignore_errors=True)
        raise HTTPException(status_code=400, detail=f"无法解析文档: {str(e)}")

    with open(docx_path, "rb") as f:
        docx_bytes = f.read()
    media_anchors = _extract_media_anchors(docx_bytes)
    original_html = _build_original_html(docx_bytes)

    meta = {
        "doc_id": doc_id,
        "filename": file.filename,
        "was_doc": was_doc,
        "raw_path": raw_path,
        "docx_path": docx_path,
        "paragraphs": paragraphs,
        "media_anchors": media_anchors,
        "original_html": original_html,
    }
    with open(os.path.join(sess, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False)

    return {
        "doc_id": doc_id,
        "filename": file.filename,
        "was_doc": was_doc,
        "paragraphs": paragraphs,
        "media_anchors": media_anchors,
        "original_html": original_html,
    }


@app.get("/session/{doc_id}")
def get_session(doc_id: str):
    sess = _session_dir(doc_id)
    meta_path = os.path.join(sess, "meta.json")
    if not os.path.isfile(meta_path):
        raise HTTPException(status_code=404, detail="session not found or expired")
    _touch(sess)
    _touch(meta_path)
    with open(meta_path, "r", encoding="utf-8") as f:
        meta = json.load(f)
    return {
        "doc_id": meta["doc_id"],
        "filename": meta["filename"],
        "was_doc": meta["was_doc"],
        "paragraphs": meta["paragraphs"],
        "media_anchors": meta["media_anchors"],
        "original_html": meta["original_html"],
    }


@app.delete("/session/{doc_id}")
def delete_session(doc_id: str):
    sess = _session_dir(doc_id)
    shutil.rmtree(sess, ignore_errors=True)
    return {"ok": True}


@app.post("/original-preview")
async def original_preview(file: UploadFile = File(...)):
    content = await file.read()
    return {"html": _build_original_html(content)}


class PreviewRequest(BaseModel):
    text: str
    level: str
    template: str = "report_cn.yaml"


@app.post("/preview")
def preview(req: PreviewRequest):
    tpl = load_template(os.path.join(TEMPLATES_DIR, req.template))
    html = render_paragraph_html(req.text, req.level, tpl)
    return {"html": html}


class PreviewBatchItem(BaseModel):
    text: str
    level: str


class PreviewBatchRequest(BaseModel):
    items: list[PreviewBatchItem]
    template: str = "report_cn.yaml"


@app.post("/preview/batch")
def preview_batch(req: PreviewBatchRequest):
    tpl = load_template(os.path.join(TEMPLATES_DIR, req.template))
    return {"htmls": [render_paragraph_html(it.text, it.level, tpl) for it in req.items]}


def _replace_para_text(para: Paragraph, new_text: str):
    """把段落的所有 run 合并成一个，文本替换为 new_text，保留第一个 run 的 rPr。"""
    runs = para.runs
    if not runs:
        para.add_run(new_text)
        return
    runs[0].text = new_text
    for r in runs[1:]:
        r._element.getparent().remove(r._element)


def _insert_para_after(para: Paragraph, text: str) -> Paragraph:
    """在 para 之后插入一个新 w:p，复制 para 的 pPr，文本为 text。返回新 Paragraph。"""
    new_p = OxmlElement("w:p")
    pPr = para._element.find(qn_w("pPr"))
    if pPr is not None:
        from copy import deepcopy
        new_p.append(deepcopy(pPr))
    para._element.addnext(new_p)
    new_para = Paragraph(new_p, para._parent)
    new_para.add_run(text)
    return new_para


def qn_w(tag: str) -> str:
    return "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}" + tag


@app.post("/export")
async def export_doc(
    background_tasks: BackgroundTasks,
    paragraphs: str = Form(...),
    template: str = Form("report_cn.yaml"),
    doc_id: str = Form(...),
):
    items = json.loads(paragraphs)
    tpl = load_template(os.path.join(TEMPLATES_DIR, template))

    sess = _session_dir(doc_id)
    meta_path = os.path.join(sess, "meta.json")
    if not os.path.isfile(meta_path):
        raise HTTPException(status_code=404, detail="session not found or expired")
    _touch(sess)
    _touch(meta_path)
    with open(meta_path, "r", encoding="utf-8") as f:
        meta = json.load(f)

    src_docx = meta["docx_path"]
    filename = meta["filename"]
    was_doc = meta["was_doc"]

    out_dir = tempfile.mkdtemp()
    background_tasks.add_task(shutil.rmtree, out_dir, True)

    base = os.path.splitext(filename)[0]
    out_docx = os.path.join(out_dir, f"{base}_formatted.docx")
    shutil.copy2(src_docx, out_docx)

    doc = Document(out_docx)
    apply_page_margins(doc, tpl)

    para_list = list(doc.paragraphs)
    by_index = {int(it["index"]): it for it in items if "index" in it}

    for idx, para in enumerate(para_list):
        item = by_index.get(idx)
        if item is None:
            continue
        level = item.get("detected_level", "Body")

        if item.get("is_split") and item.get("split_title") and item.get("split_body"):
            _replace_para_text(para, item["split_title"])
            apply_format_to_paragraph(para, level, tpl)
            body_para = _insert_para_after(para, item["split_body"])
            apply_format_to_paragraph(body_para, "Body", tpl)
        else:
            apply_format_to_paragraph(para, level, tpl)

    center_tables_and_images(doc)

    doc.save(out_docx)

    if was_doc:
        out_path = to_doc(out_docx, out_dir)
        media_type = "application/msword"
    else:
        out_path = out_docx
        media_type = ("application/vnd.openxmlformats-officedocument"
                      ".wordprocessingml.document")

    return FileResponse(out_path, media_type=media_type,
                        filename=os.path.basename(out_path))
