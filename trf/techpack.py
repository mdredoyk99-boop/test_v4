"""Techpack reader: header info, comparison text and the care-instruction screenshot."""
from __future__ import annotations
import io
import re
import subprocess
import tempfile
from dataclasses import dataclass, field
from xml.etree import ElementTree as ET
from PIL import Image
from pdf2image import convert_from_bytes

DPI = 300


@dataclass
class Techpack:
    is_pdf: bool
    text: str = ""
    season: str = ""
    program: str = ""
    item_no: str = ""
    style: str = ""
    composition: str = ""
    care_text: list = field(default_factory=list)
    care_image: Image.Image | None = None
    page_image: Image.Image | None = None
    crop_pts: tuple | None = None       # (left, top, right, bottom) in PDF points
    page_size_pts: tuple | None = None


def _run(cmd, data: bytes) -> str:
    with tempfile.NamedTemporaryFile(suffix=".pdf") as f:
        f.write(data); f.flush()
        return subprocess.run(cmd + [f.name, "-"], capture_output=True, text=True,
                              check=True).stdout


def _words(data: bytes, page=1):
    xml = _run(["pdftotext", "-f", str(page), "-l", str(page), "-bbox"], data)
    xml = re.sub(r"<!DOCTYPE[^>]*>", "", xml)
    xml = re.sub(r'xmlns="[^"]+"', "", xml)
    root = ET.fromstring(xml)
    pg = root.find(".//page")
    size = (float(pg.get("width")), float(pg.get("height")))
    words = [(float(w.get("xMin")), float(w.get("yMin")), float(w.get("xMax")),
              float(w.get("yMax")), w.text or "") for w in root.iter("word")]
    return size, words


def detect_care_box(size, words):
    """Care band = colour strip starting at the 'Washing' header down to 'Care Status'."""
    wash = [w for w in words if w[4] == "Washing"]
    status = [w for w in words if w[4] == "Status"]
    if not wash:
        return None
    top = min(w[1] for w in wash) - 12
    if status:
        bottom = max(w[3] for w in status) + 18
    else:
        bottom = top + 85
    return (max(size[0] * 0.0, wash[0][0] - 8), top, size[0] - 14, bottom)


def crop_page(page: Image.Image, size_pts, box_pts) -> Image.Image:
    sx, sy = page.width / size_pts[0], page.height / size_pts[1]
    l, t, r, b = box_pts
    return page.crop((int(l * sx), int(t * sy), int(r * sx), int(b * sy)))


def care_lines(words, box):
    """Text lines inside the care band that are NOT the column headers (WARNING, Care Status...)."""
    l, t, r, b = box
    inside = [w for w in words if w[0] >= l and w[2] <= r and w[1] >= t and w[3] <= b]
    header_y = min((w[1] for w in inside if w[4] == "Washing"), default=t)
    rows = {}
    for w in inside:
        if w[1] > header_y + 12:           # below the header strip
            rows.setdefault(round(w[1] / 4), []).append(w)
    return [" ".join(x[4] for x in sorted(v)) for _, v in sorted(rows.items())]


def read_techpack(data: bytes, filename: str) -> Techpack:
    if not filename.lower().endswith(".pdf"):
        img = Image.open(io.BytesIO(data)).convert("RGB")
        return Techpack(is_pdf=False, care_image=img, page_image=img)

    tp = Techpack(is_pdf=True)
    tp.text = _run(["pdftotext", "-layout"], data)
    head = re.search(r"\b([A-Z]{2,3}\d{2})\s*-\s*(U[A-Z0-9\-]+)\s*-\s*S\d+\s*\((I\d+)\)\s*-\s*(.+)", tp.text)
    if head:
        tp.season, tp.program, tp.item_no = head.group(1), head.group(2), head.group(3)
        tp.style = head.group(4).strip()
    comp = re.search(r"Composition\s*:\s*(.+)", tp.text)
    tp.composition = comp.group(1).strip() if comp else ""

    size, words = _words(data, 1)
    tp.page_size_pts = size
    tp.page_image = convert_from_bytes(data, dpi=DPI, first_page=1, last_page=1)[0]
    box = detect_care_box(size, words)
    if box:
        tp.crop_pts = box
        tp.care_image = crop_page(tp.page_image, size, box)
        tp.care_text = care_lines(words, box)
    return tp


def recrop(tp: Techpack, box_pts) -> Image.Image:
    return crop_page(tp.page_image, tp.page_size_pts, box_pts)
