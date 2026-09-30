"""Fill the tagged TRF template: string replacement only (no bookmarks), logo untouched."""
from __future__ import annotations
import copy
import hashlib
import io
import math
import re
import zipfile
from dataclasses import dataclass, field
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Emu
from PIL import Image
from . import config as C
from .docx_utils import (TOKEN_RE, own_runs, replace_tokens_in_paragraph, set_symbol_font,
                         run_text, para_text, make_token_run)

CARE_WIDTH_CM = 10.76          # same width as the picture in the original template
LINE_HEIGHT_CM = 0.42          # 9 pt paragraph lines in the care-instruction cell


@dataclass
class TRFData:
    text: dict = field(default_factory=dict)       # SEASON, COLOR, PROGRAM, STYLE, PO, GSM, FIBRE, CONSTRUCTION
    checked: set = field(default_factory=set)      # token keys to tick, e.g. {"PROTO_KNIT_FABRIC_1"}
    care_image: Image.Image | None = None


def has_tokens(template: bytes) -> bool:
    d = Document(io.BytesIO(template))
    return any("{{" in (t.text or "") for t in d.element.iter(qn("w:t")))


# ----------------------------------------------------------------- care image
def _inline_to_left_anchor(inline, v_offset_emu=31750):
    """Turn python-docx's inline picture into a floating one pinned to the LEFT edge of its cell
    (same technique as the original template, but with explicit align=left)."""
    anchor = parse_xml(
        f'<wp:anchor {nsdecls("wp", "a", "pic", "r")} distT="0" distB="0" distL="114300" '
        f'distR="114300" simplePos="0" relativeHeight="251671040" behindDoc="0" locked="0" '
        f'layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="column"><wp:align>left</wp:align></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{v_offset_emu}</wp:posOffset></wp:positionV>'
        f'</wp:anchor>')
    anchor.append(inline.find(qn("wp:extent")))
    anchor.append(parse_xml(f'<wp:effectExtent {nsdecls("wp")} l="0" t="0" r="0" b="0"/>'))
    anchor.append(parse_xml(f'<wp:wrapNone {nsdecls("wp")}/>'))
    anchor.append(inline.find(qn("wp:docPr")))
    cnv = inline.find(qn("wp:cNvGraphicFramePr"))
    if cnv is not None:
        anchor.append(cnv)
    anchor.append(inline.find(qn("a:graphic")))
    inline.getparent().replace(inline, anchor)


def insert_care_image(doc, paragraph_el, image: Image.Image):
    from docx.text.paragraph import Paragraph
    para = Paragraph(paragraph_el, doc._body)
    for r in own_runs(paragraph_el):                       # drop the {{CARE_IMAGE}} text run
        if r.find(qn("w:t")) is not None:
            r.getparent().remove(r)
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT               # strictly left-aligned
    buf = io.BytesIO()
    image.convert("RGB").save(buf, "PNG", optimize=True)
    buf.seek(0)
    width = Cm(CARE_WIDTH_CM)
    shape = para.add_run().add_picture(buf, width=width)
    height_cm = shape.height / 360000
    _inline_to_left_anchor(shape._inline)
    # a floating picture does not grow its cell -> reserve enough empty lines below it
    cell = paragraph_el.getparent()
    need = math.ceil((height_cm + 0.1) / LINE_HEIGHT_CM)
    paras = cell.findall(qn("w:p"))
    while len(paras) < need:
        clone = copy.deepcopy(paras[-1])
        for r in clone.findall(qn("w:r")):
            clone.remove(r)
        paras[-1].addnext(clone)
        paras = cell.findall(qn("w:p"))


# ----------------------------------------------------------------- main fill
def fill_template(template: bytes, data: TRFData) -> tuple[bytes, dict]:
    doc = Document(io.BytesIO(template))
    roots = [doc.element]
    for s in doc.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer,
                     s.even_page_header, s.even_page_footer):
            if part is not None and not part.is_linked_to_previous:
                roots.append(part._element)

    unresolved, care_p = [], []

    def resolver(tok):
        if tok == C.CARE_IMAGE_TOKEN:
            return None                                    # handled below
        if tok.startswith("CHK_"):
            return C.CHECKED if tok[4:] in data.checked else C.UNCHECKED
        if tok in data.text:
            return str(data.text[tok])
        unresolved.append(tok)
        return None

    for root in roots:
        for p in list(root.iter(qn("w:p"))):
            for run, tok in replace_tokens_in_paragraph(p, resolver):
                if tok.startswith("CHK_"):
                    set_symbol_font(run, C.SYMBOL_FONT)
            if "{{%s}}" % C.CARE_IMAGE_TOKEN in para_text(p):
                care_p.append(p)

    placed = False
    if care_p and data.care_image is not None:
        insert_care_image(doc, care_p[0], data.care_image)
        placed = True
    elif care_p:                                            # no image supplied: leave cell empty
        for r in own_runs(care_p[0]):
            care_p[0].remove(r)

    out = io.BytesIO()
    doc.save(out)
    result = out.getvalue()
    report = {"unresolved_tokens": sorted(set(unresolved)), "care_image_placed": placed,
              "logo_preserved": verify_logo(template, result)}
    return result, report


# ----------------------------------------------------------------- safety checks
def _media_hashes(z):
    return {hashlib.sha1(z.read(n)).hexdigest() for n in z.namelist() if n.startswith("word/media/")}


def _header_drawings(z):
    n = 0
    for name in z.namelist():
        if re.match(r"word/header\d*\.xml$", name):
            x = z.read(name).decode("utf8", "ignore")
            n += x.count("<w:drawing") + x.count("<w:pict")
    return n


def verify_logo(before: bytes, after: bytes) -> bool:
    """True when every header picture (Carhartt logo) and media file of the template survives."""
    with zipfile.ZipFile(io.BytesIO(before)) as a, zipfile.ZipFile(io.BytesIO(after)) as b:
        return (_header_drawings(a) == _header_drawings(b)
                and _media_hashes(a) <= _media_hashes(b))


def output_filename(style: str, program: str, color: str = "") -> str:
    name = C.FILENAME_PATTERN.format(style=style, program=program, color=color)
    name = name.replace("/", "").replace("_", " ")           # 'S/S' -> 'SS', never underscores
    name = re.sub(r'[\\:*?"<>|]', " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name + ".docx"
