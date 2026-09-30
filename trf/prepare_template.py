"""Turn the original TRF (legacy tick-boxes + sample text) into a TAGGED template.

  * Protocol / Stage / Gender / Garment-Wash tick-boxes -> {{CHK_<KEY>}}
  * Colour, Program, Style, PO, Fibre, Construction, GSM, Season -> {{COLOR}} ...
  * Old care-label picture in the "Care Instructions" cell -> {{CARE_IMAGE}}

Nothing else is touched: header (Carhartt logo), footers, text boxes, tables and all other
tick-boxes stay byte-for-byte as in your original file. No bookmarks are created.

Usage:  python -m trf.prepare_template "original TRF.docx" templates/TRF_TEMPLATE.docx
"""
from __future__ import annotations
import io
import re
import sys
from docx import Document
from docx.oxml.ns import qn
from . import config as C
from .docx_utils import (own_runs, para_text, run_text, make_token_run, set_run_text,
                         cell_text, find_cell, next_cell)


def _norm(s):
    return re.sub(r"[^a-z0-9]", "", s.casefold())


def _is_cb_begin(r):
    fc = r.find(qn("w:fldChar"))
    return (fc is not None and fc.get(qn("w:fldCharType")) == "begin"
            and fc.find(".//" + qn("w:checkBox")) is not None)


def _field_runs(runs, i):
    """Runs i..j that make up one legacy check-box field (begin ... end)."""
    out, depth = [], 0
    for r in runs[i:]:
        out.append(r)
        fc = r.find(qn("w:fldChar"))
        if fc is not None:
            t = fc.get(qn("w:fldCharType"))
            depth += 1 if t == "begin" else -1 if t == "end" else 0
            if depth == 0:
                break
    return out


def tag_checkboxes(paragraphs, patterns, side="after") -> list[str]:
    """Walk paragraphs in order, label each tick-box with its neighbouring text, tag matches.

    side="after"  : label is the text AFTER the box   ("[ ] Development")
    side="before" : label is the text BEFORE the box  ("yes [ ]  no [ ]")
    Labels never cross a paragraph boundary.
    """
    events = []                        # ('cb', [runs]) | ('t', str) | ('p',)
    for p in paragraphs:
        events.append(("p",))
        runs = own_runs(p)
        i = 0
        while i < len(runs):
            r = runs[i]
            if _is_cb_begin(r):
                fr = _field_runs(runs, i)
                events.append(("cb", fr)); i += len(fr); continue
            events.append(("t", run_text(r))); i += 1
    tagged = []
    for idx, ev in enumerate(events):
        if ev[0] != "cb":
            continue
        label = ""
        if side == "after":
            j = idx + 1
            while j < len(events) and events[j][0] == "t":
                label += events[j][1]; j += 1
        else:
            j = idx - 1
            while j >= 0 and events[j][0] == "t":
                label = events[j][1] + label; j -= 1
        key = _norm(label)
        for pat, token in patterns:
            if re.fullmatch(pat, key):
                fr = ev[1]
                new = make_token_run(fr[0], "CHK_" + token)
                fr[0].addprevious(new)
                for r in fr:
                    r.getparent().remove(r)
                tagged.append(token)
                break
    return tagged


def _paras_between(root, start_txt, end_txt):
    ps = list(root.iter(qn("w:p")))
    s = next(i for i, p in enumerate(ps) if para_text(p).strip() == start_txt)
    e = next(i for i, p in enumerate(ps) if i > s and end_txt in para_text(p))
    return ps[s:e + 1]


def _set_cell_token(tc, text):
    p = tc.find(qn("w:p"))
    runs = [r for r in own_runs(p) if r.find(qn("w:t")) is not None]
    if runs:
        set_run_text(runs[0], text)
        for r in runs[1:]:
            r.getparent().remove(r)
    else:
        p.append(make_token_run(None, text.strip("{}")))


def prepare(src: bytes) -> tuple[bytes, dict]:
    doc = Document(io.BytesIO(src))
    root = doc.element
    report = {"checkbox": [], "text": [], "care": False}

    # 1) Protocol + Stage
    scope = _paras_between(root, "PROTOCOL:", "Sealed Sample")
    report["checkbox"] += tag_checkboxes(scope, C.PROTOCOL_LABEL_PATTERNS + C.STAGE_LABEL_PATTERNS)

    # 2) Garment wash + gender (inside the sample-information table only)
    tc = find_cell(root, lambda t: "Garment Wash" in t)
    if tc is not None:
        report["checkbox"] += tag_checkboxes(tc.findall(qn("w:p")), C.WASH_LABEL_PATTERNS, side="before")
    tc = find_cell(root, lambda t: "Women" in t and "Infant" in t)
    if tc is not None:
        report["checkbox"] += tag_checkboxes(tc.findall(qn("w:p")), C.GENDER_LABEL_PATTERNS)

    # 3) Text fields: value cell = the cell to the right of the label cell
    for key, label in C.TEXT_FIELDS.items():
        lab = find_cell(root, lambda t, l=label: t.strip() == l)
        val = next_cell(lab) if lab is not None else None
        if val is None:
            continue
        _set_cell_token(val, "{{GSM}} GSM" if key == "GSM" else "{{%s}}" % key)
        report["text"].append(key)

    # 4) Care-instruction picture cell (first cell of the row below the heading)
    head = find_cell(root, lambda t: t.startswith("#Care Instructions"))
    if head is not None:
        row = head.getparent()
        below = row.getnext()
        if below is not None:
            cell = below.find(qn("w:tc"))
            p = cell.find(qn("w:p"))
            for r in own_runs(p):
                if r.find(qn("w:drawing")) is not None:      # old screenshot
                    r.getparent().remove(r)
            p.append(make_token_run(None, C.CARE_IMAGE_TOKEN))
            report["care"] = True

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue(), report


if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    data, rep = prepare(open(src, "rb").read())
    open(dst, "wb").write(data)
    print("tagged checkboxes:", len(rep["checkbox"]), "| text fields:", rep["text"], "| care:", rep["care"])
