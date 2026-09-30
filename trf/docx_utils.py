"""Low-level helpers for WordprocessingML (lxml elements from python-docx)."""
from __future__ import annotations
import re
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

TOKEN_RE = re.compile(r"\{\{([A-Z0-9_]+)\}\}")


def nearest_p(el):
    while el is not None and el.tag != qn("w:p"):
        el = el.getparent()
    return el


def own_runs(p):
    """Runs that belong to paragraph p itself (not to paragraphs nested in text boxes)."""
    return [r for r in p.iter(qn("w:r")) if nearest_p(r.getparent()) is p]


def run_text(r) -> str:
    return "".join(t.text or "" for t in r.findall(qn("w:t")))


def para_text(p) -> str:
    return "".join(run_text(r) for r in own_runs(p))


def set_run_text(r, text: str):
    ts = r.findall(qn("w:t"))
    if not ts:
        t = OxmlElement("w:t"); r.append(t); ts = [t]
    ts[0].text = text
    ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for extra in ts[1:]:
        r.remove(extra)


def make_token_run(rpr_source_run, token: str):
    """New run carrying {{TOKEN}}, copying character formatting from another run."""
    r = OxmlElement("w:r")
    if rpr_source_run is not None:
        rpr = rpr_source_run.find(qn("w:rPr"))
        if rpr is not None:
            import copy
            r.append(copy.deepcopy(rpr))
    t = OxmlElement("w:t")
    t.text = "{{%s}}" % token
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    return r


def set_symbol_font(r, font: str):
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr"); r.insert(0, rpr)
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
    for a in ("ascii", "hAnsi", "eastAsia", "cs"):
        rf.set(qn("w:" + a), font)
    for a in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        rf.attrib.pop(qn("w:" + a), None)


def replace_tokens_in_paragraph(p, resolver):
    """Replace every {{TOKEN}} in p, even when Word split it over several runs.

    resolver(token) -> replacement string, or None to leave the token untouched.
    Returns list of (run, token) for runs that received a replacement.
    """
    runs = [r for r in own_runs(p) if r.find(qn("w:t")) is not None]
    texts = [run_text(r) for r in runs]
    full = "".join(texts)
    if "{{" not in full:
        return []
    starts, pos = [], 0
    for t in texts:
        starts.append(pos); pos += len(t)
    hits = []
    for m in TOKEN_RE.finditer(full):
        rep = resolver(m.group(1))
        if rep is None:
            continue
        hits.append((m.start(), m.end(), rep, m.group(1)))
    if not hits:
        return []
    new = list(texts)
    done = []
    for s, e, rep, tok in reversed(hits):            # right-to-left keeps offsets valid
        first = None
        for i, (rs, t) in enumerate(zip(starts, texts)):
            re_ = rs + len(t)
            if re_ <= s or rs >= e:
                continue
            a, b = max(s, rs) - rs, min(e, re_) - rs
            if first is None:
                new[i] = new[i][:a] + rep + new[i][b:]
                first = i
            else:
                new[i] = new[i][:a] + new[i][b:]
        done.append((runs[first], tok))
    for r, t in zip(runs, new):
        if t != run_text(r):
            set_run_text(r, t)
    return done


def cell_text(tc) -> str:
    return " ".join(para_text(p) for p in tc.findall(qn("w:p"))).strip()


def find_cell(root, predicate):
    for tc in root.iter(qn("w:tc")):
        if predicate(cell_text(tc)):
            return tc
    return None


def next_cell(tc):
    sib = tc.getnext()
    while sib is not None and sib.tag != qn("w:tc"):
        sib = sib.getnext()
    return sib
