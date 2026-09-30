"""Read the order sheet, de-duplicate colours, pull PO / GSM / fibre from a row."""
from __future__ import annotations
import io
import re
from dataclasses import dataclass, field
import pandas as pd
from .config import COLUMN_ALIASES


def _clean(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"[ \t\r\f\v]+", " ", str(v)).strip()


def norm(v) -> str:
    """Case/space/punctuation-insensitive key used for matching & de-duplication."""
    return re.sub(r"[^a-z0-9]+", "", _clean(v).casefold())


def _find_col(df: pd.DataFrame, key: str):
    wanted = [a.upper() for a in COLUMN_ALIASES[key]]
    for c in df.columns:
        if str(c).strip().upper() in wanted:
            return c
    return None


@dataclass
class OrderSheet:
    df: pd.DataFrame
    cols: dict = field(default_factory=dict)   # logical name -> real column name

    def col(self, key):
        return self.cols.get(key)

    def programs(self) -> list[str]:
        return unique_ordered(self.df[self.cols["program"]])

    def styles(self, program: str) -> list[str]:
        d = self.df[self.df[self.cols["program"]].map(norm) == norm(program)]
        return unique_ordered(d[self.cols["style"]])

    def rows(self, program: str, style: str) -> pd.DataFrame:
        d = self.df
        d = d[d[self.cols["program"]].map(norm) == norm(program)]
        return d[d[self.cols["style"]].map(norm) == norm(style)]

    def colors(self, program: str, style: str) -> list[str]:
        """Unique colours (each appears once, first spelling wins, whitespace/case ignored)."""
        return unique_ordered(self.rows(program, style)[self.cols["color"]])

    def row_for_color(self, program: str, style: str, color: str) -> pd.Series:
        r = self.rows(program, style)
        return r[r[self.cols["color"]].map(norm) == norm(color)].iloc[0]

    def po_columns(self) -> list[str]:
        wanted = [a.upper() for a in COLUMN_ALIASES["po_candidates"]]
        return [c for c in self.df.columns if str(c).strip().upper() in wanted]


def unique_ordered(series) -> list[str]:
    seen, out = set(), []
    for v in series:
        s = _clean(v)
        k = norm(s)
        if s and k not in seen:
            seen.add(k)
            out.append(s)
    return out


def load_orders(data: bytes, sheet_name=0) -> OrderSheet:
    df = pd.read_excel(io.BytesIO(data), sheet_name=sheet_name, dtype=object)
    df = df.dropna(how="all")
    cols = {k: _find_col(df, k) for k in COLUMN_ALIASES if k != "po_candidates"}
    missing = [k for k in ("program", "style", "color", "fabrication") if not cols.get(k)]
    if missing:
        raise ValueError(f"Excel is missing required column(s): {', '.join(missing)}. "
                         f"Found: {', '.join(map(str, df.columns))}")
    return OrderSheet(df=df.reset_index(drop=True), cols=cols)


# ------------------------------------------------------------------ parsing
def extract_gsm(fabrication: str) -> tuple[str, str]:
    """Return (gsm, note). Takes ONLY the number at the very end of the string.

    '60% Cotton 40% Polyester S/J 180 GSM'          -> 180
    '100% Cotton Combed Single Jersey, 235 g/m²'    -> 235
    '... 220 gsm(after OD 230 GSM)'                 -> 230
    '... Fleece 390 (AW 420)'                       -> 420
    """
    s = _clean(fabrication)
    if not s:
        return "", "Fabrication is empty in Excel."
    m = re.search(r"(\d+(?:\.\d+)?)\D*$", s)
    if not m or not (50 <= float(m.group(1)) <= 1000):
        return "", "No plausible GSM number found at the end of Fabrication - please type it."
    val = m.group(1)
    val = val[:-2] if val.endswith(".0") else val
    note = ""
    if "\n" in str(fabrication).strip():
        note = "Fabrication has several lines (e.g. body + rib); the number at the very end was used."
    return val, note


_CONSTRUCTIONS = [
    (r"\bs\.?\s?/?\s?j\b|single\s+jersey", "Single Jersey"),
    (r"slub\s+jersey", "Slub Jersey"),
    (r"waffle", "Waffle"), (r"pique|piqu\u00e9", "Pique"),
    (r"loop\s*back\s*terry|terry", "Terry"), (r"fleece", "Fleece"),
    (r"mesh", "Mesh"), (r"\brib\b", "Rib"),
]


def parse_fabrication(fabrication: str) -> dict:
    """Best-effort fibre + construction guess. Always shown to the user as editable."""
    s = _clean(fabrication)
    gsm, note = extract_gsm(fabrication)
    construction, cut = "", len(s)
    for pat, name in _CONSTRUCTIONS:
        m = re.search(pat, s, re.I)
        if m:
            construction, cut = name, m.start()
            break
    fibre = s[:cut].strip(" ,;-\n") or s.split("\n")[0]
    fibre = re.sub(r"^(body|main)\s*-\s*", "", fibre, flags=re.I)
    return {"gsm": gsm, "gsm_note": note, "fibre": fibre, "construction": construction}


def guess_gender(style: str, explicit: str = "") -> str:
    e = _clean(explicit).casefold()
    for g in ("women", "men", "children", "infant"):
        if e.startswith(g[:3]) and e:
            return g.capitalize()
    return "Women" if re.match(r"^\s*w['\u2019`]", style or "", re.I) else "Men"


def guess_wash(wash_type: str) -> str:
    return "No" if re.search(r"non\s*wash", wash_type or "", re.I) else "Yes"
