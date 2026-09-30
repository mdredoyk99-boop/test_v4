"""TRF Generator - Streamlit UI.   Run:  streamlit run app.py"""
import io
from pathlib import Path

import pandas as pd
import streamlit as st

from trf import config as C
from trf.docx_filler import TRFData, fill_template, has_tokens, output_filename
from trf.excel_reader import (guess_gender, guess_wash, load_orders, norm, parse_fabrication)
from trf.prepare_template import prepare
from trf.techpack import read_techpack, recrop

st.set_page_config(page_title="TRF Generator", page_icon="🧵", layout="wide")
DEFAULT_TEMPLATE = Path(__file__).parent / "templates" / "TRF_TEMPLATE.docx"


@st.cache_data(show_spinner=False)
def _orders(data: bytes):
    return load_orders(data)


@st.cache_resource(show_spinner="Reading techpack...")
def _techpack(data: bytes, name: str):
    return read_techpack(data, name)


def _status(ok):
    return "✅ match" if ok else "⚠️ check"


st.title("🧵 Test Request Form (TRF) Generator")

# ------------------------------------------------------------------ 1. uploads
c1, c2 = st.columns(2)
xl_file = c1.file_uploader("1 · Order Excel (PO, Style, Color, Program, Fabrication…)", type=["xlsx", "xls"])
tp_file = c2.file_uploader("2 · Techpack (PDF, or a screenshot of the care label)", type=["pdf", "png", "jpg", "jpeg"])
with st.sidebar:
    st.header("Template")
    tpl_up = st.file_uploader("Use a different TRF template (.docx)", type=["docx"])
    st.caption("Default: templates/TRF_TEMPLATE.docx (tagged version of your demo TRF). "
               "An untagged original is tagged automatically.")

if not (xl_file and tp_file):
    st.info("Upload the Excel sheet and the Techpack to begin.")
    st.stop()

try:
    sheet = _orders(xl_file.getvalue())
except Exception as e:
    st.error(f"Could not read the Excel file: {e}"); st.stop()
tp = _techpack(tp_file.getvalue(), tp_file.name)

# ------------------------------------------------------------------ 2. pick the garment
st.subheader("3 · Select program, style & colour")
programs = sheet.programs()
p_default = next((i for i, p in enumerate(programs) if norm(p) == norm(tp.program)), 0)
s1, s2, s3 = st.columns(3)
program = s1.selectbox("Program", programs, index=p_default)
styles = sheet.styles(program)
s_default = next((i for i, s in enumerate(styles) if norm(s) == norm(tp.style)), 0)
style = s2.selectbox("Style", styles, index=s_default)
colors = sheet.colors(program, style)            # duplicates already removed
color = s3.selectbox("Colour (from Excel, unique)", colors)
row = sheet.row_for_color(program, style, color)

fab = str(row[sheet.col("fabrication")] or "")
parsed = parse_fabrication(fab)
po_cols = sheet.po_columns()

# ------------------------------------------------------------------ 3. verification view
st.subheader("4 · Verify Excel data against the Techpack")
tp_norm = norm(tp.text)
wash_txt = str(row[sheet.col("wash")]) if sheet.col("wash") else ""
item_no = str(row[sheet.col("item_no")]).strip() if sheet.col("item_no") else ""
rows = [
    ("Program", program, tp.program, norm(program) == norm(tp.program)),
    ("Style", style, tp.style, norm(style) == norm(tp.style)),
    ("Item No", item_no, tp.item_no, bool(item_no) and norm(item_no) == norm(tp.item_no)),
    ("Colour", color, "found in techpack" if norm(color) in tp_norm else "not found", norm(color) in tp_norm),
    ("Fibre / composition", parsed["fibre"], tp.composition,
     bool(parsed["fibre"]) and norm(parsed["fibre"]) in norm(tp.composition)),
    ("Construction", parsed["construction"], tp.composition,
     bool(parsed["construction"]) and norm(parsed["construction"]) in norm(tp.composition)),
    ("Wash type (Excel)", wash_txt, " / ".join(tp.care_text[:1]) or "-", not tp.care_text and False),
]
if not tp.is_pdf:                                   # image techpack: only the care label is available
    rows = [r for r in rows if r[0] in ("Wash type (Excel)",)]
df = pd.DataFrame([(a, b, c, _status(d) if a != "Wash type (Excel)" else "👁 compare") for a, b, c, d in rows],
                  columns=["Field", "Excel", "Techpack", "Status"])
st.dataframe(df, hide_index=True, use_container_width=True)
if any("⚠️" in s for s in df["Status"]):
    st.warning("Some fields differ between Excel and the Techpack (e.g. techpack composition may still say "
               "'In Development'). Confirm the Excel values are the ones you want on the TRF.")

v1, v2 = st.columns([3, 2])
with v1:
    st.markdown("**Care instructions (from Techpack) - this image goes into the TRF, left-aligned**")
    care_img = tp.care_image
    if tp.is_pdf and tp.crop_pts:
        with st.expander("Adjust the crop (PDF points)"):
            l, t, r, b = tp.crop_pts
            W, H = tp.page_size_pts
            cl, ct, cr, cb = st.columns(4)
            l = cl.number_input("Left", 0.0, W, float(l), 1.0)
            t = ct.number_input("Top", 0.0, H, float(t), 1.0)
            r = cr.number_input("Right", 0.0, W, float(r), 1.0)
            b = cb.number_input("Bottom", 0.0, H, float(b), 1.0)
            care_img = recrop(tp, (l, t, r, b))
    elif tp.is_pdf:
        st.warning("Care band not detected automatically - showing the whole first page.")
        care_img = tp.page_image
    st.image(care_img, use_container_width=True)
with v2:
    st.markdown("**Text read from the care band**")
    st.code("\n".join(tp.care_text) or "(image techpack - no text)", language=None)
    st.markdown("**Excel wash type:** " + (wash_txt or "-"))

# ------------------------------------------------------------------ 4. dropdowns & editable fields
st.subheader("5 · Test package")
d1, d2 = st.columns(2)
protocol = d1.selectbox("PROTOCOL", list(C.PROTOCOLS), index=None, placeholder="Select protocol…")
stage = d2.selectbox("STAGE", list(C.STAGES), index=None, placeholder="Select stage…")

st.subheader("6 · Review the values that will be written")
e1, e2, e3, e4 = st.columns(4)
season = e1.text_input("Season", tp.season or "")
po_col = e2.selectbox("PO column", po_cols, index=0) if po_cols else None
po = e3.text_input("PO #", str(row[po_col]).strip() if po_col else "")
gsm = e4.text_input("GSM (number at end of Fabrication)", parsed["gsm"])
if parsed["gsm_note"]:
    st.caption("ℹ️ " + parsed["gsm_note"])
f1, f2, f3, f4 = st.columns(4)
fibre = f1.text_input("Fibre content", parsed["fibre"])
construction = f2.text_input("Construction", parsed["construction"])
gender_default = guess_gender(style, str(row[sheet.col("gender")]) if sheet.col("gender") else "")
gender = f3.selectbox("Gender", list(C.GENDERS), index=list(C.GENDERS).index(gender_default))
wash = f4.selectbox("Garment wash", list(C.WASH), index=list(C.WASH).index(guess_wash(wash_txt)))
st.caption(f"Excel Fabrication: {fab!r}")

# ------------------------------------------------------------------ 5. generate
st.subheader("7 · Generate")
ok = st.checkbox("I have verified the Excel data against the Techpack care instructions")
missing = [n for n, v in (("Protocol", protocol), ("Stage", stage), ("GSM", gsm), ("PO", po)) if not v]
if missing:
    st.caption("Still needed: " + ", ".join(missing))

if st.button("Generate TRF", type="primary", disabled=not (ok and not missing)):
    tpl = tpl_up.getvalue() if tpl_up else DEFAULT_TEMPLATE.read_bytes()
    if not has_tokens(tpl):
        tpl, rep = prepare(tpl)
        st.info(f"Template had no {{{{tags}}}} - tagged automatically ({len(rep['checkbox'])} tick-boxes).")
    data = TRFData(
        text={"SEASON": season, "COLOR": color, "PROGRAM": program, "STYLE": style, "PO": po,
              "GSM": gsm, "FIBRE": fibre, "CONSTRUCTION": construction},
        checked={C.PROTOCOLS[protocol], C.STAGES[stage], C.GENDERS[gender], C.WASH[wash]},
        care_image=care_img)
    out, rep = fill_template(tpl, data)
    name = output_filename(style, program, color)
    st.success(f"Created: {name}")
    if not rep["logo_preserved"]:
        st.error("Carhartt logo check FAILED - do not send this file.")
    if rep["unresolved_tokens"]:
        st.warning("Placeholders with no value: " + ", ".join(rep["unresolved_tokens"]))
    st.download_button("⬇️ Download TRF (.docx)", out, file_name=name,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
