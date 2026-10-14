import os
import base64
import datetime
import math
import re
import zipfile
from copy import deepcopy
from io import BytesIO

import streamlit as st
import pandas as pd
from docx import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run

st.set_page_config(page_title="SGV Group | Smart Letter Gen", layout="centered")

# ===================== CUSTOM CSS =====================
st.markdown("""
    <style>
        .stApp {
            background-color: #F9FAFB;
        }
        /* Custom Header Banner */
        .custom-banner {
            background-color: #264653;
            padding: 24px;
            border-radius: 12px;
            display: flex;
            align-items: center;
            gap: 20px;
            margin-bottom: 30px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }
        .logo-box {
            background-color: white;
            padding: 8px;
            border-radius: 10px;
            width: 80px;
            height: 80px;
            display: flex;
            justify-content: center;
            align-items: center;
            flex-shrink: 0;
        }    
        .banner-title {
            margin: 0 !important;
            font-size: 1.8rem !important;
            font-weight: 700 !important;
            color: white !important;
        }
        .banner-subtitle {
            margin: 5px 0 0 0;
            font-size: 0.95rem;
            color: #E2E8F0;
        }
        /* Section Titles */
        .section-title {
            color: #64748B;
            font-size: 0.9rem;
            font-weight: 700;
            letter-spacing: 1px;
            margin-top: 10px;
            margin-bottom: 10px;
            text-transform: uppercase;
        }
        /* Uploader Boxes */
        div[data-testid="stFileUploader"] {
            padding: 1.5rem;
            border-radius: 10px;
            background-color: #F1F5F9;
            border: 2px dashed #CBD5E1;
            transition: all 0.3s;
        }
        div[data-testid="stFileUploader"]:hover {
            border-color: #94A3B8;
            background-color: #E2E8F0;
        }
        /* Button */
        .stButton>button {
            width: 100%;
            border-radius: 8px;
            font-weight: bold;
            transition: all 0.3s;
        }
        .stButton>button:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 6px rgba(49, 130, 206, 0.2);
        }
    </style>
""", unsafe_allow_html=True)

# ===================== SIDEBAR =====================
st.sidebar.title("Excel Rules")

st.sidebar.markdown("""
### How it works
- Each Excel **column header** is a placeholder.
  Put `{Header}` anywhere in your DOCX template.
- Each Excel **row** becomes one letter.
- Works for any letter: offers, invoices, notices,
  certificates, invitations...

---

### Styling Suffixes
- `_b` → Bold  
- `_i` → Italic  
- `_u` → Underline  

You can combine them:
- `_bi`, `_bu`, `_iu`, `_biu`

---

### Number Formatting
- `_c` → Currency (₹)  
- `_comma` → Number with commas  
- `_2d` → 2 decimal places  

---

### Examples
- `Amount_c_b` → **₹5,000**
- `Score_2d_i` → *98.45*  
- `Users_comma_b` → **1,000,000**  

No suffix = plain text
""")

# ---------- SAMPLE EXCEL ----------
sample_df = pd.DataFrame({
    "RefNo": ["REF-001", "REF-002"],
    "Name_b": ["Neeraj Balodi", "Asha Rao"],
    "Address": ["12 MG Road, Delhi", "45 Park Street, Kolkata"],
    "Subject_u": ["Invoice Reminder", "Event Invitation"],
    "Amount_c_b": [5000, 12500.5],
    "Score_2d_i": [98.456, 87.1],
    "Units_comma": [1000000, 2500],
    "Date": [pd.Timestamp("2026-10-01"), pd.Timestamp("2026-10-15")],
})

sample_buffer = BytesIO()
sample_df.to_excel(sample_buffer, index=False, engine="openpyxl")
sample_buffer.seek(0)

st.sidebar.download_button(
    label="⬇️ Download Sample Excel",
    data=sample_buffer,
    file_name="SmartLetterGen_Sample.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary"
)

# ===================== MAIN =====================
# Helper to load logo as base64
logo_b64 = ""
if os.path.exists("logo.png"):
    with open("logo.png", "rb") as f:
        logo_b64 = base64.b64encode(f.read()).decode()

logo_html = f'<img src="data:image/png;base64,{logo_b64}" style="max-width: 100%; max-height: 100%;">' if logo_b64 else '<div style="color: #B91C1C; font-size: 22px; font-weight: 900; text-align: center; letter-spacing: 1px;">SGV</div>'

st.markdown(f"""
    <div class="custom-banner">
        <div class="logo-box">
            {logo_html}
        </div>
        <div>
            <h1 class="banner-title">Smart Letter Gen</h1>
            <p class="banner-subtitle">Upload your DOCX template and Excel data — personalized letters will be generated automatically.</p>
        </div>
    </div>
""", unsafe_allow_html=True)

st.markdown('<p class="section-title">📁 UPLOAD FILES</p>', unsafe_allow_html=True)

col_t, col_e = st.columns(2)
with col_t:
    template_file = st.file_uploader("DOCX Template", type=["docx"])
with col_e:
    excel_file = st.file_uploader("Excel Data (.xlsx)", type=["xlsx"])

# ===================== FORMATTERS =====================
FORMAT_TOKENS = {"c", "comma", "2d"}
STYLE_CHARS = set("biu")


def parse_key(key):
    """Read trailing suffix tokens (e.g. Salary_c_b) into a number format and a style.

    Only whole trailing tokens count, so names like "Emp_id" or "Dept_code"
    are not mistaken for italic / currency.
    """
    style = {"bold": False, "italic": False, "underline": False}
    fmt = None
    parts = key.split("_")

    while len(parts) > 1:
        token = parts[-1].lower()
        if token in FORMAT_TOKENS and fmt is None:
            fmt = token
        elif token and set(token) <= STYLE_CHARS and len(set(token)) == len(token):
            style["bold"] |= "b" in token
            style["italic"] |= "i" in token
            style["underline"] |= "u" in token
        else:
            break
        parts.pop()

    return fmt, style


def _to_number(value):
    if isinstance(value, bool):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def _plain(value):
    if isinstance(value, (pd.Timestamp, datetime.datetime)):
        if (value.hour, value.minute, value.second) == (0, 0, 0):
            return value.strftime("%d-%m-%Y")
        return value.strftime("%d-%m-%Y %H:%M")
    if isinstance(value, datetime.date):
        return value.strftime("%d-%m-%Y")
    if isinstance(value, float) or type(value).__name__ in ("float32", "float64"):
        if math.isfinite(value) and value == int(value):
            # Remove decimal .0 for whole numbers parsed as floats by pandas
            return str(int(value))
    return str(value)


def _grouped(number):
    if number == int(number):
        return f"{int(number):,}"
    return f"{number:,.2f}"


def format_value(key, value):
    if value is None or (isinstance(value, str) and value == ""):
        return ""

    fmt, _ = parse_key(key)
    number = _to_number(value)

    if number is not None:
        if fmt == "comma":
            return _grouped(number)
        if fmt == "c":
            return f"₹{_grouped(number)}"
        if fmt == "2d":
            return f"{number:.2f}"

    return _plain(value)

# ===================== PLACEHOLDER ENGINE =====================
PLACEHOLDER_RE = re.compile(r"\{\s*([^{}]+?)\s*\}")


def _replace_span(para, start, end, value, style):
    """Replace characters [start, end) of the paragraph's run text with value.

    The replacement inherits the formatting of the run where the placeholder
    begins (font, size, colour, highlight...). Suffix styles are only added on
    top; they never switch off formatting the template already has.
    """
    runs = para.runs
    bounds = []
    pos = 0
    for run in runs:
        length = len(run.text)
        bounds.append((pos, pos + length))
        pos += length

    si = next(i for i, (a, b) in enumerate(bounds) if a <= start < b)
    ei = next(i for i, (a, b) in enumerate(bounds) if a < end <= b)

    s_run, e_run = runs[si], runs[ei]
    prefix = s_run.text[: start - bounds[si][0]]
    suffix = e_run.text[end - bounds[ei][0]:]
    template_r = deepcopy(s_run._r)

    # Runs fully inside the placeholder (Word often splits "{Name}" into pieces)
    for run in runs[si + 1:ei]:
        if run.text:
            run._r.getparent().remove(run._r)

    if ei != si:
        if suffix:
            e_run.text = suffix
        else:
            e_run._r.getparent().remove(e_run._r)
        suffix = ""

    def make_run(text, extra_style=None):
        r = deepcopy(template_r)
        run = Run(r, para)
        run.text = text
        if extra_style:
            if extra_style["bold"]:
                run.bold = True
            if extra_style["italic"]:
                run.italic = True
            if extra_style["underline"]:
                run.underline = True
        return r

    new_elements = []
    if prefix:
        new_elements.append(make_run(prefix))
    new_elements.append(make_run(value, style))
    if suffix:
        new_elements.append(make_run(suffix))

    for el in new_elements:
        s_run._r.addprevious(el)
    s_run._r.getparent().remove(s_run._r)


def process_paragraph(para, data, missing):
    full_text = "".join(run.text for run in para.runs)
    if "{" not in full_text:
        return

    # Work backwards so earlier character offsets stay valid after each edit
    for match in reversed(list(PLACEHOLDER_RE.finditer(full_text))):
        key = match.group(1)
        if key not in data:
            missing.add(key)
            continue
        _, style = parse_key(key)
        value = format_value(key, data[key])
        _replace_span(para, match.start(), match.end(), value, style)


def _iter_paragraphs(doc):
    """Every paragraph in body, tables (incl. nested), text boxes, headers and footers."""
    yield from (Paragraph(p, doc._body) for p in doc.element.body.xpath(".//w:p"))

    seen = set()
    for section in doc.sections:
        for part in (
            section.header, section.first_page_header, section.even_page_header,
            section.footer, section.first_page_footer, section.even_page_footer,
        ):
            if part.is_linked_to_previous:
                continue
            el = part._element
            if id(el) in seen:
                continue
            seen.add(id(el))
            yield from (Paragraph(p, part) for p in el.xpath(".//w:p"))


def replace_placeholders(doc, data):
    missing = set()
    for para in list(_iter_paragraphs(doc)):
        process_paragraph(para, data, missing)
    return missing


def template_placeholders(doc):
    """All placeholder keys used in the template, in order of first appearance."""
    keys = []
    for para in _iter_paragraphs(doc):
        text = "".join(run.text for run in para.runs)
        for match in PLACEHOLDER_RE.finditer(text):
            if match.group(1) not in keys:
                keys.append(match.group(1))
    return keys


def safe_filename(text):
    text = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", text).strip(" .")
    return text or "letter"

# ===================== PROCESS =====================
if template_file and excel_file:
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"❌ Could not read the Excel file: {e}")
        st.stop()

    template_bytes = template_file.getvalue()
    try:
        template_keys = template_placeholders(Document(BytesIO(template_bytes)))
    except Exception as e:
        st.error(f"❌ Could not open the DOCX template: {e}")
        st.stop()

    # Clean headers so "Name " in Excel still matches {Name} in the template
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")

    # ---------- VALIDATION ----------
    if df.shape[1] < 1:
        st.error("❌ Excel has no columns.")
        st.stop()

    if df.empty:
        st.error("❌ Excel has headers but no data rows.")
        st.stop()

    if not template_keys:
        st.error("❌ No placeholders found in the template. Add `{ColumnName}` where data should go.")
        st.stop()

    matched = [k for k in template_keys if k in df.columns]
    missing = [k for k in template_keys if k not in df.columns]
    unused = [c for c in df.columns if c not in template_keys]

    if not matched:
        st.error(
            "❌ None of the template placeholders match an Excel column header. "
            "Placeholders found: " + ", ".join(f"`{{{k}}}`" for k in template_keys)
        )
        st.stop()

    st.success(
        f"✅ {len(df)} letter(s) ready — {len(matched)} of {len(template_keys)} "
        "template placeholder(s) matched to Excel columns"
    )
    if missing:
        st.warning(
            "⚠️ No Excel column for these placeholders (they will be left as-is): "
            + ", ".join(f"`{{{k}}}`" for k in missing)
        )
    if unused:
        st.caption("Excel columns not used in the template: " + ", ".join(unused))

    with st.expander("📊 Preview Excel data"):
        st.dataframe(df.head(20), hide_index=True)

    st.markdown('<p class="section-title">📝 FILE NAMES</p>', unsafe_allow_html=True)
    name_cols = st.multiselect(
        "Name each letter using these columns",
        options=list(df.columns),
        default=list(df.columns[:1]),
        help="Values are joined with '_'. If none are chosen (or values are blank), "
             "letters are numbered instead.",
    )
    name_prefix = st.text_input("Optional file name prefix", placeholder="e.g. Invitation")

    files_key = (
        template_file.name, template_file.size,
        excel_file.name, excel_file.size,
        tuple(name_cols), name_prefix,
    )
    if st.session_state.get("zip_key") != files_key:
        st.session_state.pop("zip_bytes", None)

    if st.button("🚀 Generate Letters (ZIP)", type="primary"):
        zip_buffer = BytesIO()
        used_names = set()
        width = len(str(len(df)))

        with st.spinner("Generating letters..."):
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for i, (_, row) in enumerate(df.iterrows(), start=1):
                    doc = Document(BytesIO(template_bytes))

                    row_data = {
                        col: "" if pd.isna(row[col]) else row[col]
                        for col in df.columns
                    }

                    replace_placeholders(doc, row_data)

                    doc_buffer = BytesIO()
                    doc.save(doc_buffer)

                    parts = [_plain(row_data[c]) for c in name_cols]
                    parts = [p for p in parts if p.strip()] or [f"Letter_{i:0{width}d}"]
                    if name_prefix.strip():
                        parts.insert(0, name_prefix.strip())
                    base = safe_filename("_".join(parts))
                    filename = f"{base}.docx"
                    n = 2
                    while filename.lower() in used_names:
                        filename = f"{base}_{n}.docx"
                        n += 1
                    used_names.add(filename.lower())

                    zip_file.writestr(filename, doc_buffer.getvalue())

        st.session_state.zip_key = files_key
        st.session_state.zip_bytes = zip_buffer.getvalue()

    if st.session_state.get("zip_bytes"):
        st.download_button(
            label="⬇️ Download Letters (ZIP)",
            data=st.session_state.zip_bytes,
            file_name="SmartLetterGen_Output.zip",
            mime="application/zip",
            type="primary",
            on_click="ignore",
        )
else:
    st.markdown("<br>", unsafe_allow_html=True)
    st.info("👆 Upload both your DOCX template and Excel data to get started.")
