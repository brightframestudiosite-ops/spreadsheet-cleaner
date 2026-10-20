"""Spreadsheet Cleaner - a small Streamlit tool.

Upload a CSV or Excel file, choose which columns are names, emails, phones,
dates and numbers, and download a cleaned Excel report with four sheets:
Summary, Clean data, Needs review, Duplicates removed.

Run it with:   streamlit run app.py
Install with:  pip install streamlit pandas openpyxl
"""

import io

import pandas as pd
import streamlit as st
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

EMAIL_PATTERN = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"   # text @ text . text, no spaces
NONE = "(none)"                                  # label for "this column doesn't exist"


# ---------------------------------------------------------------------------
# Cleaning logic (no Streamlit code in here, so it is easy to test)
# ---------------------------------------------------------------------------

def load_file(uploaded):
    """Read an uploaded CSV or Excel file and tidy the column names."""
    if uploaded.name.lower().endswith(".csv"):
        try:
            df = pd.read_csv(uploaded)
        except UnicodeDecodeError:               # some CSVs are not UTF-8
            uploaded.seek(0)
            df = pd.read_csv(uploaded, encoding="latin-1")
    else:
        df = pd.read_excel(uploaded)
    df.columns = df.columns.astype(str).str.strip()
    return df


def clean_data(raw, cfg):
    """Clean the table. Returns a dict with clean, review, removed, summary, excel."""
    df = raw.copy()                              # never change the original
    problems = pd.Series("", index=df.index)     # one empty note per row

    # 1. Required columns: blank values are flagged (checked on the ORIGINAL data)
    for col in cfg["required_cols"]:
        blank = raw[col].isna() | (raw[col].astype("string").str.strip() == "").fillna(False)
        problems.loc[blank] += f"missing {col}; "

    # 2. Names: trim, collapse spaces, capitalise
    for col in cfg["name_cols"]:
        df[col] = (
            df[col].astype("string").str.strip()
            .str.replace(r"\s+", " ", regex=True)
            .str.title()
            .replace("", pd.NA)
        )

    # 3. Email: trim, lowercase, check the shape (blank emails are left to "required")
    if cfg["email_col"]:
        col = cfg["email_col"]
        df[col] = df[col].astype("string").str.strip().str.lower()
        bad = df[col].notna() & ~df[col].str.match(EMAIL_PATTERN, na=False)
        problems.loc[bad] += "invalid email; "

    # 4. Phone: keep digits, add country code, validate length
    if cfg["phone_col"]:
        col = cfg["phone_col"]
        raw_phone = df[col].astype("string").str.strip()
        digits = raw_phone.str.replace(r"\D", "", regex=True)
        negative = digits.str.startswith("-", na=False)
        if cfg["country_code"]:
            not_local = (digits.str.len() != cfg["local_length"]).fillna(True)
            digits = digits.where(not_local, cfg["country_code"] + digits)
        valid = digits.str.len().between(cfg["min_digits"], cfg["max_digits"]).fillna(False)
        valid = valid.astype(bool)
        if cfg["negative_invalid"]:
            valid = valid & ~negative
        df[f"{col}_Clean"] = ("+" + digits).where(valid, pd.NA)
        df[f"{col}_Status"] = valid.map({True: "valid", False: "invalid"})
        problems.loc[~valid & raw_phone.notna()] += "invalid phone; "

    # 5. Dates: convert to real dates, flag the ones that cannot be read
    for col in cfg["date_cols"]:
        parsed = pd.to_datetime(
            df[col], format="mixed", dayfirst=cfg["dayfirst"], errors="coerce"
        )
        problems.loc[parsed.isna() & df[col].notna()] += f"invalid date ({col}); "
        df[col] = parsed

    # 6. Numbers: strip currency symbols and commas, convert, flag problems
    for col in cfg["number_cols"]:
        original = df[col]
        text = original.astype("string").str.replace(r"[€$£,\s]", "", regex=True)
        parsed = pd.to_numeric(text, errors="coerce")
        problems.loc[parsed.isna() & original.notna()] += f"invalid number ({col}); "
        problems.loc[parsed <= 0] += f"value out of range ({col}); "
        df[col] = parsed

    # 7. Duplicates (on the cleaned data). Empty key = whole-row duplicates.
    is_dup = df.duplicated(subset=cfg["dup_cols"] or None, keep="first")

    # 8. Split into three groups
    df["Problems"] = problems.str.strip("; ")
    clean = df[(df["Problems"] == "") & ~is_dup]
    review = df[(df["Problems"] != "") & ~is_dup]
    removed = df[is_dup]

    # 9. Summary sheet (includes a count for every kind of problem)
    breakdown = df.loc[~is_dup, "Problems"].str.split("; ").explode()
    breakdown = breakdown[breakdown != ""].value_counts()
    rows = [
        ("Rows in", len(raw)),
        ("Clean rows", len(clean)),
        ("Rows needing review", len(review)),
        ("Duplicates removed", len(removed)),
    ]
    rows += [(f"Problem: {name}", int(count)) for name, count in breakdown.items()]
    summary = pd.DataFrame(rows, columns=["Metric", "Count"])

    excel = to_excel_bytes({
        "Summary": summary,
        "Clean data": clean,
        "Needs review": review,
        "Duplicates removed": removed,
    })
    return {"clean": clean, "review": review, "removed": removed,
            "summary": summary, "excel": excel}


def to_excel_bytes(sheets):
    """Write the sheets to an Excel file in memory, with tidy formatting."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl", datetime_format="DD/MM/YYYY") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False)
            sheet = writer.sheets[name]
            sheet.freeze_panes = "A2"                       # keep the header visible
            for cell in sheet[1]:
                cell.font = Font(bold=True)                 # bold header row
            for i, col in enumerate(frame.columns, start=1):
                longest = frame[col].astype(str).str.len().max() if len(frame) else 0
                width = min(max(len(str(col)), longest) + 2, 40)
                sheet.column_dimensions[get_column_letter(i)].width = width
    return buffer.getvalue()


def guess(columns, keywords):
    """Return the first column whose name contains one of the keywords."""
    for c in columns:
        if any(k in c.lower() for k in keywords):
            return c
    return None


# ---------------------------------------------------------------------------
# Page (everything below is the Streamlit interface)
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Spreadsheet Cleaner", page_icon="🧹", layout="wide")
st.title("🧹 Spreadsheet Cleaner")
st.write(
    "Upload a messy CSV or Excel file. The tool standardises names, emails, phones, "
    "dates and numbers, flags problem rows, and never deletes anything silently."
)

uploaded = st.file_uploader("Upload a CSV or Excel file", type=["csv", "xlsx"])
if uploaded is None:
    st.info("Upload a file to begin. Everything is processed in memory and not stored.")
    st.stop()

raw = load_file(uploaded)
cols = list(raw.columns)

st.subheader("1. Preview")
st.write(f"{len(raw):,} rows × {len(cols)} columns")
st.dataframe(raw.head(20), use_container_width=True)

# ---- Settings (sidebar) ----
st.sidebar.header("2. Tell the tool about your columns")


def pick(label, keywords):
    options = [NONE] + cols
    found = guess(cols, keywords)
    choice = st.sidebar.selectbox(label, options, index=options.index(found) if found else 0)
    return None if choice == NONE else choice


name_default = [c for c in cols if "name" in c.lower()]
date_default = [c for c in cols if "date" in c.lower()]
number_words = ["salary", "price", "amount", "total", "qty", "quantity"]
number_default = [c for c in cols
                  if c.lower() == "age" or any(w in c.lower() for w in number_words)]
id_default = [c for c in cols if c.lower() == "id" or c.lower().endswith("_id")][:1]

name_cols = st.sidebar.multiselect("Name columns", cols, default=name_default)
email_col = pick("Email column", ["email", "e-mail"])
phone_col = pick("Phone column", ["phone", "mobile", "tel"])
date_cols = st.sidebar.multiselect("Date columns", cols, default=date_default)
number_cols = st.sidebar.multiselect("Number columns (salary, age, price...)",
                                     cols, default=number_default)

st.sidebar.header("3. Rules")
date_order = st.sidebar.radio(
    "Date order",
    ["Day first (4/2/2021 = 4 February)", "Month first (4/2/2021 = April 2)"],
)
st.sidebar.caption("Look at the preview first. If you see a date like 27/11/2021, "
                   "or 11/27/2021, that tells you which order the file uses.")

required_default = [c for c in name_cols + [email_col] + date_cols if c]
required_default += [c for c in number_cols if "salary" in c.lower() or "price" in c.lower()]
required_cols = st.sidebar.multiselect(
    "Required columns (blank = flagged for review)", cols,
    default=list(dict.fromkeys(required_default)),
)
dup_cols = st.sidebar.multiselect(
    "Duplicate key (rows sharing these values are duplicates)", cols, default=id_default,
)
st.sidebar.caption("Leave empty to treat only fully identical rows as duplicates. "
                   "Pick the key carefully: an email column can repeat for different people.")

if phone_col:
    with st.sidebar.expander("Phone settings"):
        country_code = st.text_input("Country code to add (blank = none)", "356")
        local_length = st.number_input("Local number length", 1, 15, 8)
        min_digits = st.number_input("Minimum digits", 1, 15, 8)
        max_digits = st.number_input("Maximum digits", 1, 20, 15)
        negative_invalid = st.checkbox("Treat negative numbers as invalid", value=True)
else:
    country_code, local_length, min_digits, max_digits, negative_invalid = "", 8, 8, 15, True

# ---- Run ----
st.subheader("2. Clean")
if st.button("Clean my file", type="primary"):
    cfg = {
        "name_cols": name_cols, "email_col": email_col, "phone_col": phone_col,
        "date_cols": date_cols, "number_cols": number_cols,
        "dayfirst": date_order.startswith("Day"),
        "required_cols": required_cols, "dup_cols": dup_cols,
        "country_code": country_code.strip(), "local_length": int(local_length),
        "min_digits": int(min_digits), "max_digits": int(max_digits),
        "negative_invalid": negative_invalid,
    }
    try:
        with st.spinner("Cleaning..."):
            st.session_state["result"] = clean_data(raw, cfg)
            st.session_state["result_file"] = uploaded.name
    except Exception as error:
        st.error(f"Something went wrong: {error}")

result = st.session_state.get("result")
if result and st.session_state.get("result_file") == uploaded.name:
    st.subheader("3. Results")
    st.caption("Changed a setting? Click 'Clean my file' again to refresh.")
    a, b, c, d = st.columns(4)
    a.metric("Rows in", f"{len(raw):,}")
    b.metric("Clean", f"{len(result['clean']):,}")
    c.metric("Needs review", f"{len(result['review']):,}")
    d.metric("Duplicates removed", f"{len(result['removed']):,}")

    st.download_button(
        "⬇️ Download cleaned Excel report",
        data=result["excel"],
        file_name="cleaned_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    tab1, tab2, tab3, tab4 = st.tabs(["Summary", "Needs review", "Clean data", "Duplicates removed"])
    tab1.dataframe(result["summary"], use_container_width=True, hide_index=True)
    tab2.dataframe(result["review"], use_container_width=True)
    tab3.dataframe(result["clean"], use_container_width=True)
    tab4.dataframe(result["removed"], use_container_width=True)
