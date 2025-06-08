import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

# ───── Streamlit page setup ───────────────────────────────────────────────────
st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# ───── Download Sample Template ───────────────────────────────────────────────
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ───── Helpers ─────────────────────────────────────────────────────────────────

def nationality_group(row):
    nat = str(row.get("Nationality (Country Name)", "")).strip().lower()
    pr  = str(row.get("PR", "")).strip().lower()
    if nat == "singapore":
        return 1
    elif pr in ("yes","y","pr"):
        return 2
    elif nat == "malaysia":
        return 3
    elif nat == "india":
        return 4
    else:
        return 5

def split_name(full_name):
    s = str(full_name).strip()
    if " " in s:
        i = s.find(" ")
        return pd.Series([s[:i], s[i+1:]])
    return pd.Series([s, ""])

def clean_gender(g):
    v = str(g).strip().upper()
    if v == "M": return "Male"
    if v == "F": return "Female"
    if v in ("MALE","FEMALE"): return v.title()
    return v

# ───── Core Cleaning Logic ────────────────────────────────────────────────────

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    # — Step 0: drop any “Unnamed” junk columns —
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    # — Step 1: ensure exactly these 13 by name (pad missing with blanks):
    EXPECTED = [
        "S/N",
        "Vehicle Plate Number",
        "Company Full Name",
        "Full Name As Per NRIC",
        "First Name as per NRIC",
        "Middle and Last Name as per NRIC",
        "Identification Type",
        "IC (Last 3 digits and suffix) 123A",
        "Work Permit Expiry Date",
        "Nationality (Country Name)",
        "PR",
        "Gender",
        "Mobile Number",
    ]
    for col in EXPECTED:
        if col not in df.columns:
            df[col] = ""
    # now pick only those, in order
    df = df[EXPECTED].copy()

    # — Step 2: drop fully blank rows in cols D–M
    df = df.dropna(subset=EXPECTED[3:13], how="all")

    # — Step 3: normalize nationality (including Indian→India)
    nat_map = {
        "chinese":    "China",
        "singaporean":"Singapore",
        "malaysian":  "Malaysia",
        "indian":     "India",
    }
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .astype(str)
          .str.strip()
          .str.lower()
          .replace(nat_map, regex=False)
          .str.title()
    )

    # — Step 4: sort by company → nationality-group → country → name
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df = (
        df.sort_values(
            by=[
                "Company Full Name",
                "SortGroup",
                "Nationality (Country Name)",
                "Full Name As Per NRIC",
            ],
            ignore_index=True,
        )
        .drop(columns="SortGroup")
    )

    # — Step 5: reassign S/N 1..N
    df["S/N"] = range(1, len(df) + 1)

    # — Step 6: standardize vehicle plates
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
          .str.replace(r"[\/,]", ";", regex=True)
          .str.replace(r"\s*;\s*", ";", regex=True)
          .str.strip()
          .replace("nan", "", regex=False)
    )

    # — Step 7: proper-case + split name
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = (
        df["Full Name As Per NRIC"].apply(split_name)
    )

    # — Step 8: swap IC vs WP if reversed (detect “-” in IC column)
    iccol, wpcol = EXPECTED[7], EXPECTED[8]
    if df[iccol].astype(str).str.contains("-", na=False).any():
        df[[iccol, wpcol]] = df[[wpcol, iccol]]

    # — Step 9: trim IC suffix
    df[iccol] = df[iccol].astype(str).str[-4:]

    # — Step 10: mobile → digits only
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(r"\D", "", regex=True)

    # — Step 11: normalize gender
    df["Gender"] = df["Gender"].apply(clean_gender)

    # — Step 12: format WP expiry date YYYY-MM-DD
    df[wpcol] = pd.to_datetime(df[wpcol], errors="coerce").dt.strftime("%Y-%m-%d")

    return df

# ───── Build single-sheet Excel ───────────────────────────────────────────────

def generate_visitor_only(df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        ws = writer.sheets["Visitor List"]

        # styling objects
        header_fill  = PatternFill("solid", fgColor="94B455")
        warning_fill = PatternFill("solid", fgColor="FFCCCC")
        border       = Border(Side("thin"),Side("thin"),Side("thin"),Side("thin"))
        center       = Alignment("center","center")
        normal_font  = Font("Calibri",11)
        bold_font    = Font("Calibri",11,bold=True)

        # 1) all-cell border/alignment/font
        for row in ws.iter_rows():
            for cell in row:
                cell.border    = border
                cell.alignment = center
                cell.font      = normal_font

        # 2) header row styling
        for c in range(1, ws.max_column+1):
            h = ws[f"{get_column_letter(c)}1"]
            h.fill = header_fill
            h.font = bold_font

        # 3) freeze top row
        ws.freeze_panes = ws["A2"]

        # 4) highlight ID vs PR/Nat errors
        mismatches = 0
        for r in range(2, ws.max_row+1):
            idt = str(ws[f"G{r}"].value).strip().upper()
            nat = str(ws[f"J{r}"].value).strip().title()
            pr  = str(ws[f"K{r}"].value).strip().lower()
            bad = False
            # NRIC must be SG or PR
            if idt=="NRIC" and not (nat=="Singapore" or (nat!="Singapore" and pr in ("yes","pr"))):
                bad = True
            # FIN must not be PR or SG
            if idt=="FIN" and (pr in ("yes","pr") or nat=="Singapore"):
                bad = True
            if bad:
                for col in ("G","J","K"):
                    ws[f"{col}{r}"].fill = warning_fill
                mismatches += 1

        if mismatches:
            st.warning(f"⚠️ {mismatches} potential mismatch(es) found.")

        # 5) autosize columns & row height
        for col in ws.columns:
            w = max(len(str(cell.value)) for cell in col if cell.value)
            ws.column_dimensions[get_column_letter(col[0].column)].width = w + 2
        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 20

        # 6) vehicles summary
        plates = []
        for v in df["Vehicle Plate Number"].dropna():
            plates += [p.strip() for p in str(v).split(";") if p.strip()]
        ir = ws.max_row + 2
        if plates:
            ws[f"B{ir}"].value     = "Vehicles"
            ws[f"B{ir}"].border    = border
            ws[f"B{ir}"].alignment = center
            ws[f"B{ir+1}"].value   = ";".join(sorted(set(plates)))
            ws[f"B{ir+1}"].border  = border
            ws[f"B{ir+1}"].alignment = center
            ir += 2

        # 7) total visitors
        ws[f"B{ir}"].value     = "Total Visitors"
        ws[f"B{ir}"].border    = border
        ws[f"B{ir}"].alignment = center
        ws[f"B{ir+1}"].value   = df["Company Full Name"].notna().sum()
        ws[f"B{ir+1}"].border  = border
        ws[f"B{ir+1}"].alignment = center

    buf.seek(0)
    return buf

# ───── Streamlit UI ────────────────────────────────────────────────────────────
uploaded = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])
if uploaded:
    raw_df  = pd.read_excel(uploaded, sheet_name="Visitor List")
    cleaned = clean_data(raw_df)
    out_buf = generate_visitor_only(cleaned)
    fname   = f"Cleaned_VisitorList_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        label="📥 Download Cleaned Visitor List Only",
        data=out_buf.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
