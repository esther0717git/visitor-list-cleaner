import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

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
    nat = str(row["Nationality (Country Name)"]).strip().lower()
    pr  = str(row["PR"]).strip().lower()
    if nat == "singapore":
        return 1
    elif pr in ("yes", "pr"):
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

# ───── Cleaning Logic ─────────────────────────────────────────────────────────

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    # 1) rename to exactly these 13 columns
    df.columns = [
        "S/N","Vehicle Plate Number","Company Full Name","Full Name As Per NRIC",
        "First Name as per NRIC","Middle and Last Name as per NRIC","Identification Type",
        "IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date",
        "Nationality (Country Name)","PR","Gender","Mobile Number",
    ]

    # 2) drop rows where columns D–M are all NaN
    df = df.dropna(subset=df.columns[3:13], how="all")

    # 3) nationality mapping (including “Malaysian”→“Malaysia”) + proper-case
    nat_map = {"Chinese":"China","Singaporean":"Singapore","Malaysian":"Malaysia"}
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .replace(nat_map)
          .astype(str)
          .str.title()
    )

    # 4) sort by company → nationality‐group → country → full name
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df = (
        df.sort_values(
            ["Company Full Name","SortGroup","Nationality (Country Name)","Full Name As Per NRIC"],
            ignore_index=True
        )
        .drop(columns="SortGroup")
    )

    # 5) reset S/N
    df["S/N"] = range(1, len(df)+1)

    # 6) clean vehicle plates: “/” or “,”→“;”, trim
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
          .str.replace(r"[\/,]",";",regex=True)
          .str.replace(r"\s*;\s*",";",regex=True)
          .str.strip()
          .replace("nan","",regex=False)
    )

    # 7) proper-case & split full name
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = (
        df["Full Name As Per NRIC"].apply(split_name)
    )

    # 8) swap IC vs Work Permit if mis‐placed
    iccol = "IC (Last 3 digits and suffix) 123A"
    wpcol = "Work Permit Expiry Date"
    if df[iccol].astype(str).str.contains("-", na=False).any():
        df[[iccol,wpcol]] = df[[wpcol,iccol]]

    # 9) trim IC suffix
    df[iccol] = df[iccol].astype(str).str[-4:]

    # 10) digits-only mobile
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(r"\D","",regex=True)

    # 11) normalize gender
    df["Gender"] = df["Gender"].apply(clean_gender)

    # 12) standardize permit date → YYYY-MM-DD
    df[wpcol] = pd.to_datetime(df[wpcol], errors="coerce").dt.strftime("%Y-%m-%d")

    return df

# ───── Generate single‐sheet Excel ────────────────────────────────────────────

def generate_visitor_only(df: pd.DataFrame) -> BytesIO:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        wb = writer.book
        ws = writer.sheets["Visitor List"]

        # style definitions
        header_fill  = PatternFill("solid", fgColor="94B455")
        warn_fill    = PatternFill("solid", fgColor="FFCCCC")
        border       = Border(Side("thin"),Side("thin"),Side("thin"),Side("thin"))
        center       = Alignment("center","center")
        normal_font  = Font("Calibri",11)
        bold_font    = Font("Calibri",11,bold=True)

        # 1) apply border/align/font to all cells
        for row in ws.iter_rows():
            for cell in row:
                cell.border    = border
                cell.alignment = center
                cell.font      = normal_font

        # 2) header styling
        for col in range(1, ws.max_column+1):
            c = ws[f"{get_column_letter(col)}1"]
            c.fill = header_fill
            c.font = bold_font

        # 3) freeze top row
        ws.freeze_panes = ws["A2"]

        # 4) highlight ID vs PR violations:
        #    • NRIC ok if PR is yes/PR
        #    • all others (WORK PERMIT, FIN, OTHERS) must NOT have PR filled
        bad_count = 0
        for r in range(2, ws.max_row+1):
            idt = str(ws[f"G{r}"].value).strip().upper()
            pr  = str(ws[f"K{r}"].value).strip().upper()
            # if non-NRIC and PR is non-empty → highlight K cell
            if idt not in ("NRIC",) and pr not in ("", "NO", "N"):
                ws[f"K{r}"].fill = warn_fill
                bad_count += 1

        if bad_count:
            st.warning(f"⚠️ {bad_count} non‐NRIC row(s) with a PR value detected.")

        # 5) auto-fit & set row height
        for col in ws.columns:
            w = max(len(str(c.value)) for c in col if c.value)
            ws.column_dimensions[get_column_letter(col[0].column)].width = w+2
        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 20

    return out

# ───── Streamlit UI ────────────────────────────────────────────────────────────

uploaded = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])
if uploaded:
    raw   = pd.read_excel(uploaded, sheet_name="Visitor List")
    clean = clean_data(raw)
    buf   = generate_visitor_only(clean)
    fname = f"Cleaned_VisitorList_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        label="📥 Download Cleaned Visitor List Only",
        data=buf.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
