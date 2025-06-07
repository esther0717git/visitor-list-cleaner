import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("CG1")

# ─── Sample template download ───────────────────────────────────────────────
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ─── Helper functions ────────────────────────────────────────────────────────
def nationality_group(row):
    nat = str(row["Nationality (Country Name)"]).strip().lower()
    pr  = str(row["PR"]).strip().lower()
    if nat == "singapore":
        return 1
    if pr in ("yes", "pr"):
        return 2
    if nat == "malaysia":
        return 3
    if nat == "india":
        return 4
    return 5

def split_name(full):
    parts = str(full).strip().title().split(" ", 1)
    return pd.Series([parts[0], parts[1] if len(parts)>1 else ""])

def clean_gender(g):
    u = str(g).strip().upper()
    return {"M":"Male", "F":"Female"}.get(u, u.title())

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    # 1) Normalize headers
    df.columns = [
      "S/N","Vehicle Plate Number","Company Full Name","Full Name As Per NRIC",
      "First Name as per NRIC","Middle and Last Name as per NRIC",
      "Identification Type","IC (Last 3 digits and suffix) 123A",
      "Work Permit Expiry Date","Nationality (Country Name)",
      "PR","Gender","Mobile Number"
    ]

    # 2) Drop fully blank visitor rows
    df = df.dropna(subset=df.columns[3:], how="all")

    # 3) Clean plate numbers
    df["Vehicle Plate Number"] = (
      df["Vehicle Plate Number"].astype(str)
        .str.replace(r"[\/,]", ";", regex=True)
        .str.replace(r"\s*;\s*", ";", regex=True)
        .str.strip()
        .replace("nan","",regex=False)
    )

    # 4) Proper-case names & split
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = (
      df["Full Name As Per NRIC"].apply(split_name)
    )

    # 5) Nationality mapping + title-case
    df["Nationality (Country Name)"] = (
      df["Nationality (Country Name)"]
        .replace({"Singaporean":"Singapore","Chinese":"China"})
        .astype(str).str.title()
    )

    # 6) Swap IC vs date if mis-placed
    mask = df["IC (Last 3 digits and suffix) 123A"].astype(str).str.contains(r"\d{4}-", na=False)
    df.loc[mask, ["IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date"]] = \
      df.loc[mask, ["Work Permit Expiry Date","IC (Last 3 digits and suffix) 123A"]].values

    # 7) Trim IC suffix to last 4 chars
    df["IC (Last 3 digits and suffix) 123A"] = (
      df["IC (Last 3 digits and suffix) 123A"].astype(str).str[-4:]
    )

    # 8) Strip non-digits from mobile (no decimals)
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(r"\D","",regex=True)

    # 9) Clean Gender
    df["Gender"] = df["Gender"].apply(clean_gender)

    # 10) Standardize date
    df["Work Permit Expiry Date"] = (
      pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce")
        .dt.strftime("%Y-%m-%d")
    )

    # 11) Sort by Company → nationality group (stable so original order within group)
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df.sort_values(
      ["Company Full Name","SortGroup"],
      inplace=True,
      kind="stable"
    )
    df.drop(columns="SortGroup", inplace=True)

    # 12) Re-assign serial numbers
    df["S/N"] = range(1, len(df)+1)

    return df

# ─── Excel generation ────────────────────────────────────────────────────────
def generate_excel(cleaned: pd.DataFrame) -> BytesIO:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # Only write the cleaned Visitor List
        cleaned.to_excel(writer, index=False, sheet_name="Visitor List")

        wb = writer.book
        ws = wb["Visitor List"]

        # Styling setup
        hdr_fill    = PatternFill("solid", fgColor="94B455")
        warn_fill   = PatternFill("solid", fgColor="FFCCCC")
        border      = Border(Side("thin"),Side("thin"),Side("thin"),Side("thin"))
        center      = Alignment("center","center")
        font_base   = Font("Calibri",11)
        font_bold   = Font("Calibri",11,bold=True)

        # Apply to every cell
        for row in ws.iter_rows():
            for c in row:
                c.border    = border
                c.alignment = center
                c.font      = font_base

        # Header row
        for col in range(1, ws.max_column+1):
            h = ws[f"{get_column_letter(col)}1"]
            h.fill = hdr_fill
            h.font = font_bold

        # Freeze top row
        ws.freeze_panes = ws["A2"]

        # Highlight mismatches
        mismatches = 0
        for r in range(2, ws.max_row+1):
            idt = str(ws[f"G{r}"].value or "").strip().upper()
            nat = str(ws[f"J{r}"].value or "").strip().title()
            pr  = str(ws[f"K{r}"].value or "").strip().lower()

            bad = False
            # 4 ID rules
            if nat == "Singapore" and idt != "NRIC":
                bad = True
            if idt == "NRIC" and nat != "Singapore" and pr not in ("yes","pr"):
                bad = True
            if idt == "FIN" and (nat == "Singapore" or pr in ("yes","pr")):
                bad = True

            if bad:
                for col in ("G","J","K"):
                    ws[f"{col}{r}"].fill = warn_fill
                mismatches += 1

        # Auto-fit widths & heights
        for col in ws.columns:
            width = max(len(str(c.value or "")) for c in col) + 2
            ws.column_dimensions[get_column_letter(col[0].column)].width = width
        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 20

        if mismatches:
            st.warning(f"⚠️ {mismatches} potential mismatch(es) found. Please review highlighted rows.")

    return out

# ─── Streamlit UI ────────────────────────────────────────────────────────────
uploaded = st.file_uploader("📁 Upload your Excel file", type="xlsx")
if uploaded:
    # Read only the Visitor List
    sheets = pd.read_excel(uploaded, sheet_name=None)
    vis   = sheets.get("Visitor List", pd.DataFrame())
    cleaned = clean_data(vis)

    buf = generate_excel(cleaned)
    fname = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        label="📥 Download Cleaned Excel File",
        data=buf.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
