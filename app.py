import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# ——————————————
# Download sample template
# ——————————————
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ——————————————
# Helper functions
# ——————————————
def nationality_group(row):
    nat = str(row["Nationality (Country Name)"]).strip().lower()
    pr  = str(row["PR"]).strip().lower()
    if nat == "singapore":
        return 1
    elif pr in ("yes", "pr"):
        return 2
    else:
        return 3

def split_name(name):
    s = str(name).strip()
    if " " in s:
        i = s.index(" ")
        return pd.Series([s[:i], s[i+1:]])
    return pd.Series([s, ""])

def clean_gender(v):
    v = str(v).strip().upper()
    if v == "M": return "Male"
    if v == "F": return "Female"
    if v in ("MALE","FEMALE"): return v.title()
    return v

# ——————————————
# Core cleaning
# ——————————————
def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    # rename columns
    df.columns = [
        "S/N", "Vehicle Plate Number", "Company Full Name", "Full Name As Per NRIC",
        "First Name as per NRIC", "Middle and Last Name as per NRIC", "Identification Type",
        "IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date",
        "Nationality (Country Name)", "PR", "Gender", "Mobile Number"
    ]

    # drop fully blank visitor rows
    df = df.dropna(subset=df.columns[3:], how="all")

    # sort into 3 buckets: SG → PR → Others, then by company & name
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df.sort_values(
        by=["Company Full Name", "SortGroup", "Full Name As Per NRIC"],
        inplace=True
    )
    df.drop(columns="SortGroup", inplace=True)

    # reset serial
    df["S/N"] = range(1, len(df)+1)

    # vehicles: slash/comma → semicolon, trim, drop 'nan'
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
          .str.replace(r"[\/,]", ";", regex=True)
          .str.replace(r"\s*;\s*", ";", regex=True)
          .str.strip()
          .replace("nan", "", regex=False)
    )

    # proper-case names + split
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = (
        df["Full Name As Per NRIC"].apply(split_name)
    )

    # nationality mapping + title-case
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .replace({"Chinese":"China","Singaporean":"Singapore"})
          .astype(str).str.title()
    )

    # swap if columns got reversed (we look for a dash in the IC field)
    ic_col = df["IC (Last 3 digits and suffix) 123A"].astype(str)
    if ic_col.str.contains("-", na=False).any():
        df[["IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date"]] = (
            df[["Work Permit Expiry Date","IC (Last 3 digits and suffix) 123A"]]
        )

    # last-4 of IC
    df["IC (Last 3 digits and suffix) 123A"] = (
        df["IC (Last 3 digits and suffix) 123A"].astype(str).str[-4:]
    )

    # clean mobile → digits only
    df["Mobile Number"] = (
        df["Mobile Number"].astype(str).str.replace(r"\D+", "", regex=True)
    )

    # gender normalize
    df["Gender"] = df["Gender"].apply(clean_gender)

    # date format
    df["Work Permit Expiry Date"] = (
        pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce")
          .dt.strftime("%Y-%m-%d")
    )

    return df

# ——————————————
# Build output Excel
# ——————————————
def generate_excel(df: pd.DataFrame) -> BytesIO:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # write cleaned Visitor List only
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        wb = writer.book
        ws = writer.sheets["Visitor List"]

        # styling
        header_fill  = PatternFill("solid", fgColor="94B455")
        warn_fill    = PatternFill("solid", fgColor="FFCCCC")
        thin_border  = Border(*(Side("thin"),)*4)
        center_align = Alignment("center","center")
        font_body    = Font(name="Calibri", size=11)
        font_bold    = Font(name="Calibri", size=11, bold=True)

        # apply border + align + font to all
        for row in ws.iter_rows():
            for cell in row:
                cell.border    = thin_border
                cell.alignment = center_align
                cell.font      = font_body

        # header row
        for col in range(1, ws.max_column+1):
            c = ws[f"{get_column_letter(col)}1"]
            c.fill = header_fill
            c.font = font_bold

        ws.freeze_panes = "A2"

        # validation highlights
        mismatches = 0
        for r in range(2, ws.max_row+1):
            it  = str(ws[f"G{r}"].value).strip().upper()
            nat = str(ws[f"J{r}"].value).strip().title()
            pr  = str(ws[f"K{r}"].value).strip().title()

            bad = False
            if nat == "Singapore" and it != "NRIC":
                bad = True
            if it == "NRIC" and not (nat=="Singapore" or pr in ("Yes","Pr")):
                bad = True

            if bad:
                mismatches +=1
                for col in ("G","J","K"):
                    ws[f"{col}{r}"].fill = warn_fill

        # vehicles summary
        vehicles = []
        for v in df["Vehicle Plate Number"].dropna():
            vehicles += [x.strip() for x in str(v).split(";") if x.strip()]
        vehicles = sorted(set(vehicles))

        if vehicles:
            start = ws.max_row + 2
            ws[f"B{start}"].value = "Vehicles"
            ws[f"B{start+1}"].value = ";".join(vehicles)
            for rr in (start, start+1):
                cell = ws[f"B{rr}"]
                cell.border    = thin_border
                cell.alignment = center_align

        # total visitors
        tv = df["Company Full Name"].notna().sum()
        trow = ws.max_row + 4 if vehicles else ws.max_row+2
        ws[f"B{trow}"].value   = "Total Visitors"
        ws[f"B{trow+1}"].value = tv
        for rr in (trow, trow+1):
            c = ws[f"B{rr}"]
            c.border    = thin_border
            c.alignment = center_align

    out.seek(0)
    return out

# ——————————————
# Streamlit UI
# ——————————————
uploaded = st.file_uploader("📁 Upload your Excel file", type="xlsx")
if uploaded:
    raw = pd.read_excel(uploaded, sheet_name="Visitor List")
    clean = clean_data(raw)
    excel_io = generate_excel(clean)

    name = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        label="📥 Download Cleaned Excel File",
        data=excel_io,
        file_name=name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
