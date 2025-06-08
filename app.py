import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# -- Download sample template --
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f.read(),
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -- Helper functions --
def nationality_group(row):
    nat = str(row["Nationality (Country Name)"]).strip().lower()
    pr  = str(row["PR"]).strip().lower()
    if nat == "singapore":
        return 1
    if pr in ("yes", "pr"):
        return 2
    # (you can add more fixed groups if needed)
    return 3

def split_name(full):
    s = str(full).strip()
    if " " in s:
        i = s.find(" ")
        return pd.Series([s[:i], s[i+1:]])
    return pd.Series([s, ""])

def clean_gender(g):
    g = str(g).strip().upper()
    if g in ("M", "MALE"):
        return "Male"
    if g in ("F", "FEMALE"):
        return "Female"
    return g.title()

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    # -- standardise headings --
    df = df.copy()
    df.columns = [
        "S/N", "Vehicle Plate Number", "Company Full Name", "Full Name As Per NRIC",
        "First Name as per NRIC", "Middle and Last Name as per NRIC", "Identification Type",
        "IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date",
        "Nationality (Country Name)", "PR", "Gender", "Mobile Number"
    ][: len(df.columns)]
    # drop fully blank rows (columns D–M)
    df = df.dropna(subset=df.columns[3:], how="all")

    # fix swapped IC / WP columns dynamically
    ic_col = next((c for c in df.columns if c.lower().startswith("ic (last 3")), None)
    wp_col = next((c for c in df.columns if "work permit expiry" in c.lower()), None)
    if ic_col and wp_col:
        mask = (
            df[ic_col].astype(str).str.match(r"\d{4}-\d{2}-\d{2}", na=False)
            | df[wp_col].astype(str).str.match(r".{3}\d", na=False)
        )
        df.loc[mask, [ic_col, wp_col]] = df.loc[mask, [wp_col, ic_col]].values

    # parse & format
    df["IC (Last 3 digits and suffix) 123A"] = (
        df["IC (Last 3 digits and suffix) 123A"]
        .astype(str).str[-4:]
    )
    df["Work Permit Expiry Date"] = (
        pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce")
          .dt.strftime("%Y-%m-%d")
    )

    # standardise other fields
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
          .str.replace(r"[\/,]", ";", regex=True)
          .str.replace(r"\s*;\s*", ";", regex=True)
          .str.strip().replace("nan", "")
    )
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = (
        df["Full Name As Per NRIC"].apply(split_name)
    )
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .astype(str)
          .replace({"Chinese": "China", "Singaporean": "Singapore"})
          .str.title()
    )
    df["Gender"] = df["Gender"].apply(clean_gender)
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(r"\D", "", regex=True)

    # sort by company → nationality group → nationality → name
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df.sort_values(
        by=["Company Full Name", "SortGroup", "Nationality (Country Name)", "Full Name As Per NRIC"],
        inplace=True
    )
    df.drop(columns="SortGroup", inplace=True)
    df.reset_index(drop=True, inplace=True)

    # re-assign serial numbers
    df["S/N"] = range(1, len(df)+1)
    return df

def generate_excel(df: pd.DataFrame):
    wb = Workbook()
    ws = wb.active
    ws.title = "Visitor List"

    # write dataframe to sheet
    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        for c_idx, val in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=val)

    # write headers
    header = list(df.columns)
    for c_idx, h in enumerate(header, start=1):
        cell = ws.cell(row=1, column=c_idx, value=h)
        cell.fill = PatternFill("solid", fgColor="94B455")
        cell.font = Font(bold=True, name="Calibri", size=11)
        cell.alignment = Alignment("center","center")

    # global styles: borders, alignment, default font
    thin = Side("thin")
    border = Border(thin,thin,thin,thin)
    default_font = Font(name="Calibri", size=11)
    for row in ws.iter_rows(min_row=2, max_row=1+len(df), max_col=len(header)):
        for cell in row:
            cell.border = border
            cell.font = default_font
            cell.alignment = Alignment("center","center")

    # freeze header
    ws.freeze_panes = "A2"

    # highlight ID mismatches
    red_fill = PatternFill("solid", fgColor="FFCCCC")
    mismatches = 0
    for r in range(2, 2+len(df)):
        nat = ws[f"J{r}"].value or ""
        idt = str(ws[f"G{r}"].value or "").strip().upper()
        # if Singapore but not NRIC, or non-Singapore but NRIC
        if (nat=="Singapore" and idt!="NRIC") or (nat!="Singapore" and idt=="NRIC"):
            for col in ("G","J"):
                ws[f"{col}{r}"].fill = red_fill
            mismatches += 1
    if mismatches:
        st.warning(f"⚠️ {mismatches} Identification/Nationality mismatch(es) highlighted.")

    # auto-fit widths
    for col in ws.columns:
        max_len = max(len(str(c.value)) for c in col if c.value) + 2
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_len

    # set uniform row height
    for r in range(1, 2+len(df)):
        ws.row_dimensions[r].height = 20

    # append Vehicles summary
    plates = ";".join(sorted(set(
        v.strip() for val in df["Vehicle Plate Number"] for v in str(val).split(";") if v.strip()
    )))
    vt_r = 2 + len(df) + 1
    ws.cell(vt_r, 2, "Vehicles").border = border
    ws.cell(vt_r+1,2, plates).border = border

    # append Total Visitors
    tv_r = vt_r + 3
    ws.cell(tv_r,2, "Total Visitors").border = border
    ws.cell(tv_r+1,2, len(df)).border = border

    # return Byte stream
    bio = BytesIO()
    wb.save(bio)
    return bio

# -- Streamlit UI --
uploaded = st.file_uploader("📁 Upload your Excel file", type="xlsx")
if uploaded:
    all_sheets = pd.read_excel(uploaded, sheet_name=None)
    # pick sheet with "visitor" in its name
    raw = next(
        df for name, df in all_sheets.items()
        if "visitor" in name.lower()
    )
    cleaned = clean_data(raw)
    excel_io = generate_excel(cleaned)

    fn = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        label="📥 Download Cleaned Excel File",
        data=excel_io.getvalue(),
        file_name=fn,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
