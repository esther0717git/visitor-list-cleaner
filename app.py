import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# —————————————————————————————
# Download sample template
# —————————————————————————————
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# —————————————————————————————
# Helpers
# —————————————————————————————
def nationality_group(row):
    nat = str(row["Nationality (Country Name)"]).strip().lower()
    pr  = str(row["PR"]).strip().lower()
    if nat == "singapore":
        return 1
    if pr in ("yes","pr"):
        return 2
    if nat == "malaysia":
        return 3
    if nat == "india":
        return 4
    return 5

def clean_gender(val):
    v = str(val).strip().upper()
    return {"M":"Male","F":"Female"}.get(v, v.title())

# —————————————————————————————
# Main cleaning
# —————————————————————————————
def clean_data(df):
    # 1) Standardize column names
    df.columns = [
        "S/N","Vehicle Plate Number","Company Full Name","Full Name As Per NRIC",
        "First Name as per NRIC","Middle and Last Name as per NRIC","Identification Type",
        "IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date",
        "Nationality (Country Name)","PR","Gender","Mobile Number"
    ]

    # 2) Drop rows where D–M all blank
    df = df.dropna(subset=df.columns[3:], how="all")

    # 3) Clean vehicle plate formatting
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
          .str.replace(r"[\/,]", ";", regex=True)
          .str.replace(r"\s*;\s*", ";", regex=True)
          .replace("nan","",regex=False)
    )

    # 4) Title-case full name, then split into exactly two columns
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    names = df["Full Name As Per NRIC"] \
        .str.split(r"\s+", n=1, expand=True)
    # if only one part, pandas gives NaN for second; fill it
    names.columns = ["First Name as per NRIC","Middle and Last Name as per NRIC"]
    names["Middle and Last Name as per NRIC"].fillna("", inplace=True)
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = names

    # 5) Normalize nationality values
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .replace({"Chinese":"China","Singaporean":"Singapore"})
          .astype(str).str.title()
    )

    # 6) Swap IC & expiry if mis-placed (detect dash in IC column)
    if df["IC (Last 3 digits and suffix) 123A"].astype(str).str.contains("-", na=False).any():
        df[["IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date"]] = (
            df[["Work Permit Expiry Date","IC (Last 3 digits and suffix) 123A"]]
        )

    # 7) Trim IC suffix to last 4 chars
    df["IC (Last 3 digits and suffix) 123A"] = (
        df["IC (Last 3 digits and suffix) 123A"].astype(str).str[-4:]
    )

    # 8) Mobile → digits only
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(r"\D","",regex=True)

    # 9) Clean gender
    df["Gender"] = df["Gender"].apply(clean_gender)

    # 10) Date formatting
    df["Work Permit Expiry Date"] = (
        pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce")
          .dt.strftime("%Y-%m-%d")
    )

    # 11) Sort & renumber
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df.sort_values(
        by=["Company Full Name","SortGroup","Nationality (Country Name)","Full Name As Per NRIC"],
        inplace=True
    )
    df.drop(columns="SortGroup", inplace=True)
    df["S/N"] = range(1, len(df)+1)

    return df

# —————————————————————————————
# Write back to Excel
# —————————————————————————————
def generate_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        wb = writer.book
        ws = writer.sheets["Visitor List"]

        # styles
        header_fill = PatternFill(start_color="94B455", end_color="94B455", fill_type="solid")
        warn_fill   = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        thin        = Side(style="thin")
        border      = Border(left=thin,right=thin,top=thin,bottom=thin)
        center      = Alignment(horizontal="center", vertical="center")
        base_font   = Font(name="Calibri", size=11)
        bold_font   = Font(name="Calibri", size=11, bold=True)

        # 1) apply border/align/font to all cells
        for row in ws.iter_rows():
            for cell in row:
                cell.border    = border
                cell.alignment = center
                cell.font      = base_font

        # 2) header row styling
        for col in range(1, ws.max_column+1):
            h = ws[f"{get_column_letter(col)}1"]
            h.fill = header_fill
            h.font = bold_font

        ws.freeze_panes = "A2"

        # 3) ID-Nationality validations
        mismatches = 0
        for r in range(2, ws.max_row+1):
            idt = str(ws[f"G{r}"].value).strip().lower()
            nat = str(ws[f"J{r}"].value).strip().lower()
            pr_status = str(ws[f"K{r}"].value).strip().lower()
            bad = False

            # Singaporeans MUST use NRIC
            if nat=="singapore" and idt!="nric":
                bad = True
            # Non-SG who are not PR MUST NOT use NRIC
            if nat!="singapore" and pr_status not in ("yes","pr") and idt=="nric":
                bad = True

            if bad:
                mismatches += 1
                for col in ("G","J","K"):
                    ws[f"{col}{r}"].fill = warn_fill

        # 4) auto-fit columns
        for col in ws.columns:
            width = max(len(str(cell.value)) for cell in col) + 2
            ws.column_dimensions[get_column_letter(col[0].column)].width = width

        # 5) uniform row height
        for i in range(1, ws.max_row+1):
            ws.row_dimensions[i].height = 20

        # 6) append Vehicles list
        plates = []
        for v in df["Vehicle Plate Number"].dropna():
            plates += [p.strip() for p in str(v).split(";") if p.strip()]
        if plates:
            ins = ws.max_row + 2
            ws[f"B{ins}"].value = "Vehicles"
            ws[f"B{ins+1}"].value = ";".join(sorted(set(plates)))
            for rr in (ins, ins+1):
                ws[f"B{rr}"].border    = border
                ws[f"B{rr}"].alignment = center

        # 7) append total count
        total = df["Company Full Name"].notna().sum()
        ins2 = ws.max_row + 5
        ws[f"B{ins2}"].value   = "Total Visitors"
        ws[f"B{ins2+1}"].value = total
        for rr in (ins2, ins2+1):
            ws[f"B{rr}"].border    = border
            ws[f"B{rr}"].alignment = center

        # 8) show warning if any mismatches
        if mismatches:
            st.warning(f"⚠️ Found {mismatches} ID/​Nationality mismatch(es). Please review highlighted rows.")

    return output

# —————————————————————————————
# Streamlit interaction
# —————————————————————————————
uploaded = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])
if uploaded:
    raw = pd.read_excel(uploaded, sheet_name="Visitor List")
    clean = clean_data(raw)
    out   = generate_excel(clean)

    fname = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        label="📥 Download Cleaned File",
        data=out.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
