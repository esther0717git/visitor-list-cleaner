import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# Download sample template
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def nationality_group(row):
    nat = str(row["Nationality (Country Name)"]).strip().lower()
    pr  = str(row["PR"]).strip().lower()
    # order: Singapore → PR (any nat) → Malaysia (non-PR) → India → Others
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

def split_name(full):
    parts = str(full).strip().title().split(" ",1)
    return parts if len(parts)>1 else [parts[0], ""]

def clean_data(df):
    # rename & drop fully blank rows D–M
    df.columns = [
        "S/N","Vehicle Plate Number","Company Full Name","Full Name As Per NRIC",
        "First Name as per NRIC","Middle and Last Name as per NRIC","Identification Type",
        "IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date",
        "Nationality (Country Name)","PR","Gender","Mobile Number"
    ]
    df = df.dropna(subset=df.columns[3:], how="all")

    # split names
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = (
        df["Full Name As Per NRIC"].apply(split_name)
    )

    # clean vehicles
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
          .str.replace(r"[\/,]", ";", regex=True)
          .str.replace(r"\s*;\s*", ";", regex=True)
          .str.replace("nan","")
    )

    # nationality map & title-case
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .replace({"Chinese":"China","Singaporean":"Singapore"})
          .astype(str).str.title()
    )

    # swap IC & date if swapped (detect dash in wrong col)
    if df["IC (Last 3 digits and suffix) 123A"].astype(str).str.contains("-", na=False).any():
        df[["IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date"]] = (
            df[["Work Permit Expiry Date","IC (Last 3 digits and suffix) 123A"]]
        )

    # trim IC suffix to last4
    df["IC (Last 3 digits and suffix) 123A"] = (
        df["IC (Last 3 digits and suffix) 123A"].astype(str).str[-4:]
    )

    # mobile → digits only
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(r"\D","",regex=True)

    # gender cleanup
    df["Gender"] = df["Gender"].apply(clean_gender)

    # date format
    df["Work Permit Expiry Date"] = (
        pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce")
          .dt.strftime("%Y-%m-%d")
    )

    # sort & renumber
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df.sort_values(
        by=["Company Full Name","SortGroup","Nationality (Country Name)","Full Name As Per NRIC"],
        inplace=True
    )
    df.drop(columns="SortGroup", inplace=True)
    df["S/N"] = range(1, len(df)+1)

    return df

def generate_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # only write the cleaned Visitor List sheet
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        wb = writer.book
        ws = writer.sheets["Visitor List"]

        # styles
        header_fill = PatternFill(start_color="94B455", end_color="94B455", fill_type="solid")
        warning_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        thin = Side(style="thin")
        border = Border(left=thin,right=thin,top=thin,bottom=thin)
        center = Alignment(horizontal="center",vertical="center")
        font = Font(name="Calibri",size=11)
        bold = Font(name="Calibri",size=11,bold=True)

        # apply to all cells
        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = center
                cell.font = font

        # header row
        for col in range(1, ws.max_column+1):
            c = ws[f"{get_column_letter(col)}1"]
            c.fill = header_fill
            c.font = bold

        # freeze header
        ws.freeze_panes = "A2"

        # highlight ID-Nat mismatches
        mismatch = 0
        for r in range(2, ws.max_row+1):
            idt = str(ws[f"G{r}"].value).strip().lower()
            nat = str(ws[f"J{r}"].value).strip().lower()
            pr  = str(ws[f"K{r}"].value).strip().lower()
            bad = False
            # if Singaporean → must NRIC
            if nat=="singapore" and idt!="nric":
                bad = True
            # if not S’pore & not PR → cannot NRIC
            if nat!="singapore" and pr not in ("yes","pr") and idt=="nric":
                bad = True
            if bad:
                for col in ("G","J","K"):
                    ws[f"{col}{r}"].fill = warning_fill
                mismatch += 1

        # autofit columns
        for col in ws.columns:
            width = max(len(str(cell.value)) for cell in col) + 2
            ws.column_dimensions[get_column_letter(col[0].column)].width = width

        # uniform row height
        for i in range(1, ws.max_row+1):
            ws.row_dimensions[i].height = 20

        # Vehicles summary
        plates = []
        for val in df["Vehicle Plate Number"].dropna():
            plates += [p.strip() for p in str(val).split(";") if p.strip()]
        if plates:
            ins = ws.max_row + 2
            ws[f"B{ins}"].value = "Vehicles"
            ws[f"B{ins+1}"].value = ";".join(sorted(set(plates)))
            for r in (ins,ins+1):
                ws[f"B{r}"].border = border
                ws[f"B{r}"].alignment = center

        # Total Visitors
        total = df["Company Full Name"].notna().sum()
        ins2 = ws.max_row + 5
        ws[f"B{ins2}"].value = "Total Visitors"
        ws[f"B{ins2+1}"].value = total
        for r in (ins2,ins2+1):
            ws[f"B{r}"].border = border
            ws[f"B{r}"].alignment = center

        # show warning banner if any mismatches
        if mismatch:
            st.warning(f"⚠️ Found {mismatch} ID/Nat mismatch(es). Check highlighted rows.")

    return output

# Streamlit UI
uploaded = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])
if uploaded:
    # load only Visitor List sheet
    raw = pd.read_excel(uploaded, sheet_name="Visitor List")
    clean = clean_data(raw)
    out = generate_excel(clean)
    fname = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        label="📥 Download Cleaned File",
        data=out.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
