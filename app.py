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
    nat = str(row["Nationality (Country Name)"]).lower()
    pr  = str(row["PR"]).strip().lower()
    if nat == "singapore":
        return 1
    elif pr in ("yes","pr"):
        return 2
    elif nat == "malaysia":
        return 3
    elif nat == "india":
        return 4
    else:
        return 5

def split_name(name):
    s = str(name).strip()
    if " " in s:
        i = s.find(" ")
        return pd.Series([s[:i], s[i+1:]])
    return pd.Series([s, ""])

def clean_gender(val):
    v = str(val).strip().upper()
    if v=="M": return "Male"
    if v=="F": return "Female"
    if v in ("MALE","FEMALE"): return v.title()
    return v

def clean_data(df):
    # rename & drop fully blank rows
    df.columns = [
        "S/N", "Vehicle Plate Number", "Company Full Name", "Full Name As Per NRIC",
        "First Name as per NRIC", "Middle and Last Name as per NRIC", "Identification Type",
        "IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date",
        "Nationality (Country Name)", "PR", "Gender", "Mobile Number"
    ]
    df = df.dropna(axis=0, subset=df.columns[3:], how="all")

    # sort & reindex
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df.sort_values(
        by=["Company Full Name","SortGroup","Nationality (Country Name)","Full Name As Per NRIC"],
        inplace=True
    )
    df.drop(columns="SortGroup", inplace=True)
    df["S/N"] = range(1, len(df)+1)

    # vehicles
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
          .str.replace(r"[\/\,]", ";", regex=True)
          .str.replace(r"\s*;\s*", ";", regex=True)
          .str.strip()
          .replace("nan","",regex=False)
    )

    # names
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = \
        df["Full Name As Per NRIC"].apply(split_name)

    # nationality
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .replace({"Chinese":"China","Singaporean":"Singapore"})
          .astype(str).str.title()
    )

    # swap if mistakenly placed
    if df["IC (Last 3 digits and suffix) 123A"].astype(str).str.contains("-",na=False).any():
        df[[
          "IC (Last 3 digits and suffix) 123A",
          "Work Permit Expiry Date"
        ]] = df[[
          "Work Permit Expiry Date",
          "IC (Last 3 digits and suffix) 123A"
        ]]

    df["IC (Last 3 digits and suffix) 123A"] = df["IC (Last 3 digits and suffix) 123A"].astype(str).str[-4:]
    # mobile = digits only
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(r"\D","",regex=True)
    df["Gender"] = df["Gender"].apply(clean_gender)
    df["Work Permit Expiry Date"] = pd.to_datetime(
        df["Work Permit Expiry Date"], errors="coerce"
    ).dt.strftime("%Y-%m-%d")

    return df

def generate_excel(all_sheets, cleaned_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) re-write untouched sheets
        for name, sheet in all_sheets.items():
            if "visitor" not in name.lower():
                sheet.to_excel(writer, index=False, sheet_name=name)

        # 2) write cleaned Visitor sheet
        cleaned_df.to_excel(writer, index=False, sheet_name="Visitor List")
        wb = writer.book
        ws = wb["Visitor List"]

        # styles
        header_fill    = PatternFill("solid", fgColor="94B455")
        warning_fill   = PatternFill("solid", fgColor="FFCCCC")
        border         = Border(
            left=Side("thin"), right=Side("thin"),
            top=Side("thin"),  bottom=Side("thin")
        )
        center_align   = Alignment("center","center")
        calibri        = Font("Calibri",11)
        bold_calibri   = Font("Calibri",11,bold=True)

        # apply to all cells
        for row in ws.iter_rows():
            for cell in row:
                cell.border    = border
                cell.alignment = center_align
                cell.font      = calibri

        # header row
        for col in range(1, ws.max_column+1):
            c = ws[f"{get_column_letter(col)}1"]
            c.fill = header_fill
            c.font = bold_calibri

        ws.freeze_panes = "A2"

        # highlight mismatches
        mismatch_count = 0
        for r in range(2, ws.max_row+1):
            t = str(ws[f"G{r}"].value).strip().upper()   # ID type
            n = str(ws[f"J{r}"].value).strip().title()   # nationality
            p = str(ws[f"K{r}"].value).strip().title()   # PR
            bad = (
                (t=="NRIC" and not (n=="Singapore" or (n!="Singapore" and p in ("Yes","Pr"))))
                or
                (t=="FIN" and (n=="Singapore" or p in ("Yes","Pr")))
            )
            if bad:
                for col in ("G","J","K"):
                    ws[f"{col}{r}"].fill = warning_fill
                mismatch_count += 1

        # autosize
        for col in ws.columns:
            letter = get_column_letter(col[0].column)
            maxlen = max((len(str(c.value)) for c in col if c.value), default=0)
            ws.column_dimensions[letter].width = maxlen + 2
        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 20

        # Vehicles summary
        vals = cleaned_df["Vehicle Plate Number"].dropna().str.split(";").explode().str.strip()
                # Vehicles summary
        vals = cleaned_df["Vehicle Plate Number"] \
            .dropna() \
            .str.split(";") \
            .explode() \
            .str.strip()

        # filter out empty strings, then get unique and sort
        unique_v = sorted(vals[vals != ""].unique())

        insert_row = ws.max_row + 2
        if unique_v:
            ws[f"B{insert_row}"].value     = "Vehicles"
            ws[f"B{insert_row}"].border    = border
            ws[f"B{insert_row}"].alignment = center_align

            ws[f"B{insert_row+1}"].value     = ";".join(unique_v)
            ws[f"B{insert_row+1}"].border    = border
            ws[f"B{insert_row+1}"].alignment = center_align

            insert_row += 3

        if unique_v:
            r0 = ws.max_row + 2
            ws[f"B{r0}"].value = "Vehicles"
            ws[f"B{r0}"].border    = border
            ws[f"B{r0}"].alignment = center_align
            ws[f"B{r0+1}"].value   = ";".join(unique_v)
            ws[f"B{r0+1}"].border    = border
            ws[f"B{r0+1}"].alignment = center_align
            last = r0+3
        else:
            last = ws.max_row + 2

        # Total visitors summary
        total_vis = cleaned_df["Company Full Name"].notna().sum()
        ws[f"B{last}"].value      = "Total Visitors"
        ws[f"B{last}"].border     = border
        ws[f"B{last}"].alignment  = center_align
        ws[f"B{last+1}"].value    = total_vis
        ws[f"B{last+1}"].border   = border
        ws[f"B{last+1}"].alignment= center_align

    return output, mismatch_count

# ──────────────────────────────

uploaded = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])
if uploaded:
    # read all sheets
    all_sheets = pd.read_excel(uploaded, sheet_name=None)
    # locate the Visitor sheet (case-insensitive)
    for nm in all_sheets:
        if "visitor" in nm.lower():
            raw = all_sheets[nm]
            break
    else:
        st.error("❌ No sheet with ‘Visitor’ in its name found.")
        st.stop()

    cleaned = clean_data(raw)
    out, bad = generate_excel(all_sheets, cleaned)

    if bad:
        st.warning(f"⚠️ {bad} row(s) flagged for you to review.")

    fname = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        "📥 Download Cleaned Excel File",
        data=out.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
