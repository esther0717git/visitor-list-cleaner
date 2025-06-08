import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# Download button for your template
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def nationality_group(row):
    nat = str(row["Nationality (Country Name)"]).strip().lower()
    pr = str(row["PR"]).strip().lower()
    if nat == "singapore":
        return 1
    elif pr in ["yes", "pr"]:
        return 2
    elif nat == "malaysia":
        return 3
    elif nat == "india":
        return 4
    else:
        return 5

def split_name(full):
    s = str(full).strip()
    if " " in s:
        i = s.find(" ")
        return pd.Series([s[:i], s[i+1:]])
    return pd.Series([s, ""])

def clean_gender(g):
    g = str(g).strip().upper()
    if g == "M":
        return "Male"
    if g == "F":
        return "Female"
    if g in ["MALE","FEMALE"]:
        return g.title()
    return g

def clean_data(df):
    # standardize column names
    df.columns = [
        "S/N","Vehicle Plate Number","Company Full Name","Full Name As Per NRIC",
        "First Name as per NRIC","Middle and Last Name as per NRIC","Identification Type",
        "IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date",
        "Nationality (Country Name)","PR","Gender","Mobile Number"
    ]

    # drop blank rows
    df = df.dropna(subset=df.columns[3:13], how="all")

    # sort by company → nationality group → nationality → name
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df.sort_values(
        by=["Company Full Name","SortGroup","Nationality (Country Name)","Full Name As Per NRIC"],
        inplace=True
    )
    df.drop(columns="SortGroup", inplace=True)

    # re-index S/N
    df["S/N"] = range(1, len(df)+1)

    # clean up vehicle cells
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
          .str.replace(r"[\/,]", ";", regex=True)
          .str.replace(r"\s*;\s*", ";", regex=True)
          .str.strip()
          .replace("nan","", regex=False)
    )

    # proper case full name → split
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = \
        df["Full Name As Per NRIC"].apply(split_name)

    # normalize nationality text
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .replace({"Chinese":"China","Singaporean":"Singapore"})
          .astype(str)
          .str.title()
    )

    # —— HERE’S THE NEW SWAP LOGIC —— #
    ic_raw = df["IC (Last 3 digits and suffix) 123A"].astype(str)
    wp_raw = df["Work Permit Expiry Date"].astype(str)

    # detect real dates in the IC column
    is_date_in_ic = pd.to_datetime(ic_raw, format="%Y-%m-%d", errors="coerce").notna()

    if is_date_in_ic.any():
        # swap only those rows
        df.loc[is_date_in_ic, [
            "IC (Last 3 digits and suffix) 123A",
            "Work Permit Expiry Date"
        ]] = df.loc[is_date_in_ic, [
            "Work Permit Expiry Date",
            "IC (Last 3 digits and suffix) 123A"
        ]].values

    # finalize IC suffix
    df["IC (Last 3 digits and suffix) 123A"] = (
        df["IC (Last 3 digits and suffix) 123A"]
          .astype(str)
          .str[-4:]
          .replace("nan","",regex=False)
    )

    # finalize work-permit date
    df["Work Permit Expiry Date"] = (
        pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce")
          .dt.strftime("%Y-%m-%d")
          .fillna("")
    )

    # mobile = digits only
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(r"\D","",regex=True)

    df["Gender"] = df["Gender"].apply(clean_gender)

    return df

def generate_excel(df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # write only the cleaned Visitor List sheet
        df.to_excel(writer, index=False, sheet_name="Visitor List")

        wb = writer.book
        ws = writer.sheets["Visitor List"]

        # styling objects
        header_fill   = PatternFill("solid", fgColor="94B455")
        warning_fill  = PatternFill("solid", fgColor="FFCCCC")
        thin_border   = Border(
            left=Side("thin"), right=Side("thin"),
            top=Side("thin"), bottom=Side("thin")
        )
        center_align  = Alignment("center","center")
        font_regular  = Font(name="Calibri", size=11)
        font_bold     = Font(name="Calibri", size=11, bold=True)

        # border + alignment + font
        for row in ws.iter_rows():
            for cell in row:
                cell.border = thin_border
                cell.alignment = center_align
                cell.font = font_regular

        # header row style
        for col_idx in range(1, ws.max_column+1):
            h = ws.cell(row=1, column=col_idx)
            h.fill = header_fill
            h.font = font_bold

        ws.freeze_panes = "A2"

        # highlight invalid Identification → Nat/PR combos
        warnings = []
        for r in range(2, ws.max_row+1):
            it = ws[f"G{r}"].value or ""
            nat = (ws[f"J{r}"].value or "").strip().title()
            pr  = (ws[f"K{r}"].value or "").strip().title()

            bad = False
            it_up = it.strip().upper()
            if it_up=="NRIC" and nat!="Singapore":
                bad = True
            if it_up=="WORK PERMIT" and nat=="Singapore":
                bad = True
            if it_up=="FIN" and nat=="Singapore":
                bad = True
            # (others pass)

            if bad:
                warnings.append(r)
                for col in ["G","J","K"]:
                    ws[f"{col}{r}"].fill = warning_fill

        if warnings:
            st.warning(f"⚠️ {len(warnings)} row(s) with mismatched ID/Nat—see highlights.")

        # auto‐fit columns + fixed row height
        for col in ws.columns:
            m = 0
            letter = get_column_letter(col[0].column)
            for cell in col:
                val = cell.value
                if val is not None:
                    m = max(m, len(str(val)))
            ws.column_dimensions[letter].width = m + 2

        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 20

        # append Vehicles summary
        plates = []
        for v in df["Vehicle Plate Number"].dropna():
            plates += [p.strip() for p in str(v).split(";") if p.strip()]
        if plates:
            ins = ws.max_row + 2
            ws[f"B{ins}"].value = "Vehicles"
            ws[f"B{ins}"].border = thin_border
            ws[f"B{ins}"].alignment = center_align
            ws[f"B{ins+1}"].value = ";".join(sorted(set(plates)))
            ws[f"B{ins+1}"].border = thin_border
            ws[f"B{ins+1}"].alignment = center_align

            ins += 3
        else:
            ins = ws.max_row + 2

        # append Total Visitors
        ws[f"B{ins}"].value = "Total Visitors"
        ws[f"B{ins}"].border = thin_border
        ws[f"B{ins}"].alignment = center_align
        ws[f"B{ins+1}"].value = df["Company Full Name"].notna().sum()
        ws[f"B{ins+1}"].border = thin_border
        ws[f"B{ins+1}"].alignment = center_align

    return out.getvalue()

# Streamlit UI
uploaded = st.file_uploader("📁 Upload your Excel file", type="xlsx")
if uploaded:
    raw = pd.read_excel(uploaded, sheet_name="Visitor List")
    cleaned = clean_data(raw)
    result = generate_excel(cleaned)
    fname = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        "📥 Download Cleaned Excel File",
        data=result,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
