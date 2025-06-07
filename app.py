import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

# Configure page
st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# Download sample template
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f.read(),
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Helpers
def nationality_group(row):
    nat = str(row.get("Nationality (Country Name)", "")).lower()
    pr  = str(row.get("PR", "")).strip().lower()
    if nat == "singapore":    return 1
    if pr in ["yes","pr"]:     return 2
    if nat == "malaysia":      return 3
    if nat == "india":         return 4
    return 5

def split_name(name):
    s = str(name).strip()
    if " " in s:
        i = s.find(" ")
        return pd.Series([s[:i], s[i+1:]])
    return pd.Series([s, ""])

def clean_gender(v):
    v2 = str(v).strip().upper()
    if v2 == "M":      return "Male"
    if v2 == "F":      return "Female"
    if v2 in ["MALE","FEMALE"]:
        return v2.title()
    return v2

# Core cleaning of the Visitor List DF
def clean_data(df):
    df = df.copy()
    df.columns = [
        "S/N","Vehicle Plate Number","Company Full Name","Full Name As Per NRIC",
        "First Name as per NRIC","Middle and Last Name as per NRIC","Identification Type",
        "IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date",
        "Nationality (Country Name)","PR","Gender","Mobile Number"
    ]
    # drop rows where D–M are all blank
    df = df.dropna(subset=df.columns[3:13], how="all")
    # sort + renumber
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df.sort_values(
        by=["Company Full Name","SortGroup","Nationality (Country Name)","Full Name As Per NRIC"],
        inplace=True
    )
    df.drop(columns=["SortGroup"], inplace=True)
    df["S/N"] = range(1, len(df)+1)
    # vehicles
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
          .str.replace(r"[\/,]", ";", regex=True)
          .str.replace(r"\s*;\s*", ";", regex=True)
          .str.strip()
          .replace("nan","", regex=False)
    )
    # names
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = \
        df["Full Name As Per NRIC"].apply(split_name)
    # nationality
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .replace({"Chinese":"China","Singaporean":"Singapore"})
          .astype(str)
          .str.title()
    )
    # swap IC & date if mis-placed
    col_ic = "IC (Last 3 digits and suffix) 123A"
    if df[col_ic].astype(str).str.contains("-", na=False).any():
        df[[col_ic,"Work Permit Expiry Date"]] = df[
            ["Work Permit Expiry Date", col_ic]
        ]
    # trim IC
    df[col_ic] = df[col_ic].astype(str).str[-4:]
    # mobile only digits
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(r"\D","",regex=True)
    # gender + date formatting
    df["Gender"] = df["Gender"].apply(clean_gender)
    df["Work Permit Expiry Date"] = (
        pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce")
          .dt.strftime("%Y-%m-%d")
    )
    return df

# Create Excel with ALL sheets, but only re-write & style the Visitor sheet
def generate_excel(all_sheets, cleaned_df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # 1) write every sheet exactly as uploaded
        for name, sheet in all_sheets.items():
            sheet.to_excel(writer, index=False, sheet_name=name)
        # 2) find the visitor sheet name, overwrite it with cleaned + styling
        visitor = next((n for n in all_sheets if "visitor" in n.lower()), None)
        if visitor:
            cleaned_df.to_excel(writer, index=False, sheet_name=visitor)
            ws = writer.book[visitor]
            # styling
            header_fill = PatternFill(start_color="94B455", end_color="94B455", fill_type="solid")
            warn_fill   = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            thin_border = Border(*([Side(style="thin")]*4))
            center      = Alignment(horizontal="center", vertical="center")
            normal_font = Font(name="Calibri", size=11)
            bold_font   = Font(name="Calibri", size=11, bold=True)
            # apply to all
            for row in ws.iter_rows():
                for cell in row:
                    cell.border    = thin_border
                    cell.alignment = center
                    cell.font      = normal_font
            # header row
            for col in range(1, ws.max_column+1):
                h = ws[f"{get_column_letter(col)}1"]
                h.fill = header_fill
                h.font = bold_font
            ws.freeze_panes = "A2"
            # highlight mismatches
            mismatches = 0
            for r in range(2, ws.max_row+1):
                idt = str(ws[f"G{r}"].value).strip().upper()
                nat = str(ws[f"J{r}"].value).strip().title()
                pr  = str(ws[f"K{r}"].value).strip().title()
                bad = False
                if idt=="NRIC" and not (nat=="Singapore" or pr in ["Yes","Pr"]):
                    bad = True
                if idt=="FIN" and (nat=="Singapore" or pr in ["Yes","Pr"]):
                    bad = True
                if bad:
                    mismatches += 1
                    for col in ("G","J","K"):
                        ws[f"{col}{r}"].fill = warn_fill
            # autofit
            for col in ws.columns:
                width = max((len(str(c.value)) for c in col if c.value), default=0)
                ws.column_dimensions[get_column_letter(col[0].column)].width = width+2
            for row in ws.iter_rows():
                ws.row_dimensions[row[0].row].height = 20
            # vehicle summary
            vehicles = []
            for v in cleaned_df["Vehicle Plate Number"].dropna():
                vehicles += [x.strip() for x in str(v).split(";") if x.strip()]
            ir = ws.max_row + 2
            if vehicles:
                summary = ";".join(sorted(set(vehicles)))
                ws[f"B{ir}"].value      = "Vehicles"
                ws[f"B{ir}"].border     = thin_border
                ws[f"B{ir}"].alignment  = center
                ws[f"B{ir+1}"].value    = summary
                ws[f"B{ir+1}"].border   = thin_border
                ws[f"B{ir+1}"].alignment= center
                ir += 3
            # total visitors
            total = cleaned_df["Company Full Name"].notna().sum()
            ws[f"B{ir}"].value      = "Total Visitors"
            ws[f"B{ir}"].border     = thin_border
            ws[f"B{ir}"].alignment  = center
            ws[f"B{ir+1}"].value    = total
            ws[f"B{ir+1}"].border   = thin_border
            ws[f"B{ir+1}"].alignment= center
            # warning banner
            if mismatches:
                st.warning(f"⚠️ {mismatches} potential mismatch(es) found in Identification/Nationality/PR.")
    return out

# ---- Main ----
uploaded = st.file_uploader("📁 Upload your Excel file", type="xlsx")
if uploaded:
    # read all sheets into dict
    all_sheets = pd.read_excel(uploaded, sheet_name=None)
    # find visitor sheet
    vis = next((n for n in all_sheets if "visitor" in n.lower()), None)
    if not vis:
        st.error("❌ Could not find any sheet with “visitor” in its name.")
    else:
        cleaned = clean_data(all_sheets[vis])
        excel_out = generate_excel(all_sheets, cleaned)
        fname = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        st.download_button(
            label="📥 Download Cleaned Excel File",
            data=excel_out.getvalue(),
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
