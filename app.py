import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# 1) Sample template download
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# 2) Helpers for grouping, splitting, gender cleanup
def nationality_group(row):
    c = str(row["Nationality (Country Name)"]).strip().lower()
    p = str(row["PR"]).strip().lower()
    if c == "singapore":
        return 1
    if p in ("yes", "pr"):
        return 2
    if c == "malaysia":
        return 3
    if c == "india":
        return 4
    return 5

def split_name(name):
    s = str(name).strip()
    if " " in s:
        i = s.find(" ")
        return pd.Series([s[:i], s[i+1:]])
    return pd.Series([s, ""])

def clean_gender(v):
    x = str(v).strip().upper()
    return {"M":"Male","F":"Female"}.get(x, x.title())

# 3) Core cleaning logic on Visitor List df
def clean_visitor_df(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [
        "S/N","Vehicle Plate Number","Company Full Name","Full Name As Per NRIC",
        "First Name as per NRIC","Middle and Last Name as per NRIC","Identification Type",
        "IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date",
        "Nationality (Country Name)","PR","Gender","Mobile Number"
    ]

    # drop entirely blank rows in the core columns
    df = df.dropna(subset=df.columns[3:13], how="all")

    # grouping & sort
    df["_grp"] = df.apply(nationality_group, axis=1)
    df = df.sort_values(
        by=["Company Full Name","_grp","Nationality (Country Name)","Full Name As Per NRIC"],
        ignore_index=True
    ).drop(columns="_grp")

    # reassign running serial
    df["S/N"] = range(1, len(df)+1)

    # plates cleanup
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"]
          .astype(str)
          .str.replace(r"[\/,]", ";", regex=True)
          .str.replace(r"\s*;\s*",";", regex=True)
          .str.strip()
          .replace("nan","",regex=False)
    )

    # name casing & splitting
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = \
        df["Full Name As Per NRIC"].apply(split_name)

    # nationality standardisation
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .replace({"Chinese":"China","Singaporean":"Singapore"})
          .astype(str).str.title()
    )

    # swap IC / Work-Permit columns if mis-placed
    if df["IC (Last 3 digits and suffix) 123A"].astype(str).str.contains("-", na=False).any():
        df[["IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date"]] = \
            df[["Work Permit Expiry Date","IC (Last 3 digits and suffix) 123A"]]

    # IC suffix & mobile cleanup
    df["IC (Last 3 digits and suffix) 123A"] = \
        df["IC (Last 3 digits and suffix) 123A"].astype(str).str[-4:]
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(r"\D","",regex=True)

    # gender & date formatting
    df["Gender"] = df["Gender"].apply(clean_gender)
    df["Work Permit Expiry Date"] = (
        pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce")
          .dt.strftime("%Y-%m-%d")
    )

    return df

# 4) Write out only the cleaned Visitor List sheet
def generate_clean_workbook(xlsx: dict[str,pd.DataFrame], cleaned: pd.DataFrame):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # skip writing all the other sheets → they get dropped
        cleaned.to_excel(writer, index=False, sheet_name="Visitor List")

        wb = writer.book
        ws = writer.sheets["Visitor List"]

        # styling presets
        header_fill   = PatternFill(start_color="94B455", end_color="94B455", fill_type="solid")
        light_red     = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        thin_border   = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin")
        )
        center_align  = Alignment(horizontal="center",vertical="center")
        calibri_11    = Font(name="Calibri", size=11)
        calibri_bold  = Font(name="Calibri", size=11, bold=True)

        # apply to all cells
        for row in ws.iter_rows():
            for c in row:
                c.border = thin_border
                c.alignment = center_align
                c.font = calibri_11

        # header row styling
        for col in range(1, ws.max_column+1):
            cell = ws[f"{get_column_letter(col)}1"]
            cell.fill = header_fill
            cell.font = calibri_bold

        ws.freeze_panes = "A2"

        # highlight invalid ID / nationality combos
        flags = []
        for r in range(2, ws.max_row+1):
            idt = str(ws[f"G{r}"].value).strip().lower()
            nat = str(ws[f"J{r}"].value).strip().lower()
            pr  = str(ws[f"K{r}"].value).strip().lower()
            bad = False

            # singaporeans must use NRIC
            if nat=="singapore" and idt!="nric":
                bad = True
            # non-sg & non-pr must NOT use NRIC
            if nat!="singapore" and pr not in ("yes","pr") and idt=="nric":
                bad = True

            if bad:
                for col in ("G","J","K"):
                    ws[f"{col}{r}"].fill = light_red
                flags.append(r)

        if flags:
            st.warning(f"⚠️ {len(flags)} potential ID/Nationality mismatches highlighted.")

        # auto-fit columns
        for col in ws.columns:
            max_l = 0
            letter = get_column_letter(col[0].column)
            for c in col:
                if c.value is not None:
                    max_l = max(max_l, len(str(c.value)))
            ws.column_dimensions[letter].width = max_l + 2

        # auto-fit row height
        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 20

        # Vehicles & Total Visitors summary
        plates = []
        for v in cleaned["Vehicle Plate Number"].dropna():
            plates += [p.strip() for p in str(v).split(";") if p.strip()]

        start = ws.max_row + 2
        if plates:
            uniq = sorted(set(plates))
            ws[f"B{start}"].value = "Vehicles"
            ws[f"B{start}"].border = thin_border
            ws[f"B{start}"].alignment = center_align

            ws[f"B{start+1}"].value = ";".join(uniq)
            ws[f"B{start+1}"].border = thin_border
            ws[f"B{start+1}"].alignment = center_align

            start += 3
            total = cleaned["Company Full Name"].notna().sum()
            ws[f"B{start}"].value = "Total Visitors"
            ws[f"B{start}"].border = thin_border
            ws[f"B{start}"].alignment = center_align

            ws[f"B{start+1}"].value = total
            ws[f"B{start+1}"].border = thin_border
            ws[f"B{start+1}"].alignment = center_align

    return out

# 5) Streamlit upload/download UI
uploaded = st.file_uploader("📁 Upload your Excel file", type="xlsx")
if uploaded:
    # read all tabs
    all_sheets = pd.read_excel(uploaded, sheet_name=None)
    df0 = all_sheets.get("Visitor List", pd.DataFrame())
    clean = clean_visitor_df(df0)
    book = generate_clean_workbook(all_sheets, clean)

    fn = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        label="📥 Download Cleaned Excel File",
        data=book.getvalue(),
        file_name=fn,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
