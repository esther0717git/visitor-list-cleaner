import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# — sample template download button —
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
    if v == "M":    return "Male"
    if v == "F":    return "Female"
    if v in ("MALE","FEMALE"): return v.title()
    return v

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    # 1) Standardise columns
    df.columns = [
        "S/N","Vehicle Plate Number","Company Full Name","Full Name As Per NRIC",
        "First Name as per NRIC","Middle and Last Name as per NRIC","Identification Type",
        "IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date",
        "Nationality (Country Name)","PR","Gender","Mobile Number"
    ]
    # 2) Drop rows where all of D–M are blank
    df = df.dropna(subset=df.columns[3:13], how="all")

    # 3) Sort & reindex
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df.sort_values(
        by=["Company Full Name","SortGroup","Nationality (Country Name)","Full Name As Per NRIC"],
        inplace=True
    )
    df.drop(columns="SortGroup", inplace=True)
    df["S/N"] = range(1, len(df)+1)

    # 4) Clean vehicles
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
          .str.replace(r"[\/\,]", ";", regex=True)
          .str.replace(r"\s*;\s*", ";", regex=True)
          .str.strip()
          .replace("nan","", regex=False)
    )

    # 5) Name casing & split
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = (
        df["Full Name As Per NRIC"].apply(split_name)
    )

    # 6) Nationality mapping & title case
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .replace({"Chinese":"China","Singaporean":"Singapore"})
          .astype(str).str.title()
    )

    # 7) Swap back if columns were reversed (date vs IC)
    swappable = df["IC (Last 3 digits and suffix) 123A"].astype(str).str.contains("-", na=False)
    if swappable.any():
        df[["IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date"]] = (
            df[["Work Permit Expiry Date","IC (Last 3 digits and suffix) 123A"]]
        )

    # 8) IC suffix & mobile formatting
    df["IC (Last 3 digits and suffix) 123A"] = (
        df["IC (Last 3 digits and suffix) 123A"].astype(str).str[-4:]
    )
    # strip everything except digits
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(r"\D","",regex=True)

    # 9) Gender
    df["Gender"] = df["Gender"].apply(clean_gender)

    # 10) Normalize date
    df["Work Permit Expiry Date"] = (
        pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce")
          .dt.strftime("%Y-%m-%d")
    )

    return df

def generate_excel(df: pd.DataFrame) -> BytesIO:
    """Only writes the cleaned 'Visitor List' sheet and drops all others."""
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # Write cleaned Visitor List
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        wb  = writer.book
        ws  = writer.sheets["Visitor List"]

        # — styling setup —
        header_fill    = PatternFill(start_color="94B455", end_color="94B455", fill_type="solid")
        light_red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        border         = Border(
            left=Side("thin"), right=Side("thin"),
            top=Side("thin"),  bottom=Side("thin")
        )
        center_align = Alignment(horizontal="center", vertical="center")
        base_font    = Font(name="Calibri", size=11)
        bold_font    = Font(name="Calibri", size=11, bold=True)

        # Apply to all cells
        for row in ws.iter_rows():
            for cell in row:
                cell.border    = border
                cell.alignment = center_align
                cell.font      = base_font

        # Header row
        for col in range(1, ws.max_column+1):
            h = ws[f"{get_column_letter(col)}1"]
            h.fill = header_fill
            h.font = bold_font

        # Freeze header
        ws.freeze_panes = "A2"

        # — validation highlights: Singaporeans must use NRIC —
        warnings = 0
        for r in range(2, ws.max_row+1):
            idt  = str(ws[f"G{r}"].value or "").strip().upper()
            nat  = str(ws[f"J{r}"].value or "").strip().title()

            # if nationality is Singapore but ID type is not NRIC, highlight
            if nat == "Singapore" and idt != "NRIC":
                warnings += 1
                for c in ("G","J","K"):
                    ws[f"{c}{r}"].fill = light_red_fill

        # Auto‐fit cols & rows
        for col in ws.columns:
            w = max((len(str(c.value or "")) for c in col), default=0)
            ws.column_dimensions[get_column_letter(col[0].column)].width = w+2
        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 20

        # Vehicles summary
        vals = (
            df["Vehicle Plate Number"]
              .dropna()
              .astype(str)
              .str.split(";")
              .explode()
              .str.strip()
        )
        uniq = sorted(vals[vals!=""].unique())
        ir   = ws.max_row + 2
        if uniq:
            ws[f"B{ir}"].value     = "Vehicles"
            ws[f"B{ir}"].border    = border
            ws[f"B{ir}"].alignment = center_align
            ws[f"B{ir+1}"].value     = ";".join(uniq)
            ws[f"B{ir+1}"].border    = border
            ws[f"B{ir+1}"].alignment = center_align
            ir += 3

        # Total visitors
        total = df["Company Full Name"].notna().sum()
        ws[f"B{ir}"].value     = "Total Visitors"
        ws[f"B{ir}"].border    = border
        ws[f"B{ir}"].alignment = center_align
        ws[f"B{ir+1}"].value     = total
        ws[f"B{ir+1}"].border    = border
        ws[f"B{ir+1}"].alignment = center_align

        # UI warning
        if warnings:
            st.warning(f"⚠️ {warnings} mismatch(es) found. Please review highlighted rows.")

    return out

# — main UI —
uploaded = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])
if uploaded:
    raw = pd.read_excel(uploaded, sheet_name="Visitor List")
    clean = clean_data(raw)
    excel_bytes = generate_excel(clean).getvalue()
    fname = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"

    st.download_button(
        label="📥 Download Cleaned Excel File",
        data=excel_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
