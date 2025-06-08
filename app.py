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

def split_name(full):
    full = str(full).strip()
    if " " in full:
        i = full.find(" ")
        return pd.Series([full[:i], full[i+1:]])
    return pd.Series([full, ""])

def clean_gender(g):
    g = str(g).strip().upper()
    if g=="M":   return "Male"
    if g=="F":   return "Female"
    if g in ("MALE","FEMALE"):
        return g.title()
    return g

def clean_data(df):
    # 1) normalize columns
    df.columns = [
        "S/N", "Vehicle Plate Number","Company Full Name","Full Name As Per NRIC",
        "First Name as per NRIC","Middle and Last Name as per NRIC",
        "Identification Type","IC (Last 3 digits and suffix) 123A",
        "Work Permit Expiry Date","Nationality (Country Name)","PR","Gender","Mobile Number"
    ]

    # 2) drop rows if columns D–M all blank
    df = df.dropna(subset=df.columns[3:], how="all")

    # 3) split full name
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = \
        df["Full Name As Per NRIC"].apply(split_name)

    # 4) fix nationality naming
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .replace({"Chinese":"China","Singaporean":"Singapore"})
          .astype(str)
          .str.title()
    )

    # 5) swap IC⇆Expiry if they’re flipped:
    #    dash in IC column is a quick heuristic
    mask_swap = (
        df["IC (Last 3 digits and suffix) 123A"]
          .astype(str)
          .str.contains("-", na=False)
    )
    if mask_swap.any():
        df.loc[mask_swap, ["IC (Last 3 digits and suffix) 123A",
                           "Work Permit Expiry Date"]] = \
            df.loc[mask_swap, ["Work Permit Expiry Date",
                               "IC (Last 3 digits and suffix) 123A"]].values

    # 6) IC: last 4 chars
    df["IC (Last 3 digits and suffix) 123A"] = \
        df["IC (Last 3 digits and suffix) 123A"].astype(str).str[-4:]

    # 7) mobile → digits only
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(r"\D","",regex=True)

    # 8) gender cleanup
    df["Gender"] = df["Gender"].apply(clean_gender)

    # 9) expiry date → yyyy-mm-dd
    df["Work Permit Expiry Date"] = (
        pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce")
          .dt.strftime("%Y-%m-%d")
    )

    # 10) grouping + sort
    df["__grp"] = df.apply(nationality_group, axis=1)
    df = df.sort_values(
        ["Company Full Name","__grp","Nationality (Country Name)","Full Name As Per NRIC"]
    )
    df = df.drop(columns="__grp")

    # 11) reset serial
    df["S/N"] = range(1, len(df)+1)
    return df

def generate_clean_workbook(df: pd.DataFrame):
    """
    Create a new workbook in memory, 
    write only 'Visitor List' sheet, style & autosize columns & rows, 
    append Vehicles & Total Visitors summary.
    """
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # write cleaned sheet
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        wb = writer.book
        ws = writer.sheets["Visitor List"]

        # styling presets
        header_fill    = PatternFill("solid", start_color="94B455", end_color="94B455")
        warn_fill      = PatternFill("solid", start_color="FFCCCC", end_color="FFCCCC")
        thin_border    = Border(
            Side("thin"),Side("thin"),Side("thin"),Side("thin")
        )
        center         = Alignment("center","center")
        cal  = Font("Calibri",11)
        cab  = Font("Calibri",11,bold=True)

        # style all cells
        for row in ws.iter_rows():
            for cell in row:
                cell.border = thin_border
                cell.alignment = center
                cell.font = cal

        # header row
        for col in range(1, ws.max_column+1):
            c = ws[f"{get_column_letter(col)}1"]
            c.fill = header_fill
            c.font = cab
        ws.freeze_panes = "A2"

        # highlight mismatches
        mismatches = 0
        for r in range(2, ws.max_row+1):
            it   = str(ws[f"G{r}"].value).strip().upper()
            nat  = str(ws[f"J{r}"].value).strip().title()
            pr   = str(ws[f"K{r}"].value).strip().title()
            bad = (
                (it=="NRIC" and not (nat=="Singapore" or (nat!="Singapore" and pr in ("Yes","Pr"))))
                or
                (it=="FIN" and (nat=="Singapore" or pr in ("Yes","Pr")))
            )
            if bad:
                mismatches += 1
                for col in ("G","J","K"):
                    ws[f"{col}{r}"].fill = warn_fill

        if mismatches:
            st.warning(f"⚠️ {mismatches} validation warning(s) flagged.")

        # autosize columns
        for col in ws.columns:
            max_len = 0
            name    = get_column_letter(col[0].column)
            for c in col:
                v = c.value
                if v is not None:
                    max_len = max(max_len, len(str(v)))
            ws.column_dimensions[name].width = max_len + 2

        # fixed row height
        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 20

        # vehicle summary
        plates = (
            ";".join(
                sorted(
                    set(
                        v.strip() 
                        for v in ";".join(
                            df["Vehicle Plate Number"].dropna().astype(str)
                        ).split(";") 
                        if v.strip()
                    )
                )
            )
        )
        ir = ws.max_row + 2
        if plates:
            ws[f"B{ir}"].value = "Vehicles"
            ws[f"B{ir+1}"].value = plates
            for cell in (f"B{ir}",f"B{ir+1}"):
                ws[cell].border     = thin_border
                ws[cell].alignment  = center

        # total visitors
        tv = df["Company Full Name"].notna().sum()
        ws[f"B{ir+3}"].value = "Total Visitors"
        ws[f"B{ir+4}"].value = tv
        for cell in (f"B{ir+3}",f"B{ir+4}"):
            ws[cell].border     = thin_border
            ws[cell].alignment  = center

    out.seek(0)
    return out

# ——— Streamlit UI ———
uploaded = st.file_uploader("📁 Upload your Excel file", type="xlsx")
if uploaded:
    raw = pd.read_excel(uploaded, sheet_name="Visitor List")
    clean = clean_data(raw)
    wb_io = generate_clean_workbook(clean)

    fname = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        "📥 Download Cleaned Excel File", 
        data=wb_io.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
