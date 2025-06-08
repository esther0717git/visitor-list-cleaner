import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

#
# ───── Download Sample Template Button ─────────────────────────────────────────
#
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

#
# ───── Helper Functions ────────────────────────────────────────────────────────
#

def nationality_group(row):
    nat = str(row.get("Nationality (Country Name)", "")).strip().lower()
    pr  = str(row.get("PR", "")).strip().lower()
    if nat == "singapore":
        return 1
    elif pr in ("yes", "pr"):
        return 2
    elif nat == "malaysia":
        return 3
    elif nat == "india":
        return 4
    else:
        return 5

def split_name(full_name):
    s = str(full_name).strip()
    if " " in s:
        i = s.find(" ")
        return pd.Series([s[:i], s[i+1:]])
    return pd.Series([s, ""])

def clean_gender(g):
    v = str(g).strip().upper()
    if v == "M": return "Male"
    if v == "F": return "Female"
    if v in ("MALE","FEMALE"): return v.title()
    return v

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    # 1) rename
    df.columns = [
        "S/N","Vehicle Plate Number","Company Full Name","Full Name As Per NRIC",
        "First Name as per NRIC","Middle and Last Name as per NRIC","Identification Type",
        "IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date",
        "Nationality (Country Name)","PR","Gender","Mobile Number",
    ]
    # 2) drop empty
    df = df.dropna(subset=df.columns[3:13], how="all")
    # 3) sort
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df = df.sort_values(
        ["Company Full Name","SortGroup","Nationality (Country Name)","Full Name As Per NRIC"],
        ignore_index=True
    ).drop(columns="SortGroup")
    # 4) reset S/N
    df["S/N"] = range(1, len(df)+1)
    # 5) plate cleanup
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
          .str.replace(r"[\/,]", ";", regex=True)
          .str.replace(r"\s*;\s*", ";", regex=True)
          .str.strip()
          .replace("nan","")
    )
    # 6) name casing + split
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = (
        df["Full Name As Per NRIC"].apply(split_name)
    )
    # 7) nationality map + title
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .replace({"Chinese":"China","Singaporean":"Singapore"})
          .astype(str).str.title()
    )
    # 8) swap IC/work-permit if hyphen in IC column
    iccol = "IC (Last 3 digits and suffix) 123A"
    wpcol = "Work Permit Expiry Date"
    if df[iccol].astype(str).str.contains("-", na=False).any():
        df[[iccol,wpcol]] = df[[wpcol,iccol]]
    # 9) trim IC suffix
    df[iccol] = df[iccol].astype(str).str[-4:]
    # 10) mobile → digits only
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(r"\D","",regex=True)
    # 11) gender
    df["Gender"] = df["Gender"].apply(clean_gender)
    # 12) format WP date
    df[wpcol] = pd.to_datetime(df[wpcol], errors="coerce").dt.strftime("%Y-%m-%d")
    return df

def generate_visitor_only(df: pd.DataFrame) -> BytesIO:
    """
    Builds a single-sheet Excel (Visitor List only), applies styling
    and adds Vehicles + Total Visitors summaries.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        wb = writer.book
        ws = writer.sheets["Visitor List"]

        # styling objects
        header_fill = PatternFill("solid", fgColor="94B455")
        warning_fill = PatternFill("solid", fgColor="FFCCCC")
        border = Border(
            left=Side("thin"), right=Side("thin"),
            top=Side("thin"), bottom=Side("thin")
        )
        center = Alignment("center","center")
        normal_font = Font(name="Calibri", size=11)
        bold_font   = Font(name="Calibri", size=11, bold=True)

        # 1) all cells border/alignment/font
        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = center
                cell.font      = normal_font

        # 2) header row
        for col in range(1, ws.max_column+1):
            cell = ws[f"{get_column_letter(col)}1"]
            cell.fill = header_fill
            cell.font = bold_font

        # 3) freeze top
        ws.freeze_panes = ws["A2"]

        # 4) highlight mismatches
        mismatches = 0
        for r in range(2, ws.max_row+1):
            idt = str(ws[f"G{r}"].value).strip().upper()
            nat = str(ws[f"J{r}"].value).strip().title()
            pr  = str(ws[f"K{r}"].value).strip().title()
            bad = False
            if idt=="NRIC" and not (nat=="Singapore" or (nat!="Singapore" and pr in ("Yes","Pr"))):
                bad = True
            if idt=="FIN" and (nat=="Singapore" or pr in ("Yes","Pr")):
                bad = True
            if bad:
                for col in ("G","J","K"):
                    ws[f"{col}{r}"].fill = warning_fill
                mismatches += 1

        if mismatches:
            st.warning(f"⚠️ {mismatches} potential mismatch(es) found. Check highlighted rows.")

        # 5) autofit cols + set row height
        for col in ws.columns:
            width = max(len(str(cell.value)) for cell in col if cell.value)
            ws.column_dimensions[get_column_letter(col[0].column)].width = width+2
        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 20

        # 6) Vehicles summary
        plates = []
        for v in df["Vehicle Plate Number"].dropna():
            plates += [x.strip() for x in str(v).split(";") if x.strip()]
        out_row = ws.max_row + 2
        if plates:
            ws[f"B{out_row}"].value = "Vehicles"
            ws[f"B{out_row}"].border = border
            ws[f"B{out_row}"].alignment = center
            ws[f"B{out_row+1}"].value = ";".join(sorted(set(plates)))
            ws[f"B{out_row+1}"].border = border
            ws[f"B{out_row+1}"].alignment = center
            out_row += 2

        # 7) Total Visitors
        ws[f"B{out_row}"].value = "Total Visitors"
        ws[f"B{out_row}"].border = border
        ws[f"B{out_row}"].alignment = center
        ws[f"B{out_row+1}"].value = df["Company Full Name"].notna().sum()
        ws[f"B{out_row+1}"].border = border
        ws[f"B{out_row+1}"].alignment = center

    return output

#
# ───── Streamlit UI ─────────────────────────────────────────────────────────────
#
uploaded = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])
if uploaded:
    # only read Visitor List sheet
    raw = pd.read_excel(uploaded, sheet_name="Visitor List")
    clean = clean_data(raw)
    buf   = generate_visitor_only(clean)
    fname = f"VisitorList_Cleaned_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        label="📥 Download Cleaned Visitor List Only",
        data=buf.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
