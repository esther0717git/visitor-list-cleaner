import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

# ───── Streamlit page setup ───────────────────────────────────────────────────
st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# ───── Download Sample Template Button ───────────────────────────────────────
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ───── Helper Functions ───────────────────────────────────────────────────────

def nationality_group(row):
    nat = str(row["Nationality (Country Name)"]).strip().lower()
    pr  = str(row["PR"]).strip().lower()
    if nat == "singapore":
        return 1
    elif pr in ("yes", "y", "pr"):
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

# ───── Core Cleaning Logic ────────────────────────────────────────────────────

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    # 1) Rename to exactly these columns
    df.columns = [
        "S/N","Vehicle Plate Number","Company Full Name","Full Name As Per NRIC",
        "First Name as per NRIC","Middle and Last Name as per NRIC","Identification Type",
        "IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date",
        "Nationality (Country Name)","PR","Gender","Mobile Number",
    ]

    # 2) Drop rows where all of D–M are blank
    df = df.dropna(subset=df.columns[3:13], how="all")

    # 3) Normalize nationality (including “Indian”→“India”)
    nat_map = {
        "chinese": "China",
        "singaporean": "Singapore",
        "malaysian": "Malaysia",
        "indian": "India",            # ← added
    }
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .astype(str).str.strip().str.lower()
          .replace(nat_map, regex=False)
          .str.title()
    )

    # 4) Sort by Company → nationality group → Country → Full Name
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df = (
        df.sort_values(
            ["Company Full Name","SortGroup","Nationality (Country Name)","Full Name As Per NRIC"],
            ignore_index=True
        )
        .drop(columns="SortGroup")
    )

    # 5) Reset S/N
    df["S/N"] = range(1, len(df) + 1)

    # 6) Standardize Vehicle Plate Number
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
          .str.replace(r"[\/,]", ";", regex=True)
          .str.replace(r"\s*;\s*", ";", regex=True)
          .str.strip()
          .replace("nan","", regex=False)
    )

    # 7) Proper-case & split names
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = (
        df["Full Name As Per NRIC"].apply(split_name)
    )

    # 8) Swap IC vs WP if mis-entered
    iccol, wpcol = "IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date"
    if df[iccol].astype(str).str.contains("-", na=False).any():
        df[[iccol, wpcol]] = df[[wpcol, iccol]]

    # 9) Trim IC suffix
    df[iccol] = df[iccol].astype(str).str[-4:]

    # 10) Clean mobile to digits only
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(r"\D", "", regex=True)

    # 11) Normalize gender
    df["Gender"] = df["Gender"].apply(clean_gender)

    # 12) Format Work Permit dates to YYYY-MM-DD
    df[wpcol] = pd.to_datetime(df[wpcol], errors="coerce").dt.strftime("%Y-%m-%d")

    return df

# ───── Generate Single-Sheet Excel ────────────────────────────────────────────

def generate_visitor_only(df: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        ws = writer.sheets["Visitor List"]

        # Styling definitions
        header_fill  = PatternFill("solid", fgColor="94B455")
        warning_fill = PatternFill("solid", fgColor="FFCCCC")
        border       = Border(Side("thin"),Side("thin"),Side("thin"),Side("thin"))
        center       = Alignment("center","center")
        normal_font  = Font(name="Calibri", size=11)
        bold_font    = Font(name="Calibri", size=11, bold=True)

        # 1) Apply border/alignment/font to all cells
        for row in ws.iter_rows():
            for cell in row:
                cell.border    = border
                cell.alignment = center
                cell.font      = normal_font

        # 2) Style header row
        for col in range(1, ws.max_column + 1):
            h = ws[f"{get_column_letter(col)}1"]
            h.fill = header_fill
            h.font = bold_font

        # 3) Freeze top row
        ws.freeze_panes = ws["A2"]

        # 4) Highlight ID+PR rules:
        #    • Only NRIC may have PR=Yes/Y
        #    • FIN rows must not be Singaporean or have PR
        errors = 0
        for r in range(2, ws.max_row + 1):
            idt = str(ws[f"G{r}"].value).strip().upper()
            nat = str(ws[f"J{r}"].value).strip().title()
            pr  = str(ws[f"K{r}"].value).strip().lower()

            bad = False

            # only NRIC may have PR=yes/y
            if idt != "NRIC" and pr in ("yes","y","pr"):
                bad = True

            # FIN must not be Singaporean or carry PR
            if idt == "FIN" and (nat == "Singapore" or pr in ("yes","y","pr")):
                bad = True

            if bad:
                for c in ("G","J","K"):
                    ws[f"{c}{r}"].fill = warning_fill
                errors += 1

        if errors:
            st.warning(f"⚠️ {errors} validation error(s) found.")

        # 5) Auto-fit & fixed row height
        for col in ws.columns:
            w = max(len(str(cell.value)) for cell in col if cell.value)
            ws.column_dimensions[get_column_letter(col[0].column)].width = w + 2
        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 20

        # 6) Vehicles summary
        plates = []
        for v in df["Vehicle Plate Number"].dropna():
            plates += [x.strip() for x in str(v).split(";") if x.strip()]
        ins = ws.max_row + 2
        if plates:
            ws[f"B{ins}"].value     = "Vehicles"
            ws[f"B{ins}"].border    = border
            ws[f"B{ins}"].alignment = center
            ws[f"B{ins+1}"].value   = ";".join(sorted(set(plates)))
            ws[f"B{ins+1}"].border  = border
            ws[f"B{ins+1}"].alignment = center
            ins += 2

        # 7) Total Visitors
        ws[f"B{ins}"].value     = "Total Visitors"
        ws[f"B{ins}"].border    = border
        ws[f"B{ins}"].alignment = center
        ws[f"B{ins+1}"].value   = df["Company Full Name"].notna().sum()
        ws[f"B{ins+1}"].border  = border
        ws[f"B{ins+1}"].alignment = center

    output.seek(0)
    return output

# ───── Streamlit UI: Upload & Download ────────────────────────────────────────
uploaded = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])
if uploaded:
    raw_df  = pd.read_excel(uploaded, sheet_name="Visitor List")
    cleaned = clean_data(raw_df)
    buf     = generate_visitor_only(cleaned)
    fname   = f"Cleaned_VisitorList_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        label="📥 Download Cleaned Visitor List Only",
        data=buf.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
