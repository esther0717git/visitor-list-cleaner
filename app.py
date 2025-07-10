import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

# ───── Streamlit setup ────────────────────────────────────────────────────────
st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🇸🇬 CLARITY GATE - VISITOR DATA CLEANING & VALIDATION 🫧")

# ───── Download Sample Template ───────────────────────────────────────────────
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ───── Helper functions ────────────────────────────────────────────────────────

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
    if v == "M":
        return "Male"
    if v == "F":
        return "Female"
    if v in ("MALE","FEMALE"):
        return v.title()
    return v

# ───── Core Cleaning Logic ────────────────────────────────────────────────────

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    # 1) Trim to exactly 13 cols then rename
    df = df.iloc[:, :13]
    df.columns = [
        "S/N",
        "Vehicle Plate Number",
        "Company Full Name",
        "Full Name As Per NRIC",
        "First Name as per NRIC",
        "Middle and Last Name as per NRIC",
        "Identification Type",
        "IC (Last 3 digits and suffix) 123A",
        "Work Permit Expiry Date",
        "Nationality (Country Name)",
        "PR",
        "Gender",
        "Mobile Number",
    ]

    # 2) Drop rows where all of D–M are blank
    df = df.dropna(subset=df.columns[3:13], how="all")

    # 3) Normalize nationality (incl. Indian → India)
    nat_map = {
        "chinese":     "China",
        "singaporean": "Singapore",
        "malaysian":   "Malaysia",
        "indian":      "India",
    }
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .astype(str)
          .str.strip()
          .str.lower()
          .replace(nat_map, regex=False)
          .str.title()
    )

    # 4) Sort by Company → nat-group → Country → Full Name
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df = (
        df.sort_values(
            ["Company Full Name","SortGroup","Nationality (Country Name)","Full Name As Per NRIC"],
            ignore_index=True,
        )
        .drop(columns="SortGroup")
    )

    # 5) Reset S/N
    df["S/N"] = range(1, len(df) + 1)

    # 6) Standardize Vehicle Plate Number
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"]
          .astype(str)
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

    # 8) Swap IC vs WP if reversed
    iccol, wpcol = "IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date"
    if df[iccol].astype(str).str.contains("-", na=False).any():
        df[[iccol, wpcol]] = df[[wpcol, iccol]]

    # 9) Trim IC suffix
    df[iccol] = df[iccol].astype(str).str[-4:]

# 10) Fix Mobile Number back to the original 8 digits—
    def fix_mobile(x):
        d = re.sub(r"\D", "", str(x))
        # if too long...
        if len(d) > 8:
            extra = len(d) - 8
            # if the extras are just decimal zeros, strip from the right
            if d.endswith("0" * extra):
                d = d[:-extra]
            else:
                # otherwise assume it's a country code and drop from the left
                d = d[-8:]
        # if too short, left-pad with zeros
        if len(d) < 8:
            d = d.zfill(8)
        return d

    df["Mobile Number"] = df["Mobile Number"].apply(fix_mobile)

    # 11) Normalize gender
    df["Gender"] = df["Gender"].apply(clean_gender)

    # 12) Format Work Permit Expiry Date → YYYY-MM-DD
    df[wpcol] = pd.to_datetime(df[wpcol], errors="coerce").dt.strftime("%Y-%m-%d")

    return df

# ───── Build & style the single sheet Excel ────────────────────────────────────

def generate_visitor_only(df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        ws = writer.sheets["Visitor List"]

        # styling objects
        header_fill  = PatternFill("solid", fgColor="94B455")
        warning_fill = PatternFill("solid", fgColor="FFCCCC")
        border       = Border(Side("thin"),Side("thin"),Side("thin"),Side("thin"))
        center       = Alignment("center","center")
        normal_font  = Font(name="Calibri", size=9)
        bold_font    = Font(name="Calibri", size=9, bold=True)

        # 1) Apply borders, alignment, font
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

        # 4) Validation & highlight errors
        errors = 0
        for r in range(2, ws.max_row + 1):
            idt = str(ws[f"G{r}"].value).strip().upper()
            nat = str(ws[f"J{r}"].value).strip().title()
            pr  = str(ws[f"K{r}"].value).strip().lower()
            bad = False

            # → only NRIC may have PR=yes
            if idt != "NRIC" and pr in ("yes","y","pr"):
                bad = True

            # → FIN must not be Singapore or carry PR
            if idt == "FIN" and (nat == "Singapore" or pr in ("yes","y","pr")):
                bad = True

            # → NRIC must be either Singapore or foreign+PR
            if idt == "NRIC" and not (nat == "Singapore" or pr in ("yes","y","pr")):
                bad = True

            if bad:
                for c in ("G","J","K"):
                    ws[f"{c}{r}"].fill = warning_fill
                errors += 1

        if errors:
            st.warning(f"⚠️ {errors} validation error(s) found.")

        # 5) Auto-fit columns & set row height
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

    buf.seek(0)
    return buf

# ───── Streamlit UI: Upload & Download ────────────────────────────────────────
uploaded = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])
if uploaded:
    # 1) Read the "Visitor List" sheet
    raw_df = pd.read_excel(uploaded, sheet_name="Visitor List")

    # 2) Capture Company Name in cell C2 (excel row 2, column C → pandas row 0, col 2)
    company_cell = raw_df.iloc[0, 2]
    company = (
        str(company_cell).strip()
        if pd.notna(company_cell) and str(company_cell).strip()
        else "VisitorList"
    )

    # 3) Clean & generate output
    cleaned = clean_data(raw_df)
    out_buf = generate_visitor_only(cleaned)

    # 4) Build filename: CompanyName_YYYYMMDD.xlsx in Asia/Singapore time
    today = datetime.now(ZoneInfo("Asia/Singapore")).strftime("%Y%m%d")
    fname = f"{company}_{today}.xlsx"

    # 5) Serve download
    st.download_button(
        label="📥 Download Cleaned Visitor List",
        data=out_buf.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
