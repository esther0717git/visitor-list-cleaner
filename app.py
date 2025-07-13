import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

# ───── Streamlit setup ────────────────────────────────────────────────────────
st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🇸🇬 CLARITY GATE – VISITOR DATA CLEANING & VALIDATION 🫧")

# ───── 1) Info Banner: Data Integrity Foundation ───────────────────────────────
st.info(
    """
    **Data Integrity Is Our Foundation**  
    At every step—from file upload to final report—we enforce strict validation to guarantee your visitor data is accurate, complete, and compliant.  
    Maintaining integrity not only expedites gate clearance, it protects our facilities and ensures we meet all regulatory requirements.
    """
)

# ───── 2) Dedicated “Why Data Integrity?” Section ─────────────────────────────
with st.expander("Why is Data Integrity Important?"):
    st.write(
        """
        - **Accuracy**: Correct visitor details reduce clearance delays.  
        - **Security**: Reliable ID checks prevent unauthorized access.  
        - **Compliance**: Audit-ready records ensure regulatory adherence.  
        - **Efficiency**: Trustworthy data powers faster reporting and analytics.
        """
    )

# ───── 3) Inline Callout Above Uploader ───────────────────────────────────────
st.markdown("### ⚠️ **Please ensure your spreadsheet has no missing or malformed fields.**")
uploaded = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])

# ───── Estimate Clearance Date (below uploader) ──────────────────────────────
now = datetime.now(ZoneInfo("Asia/Singapore"))
formatted_now = now.strftime("%A %d %B, %I:%M%p").lstrip("0")
st.markdown("### 📦 Estimate Clearance Date")
st.write(f"**Today is:** {formatted_now}")

if st.button("▶️ Calculate Estimated Delivery"):
    # 1) Determine submission date (bump if 3 PM+)
    sub_date = now.date()
    if now.hour >= 15:
        sub_date += timedelta(days=1)

    # 2) Count two working days, skipping weekends
    days_added = 0
    current = sub_date
    while days_added < 2:
        current += timedelta(days=1)
        if current.weekday() < 5:  # Mon=0 … Fri=4
            days_added += 1

    # 3) Clearance = the day *after* the 2nd working day
    clearance_date = current + timedelta(days=1)
    #    and if that lands on Sat/Sun, push to Monday
    while clearance_date.weekday() >= 5:
        clearance_date += timedelta(days=1)

    # Format as "Sunday 13 July" (no leading zero on day)
    formatted_clearance = f"{clearance_date:%A} {clearance_date.day} {clearance_date:%B}"
    st.success(f"✓ Earliest clearance: **{formatted_clearance}**")

# ───── Helper functions ────────────────────────────────────────────────────────

def nationality_group(row):
    nat = str(row["Nationality (Country Name)"]).strip().lower()
    pr  = str(row["PR"]).strip().lower()
    if nat == "singapore":
        return 1
    elif pr in ("yes","y","pr"):
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
    df = df.iloc[:, :13]
    df.columns = [
        "S/N","Vehicle Plate Number","Company Full Name","Full Name As Per NRIC",
        "First Name as per NRIC","Middle and Last Name as per NRIC","Identification Type",
        "IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date",
        "Nationality (Country Name)","PR","Gender","Mobile Number",
    ]
    df = df.dropna(subset=df.columns[3:13], how="all")

    nat_map = {"chinese":"China","singaporean":"Singapore","malaysian":"Malaysia","indian":"India"}
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .astype(str).str.strip().str.lower()
          .replace(nat_map, regex=False).str.title()
    )

    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df = (
        df.sort_values(
            ["Company Full Name","SortGroup","Nationality (Country Name)","Full Name As Per NRIC"],
            ignore_index=True
        )
        .drop(columns="SortGroup")
    )
    df["S/N"] = range(1, len(df) + 1)

    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"]
          .astype(str)
          .str.replace(r"[\/,]", ";", regex=True)
          .str.replace(r"\s*;\s*", ";", regex=True)
          .str.strip()
          .replace("nan","", regex=False)
    )

    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = (
        df["Full Name As Per NRIC"].apply(split_name)
    )

    iccol, wpcol = "IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date"
    if df[iccol].astype(str).str.contains("-", na=False).any():
        df[[iccol, wpcol]] = df[[wpcol, iccol]]
    df[iccol] = df[iccol].astype(str).str[-4:]

    def fix_mobile(x):
        d = re.sub(r"\D", "", str(x))
        if len(d) > 8:
            extra = len(d) - 8
            if d.endswith("0"*extra): d = d[:-extra]
            else: d = d[-8:]
        if len(d) < 8: d = d.zfill(8)
        return d

    df["Mobile Number"] = df["Mobile Number"].apply(fix_mobile)
    df["Gender"] = df["Gender"].apply(clean_gender)
    df[wpcol] = pd.to_datetime(df[wpcol], errors="coerce").dt.strftime("%Y-%m-%d")
    return df

def generate_visitor_only(df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        ws = writer.sheets["Visitor List"]

        header_fill  = PatternFill("solid", fgColor="94B455")
        warning_fill = PatternFill("solid", fgColor="FFCCCC")
        border       = Border(Side("thin"),Side("thin"),Side("thin"),Side("thin"))
        center       = Alignment("center","center")
        normal_font  = Font(name="Calibri", size=9)
        bold_font    = Font(name="Calibri", size=9, bold=True)

        for row in ws.iter_rows():
            for cell in row:
                cell.border    = border
                cell.alignment = center
                cell.font      = normal_font

        for col in range(1, ws.max_column + 1):
            h = ws[f"{get_column_letter(col)}1"]
            h.fill = header_fill
            h.font = bold_font

        ws.freeze_panes = ws["A2"]

        errors = 0
        for r in range(2, ws.max_row + 1):
            idt = str(ws[f"G{r}"].value).strip().upper()
            nat = str(ws[f"J{r}"].value).strip().title()
            pr  = str(ws[f"K{r}"].value).strip().lower()
            bad = False

            if idt != "NRIC" and pr in ("yes","y","pr"): bad = True
            if idt == "FIN" and (nat=="Singapore" or pr in ("yes","y","pr")): bad = True
            if idt == "NRIC" and not (nat=="Singapore" or pr in ("yes","y","pr")): bad = True

            if bad:
                for c in ("G","J","K"):
                    ws[f"{c}{r}"].fill = warning_fill
                errors += 1

        if errors:
            st.warning(f"⚠️ {errors} validation error(s) found.")

        for col in ws.columns:
            w = max(len(str(cell.value)) for cell in col if cell.value)
            ws.column_dimensions[get_column_letter(col[0].column)].width = w + 2
        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 20

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

        ws[f"B{ins}"].value     = "Total Visitors"
        ws[f"B{ins}"].border    = border
        ws[f"B{ins}"].alignment = center
        ws[f"B{ins+1}"].value   = df["Company Full Name"].notna().sum()
        ws[f"B{ins+1}"].border  = border
        ws[f"B{ins+1}"].alignment = center

    buf.seek(0)
    return buf

# ───── Read, Clean & Download ────────────────────────────────────────────────
if uploaded:
    raw_df = pd.read_excel(uploaded, sheet_name="Visitor List")
    company_cell = raw_df.iloc[0, 2]
    company = (
        str(company_cell).strip()
        if pd.notna(company_cell) and str(company_cell).strip()
        else "VisitorList"
    )

    cleaned = clean_data(raw_df)
    out_buf = generate_visitor_only(cleaned)

    today_str = datetime.now(ZoneInfo("Asia/Singapore")).strftime("%Y%m%d")
    fname = f"{company}_{today_str}.xlsx"

    st.download_button(
        label="📥 Download Cleaned Visitor List",
        data=out_buf.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ───── 4) Footer Reminder After Download ─────────────────────────────────
    st.caption(
        "✅ Your data has been validated. Please double-check critical fields before sharing with security teams."
    )
