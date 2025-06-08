import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# ───── Download Sample Template ───────────────────────────────────────────────
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ───── Helpers ─────────────────────────────────────────────────────────────────

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
    if v=="M":      return "Male"
    if v=="F":      return "Female"
    if v in ("MALE","FEMALE"): return v.title()
    return v

# ───── Cleaning Logic ─────────────────────────────────────────────────────────

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    # 1) Rename to the 13 columns we expect
    df.columns = [
        "S/N","Vehicle Plate Number","Company Full Name","Full Name As Per NRIC",
        "First Name as per NRIC","Middle and Last Name as per NRIC","Identification Type",
        "IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date",
        "Nationality (Country Name)","PR","Gender","Mobile Number",
    ]

    # 2) Drop rows where all of D–M are empty
    df = df.dropna(subset=df.columns[3:13], how="all")

    # 3) Normalize & title-case Nationality; map common variants
    nat_map = {"chinese":"China","singaporean":"Singapore","malaysian":"Malaysia"}
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .astype(str)
          .str.strip()
          .replace(nat_map, regex=False)
          .str.title()
    )

    # 4) Sort by company → nationality-group → country → name
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df = (
        df.sort_values(
            ["Company Full Name","SortGroup","Nationality (Country Name)","Full Name As Per NRIC"],
            ignore_index=True
        )
        .drop(columns="SortGroup")
    )

    # 5) Reset serial #
    df["S/N"] = range(1, len(df)+1)

    # 6) Standardize vehicle plates
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
          .str.replace(r"[\/,]",";",regex=True)
          .str.replace(r"\s*;\s*",";",regex=True)
          .str.strip()
          .replace("nan","",regex=False)
    )

    # 7) Title-case full name + split into first / rest
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = (
        df["Full Name As Per NRIC"].apply(split_name)
    )

    # 8) Detect & swap IC vs WorkPermit if mis-entered
    ic_col = "IC (Last 3 digits and suffix) 123A"
    wp_col = "Work Permit Expiry Date"
    if df[ic_col].astype(str).str.contains("-", na=False).any():
        df[[ic_col,wp_col]] = df[[wp_col,ic_col]]

    # 9) Keep only last 4 chars of IC suffix
    df[ic_col] = df[ic_col].astype(str).str[-4:]

    # 10) Clean mobile to digits only
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(r"\D","",regex=True)

    # 11) Normalize gender
    df["Gender"] = df["Gender"].apply(clean_gender)

    # 12) Format Work Permit date
    df[wp_col] = pd.to_datetime(df[wp_col], errors="coerce").dt.strftime("%Y-%m-%d")

    return df

# ───── Build a single-sheet Excel ──────────────────────────────────────────────

def generate_visitor_only(df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        wb = writer.book
        ws = writer.sheets["Visitor List"]

        # style definitions
        header_fill = PatternFill("solid", fgColor="94B455")
        warn_fill   = PatternFill("solid", fgColor="FFCCCC")
        border      = Border(Side("thin"),Side("thin"),Side("thin"),Side("thin"))
        center      = Alignment("center","center")
        normal_font = Font("Calibri",11)
        bold_font   = Font("Calibri",11,bold=True)

        # 1) apply border/alignment/font to all cells
        for row in ws.iter_rows():
            for cell in row:
                cell.border    = border
                cell.alignment = center
                cell.font      = normal_font

        # 2) style header row
        for c in range(1, ws.max_column+1):
            h = ws[f"{get_column_letter(c)}1"]
            h.fill = header_fill
            h.font = bold_font

        # 3) freeze pane
        ws.freeze_panes = ws["A2"]

        # ───── Validations & Highlights ────────────────────────────────

        nationality_allowed = {"Singapore","India","Thailand","Malaysia","China"}  # add as needed
        id_errors = 0
        nat_errors = 0
        dup_errors = 0

        # find duplicates by (Company, FullName)
        dup_mask = df.duplicated(subset=["Company Full Name","Full Name As Per NRIC"], keep=False)

        for r in range(2, ws.max_row+1):
            idt = str(ws[f"G{r}"].value).strip().upper()
            pr  = str(ws[f"K{r}"].value).strip().lower()
            nat = str(ws[f"J{r}"].value).strip().title()

            # 1) ID-type logic
            bad_id = False
            if idt in ("NRIC","PR"):
                # must be Singapore
                if nat != "Singapore":
                    bad_id = True
            elif idt == "FIN":
                # must NOT be Singapore
                if nat == "Singapore":
                    bad_id = True
            elif idt == "WORK PERMIT":
                # permit date must exist
                if not ws[f"I{r}"].value:
                    bad_id = True
            else:  # OTHERS
                # just require nationality non-blank
                if nat=="": bad_id = True

            if bad_id:
                # highlight G (ID) and J (Nation)
                ws[f"G{r}"].fill = warn_fill
                ws[f"J{r}"].fill = warn_fill
                id_errors += 1

            # 2) PR column logic: only NRIC/PR may have Yes/Y
            if idt not in ("NRIC","PR") and pr in ("yes","y"):
                ws[f"K{r}"].fill = warn_fill
                id_errors += 1

            # 3) Nationality‐required logic
            if nat not in nationality_allowed:
                ws[f"J{r}"].fill = warn_fill
                nat_errors += 1

            # 4) Duplicates (by Company + FullName)
            if dup_mask[r-2]:
                # highlight the whole row lightly
                for c in range(1, ws.max_column+1):
                    ws[f"{get_column_letter(c)}{r}"].fill = warn_fill
                dup_errors += 1

        # 4) show warnings
        if id_errors:
            st.warning(f"⚠️ {id_errors} ID/PR validation error(s) found.")
        if nat_errors:
            st.warning(f"⚠️ {nat_errors} invalid Nationality entry(ies) found.")
        if dup_errors:
            st.warning(f"⚠️ {dup_errors} duplicate visitor name(s) found.")

        # 5) autosize + row height
        for col in ws.columns:
            w = max(len(str(c.value)) for c in col if c.value)
            ws.column_dimensions[get_column_letter(col[0].column)].width = w+2
        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 20

    buf.seek(0)
    return buf

# ───── Streamlit UI ────────────────────────────────────────────────────────────

uploaded = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])
if uploaded:
    raw_df = pd.read_excel(uploaded, sheet_name="Visitor List")
    clean_df = clean_data(raw_df)
    out_buf   = generate_visitor_only(clean_df)
    fname     = f"Cleaned_VisitorList_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        label="📥 Download Cleaned Visitor List Only",
        data=out_buf.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
