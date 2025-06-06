import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
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
    """
    Assigns a sort‐group integer based on:
      1 = Singapore
      2 = Malaysia (PR = Yes or 'PR')
      3 = Malaysia (non-PR)
      4 = India
      5 = All other nationalities
    """
    nationality = str(row.get("Nationality (Country Name)")).lower()
    pr_status = str(row.get("PR")).strip().lower()

    if nationality == "singapore":
        return 1
    elif pr_status in ["yes", "pr"]:
        return 2
    elif nationality == "malaysia":
        return 3
    elif nationality == "india":
        return 4
    else:
        return 5


def split_name(name):
    """
    Splits a single “Full Name” (proper-cased) into [First, MiddleAndLast].
    If no space found, returns [name, ""]. 
    """
    name = str(name).strip()
    if " " in name:
        first_space = name.find(" ")
        return pd.Series([name[:first_space], name[first_space + 1 :]])
    return pd.Series([name, ""])


def clean_gender(val):
    """
    Normalizes gender values:
      - "M"    → "Male"
      - "F"    → "Female"
      - "MALE" → "Male"
      - "FEMALE" → "Female"
      - Others remain unchanged
    """
    val = str(val).strip().upper()
    if val == "M":
        return "Male"
    elif val == "F":
        return "Female"
    elif val in ["MALE", "FEMALE"]:
        return val.title()
    return val


def clean_data(df):
    """
    Takes the raw Visitor List DataFrame, renames columns, drops empty rows,
    sorts, and applies all cleaning rules:
      • Proper-case Full Name
      • Split Full Name into First / Middle+Last
      • Nationality mapping + proper-case
      • Swap IC suffix vs. Work Permit Date if user reversed them
      • Extract last 4 chars of IC suffix
      • Force mobile to digits only
      • Normalize Gender
      • Format dates to YYYY-MM-DD
      • Reassign running “S/N”
    """
    # 1) Rename columns exactly as expected
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

    # 2) Drop any row where D–M are all NaN (i.e., columns [3:13])
    df = df.dropna(subset=df.columns[3:13], how="all")

    # 3) Sort by company, then nationality/PR group, then country, then full name
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df.sort_values(
        by=[
            "Company Full Name",
            "SortGroup",
            "Nationality (Country Name)",
            "Full Name As Per NRIC",
        ],
        inplace=True,
    )
    df.drop(columns=["SortGroup"], inplace=True)

    # 4) Reset Serial # (1..N)
    df["S/N"] = range(1, len(df) + 1)

    # 5) Clean Vehicle Plate Number → replace “/” or “,” with “;”, trim spaces
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"]
        .astype(str)
        .str.replace(r"[\/\,]", ";", regex=True)
        .str.replace(r"\s*;\s*", ";", regex=True)
        .str.strip()
        .replace("nan", "", regex=False)
    )

    # 6) Proper-case full name (D), then split into E/F
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC", "Middle and Last Name as per NRIC"]] = df[
        "Full Name As Per NRIC"
    ].apply(split_name)

    # 7) Nationality transform: map “Chinese”→“China”, “Singaporean”→“Singapore”
    nationality_map = {"Chinese": "China", "Singaporean": "Singapore"}
    df["Nationality (Country Name)"] = df["Nationality (Country Name)"].replace(
        nationality_map
    )
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"].astype(str).str.title()
    )

    # 8) If IC and WorkPermit columns were reversed (look for a “-” in IC column),
    #    then swap them back:
    if df["IC (Last 3 digits and suffix) 123A"].astype(str).str.contains("-", na=False).any():
        df[
            ["IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date"]
        ] = df[["Work Permit Expiry Date", "IC (Last 3 digits and suffix) 123A"]]

    # 9) Extract only the last 4 chars of IC suffix
    df["IC (Last 3 digits and suffix) 123A"] = df[
        "IC (Last 3 digits and suffix) 123A"
    ].astype(str).str[-4:]

    # 10) Force Mobile Number → digits only (drop decimals / non-digits)
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(
        r"\D", "", regex=True
    )

    # 11) Clean Gender
    df["Gender"] = df["Gender"].apply(clean_gender)

    # 12) Standardize Work Permit Expiry Date → YYYY-MM-DD (if parseable)
    df["Work Permit Expiry Date"] = (
        pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce")
        .dt.strftime("%Y-%m-%d")
    )

    return df


def generate_excel(xlsx, df):
    """
    Rebuilds a new Excel file:
      • Writes back every original sheet except Visitor List untouched
      • Overwrites Visitor List with our cleaned df + styling + summary rows
      • Returns (BytesIO, mismatch_count)
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Write all sheets except “Visitor List” → untouched
        for sheet_name, sheet_df in xlsx.items():
            if sheet_name != "Visitor List":
                sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)

        # 2) Overwrite “Visitor List” with cleaned DataFrame
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        workbook = writer.book
        worksheet = writer.sheets["Visitor List"]

        # 3) Define styling objects
        header_fill = PatternFill(start_color="94B455", end_color="94B455", fill_type="solid")
        light_red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )
        center_align = Alignment(horizontal="center", vertical="center")
        font_style = Font(name="Calibri", size=11)
        bold_font = Font(name="Calibri", size=11, bold=True)

        # 4) Apply border/alignment/font to every cell in Visitor List
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = center_align
                cell.font = font_style

        # 5) Style Header Row (row 1)
        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet[f"{get_column_letter(col_idx)}1"]
            cell.fill = header_fill
            cell.font = bold_font

        # 6) Freeze the top row
        worksheet.freeze_panes = worksheet["A2"]

        mismatch_count = 0
        warning_rows = []

        # 7) Highlight any ID vs. Nationality/PR mismatches in light red
        for row_idx in range(2, worksheet.max_row + 1):
            id_type = str(worksheet[f"G{row_idx}"].value).strip().upper()
            nationality = str(worksheet[f"J{row_idx}"].value).strip().title()
            pr_status = str(worksheet[f"K{row_idx}"].value).strip().title()

            highlight = False
            # 7a) If ID = “NRIC” but not (Sing or PR)
            if id_type == "NRIC" and not (
                nationality == "Singapore" or (nationality != "Singapore" and pr_status in ["Yes", "Pr"])
            ):
                highlight = True
            # 7b) If ID = “FIN” but nationality is Singapore or PR
            if id_type == "FIN" and (nationality == "Singapore" or pr_status in ["Yes", "Pr"]):
                highlight = True

            if highlight:
                warning_rows.append(row_idx)
                worksheet[f"G{row_idx}"].fill = light_red_fill
                worksheet[f"J{row_idx}"].fill = light_red_fill
                worksheet[f"K{row_idx}"].fill = light_red_fill
                mismatch_count += 1

        # 8) Auto‐fit column widths
        for col in worksheet.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            worksheet.column_dimensions[col_letter].width = max_length + 2

        # 9) Set a fixed row height (20) for each row
        for row in worksheet.iter_rows():
            worksheet.row_dimensions[row[0].row].height = 20

        # 10) Build “Vehicles” summary below all data (if any)
        vehicles = []
        for val in df["Vehicle Plate Number"].dropna():
            vehicles.extend([v.strip() for v in str(val).split(";") if v.strip()])

        insert_row = worksheet.max_row + 2
        if vehicles:
            summary = ";".join(sorted(set(vehicles)))
            worksheet[f"B{insert_row}"].value = "Vehicles"
            worksheet[f"B{insert_row}"].border = border
            worksheet[f"B{insert_row}"].alignment = center_align
            worksheet[f"B{insert_row + 1}"].value = summary
            worksheet[f"B{insert_row + 1}"].border = border
            worksheet[f"B{insert_row + 1}"].alignment = center_align
            insert_row += 3

        # 11) Add “Total Visitors” (count of non-blank Company Names)
        total_visitors = df["Company Full Name"].notna().sum()
        worksheet[f"B{insert_row}"].value = "Total Visitors"
        worksheet[f"B{insert_row}"].alignment = center_align
        worksheet[f"B{insert_row}"].border = border
        worksheet[f"B{insert_row + 1}"].value = total_visitors
        worksheet[f"B{insert_row + 1}"].alignment = center_align
        worksheet[f"B{insert_row + 1}"].border = border

        # 12) If there were any mismatches, show a Streamlit warning
        if warning_rows:
            st.warning(
                f"⚠️ {len(warning_rows)} potential mismatch(es) found in Identification Type and Nationality/PR. Please check highlighted rows."
            )

    return output, mismatch_count


#
# ───── Streamlit UI: File Uploader & Download Button ───────────────────────────
#
uploaded_file = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Read all sheets into a dict
    xlsx = pd.read_excel(uploaded_file, sheet_name=None)
    # Extract just the Visitor List sheet for cleaning
    df = xlsx.get("Visitor List", pd.DataFrame())

    # Run the cleaning logic
    df_cleaned = clean_data(df)

    # Generate a new Excel file with all sheets preserved + cleaned Visitor List
    output, mismatch_count = generate_excel(xlsx, df_cleaned)

    # Attach a timestamped filename
    filename = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        label="📥 Download Cleaned Excel File",
        data=output.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
