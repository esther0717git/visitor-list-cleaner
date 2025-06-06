import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# ——————————————————————————————
# DOWNLOAD A FRESH “sample_template.xlsx”
# ——————————————————————————————
# This button serves a static sample template for users to download.
# Just make sure your “sample_template.xlsx” file is in the same folder as app.py.
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f.read(),
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )



# ——————————————————————————————
# UTILITY FUNCTIONS FOR CLEANING
# ——————————————————————————————
def nationality_group(row):
    """
    Assigns a sort key so that rows are grouped:
      1. Singapore nationals
      2. PR (“Yes” or “Pr” in PR column), regardless of nationality
      3. Malaysia (non-PR)
      4. India
      5. All others
    """
    nationality = str(row["Nationality (Country Name)"]).strip().lower()
    pr_status = str(row["PR"]).strip().lower()
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

def split_name(full_name: str):
    """
    Splits “Full Name As Per NRIC” into [First Name, Middle+Last].
    If no space, returns [full_name, ""]
    """
    name = str(full_name).strip()
    if " " in name:
        idx = name.find(" ")
        return pd.Series([name[:idx], name[idx+1:]])
    return pd.Series([name, ""])

def clean_gender(val):
    """
    Standardizes Gender:
      “M” → “Male”
      “F” → “Female”
      “MALE” → “Male”
      “FEMALE” → “Female”
      Otherwise, return unchanged string-title.
    """
    v = str(val).strip().upper()
    if v == "M":
        return "Male"
    elif v == "F":
        return "Female"
    elif v in ["MALE", "FEMALE"]:
        return v.title()
    return v.title()

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Takes in the “Visitor List” dataframe and applies all cleaning rules:
      • Drop rows that are entirely blank between columns D..M (Full Name to Mobile Number).
      • Rename columns into our canonical set.
      • Sort by Company, then by nationality/pr grouping, then by name.
      • Re‐index S/N as 1…N.
      • Clean “Vehicle Plate Number”: replace “/” or “,” with “;”, trim spaces.
      • Proper‐case Full Name.
      • Split Full Name into First / Middle+Last.
      • Map “Chinese”→“China”, “Singaporean”→“Singapore” and ensure Title‐Case nationality.
      • If the “IC (Last 3 digits and suffix)” column contains a “-” (i.e. looks like a date),
        swap it with “Work Permit Expiry Date”.
      • Extract only the last 4 characters of the IC column.
      • Force “Mobile Number” to integer (removing decimals/spaces).
      • Clean “Gender” per above.
      • Force “Work Permit Expiry Date” into YYYY-MM-DD (drop time).
    """
    # 1) Rename to the exact column list we expect:
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
        "Mobile Number"
    ]

    # 2) Drop rows where columns D..M (index 3..12) are all blank:
    df = df.dropna(subset=df.columns[3:13], how="all")

    # 3) Sort/Group by Company name → (Nationality+PR) → nationality → Full Name
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df.sort_values(
        by=["Company Full Name", "SortGroup", "Nationality (Country Name)", "Full Name As Per NRIC"],
        inplace=True
    )
    df.drop(columns=["SortGroup"], inplace=True)

    # 4) Reset the “S/N” column to run 1..len(df)
    df["S/N"] = range(1, len(df) + 1)

    # 5) Clean Vehicle Plate Number → replace “/” or “,” with “;”, strip
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
        .str.replace(r"[\/\,]", ";", regex=True)
        .str.replace(r"\s*;\s*", ";", regex=True)
        .str.strip()
        .replace("nan", "", regex=False)
    )

    # 6) Proper‐case “Full Name As Per NRIC”
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()

    # 7) Split into First / Middle+Last
    df[["First Name as per NRIC", "Middle and Last Name as per NRIC"]] = df["Full Name As Per NRIC"].apply(split_name)

    # 8) Map nationality: Chinese→China, Singaporean→Singapore, then Title‐Case
    nationality_map = {"Chinese": "China", "Singaporean": "Singapore"}
    df["Nationality (Country Name)"] = df["Nationality (Country Name)"].replace(nationality_map)
    df["Nationality (Country Name)"] = df["Nationality (Country Name)"].astype(str).str.title()

    # 9) Detect if “IC (Last 3 digits and suffix)” actually contains “‐” (like a date)
    #    If so, swap it with “Work Permit Expiry Date”
    if df["IC (Last 3 digits and suffix) 123A"].astype(str).str.contains("-", na=False).any():
        df[["IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date"]] = df[
            ["Work Permit Expiry Date", "IC (Last 3 digits and suffix) 123A"]
        ]

    # 10) Extract the last 4 chars of the IC suffix column
    df["IC (Last 3 digits and suffix) 123A"] = df["IC (Last 3 digits and suffix) 123A"].astype(str).str[-4:]

    # 11) Force Mobile Number → remove any non-digit, coerce to int, then back to str
    df["Mobile Number"] = (
        pd.to_numeric(df["Mobile Number"].astype(str).str.replace(r"\D", "", regex=True), errors="coerce")
        .fillna(0)
        .astype(int)
        .astype(str)
    )

    # 12) Clean Gender per the helper
    df["Gender"] = df["Gender"].apply(clean_gender)

    # 13) Standardize “Work Permit Expiry Date” → YYYY-MM-DD only (drop time)
    df["Work Permit Expiry Date"] = (
        pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce")
        .dt.strftime("%Y-%m-%d")
    )

    return df


# ——————————————————————————————
# GENERATE EXCEL (ALL TABS)
# ——————————————————————————————
def generate_excel(all_sheets_dict: dict, cleaned_df: pd.DataFrame):
    """
    Accepts:
      • all_sheets_dict: a dict of sheet_name → dataframe (read by pd.read_excel(..., sheet_name=None))
      • cleaned_df: the cleaned “Visitor List” dataframe

    Writes out a new in-memory workbook:
      • For every sheet ≠ “Visitor List”, write it untouched.
      • For “Visitor List”, overwrite with cleaned_df + styling.
      • Return BytesIO buffer + mismatch_count.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Write all untouched sheets first (except “Visitor List”)
        for sheet_name, sheet_df in all_sheets_dict.items():
            if sheet_name != "Visitor List":
                sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)

        # 2) Now write the cleaned “Visitor List”
        cleaned_df.to_excel(writer, index=False, sheet_name="Visitor List")

        # Grab workbook + worksheet objects
        workbook = writer.book
        worksheet = writer.sheets["Visitor List"]

        # Set up common styles
        header_fill   = PatternFill(start_color="94B455", end_color="94B455", fill_type="solid")
        light_red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )
        center_align = Alignment(horizontal="center", vertical="center")
        calibri_font = Font(name="Calibri", size=11)
        bold_calibri = Font(name="Calibri", size=11, bold=True)

        # 3) Apply borders, alignment, font to every cell
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = thin_border
                cell.alignment = center_align
                cell.font = calibri_font

        # 4) Header row (Row 1) gets the colored fill + bold font
        for col_idx in range(1, worksheet.max_column + 1):
            hdr_cell = worksheet[f"{get_column_letter(col_idx)}1"]
            hdr_cell.fill = header_fill
            hdr_cell.font = bold_calibri

        # 5) Freeze the top row
        worksheet.freeze_panes = worksheet["A2"]

        # 6) Highlight logic: ID vs Nationality/PR mismatches
        mismatch_count = 0
        warning_rows = []
        for r in range(2, worksheet.max_row + 1):
            id_type   = str(worksheet[f"G{r}"].value).strip().upper()
            nationality = str(worksheet[f"J{r}"].value).strip().title()
            pr_status = str(worksheet[f"K{r}"].value).strip().title()

            highlight = False
            # If “NRIC” but NOT (Singapore OR foreign PR) → highlight
            if id_type == "NRIC" and not (
                nationality == "Singapore" or (nationality != "Singapore" and pr_status in ["Yes", "Pr"])
            ):
                highlight = True
            # If “FIN” but (Nationality=Singapore OR PR=Yes) → highlight
            if id_type == "FIN" and (nationality == "Singapore" or pr_status in ["Yes", "Pr"]):
                highlight = True

            if highlight:
                warning_rows.append(r)
                worksheet[f"G{r}"].fill = light_red_fill
                worksheet[f"J{r}"].fill = light_red_fill
                worksheet[f"K{r}"].fill = light_red_fill
                mismatch_count += 1

        # 7) Auto-fit column widths
        for col in worksheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
            worksheet.column_dimensions[col_letter].width = max_len + 2

        # 8) Auto-fit row heights to a fixed 20 points
        for r in worksheet.iter_rows():
            worksheet.row_dimensions[r[0].row].height = 20

        # 9) “Vehicles” summary: gather every unique plate (split on “;”), sort, join by “;”
        all_plates = []
        for val in cleaned_df["Vehicle Plate Number"].dropna():
            for chunk in str(val).split(";"):
                chunk = chunk.strip()
                if chunk:
                    all_plates.append(chunk)
        unique_sorted = sorted(set(all_plates))

        insert_row = worksheet.max_row + 2
        if unique_sorted:
            worksheet[f"B{insert_row}"].value = "Vehicles"
            worksheet[f"B{insert_row}"].border = thin_border
            worksheet[f"B{insert_row}"].alignment = center_align

            worksheet[f"B{insert_row+1}"].value = ";".join(unique_sorted)
            worksheet[f"B{insert_row+1}"].border = thin_border
            worksheet[f"B{insert_row+1}"].alignment = center_align
            insert_row += 3

        # 10) “Total Visitors” = count of non-blank “Company Full Name”
        total_visitors = cleaned_df["Company Full Name"].notna().sum()
        worksheet[f"B{insert_row}"].value = "Total Visitors"
        worksheet[f"B{insert_row}"].alignment = center_align
        worksheet[f"B{insert_row}"].border = thin_border

        worksheet[f"B{insert_row+1}"].value = total_visitors
        worksheet[f"B{insert_row+1}"].alignment = center_align
        worksheet[f"B{insert_row+1}"].border = thin_border

        # 11) If there were any mismatches, show a Streamlit warning
        if warning_rows:
            st.warning(f"⚠️ {len(warning_rows)} mismatch(es) found in Identification Type vs Nationality/PR. Please correct highlighted rows.")

    return output, mismatch_count



# ——————————————————————————————
# STREAMLIT Uploader / Download Button
# ——————————————————————————————
uploaded_file = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])
if uploaded_file:
    # 1) Read all sheets into a dict
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None)

    # 2) Take the “Visitor List” sheet and clean it
    df_visitor = all_sheets.get("Visitor List", pd.DataFrame())
    df_cleaned = clean_data(df_visitor)

    # 3) Generate the new in-memory Excel with all tabs
    output_buffer, mismatch_ct = generate_excel(all_sheets, df_cleaned)

    # 4) Provide a timestamped download
    filename = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        label="📥 Download Cleaned Excel File",
        data=output_buffer.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
