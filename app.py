import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# Download link for the sample template
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f.read(),
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )



def nationality_group(row):
    """
    Assigns a sort key based on nationality + PR status:
      1 = Singaporean
      2 = PR (any nationality other than Singapore)
      3 = Malaysian (non-PR)
      4 = Indian
      5 = Others
    """
    nationality = str(row["Nationality (Country Name)"]).strip().lower()
    pr_status   = str(row["PR"]).strip().lower()

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
    Splits a full name into first name and remainder (middle + last).
    If there is no space, treats the entire string as 'first name' and leaves remainder blank.
    """
    s = str(name).strip()
    if " " in s:
        idx = s.find(" ")
        return pd.Series([s[:idx], s[idx+1:]])
    return pd.Series([s, ""])

def clean_gender(val):
    """
    Normalizes Gender field:
      - "M" -> "Male"
      - "F" -> "Female"
      - "MALE" -> "Male"
      - "FEMALE" -> "Female"
      - Otherwise returns the string as-is (title-cased if already fully upper-case).
    """
    s = str(val).strip().upper()
    if s == "M":
        return "Male"
    if s == "F":
        return "Female"
    if s in ["MALE", "FEMALE"]:
        return s.title()
    return val


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Takes the Visitor List DataFrame, applies all cleaning rules, and returns the cleaned DF.
    """
    # 1) Rename columns explicitly
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

    # 2) Remove any rows that are entirely blank between columns D–M
    df = df.dropna(subset=df.columns[3:13], how="all")

    # 3) Sort by Company → Nationality group → Full Name
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df.sort_values(
        by=["Company Full Name", "SortGroup", "Nationality (Country Name)", "Full Name As Per NRIC"],
        inplace=True
    )
    df.drop(columns=["SortGroup"], inplace=True)

    # 4) Reassign S/N as a running 1..N
    df["S/N"] = range(1, len(df) + 1)

    # 5) Clean Vehicle Plate Number: replace “/” or “,” with “;”, trim spaces, drop literal “nan”
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"]
        .astype(str)
        .str.replace(r"[\/\,]", ";", regex=True)
        .str.replace(r"\s*;\s*", ";", regex=True)
        .str.strip()
        .replace("nan", "", regex=False)
    )

    # 6) Proper‐case Full Name
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()

    # 7) Split into “First Name” + “Middle and Last Name”
    df[["First Name as per NRIC", "Middle and Last Name as per NRIC"
       ]] = df["Full Name As Per NRIC"].apply(split_name)

    # 8) Nationality mapping (“Chinese”→“China”, “Singaporean”→“Singapore”), then title‐case
    nationality_map = {"Chinese": "China", "Singaporean": "Singapore"}
    df["Nationality (Country Name)"] = df["Nationality (Country Name)"].replace(nationality_map)
    df["Nationality (Country Name)"] = df["Nationality (Country Name)"].astype(str).str.title()

    # 9) Fix swapped columns if any row's IC field contains “-”. 
    #    (Work Permit Expiry and IC suffix sometimes reversed by vendors)
    if df["IC (Last 3 digits and suffix) 123A"].astype(str).str.contains("-", na=False).any():
        df[["IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date"
           ]] = df[["Work Permit Expiry Date", "IC (Last 3 digits and suffix) 123A"]]

    # 10) Extract last 4 characters of IC suffix
    df["IC (Last 3 digits and suffix) 123A"] = (
        df["IC (Last 3 digits and suffix) 123A"]
        .astype(str)
        .str[-4:]
    )

    # 11) Mobile Number: drop all non‐digits, coerce to int, cast back to string to remove decimals
    df["Mobile Number"] = (
        pd.to_numeric(df["Mobile Number"], errors="coerce")
          .fillna(0)
          .astype(int)
          .astype(str)
    )

    # 12) Gender normalization
    df["Gender"] = df["Gender"].apply(clean_gender)

    # 13) Work Permit Expiry Date → standard YYYY-MM-DD (ignore time)
    df["Work Permit Expiry Date"] = (
        pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce")
          .dt.strftime("%Y-%m-%d")
    )

    return df



def generate_excel(all_sheets: dict, cleaned_df: pd.DataFrame):
    """
    - all_sheets: dictionary from pd.read_excel(..., sheet_name=None)
    - cleaned_df: the cleaned Visitor List DataFrame
    Writes all original sheets except “Visitor List” verbatim, then overwrites “Visitor List” with cleaned_df,
    applies formatting (header fill, borders, highlights, vehicles summary, total visitors).
    Returns: BytesIO buffer of the new .xlsx and the mismatch_count.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Write every sheet except the one that has “visitor” in its name (case‐insensitive)
        target_sheet_name = None
        for sheet_name, sheet_df in all_sheets.items():
            lower = sheet_name.lower()
            if "visitor" not in lower:
                sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)
            else:
                # remember the exact tab name, e.g. “Visitor List”
                target_sheet_name = sheet_name

        # 2) Write the cleaned DF under exactly that tab name
        cleaned_df.to_excel(writer, index=False, sheet_name=target_sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[target_sheet_name]

        # 3) Styles: header fill in RGB “94B455”, thin borders, center alignment, Calibri 11
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
        bold_font    = Font(name="Calibri", size=11, bold=True)

        # 4) Apply border, center alignment, Calibri to all cells in Visitor List
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = thin_border
                cell.alignment = center_align
                cell.font = calibri_font

        # 5) Header row formatting (row 1)
        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet[f"{get_column_letter(col_idx)}1"]
            cell.fill = header_fill
            cell.font = bold_font

        # 6) Freeze top row
        worksheet.freeze_panes = worksheet["A2"]

        # 7) Highlight mismatches in Identification Type vs. Nationality/PR
        mismatch_count = 0
        warning_rows  = []

        for r in range(2, worksheet.max_row + 1):
            id_type    = str(worksheet[f"G{r}"].value).strip().upper()
            nationality = str(worksheet[f"J{r}"].value).strip().title()
            pr_status   = str(worksheet[f"K{r}"].value).strip().title()

            highlight = False
            # (a) If ID = “NRIC” but Not (Singapore or PR): highlight
            if id_type == "NRIC" and not (
                nationality == "Singapore" or pr_status in ["Yes", "Pr"]
            ):
                highlight = True

            # (b) If ID = “FIN” but (Nationality = Singapore or PR): highlight
            if id_type == "FIN" and (
                nationality == "Singapore" or pr_status in ["Yes", "Pr"]
            ):
                highlight = True

            if highlight:
                warning_rows.append(r)
                worksheet[f"G{r}"].fill = light_red_fill
                worksheet[f"J{r}"].fill = light_red_fill
                worksheet[f"K{r}"].fill = light_red_fill
                mismatch_count += 1

        # 8) Auto‐fit column widths (approximation)
        for col in worksheet.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    length = len(str(cell.value))
                    if length > max_length:
                        max_length = length
            # add a little padding
            worksheet.column_dimensions[col_letter].width = max_length + 2

        # 9) Set all row heights to 20
        for row in worksheet.iter_rows():
            worksheet.row_dimensions[row[0].row].height = 20

        # 10) Vehicle summary (if any exist)
        vehicles = []
        for val in cleaned_df["Vehicle Plate Number"].dropna():
            vehicles.extend([v.strip() for v in str(val).split(";") if v.strip()])

        insert_row = worksheet.max_row + 2
        if vehicles:
            summary = ";".join(sorted(set(vehicles)))
            # “Vehicles” label
            cell = worksheet[f"B{insert_row}"]
            cell.value = "Vehicles"
            cell.border = thin_border
            cell.alignment = center_align

            # Joined list
            cell2 = worksheet[f"B{insert_row+1}"]
            cell2.value = summary
            cell2.border = thin_border
            cell2.alignment = center_align

            insert_row += 3

        # 11) Total Visitors = count of non‐null “Company Full Name”
        total_visitors = cleaned_df["Company Full Name"].notna().sum()
        worksheet[f"B{insert_row}"].value = "Total Visitors"
        worksheet[f"B{insert_row}"].alignment = center_align
        worksheet[f"B{insert_row}"].border = thin_border

        worksheet[f"B{insert_row+1}"].value = total_visitors
        worksheet[f"B{insert_row+1}"].alignment = center_align
        worksheet[f"B{insert_row+1}"].border = thin_border

        # 12) If any mismatches found, show a Streamlit warning
        if warning_rows:
            st.warning(f"⚠️ {len(warning_rows)} mismatch(es) found in Visitor List (Identification Type vs. Nationality/PR). Please check highlighted rows.")

    return output, mismatch_count



# ----------------------------
# Main Streamlit workflow
# ----------------------------
uploaded_file = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Read all sheets into a dict
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None)

    # Find the sheet whose name contains “visitor” (case-insensitive)
    visitor_sheet = None
    for name in all_sheets.keys():
        if "visitor" in name.lower():
            visitor_sheet = name
            break

    if visitor_sheet is None:
        st.error("❌ Could not find a sheet whose name contains “Visitor”. Please ensure your file has a “Visitor” tab.")
    else:
        raw_df = all_sheets[visitor_sheet]
        cleaned_df = clean_data(raw_df)

        # Generate the new Excel file (BytesIO)
        output_bytes, mismatch_count = generate_excel(all_sheets, cleaned_df)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Cleaned_Visitor_List_{timestamp}.xlsx"

        st.success("✅ Visitor List cleaned successfully!")
        st.download_button(
            label="📥 Download Cleaned Excel File",
            data=output_bytes.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
