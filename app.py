import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

# -----------------------
# Page configuration
# -----------------------
st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🫧CLARITY GATE - Data Validation and Cleaning")

# Provide a “Download Sample Template” button
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =======================
# Helper functions
# =======================
def nationality_group(row):
    """
    Assign a sort group based on nationality/PR logic:
      - 1 = Singaporean
      - 2 = Any PR (“PR” column is “Yes” or “Pr”), regardless of nationality
      - 3 = Malaysian non-PR
      - 4 = Indian
      - 5 = All others
    """
    nationality = str(row["Nationality (Country Name)"]).lower()
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


def split_name(name):
    """
    Split “Full Name As Per NRIC” into [First, Middle+Last].
    """
    name = str(name).strip()
    if " " in name:
        first_space = name.find(" ")
        return pd.Series([name[:first_space], name[first_space + 1 :]])
    return pd.Series([name, ""])


def clean_gender(val):
    """
    Normalize Gender column:
      - “M” → “Male”
      - “F” → “Female”
      - “MALE”/“FEMALE” → title case
    """
    val = str(val).strip().upper()
    if val == "M":
        return "Male"
    elif val == "F":
        return "Female"
    elif val in ["MALE", "FEMALE"]:
        return val.title()
    return val


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    This function cleans **only** the Visitor List sheet. It:
      1. Renames columns to a known standard.
      2. Drops any rows where columns D–M (Full Name, etc.) are all blank.
      3. Sorts by (Company, nationality group, nationality, full name).
      4. Re‐runs S/N as 1, 2, 3, … in the final order.
      5. Cleans “Vehicle Plate Number” → replaces “/” or “,” with “;”, trims spaces.
      6. Capitalizes “Full Name As Per NRIC” (Proper case) and re‐splits into col E/F.
      7. Standardizes Nationality → maps “Chinese”→“China” etc., and title‐cases.
      8. If the IC‐suffix and Work Permit columns are swapped (i.e. one contains “YYYY-MM-DD”),
         it flips them back into the correct order.
      9. Extracts the last 4 chars of IC suffix.
     10. Forces “Mobile Number” into integer→string (no decimals).
     11. Cleans Gender (M/F → Male/Female).
     12. Reformats “Work Permit Expiry Date” to “YYYY-MM-DD” (drops any time portion).
    """
    # 1) Rename columns to known standard:
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

    # 2) Drop rows where columns D–M are all blank:
    df = df.dropna(subset=df.columns[3:13], how="all")

    # 3) Sort by (Company, nationality group, nationality, full name):
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

    # 4) Re‐assign running S/N as 1..N in final order:
    df["S/N"] = range(1, len(df) + 1)

    # 5) Clean “Vehicle Plate Number”:
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"]
        .astype(str)
        .str.replace(r"[\/\,]", ";", regex=True)   # replace “/” or “,” with “;”
        .str.replace(r"\s*;\s*", ";", regex=True)  # no spaces around “;”
        .str.strip()
        .replace("nan", "", regex=False)
    )

    # 6) Proper‐case full name and re‐split into col E/F:
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[
        ["First Name as per NRIC", "Middle and Last Name as per NRIC"]
    ] = df["Full Name As Per NRIC"].apply(split_name)

    # 7) Nationality mapping and title‐case:
    nationality_map = {"Chinese": "China", "Singaporean": "Singapore"}
    df["Nationality (Country Name)"] = df["Nationality (Country Name)"].replace(
        nationality_map
    )
    df["Nationality (Country Name)"] = df["Nationality (Country Name)"].astype(
        str
    ).str.title()

    # 8) If IC suffix / Work Permit got swapped, swap back:
    #    We assume an IC suffix always has a trailing letter, not a date “YYYY-MM-DD”.
    if df["IC (Last 3 digits and suffix) 123A"].astype(str).str.contains("-", na=False).any():
        # Swap these two columns:
        df[
            ["IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date"]
        ] = df[["Work Permit Expiry Date", "IC (Last 3 digits and suffix) 123A"]]

    # 9) Truncate IC suffix to last 4 chars:
    df["IC (Last 3 digits and suffix) 123A"] = df[
        "IC (Last 3 digits and suffix) 123A"
    ].astype(str).str[-4:]

    # 10) Mobile Number → drop any non‐digits, then to int→str (no decimals):
    df["Mobile Number"] = (
        pd.to_numeric(df["Mobile Number"], errors="coerce")
        .fillna(0)
        .astype(int)
        .astype(str)
    )

    # 11) Clean Gender column:
    df["Gender"] = df["Gender"].apply(clean_gender)

    # 12) Work Permit Expiry Date → format “YYYY-MM-DD”:
    df["Work Permit Expiry Date"] = pd.to_datetime(
        df["Work Permit Expiry Date"], errors="coerce"
    ).dt.strftime("%Y-%m-%d")

    return df


def generate_excel(xlsx_dict: dict, cleaned_df: pd.DataFrame):
    """
    Writes back all sheets to a new Excel in memory:
      - Tabs 2 & 3 (Delivery Information, Serial Number For Shipment) are written unchanged.
      - Tab 1 (“Visitor List”) is overwritten with cleaned_df + formatting & highlights.
    Returns (BytesIO, mismatch_count).
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Write all original sheets except “Visitor List” unchanged:
        for sheet_name, sheet_df in xlsx_dict.items():
            if sheet_name != "Visitor List":
                sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)

        # 2) Overwrite “Visitor List” with the cleaned version:
        cleaned_df.to_excel(writer, index=False, sheet_name="Visitor List")
        workbook = writer.book
        worksheet = writer.sheets["Visitor List"]

        # -----------------------
        # Styles & Formats
        # -----------------------
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

        # Apply border / center align / font to all cells:
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = center_align
                cell.font = font_style

        # Header row (row 1): set fill + bold:
        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet[f"{get_column_letter(col_idx)}1"]
            cell.fill = header_fill
            cell.font = bold_font

        # Freeze the top row:
        worksheet.freeze_panes = worksheet["A2"]

        # -----------------------
        # Highlight mismatches (Identification Type vs Nationality/PR):
        # -----------------------
        mismatch_count = 0
        warning_rows = []
        for row_idx in range(2, worksheet.max_row + 1):
            id_type = str(worksheet[f"G{row_idx}"].value).strip().upper()
            nationality = str(worksheet[f"J{row_idx}"].value).strip().title()
            pr_status = str(worksheet[f"K{row_idx}"].value).strip().title()

            highlight = False
            # Rule: If ID type = “NRIC” → nationality must be Singapore OR PR=Yes. Otherwise highlight.
            if id_type == "NRIC" and not (
                nationality == "Singapore" or pr_status in ["Yes", "Pr"]
            ):
                highlight = True

            # Rule: If ID type = “FIN”, they cannot be Singapore citizen or PR → highlight if they are
            if id_type == "FIN" and (
                nationality == "Singapore" or pr_status in ["Yes", "Pr"]
            ):
                highlight = True

            if highlight:
                warning_rows.append(row_idx)
                worksheet[f"G{row_idx}"].fill = light_red_fill
                worksheet[f"J{row_idx}"].fill = light_red_fill
                worksheet[f"K{row_idx}"].fill = light_red_fill
                mismatch_count += 1

        # -----------------------
        # Autofit column widths & row heights:
        # -----------------------
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

        for row in worksheet.iter_rows():
            worksheet.row_dimensions[row[0].row].height = 20

        # -----------------------
        # Vehicles summary (if any)
        # -----------------------
        vehicles = []
        for val in cleaned_df["Vehicle Plate Number"].dropna():
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

        # -----------------------
        # Total visitors summary (count of “Company Full Name” rows)
        # -----------------------
        total_visitors = cleaned_df["Company Full Name"].notna().sum()
        worksheet[f"B{insert_row}"].value = "Total Visitors"
        worksheet[f"B{insert_row}"].alignment = center_align
        worksheet[f"B{insert_row}"].border = border
        worksheet[f"B{insert_row + 1}"].value = total_visitors
        worksheet[f"B{insert_row + 1}"].alignment = center_align
        worksheet[f"B{insert_row + 1}"].border = border

        # If any mismatches, show a warning in Streamlit:
        if warning_rows:
            st.warning(
                f"⚠️ {len(warning_rows)} potential mismatch(es) found "
                "in Identification Type vs Nationality/PR. Please check highlighted rows."
            )

    return output, mismatch_count


# =======================
# Main Streamlit flow
# =======================
uploaded_file = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Read all sheets into a dict:
    xlsx_dict = pd.read_excel(uploaded_file, sheet_name=None)

    # Extract the “Visitor List” sheet as a DataFrame:
    if "Visitor List" not in xlsx_dict:
        st.error("❌ Worksheet named 'Visitor List' not found. Please ensure your file has a tab called 'Visitor List'.")
    else:
        original_df = xlsx_dict["Visitor List"]

        # Clean only the visitor list:
        df_cleaned = clean_data(original_df)

        # Generate a new Excel in memory (all sheets):
        output_bytes, mismatch_ct = generate_excel(xlsx_dict, df_cleaned)

        # Offer the cleaned file for download:
        filename = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        st.success("✅ File cleaned and ready for download!")
        st.download_button(
            label="📥 Download Cleaned Excel File",
            data=output_bytes.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
