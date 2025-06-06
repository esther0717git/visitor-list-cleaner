import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🫧 CLARITY GATE - Data Cleaning")

# Download button for the sample template
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def nationality_group(row):
    """Assign sorting groups based on nationality and PR status."""
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
    """Split a full name into first name and middle+last name."""
    name = str(name).strip()
    if " " in name:
        first_space = name.find(" ")
        return pd.Series([name[:first_space], name[first_space+1:]])
    return pd.Series([name, ""])

def clean_gender(val):
    """Normalize gender values to 'Male' or 'Female'."""
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
    Perform all the Visitor List cleaning steps:
    - Rename columns
    - Drop rows where columns D–M are all blank
    - Sort by Company, nationality/PR grouping, and name
    - Reassign serial numbers
    - Normalize Vehicle Plate Number, Full Name, split into E/F
    - Normalize Nationality, swap columns if IC/Work Permit are reversed
    - Truncate IC suffix to last 4 chars
    - Ensure Mobile Number has no decimals (stored as string of digits)
    - Normalize Gender
    - Format Work Permit Expiry Date as YYYY-MM-DD
    """
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

    # Drop rows where columns D through M (indexes 3–12) are all NaN
    df = df.dropna(subset=df.columns[3:13], how="all")

    # Sort by Company, then nationality/PR group, then nationality, then full name
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df.sort_values(
        by=["Company Full Name", "SortGroup", "Nationality (Country Name)", "Full Name As Per NRIC"],
        inplace=True
    )
    df.drop(columns=["SortGroup"], inplace=True)

    # Reassign S/N (serial numbers)
    df["S/N"] = range(1, len(df) + 1)

    # Normalize Vehicle Plate Number: replace slashes/commas with ';', remove surrounding spaces, drop "nan"
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"]
        .astype(str)
        .str.replace(r"[\/,]", ";", regex=True)
        .str.replace(r"\s*;\s*", ";", regex=True)
        .str.strip()
        .replace("nan", "", regex=False)
    )

    # Proper-case the Full Name and split into first vs middle+last
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC", "Middle and Last Name as per NRIC"]] = df[
        "Full Name As Per NRIC"
    ].apply(split_name)

    # Normalize Nationality: map certain adjectives to country names, proper-case everything
    nationality_map = {"Chinese": "China", "Singaporean": "Singapore"}
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
        .replace(nationality_map)
        .astype(str)
        .str.title()
    )

    # If IC suffix and Work Permit columns were swapped, detect by looking for '-' in the IC field
    if df["IC (Last 3 digits and suffix) 123A"].astype(str).str.contains("-", na=False).any():
        df[["IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date"]] = df[
            ["Work Permit Expiry Date", "IC (Last 3 digits and suffix) 123A"]
        ]

    # Truncate IC suffix to last 4 characters
    df["IC (Last 3 digits and suffix) 123A"] = df[
        "IC (Last 3 digits and suffix) 123A"
    ].astype(str).str[-4:]

    # Remove all non-digit characters from Mobile Number, coerce to numeric then back to string
    df["Mobile Number"] = (
        pd.to_numeric(df["Mobile Number"], errors="coerce")
        .fillna(0)
        .astype(int)
        .astype(str)
    )

    # Normalize Gender
    df["Gender"] = df["Gender"].apply(clean_gender)

    # Format Work Permit Expiry Date as YYYY-MM-DD (or blank if invalid)
    df["Work Permit Expiry Date"] = pd.to_datetime(
        df["Work Permit Expiry Date"], errors="coerce"
    ).dt.strftime("%Y-%m-%d").fillna("")

    return df

def generate_excel(xlsx_dict: dict, df_cleaned: pd.DataFrame):
    """
    Recombine all tabs from the uploaded workbook:
    - Write every sheet except "Visitor List" untouched.
    - Overwrite the "Visitor List" sheet with the cleaned DataFrame.
    - Apply styling (header color, borders, auto-fit, row height, highlights, summaries).
    """
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Write each original sheet (except Visitor List) as-is
        for sheet_name, sheet_df in xlsx_dict.items():
            if sheet_name != "Visitor List":
                sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)

        # Overwrite Visitor List with cleaned data
        df_cleaned.to_excel(writer, index=False, sheet_name="Visitor List")
        workbook = writer.book
        worksheet = writer.sheets["Visitor List"]

        # Define styling objects
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

        # Apply border, alignment, and font to all cells
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = center_align
                cell.font = font_style

        # Style the header row: fill color + bold
        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet[f"{get_column_letter(col_idx)}1"]
            cell.fill = header_fill
            cell.font = bold_font

        # Freeze the top row
        worksheet.freeze_panes = worksheet["A2"]

        # Validation: highlight mismatches between ID type, nationality, PR
        mismatch_count = 0
        warning_rows = []
        for row_idx in range(2, worksheet.max_row + 1):
            id_type = str(worksheet[f"G{row_idx}"].value).strip().upper()
            nationality = str(worksheet[f"J{row_idx}"].value).strip().title()
            pr_status = str(worksheet[f"K{row_idx}"].value).strip().title()

            highlight = False
            # NRIC must be Singapore or foreign PR
            if id_type == "NRIC" and not (
                (nationality == "Singapore") or (nationality != "Singapore" and pr_status in ["Yes", "Pr"])
            ):
                highlight = True
            # FIN cannot be Singapore or PR
            if id_type == "FIN" and (
                (nationality == "Singapore") or (pr_status in ["Yes", "Pr"])
            ):
                highlight = True

            if highlight:
                warning_rows.append(row_idx)
                worksheet[f"G{row_idx}"].fill = light_red_fill
                worksheet[f"J{row_idx}"].fill = light_red_fill
                worksheet[f"K{row_idx}"].fill = light_red_fill
                mismatch_count += 1

        # Auto-fit column widths
        for col in worksheet.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            worksheet.column_dimensions[col_letter].width = max_length + 2

        # Set a fixed row height for all rows
        for row in worksheet.iter_rows():
            worksheet.row_dimensions[row[0].row].height = 20

        # Collect all unique vehicle plates for the summary
        vehicles = []
        for val in df_cleaned["Vehicle Plate Number"].dropna():
            vehicles.extend([v.strip() for v in str(val).split(";") if v.strip()])

        # Insert "Vehicles" label + summary two rows below the last data row
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

        # Insert "Total Visitors" label + count one row below the vehicles summary
        total_visitors = df_cleaned["Company Full Name"].notna().sum()
        worksheet[f"B{insert_row}"].value = "Total Visitors"
        worksheet[f"B{insert_row}"].alignment = center_align
        worksheet[f"B{insert_row}"].border = border

        worksheet[f"B{insert_row + 1}"].value = total_visitors
        worksheet[f"B{insert_row + 1}"].alignment = center_align
        worksheet[f"B{insert_row + 1}"].border = border

        # If any warnings exist, show a Streamlit warning message
        if warning_rows:
            st.warning(
                f"⚠️ {len(warning_rows)} potential mismatch(es) found in Identification Type and Nationality/PR. "
                "Please check highlighted rows."
            )

    return output, mismatch_count

# Streamlit file uploader
uploaded_file = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Read all sheets into a dict
    xlsx_dict = pd.read_excel(uploaded_file, sheet_name=None)
    # Extract the Visitor List sheet as a DataFrame
    df_original = xlsx_dict.get("Visitor List", pd.DataFrame())
    # Clean the Visitor List
    df_cleaned = clean_data(df_original)
    # Regenerate a new workbook (including tabs 2 & 3 untouched)
    output, mismatch_count = generate_excel(xlsx_dict, df_cleaned)

    # Build dynamic filename
    filename = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    # Provide download button
    st.download_button(
        label="📥 Download Cleaned Excel File",
        data=output.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
