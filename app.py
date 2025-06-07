import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# Provide a download button for the sample template
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def nationality_group(row):
    """
    Assign a sort group based on nationality and PR status:
    1 = Singaporean (any PR flag ignored because they're already Singapore)
    2 = PR (non-Singaporean nationals with PR 'Yes' or 'Pr')
    3 = Malaysian (non-PR)
    4 = Indian (non-PR)
    5 = Others
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


def split_name(name):
    """
    Split a full name string (e.g. "John Doe Smith") into:
      - First Name = everything before the first space
      - Middle+Last = everything after the first space
    """
    name = str(name).strip()
    if " " in name:
        first_space = name.find(" ")
        return pd.Series([name[:first_space], name[first_space + 1 :]])
    return pd.Series([name, ""])


def clean_gender(val):
    """
    Normalize gender entries:
      - "M" or "Male" (case-insensitive) => "Male"
      - "F" or "Female" (case-insensitive) => "Female"
      - Otherwise, return the raw (title‐cased) string
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
    Perform all cleaning on the 'Visitor List' sheet:
      1. Rename columns to a known fixed schema.
      2. Drop rows that are blank in columns D–M (Full Name through Mobile).
      3. Sort by Company, then nationality/PR groups, then name.
      4. Re‐index S/N as consecutive integers starting from 1.
      5. Clean vehicle plates: replace "/" or "," with ";" and remove extra spaces.
      6. Proper‐case the full name, and re‐derive first/middle‐last from it.
      7. Map 'Chinese' → 'China', 'Singaporean' → 'Singapore'; then title‐case nationality.
      8. If IC and Work Permit columns are swapped (detected by a "-" in IC), swap them back.
      9. Extract only the last 4 characters of the IC string.
     10. Ensure Mobile is numeric (drop decimals) and convert to string.
     11. Normalize Gender.
     12. Standardize Work Permit Expiry as "YYYY-MM-DD".
    """
    # 1. Rename columns exactly
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

    # 2. Drop rows where columns D–M (index 3:13) are all NaN/empty
    df = df.dropna(subset=df.columns[3:13], how="all")

    # 3. Sort by Company, then nationality group, then nationality text, then full name
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

    # 4. Re‐index S/N sequentially
    df["S/N"] = range(1, len(df) + 1)

    # 5. Clean Vehicle Plate Number
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"]
        .astype(str)
        .str.replace(r"[\/\,]", ";", regex=True)        # "/" or "," → ";"
        .str.replace(r"\s*;\s*", ";", regex=True)       # remove spaces around ";"
        .str.strip()
        .replace("nan", "", regex=False)
    )

    # 6. Proper‐case Full Name As Per NRIC
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    # Re‐split into first + middle/last
    df[
        ["First Name as per NRIC", "Middle and Last Name as per NRIC"]
    ] = df["Full Name As Per NRIC"].apply(split_name)

    # 7. Nationality mapping + proper‐case
    nationality_map = {"Chinese": "China", "Singaporean": "Singapore"}
    df["Nationality (Country Name)"] = df["Nationality (Country Name)"].replace(
        nationality_map
    )
    df["Nationality (Country Name)"] = df["Nationality (Country Name)"].astype(
        str
    ).str.title()

    # 8. Detect if IC and Work Permit are swapped: if IC column contains "-" then swap
    if df["IC (Last 3 digits and suffix) 123A"].astype(str).str.contains("-", na=False).any():
        df[
            ["IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date"]
        ] = df[
            ["Work Permit Expiry Date", "IC (Last 3 digits and suffix) 123A"]
        ]

    # 9. Take last 4 characters of IC suffix
    df["IC (Last 3 digits and suffix) 123A"] = df[
        "IC (Last 3 digits and suffix) 123A"
    ].astype(str).str[-4:]

    # 10. Mobile Number → numeric, drop any decimals, then back to string
    df["Mobile Number"] = (
        pd.to_numeric(df["Mobile Number"], errors="coerce").fillna(0).astype(int).astype(str)
    )

    # 11. Normalize Gender
    df["Gender"] = df["Gender"].apply(clean_gender)

    # 12. Standardize Work Permit Expiry Date to "YYYY-MM-DD"
    df["Work Permit Expiry Date"] = (
        pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce")
        .dt.strftime("%Y-%m-%d")
    )

    return df


def generate_excel(xlsx: dict, df_cleaned: pd.DataFrame):
    """
    Rebuild an in-memory Excel file that:
      • Leaves all sheets except any sheet whose name contains "visitor" (case‐insensitive) untouched.
      • Overwrites the Visitor sheet (detected by "visitor" in its name) with df_cleaned.
      • Applies styling (header fill, borders, freeze pane, auto‐fit, mismatch highlighting).
      • Appends Vehicles summary and Total Visitors at the bottom of the Visitor sheet.
    Returns: (BytesIO_buffer, mismatch_count)
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1. Write every sheet except "Visitor" (by name detection) exactly as it was:
        for sheet_name, sheet_df in xlsx.items():
            if "visitor" not in sheet_name.lower():
                sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)

        # 2. Overwrite the Visitor sheet (match by name containing "visitor", case‐insensitive)
        #    If there are multiple "visitor" tabs, pick the first one encountered.
        visitor_sheet_name = None
        for name in xlsx.keys():
            if "visitor" in name.lower():
                visitor_sheet_name = name
                break

        if visitor_sheet_name is None:
            st.error("❌ No sheet containing 'Visitor' found in the uploaded workbook.")
            return None, 0

        # Write the cleaned DataFrame under exactly that sheet name
        df_cleaned.to_excel(writer, index=False, sheet_name=visitor_sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[visitor_sheet_name]

        # 3. Define styles
        header_fill = PatternFill(
            start_color="94B455", end_color="94B455", fill_type="solid"
        )
        light_red_fill = PatternFill(
            start_color="FFCCCC", end_color="FFCCCC", fill_type="solid"
        )
        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )
        center_align = Alignment(horizontal="center", vertical="center")
        font_style = Font(name="Calibri", size=11)
        bold_font = Font(name="Calibri", size=11, bold=True)

        # 4. Apply border+alignment+font to ALL cells
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = center_align
                cell.font = font_style

        # 5. Header row styling (row 1)
        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet[f"{get_column_letter(col_idx)}1"]
            cell.fill = header_fill
            cell.font = bold_font

        # 6. Freeze the top row
        worksheet.freeze_panes = worksheet["A2"]

        # 7. Highlight mismatches in Identification Type vs Nationality/PR
        mismatch_count = 0
        warning_rows = []
        for row_idx in range(2, worksheet.max_row + 1):
            id_type = str(worksheet[f"G{row_idx}"].value).strip().upper()
            nationality = str(worksheet[f"J{row_idx}"].value).strip().title()
            pr_status = str(worksheet[f"K{row_idx}"].value).strip().title()

            highlight = False
            if id_type == "NRIC":
                # NRIC → valid only if nationality == "Singapore" OR (nationality != "Singapore" AND PR == Yes/Pr)
                if not (
                    nationality == "Singapore"
                    or (nationality != "Singapore" and pr_status in ["Yes", "Pr"])
                ):
                    highlight = True
            if id_type == "FIN":
                # FIN → invalid if nationality=="Singapore" OR PR == Yes/Pr
                if nationality == "Singapore" or pr_status in ["Yes", "Pr"]:
                    highlight = True

            if highlight:
                warning_rows.append(row_idx)
                worksheet[f"G{row_idx}"].fill = light_red_fill
                worksheet[f"J{row_idx}"].fill = light_red_fill
                worksheet[f"K{row_idx}"].fill = light_red_fill
                mismatch_count += 1

        # 8. Auto‐fit column widths
        for col in worksheet.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            worksheet.column_dimensions[col_letter].width = max_length + 2

        # 9. Auto‐fit row heights (set a default height if content exists)
        for row in worksheet.iter_rows():
            worksheet.row_dimensions[row[0].row].height = 20

        # 10. Vehicles summary: collect unique plates (split by ";" and strip), sort, join
        vehicles = []
        for val in df_cleaned["Vehicle Plate Number"].dropna():
            vehicles.extend([v.strip() for v in str(val).split(";") if v.strip()])
        if vehicles:
            insert_row = worksheet.max_row + 2
            unique_sorted = sorted(set(vehicles))
            summary = ";".join(unique_sorted)

            worksheet[f"B{insert_row}"].value = "Vehicles"
            worksheet[f"B{insert_row}"].border = border
            worksheet[f"B{insert_row}"].alignment = center_align

            worksheet[f"B{insert_row + 1}"].value = summary
            worksheet[f"B{insert_row + 1}"].border = border
            worksheet[f"B{insert_row + 1}"].alignment = center_align

            insert_row += 3
        else:
            insert_row = worksheet.max_row + 1

        # 11. Total Visitors: count non‐NA "Company Full Name"
        total_visitors = df_cleaned["Company Full Name"].notna().sum()
        worksheet[f"B{insert_row}"].value = "Total Visitors"
        worksheet[f"B{insert_row}"].alignment = center_align
        worksheet[f"B{insert_row}"].border = border

        worksheet[f"B{insert_row + 1}"].value = total_visitors
        worksheet[f"B{insert_row + 1}"].alignment = center_align
        worksheet[f"B{insert_row + 1}"].border = border

        # 12. If any mismatches found, display a Streamlit warning
        if warning_rows:
            st.warning(
                f"⚠️ {len(warning_rows)} potential mismatch(es) found in Identification Type vs Nationality/PR. Rows highlighted in light red."
            )

    return output, mismatch_count


# --- Streamlit UI: File Uploader + Clean → Download ---
uploaded_file = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Read all sheets into a dict
    xlsx_dict = pd.read_excel(uploaded_file, sheet_name=None)
    # Identify the visitor sheet by checking for "visitor" in the sheet name (case‐insensitive)
    visitor_sheet = None
    for name in xlsx_dict.keys():
        if "visitor" in name.lower():
            visitor_sheet = name
            break

    if visitor_sheet is None:
        st.error("❌ Uploaded workbook does not contain any sheet with 'Visitor' in its name.")
    else:
        # Clean only the visitor sheet
        df_raw = xlsx_dict[visitor_sheet]
        df_cleaned = clean_data(df_raw)

        # Rebuild the Excel with unchanged tabs 2 & 3, and the cleaned Visitor tab
        output_buffer, mismatch_count = generate_excel(xlsx_dict, df_cleaned)

        # Offer the cleaned file for download
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Cleaned_Visitor_List_{timestamp}.xlsx"
        st.download_button(
            label="📥 Download Cleaned Excel File",
            data=output_buffer.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
