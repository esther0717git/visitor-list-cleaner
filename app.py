import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# “Download Sample Template” button
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

#
# ─── CLEANING HELPERS ────────────────────────────────────────────────────────────
#
def nationality_group(row):
    """Return a sort‐group integer based on nationality/PR."""
    nationality = str(row["Nationality (Country Name)"]).lower()
    pr_status = str(row["PR"]).strip().lower()
    if nationality == "singapore":
        return 1
    elif pr_status in ["yes", "pr"]:
        # Any PR (including Malaysian PR, etc.)
        return 2
    elif nationality == "malaysia":
        return 3
    elif nationality == "india":
        return 4
    else:
        return 5

def split_name(name):
    """Split a full name into (first, rest). If no space, rest = ''."""
    name = str(name).strip()
    if " " in name:
        first_space = name.find(" ")
        return pd.Series([name[:first_space], name[first_space+1:]])
    return pd.Series([name, ""])

def clean_gender(val):
    """Normalize gender: M → 'Male', F → 'Female', 'MALE' → 'Male', 'FEMALE' → 'Female'."""
    v = str(val).strip().upper()
    if v == "M":
        return "Male"
    elif v == "F":
        return "Female"
    elif v in ["MALE", "FEMALE"]:
        return v.title()
    return v

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    1. Rename columns to a consistent set.
    2. Drop fully‐blank rows (columns D–M all blank).
    3. Sort by Company, then nationality+PR, then name.
    4. Re‐index S/N as 1..N.
    5. Clean Vehicle Plate (replace '/', ',' → ';', remove extra spaces).
    6. Title‐case Full Name, then split into first/rest.
    7. Normalize nationality map (e.g. 'Chinese'→'China', 'Singaporean'→'Singapore') + title‐case.
    8. If Work Permit Date/IC suffix are swapped (detect a '-' in IC column), swap them back.
    9. Truncate IC to last 4 chars.
    10. Force Mobile Number → integer string (no decimals, no extra characters).
    11. Clean gender.
    12. Re‐format Work Permit Date to YYYY-MM-DD (drop any time).
    """
    # 1. Rename columns (expects exactly 13 columns in Visitor List)
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

    # 2. Drop rows where columns D–M (index 3..12) are all NaN/blank
    df = df.dropna(subset=df.columns[3:13], how="all")

    # 3. Sort by Company, then nationality grouping, then nationality text, then name
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

    # 4. Reassign S/N = 1..N in sorted order
    df["S/N"] = range(1, len(df) + 1)

    # 5. Clean Vehicle Plate Number: replace '/',','→';', remove leading/trailing spaces around ';'
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"]
        .astype(str)
        .str.replace(r"[\/\,]", ";", regex=True)
        .str.replace(r"\s*;\s*", ";", regex=True)
        .str.strip()
        .replace("nan", "", regex=False)
    )

    # 6. Title-case Full Name; then re-split into first/rest
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC", "Middle and Last Name as per NRIC"]] = df[
        "Full Name As Per NRIC"
    ].apply(split_name)

    # 7. Normalize nationality map + title-case everything
    nationality_map = {"Chinese": "China", "Singaporean": "Singapore"}
    df["Nationality (Country Name)"] = df["Nationality (Country Name)"].replace(
        nationality_map
    )
    df["Nationality (Country Name)"] = df["Nationality (Country Name)"].astype(str).str.title()

    # 8. Detect if IC suffix and Work Permit Date got swapped. If IC column contains a '-' (a date),
    #    then swap the two columns back.
    if df["IC (Last 3 digits and suffix) 123A"].astype(str).str.contains("-", na=False).any():
        df[
            ["IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date"]
        ] = df[
            ["Work Permit Expiry Date", "IC (Last 3 digits and suffix) 123A"]
        ]

    # 9. Truncate IC to last 4 characters
    df["IC (Last 3 digits and suffix) 123A"] = df[
        "IC (Last 3 digits and suffix) 123A"
    ].astype(str).str[-4:]

    # 10. Clean Mobile Number: remove any non-digit, then coerce to integer, then back to string
    df["Mobile Number"] = (
        pd.to_numeric(df["Mobile Number"], errors="coerce")
        .fillna(0)
        .astype(int)
        .astype(str)
    )

    # 11. Clean Gender
    df["Gender"] = df["Gender"].apply(clean_gender)

    # 12. Format Work Permit Expiry Date to YYYY-MM-DD (drop times)
    df["Work Permit Expiry Date"] = (
        pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce")
        .dt.strftime("%Y-%m-%d")
    )

    return df

#
# ─── EXCEL‐WRITING (ALL SHEETS) ────────────────────────────────────────────────────
#
def generate_excel(xlsx: dict, df_cleaned: pd.DataFrame):
    """
    - xlsx: entire dict of {sheet_name: DataFrame}
    - df_cleaned: cleaned Visitor List DataFrame
    Writes back:
      • All original sheets exactly as they were, EXCEPT
      • Overwrite “Visitor List” with df_cleaned plus formatting/highlighting.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Write every sheet except “Visitor List” exactly as it was:
        for sheet_name, sheet_df in xlsx.items():
            if sheet_name.lower() != "visitor list":
                sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)

        # 2) Write the cleaned “Visitor List”:
        df_cleaned.to_excel(writer, index=False, sheet_name="Visitor List")
        workbook = writer.book
        worksheet = writer.sheets["Visitor List"]

        # ── Formatting “Visitor List” header + cells:
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

        # Apply border/center/font to every cell:
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = center_align
                cell.font = font_style

        # Header row formatting (row 1):
        for col in range(1, worksheet.max_column + 1):
            cell = worksheet[f"{get_column_letter(col)}1"]
            cell.fill = header_fill
            cell.font = bold_font

        # Freeze the top row:
        worksheet.freeze_panes = worksheet["A2"]

        # ── Highlight logic (“Identification Type” vs “Nationality/PR” mismatch):
        mismatch_count = 0
        warning_rows = []
        for r in range(2, worksheet.max_row + 1):
            id_type = str(worksheet[f"G{r}"].value).strip().upper()
            nationality = str(worksheet[f"J{r}"].value).strip().title()
            pr_status = str(worksheet[f"K{r}"].value).strip().title()

            highlight = False
            # If ID=NRIC but (not Singapore AND not PR), highlight
            if id_type == "NRIC" and not (
                nationality == "Singapore" or pr_status in ["Yes", "Pr"]
            ):
                highlight = True
            # If ID=FIN but (Singapore or PR), highlight
            if id_type == "FIN" and (nationality == "Singapore" or pr_status in ["Yes", "Pr"]):
                highlight = True

            if highlight:
                warning_rows.append(r)
                worksheet[f"G{r}"].fill = light_red_fill
                worksheet[f"J{r}"].fill = light_red_fill
                worksheet[f"K{r}"].fill = light_red_fill
                mismatch_count += 1

        # ── Autofit column widths and row heights in “Visitor List”:
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

        # ── Vehicles summary (only if any vehicles exist)
        vehicles = []
        for val in df_cleaned["Vehicle Plate Number"].dropna():
            vehicles.extend([v.strip() for v in str(val).split(";") if v.strip()])

        insert_row = worksheet.max_row + 2
        if vehicles:
            summary = ";".join(sorted(set(vehicles)))
            # “Vehicles” label:
            worksheet[f"B{insert_row}"].value = "Vehicles"
            worksheet[f"B{insert_row}"].border = border
            worksheet[f"B{insert_row}"].alignment = center_align
            # joined list in the next row:
            worksheet[f"B{insert_row + 1}"].value = summary
            worksheet[f"B{insert_row + 1}"].border = border
            worksheet[f"B{insert_row + 1}"].alignment = center_align
            insert_row += 3

        # ── Total Visitors summary:
        total_visitors = df_cleaned["Company Full Name"].notna().sum()
        worksheet[f"B{insert_row}"].value = "Total Visitors"
        worksheet[f"B{insert_row}"].alignment = center_align
        worksheet[f"B{insert_row}"].border = border
        worksheet[f"B{insert_row + 1}"].value = total_visitors
        worksheet[f"B{insert_row + 1}"].alignment = center_align
        worksheet[f"B{insert_row + 1}"].border = border

        # If there are any mismatches, show a warning in Streamlit:
        if warning_rows:
            st.warning(
                f"⚠️ {len(warning_rows)} potential mismatch(es) found in Identification Type vs Nationality/PR. "
                "Please check the highlighted rows in “Visitor List.”"
            )

    return output, mismatch_count

#
# ─── STREAMLIT Uploader & Download Button ───────────────────────────────────────────
#
uploaded_file = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Read **all** sheets into a dict: sheet_name → DataFrame
    xlsx = pd.read_excel(uploaded_file, sheet_name=None)

    # Identify “Visitor List” sheet (case-insensitive search for “visitor”)
    visitor_sheet_name = None
    for name in xlsx.keys():
        if "visitor" in name.lower():
            visitor_sheet_name = name
            break

    if visitor_sheet_name is None:
        st.error("No sheet named “Visitor List” found. Please ensure your file has a “Visitor List” tab.")
    else:
        # Extract just that DataFrame:
        df_raw = xlsx[visitor_sheet_name]
        df_cleaned = clean_data(df_raw)

        # Rebuild the output workbook: pass entire xlsx dict + cleaned df
        output_bytes, mismatch_count = generate_excel(xlsx, df_cleaned)

        # Generate a timestamped filename:
        filename = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        st.success("✅ File cleaned and ready for download!")
        st.download_button(
            label="📥 Download Cleaned Excel File",
            data=output_bytes.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
