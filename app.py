import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f,
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def nationality_group(row):
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
    name = str(name).strip()
    if " " in name:
        first_space = name.find(" ")
        return pd.Series([name[:first_space], name[first_space+1:]])
    return pd.Series([name, ""])

def clean_gender(val):
    val = str(val).strip().upper()
    if val == "M":
        return "Male"
    elif val == "F":
        return "Female"
    elif val in ["MALE", "FEMALE"]:
        return val.title()
    return val

def clean_data(df):
    df.columns = [
        "S/N", "Vehicle Plate Number", "Company Full Name", "Full Name As Per NRIC",
        "First Name as per NRIC", "Middle and Last Name as per NRIC", "Identification Type",
        "IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date",
        "Nationality (Country Name)", "PR", "Gender", "Mobile Number"
    ]

    df = df.dropna(subset=df.columns[3:13], how="all")

    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df.sort_values(by=["Company Full Name", "SortGroup", "Nationality (Country Name)", "Full Name As Per NRIC"], inplace=True)
    df.drop(columns=["SortGroup"], inplace=True)

    df["S/N"] = range(1, len(df) + 1)

    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
        .str.replace(r"[\/\,]", ";", regex=True)
        .str.replace(r"\s*;\s*", ";", regex=True)
        .str.strip()
        .replace("nan", "", regex=False)
    )

    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC", "Middle and Last Name as per NRIC"]] = df["Full Name As Per NRIC"].apply(split_name)

    nationality_map = {"Chinese": "China", "Singaporean": "Singapore"}
    df["Nationality (Country Name)"] = df["Nationality (Country Name)"].replace(nationality_map)
    df["Nationality (Country Name)"] = df["Nationality (Country Name)"].astype(str).str.title()

    # Ensure columns are in correct order if swapped (Work Permit Expiry and IC Suffix)
    if df["IC (Last 3 digits and suffix) 123A"].str.contains("-", na=False).any():
        df[["IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date"]] = df[["Work Permit Expiry Date", "IC (Last 3 digits and suffix) 123A"]]

    df["IC (Last 3 digits and suffix) 123A"] = df["IC (Last 3 digits and suffix) 123A"].astype(str).str[-4:]
    df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(" ", "", regex=False)
    df["Gender"] = df["Gender"].apply(clean_gender)

    df["Work Permit Expiry Date"] = pd.to_datetime(df["Work Permit Expiry Date"], errors="coerce").dt.strftime("%Y-%m-%d")

    return df

def generate_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        workbook = writer.book
        worksheet = writer.sheets["Visitor List"]

        header_fill = PatternFill(start_color="FEC100", end_color="FEC100", fill_type="solid")
        light_red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        center_align = Alignment(horizontal="center", vertical="center")
        font_style = Font(name="Calibri", size=11)
        bold_font = Font(name="Calibri", size=11, bold=True)

        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = center_align
                cell.font = font_style

        for col in range(1, worksheet.max_column + 1):
            cell = worksheet[f"{get_column_letter(col)}1"]
            cell.fill = header_fill
            cell.font = bold_font

        worksheet.freeze_panes = worksheet["A2"]

        mismatch_count = 0
        warning_rows = []
        for row in range(2, worksheet.max_row + 1):
            id_type = str(worksheet[f"G{row}"].value).strip().upper()
            nationality = str(worksheet[f"J{row}"].value).strip().title()
            pr_status = str(worksheet[f"K{row}"].value).strip().title()

            highlight = False
            if id_type == "NRIC" and not (nationality == "Singapore" or (nationality != "Singapore" and pr_status in ["Yes", "Pr"])):
                highlight = True
            if id_type == "FIN" and (nationality == "Singapore" or pr_status in ["Yes", "Pr"]):
                highlight = True

            if highlight:
                warning_rows.append(row)
                worksheet[f"G{row}"].fill = light_red_fill
                worksheet[f"J{row}"].fill = light_red_fill
                worksheet[f"K{row}"].fill = light_red_fill
                mismatch_count += 1

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

        total_visitors = df["Company Full Name"].notna().sum()
        worksheet[f"B{insert_row}"].value = "Total Visitors"
        worksheet[f"B{insert_row}"].alignment = center_align
        worksheet[f"B{insert_row}"].border = border
        worksheet[f"B{insert_row + 1}"].value = total_visitors
        worksheet[f"B{insert_row + 1}"].alignment = center_align
        worksheet[f"B{insert_row + 1}"].border = border

            if warning_rows:
            st.warning(f"⚠️ {len(warning_rows)} potential mismatch(es) found in Identification Type and Nationality/PR. Please check highlighted rows.")} potential mismatch(es) found in Identification Type and Nationality/PR. Please check highlighted rows.")

    return output, mismatch_count

uploaded_file = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="Visitor List")
    df_cleaned = clean_data(df)
    output, mismatch_count = generate_excel(df_cleaned)

    filename = f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        label="📥 Download Cleaned Excel File",
        data=output.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
