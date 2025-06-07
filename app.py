import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🧼 Visitor List Excel Cleaner")

# Provide a download button for the sample template
with open("sample_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download Sample Template",
        data=f.read(),
        file_name="sample_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def nationality_group(row):
    nationality = str(row.get("Nationality (Country Name)", "")).lower()
    pr_status = str(row.get("PR", "")).strip().lower()
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
        idx = name.find(" ")
        return pd.Series([name[:idx], name[idx+1:]])
    return pd.Series([name, ""])


def clean_gender(val):
    v = str(val).strip().upper()
    if v == "M": return "Male"
    if v == "F": return "Female"
    if v in ["MALE", "FEMALE"]: return v.title()
    return val


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        "S/N", "Vehicle Plate Number", "Company Full Name", "Full Name As Per NRIC",
        "First Name as per NRIC", "Middle and Last Name as per NRIC", "Identification Type",
        "IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date",
        "Nationality (Country Name)", "PR", "Gender", "Mobile Number"
    ]
    df = df.dropna(subset=df.columns[3:13], how="all")
    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df.sort_values(by=["Company Full Name","SortGroup","Nationality (Country Name)","Full Name As Per NRIC"], inplace=True)
    df.drop(columns=["SortGroup"], inplace=True)
    df["S/N"] = range(1, len(df) + 1)
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"].astype(str)
        .str.replace(r"[\/\,]",";",regex=True)
        .str.replace(r"\s*;\s*",";",regex=True)
        .str.strip()
        .replace("nan","",regex=False)
    )
    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] = df["Full Name As Per NRIC"].apply(split_name)
    nat_map = {"Chinese":"China","Singaporean":"Singapore"}
    df["Nationality (Country Name)"] = df["Nationality (Country Name)"].replace(nat_map).astype(str).str.title()
    if df["IC (Last 3 digits and suffix) 123A"].astype(str).str.contains("-",na=False).any():
        df[["IC (Last 3 digits and suffix) 123A","Work Permit Expiry Date"]] = df[["Work Permit Expiry Date","IC (Last 3 digits and suffix) 123A"]]
    df["IC (Last 3 digits and suffix) 123A"] = df["IC (Last 3 digits and suffix) 123A"].astype(str).str[-4:]
    df["Mobile Number"] = pd.to_numeric(df["Mobile Number"],errors="coerce").fillna(0).astype(int).astype(str)
    df["Gender"] = df["Gender"].apply(clean_gender)
    df["Work Permit Expiry Date"] = pd.to_datetime(df["Work Permit Expiry Date"],errors="coerce").dt.strftime("%Y-%m-%d")
    return df


def generate_excel(all_sheets: dict, df_cleaned: pd.DataFrame):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # write non-visitor tabs unchanged
        visitor_tab = None
        for name, sheet in all_sheets.items():
            if "visitor" not in name.lower():
                sheet.to_excel(writer,index=False,sheet_name=name)
            else:
                visitor_tab = name
        # overwrite visitor tab
        df_cleaned.to_excel(writer,index=False,sheet_name=visitor_tab)
        ws = writer.book[sheet_name]
        worksheet = writer.sheets[visitor_tab]
        # styling
        header_fill = PatternFill(start_color="94B455",end_color="94B455",fill_type="solid")
        light_red = PatternFill(start_color="FFCCCC",end_color="FFCCCC",fill_type="solid")
        thin = Border(left=Side(style="thin"),right=Side(style="thin"),top=Side(style="thin"),bottom=Side(style="thin"))
        center = Alignment(horizontal="center",vertical="center")
        font = Font(name="Calibri",size=11)
        bold = Font(name="Calibri",size=11,bold=True)
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border=thin;cell.alignment=center;cell.font=font
        for col in range(1,worksheet.max_column+1):
            cell=worksheet[f"{get_column_letter(col)}1"];cell.fill=header_fill;cell.font=bold
        worksheet.freeze_panes=worksheet["A2"]
        mismatches=0;warn_rows=[]
        for r in range(2,worksheet.max_row+1):
            idt=worksheet[f"G{r}"].value;nat=worksheet[f"J{r}"].value;pr=worksheet[f"K{r}"].value
            idt=str(idt).strip().upper(); nat=str(nat).strip().title(); pr=str(pr).strip().title()
            hl=False
            if idt=="NRIC" and not(nat=="Singapore" or pr in ["Yes","Pr"]): hl=True
            if idt=="FIN" and (nat=="Singapore" or pr in ["Yes","Pr"]): hl=True
            if hl:
                warn_rows.append(r);mismatches+=1
                for c in ["G","J","K"]:worksheet[f"{c}{r}"].fill=light_red
        # autofit
        for col in worksheet.columns:
            ml=0;col_letter=get_column_letter(col[0].column)
            for c in col:
                if c.value: ml=max(ml,len(str(c.value)))
            worksheet.column_dimensions[col_letter].width=ml+2
        for row in worksheet.iter_rows():worksheet.row_dimensions[row[0].row].height=20
        # vehicles summary
        veh=[]
        for v in df_cleaned["Vehicle Plate Number"].dropna(): veh.extend([x.strip() for x in str(v).split(";") if x.strip()])
        ir=worksheet.max_row+2
        if veh:
            summ=";".join(sorted(set(veh)))
            worksheet[f"B{ir}"].value="Vehicles";worksheet[f"B{ir}"].border=thin;worksheet[f"B{ir}"].alignment=center
            worksheet[f"B{ir+1}"].value=summ;worksheet[f"B{ir+1}"].border=thin;worksheet[f"B{ir+1}"].alignment=center
            ir+=3
        # total visitors
        tv=df_cleaned["Company Full Name"].notna().sum()
        worksheet[f"B{ir}"].value="Total Visitors";worksheet[f"B{ir}"].border=thin;worksheet[f"B{ir}"].alignment=center
        worksheet[f"B{ir+1}"].value=tv;worksheet[f"B{ir+1}"].border=thin;worksheet[f"B{ir+1}"].alignment=center
        if warn_rows: st.warning(f"⚠️ {len(warn_rows)} mismatch(es) found.")
    return output, mismatches

# main
uploaded_file=st.file_uploader("📁 Upload your Excel file",type=["xlsx"])
if uploaded_file:
    xlsx=pd.read_excel(uploaded_file,sheet_name=None)
    vt=None
    for n in xlsx.keys():
        if "visitor" in n.lower(): vt=n;break
    if not vt:
        st.error("Upload must contain a sheet with 'visitor' in its name.")
    else:
        raw=xlsx[vt]
        cleaned=clean_data(raw)
        out,mc=generate_excel(xlsx,cleaned)
        fn=f"Cleaned_Visitor_List_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        st.download_button("📥 Download Cleaned Excel File",data=out.getvalue(),file_name=fn,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
