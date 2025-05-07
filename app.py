
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl.styles import PatternFill
from openpyxl import load_workbook
from openpyxl.writer.excel import save_virtual_workbook

st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("🔧 Visitor List Excel Cleaner")

# --- Step 0: Offer Sample Template ---
st.markdown("**Download a blank template** to ensure your sheet has the correct headers before uploading.")
with open("sample_template.xlsx", "rb") as f:
    st.download_button("📥 Download Template", f, file_name="sample_template.xlsx")

st.markdown("---")

# --- Step 1: Upload ---
st.header("Step 1: Upload Your Excel File")
uploaded = st.file_uploader("Upload an .xlsx with sheet named **Visitor List**", type=["xlsx"])
if not uploaded:
    st.info("Awaiting file upload…")
    st.stop()

# --- Read & Preview Original ---
df = pd.read_excel(uploaded, sheet_name="Visitor List")
st.subheader("Preview: Original Data")
st.dataframe(df.head(), height=200)

st.markdown("""
**Validation Rules**  
- **Column G** vs **Column J**:  
  - If ID Type = “NRIC” → Nationality must be “Singapore”  
  - If ID Type = “FIN” → Nationality must NOT be “Singapore”  
- **Column D**: Proper case (like `=PROPER()`)  
- **Columns E/F**: Split at first space so E+F exactly equals D  
- **Column B**: Replace `/` or `,` with `;`, no spaces around `;`  
- **Column H**: Keep only last 4 characters  
- **Column J**: Map “Chinese”→“China”, “Singaporean”→“Singapore”, then proper case  
- **Column L**: “M”→“Male”, “F”→“Female”, then proper case  
- **Column M**: Remove all spaces  
""")

# --- Cleaning Logic ---
df = df.copy()
df.columns = [
    "S/N", "Vehicle Plate Number", "Company Full Name", "Full Name As Per NRIC",
    "First Name as per NRIC", "Middle and Last Name as per NRIC", "Identification Type",
    "IC (Last 3 digits and suffix)", "Work Permit Expiry Date",
    "Nationality", "PR", "Gender", "Mobile Number"
]

# B: Vehicle plates
df["Vehicle Plate Number"] = (
    df["Vehicle Plate Number"].astype(str)
    .str.replace(r"[\/,]", ";", regex=True)
    .str.replace(r"\s*;\s*", ";", regex=True)
    .str.replace(r"^nan$", "", regex=True)
)

# D: Proper case full name
df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()

# E/F: Split names
def split_name(x):
    x = str(x).strip()
    if " " in x:
        i = x.find(" ")
        return x[:i], x[i+1:]
    return x, ""
df[["First Name as per NRIC","Middle and Last Name as per NRIC"]] =     df["Full Name As Per NRIC"].apply(lambda x: pd.Series(split_name(x)))

# J: Nationality map + proper case
nat_map = {"Chinese":"China","Singaporean":"Singapore"}
df["Nationality"] = df["Nationality"].replace(nat_map).astype(str).str.title()

# H: Last 4 chars of IC
df["IC (Last 3 digits and suffix)"] = df["IC (Last 3 digits and suffix)"].astype(str).str[-4:]

# L: Gender cleanup
def clean_gender(v):
    v = str(v).strip().upper()
    return "Male" if v in ["M","MALE"] else "Female" if v in ["F","FEMALE"] else v.title()
df["Gender"] = df["Gender"].apply(clean_gender)

# M: Mobile no spaces
df["Mobile Number"] = df["Mobile Number"].astype(str).str.replace(" ", "", regex=False)

# Prepare Excel in memory
output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Visitor List")
    wb = writer.book
    ws = writer.sheets["Visitor List"]
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Validation highlights
    mismatches = 0
    for r in range(2, ws.max_row+1):
        idt = str(ws[f"G{r}"].value).strip().upper()
        nat = str(ws[f"J{r}"].value).strip().title()
        bad = (idt=="NRIC" and nat!="Singapore") or (idt=="FIN" and nat=="Singapore")
        if bad:
            ws[f"G{r}"].fill = yellow
            ws[f"J{r}"].fill = yellow
            mismatches += 1

    # Vehicle summary
    all_vehicles = []
    for v in df["Vehicle Plate Number"].dropna():
        all_vehicles += [x.strip() for x in v.split(";") if x.strip()]
    summary = ";".join(all_vehicles)
    nxt = ws.max_row + 2
    ws[f"B{nxt}"] = "Vehicles"
    ws[f"B{nxt+1}"] = summary

# --- After Cleaning Preview & Feedback ---
st.subheader("Preview: Cleaned Data")
st.dataframe(df.head(), height=200)

st.markdown(f"**✅ Total rows flagged for ID/Nationality mismatch:** {mismatches}")

# --- Download Button ---
fn = f"Cleaned_VisitorList_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
st.download_button(
    "📥 Download Cleaned File",
    data=output.getvalue(),
    file_name=fn,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
