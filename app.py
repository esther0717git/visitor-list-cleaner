def generate_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Write only the cleaned "Visitor List" sheet
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        workbook  = writer.book
        worksheet = writer.sheets["Visitor List"]

        # — styling boilerplate —
        header_fill    = PatternFill(start_color="94B455", end_color="94B455", fill_type="solid")
        light_red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        border         = Border(left=Side("thin"), right=Side("thin"),
                                top=Side("thin"),  bottom=Side("thin"))
        center_align   = Alignment(horizontal="center", vertical="center")
        font_style     = Font(name="Calibri", size=11)
        bold_font      = Font(name="Calibri", size=11, bold=True)

        # Apply borders, alignment & font to all cells
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border    = border
                cell.alignment = center_align
                cell.font      = font_style

        # Header row styling
        for col in range(1, worksheet.max_column+1):
            h = worksheet[f"{get_column_letter(col)}1"]
            h.fill = header_fill
            h.font = bold_font

        # Freeze header
        worksheet.freeze_panes = "A2"

        # Validation highlight pass
        warning_rows = []
        for r in range(2, worksheet.max_row+1):
            idt = str(worksheet[f"G{r}"].value).strip().upper()
            nat = str(worksheet[f"J{r}"].value).strip().title()
            pr  = str(worksheet[f"K{r}"].value).strip().title()

            bad = (
                (idt == "NRIC" and not (nat == "Singapore" or pr in ["Yes","Pr"])) or
                (idt == "FIN"  and (nat == "Singapore" or pr in ["Yes","Pr"]))
            )
            if bad:
                warning_rows.append(r)
                for c in ("G","J","K"):
                    ws = worksheet[f"{c}{r}"]
                    ws.fill = light_red_fill

        # Auto-fit columns & rows
        for col in worksheet.columns:
            ml = max((len(str(cell.value or "")) for cell in col), default=0)
            worksheet.column_dimensions[get_column_letter(col[0].column)].width = ml + 2
        for row in worksheet.iter_rows():
            worksheet.row_dimensions[row[0].row].height = 20

        # Vehicles summary
        vals = (
            df["Vehicle Plate Number"]
            .dropna()
            .astype(str)
            .str.split(";")
            .explode()
            .str.strip()
        )
        unique_v = sorted(vals[vals != ""].unique())
        ir = worksheet.max_row + 2
        if unique_v:
            worksheet[f"B{ir}"].value     = "Vehicles"
            worksheet[f"B{ir}"].border    = border
            worksheet[f"B{ir}"].alignment = center_align
            worksheet[f"B{ir+1}"].value     = ";".join(unique_v)
            worksheet[f"B{ir+1}"].border    = border
            worksheet[f"B{ir+1}"].alignment = center_align
            ir += 3

        # Total visitors
        total = df["Company Full Name"].notna().sum()
        worksheet[f"B{ir}"].value     = "Total Visitors"
        worksheet[f"B{ir}"].border    = border
        worksheet[f"B{ir}"].alignment = center_align
        worksheet[f"B{ir+1}"].value     = total
        worksheet[f"B{ir+1}"].border    = border
        worksheet[f"B{ir+1}"].alignment = center_align

        # Optional warning message in UI
        if warning_rows:
            st.warning(f"⚠️ {len(warning_rows)} potential mismatch(es) found. Please review highlighted rows.")

    return output
