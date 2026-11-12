import os
import re
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright


def extract_description_point_1(text):
    """Extracts Point 1 from Description and removes the '*1. WCAG SC:*' prefix."""
    if pd.isna(text):
        return ""
    text = str(text).strip().lstrip('"')
    match = re.search(
        r'^[\s\*_#]*1\.?\s*.*?(?=\n\s*[\*_#]*2\.|\n\s*[\*_#]*[A-Za-z][A-Za-z\s]+:[\*_]*|\Z)',
        text,
        flags=re.DOTALL
    )
    if not match:
        return ""
    point1 = match.group(0).strip()
    # Remove '*1. WCAG SC:*' (including variations like *_1. WCAG SC_*_:_, **1. WCAG SC:**, etc.)
    cleaned = re.sub(
        r'^[\s\*_#]*1\.?\s*WCAG\s*SC[\s\*_:]*',
        '',
        point1,
        flags=re.IGNORECASE
    ).strip()
    return cleaned


def check_sc_match(success_criteria, desc_val):
    """Checks if Success Criteria matches the cleaned Description value."""
    if pd.isna(success_criteria) or not desc_val:
        return "No"

    sc_str = str(success_criteria).strip()
    sc_num = re.sub(r'^A{1,3}-', '', sc_str, flags=re.IGNORECASE).strip()
    clean_desc = str(desc_val).strip()

    if sc_str.lower() == clean_desc.lower() or sc_num.lower() == clean_desc.lower():
        return "Yes"
    return "No"


def extract_acceptance_criteria(text):
    """Splits Custom field (Acceptance Criteria) into 4 separate fields."""
    if pd.isna(text):
        return ("", "", "", "")
    text = str(text)

    p1 = re.search(
        r'(?:[\*_#]+\s*1\.?\s*(?:Actu[a-z]*\s*:?)?[\*_]*|1\.\s*Actual\s*:?)\s*(.*?)(?=[\*_#]*2\.?\s*Expected\s*:?[\*_]*|#\s*[\*_]*Expected\s*:?[\*_]*|\Z)',
        text,
        flags=re.DOTALL | re.IGNORECASE
    )
    p2 = re.search(
        r'(?:[\*_#]*2\.?\s*Expected\s*:?[\*_]*|#\s*[\*_]*Expected\s*:?[\*_]*)\s*(.*?)(?=[\*_#]*3\.?\s*(?:Rationale|User Impact)\s*:?[\*_]*|#\s*[\*_]*(?:Rationale|User Impact)\s*:?[\*_]*|\Z)',
        text,
        flags=re.DOTALL | re.IGNORECASE
    )
    p3 = re.search(
        r'(?:[\*_#]*3\.?\s*(?:Rationale|User Impact)\s*:?[\*_]*|#\s*[\*_]*(?:Rationale|User Impact)\s*:?[\*_]*)\s*(.*?)(?=[\*_#]*[34]\.?\s*Recommendation\s*:?[\*_]*|#\s*[\*_]*Recommendation\s*:?[\*_]*|\Z)',
        text,
        flags=re.DOTALL | re.IGNORECASE
    )
    p4 = re.search(
        r'(?:[\*_#]*[34]\.?\s*Recommendation\s*:?[\*_]*|#\s*[\*_]*Recommendation\s*:?[\*_]*)\s*(.*?)(?=[\*_#]*5\.|[\*_#]*(?:Reference|Affected Code|Suggested Code|Note)\s*:?[\*_]*|\Z)',
        text,
        flags=re.DOTALL | re.IGNORECASE
    )

    return (
        p1.group(1).strip() if p1 else "",
        p2.group(1).strip() if p2 else "",
        p3.group(1).strip() if p3 else "",
        p4.group(1).strip() if p4 else ""
    )


def style_worksheet(ws):
    """Applies clean header styling, text wrapping, frozen top row, and filters."""
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F4F7FA", end_color="F4F7FA", fill_type="solid")
    yes_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    no_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    # Column widths for: Issue key, Issue id, Success Criteria, Description, Matched, Actual, Expected, Rationale, Recommendation
    col_widths = [18, 15, 22, 22, 14, 55, 55, 55, 55]
    for col_idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        for cell in row:
            cell.border = thin_border
            if row_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                if row_idx % 2 == 0:
                    cell.fill = zebra_fill
                cell.font = Font(name="Calibri", size=10)
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

                # Highlight Yes/No in the 'Matched' column (column 5)
                if cell.column == 5:
                    cell.alignment = Alignment(horizontal="center", vertical="top")
                    if cell.value == "Yes":
                        cell.fill = yes_fill
                        cell.font = Font(name="Calibri", size=10, bold=True, color="375623")
                    elif cell.value == "No":
                        cell.fill = no_fill
                        cell.font = Font(name="Calibri", size=10, bold=True, color="C00000")


def process_jira_excel(input_excel_path):
    if not os.path.exists(input_excel_path):
        raise FileNotFoundError(f"File not found: {input_excel_path}")

    folder = os.path.dirname(input_excel_path)
    base_name = os.path.splitext(os.path.basename(input_excel_path))[0]
    output_excel_path = os.path.join(folder, f"{base_name}_output.xlsx")

    # Read Excel in its original row order
    df = pd.read_excel(input_excel_path)

    # Extract Description Point 1 without '*1. WCAG SC:*'
    desc_p1 = df["Description"].apply(extract_description_point_1)

    # Check if Success Criteria and Description match (Yes/No)
    matched_col = [
        check_sc_match(sc, d1)
        for sc, d1 in zip(df["Success Criteria"], desc_p1)
    ]

    # Extract Acceptance Criteria into 4 separate columns
    ac_extracted = df["Custom field (Acceptance Criteria)"].apply(extract_acceptance_criteria)

    # Build single DataFrame with Issue key and Issue id included
    out_df = pd.DataFrame({
        "Issue key": df["Issue key"],
        "Issue id": df["Issue id"],
        "Success Criteria": df["Success Criteria"],
        "Description": desc_p1,
        "Matched": matched_col,
        "*1. Actual:*": [item[0] for item in ac_extracted],
        "*2. Expected:*": [item[1] for item in ac_extracted],
        "*3. Rationale:*": [item[2] for item in ac_extracted],
        "*4. Recommendation:*": [item[3] for item in ac_extracted],
    })

    # Write to a single sheet in one Excel file
    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="Output", index=False)
        style_worksheet(writer.book["Output"])

    print(f"Done! Output Excel saved at:\n{output_excel_path}")
    print(f"Total rows processed (original order preserved): {len(out_df)}")
    print(f"Matched counts: {out_df['Matched'].value_counts().to_dict()}")
    return output_excel_path


if __name__ == "__main__":
    manual_excel_path = input("Enter the full path to the Excel file: ").strip().strip('"')
    process_jira_excel(manual_excel_path)