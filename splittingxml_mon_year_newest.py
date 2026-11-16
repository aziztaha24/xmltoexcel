"""
Invoice QC Toolkit — unified Streamlit app
Combines all 4 pipeline steps used by finance:
  1. XML → XML (split master XML by company)
  2. XML → Excel (single consolidated data file)
  3. Excel + NetSuite CSV merger
  4. Excel split by company (with formulas + conditional formatting)
"""

import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import io
import zipfile
import copy
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule


# ────────────────────────────────────────────────────────────────────────────
# PAGE SETUP
# ────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Invoice QC Toolkit",
    page_icon="🧾",
    layout="wide",
)

st.title("🧾 Invoice QC Toolkit")
st.caption("Full XML → Excel pipeline for finance. Each tab is one step in the workflow.")


# ────────────────────────────────────────────────────────────────────────────
# SHARED HELPERS
# ────────────────────────────────────────────────────────────────────────────
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

EXPECTED_COLUMNS = [
    "Name", "Identifier", "FirstName", "MiddleName", "LastName",
    "PlanCost", "EmploymentStatus", "HireDate", "HiredOn",
    "TerminationDate", "TerminatedOn", "StartDate",
    "EnrolledOn", "EndDate", "EndedOn", "CoverageLevel",
    "CarrierPlanCode", "PriorCoverageStartDate"
]


def month_year_inputs(key_prefix: str):
    """Render month + year inputs side by side. Returns (month, year)."""
    current_month_idx = datetime.now().month - 1
    current_year = datetime.now().year
    c1, c2 = st.columns(2)
    with c1:
        month = st.selectbox(
            "Month",
            MONTHS,
            index=current_month_idx,
            key=f"{key_prefix}_month",
        )
    with c2:
        year = st.text_input(
            "Year",
            value=str(current_year),
            key=f"{key_prefix}_year",
        )
    return month, year.strip()


# ────────────────────────────────────────────────────────────────────────────
# STEP 1 — XML → XML (split by company)
# ────────────────────────────────────────────────────────────────────────────
def split_xml_by_company(xml_bytes: bytes, month: str, year: str):
    tree = ET.parse(io.BytesIO(xml_bytes))
    root = tree.getroot()

    header_elements = root.findall(".//Header")
    companies = root.findall(".//Company")

    zip_buffer = io.BytesIO()
    files_created = []
    skipped = 0

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for company in companies:
            name_elem = company.find(".//Name")
            if name_elem is None or not name_elem.text:
                skipped += 1
                continue

            company_name = name_elem.text.replace("/", "_").replace("\\", "_")

            new_root = ET.Element(root.tag)
            for header in header_elements:
                new_root.append(copy.deepcopy(header))
            new_root.append(copy.deepcopy(company))

            xml_str = ET.tostring(new_root, encoding="utf-8", xml_declaration=True)
            filename = f"{company_name} ({month}-{year}).xml"
            zf.writestr(filename, xml_str)
            files_created.append(filename)

    return zip_buffer.getvalue(), files_created, skipped


# ────────────────────────────────────────────────────────────────────────────
# STEP 2 — XML → Excel (single data file)
# ────────────────────────────────────────────────────────────────────────────
import pandas as pd
import xml.etree.ElementTree as ET

# === FILE PATHS ===
xml_file = r"C:\Users\LENOVO\OneDrive\Documents\REDH\Invoice QC\2026\9. Sept 2026\EN XML & XLS Data Sep-2026\Redirect_Health_Production_20260816_031142_1121345.xml"
output_file = r"C:\Users\LENOVO\OneDrive\Documents\REDH\Invoice QC\2026\9. Sept 2026\EN XML & XLS Data Sep-2026\Single File before merger.xlsx"
# === EXPECTED COLUMNS ===
EXPECTED_COLUMNS = [
    "Name", "Identifier", "FirstName", "MiddleName", "LastName",
    "PlanCost", "EmploymentStatus", "HireDate", "HiredOn",
    "TerminationDate", "TerminatedOn", "StartDate",
    "EnrolledOn", "EndDate", "EndedOn", "CoverageLevel",
    "CarrierPlanCode", "PriorCoverageStartDate"
]

# === FUNCTION ===
def xml_to_exact_excel(xml_file_path, output_excel_path):
    context = ET.iterparse(xml_file_path, events=("start", "end"))
    data_list = []
    header_info = {}

    for event, elem in context:
        if event == "end":
            if elem.tag == "Header":
                header_info = {
                    "Disclaimer": elem.findtext("Disclaimer", ""),
                    "ExchangeName": elem.findtext("ExchangeName", ""),
                    "VendorName": elem.findtext("VendorName", ""),
                    "RunDate": elem.findtext("RunDate", ""),
                }

            elif elem.tag == "Company":
                company_data = {
                    "Identifier": elem.findtext("Identifier", ""),
                    "Name": elem.findtext("Name", ""),
                }
                full_data = {**header_info, **company_data}
                employees = elem.findall("Employees/Employee")
                first_employee = True

                for emp in employees:
                    emp_data = {
                        "FirstName": emp.findtext("FirstName", ""),
                        "MiddleName": emp.findtext("MiddleName", ""),
                        "LastName": emp.findtext("LastName", ""),
                        "EmploymentStatus": emp.findtext("EmploymentStatus", ""),
                        "HireDate": emp.findtext("HireDate", ""),
                        "HiredOn": emp.findtext("HiredOn", ""),
                        "TerminationDate": emp.findtext("TerminationDate", ""),
                        "TerminatedOn": emp.findtext("TerminatedOn", ""),
                    }

                    enrollments = emp.findall("Enrollments/Enrollment")
                    if enrollments:
                        for enroll in enrollments:
                            enroll_data = {
                                "PlanCost": enroll.findtext("PlanCost", ""),
                                "StartDate": enroll.findtext("StartDate", ""),
                                "EnrolledOn": enroll.findtext("EnrolledOn", ""),
                                "EndDate": enroll.findtext("EndDate", ""),
                                "EndedOn": enroll.findtext("EndedOn", ""),
                                "CoverageLevel": enroll.findtext("CoverageLevel", ""),
                                "CarrierPlanCode": enroll.findtext("CarrierPlanCode", ""),
                                "PriorCoverageStartDate": enroll.findtext("PriorCoverageStartDate", ""),
                            }

                            if first_employee:
                                final_data = {**full_data, **emp_data, **enroll_data}
                                first_employee = False
                            else:
                                final_data = {**{k: "" for k in full_data}, **emp_data, **enroll_data}

                            cleaned_data = {col: final_data.get(col, "") for col in EXPECTED_COLUMNS}
                            data_list.append(cleaned_data)
                    else:
                        if first_employee:
                            final_data = {**full_data, **emp_data}
                            first_employee = False
                        else:
                            final_data = {**{k: "" for k in full_data}, **emp_data}

                        cleaned_data = {col: final_data.get(col, "") for col in EXPECTED_COLUMNS}
                        data_list.append(cleaned_data)
                elem.clear()

    df = pd.DataFrame(data_list, columns=EXPECTED_COLUMNS)
    df.to_excel(output_excel_path, index=False)
    print(f"✅ Excel file saved at: {output_excel_path}")

# === EXECUTE ===
xml_to_exact_excel(xml_file, output_file)


# ────────────────────────────────────────────────────────────────────────────
# STEP 3 — Excel + CSV merger
# ────────────────────────────────────────────────────────────────────────────

import pandas as pd

# === FILE PATHS ===
master_file = r"C:\Users\LENOVO\OneDrive\Documents\REDH\Invoice QC\2026\11.Nov 2026\EN XML & XLS Data Nov-2026\Single Data file before merger.xlsx"
cost_file = r"C:\Users\LENOVO\OneDrive\Documents\REDH\Invoice QC\2026\11.Nov 2026\EN XML & XLS Data Nov-2026\NetSuiteInvoiceReportMonthly_20261017.csv"
output_file = r"C:\Users\LENOVO\OneDrive\Documents\REDH\Invoice QC\2026\11.Nov 2026\EN XML & XLS Data Nov-2026\Single Data after merger.xlsx"
master_df = pd.read_excel(master_file)
cost_df = pd.read_csv(cost_file, encoding='latin1')

# === NORMALIZE FUNCTION ===
def normalize(val):
    return str(val).strip().lower()

# === RENAME COST FILE COLUMNS TO MATCH MASTER ===
cost_df.rename(columns={
    "Member First Name": "FirstName",
    "Member Last Name": "LastName",
    "Group Number": "Identifier"
}, inplace=True)

# === NORMALIZE COLUMNS FOR MATCHING ===
for col in ['FirstName', 'LastName', 'Identifier']:
    if col in master_df.columns:
        master_df[col] = master_df[col].apply(normalize)
    if col in cost_df.columns:
        cost_df[col] = cost_df[col].apply(normalize)

# === MERGE ON FIRSTNAME, LASTNAME, IDENTIFIER ===

merged_df = master_df.merge(
    cost_df[['FirstName', 'LastName', 'Identifier', 'Premium ($)', 'Coverage/Plan Type/Plan Tier', 'Plan SKU']],
    on=['FirstName', 'LastName', 'Identifier'],
    how='left'
)

# === ASSIGN FIELDS ===
merged_df['Charge'] = merged_df['Premium ($)']
merged_df['CoverageCSV'] = merged_df['Coverage/Plan Type/Plan Tier'].fillna('0')
merged_df['PlanSKUCSV'] = merged_df['Plan SKU'].fillna ('0')

merged_df['Identifier'] = merged_df['Identifier'].str.upper()

# === DROP UNUSED COLUMNS ===
merged_df.drop(columns=['Premium ($)', 'Coverage/Plan Type/Plan Tier', 'Plan SKU'], inplace=True)

# === SAVE TO EXCEL ===
merged_df.to_excel(output_file, index=False)
print(f"✅ File saved: {output_file}")

# ────────────────────────────────────────────────────────────────────────────
# STEP 4 — Excel split by company (with formulas + duplicates highlight)
# ────────────────────────────────────────────────────────────────────────────

import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule

def split_excel_by_name(excel_file, month, year):
    df = pd.read_excel(excel_file)

    if 'Name' not in df.columns:
        raise ValueError("Column 'Name' not found in the Excel file.")

    output_folder = f"split_excel_{month}_{year}"
    os.makedirs(output_folder, exist_ok=True)

    for name, group in df.groupby('Name'):
        if pd.isna(name):
            print("Name missing in some rows, skipping...")
            continue

        safe_name = str(name).replace("/", "_").replace("\\", "_")
        output_file = os.path.join(output_folder, f"{safe_name} ({month}-{year}).xlsx")

        # Ensure 'Diff' column exists
        if 'Diff' not in group.columns:
            group['Diff'] = ''  # Placeholder for formula

        group.to_excel(output_file, index=False)

        wb = load_workbook(output_file)
        ws = wb.active

        # Freeze header row and apply autofilter
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        # Get column headers and their Excel column letters
        headers = {cell.value: cell.column for cell in ws[1]}

        # Add formula in 'Diff' column
        if 'PlanCost' in headers and 'Charge' in headers and 'Diff' in headers:
            plan_cost_col = get_column_letter(headers['PlanCost'])
            charge_col = get_column_letter(headers['Charge'])
            diff_col = get_column_letter(headers['Diff'])

            for row in range(2, ws.max_row + 1):
                formula = f"={plan_cost_col}{row}-{charge_col}{row}"
                ws[f"{diff_col}{row}"] = formula

        # Apply conditional formatting to highlight duplicates in 'Lookup Criteria'
        if 'Lookup Criteria' in headers:
            lookup_col_letter = get_column_letter(headers['Lookup Criteria'])
            start_row = 2
            end_row = ws.max_row
            data_range = f"${lookup_col_letter}${start_row}:${lookup_col_letter}${end_row}"

            # COUNTIF formula to check for duplicates
            formula = f'=COUNTIF({data_range}, {lookup_col_letter}{start_row})>1'

            duplicate_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            rule = FormulaRule(formula=[formula], fill=duplicate_fill)
            ws.conditional_formatting.add(data_range, rule)

        wb.save(output_file)
        print(f"Created: {output_file}")

if __name__ == "__main__":
    excel_file = r"C:\Users\LENOVO\OneDrive\Documents\REDH\Invoice QC\2026\8. Aug 2026\EN XML & XLS Data Aug-2026\Single Data File After Merger.xlsx"
    month = input("Enter the month (e.g., Jan, Feb, etc.): ").strip()
    year = input("Enter the year (e.g., 2024): ").strip()

    split_excel_by_name(excel_file, month, year)

# ────────────────────────────────────────────────────────────────────────────
# UI — TABS
# ────────────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "1️⃣  XML → XML (split by company)",
    "2️⃣  XML → Excel (single data file)",
    "3️⃣  Excel + CSV merger",
    "4️⃣  Excel split by company",
])


# ── TAB 1 ───────────────────────────────────────────────────────────────────
with tab1:
    st.subheader("Split master XML into one XML per company")
    st.markdown(
        "Upload the master XML. Each company is written to its own XML file "
        "(with the Header preserved). All files are bundled into a single ZIP."
    )

    month1, year1 = month_year_inputs("tab1")
    xml_file_1 = st.file_uploader("Upload master XML", type=["xml"], key="xml1")

    if xml_file_1 and month1 and year1:
        try:
            with st.spinner("Splitting XML by company…"):
                zip_bytes, files, skipped = split_xml_by_company(
                    xml_file_1.read(), month1, year1
                )

            st.success(f"✅ Created {len(files)} XML file(s).")
            if skipped:
                st.warning(f"Skipped {skipped} company entries with missing Name.")

            with st.expander("Files in ZIP"):
                for f in files:
                    st.text(f)

            st.download_button(
                label="⬇️ Download ZIP of XML files",
                data=zip_bytes,
                file_name=f"Split_XML_{month1}_{year1}.zip",
                mime="application/zip",
            )
        except ET.ParseError as e:
            st.error(f"XML parse error: {e}")
        except Exception as e:
            st.error(f"Something went wrong: {e}")


# ── TAB 2 ───────────────────────────────────────────────────────────────────
with tab2:
    st.subheader("Convert master XML to a single Excel data file")
    st.markdown(
        "Flattens Header → Company → Employees → Enrollments into a tabular Excel "
        "with the standard column set."
    )

    xml_file_2 = st.file_uploader("Upload master XML", type=["xml"], key="xml2")

    if xml_file_2:
        try:
            with st.spinner("Parsing XML and building Excel…"):
                df = xml_to_single_excel(xml_file_2.read())

            st.success(f"✅ Extracted {len(df):,} row(s).")
            st.write("Preview:")
            st.dataframe(df.head(20))

            st.download_button(
                label="⬇️ Download Excel (Single Data file before merger)",
                data=df_to_excel_bytes(df),
                file_name="Single Data file before merger.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except ET.ParseError as e:
            st.error(f"XML parse error: {e}")
        except Exception as e:
            st.error(f"Something went wrong: {e}")


# ── TAB 3 ───────────────────────────────────────────────────────────────────
with tab3:
    st.subheader("Merge the Excel data file with the monthly NetSuite CSV")
    st.markdown(
        "Match keys: **FirstName + LastName + Identifier** (case/space normalized). "
        "Adds `Charge`, `CoverageCSV`, `PlanSKUCSV` from the CSV."
    )

    c1, c2 = st.columns(2)
    with c1:
        master_xlsx = st.file_uploader(
            "Upload 'Single Data file before merger.xlsx'",
            type=["xlsx"],
            key="master3",
        )
    with c2:
        cost_csv = st.file_uploader(
            "Upload NetSuite monthly CSV",
            type=["csv"],
            key="csv3",
        )

    if master_xlsx and cost_csv:
        try:
            with st.spinner("Merging…"):
                merged_df = merge_excel_and_csv(master_xlsx.read(), cost_csv.read())

            matched = merged_df["Charge"].notna().sum()
            unmatched = merged_df["Charge"].isna().sum()
            st.success(f"✅ Merged. Rows: {len(merged_df):,} | matched: {matched:,} | unmatched: {unmatched:,}")

            st.write("Preview:")
            st.dataframe(merged_df.head(20))

            st.download_button(
                label="⬇️ Download merged Excel (XML.xlsx)",
                data=df_to_excel_bytes(merged_df),
                file_name="XML.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Something went wrong: {e}")


# ── TAB 4 ───────────────────────────────────────────────────────────────────
with tab4:
    st.subheader("Split the merged Excel into one file per company")
    st.markdown(
        "Each company file gets: frozen header, autofilter, `Diff = PlanCost − Charge` formula, "
        "and duplicate highlighting in **Lookup Criteria** (if the column exists)."
    )

    month4, year4 = month_year_inputs("tab4")
    merged_xlsx = st.file_uploader("Upload merged Excel (XML.xlsx)", type=["xlsx"], key="merged4")

    if merged_xlsx and month4 and year4:
        try:
            with st.spinner("Splitting by company and applying formatting…"):
                zip_bytes, files, skipped = split_excel_by_name(
                    merged_xlsx.read(), month4, year4
                )

            st.success(f"✅ Created {len(files)} Excel file(s).")
            if skipped:
                st.warning(f"Skipped {skipped} rows with missing Name.")

            with st.expander("Files in ZIP"):
                for f in files:
                    st.text(f)

            st.download_button(
                label="⬇️ Download ZIP of per-company Excel files",
                data=zip_bytes,
                file_name=f"split_excel_{month4}_{year4}.zip",
                mime="application/zip",
            )
        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Something went wrong: {e}")


# ────────────────────────────────────────────────────────────────────────────
# FOOTER
# ────────────────────────────────────────────────────────────────────────────
st.divider()
st.caption("Pipeline order: 1 → 2 → 3 → 4. Each step's output feeds the next.")
