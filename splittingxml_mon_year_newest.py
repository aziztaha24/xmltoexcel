# import streamlit as st
# import pandas as pd
# import xml.etree.ElementTree as ET
# import tempfile



# EXPECTED_COLUMNS = [
#     "Name", "Identifier", "FirstName", "MiddleName", "LastName",
#     "PlanCost", "EmploymentStatus", "HireDate", "HiredOn",
#     "TerminationDate", "TerminatedOn", "StartDate",
#     "EnrolledOn", "EndDate", "EndedOn"
# ]

# def xml_to_exact_excel(xml_file):
#     # Parse XML
#     context = ET.iterparse(xml_file, events=("start", "end"))
#     data_list = []
#     header_info = {}

#     for event, elem in context:
#         if event == "end":
#             if elem.tag == "Header":
#                 header_info = {
#                     "Disclaimer": elem.findtext("Disclaimer", ""),
#                     "ExchangeName": elem.findtext("ExchangeName", ""),
#                     "VendorName": elem.findtext("VendorName", ""),
#                     "RunDate": elem.findtext("RunDate", ""),
#                 }

#             elif elem.tag == "Company":
#                 company_data = {
#                     "Identifier": elem.findtext("Identifier", ""),
#                     "Name": elem.findtext("Name", ""),
#                 }
#                 full_data = {**header_info, **company_data}
#                 employees = elem.findall("Employees/Employee")
#                 first_employee = True
                
#                 for emp in employees:
#                     emp_data = {
#                         "FirstName": emp.findtext("FirstName", ""),
#                         "MiddleName": emp.findtext("MiddleName", ""),
#                         "LastName": emp.findtext("LastName", ""),
#                         "EmploymentStatus": emp.findtext("EmploymentStatus", ""),
#                         "HireDate": emp.findtext("HireDate", ""),
#                         "HiredOn": emp.findtext("HiredOn", ""),
#                         "TerminationDate": emp.findtext("TerminationDate", ""),
#                         "TerminatedOn": emp.findtext("TerminatedOn", ""),
#                     }
                    
#                     enrollments = emp.findall("Enrollments/Enrollment")
#                     if enrollments:
#                         for enroll in enrollments:
#                             enroll_data = {
#                                 "PlanCost": enroll.findtext("PlanCost", ""),
#                                 "StartDate": enroll.findtext("StartDate", ""),
#                                 "EnrolledOn": enroll.findtext("EnrolledOn", ""),
#                                 "EndDate": enroll.findtext("EndDate", ""),
#                                 "EndedOn": enroll.findtext("EndedOn", ""),
#                             }
                            
#                             if first_employee:
#                                 final_data = {**full_data, **emp_data, **enroll_data}
#                                 first_employee = False
#                             else:
#                                 final_data = {**{k: "" for k in full_data}, **emp_data, **enroll_data}
                            
#                             cleaned_data = {col: final_data.get(col, "") for col in EXPECTED_COLUMNS}
#                             data_list.append(cleaned_data)
#                     else:
#                         if first_employee:
#                             final_data = {**full_data, **emp_data}
#                             first_employee = False
#                         else:
#                             final_data = {**{k: "" for k in full_data}, **emp_data}
                        
#                         cleaned_data = {col: final_data.get(col, "") for col in EXPECTED_COLUMNS}
#                         data_list.append(cleaned_data)
#                 elem.clear()
    
#     df = pd.DataFrame(data_list, columns=EXPECTED_COLUMNS)
#     return df

# st.title("XML to Excel Converter")

# uploaded_xml = st.file_uploader("Upload XML File", type=["xml"])

# if uploaded_xml:
#     with tempfile.NamedTemporaryFile(delete=False) as tmp_xml:
#         tmp_xml.write(uploaded_xml.read())
#         tmp_xml.close()
        
#         df = xml_to_exact_excel(tmp_xml.name)
#         st.write("Preview of Extracted Data:")
#         st.dataframe(df.head())
        
#         with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
#             df.to_excel(tmp_out.name, index=False)
#             st.download_button(
#                 label="XML Converted! Download Excel",
#                 data=open(tmp_out.name, "rb").read(),
#                 file_name="converted.xlsx",
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#             )
import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import tempfile
import os
import zipfile
from datetime import datetime

EXPECTED_COLUMNS = [
    "Name", "Identifier", "FirstName", "MiddleName", "LastName",
    "PlanCost", "EmploymentStatus", "HireDate", "HiredOn",
    "TerminationDate", "TerminatedOn", "StartDate",
    "EnrolledOn", "EndDate", "EndedOn"
]

def xml_to_exact_excel(xml_file):
    context = ET.iterparse(xml_file, events=("start", "end"))
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
    return df

def split_excel_by_name(df, month, year):
    temp_dir = tempfile.mkdtemp()
    file_paths = []
    
    for name, group in df.groupby('Name'):
        if pd.isna(name):
            continue
        safe_name = str(name).replace("/", "_").replace("\\", "_")
        filename = f"{safe_name} ({month}-{year}).xlsx"
        filepath = os.path.join(temp_dir, filename)
        group.to_excel(filepath, index=False)
        file_paths.append(filepath)
    
    zip_path = os.path.join(temp_dir, f"Split_Companies_{month}_{year}.zip")
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file in file_paths:
            zipf.write(file, os.path.basename(file))
    
    return zip_path

# Streamlit UI
st.title("📤 XML to Excel Converter with Optional Split")

uploaded_xml = st.file_uploader("Upload XML File", type=["xml"])

if uploaded_xml:
    with tempfile.NamedTemporaryFile(delete=False) as tmp_xml:
        tmp_xml.write(uploaded_xml.read())
        tmp_xml.close()

        df = xml_to_exact_excel(tmp_xml.name)
        st.success("XML converted to Excel!")
        st.write("Preview of Extracted Data:")
        st.dataframe(df.head())

        # Offer main Excel download
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
            df.to_excel(tmp_out.name, index=False)
            st.download_button(
                label="📥 Download Full Excel",
                data=open(tmp_out.name, "rb").read(),
                file_name="converted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # Optionally split Excel by company name
        if st.checkbox("Do you want to split the Excel by company name ('Name')?"):
            today = datetime.today()
            default_month = today.strftime("%b")
            default_year = today.year

            col1, col2 = st.columns(2)
            with col1:
                month = st.text_input("Enter Month", default_month)
            with col2:
                year = st.text_input("Enter Year", str(default_year))

            if st.button("🔀 Split and Download Zip"):
                zip_file = split_excel_by_name(df, month, year)
                with open(zip_file, "rb") as zf:
                    st.download_button(
                        label="📦 Download Split Files (ZIP)",
                        data=zf.read(),
                        file_name=f"Split_Companies_{month}_{year}.zip",
                        mime="application/zip"
                    )

