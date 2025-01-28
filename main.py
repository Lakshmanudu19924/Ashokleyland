import streamlit as st
import pyrebase
import firebase_admin
from firebase_admin import credentials, auth
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


firebaseConfig = {
  "apiKey": "AIzaSyAcQpRtTUZ_LH9MqKuV0IJ51zqD_tAU410",
  "authDomain": "ashok-88fd9.firebaseapp.com",
  "projectId": "ashok-88fd9",
  "storageBucket": "ashok-88fd9.firebasestorage.app",
  "messagingSenderId": "249675153769",
  "appId": "1:249675153769:web:c59ed35637fcc4b9dc7df8",
  "measurementId": "G-PQFTZ53R4M",
  "databaseURL": "https://ashok-88fd9-default-rtdb.firebaseio.com"
}

firebase = pyrebase.initialize_app(firebaseConfig)
auth_client = firebase.auth()

if not firebase_admin._apps:
    cred = credentials.Certificate("ashok-88fd9-firebase-adminsdk-fbsvc-a87adc24e7.json")  # Replace with your Firebase Admin SDK JSON file
    firebase_admin.initialize_app(cred)

def login():
    st.title("Login")

    email = st.text_input("Email")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        try:
            user = auth_client.sign_in_with_email_and_password(email, password)
            st.success("Successfully logged in!")
            st.session_state["user"] = user
        except Exception as e:
            st.error(f"Error: {e}")


def register():
    st.title("Register")

    email = st.text_input("Email")
    password = st.text_input("Password", type="password")
    confirm_password = st.text_input("Confirm Password", type="password")

    if st.button("Register"):
        # Check if the email domain matches
        if not email.endswith("@ashokleyland.com"):
            st.error("Registration is restricted to @ashokleyland.com email addresses.")
        elif password != confirm_password:
            st.error("Passwords do not match!")
        else:
            try:
                auth_client.create_user_with_email_and_password(email, password)
                st.success("Successfully registered! Please log in.")
            except Exception as e:
                st.error(f"Error: {e}")


def logout():
    if "user" in st.session_state:
        del st.session_state["user"]
        st.success("Logged out successfully!")
        
def app_functionality():
    if "user" not in st.session_state:
        st.warning("Please log in to access the app.")
        return

    st.sidebar.title("Choose a Functionality")
    options = {
        "Made here parts calculation": process_part_matrix_master,
        "Priority Sheet": Priority_Analysis_P_NO_with_WIP_Description_and_SUB1_Mapping,
        "Mapped set avilable with out consider alternates": map_wout_alt,
        "Mapped set avilable with consider alternates": map_w_alt,
        "Month GB Requirement After OS": Month,
        "GB Requirement For Bal Month": Gbreq
    }

    choice = st.sidebar.radio("Select a process", list(options.keys()))
    options[choice]()


def Gbreq():
    st.title("GB Requirement For Bal Month")

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

    if uploaded_file:
        try:
            workbook = pd.ExcelFile(uploaded_file)

            if "Monthly Opening Stock" not in workbook.sheet_names or "3 Month Plan" not in workbook.sheet_names:
                st.error("Ensure the Excel file has sheets named 'Monthly Opening Stock' and '3 Month Plan'.")
            else:
                os_sheet = pd.read_excel(workbook, sheet_name="Monthly Opening Stock")
                plan_sheet = pd.read_excel(workbook, sheet_name="3 Month Plan")

                # Extract months from headers
                month_headers = list(
                    set(
                        [
                            col.split(" W")[0]
                            for col in plan_sheet.columns
                            if "W" in col
                        ]
                    )
                )

                selected_month = st.selectbox("Select Month", month_headers)

                if selected_month:
                    st.subheader(f"Results for {selected_month}")

                    results = []

                    for i, row in os_sheet.iterrows():
                        gb_value = row.get("GB", 0)
                        opening_stock = row.get("Opening Stock", 0)

                        w1_plan = plan_sheet.get(f"{selected_month} W1", pd.Series([0])).iloc[i] if i < len(plan_sheet) else 0
                        w2_plan = plan_sheet.get(f"{selected_month} W2", pd.Series([0])).iloc[i] if i < len(plan_sheet) else 0
                        w3_plan = plan_sheet.get(f"{selected_month} W3", pd.Series([0])).iloc[i] if i < len(plan_sheet) else 0
                        w4_plan = plan_sheet.get(f"{selected_month} W4", pd.Series([0])).iloc[i] if i < len(plan_sheet) else 0

                        w1_rev = max(0, w1_plan - opening_stock)
                        w1_excess = w1_rev - w1_plan

                        w2_rev_plan = w2_plan + w1_excess
                        w2_rev_plan_wn = max(0, w2_rev_plan)
                        week2_excess = w2_rev_plan - w2_plan

                        w3_rev_plan = w3_plan + week2_excess
                        w3_rev_plan_wn = max(0, w3_rev_plan)
                        week3_excess = w3_rev_plan - w3_plan

                        w4_rev_plan = w4_plan + week3_excess
                        w4_rev_plan_wn = max(0, w4_rev_plan)

                        results.append(
                            {
                                "GB": gb_value,
                                "Total": w1_plan + w2_plan + w3_plan + w4_plan,
                                "W1": w1_plan,
                                "W2": w2_plan,
                                "W3": w3_plan,
                                "W4": w4_plan,
                                "W1 Rev": w1_rev,
                                "W1 Excess": w1_excess,
                                "W2 Rev Plan": w2_rev_plan,
                                "W2 Rev Plan w/o Negative": w2_rev_plan_wn,
                                "Week 2 Excess / Less": week2_excess,
                                "W3 Rev Plan": w3_rev_plan,
                                "W3 Rev Plan w/o Negative": w3_rev_plan_wn,
                                "Week 3 Excess / Less": week3_excess,
                                "W4 Rev Plan": w4_rev_plan,
                                "W4 Rev Plan w/o Negative": w4_rev_plan_wn,
                            }
                        )

                    results_df_gb = pd.DataFrame(results)

                    st.dataframe(results_df_gb, use_container_width=True)

                    output = io.BytesIO()
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Processed Data"

                        # Write DataFrame to the worksheet
                    for row in dataframe_to_rows(results_df_gb, index=False, header=True):
                        ws.append(row)

                        # Save the workbook to the BytesIO object
                    wb.save(output)
                    processed_file = output.getvalue()

                    st.download_button(
                        label="Download Processed Excel",
                        data=processed_file,
                        file_name="GB Requirement For Bal Month.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"Error processing file: {e}")

def Month():

    st.title("Monthly GB Requirement After OS")

# File upload
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

    if uploaded_file:
        try:
            workbook = pd.ExcelFile(uploaded_file)

            # Check for required sheets
            if "3 Month Plan" not in workbook.sheet_names or "Monthly Opening Stock" not in workbook.sheet_names:
                st.error("Ensure the Excel file has sheets named 'Monthly Opening Stock' and '3 Month Plan'.")
            else:
                # Load sheets
                plan_df = pd.read_excel(workbook, sheet_name="3 Month Plan")
                os_df = pd.read_excel(workbook, sheet_name="Monthly Opening Stock")

                # Extract unique months
                month_headers = list(
                    {
                        header.split(" W")[0].strip()
                        for header in plan_df.columns
                        if "W" in header
                    }
                )

                selected_month = st.selectbox("Select Month", month_headers)

                if selected_month:
                    st.subheader(f"Results for {selected_month}")
                    processed_data = []

                    for _, row in os_df.iterrows():
                        gb_value = row.get("GB", 0)
                        opening_stock = row.get("Opening Stock", 0)
                        remaining_stock = opening_stock

                        row_result = {"GB": gb_value, "Opening Stock": opening_stock}

                        for week in ["W1", "W2", "W3", "W4"]:
                            header = f"{selected_month} {week}"
                            if header in plan_df.columns:
                                week_plan = plan_df.loc[_, header] if _ < len(plan_df) else 0
                                fulfilled_plan = min(week_plan, remaining_stock)
                                unmet_plan = week_plan - fulfilled_plan
                                remaining_stock -= fulfilled_plan

                                row_result[header] = week_plan
                                row_result[f"Plan for {week}"] = unmet_plan

                        processed_data.append(row_result)

                    # Rearrange columns
                    column_order = [
                        "GB", "Opening Stock",
                        f"{selected_month} W1", f"{selected_month} W2", f"{selected_month} W3", f"{selected_month} W4",
                        "Plan for W1", "Plan for W2", "Plan for W3", "Plan for W4"
                    ]
                    results_df = pd.DataFrame(processed_data)[column_order]

                    st.write("### Calculated Results")
                    st.dataframe(results_df, use_container_width=True)

                    # Download button
                    output = io.BytesIO()
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Processed Data"

                        # Write DataFrame to the worksheet
                    for row in dataframe_to_rows(results_df, index=False, header=True):
                        ws.append(row)

                        # Save the workbook to the BytesIO object
                    wb.save(output)
                    processed_file = output.getvalue()

                    st.download_button(
                        label="Download Processed Excel",
                        data=processed_file,
                        file_name="Month GB Requirement After OS.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"An error occurred: {e}")
                
def map_wout_alt():
    # Title for the Streamlit app
    st.title("Mapped set avilable with out consider alternates ")

    # File uploader for user to upload an Excel file
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"],key="mapset")

    if uploaded_file:
        try:
            # Load the Excel file
            data = pd.ExcelFile(uploaded_file)

            # Check for required sheets
            required_sheets = [
                "Today's Tentative Plan", "Part Raw Data", "Nomenclature Master", "Made Here Parts Calc"
            ]

            # Validate sheet names
            if not all(sheet in data.sheet_names for sheet in required_sheets):
                missing_sheets = [sheet for sheet in required_sheets if sheet not in data.sheet_names]
                st.error(f"Missing sheets: {', '.join(missing_sheets)}")
            else:
                # Load sheets
                tentative_plan_df = data.parse("Today's Tentative Plan")
                nomenclature_master_df = data.parse("Nomenclature Master")
                part_raw_data_df = data.parse("Part Raw Data")
                made_here_parts_calc_df = data.parse("Made Here Parts Calc")

                # Standardize column names
                tentative_plan_df.columns = tentative_plan_df.columns.str.strip().str.upper()
                nomenclature_master_df.columns = nomenclature_master_df.columns.str.strip().str.upper()
                part_raw_data_df.columns = part_raw_data_df.columns.str.strip().str.upper()
                made_here_parts_calc_df.columns = made_here_parts_calc_df.columns.str.strip().str.upper()

                # Ensure required columns exist
                required_columns = {
                    "Today's Tentative Plan": ["MODEL", "QTY"],
                    "Nomenclature Master": ["MODEL", "SPE"],
                    "Part Raw Data": [
                        "SPE", "1ST ON MS", "2ND ON MS", "3RD ON MS", "4TH ON MS", "5TH ON MS", "REV ON MS", "CM ON LS",
                        "REV IDLER", "3RD ON LS", "4TH ON LS", "5TH ON LS", "INPUT SHAFT", "MAIN SHAFT", "LAY SHAFT",
                        "HUB 1/ 2", "HUB 3/4", "HUB 5/6", "FDR", "SLEEVE 1/ 2", "SLEEVE 3/4", "SLEEVE 5/6",
                        "CONE 1/2", "CONE 3/4", "CONE 5/6", "CONE 3", "CONE 4"
                    ],
                    "Made Here Parts Calc": ["P.NO", "CURRENT MH"]
                }

                missing_columns = []
                for sheet_name, cols in required_columns.items():
                    df = {
                        "Today's Tentative Plan": tentative_plan_df,
                        "Nomenclature Master": nomenclature_master_df,
                        "Part Raw Data": part_raw_data_df,
                        "Made Here Parts Calc": made_here_parts_calc_df
                    }[sheet_name]

                    for col in cols:
                        if col not in df.columns:
                            missing_columns.append(f"{col} in {sheet_name}")

                if missing_columns:
                    st.error(f"Missing columns: {', '.join(missing_columns)}")
                else:
                    # Step 1: Map MODEL to SPE
                    tentative_plan_df = tentative_plan_df.merge(
                        nomenclature_master_df[["MODEL", "SPE"]], on="MODEL", how="left"
                    )

                    # Step 2: Map SPE to columns
                    columns_to_process = [
                        "1ST ON MS", "2ND ON MS", "3RD ON MS", "4TH ON MS", "5TH ON MS", "REV ON MS", "CM ON LS",
                        "REV IDLER", "3RD ON LS", "4TH ON LS", "5TH ON LS", "INPUT SHAFT", "MAIN SHAFT", "LAY SHAFT",
                        "HUB 1/ 2", "HUB 3/4", "HUB 5/6", "FDR", "SLEEVE 1/ 2", "SLEEVE 3/4", "SLEEVE 5/6",
                        "CONE 1/2", "CONE 3/4", "CONE 5/6", "CONE 3", "CONE 4"
                    ]

                    tentative_plan_df = tentative_plan_df.merge(
                        part_raw_data_df[["SPE"] + columns_to_process], on="SPE", how="left"
                    )

                    # Replace None or NaN values with 0
                    tentative_plan_df.fillna(0, inplace=True)

                    # Step 3: Calculate CURRENT MH and REMAINING MH row-wise for each column
                    remaining_stock = made_here_parts_calc_df.set_index("P.NO")["CURRENT MH"].to_dict()

                    def calculate_row_remaining(row, column):
                        key = row[column]
                        if key in remaining_stock:
                            available_mh = remaining_stock[key]
                            used_mh = min(row["QTY"], available_mh)
                            remaining_stock[key] -= used_mh
                            return used_mh
                        return 0

                    for col in columns_to_process:
                        tentative_plan_df[f"CURRENT MH ({col})"] = tentative_plan_df.apply(
                            lambda row: calculate_row_remaining(row, col), axis=1
                        )

                        tentative_plan_df[f"REMAINING MH ({col})"] = (
                            tentative_plan_df["QTY"] - tentative_plan_df[f"CURRENT MH ({col})"].fillna(0)
                        ).clip(lower=0)

                    # Step 4: Calculate the minimum CURRENT MH for each row, excluding zero-value columns
                    def calculate_min_current_mh(row):
                        non_zero_values = [
                            row[f"CURRENT MH ({col})"] for col in columns_to_process
                            if row[col] != 0
                        ]
                        return min(non_zero_values) if non_zero_values else 0

                    tentative_plan_df["MINIMUM CURRENT MH"] = tentative_plan_df.apply(
                        calculate_min_current_mh, axis=1
                    )

                    # Step 5: Select final columns for output
                    final_columns = [
                        "MODEL", "SPE", "QTY", "MINIMUM CURRENT MH"
                    ]

                    for col in columns_to_process:
                        final_columns.extend([
                            col, f"CURRENT MH ({col})", f"REMAINING MH ({col})"
                        ])

                    final_df = tentative_plan_df[final_columns]

                    # Add serial numbers starting from 1
                    final_df.reset_index(inplace=True, drop=True)
                    final_df.index = final_df.index + 1
                    final_df.index.name = "Serial Number"

                    # Step 6: Add a Total Row for MINIMUM CURRENT MH only
                    total_row = {col: "" for col in final_df.columns}  # Initialize with empty strings
                    total_row["MODEL"] = "TOTAL"  # Add label in the MODEL column
                    total_row["MINIMUM CURRENT MH"] = final_df["MINIMUM CURRENT MH"].sum()  # Sum for MINIMUM CURRENT MH

                    # Append the total row to the DataFrame
                    final_df_wo = pd.concat([final_df, pd.DataFrame([total_row])], ignore_index=True)

                    # Display the final DataFrame
                    st.write("### Processed Data (Detailed):")
                    st.dataframe(final_df_wo)
                    # Option to download the processed file using openpyxl
                    output = io.BytesIO()
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Processed Data"

                    # Write DataFrame to the worksheet
                    for row in dataframe_to_rows(final_df_wo, index=False, header=True):
                        ws.append(row)

                    # Save the workbook to the BytesIO object
                    wb.save(output)
                    processed_file = output.getvalue()

                    st.download_button(
                        label="Download Processed Excel",
                        data=processed_file,
                        file_name="Without Alternate.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.info("Please upload an Excel file to get started.")
        
    
def Priority_Analysis_P_NO_with_WIP_Description_and_SUB1_Mapping():
    
    st.title("Priority Sheet")

# File uploader
    data_file = st.file_uploader("Upload the Excel file", type=["xlsx"])

    if data_file is not None:
        try:
            # Read the uploaded Excel file
            excel_data = pd.ExcelFile(data_file)

            # Check if 'Priority format' sheet exists
            if 'Priority format' in excel_data.sheet_names:
                priority_df = excel_data.parse('Priority format')

                # Check if 'P.NO' column exists
                if 'P.NO' in priority_df.columns:
                    part_no_column = priority_df[['P.NO']].drop_duplicates().reset_index(drop=True)
                else:
                    st.error("The 'P.NO' column was not found in the 'Priority format' sheet.")
                    part_no_column = None
            else:
                st.error("The 'Priority format' sheet was not found in the uploaded Excel file.")
                part_no_column = None

            # Check if 'Made Here Parts Calc' sheet exists
            if 'Made Here Parts Calc' in excel_data.sheet_names:
                made_here_df = excel_data.parse('Made Here Parts Calc')

                # Check if required columns exist
                required_columns = ['P.NO', 'Hard WIP', 'HT WIP', 'Soft WIP', 'Rough WIP', 'Desc']
                missing_columns = [col for col in required_columns if col not in made_here_df.columns]

                if not missing_columns:
                    # Use the columns required for mapping
                    wip_data = made_here_df[required_columns]
                    # Replace None values with 0
                    wip_data.fillna(0, inplace=True)
                else:
                    st.error(f"The following columns were not found in the 'Made Here Parts Calc' sheet: {', '.join(missing_columns)}")
                    wip_data = None
            else:
                st.error("The 'Made Here Parts Calc' sheet was not found in the uploaded Excel file.")
                wip_data = None

            # Check if 'Alternate Part Master' sheet exists
            if 'Alternate Part Master' in excel_data.sheet_names:
                alternate_part_master_df = excel_data.parse('Alternate Part Master')

                # Check if 'P.NO' and 'SUB1' columns exist
                required_sub1_columns = ['P.NO', 'SUB1']
                missing_sub1_columns = [col for col in required_sub1_columns if col not in alternate_part_master_df.columns]

                if not missing_sub1_columns:
                    sub1_data = alternate_part_master_df[['P.NO', 'SUB1']].drop_duplicates().reset_index(drop=True)
                else:
                    st.error(f"The following columns were not found in the 'Alternate Part Master' sheet: {', '.join(missing_sub1_columns)}")
                    sub1_data = None
            else:
                st.error("The 'Alternate Part Master' sheet was not found in the uploaded Excel file.")
                sub1_data = None

            # Check if 'Without_Alternate' sheet exists
            if 'Without_Alternate' in excel_data.sheet_names:
                without_alternate_df = excel_data.parse('Without_Alternate')

                # Check if required columns exist
                remaining_mh_columns = [
                    '1ST ON MS', '2ND ON MS', '3RD ON MS', '4TH ON MS', '5TH ON MS', 'REV ON MS', 'CM ON LS', 'REV IDLER',
                    '3RD ON LS', '4TH ON LS', '5TH ON LS', 'INPUT SHAFT', 'MAIN SHAFT', 'LAY SHAFT', 'HUB 1/ 2', 'HUB 3/4',
                    'HUB 5/6', 'FDR', 'SLEEVE 1/ 2', 'SLEEVE 3/4', 'SLEEVE 5/6', 'CONE 1/2', 'CONE 3/4', 'CONE 5/6',
                    'CONE 3', 'CONE 4'
                ]

                stacked_data = pd.DataFrame(columns=['P.NO', 'REMAINING MH'])

                for column in remaining_mh_columns:
                    remaining_column = f"REMAINING MH ({column})"
                    if column in without_alternate_df.columns and remaining_column in without_alternate_df.columns:
                        # Extract relevant data
                        ms_data = without_alternate_df[[column, remaining_column]]

                        # Drop rows with missing values
                        ms_data.dropna(subset=[column, remaining_column], inplace=True)

                        # Rename columns to standardize
                        ms_data.columns = ['P.NO', 'REMAINING MH']

                        # Append to the stacked data
                        stacked_data = pd.concat([stacked_data, ms_data], ignore_index=True)
                    else:
                        st.warning(f"Columns '{column}' or '{remaining_column}' are missing in the 'Without_Alternate' sheet.")

                # Aggregate the stacked data by summing up REMAINING MH for each P.NO
                aggregated_stacked_data = stacked_data.groupby('P.NO', as_index=False).agg({
                    'REMAINING MH': 'sum'
                })

                # Display the aggregated stacked data
                # st.subheader("Aggregated Stacked Data: P.NO with Total Remaining MH")
                # st.write(aggregated_stacked_data)

            # Map the data
            if part_no_column is not None and wip_data is not None and sub1_data is not None:
                # Replace None values with 0 in the part_no_column
                part_no_column.fillna(0, inplace=True)

                # Merge part.no with corresponding WIP columns and Desc column
                mapped_data = part_no_column.merge(wip_data, on='P.NO', how='left')

                # Merge with SUB1 column
                mapped_data = mapped_data.merge(sub1_data, on='P.NO', how='left')

                # Replace None values with 0 in the mapped data
                mapped_data.fillna(0, inplace=True)

                # Use SUB1 values to extract corresponding WIP data from 'Made Here Parts Calc'
                sub1_wip_data = made_here_df[['P.NO', 'Hard WIP', 'HT WIP', 'Soft WIP', 'Rough WIP']]
                sub1_wip_data.columns = ['SUB1', 'Hard WIP (2)', 'HT WIP (2)', 'Soft WIP (2)', 'Rough WIP (2)']

                # Merge SUB1 WIP data into the mapped data
                mapped_data = mapped_data.merge(sub1_wip_data, on='SUB1', how='left')

                # Replace None values with 0 in the final data
                mapped_data.fillna(0, inplace=True)

                # Update Remaining MH based on aggregated stacked data
                def update_remaining_mh(row):
                    matched_row = aggregated_stacked_data.loc[aggregated_stacked_data['P.NO'] == row['P.NO']]
                    if not matched_row.empty:
                        return matched_row.iloc[0]['REMAINING MH']
                    return row['Remaining MH']

                mapped_data['Remaining MH'] = mapped_data.apply(update_remaining_mh, axis=1)

                # Final column order with renamed appended columns to avoid duplicates
                final_columns = ['P.NO', 'Remaining MH', 'Hard WIP', 'HT WIP', 'Soft WIP', 'Rough WIP', 'Desc', 'SUB1', 
                                'Hard WIP (2)', 'HT WIP (2)', 'Soft WIP (2)', 'Rough WIP (2)', 
                                'P.NO (Appended)', 'Hard WIP (Appended)']

                # Renaming the appended columns
                mapped_data['P.NO (Appended)'] = mapped_data['P.NO']
                mapped_data['Hard WIP (Appended)'] = mapped_data['Hard WIP']

                # Selecting the updated columns for final output
                mapped_data = mapped_data[final_columns]

                # Display the mapped data
                st.subheader("Mapped Data: P.NO with WIP, Description, and SUB1 Columns")
                st.write(mapped_data)
            elif part_no_column is not None:
                # Replace None values with 0 in the part_no_column
                part_no_column.fillna(0, inplace=True)

                st.subheader("Extracted 'P.NO' Column from Priority format Sheet (Duplicates Removed)")
                st.write(part_no_column)
            elif wip_data is not None:
                st.subheader("Extracted WIP and Desc Columns from Made Here Parts Calc Sheet (Duplicates Removed)")
                st.write(wip_data)
            elif sub1_data is not None:
                st.subheader("Extracted SUB1 Column from Alternate Part Master Sheet (Duplicates Removed)")
                st.write(sub1_data)

        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "Processed Data"

                    # Write DataFrame to the worksheet
        for row in dataframe_to_rows(mapped_data, index=False, header=True):
            ws.append(row)

                    # Save the workbook to the BytesIO object
            wb.save(output)
            processed_file = output.getvalue()

        st.download_button(
            label="Download Processed Excel",
            data=processed_file,
            file_name="Priority Sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )    
    else:
        st.info("Please upload an Excel file to proceed.")

    
def process_part_matrix_master():
    st.title("Made here parts calculation")
    st.write("UEpload an Excel file, and we'll process the 'Part Matrix Master', 'GB Requirement for Bal Month', and 'Date wise made here' sheets for you.")

    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

    if uploaded_file:
        # try:
            # Load the Excel file
            excel_data = pd.ExcelFile(uploaded_file)

            # Check if necessary sheets exist
            required_sheets = ['Part Matrix Master', 'GB Requirement for Bal Month', 'Date wise made here']
            if all(sheet in excel_data.sheet_names for sheet in required_sheets):
                # Read the necessary sheets
                part_matrix_df = pd.read_excel(excel_data, sheet_name='Part Matrix Master')
                gb_requirement_df = pd.read_excel(excel_data, sheet_name='GB Requirement for Bal Month')
                date_wise_df = pd.read_excel(excel_data, sheet_name='Date wise made here')

                # Fill empty values with 0
                part_matrix_df.fillna(0, inplace=True)
                gb_requirement_df.fillna(0, inplace=True)
                date_wise_df.fillna(0, inplace=True)

                # Ensure the 'W1 Rev', 'W2 Rev', 'W3 Rev', and 'W4 Rev' columns exist
                required_columns = ['W1 Rev', 'W2 Rev', 'W3 Rev', 'W4 Rev']
                for col in required_columns:
                    if col not in gb_requirement_df.columns:
                        st.error(f"'{col}' column not found in 'GB Requirement for Bal Month' sheet.")
                        return

                # Create mappings from GB Requirement sheet
                gb_mappings = {
                    'W1': gb_requirement_df.set_index(gb_requirement_df.columns[0])['W1 Rev'].to_dict(),
                    'W2': gb_requirement_df.set_index(gb_requirement_df.columns[0])['W2 Rev'].to_dict(),
                    'W3': gb_requirement_df.set_index(gb_requirement_df.columns[0])['W3 Rev'].to_dict(),
                    'W4': gb_requirement_df.set_index(gb_requirement_df.columns[0])['W4 Rev'].to_dict(),
                }

                # Duplicate columns and compute for W1, W2, W3, and W4
                duplicate_columns = [
                    "A1577000", "A5P35700", "A1580300", "A5P41600", "A5P15900", "A5P43400", "A5P36500", "A5P46200",
                    "A5P64100", "A5P56000", "A5P25000", "A5P41800", "A5P27700", "A5P53200", "A5P50700", "A5P07100",
                    "A5P71000", "A5P71300", "A5P75900", "A5P73600", "A5P41400", "A5P58800", "A5P66600", "A5P50100",
                    "A5P47800", "A5P76200", "A1571600", "A5P56700", "A5P72100", "A5P74000", "A5P72300", "MM22000090",
                    "MM22000091", "MM22000111", "MM22000092", "MM22000114", "MM22000130", "MM22000136", "MM22000228",
                    "MM22000163", "MM22000165", "MM22000181", "MM22000170", "MM22000179", "MM22000150", "MM22000172",
                    "MM22000214", "MM22000239", "MM22000353", "MM22000216", "MM22000233", "MM22000235", "MM22000191",
                    "MM22000241", "MM22000253", "MM22000256", "MM22000259", "MM22000260", "MM22000261", "MM22000284",
                    "MM22000287", "MM22000288", "MM22000289", "MM22000290", "MM22000294", "MM22000307", "A5P45600",
                    "GB105/32", "A5P51000", "A5P36200", "A5P50200", "A5P72900"
                ]

                for week, mapping in gb_mappings.items():
                    for col in duplicate_columns:
                        if col in part_matrix_df.columns:
                            # Replace negative values in Part Matrix Master with 0
                            part_matrix_df[col] = part_matrix_df[col].apply(lambda x: max(x, 0))

                            # Create duplicate column and apply mapping
                            new_col_name = f"{col}_{week.lower()}"
                            part_matrix_df[new_col_name] = part_matrix_df[col] * mapping.get(col, 1)
                            part_matrix_df[new_col_name] = part_matrix_df[new_col_name].apply(lambda x: max(x, 0))

                # Calculate W1, W2, W3, and W4 as the sum of respective columns
                for week in gb_mappings.keys():
                    week_cols = [f"{col}_{week.lower()}" for col in duplicate_columns if f"{col}_{week.lower()}" in part_matrix_df.columns]
                    part_matrix_df[week] = part_matrix_df[week_cols].sum(axis=1)

                # Process unique dates from 'Date wise made here'
                if 'Date' in date_wise_df.columns:
                    unique_dates = date_wise_df['Date'].drop_duplicates().sort_values()

                    selected_date = st.selectbox("Select a Date", unique_dates)
                    st.write(f"You selected: {selected_date}")

                    # Filter rows based on the selected date
                    filtered_date_wise_df = date_wise_df[date_wise_df['Date'] == selected_date]

                    # Add the required columns to 'Part Matrix Master'
                    if {'Current MH', 'Hard WIP', 'HT WIP', 'Soft WIP', 'Rough WIP'}.issubset(filtered_date_wise_df.columns):
                        part_matrix_df['Current MH'] = filtered_date_wise_df['Current MH'].values
                        part_matrix_df['Hard WIP'] = filtered_date_wise_df['Hard WIP'].values
                        part_matrix_df['HT WIP'] = filtered_date_wise_df['HT WIP'].values
                        part_matrix_df['Soft WIP'] = filtered_date_wise_df['Soft WIP'].values
                        part_matrix_df['Rough WIP'] = filtered_date_wise_df['Rough WIP'].values

                        # Rename 'Current MH' to 'Store Finished'
                        part_matrix_df.rename(columns={'Current MH': 'Store Finished'}, inplace=True)
                    else:
                        st.warning("Some required columns are missing in 'Date wise made here'.")
                else:
                    st.error("'Date' column not found in 'Date wise made here' sheet.")

                # Display the processed DataFrame
                st.subheader("Processed 'Part Matrix Master' Sheet")
                st.dataframe(part_matrix_df)

                # Option to download the processed file using openpyxl
                output = io.BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Processed Data"

                # Write DataFrame to the worksheet
                for row in dataframe_to_rows(part_matrix_df, index=False, header=True):
                    ws.append(row)

                # Save the workbook to the BytesIO object
                wb.save(output)
                processed_file = output.getvalue()

                st.download_button(
                    label="Download Processed Excel",
                    data=processed_file,
                    file_name="Made Here Parts Calculation.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            else:
                st.error("The required sheets ('Part Matrix Master', 'GB Requirement for Bal Month', and 'Date wise made here') were not found in the uploaded file.")

        # except Exception as e:
        #     st.error(f"An error occurred: {e}")

def map_w_alt():
    st.title("Mapped set avilable with consider alternates")
    

# File uploader for user to upload an Excel file
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

    if uploaded_file:
        try:
            # Load the Excel file
            data = pd.ExcelFile(uploaded_file)

            # Check for required sheets
            required_sheets = [
                "Today's Tentative Plan", "Part Raw Data", "Nomenclature Master", "Made Here Parts Calc", "Alternate Part Master"
            ]

            # Validate sheet names
            if not all(sheet in data.sheet_names for sheet in required_sheets):
                missing_sheets = [sheet for sheet in required_sheets if sheet not in data.sheet_names]
                st.error(f"Missing sheets: {', '.join(missing_sheets)}")
            else:
                # Load sheets
                tentative_plan_df = data.parse("Today's Tentative Plan")
                nomenclature_master_df = data.parse("Nomenclature Master")
                part_raw_data_df = data.parse("Part Raw Data")
                made_here_parts_calc_df = data.parse("Made Here Parts Calc")
                alternate_part_master_df = data.parse("Alternate Part Master")

                # Standardize column names
                tentative_plan_df.columns = tentative_plan_df.columns.str.strip().str.upper()
                nomenclature_master_df.columns = nomenclature_master_df.columns.str.strip().str.upper()
                part_raw_data_df.columns = part_raw_data_df.columns.str.strip().str.upper()
                made_here_parts_calc_df.columns = made_here_parts_calc_df.columns.str.strip().str.upper()
                alternate_part_master_df.columns = alternate_part_master_df.columns.str.strip().str.upper()

                # Ensure required columns exist
                required_columns = {
                    "Today's Tentative Plan": ["MODEL", "QTY"],
                    "Nomenclature Master": ["MODEL", "SPE"],
                    "Part Raw Data": [
                        "SPE", "1ST ON MS", "2ND ON MS", "3RD ON MS", "4TH ON MS", "5TH ON MS", "REV ON MS", "CM ON LS",
                        "REV IDLER", "3RD ON LS", "4TH ON LS", "5TH ON LS", "INPUT SHAFT", "MAIN SHAFT", "LAY SHAFT",
                        "HUB 1/ 2", "HUB 3/4", "HUB 5/6", "FDR", "SLEEVE 1/ 2", "SLEEVE 3/4", "SLEEVE 5/6",
                        "CONE 1/2", "CONE 3/4", "CONE 5/6", "CONE 3", "CONE 4"
                    ],
                    "Made Here Parts Calc": ["P.NO", "CURRENT MH"],
                    "Alternate Part Master": ["P.NO", "SUB1", "SUB2"]
                }

                missing_columns = []
                for sheet_name, cols in required_columns.items():
                    df = {
                        "Today's Tentative Plan": tentative_plan_df,
                        "Nomenclature Master": nomenclature_master_df,
                        "Part Raw Data": part_raw_data_df,
                        "Made Here Parts Calc": made_here_parts_calc_df,
                        "Alternate Part Master": alternate_part_master_df
                    }[sheet_name]

                    for col in cols:
                        if col not in df.columns:
                            missing_columns.append(f"{col} in {sheet_name}")

                if missing_columns:
                    st.error(f"Missing columns: {', '.join(missing_columns)}")
                else:
                    # Step 1: Map MODEL to SPE
                    tentative_plan_df = tentative_plan_df.merge(
                        nomenclature_master_df[["MODEL", "SPE"]], on="MODEL", how="left"
                    )

                    # Step 2: Map SPE to columns
                    columns_to_process = [
                        "1ST ON MS", "2ND ON MS", "3RD ON MS", "4TH ON MS", "5TH ON MS", "REV ON MS", "CM ON LS",
                        "REV IDLER", "3RD ON LS", "4TH ON LS", "5TH ON LS", "INPUT SHAFT", "MAIN SHAFT", "LAY SHAFT",
                        "HUB 1/ 2", "HUB 3/4", "HUB 5/6", "FDR", "SLEEVE 1/ 2", "SLEEVE 3/4", "SLEEVE 5/6",
                        "CONE 1/2", "CONE 3/4", "CONE 5/6", "CONE 3", "CONE 4"
                    ]

                    tentative_plan_df = tentative_plan_df.merge(
                        part_raw_data_df[["SPE"] + columns_to_process], on="SPE", how="left"
                    )

                    # Replace None or NaN values with 0
                    tentative_plan_df.fillna(0, inplace=True)

                    # Step 3: Map P.NO to SUB1 and SUB2 for each column
                    alternate_part_dict = alternate_part_master_df.set_index("P.NO")[["SUB1", "SUB2"]].to_dict("index")

                    for col in columns_to_process:
                        tentative_plan_df[f"SUB1 ({col})"] = tentative_plan_df[col].map(
                            lambda x: alternate_part_dict[x]["SUB1"] if x in alternate_part_dict else 0
                        )
                        tentative_plan_df[f"SUB2 ({col})"] = tentative_plan_df[col].map(
                            lambda x: alternate_part_dict[x]["SUB2"] if x in alternate_part_dict else 0
                        )

                    # Step 4: Calculate CURRENT MH row-wise for each column
                    remaining_stock = made_here_parts_calc_df.set_index("P.NO")["CURRENT MH"].to_dict()

                    def calculate_row_current(row, column):
                        key = row[column]
                        if key in remaining_stock:
                            available_mh = remaining_stock[key]
                            used_mh = min(row["QTY"], available_mh)
                            remaining_stock[key] -= used_mh
                            return used_mh
                        return 0

                    for col in columns_to_process:
                        tentative_plan_df[f"CURRENT MH ({col})"] = tentative_plan_df.apply(
                            lambda row: calculate_row_current(row, col), axis=1
                        )

                    # Step 5: Calculate CURRENT MH for SUB1
                    for col in columns_to_process:
                        tentative_plan_df[f"CURRENT MH (SUB1 {col})"] = tentative_plan_df.apply(
                            lambda row: max(row["QTY"] - row[f"CURRENT MH ({col})"], 0), axis=1
                        )

                    # Step 6: Calculate CURRENT MH for SUB2
                    for col in columns_to_process:
                        tentative_plan_df[f"CURRENT MH (SUB2 {col})"] = tentative_plan_df.apply(
                            lambda row: max(row[f"CURRENT MH ({col})"] + row[f"CURRENT MH (SUB1 {col})"] - row["QTY"], 0), axis=1
                        )

                    # Step 7: Calculate the minimum CURRENT MH for each row, excluding zero-value columns
                    def calculate_min_current_mh(row):
                        non_zero_values = [
                            row[f"CURRENT MH ({col})"] for col in columns_to_process
                            if row[col] != 0
                        ]
                        return min(non_zero_values) if non_zero_values else 0

                    tentative_plan_df["MINIMUM CURRENT MH"] = tentative_plan_df.apply(
                        calculate_min_current_mh, axis=1
                    )

                    # Step 8: Select final columns for output
                    final_columns = [
                        "MODEL", "SPE", "QTY", "MINIMUM CURRENT MH"
                    ]

                    for col in columns_to_process:
                        final_columns.extend([
                            col, f"CURRENT MH ({col})", f"SUB1 ({col})", f"CURRENT MH (SUB1 {col})", f"SUB2 ({col})", f"CURRENT MH (SUB2 {col})"
                        ])

                    final_df = tentative_plan_df[final_columns]

                    # Replace None or NaN values in the final DataFrame with 0
                    final_df.fillna(0, inplace=True)

                    # Add serial numbers starting from 1
                    final_df.reset_index(inplace=True, drop=True)
                    final_df.index = final_df.index + 1
                    final_df.index.name = "Serial Number"

                    # Step 9: Add a Total Row for MINIMUM CURRENT MH only
                    total_row = {col: 0 for col in final_df.columns}  # Initialize with zeros
                    total_row["MODEL"] = "TOTAL"  # Add label in the MODEL column
                    total_row["MINIMUM CURRENT MH"] = final_df["MINIMUM CURRENT MH"].sum()  # Sum for MINIMUM CURRENT MH

                    # Append the total row to the DataFrame
                    final_df = pd.concat([final_df, pd.DataFrame([total_row])], ignore_index=True)

                    # Display the final DataFrame
                    st.write("### Processed Data (Detailed):")
                    st.dataframe(final_df)
                    output = io.BytesIO()
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Processed Data"

                    # Write DataFrame to the worksheet
                    for row in dataframe_to_rows(final_df, index=False, header=True):
                        ws.append(row)

                    # Save the workbook to the BytesIO object
                    wb.save(output)
                    processed_file = output.getvalue()

                    st.download_button(
                        label="Download Processed Excel",
                        data=processed_file,
                        file_name="With Alternative.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.info("Please upload an Excel file to get started.")
        


def main():
    st.sidebar.title("Navigation")
    menu = ["Login", "Register", "Logout", "App Functionality"]
    choice = st.sidebar.selectbox("Menu", menu)

    if choice == "Login":
        login()
    elif choice == "Register":
        register()
    elif choice == "Logout":
        logout()
    elif choice == "App Functionality":
        app_functionality()

# Define the other functions here (e.g., Gbreq, Month, etc.)

if __name__ == "__main__":
    main()