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
        st.stop()  # Use st.stop() instead of return in Streamlit

    # Initialize session state for tracking the selected option
    if "selected_option" not in st.session_state:
        st.session_state.selected_option = None

    # Sidebar title
    st.sidebar.title("Choose a Functionality")

    # Define functionalities
    options = {
        "Made here parts calculation": process_part_matrix_master,
        "Priority Sheet": Priority_Analysis_P_NO_with_WIP_Description_and_SUB1_Mapping,
        "Month GB Requirement After OS": Month,
        "GB Requirement For Bal Month": Gbreq,
        "Mapped set available without considering alternates": map_wout_alt,
        "Mapped set available considering alternates": map_w_alt,
        "MPS Plan - 2 Weeks": None,  # Placeholder
        "MPS Plan - 4 Weeks": None   # Placeholder
    }

    # Function to update selection
    def update_selection(selection):
        st.session_state.selected_option = selection

    # Creating collapsible sections (dropdowns) with buttons
    with st.sidebar.expander("Matched Set"):
        if st.button("Made here parts calculation"):
            update_selection("Made here parts calculation")
        if st.button("Priority Sheet"):
            update_selection("Priority Sheet")
        if st.button("Month GB Requirement After OS"):
            update_selection("Month GB Requirement After OS")
        if st.button("GB Requirement For Bal Month"):
            update_selection("GB Requirement For Bal Month")

    with st.sidebar.expander("Against Tentative Plan"):
        if st.button("Matched set available without considering alternates"):
            update_selection("Mapped set available without considering alternates")
        if st.button("Matched set available considering alternates"):
            update_selection("Mapped set available considering alternates")

    with st.sidebar.expander("Against MPS -2 weeks"):
        if st.button("MPS Plan - 2 Weeks"):
            update_selection("MPS Plan - 2 Weeks")

    with st.sidebar.expander("Against MPS -4 weeks"):
        if st.button("MPS Plan - 4 Weeks"):
            update_selection("MPS Plan - 4 Weeks")

    # Render only the selected function
    if st.session_state.selected_option in options and options[st.session_state.selected_option]:
        options[st.session_state.selected_option]()



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
    st.title("Matched set avilable with out consider alternates ")

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
    
    st.title("Priority Analysis - P.NO with WIP, Description, and SUB1 Mapping")

    # File uploader
    data_file = st.file_uploader("Upload the Excel file", type=["xlsx"])

    if data_file is not None:
        try:
            # Read the uploaded Excel file
            excel_data = pd.ExcelFile(data_file)

            # Convert sheet names to uppercase for case-insensitive matching
            sheet_names_upper = {sheet_name.upper(): sheet_name for sheet_name in excel_data.sheet_names}

            # Check if 'Priority format' sheet exists
            if 'PRIORITY FORMAT' in sheet_names_upper:
                priority_df = excel_data.parse(sheet_names_upper['PRIORITY FORMAT'])
                priority_df.columns = priority_df.columns.str.strip().str.upper()

                if 'P.NO' in priority_df.columns:
                    part_no_column = priority_df[['P.NO']].drop_duplicates().reset_index(drop=True)
                    part_no_column.index += 1  # Set serial numbers starting from 1
                    part_no_column.index.name = "Serial Number"
                else:
                    st.error("The 'P.NO' column was not found in the 'Priority format' sheet.")
                    part_no_column = None
            else:
                st.error("The 'Priority format' sheet was not found in the uploaded Excel file.")
                part_no_column = None

            # Check if 'Made Here Parts Calc' sheet exists
            if 'MADE HERE PARTS CALC' in sheet_names_upper:
                made_here_df = excel_data.parse(sheet_names_upper['MADE HERE PARTS CALC'])
                made_here_df.columns = made_here_df.columns.str.strip().str.upper()

                required_columns = ['P.NO', 'HARD WIP', 'HT WIP', 'SOFT WIP', 'ROUGH WIP','WFT', 'DESC']
                missing_columns = [col for col in required_columns if col not in made_here_df.columns]

                if not missing_columns:
                    wip_data = made_here_df[required_columns].fillna(0)
                else:
                    st.error(f"Missing columns in 'Made Here Parts Calc': {', '.join(missing_columns)}")
                    wip_data = None
            else:
                st.error("The 'Made Here Parts Calc' sheet was not found.")
                wip_data = None

            # Check if 'Alternate Part Master' sheet exists
            if 'ALTERNATE PART MASTER' in sheet_names_upper:
                alternate_part_master_df = excel_data.parse(sheet_names_upper['ALTERNATE PART MASTER'])
                alternate_part_master_df.columns = alternate_part_master_df.columns.str.strip().str.upper()

                required_sub1_columns = ['P.NO', 'SUB1']
                missing_sub1_columns = [col for col in required_sub1_columns if col not in alternate_part_master_df.columns]

                if not missing_sub1_columns:
                    sub1_data = alternate_part_master_df[['P.NO', 'SUB1']].drop_duplicates().reset_index(drop=True)
                else:
                    st.error(f"Missing columns in 'Alternate Part Master': {', '.join(missing_sub1_columns)}")
                    sub1_data = None
            else:
                st.error("The 'Alternate Part Master' sheet was not found.")
                sub1_data = None

            # Mapping Data
            if part_no_column is not None and wip_data is not None and sub1_data is not None:
                mapped_data = part_no_column.merge(wip_data, on='P.NO', how='left')
                mapped_data = mapped_data.merge(sub1_data, on='P.NO', how='left')
                mapped_data.fillna(0, inplace=True)

                sub1_wip_data = made_here_df[['P.NO', 'HARD WIP', 'HT WIP', 'SOFT WIP', 'ROUGH WIP','WFT']]
                sub1_wip_data.columns = ['SUB1', 'HARD WIP (2)', 'HT WIP (2)', 'SOFT WIP (2)', 'ROUGH WIP (2)','WFT (2)']
                mapped_data = mapped_data.merge(sub1_wip_data, on='SUB1', how='left').fillna(0)

                # Load Cycle Time Sheet if exists
                if 'CYCLE TIME SHEET' in sheet_names_upper:
                    cycle_time_df = excel_data.parse(sheet_names_upper['CYCLE TIME SHEET'])
                    cycle_time_df.columns = cycle_time_df.columns.str.strip().str.upper()

                    if {'P.NO', 'CYCLE TIME'}.issubset(cycle_time_df.columns):
                        cycle_time_mapping = cycle_time_df.set_index('P.NO')['CYCLE TIME'].to_dict()
                    else:
                        cycle_time_mapping = {}
                else:
                    cycle_time_mapping = {}

                # Calculate 1st Priority
                def calculate_1st_priority(row):
                    cycle_time = cycle_time_mapping.get(row['SUB1'], None)
                    if cycle_time is not None:
                        if cycle_time < row['WFT'] or cycle_time == row['WFT']:
                            return "Hard-TG"
                    if row['WFT'] > 100:
                        return "Hard-TG"
                    return ""

                mapped_data['1st Priority'] = mapped_data.apply(calculate_1st_priority, axis=1)

                # Calculate 2nd Priority
                def calculate_2nd_priority(row):
                    if row['HT WIP'] > row['HARD WIP']:
                        return "Hard"
                    elif row['SOFT WIP'] + row['HT WIP'] > row['HARD WIP']:
                        return "HT"
                    elif row['SOFT WIP'] + row['HT WIP'] + row['ROUGH WIP'] + row['WFT'] > row['HARD WIP']:
                        return "Soft & HT"
                    elif row['SOFT WIP'] + row['HT WIP'] + row['ROUGH WIP'] + row['WFT'] < row['HARD WIP']:
                        return "Soft"
                    return "Rough"

                mapped_data['2nd Priority'] = mapped_data.apply(calculate_2nd_priority, axis=1)
                
                # Combine 1st & 2nd Priority
                mapped_data['1st&2nd Priority'] = mapped_data['1st Priority'] + " & " + mapped_data['2nd Priority']

                final_columns = ['P.NO', 'HARD WIP', 'HT WIP', 'SOFT WIP', 'ROUGH WIP','WFT', 'DESC', 'SUB1', 
                                'HARD WIP (2)', 'HT WIP (2)', 'SOFT WIP (2)', 'ROUGH WIP (2)','WFT (2)',
                                '1st Priority', '2nd Priority', '1st&2nd Priority']

                mapped_data = mapped_data[final_columns]
                mapped_data.reset_index(inplace=True, drop=True)
                mapped_data.index += 1
                mapped_data.index.name = "Serial Number"

                st.subheader("Mapped Data: P.NO with WIP, Description, and SUB1 Columns")
                st.write(mapped_data)

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
                    label="Download Priority sheet Excel",
                    data=processed_file,
                    file_name="Priority sheet.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    else:
        st.info("Please upload an Excel file to proceed.")
def process_part_matrix_master():
    st.title("made here part calculation")
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
                    if {'Current MH', 'Hard WIP', 'HT WIP', 'Soft WIP', 'Rough WIP','Hard Wating For teeth'}.issubset(filtered_date_wise_df.columns):
                        part_matrix_df['Current MH'] = filtered_date_wise_df['Current MH'].values
                        part_matrix_df['Hard WIP'] = filtered_date_wise_df['Hard WIP'].values
                        part_matrix_df['HT WIP'] = filtered_date_wise_df['HT WIP'].values
                        part_matrix_df['Soft WIP'] = filtered_date_wise_df['Soft WIP'].values
                        part_matrix_df['Rough WIP'] = filtered_date_wise_df['Rough WIP'].values
                        part_matrix_df['Hard Wating For teeth'] = filtered_date_wise_df['Hard Wating For teeth'].values

                        # Rename 'Current MH' to 'Store Finished'
                        part_matrix_df.rename(columns={'Hard Wating For teeth': 'WFT'}, inplace=True)
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
                    file_name="processed_part_matrix_master.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            else:
                st.error("The required sheets ('Part Matrix Master', 'GB Requirement for Bal Month', and 'Date wise made here') were not found in the uploaded file.")

        # except Exception as e:
        #     st.error(f"An error occurred: {e}")

def map_w_alt():
    st.title("Matched set avilable with consider alternates")
    

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
    st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: #ffffff;
    }
    
    [data-testid="stSidebar"] {
        background: linear-gradient(195deg, #1a1a2e 0%, #16213e 100%) !important;
        border-right: 1px solid #2a529850;
    }

    /* Text color adjustments */
    .stTextInput>label, .stNumberInput>label, .stSelectbox>label,
    .stRadio>label, .stMarkdown, .stTitle {
        color: #ffffff !important;
    }

    /* Input field hover and focus effects */
    .stTextInput input, .stNumberInput input, .stTextArea textarea {
        transition: all 0.3s ease !important;
        background-color: rgba(125, 214, 235, 0.57) !important;
        border: 1px solidrgba(200, 12, 217, 0.31) !important;
    }

    .stTextInput input:hover, .stNumberInput input:hover, .stTextArea textarea:hover {
        background-color: rgba(255, 255, 255, 0.94) !important;
        transform: scale(1.02);
    }

    .stTextInput input:focus, .stNumberInput input:focus, .stTextArea textarea:focus {
        color:black;
        
       
        transform: scale(1.02);
    }

    /* Button styling */
.stButton>button {
    background: linear-gradient(45deg, #4CAF50 0%, #45a049 100%);
    color: white !important;
    border: none;
    border-radius: 5px;
    padding: 10px 24px;
    transition: all 0.3s ease !important;
}

.stButton>button:hover {
    background: linear-gradient(45deg, #2196F3 0%, #1976D2 100%) !important;
    transform: scale(1.05);
    opacity: 0.9;
}

    /* Container styling */
    .stContainer {
        background-color: rgba(255, 255, 255, 0.1) !important;
        border-radius: 10px;
        padding: 20px;
        margin: 10px 0;
        backdrop-filter: blur(5px);
    }

    /* Dataframe styling */
    .dataframe {
        background-color: rgba(23, 137, 195, 0.1) !important;
    }

    /* Hover effects */
    .stButton>button:hover {
        transform: scale(1.05);
        opacity: 0.9;
    }
    </style>
    """,
    unsafe_allow_html=True
    )
    
    
    st.sidebar.title("Navigation")
    menu = ["Login", "Register", "Logout", "Dashboard"]
    choice = st.sidebar.selectbox("Menu", menu)

    if choice == "Login":
        login()
    elif choice == "Register":
        register()
    elif choice == "Logout":
        logout()
    elif choice == "Dashboard":
        app_functionality()

# Define the other functions here (e.g., Gbreq, Month, etc.)

if __name__ == "__main__":
    main()
