import json
import argparse
import datetime
from src import export_excel, modify_excel

# Transforms SQL table var names to Excel Header Names
headers_mapping = {
    "control_account_number": "CONTROL/ACCOUNT #",
    "last_name": "LAST NAME",
    "first_name": "FIRST NAME",
    "middle": "MIDDLE",
    "date_of_birth": "DATE OF BIRTH",
    "medicaid_id": "MEDICAID ID",
    "coverage_expiration_date": "COVERAGE EXPIRATION DATE",
    "date_of_service": "DATE OF SERVICE",
    "cpt_hcpcs_dental_code": "CPT/HCPCS/DENTAL CODE",
    "service_code_modifier": "SERVICE CODE MODIFIER",
    "billed_amount": "BILLED AMOUNT",
    "spend_down": "SPEND DOWN",
    "tpl_amount": "TPL AMOUNT",
    "tpl": "TPL"
}

# Used to convert date columns to have datetime objects recognized by Excel
date_columns = [
    "DATE OF BIRTH", "COVERAGE EXPIRATION DATE", "DATE OF SERVICE"
]

# Columns to insert and their insertion position
inserted_columns = {
    "GRAND TOTAL": 11, 
    "AMOUNT DUE": 12, 
    "LOCAL SHARE": 13, 
    "FEDERAL SHARE": 14, 
    "CONTRACTUAL ADJUSTMENT": 18, 
    "ADJUSTMENT REASON": 19, 
    "NOTE": 20
}

# Columns to unprotect
unprotected_columns = [
    "AMOUNT DUE", "SPEND DOWN", "CONTRACTUAL ADJUSTMENT", "ADJUSTMENT REASON", "NOTE"
                       ]

# Loading configuration file with formatting for each column
with open("config.json") as f:
    config_dict =  json.load(f)

def main(excel_file_name: str):

    """ 
    Creating Excel
    """

    # Read in dataframe and format data
    server_connection_string = "data/medical_data.db"
    raw_dataframe = export_excel.create_dataframe(server_connection_string)

    # Rename the headers of the dataframe
    renamed_headers = export_excel.transform_header(raw_dataframe, mapping_dict=headers_mapping)
    # Reformat the date columns
    reformatted_date = export_excel.format_date_columns(renamed_headers, date_columns)

    # Insert columns for user entry
    final_df = export_excel.insert_headers(reformatted_date, inserted_columns)

    # Get number of samples from query
    num_rows = final_df.shape[0]

    # Save dataframe to excel file
    # ex_name = 'python_CTS_example.xlsx'
    # path = export_excel.create_sheet(final_df, file_name=excel_file_name)

    export_excel.insert_into_template(final_df)

    # """
    # Modifying Excel
    # """

    # # Open up excel file to modify
    # ws_to_modify = modify_excel.CustomSpreadsheet(filepath=path, row_range=num_rows+50) # Dynamically define row_range to match query size + 50 extra

    # # Set the sheet to work on
    # ws_to_modify.set_sheet("Sheet1")

    # # Style the worksheet
    # ws_to_modify.set_header_style(color_code="4472c4", font_size=8, font_name="Quattrocento Sans")
    # ws_to_modify.set_header_height(header_row_height=22.9)
    # ws_to_modify.set_column_width(multiplier=1.1)
    # ws_to_modify.set_alternating_fill()
    # ws_to_modify.set_number_of_columns(num_columns=22)

    # # Add formatting and data validation to spreadsheet using the configuration dictionary
    # modify_excel.formatting_handler(ws_to_modify, config_dict)

    # # Set password protection for columns
    # password = "test"
    # modify_excel.protection_handler(ws_to_modify, cols_to_unprotect=unprotected_columns, password=password)
    
    # # Save the modified Excel file
    # ws_to_modify.workbook.save(path)
    # print(f"\nFormatted excel file saved to {path}")

if __name__ == "__main__":

    # Create defualt file name based on current datetime
    default_file_name = datetime.datetime.now().strftime("%m-%d-%Y_%H-%M-%S")

    # Define parser and arguments
    parser = argparse.ArgumentParser(prog="main.py")
    parser.add_argument("-n", "--name", type=str, help="Specify name of the excel file", required=False)

    args = parser.parse_args()

    # Use defined command line name if defined, else use default
    if args.name:
        file_name = args.name

        # Raise error if file name foes not have .xlsx handle
        if not file_name.endswith(".xlsx"):
            raise ValueError(f"Specified name: '{file_name}' is not formatted correctly. Filename should end with '.xlsx'")
    else:
        file_name = default_file_name + ".xlsx"

    main(file_name)
