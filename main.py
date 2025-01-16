import json
import argparse
import datetime
from src import export_excel

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
    "AMOUNT DUE": 13, 
    "LOCAL SHARE": 12, 
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
with open("config.json", encoding='utf-8') as f:
    config_dict =  json.load(f)

def main(excel_file_name: str):

    """ 
    Exporting Excel
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

    # Ingest Data into spreadsheet
    workbook = export_excel.insert_into_template(final_df, validation_format_dict=config_dict)

    # Apply protection to sheet
    password = "test"
    export_excel.protection_handler(workbook, unprotected_columns, password, num_rows+50)

    # Save Workbook
    export_excel.save_workbook(workbook, excel_file_name)

if __name__ == "__main__":

    # Create default file name based on current datetime
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
