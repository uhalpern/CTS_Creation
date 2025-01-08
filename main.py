import json
from src import export_excel, modify_excel

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

date_columns = [
    "DATE OF BIRTH", "COVERAGE EXPIRATION DATE", "DATE OF SERVICE"
]

inserted_columns = {
    "GRAND TOTAL": 11, 
    "AMOUNT DUE": 12, 
    "LOCAL SHARE": 13, 
    "FEDERAL SHARE": 14, 
    "CONTRACTUAL ADJUSTMENT": 18, 
    "ADJUSTMENT REASON": 19, 
    "NOTE": 20
}

with open("config.json") as f:
    config_dict =  json.load(f)

def main():

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

    # Save dataframe to excel file
    # TODO: Add functionality to ingest and use unique file name from other script
    ex_name = 'python_CTS_example.xlsx'
    path = export_excel.create_sheet(final_df, file_name=ex_name)

    """
    Modifying Excel
    """

    # Open up excel file to modify
    ws_to_modify = modify_excel.CustomSpreadsheet(filepath=path, row_range=5000)

    # Set the sheet to work on
    ws_to_modify.set_sheet("Sheet1")

    # Style the worksheet
    ws_to_modify.set_header_style(color_code="4472c4", font_size=8, font_name="Quattrocento Sans")
    ws_to_modify.set_header_height(header_row_height=22.9)
    ws_to_modify.set_column_width(multiplier=1.1)
    ws_to_modify.set_alternating_fill()
    ws_to_modify.set_number_of_columns(num_columns=22)

    # Add formatting and data validation to spreadsheet using the configuration dictionary
    modify_excel.formatting_handler(ws_to_modify, config_dict)

    # Save the modified Excel file
    ws_to_modify.workbook.save(path)
    print(f"\nFormatted excel file saved to {path}")

main()
