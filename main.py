from src import export_excel, modify_excel
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting import Rule

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
    "LOCAL_SHARE": 13, 
    "FEDERAL_SHARE": 14, 
    "CONTRACTUAL ADJUSTMENT": 18, 
    "ADJUSTMENT REASON": 19, 
    "NOTE": 20
}

if __name__ == "__main__":

    # Read in dataframe and format data
    server_connection_string = "data/medical_data.db"
    raw_dataframe = export_excel.create_dataframe(server_connection_string)

    # Rename the headers of the dataframe
    renamed_headers = export_excel.transform_header(raw_dataframe, mapping_dict=headers_mapping)
    print(renamed_headers.head(4))

    # Reformat the date columns
    reformatted_date = export_excel.format_date_columns(renamed_headers, date_columns)
    dates = reformatted_date[["DATE OF BIRTH", "COVERAGE EXPIRATION DATE", "DATE OF SERVICE"]]
    print(dates.head(3))

    # Insert columns for user entry
    final_df = export_excel.insert_headers(reformatted_date, inserted_columns)
    print(final_df.head(3))

    # Save dataframe to excel file
    path = export_excel.create_sheet(final_df)