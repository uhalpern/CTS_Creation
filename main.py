from src import export_excel, modify_excel
import os


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

validation_dict = {
    "MEDICAID ID": {
        "val_formula": '=OR(AND(LEN(F1)=12,ISNUMBER(VALUE(LEFT(F1,2))),MID(F1,3,1)="-",ISNUMBER(VALUE(MID(F1,4,6))),MID(F1,10,1)="-",ISNUMBER(VALUE(RIGHT(F1,2)))),AND(LEN(F1)=10,NOT(VALUE(LEFT(F1))=0),ISNUMBER(VALUE(F1))))',
        "error_msg": (
            "Input must be in valid medicaid id format (**-******-**)"
        ),
    },
    "DATE OF BIRTH": {
        "val_formula": '=AND(ISNUMBER(E1), E1 > DATE(1900, 1, 1))',
        "error_msg": (
            "Invalid Date Format - Enter date as MM/DD/YYYY"
        ),
    },

    "COVERAGE EXPIRATION DATE": {
        "val_formula": '=AND(ISNUMBER(G1), G1 > DATE(1900, 1, 1))',
        "error_msg": (
            "Invalid Date Format - Enter date as MM/DD/YYYY"
        ),
    },

    "DATE OF SERVICE": {
        "val_formula": '=AND(ISNUMBER(H1), H1 > DATE(1900, 1, 1))',
        "error_msg": (
            "Invalid Date Format - Enter date as MM/DD/YYYY"
        )
    },

    "CPT/HCPCS/DENTAL CODE": {
        "val_formula": '=AND(IF(ISERROR(FIND(",",I1,1)),TRUE,FALSE),IF(ISERROR(FIND("-",I1,1)),TRUE,FALSE),IF(ISERROR(FIND(";",I1,1)),TRUE,FALSE),IF(ISERROR(FIND(CHAR(10),I1,1)),TRUE,FALSE))',
        "error_msg": (
            "Invalid entry: only enter one service code per record"
        )
    },

    "SERVICE CODE MODIFIER": {
        "val_formula": '=AND(LEN(J2)=2,IF(ISERROR(FIND(",",J2,1)),TRUE,FALSE),IF(ISERROR(FIND("-",J2,1)),TRUE,FALSE),IF(ISERROR(FIND(";",J2,1)),TRUE,FALSE),IF(ISERROR(FIND(CHAR(10),J2,1)),TRUE,FALSE))',
        "error_msg": (
            "Invalid entry: enter a single modifier two characters long"
        )
    },

    "AMOUNT DUE": {
        "val_formula": '=AND(ISNUMBER(M2), M2 >=0, M2 <= K2)',
        "error_msg": (
            "Invalid entry: amount due must not exceed billed amount"
        )
    }
}

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

    """
    Creating Excel
    """

    # Read in dataframe and format data
    server_connection_string = "data/medical_data.db"
    raw_dataframe = export_excel.create_dataframe(server_connection_string)

    # Rename the headers of the dataframe
    renamed_headers = export_excel.transform_header(raw_dataframe, mapping_dict=headers_mapping)
    #print(renamed_headers.head(4))

    # Reformat the date columns
    reformatted_date = export_excel.format_date_columns(renamed_headers, date_columns)
    dates = reformatted_date[["DATE OF BIRTH", "COVERAGE EXPIRATION DATE", "DATE OF SERVICE"]]
    #print(dates.head(3))

    # Insert columns for user entry
    final_df = export_excel.insert_headers(reformatted_date, inserted_columns)
    #print(final_df.head(3))

    # Save dataframe to excel file
    path = export_excel.create_sheet(final_df)

    """
    Modifying Excel
    """

    # Open up excel file to modify
    ws_to_modify = modify_excel.CustomSpreadsheet(filepath=path, row_range=1000)

    # Set the sheet to work on
    ws_to_modify.set_sheet("Sheet1")

    # Add data validation rules to the worksheet
    modify_excel.data_validation_handler(ws_to_modify, validation_dict)

    # Save the modified Excel file
    ws_to_modify.workbook.save("modified_output.xlsx")

    # Delete the unformatted Excel file
    os.remove(path)

    