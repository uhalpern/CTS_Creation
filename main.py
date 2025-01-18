import json
import argparse
import datetime
from src import export_excel, setup_dataframe


def main(excel_file_name: str):

    # Loading configuration file with formatting parameters
    with open("config.json", encoding='utf-8') as f:
        config_dict =  json.load(f)

    """ 
    Setting Up Dataframe
    """

    # Read in dataframe and format data
    server_connection_string = "data/medical_data.db"
    raw_dataframe = setup_dataframe.create_dataframe(server_connection_string)

    # Rename the headers of the dataframe
    renamed_headers = setup_dataframe.transform_header(raw_dataframe, mapping_dict=config_dict["database_fields_to_headers"])
    # Reformat the date columns
    reformatted_date = setup_dataframe.format_date_columns(renamed_headers, config_dict["date_columns"])

    # Insert columns for user entry
    final_df = setup_dataframe.insert_headers(reformatted_date, config_dict["inserted_columns"])

    """ 
    Exporting Excel
    """

    # Get number of samples from query
    num_rows = final_df.shape[0]

    # Ingest Data into spreadsheet
    workbook = export_excel.insert_into_template(final_df, validation_format_dict=config_dict["formatting"])

    # Apply protection to sheet
    password = "test"
    export_excel.protection_handler(workbook, config_dict["unprotected_columns"], password, num_rows)

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

        # Raise error if file name does not have .xlsx handle
        if not file_name.endswith(".xlsx"):
            raise ValueError(f"Specified name: '{file_name}' is not formatted correctly. Filename should end with '.xlsx'")
    else:
        file_name = default_file_name + ".xlsx"

    main(file_name)
