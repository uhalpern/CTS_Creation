"""
Module: export_excel
Description: This module handles the creation of a claims transmittal spreadsheet
             from a remote SQL table. It provides functions for defining header
             names, and correct formatting of the data before entering the spreadsheet.
             It also provides funcitonality for ingesting data into a spreadsheet template.

Author: Urban Halpern
Version: 2024-01-16
"""

import os

import sqlite3
import pandas as pd

import openpyxl
import openpyxl.workbook
from openpyxl import load_workbook
from openpyxl.styles import Protection, Alignment


def create_dataframe(connection_string: str) -> pd.DataFrame:
    """
    Connects to MS SQL database and queries table information into dataframe.
    After reading in the data, close the connection to the SQL server

    Note: For now, made up data will be added into the spreadsheet

    Args:
        connection_string (str): in this case, it is just a path but represents sql server connection str
    Returns:
        raw_dataframe (pd.DataFrame): Dataframe that has the raw, un-formatted data
        from the SQL database. Each column will likely be objects.
    """

    connection = sqlite3.connect(connection_string)

    # Query the database
    query = "SELECT * FROM medical_data;"
    df = pd.read_sql_query(query, connection)

    return df


def transform_header(df: pd.DataFrame, mapping_dict: dict) -> pd.DataFrame:
    """
    Transforms the headers of a DataFrame using a mapping dictionary. The
    SQL table variables are usually formatted with snake_case. This function
    will transform these headers into the specified format for the spreadsheet
    using the mapping dict.

    Args:
        df (pd.Dataframe): dataframe with headers to change
        mapping_dict (dict): Has the mapping between SQL table var names and excel header names
    Returns:
        new_headers (pd.Dataframe): copy of df with updated header names
    """
    new_headers = df.rename(columns=mapping_dict, inplace=False)
    return new_headers


def format_date_columns(df: pd.DataFrame, column_names_list: list) -> pd.DataFrame:
    """
    The DATE var type in SQL has the format YYYY-MM-DD. For the spreadsheet,
    we want the date to be in the format MM/DD/YY. This function will make use
    of the format_date column to convert all the columns in column_names_list
    into the correct format.

    Note: For now a fake SQL table will be read using SQLite

    Args:
        df (pd.DataFrame): dataframe with date columns to format
        column_names_list (list): list of all the column names (str) to format
    Returns:
        formatted_dates (pd.DataFrame): Copy of df with formatted dates
    """

    # Format each date column in place
    for name in column_names_list:
        # Convert columns to datetime
        df[name] = pd.to_datetime(df[name])

    return df


def insert_headers(df: pd.DataFrame, columns_to_insert: dict) -> pd.DataFrame:
    """
    Some of the columns in the output spreadsheet are not extracted from the SQL
    database. Instead, these columns will be for user entry. This function
    will insert these columns into a copy of df and return it.

    Args:
        df (pd.DataFrame): dataframe with date columns to format
        columns_to_insert (dict): dict of column_name: position mapping
    Returns:
        columns_inserted (pd.DataFrame): copy of df with inserted columns
    """

    columns_inserted = df.copy()

    for new_column in columns_to_insert:
        columns_inserted.insert(columns_to_insert[new_column], new_column, None)

    return columns_inserted


def create_sheet(final_df: pd.DataFrame, sheet_name: str = "Sheet1",
                 file_name: str = "output.xlsx", file_path = ".\\generated_sheets") -> None:
    """
    Saves the dataframe to an Excel file using the openpyxl engine. The Excel file
    name should be unique to not conflict with other generated Excel files. The function
    will make sure the file is unique in the specified directory by raising an exception

    Args:
        final_df: The final dataframe with all the formatted data to convert into a spreadsheet.
        sheet_name (str): The name of the sheet to create.
        file_name (str): Name of the excel workbook

    Returns:
        full_path (str): Path where spreadsheet was saved
    """
  
    # Create directory if it does not already exist
    if not os.path.exists(file_path):
        os.mkdir(file_path)

    # Combine path and file name
    full_path = os.path.join(file_path, file_name)
 
    # Error handling for sheet already existing
    if os.path.exists(full_path):
        raise FileExistsError(f'The file already exists: {full_path}')

    # Save the sheet using the full path
    final_df.to_excel(full_path, sheet_name=sheet_name, engine='openpyxl', index=False)
    print(f'Sheet saved to {file_name} with sheet name: {sheet_name} at {full_path}')

    return full_path


def get_format(sheet: openpyxl.worksheet.worksheet, cell: openpyxl.cell.cell.Cell,
               validation_format_dict: dict, header: str) -> None:
    """
    Returns the format string of a cell from the validation_format_dict

    Args:
        sheet: (openpyxl.worksheet.worksheet): spreadsheet object for accessing header cell
        cell (openpyxl.cell.cell.Cell): cell to get formatting for
        validation_format_dict (dict): dictionary that holds formatting for each column
        header (str): header name to lookup in validation_format_dict

    """

    # # Get the columns header from the sheet
    # column_letter = cell.column_letter
    # header_cell = sheet[f"{column_letter}1"]
    # header = header_cell.value

    # Look for the header in the validation_format_dict
    format_rules = validation_format_dict.get(header)

    # Apply style formatting if the header and style_format field is found
    if format_rules is not None:
        style_format = format_rules.get("style_format")
        alignment = format_rules.get("alignment")

        # Style Formatting
        if style_format is not None:
            cell.number_format = style_format

        # Alignment
        if alignment is not None:
            cell.alignment = Alignment(horizontal=alignment)


def insert_into_template(final_df: pd.DataFrame, validation_format_dict: dict) -> openpyxl.workbook.workbook.Workbook:
    """
    Inserts data into the template spreadsheet using data from the final_df row by row

    Args:
        final_df (pandas.dataframe): dataframe which holds transformed data from SQL query
        validation_format_dict (dict): dictionary that holds formatting for each column
        workbook_name (str): name of workbook that will be saved after data is ingested

    Returns:
        workbook (openpyxl.workbook.workbook.Workbook): workbook with ingested data

    """

    # Get parent dir of repo to access generated_sheets dir
    parent_dir = os.path.abspath(os.path.join(os.getcwd()))

    # Access the template file
    template_file_path = os.path.join(parent_dir, 'CTS_Example_Template.xlsx')

    workbook = load_workbook(template_file_path)
    sheet = workbook["MAP or COFA"]

    # Iterate though the columns inthe dataframe
    for col_name in final_df.columns:

        # Find the column in the sheet
        col_letter = get_column_letter(sheet, col_name)
        col_data = final_df[col_name]

        # iterate through the rows in the column, skipping the header cell and using one-based indexing
        for row_idx, value in enumerate(col_data, start=2):

            cell = sheet[f"{col_letter}{row_idx}"]

            # Skip cells with formulas
            if not cell.data_type == "f":
                cell.value = value

                get_format(sheet, cell, validation_format_dict, col_name)

    return workbook


def save_workbook(workbook: openpyxl.workbook.Workbook, workbook_name: str = "CTS_Insert_Example") -> None:
    """
    Saves the workbook to the specified path and checks if file already exists

    Args:
        workbook (openpyxl.workbook.Workbook): The workbook object to save

    """
    # Get parent dir of repo to access generated_sheets dir
    parent_dir = os.path.abspath(os.path.join(os.getcwd()))
    sheets_directory = os.path.join(parent_dir, 'generated_sheets')

    # Define path to save workbook
    save_path = os.path.join(sheets_directory, workbook_name)
    if os.path.exists(save_path):
        raise FileExistsError(f'The file already exists: {save_path}')

    # Save workbook
    workbook.save(save_path)
    print(f'Sheet saved to {workbook_name} at {save_path}') 

    return save_path


def protection_handler(workbook: openpyxl.workbook.Workbook, cols_to_unprotect: list,
                       password: str = "test", row_range: int = 50) -> None:
    """
    Un-protects columns that do not need protection

    Args:
        workbook (openpyxl.workbook.Workbook): The workbook with columns to unprotect
        cols_to_unprotect (list): List of column headers to unprotect
        password (str): password to unlock the sheet
        range (int): range of cells in column to unprotect

    Returns:
        None
    """

    # Protect all cells and set password
    sheet = workbook["MAP or COFA"]
    sheet.protection.enable()
    sheet.protection.password = password

    for column in cols_to_unprotect:
        col_letter = get_column_letter(sheet, column)

        unlock_column(sheet, col_letter, row_range)


def get_column_letter(sheet: openpyxl.worksheet.worksheet.Worksheet, column_name: str) -> str:
    """
    Helper function to find the column letter from a specified column_name. If no matching
    header is found, will return None.

    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): worksheet object to extract column letter from
        column_name (str): column name to extract header from

    Returns:
        column_letter (str): column letter associated with header.
    """

    column_letter = None

    # Iterate through all of the columns in the spreadsheet returning cells only up to row 2 for each column
    for col in sheet.iter_cols(max_row=2):

        # Extract the first row cell
        header_cell = col[0]
      
        # If header cell name found, return column letter
        if header_cell.value == column_name:
            return header_cell.column_letter
  
    # Raise error if the column was not found in the sheet.
    if column_letter is None:
        raise ValueError(f'Specified Column: {column_name} not found in sheet.')


def unlock_column(sheet: openpyxl.worksheet.worksheet.Worksheet, column_to_unlock: str, row_range: int):
    """
    This method unlocks the cells in a column to allow user entry
    assuming the sheet was already set to protection mode.
    The header cell will remain locked for each column.

    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): sheet to unlock columns in
        column_to_unlock (str): letter of the column to unlock
        range (int): Range of cells to unlock
    """

    # iterate through cells in the column, skipping the header cell
    for cell in sheet[column_to_unlock][1:row_range]:
        cell.protection = Protection(locked=False)  # Unlock the cell
