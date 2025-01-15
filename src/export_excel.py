"""
Module: export_excel
Description: This module handles the creation of a claims transmittal spreadsheet
             from a remote SQL table. It provides functions for defining header
             names, and correct formatting of the DATA before entering the spreadsheet.
             Note: this is not meant to affect the format of the spreadsheet
             itself (datavalidation, conditional formatting).

Author: Urban Halpern
Date: 2024-12-24
"""

import sqlite3
import os
import openpyxl.cell
import openpyxl.workbook
import openpyxl.worksheet
import openpyxl.worksheet.worksheet
import pandas as pd
import openpyxl
from openpyxl import load_workbook



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


def process_row(sheet: openpyxl.worksheet.worksheet, dataframe_row: list, sheet_row: tuple,
                validation_format_dict: dict) -> None:
    """
    Ingests a single dataframe row into a single spreadsheet row in the provided sheet.

    Args:
        sheet: Sheet object that dataframe is ingesting into
        dataframe_row (list): preprocessed dataframe row to ingest into sheet_row
        sheet_row (tuple): tuple of cells representing a spreadsheet row

    """

    for cell in sheet_row:
        
        # Prevent overwriting of cells that have formulas
        if cell.value is None:

            # Get value from dataframe by indexing with the cell column number
            val_to_insert = dataframe_row[cell.column]

            # Insert the Value
            cell.value = val_to_insert

            # Add formatting
            get_format(sheet, cell, validation_format_dict)


def get_format(sheet: openpyxl.worksheet.worksheet, cell: openpyxl.cell.cell.Cell, 
               validation_format_dict: dict) -> None:
    """
    Returns the format string of a cell from the validation_format_dict

    Args:
        sheet: (openpyxl.worksheet.worksheet): spreadsheet object for accessing header cell
        cell (openpyxl.cell.cell.Cell): cell to get formatting for
        validation_format_dict (dict): dictionary that holds formatting for each column

    """

    # Get the columns header from the sheet
    column_letter = cell.column_letter
    header = sheet[f"{column_letter}1"]

    # Look for the header in the validation_format_dict
    format_rules = validation_format_dict.get(header)

    # Apply style formatting if the header and style_format field is found
    if format_rules is not None:
        style_format = validation_format_dict[header].get("style_format")

        if style_format is not None:
            cell.number_format = style_format


def insert_into_template(final_df: pd.DataFrame, validation_format_dict: dict, workbook_name: str = "CTS_Insert_Example.xlsx"):
    """
    Inserts data into the template spreadsheet using data from the final_df row by row

    Args:
        final_df (pandas.dataframe): dataframe which holds transformed data from SQL query
        validation_format_dict (dict): dictionary that holds formatting for each column
        workbook_name (str): name of workbook that will be saved after data is ingested

    """

    parent_dir = os.path.abspath(os.path.join(os.getcwd()))
    sheets_directory = os.path.join(parent_dir, 'generated_sheets')

    template_file_path = os.path.join(parent_dir, 'CTS_Example_Template')

    workbook = load_workbook()

    pass