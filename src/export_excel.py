"""
Module: export_excel
Description: This module handles the creation of a claims transmittal spreadsheet
             from a remote SQL table. It provides functions for defining header
             names, and correct formatting of the DATA before entering the spreadsheet.
             Note: this is not meant to affect the format of the spreadsheet
             itself (datavalidation, conditional formatting).

Author: Urban Halpern
Date: 2024-12-24
Version: 0.5
"""

import sqlite3, os
import pandas as pd



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


def format_date(df: pd.DataFrame, column_name: str) -> None:
    """
    The DATE var type in SQL has the format YYYY-MM-DD. For the spreadsheet,
    we want the date to be in the format MM/DD/YY. This function will change
    the column passed in by the column_name parameter. These transformations
    will be done in place and will return nothing.

    Args:
        df (pd.DataFrame): dataframe with date column to format
        column_name (str): column to format
    Returns:
        None: This function modifies the Dataframe in place
    """

    # Convert columns to datetime
    df[column_name] = pd.to_datetime(df[column_name])

    # Extract the date property out of the datetime object
    df[column_name] = df[column_name].dt.date
    

    return None


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
        format_date(df, name)
    
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
