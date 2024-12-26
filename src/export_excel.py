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

import pandas as pd


def create_dataframe() -> pd.DataFrame:
    """
    Connects to MS SQL database and queries table information into dataframe.
    After reading in the data, close the connection to the SQL server

    Args:

    Returns:
        raw_dataframe (pd.DataFrame): Dataframe that has the raw, un-formatted data
        from the SQL database. Each column will likely be objects.
    """

    pass


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
    pass


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
    # convert datetime objects back to string with specified format
    df[column_name] = df[column_name].dt.strftime('%m/%d/%y')

    return None


def format_date_columns(df: pd.DataFrame, column_names_list: list) -> pd.DataFrame:
    """
    The DATE var type in SQL has the format YYYY-MM-DD. For the spreadsheet,
    we want the date to be in the format MM/DD/YY. This function will make use
    of the format_date column to convert all the columns in column_names_list
    into the correct format.

    Args:
        df (pd.DataFrame): dataframe with date columns to format
        column_names_list (list): list of all the column names (str) to format
    Returns:
        formatted_dates (pd.DataFrame): Copy of df with formatted dates
    """

    pass


def insert_headers(df: pd.DataFrame, columns_to_insert: list) -> pd.DataFrame:
    """
    Some of the columns in the output spreadsheet are not extracted from the SQL
    database. Instead, these columns will be for user entry. This function
    will insert these columns into a copy of df and return it.

    Args:
        df (pd.DataFrame): dataframe with date columns to format
        columns_to_insert (list): list of all the column names (str) to insert
    Returns:
        columns_inserted (pd.DataFrame): copy of df with inserted columns
    """

    pass


def create_sheet(final_df: pd.DataFrame, sheet_name: str) -> None:
    """
    Saves the dataframe to an Excel file using the openpyxl engine. The Excel file
    name should be unique to not conflict with other generated Excel files.

    Args:
        final_df: The final dataframe with all the formatted data to convert into a spreadsheet.
        sheet_name (str): The name of the sheet to create.

    Returns:
        None
    """
    pass
