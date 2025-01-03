"""
Module: modify_excel
Description: This module handles the implementation of Data Validation and formatting
             for the generated Excel spreadsheet from export_excel.py. The module will
             make use of the openpyxl library which provides methods for modifying an Excel
             workbook. We use a class instantiation of an openpyxl.workbook and create
             custom methods to create the spreadsheet with the desired formatting and
             data validation.

Author: Urban Halpern
Date: 2024-12-26
Version: 0.5
"""

import os
import openpyxl.worksheet.datavalidation
from openpyxl import Workbook, load_workbook
from openpyxl.formatting import Rule
from openpyxl.styles import Color, PatternFill, Font, Border, numbers
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule 



class CustomSpreadsheet:
    def __init__(self, filepath: str = None, row_range: int = 1000):
        """
        Initializes a CustomSpreadsheet object to modify or create a workbook.

        Args:
            filepath (str): The path to an existing Excel file to load.
                            If not provided, a new workbook will be created.
            row_range: The number of rows that should be formatted

        Attributes:
            workbook (Workbook): The openpyxl Workbook object representing the loaded or newly created workbook.
            sheet (Worksheet): The active worksheet in the workbook.

        Raises:
            FileNotFoundError: If the filepath is provided but the file does not exist.
        """

        if filepath:
            if not os.path.exists(filepath):
                raise FileNotFoundError(f"The file '{filepath}' does not exist.")
            try:
                self.workbook = load_workbook(filepath)
            except Exception as e:
                raise IOError(f"An error occurred while loading the file: {e}")  # __init__ terminates immediately

        self.sheet = None
        self.range = row_range

    # Simple function to set the active sheet using the sheet name
    def set_sheet(self, sheet_name: str):
        self.sheet = self.workbook[sheet_name]
      

    # Applies the specified background color to the header row
    def set_header_color(self, color_code: str = "4472c4"):
        pass

    # Sets the specified height (in pixels) of the header row
    def set_header_height(self, header_row_height: float = 22.9):
        pass

    def set_column_width(self, multiplier: float = 1.2):
        """
        Sets the column widths based on the column name. The function will iterate
        through all the columns with data in them. The column width will be modified
        according to the cell with the longest value. This will fit each column length
        so that everything is displayed neatly and no information is cut off

        Args:
            multiplier (float): Controls extra whitespace to add to the column widths

        """

        pass

    def set_alternating_fill(self, stop_row: int = 1000, first_color: str = "d9e1f2", second_color: str = "b4c6e7"):
        """
        Creates an alternating pattern on the spreadsheet of filled rows and non-filled rows.
        The row colors will be formatted up to the defined stop_row with the specified color.

        Args:
            stop_row (int): formatting will apply to up to this row in the spreadsheet.
            first_color (str): to fill all the even rows
            second_color (str): to fill all the odd rows

        """
        pass

    def set_accounting_format(self, column_num: int, stop_row: int = 1000):
        """
        Sets a spreadsheet column to have a format equivalent to the accounting format in Excel
        Must be defined manually since openpyxl does not have an implementation for this format.
        stop_row must also be defined since formatting must be applied to each cell individually

        Args:
            stop_row (int): formatting will apply to up to this row in the spreadsheet.
            column_num (int): column to apply formatting to
        """

        pass

    def set_number_of_columns(self, num_columns: int = 22):
        """
        This method sets the number of columns to display on the spreadsheet. It will
        hide every column past the num_columns specified. It will have to iterate all
        the way up to the max column for Excel spreadsheets: 16285

        Args:
            num_columns (int): number of columns to display
        """

        pass

    def add_data_validation(self, formula: str,
                            col_to_validate: str,
                            error_message: str = "Error, the data you entered violates the data validation rules rule set for this cell."):
        """
        Adds a defined datavalidation object to the sheet based on a defined formula 
        The function will apply the data validation to all rows in the columns using self.range

        Args:
            formula (str): data validation formula to add to sheet
            col_to_validate (str): the letter of the column to apply the validation to
            error_message (str): error message to display when data validation is violated

        """

        # Define a data validation object
        dv = DataValidation(type="custom",
                            formula1=formula,
                            showErrorMessage=True)

        dv.error = error_message
        add_str = f'{col_to_validate}2:{col_to_validate}{self.range}'

        dv.add(add_str) # Apply the validation to range (EX: M1:M1000)

        # Add validation object to the sheet
        self.sheet.add_data_validation(dv)

    def add_conditional_formatting(self, formula: str, col_to_validate: int):
        """
        Adds a defined rule object to the sheet. The rule object should
        already be defined and the column to apply the formatting rule
         should be specified. The function will apply the formatting rule
         to all cells in the column

        Args:
            formula (str): conditional formatting formula used to creat Rule object
            col_to_validate (int): the integer value of the column to apply the validation to


        """

        # Define conditional formatting rule to add
        yellow_highlight = PatternFill(start_color='ffff00',
               end_color='ffff00',
               fill_type='solid')
        rule = FormulaRule(formula=[formula], stopIfTrue=True, fill=yellow_highlight)

        add_str = f'{col_to_validate}2:{col_to_validate}{self.range}'

        # Add conditional formatting to sheet
        self.sheet.conditional_formatting.add(add_str, rule)


    def get_column_letter(self, column_name: str) -> str:
        """
        Helper function to find the column letter from a specified column_name. If no matching
        header is found, will return None.

        Args:
            column_name (str): column name to extract header from

        Returns:
            column_letter (str): column letter associated with header.
        """

        column_letter = None

        # Iterate through all of the columns in the spreadsheet returning cells only up to row 2 for each column
        for col in self.sheet.iter_cols(max_row=2):

            # Extract the first row cell
            header_cell = col[0]
         
            # If header cell name found, return column letter
            if header_cell.value == column_name:
                return header_cell.column_letter

        return column_letter


def data_validation_handler(workbook: CustomSpreadsheet, validation_format_dict: dict) -> None:
    """
        This will ingest all of the data validation definitions for each of the columns
        into a CustomSpreadsheet object.

        Args:
            workbook (CustomSpreadsheet): The workbook to add the data validation rules to
            validation__format dict (dict): Holds the defined data validation and conditional formatting rules for each column.
                                            {header_name: {validation_params}}

        Returns:
            None
    """

    # For each header in the validation_format_dict, add the corresponding data validation
    for header in validation_format_dict:

        # Get the column letter from the header_name
        col = workbook.get_column_letter(header)

        # Raise exception for header not found
        if col is None:
            raise ValueError(f'{header} returned None as column letter. Check to make sure header exists in sheet')
      
        # Add a new data validation rule
        # Check if key exists
        try:
            val_formula = validation_format_dict[header]["val_formula"]
            error_message = validation_format_dict[header]["error_msg"]
        except KeyError:
            val_formula = None
            error_message = None

        # print(formula)
        # print(error_message)

        # Add data validation rule to sheet
        if val_formula is not None:
            workbook.add_data_validation(val_formula, col, error_message=error_message)

        # Create a new conditional formatting rule
        # Check if key exists
        try:
            format_formula = validation_format_dict[header]["format_formula"]
        except KeyError:
            format_formula = None

        # Add conditional formatting rule to sheet
        if format_formula is not None:
            workbook.add_conditional_formatting(format_formula, col)