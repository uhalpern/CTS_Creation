import os
import pandas as pd
import pytest
from src import modify_excel, setup_dataframe


class TestExcelModification:
    @classmethod
    def setup_class(cls):
        """
        Setup logic shared by all tests in the class.
        This runs once before any test methods are executed.
        """
        cls.test_data = {
            "date_column": ["2024-12-30"],
            "value": [42]
        }

        cls.test_df = pd.DataFrame(cls.test_data)
        cls.file_name = "test.xlsx"
        cls.sheet_name = "Sheet1"
        cls.row_range = 10

        # Create the Excel sheet and initialize the spreadsheet object
        cls.filepath = setup_dataframe.create_sheet(cls.test_df, file_name=cls.file_name)
        cls.spreadsheet = modify_excel.CustomSpreadsheet(filepath=cls.filepath, row_range=cls.row_range)
        cls.spreadsheet.set_sheet(cls.sheet_name)


    def test_data_validation(self):

        test_val_dict = {
            "date_column": {
                "data_validation": "=AND(ISNUMBER(H1), H1 > DATE(1900, 1, 1))",
                "error_msg": "Invalid Date Format - Enter date as MM/DD/YYYY",
                "conditional_format_formula": "=AND(ISNUMBER(H1), H1 > DATE(1900, 1, 1))"
            }
        }

        test_incorrect_val = {
            "missing_column": {
                "data_validation": "=AND(ISNUMBER(H1), H1 > DATE(1900, 1, 1))",
                "error_msg": "Invalid Date Format - Enter date as MM/DD/YYYY"
            }
        }

        # Add data validation rules to the worksheet
        modify_excel.formatting_handler(self.spreadsheet, test_val_dict)

        try:
            # Assert the data validation was added
            assert self.spreadsheet.sheet.data_validations.count == 1
        except AssertionError:
            print((
                f"Incorrect number of data validation of objects added: "
                f"{self.spreadsheet.sheet.data_validations.count} present\n"
                f"1 expected"
            ))
            os.remove(self.filepath)
            raise

        try:
            # Assert formatting rule was added
            assert len(self.spreadsheet.sheet.conditional_formatting) == 1
        except AssertionError:
            print((
                f"Incorrect number of conditional formatting rules added: "
                f"{len(self.spreadsheet.sheet.conditional_formatting)} present\n"
                f"1 expected"
            ))
            raise

        # Test error handling for adding validaiton to missing column
        with pytest.raises(ValueError):
            modify_excel.formatting_handler(self.spreadsheet, test_incorrect_val)


    def test_add_value_formula(self):

        value_formula = "=FLOOR($B{incorrect}*0.5*$B{nan})"
        
        # Test error handling of passing in a formula without the correct placeholders
        with pytest.raises(ValueError):
           self.spreadsheet.add_value_formula(value_formula=value_formula, col="value")

    
    def test_add_protection(self):
        
        unprotected = ["value"]
        modify_excel.protection_handler(self.spreadsheet, unprotected, password="test")

        assert self.spreadsheet.sheet["A2"].protection.locked is True
        assert self.spreadsheet.sheet["B2"].protection.locked is False

    def test_delete(self):
        # Delete the unformatted Excel file
        os.remove(self.filepath)
