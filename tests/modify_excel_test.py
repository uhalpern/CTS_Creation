import os
import pandas as pd
import pytest
from src import modify_excel, export_excel


def test_data_validation():

    test_data = {
        "date_column": ["2024-12-30"],
        "value": [42]
    }

    test_val_dict = {
        "date_column": {
            "val_formula": "=AND(ISNUMBER(H1), H1 > DATE(1900, 1, 1))",
            "error_msg": "Invalid Date Format - Enter date as MM/DD/YYYY",
            "format_formula": "=AND(ISNUMBER(H1), H1 > DATE(1900, 1, 1))"
        }
    }

    test_incorrect_val = {
         "missing_column": {
            "val_formula": "=AND(ISNUMBER(H1), H1 > DATE(1900, 1, 1))",
            "error_msg": "Invalid Date Format - Enter date as MM/DD/YYYY"
        }
    }

    test_df = pd.DataFrame(test_data)
    path = export_excel.create_sheet(test_df, file_name="test.xlsx")

    test_to_modify = modify_excel.CustomSpreadsheet(filepath=path, row_range=1000)

    # Set the sheet to work on
    test_to_modify.set_sheet("Sheet1")

    # Add data validation rules to the worksheet
    modify_excel.data_validation_handler(test_to_modify, test_val_dict)

    try:
        # Assert the data validation was added
        assert test_to_modify.sheet.data_validations.count == 1
    except AssertionError:
        print((
            f"Incorrect number of data validation of objects added: "
            f"{len(test_to_modify.sheet.data_validations.count)} present\n"
            f"1 expected"
        ))
        os.remove(path)
        raise

    try:
        # Assert formatting rule was added
        assert len(test_to_modify.sheet.conditional_formatting) == 1
    except AssertionError:
        print((
            f"Incorrect number of conditional formatting rules added: "
            f"{len(test_to_modify.sheet.conditional_formatting)} present\n"
            f"1 expected"
        ))
        raise

    # Test error handling for adding validaiton to missing column
    with pytest.raises(ValueError):
        modify_excel.data_validation_handler(test_to_modify, test_incorrect_val)

    # Delete the unformatted Excel file
    os.remove(path)

