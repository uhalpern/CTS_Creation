from src import setup_dataframe
import pandas as pd
import os
import pytest

def test_create_dataframe():
    test_df = setup_dataframe.create_dataframe("data/test_medical_data.db")

    try:
        # Assert that the dataframe has 14 columns
        assert test_df.shape[1] == 14
    except AssertionError:
        # Print the details when the assertion fails
        print(f"Assertion failed: Expected 14 columns, but got {test_df.shape[1]} columns.")
        print(f"Columns: {test_df.columns.tolist()}")  # Prints column names for debugging
        raise  # Re-raise the exception so the test still fails

def test_create_sheet():

    test_data = {
        "date_column": ["2024-12-30"],
        "value": [42]
    }

    test_df = pd.DataFrame(test_data)

    # Test successful creation:
    
    path = setup_dataframe.create_sheet(test_df, file_name="test.xlsx")

    try:
        # Assert the file was created successfully
        assert os.path.isfile(path)
    except AssertionError:
        print(f"The file was not ceated successfully: {path}")
        raise

    # Test error handling for duplicate file
    with pytest.raises(FileExistsError):
            setup_dataframe.create_sheet(test_df, file_name="test.xlsx")

    # Cleanup the created file and directory
    if os.path.exists(path):
        os.remove(path)