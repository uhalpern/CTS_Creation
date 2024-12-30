from src import export_excel
import pandas as pd
from datetime import datetime


def test_create_dataframe():
    test_df = export_excel.create_dataframe("data/medical_data.db")

    try:
        # Assert that the dataframe has 14 columns
        assert test_df.shape[1] == 14
    except AssertionError:
        # Print the details when the assertion fails
        print(f"Assertion failed: Expected 14 columns, but got {test_df.shape[1]} columns.")
        print(f"Columns: {test_df.columns.tolist()}")  # Prints column names for debugging
        raise  # Re-raise the exception so the test still fails


def test_date_formatting():

    test_data = {
        "date_column": ["2024-12-30"],
        "value": [42]
    }

    test_df = pd.DataFrame(test_data)
    formatted_df = export_excel.format_date_columns(test_df, ["date_column"])

    format_str = '%m/%d/%y'
    date_str = formatted_df["date_column"].loc[0]
    print(date_str)

    try:
        # Assert that date format was changed correctly
        datetime.strptime(date_str, format_str)
    except ValueError:
        print("the date was not formatted correctly")
        raise
