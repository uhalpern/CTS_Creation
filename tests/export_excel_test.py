from src import export_excel


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
