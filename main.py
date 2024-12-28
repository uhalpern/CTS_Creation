from src import export_excel, modify_excel
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting import Rule

if __name__ == "__main__":

    # Read in dataframe and format data
    raw_dataframe = export_excel.create_dataframe()
