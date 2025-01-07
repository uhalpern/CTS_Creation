from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

workbook = load_workbook("CTS_Example.xlsx")

sheet = workbook["MAP or COFA"]

formula = "=AND(ISNUMBER(T1), T1 > DATE(1900, 1, 1))"
# dv = DataValidation(type="decimal",
#                     operator="between",
#                     formula1="0",
#                     formula2="1",
#                     showErrorMessage=True)

dv = DataValidation(type = "custom", 
                    formula1=formula,
                    showErrorMessage=True)
dv.error = "Error"
dv.add("T2:T1000")
sheet.add_data_validation(dv)

print(sheet.data_validations)

workbook.save("CTS_copy.xlsx")
