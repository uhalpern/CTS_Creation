import json

validation_format_dict = {
    'data_validation': {
        "=AND(ISNUMBER(H1), H1 > DATE(1900, 1, 1))": {
            "columns": ["DATE OF BIRTH", "COVERAGE EXPIRATION DATE", "DATE OF SERVICE"],
            "error_msg": "Invalid Date Format - Enter date as MM/DD/YYYY"
        },
        '=AND(ISNUMBER(M2), M2 >=0, M2 <= $K2)': {
            "columns": ["AMOUNT DUE", "CONTRACTUAL ADJUSTMENT", "SPEND DOWN"],
            "error_msg": "Invalid entry: entered value must not exceed billed amount"
        }
    }
}