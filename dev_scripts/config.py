import json
import os

validation_format_dict = {
    'data_validation': {
        "=AND(ISNUMBER(H1), H1 > DATE(1900, 1, 1))": {
            "columns": ["DATE OF BIRTH", "COVERAGE EXPIRATION DATE", "DATE OF SERVICE"],
            "error_msg": "Invalid Date Format - Enter date as MM/DD/YYYY"
        },
        '=AND(ISNUMBER(M2), M2 >=0, M2 <= $K2)': {
            "columns": ["AMOUNT DUE", "CONTRACTUAL ADJUSTMENT", "SPEND DOWN"],
            "error_msg": "Invalid entry: entered value must not exceed billed amount"
        },
        '=OR(AND(LEN(F2)=12,ISNUMBER(VALUE(LEFT(F2,2))),MID(F2,3,1)="-",ISNUMBER(VALUE(MID(F2,4,6))),MID(F2,10,1)="-",ISNUMBER(VALUE(RIGHT(F2,2)))),AND(LEN(F2)=10,NOT(VALUE(LEFT(F2))=0),ISNUMBER(VALUE(F2))))': {
            "columns": ["MEDICAID ID"],
            "error_msg": "Input must be in valid medicaid id format (**-******-**)"
        },
        '=AND(IF(ISERROR(FIND(",",I1,1)),TRUE,FALSE),IF(ISERROR(FIND("-",I1,1)),TRUE,FALSE),IF(ISERROR(FIND(";",I1,1)),TRUE,FALSE),IF(ISERROR(FIND(CHAR(10),I1,1)),TRUE,FALSE))': {
            "columns": ["CPT/HCPCS/DENTAL CODE"],
            "error_msg": "Invalid entry: only enter one service code per record"
        },
        '=AND(LEN(J2)=2,IF(ISERROR(FIND(",",J2,1)),TRUE,FALSE),IF(ISERROR(FIND("-",J2,1)),TRUE,FALSE),IF(ISERROR(FIND(";",J2,1)),TRUE,FALSE),IF(ISERROR(FIND(CHAR(10),J2,1)),TRUE,FALSE))': {
            "columns": ["SERVICE CODE MODIFIER"],
            "error_msg": "Invalid entry: enter a single modifier two characters long",
        }
    },

    "conditional_formatting": {
        '=AND(ISBLANK(A2),OR(NOT(ISBLANK(B2)),NOT(ISBLANK(C2))))': {
            "columns": ["CONTROL/ACCOUNT #"]
        },
        '=AND(NOT(AND(LEN(F2)=12,ISNUMBER(VALUE(LEFT(F2,2))),MID(F2,3,1)="-",ISNUMBER(VALUE(MID(F2,4,6))),MID(F2,10,1)="-",ISNUMBER(VALUE(RIGHT(F2,2))))),NOT(AND(LEN(F2)=10,ISNUMBER(VALUE(F2)))),AND(ISBLANK(F2),NOT(ISBLANK(G2))))': {
            "columns": ["MEDICAID ID"]
        },
        '=AND(NOT(LEN(J2)=2),NOT(ISBLANK(J2)))': {
            "columns": ["SERVICE CODE MODIFIER"]
        },
        '=OR(NOT(AND(IFERROR(M2<=$K2,FALSE),IFERROR($M2+$P2+$Q2+$S2<=$K2,FALSE))),AND(NOT(ISBLANK(M2)),NOT(ISNUMBER(M2))))': {
            "columns": ["AMOUNT DUE, CONTRACTUAL ADJUSTMENT, SPEND DOWN"]
        },
        '=OR(AND(NOT(ISBLANK(Q2)),NOT(ISNUMBER(Q2))),AND(ISBLANK(Q2),NOT(ISBLANK(R2))))': {
            "columns": ["TPL AMOUNT"]
        }
    },

    "style_formatting": {
        '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)': {
            "columns": ['BILLED AMOUNT', 'GRAND TOTAL', 'AMOUNT DUE', 
                     'LOCAL SHARE', 'FEDERAL SHARE', 'SPEND DOWN', 
                     'TPL AMOUNT']
        },
        '00-000000-00': {
            "columns": ["MEDICAID ID"]
        }
    },

    "value_formatting": {
        '=SUM(M{row},P{row},Q{row},S{row})': {
            "columns": ["GRAND TOTAL"]
        },
        '=FLOOR($M{row}*0.17,0.01)': {
            "columns": ["LOCAL SHARE"]
        },
        '=FLOOR($M{row}*0.83,0.01)': {
            "columns": ["FEDERAL SHARE"]
        }
    }
}

validation_format_dict2 = {
    "CONTROL/ACCOUNT #": {
        "conditional_format_formula": '=AND(ISBLANK(A2),OR(NOT(ISBLANK(B2)),NOT(ISBLANK(C2))))'
    },

    "MEDICAID ID": {
        "data_validation": '=OR(AND(LEN(F2)=12,ISNUMBER(VALUE(LEFT(F2,2))),MID(F2,3,1)="-",ISNUMBER(VALUE(MID(F2,4,6))),MID(F2,10,1)="-",ISNUMBER(VALUE(RIGHT(F2,2)))),AND(LEN(F2)=10,NOT(VALUE(LEFT(F2))=0),ISNUMBER(VALUE(F2))))',
        "error_msg": "Input must be in valid medicaid id format (**-******-**)",
        "conditional_format_formula": '=AND(NOT(AND(LEN(F2)=12,ISNUMBER(VALUE(LEFT(F2,2))),MID(F2,3,1)="-",ISNUMBER(VALUE(MID(F2,4,6))),MID(F2,10,1)="-",ISNUMBER(VALUE(RIGHT(F2,2))))),NOT(AND(LEN(F2)=10,ISNUMBER(VALUE(F2)))),AND(ISBLANK(F2),NOT(ISBLANK(G2))))',
        "style_format": '00-000000-00',
    },

    "DATE OF BIRTH": {
        "data_validation": '=AND(ISNUMBER(E2), E2 > DATE(1900, 1, 1))',
        "error_msg": "Invalid Date Format - Enter date as MM/DD/YYYY",
        "style_format": 'MM/DD/YY'
    },

    "COVERAGE EXPIRATION DATE": {
        "data_validation": '=AND(ISNUMBER(G2), G2 > DATE(1900, 1, 1))',
        "error_msg": "Invalid Date Format - Enter date as MM/DD/YYYY",
        "style_format": 'MM/DD/YY'
    },

    "DATE OF SERVICE": {
        "data_validation": '=AND(ISNUMBER(H2), H2 > DATE(1900, 1, 1))',
        "error_msg": "Invalid Date Format - Enter date as MM/DD/YYYY",
        "style_format": 'MM/DD/YY'
    },

    "CPT/HCPCS/DENTAL CODE": {
        "data_validation": '=AND(IF(ISERROR(FIND(",",I2,1)),TRUE,FALSE),IF(ISERROR(FIND("-",I2,1)),TRUE,FALSE),IF(ISERROR(FIND(";",I2,1)),TRUE,FALSE),IF(ISERROR(FIND(CHAR(10),I2,1)),TRUE,FALSE))',
        "error_msg": "Invalid entry: only enter one service code per record"
    },

    "SERVICE CODE MODIFIER": {
        "data_validation": '=AND(LEN(J2)=2,IF(ISERROR(FIND(",",J2,1)),TRUE,FALSE),IF(ISERROR(FIND("-",J2,1)),TRUE,FALSE),IF(ISERROR(FIND(";",J2,1)),TRUE,FALSE),IF(ISERROR(FIND(CHAR(10),J2,1)),TRUE,FALSE))',
        "error_msg": "Invalid entry: enter a single modifier two characters long",
        "conditional_format_formula": '=AND(NOT(LEN(J2)=2),NOT(ISBLANK(J2)))'
    },

    "BILLED AMOUNT": {
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    },

    "GRAND TOTAL": {
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
        "value_formula": "=SUM(M{row},P{row},Q{row},S{row})"
    },

    "AMOUNT DUE": {
        "data_validation": '=AND(ISNUMBER(M2), M2 >=0, M2 <= $K2)',
        "error_msg": "Invalid entry: amount due must not exceed billed amount",
        "conditional_format_formula": '=OR(NOT(AND(IFERROR(M2<=$K2,FALSE),IFERROR($M2+$P2+$Q2+$S2<=$K2,FALSE))),AND(NOT(ISBLANK(M2)),NOT(ISNUMBER(M2))))',
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    },

    "LOCAL SHARE": {
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
        "value_formula": '=FLOOR($M{row}*0.17,0.01)',
    },

    "FEDERAL SHARE": {
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
        "value_formula": '=FLOOR($M{row}*0.83,0.01)'
    },

    "SPEND DOWN": {
        "data_validation": '=AND(ISNUMBER(P2), P2 >=0, P2 <= $K2)',
        "error_msg": "Invalid entry: amount due must not exceed billed amount",
        "conditional_format_formula": '=OR(NOT(AND(IFERROR(M2<=$K2,FALSE),IFERROR($M2+$P2+$Q2+$S2<=$K2,FALSE))),AND(NOT(ISBLANK(M2)),NOT(ISNUMBER(M2))))',
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    },

    "TPL AMOUNT": {
        "format formula": '=OR(AND(NOT(ISBLANK(Q2)),NOT(ISNUMBER(Q2))),AND(ISBLANK(Q2),NOT(ISBLANK(R2))))',
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    },

    "CONTRACTUAL ADJUSTMENT": {
        "data_validation": '=AND(ISNUMBER(S2), S2 >=0, S2 <= $K2)',
        "error_msg": "Invalid entry: amount due must not exceed billed amount",
        "conditional_format_formula": '=OR(NOT(AND(IFERROR(M2<=$K2,FALSE),IFERROR($M2+$P2+$Q2+$S2<=$K2,FALSE))),AND(NOT(ISBLANK(M2)),NOT(ISNUMBER(M2))))',
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    },
}

validation_format_dict3 = {
    "MEDICAID ID": {
        "style_format": '00-000000-00',
        "alignment": "right"
    },

    "DATE OF BIRTH": {
        "style_format": 'MM/DD/YY'
    },

    "COVERAGE EXPIRATION DATE": {
        "style_format": 'MM/DD/YY'
    },

    "DATE OF SERVICE": {
        "style_format": 'MM/DD/YY'
    },

    "CPT/HCPCS/DENTAL CODE": {
        "alignment": "right"
    },

    "SERVICE CODE MODIFIER": {
        "alignment": "right"
    },

    "BILLED AMOUNT": {
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    },

    "GRAND TOTAL": {
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    },

    "AMOUNT DUE": {
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    },

    "LOCAL SHARE": {
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    },

    "FEDERAL SHARE": {
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    },

    "SPEND DOWN": {
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    },

    "TPL AMOUNT": {
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    },

    "CONTRACTUAL ADJUSTMENT": {
        "style_format": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    },
}

# Get path to parent directory
parent_dir = os.path.abspath(os.path.join(os.getcwd(), ".."))

# Define the path one directory up
file_path = os.path.join(parent_dir, 'config.json')

# Save the file in parent directory
with open(file_path, 'w', encoding="utf-8") as f:
    json.dump(validation_format_dict3, f, indent=4)
    