## CONTROL/ACCOUNT \#
### Conditional Formatting
Highlight cells that are empty while the first or last name fields are filled.

```
# Expanded format
=AND(
    ISBLANK(A2),
    OR(
        NOT(ISBLANK(B2)),
        NOT(ISBLANK(C2))
    )
)

# Copy paste format
=AND(ISBLANK(A2),OR(NOT(ISBLANK(B2)),NOT(ISBLANK(C2))))
```
## MEDICAID ID
### Data Validation
Raise an error if the value entered does not follow the ##-######-## Medicaid ID format, or a 10-digit value that can be auto-formated into the above standard

```
# Expanded format
=OR(
	    AND(
            LEN(F1)=12,
            ISNUMBER(VALUE(LEFT(F1,2))),
            MID(F1,3,1)="-",
            ISNUMBER(VALUE(MID(F1,4,6))),
            MID(F1,10,1)="-",
            ISNUMBER(VALUE(RIGHT(F1,2)))
        ),
        AND(
    		LEN(F1)=10,
            NOT(VALUE(LEFT(F1))=0),
	    	ISNUMBER(VALUE(F1)
        )
	)
)

# Copy paste format
=OR(AND(LEN(F1)=12,ISNUMBER(VALUE(LEFT(F1,2))),MID(F1,3,1)="-",ISNUMBER(VALUE(MID(F1,4,6))),MID(F1,10,1)="-",ISNUMBER(VALUE(RIGHT(F1,2)))),AND(LEN(F1)=10,NOT(VALUE(LEFT(F1))=0),ISNUMBER(VALUE(F1))))
```

### Conditional Formatting
Highlight cells that do not follow the ##-######-## Medicaid ID format, or a 10-digit value that can be auto-formated into the above standard.

```
# Expanded format
=AND(
    NOT(AND(
            LEN(F2)=12,
            ISNUMBER(VALUE(LEFT(F2,2))),
            MID(F2,3,1)="-",
            ISNUMBER(VALUE(MID(F2,4,6))),
            MID(F2,10,1)="-",
            ISNUMBER(VALUE(RIGHT(F2,2)))
    )),
    NOT(AND(
            LEN(F2)=10,
            ISNUMBER(VALUE(F2))
    )),
    AND(
        ISBLANK(F2),
        NOT(ISBLANK(G2))
    )
)

# Copy paste format
=AND(NOT(AND(LEN(F2)=12,ISNUMBER(VALUE(LEFT(F2,2))),MID(F2,3,1)="-",ISNUMBER(VALUE(MID(F2,4,6))),MID(F2,10,1)="-",ISNUMBER(VALUE(RIGHT(F2,2))))),NOT(AND(LEN(F2)=10,ISNUMBER(VALUE(F2)))),AND(ISBLANK(F2),NOT(ISBLANK(G2))))
```


## Dates (DATE OF BIRTH, COVERAGE EXPIRATION DATE, DATE OF SERVICE)
### Data Validation
Raise error if any valid non-date value is entered. Valid dates start after 1/1/1900

```
# Expanded format
=AND(
    ISNUMBER(G1),
    G1 > DATE(1900, 1, 1)
)

# Copy Paste Format
=AND(ISNUMBER(G1), G1 > DATE(1900, 1, 1))
```

### Conditional Formatting
Highlight cell if it does not contain a number (that is converted into a date) or if it is empty while the neighboring cell is filled.

Note: this is applied to ranges E2:E10000,G2:G10000,H2:H10000 (DOB, Coverage Expiration Date, and Date of Service, respectively)
```
# Expanded format
=OR(
    AND(
        NOT(ISBLANK(E2)),
        NOT(ISNUMBER(E2))
    ),
    AND(
        ISBLANK(E2),
        NOT(ISBLANK(F2))
    )
)

# Copy and paste
=OR(AND(NOT(ISBLANK(E2)),NOT(ISNUMBER(E2))), AND(ISBLANK(E2),NOT(ISBLANK(F2))))
```

## CHECK FOR MULTIPLE CODES

### Data Validation
Raise error if commas, dashes, semicolons, and new line characters to prevent common "list" separators.

```
# Expanded format
=AND(
	IF(ISERROR(FIND(",",I1,1)),TRUE,FALSE),
	IF(ISERROR(FIND("-",I1,1)),TRUE,FALSE),
	IF(ISERROR(FIND(";",I1,1)),TRUE,FALSE),
	IF(ISERROR(FIND(CHAR(10),I1,1)),TRUE,FALSE)
)

# Copy paste format
=AND(IF(ISERROR(FIND(",",I1,1)),TRUE,FALSE),IF(ISERROR(FIND("-",I1,1)),TRUE,FALSE),IF(ISERROR(FIND(";",I1,1)),TRUE,FALSE),IF(ISERROR(FIND(CHAR(10),I1,1)),TRUE,FALSE))
```
### Conditional formatting
Highlight commas, dashes, semicolons, and new line characters to prevent common "list" separators.

```
# Expanded format
=OR(
	AND(
		NOT(ISBLANK(I2)),
		NOT(AND(
			IF(ISERROR(FIND(",",I2,1)),TRUE,FALSE),
			IF(ISERROR(FIND("-",I2,1)),TRUE,FALSE),
			IF(ISERROR(FIND(";",I2,1)),TRUE,FALSE),
			IF(ISERROR(FIND(CHAR(10),I2,1)),TRUE,FALSE))
		)
    ), 
	AND(
		ISBLANK(I2),
		NOT(ISBLANK(J2))
    )
)

# Copy paste format
=OR(AND(NOT(ISBLANK(I2)),NOT(AND(IF(ISERROR(FIND(",",I2,1)),TRUE,FALSE),IF(ISERROR(FIND("-",I2,1)),TRUE,FALSE),IF(ISERROR(FIND(";",I2,1)),TRUE,FALSE),IF(ISERROR(FIND(CHAR(10),I2,1)),TRUE,FALSE)))), AND(ISBLANK(I2),NOT(ISBLANK(J2))))
```

## SERVICE CODE MODIFIER

### Data Validation
Raise error if the cell is not two characters long and contains commas, dashes, or semicolons (characters that typically denote ranges or lists).

```
# Expanded format
=AND(
    LEN(J2)=2,
    IF(ISERROR(FIND(",",J2,1)),TRUE,FALSE),
    IF(ISERROR(FIND("-",J2,1)),TRUE,FALSE),
    IF(ISERROR(FIND(";",J2,1)),TRUE,FALSE),
    IF(ISERROR(FIND(CHAR(10),J2,1)),TRUE,FALSE)
)

# Copy paste format
=AND(LEN(J2)=2,IF(ISERROR(FIND(",",J2,1)),TRUE,FALSE),IF(ISERROR(FIND("-",J2,1)),TRUE,FALSE),IF(ISERROR(FIND(";",J2,1)),TRUE,FALSE),IF(ISERROR(FIND(CHAR(10),J2,1)),TRUE,FALSE))

```
### Conditional Formatting
Highlight cells that are not two characters long and not blank.

```
# Expanded format
=AND(
    NOT(LEN(J2)=2),
    NOT(ISBLANK(J2))
)

# Copy paste format
=AND(NOT(LEN(J2)=2),NOT(ISBLANK(J2)))
```

## BILLED AMOUNT
### Data Validation
>

## Contractual Adjustment, Amount Due, Spend Down

### Data Validation
Raise an error if any of the the entered values are greater than the billed amount.

```
# Expanded format
=AND(
    ISNUMBER(M2),
    M2 >=0,
    M2 <= $K2
)

# Copy Paste Format
=AND(ISNUMBER(M2), M2 >=0, M2 <= $K2)
```

### Conditional formatting
Highlight individual amount cells that are greater than the billed amount or all of the amount cells if their sum is greater than the billed amount.

Note: This is applied to ranges L2:L10000, O2:O10000, P2:P10000 (Contractual Adjustment, Amount Due, and Spend Down columns, respectively) and will replace L2 for all ranges.
```
# Expanded format
=OR(
	NOT(AND(
		IFERROR(M2<=$K2,FALSE),
		IFERROR($M2+$P2+$Q2+$S2<=$K2,FALSE)
    )),
	AND(
		NOT(ISBLANK(M2)),
		NOT(ISNUMBER(M2))
    )
)

# Copy paste format
=OR(NOT(AND(IFERROR(M2<=$K2,FALSE),IFERROR($M2+$P2+$Q2+$S2<=$K2,FALSE))),AND(NOT(ISBLANK(M2)),NOT(ISNUMBER(M2))))
```

## TPL AMOUNT

### Conditional formatting
Highlight cell if it does not contain a number or if it is empty while the neighboring cell is filled.

```
# Expanded format
=OR(
	AND(
		NOT(ISBLANK(Q2)),
		NOT(ISNUMBER(Q2))
	), 
	AND(
		ISBLANK(Q2),
		NOT(ISBLANK(R2))
	)
)
# Copy paste format
=OR(AND(NOT(ISBLANK(Q2)),NOT(ISNUMBER(Q2))),AND(ISBLANK(Q2),NOT(ISBLANK(R2))))
```

