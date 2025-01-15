import sqlite3
import os

# check if table already exists
if not os.path.exists("data/medical_data.db"):

    # Create a connection to a local SQLite database file
    connection = sqlite3.connect("data/medical_data.db")  # This will create 'medical_services.db' in working directory 
    cursor = connection.cursor()

    # Create the table
    cursor.execute('''
    CREATE TABLE medical_data (
        control_account_number VARCHAR(20),
        last_name VARCHAR(60),
        first_name VARCHAR(35),
        middle VARCHAR(25),
        date_of_birth DATE,
        medicaid_id VARCHAR(20),
        coverage_expiration_date DATE,
        date_of_service DATE,
        cpt_hcpcs_dental_code VARCHAR(48),
        service_code_modifier CHAR(3),
        billed_amount MONEY,
        spend_down MONEY,
        tpl_amount MONEY,
        tpl VARCHAR(60)
    );
    ''')

    # Insert sample data
    data = [
        ('1234567890', 'Smith', 'John', 'A', '1980-01-01', 'MED123456', '2024-12-31', '2024-06-15', 'D1234', 'MOD1', 150.73, 50.00, 20.00, 'Private Insurance'),
        ('0987654321', 'Doe', 'Jane', 'B', '1975-05-15', 'MED654321', '2024-11-30', '2024-06-20', 'C4567', 'MOD2', 200.23, 75.00, 30.00, 'Medicare'),
        ('5678901234', 'Brown', 'Michael', 'C', '1990-09-20', 'MED789012', '2025-01-15', '2024-06-25', 'H7890', 'MOD3', 300.47, 100.00, 40.00, 'None')
    ]

    data2 = [
        ('12345A', 'Doe', 'John', 'H', '2001-06-09', '13-015294-00', '2024-05-10', '2024-03-08', '10120', '', 132.79, 0.00, 0.00, ''),
        ('3938743', 'Du Bois', 'Harry', '', '1980-12-15', '13-002648-00', '2024-04-01', '2024-02-12', 'A4649', '', 77.90, 0.00, 0.00, ''),
        ('3938743', 'Du Bois', 'Harry', '', '1980-12-15', '13-002648-00', '2024-04-01', '2024-02-13', 'A4928', '', 91.89, 0.00, 0.00, ''),
        ('18393A', 'Kitsuragi', 'Kim', 'Kimball', '1985-08-26', '13-023421-00', '2024-06-23', '2024-02-12', '96116', '', 80.00, 0.00, 0.00, ''),
        ('18393A', 'Kitsuragi', 'Kim', 'Kimball', '1985-08-26', '13-023421-00', '2024-06-23', '2024-02-12', '96312', '', 90.20, 0.00, 0.00, ''),
        ('93649A', 'Jordan', 'Michael', 'B', '1987-02-09', '14-093843-00', '2024-07-02', '2024-05-05', 'V2799', '', 30.00, 0.00, 0.00, ''),
    ]

    cursor.executemany("INSERT INTO medical_data VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", data)

    # Save and close
    connection.commit()
    print("SQLite database created and populated!")

else:
    print("SQLite database already exists")

connection.close()
