import sqlite3
import os

# Create a connection to a local SQLite database file
connection = sqlite3.connect("data/medical_data.db")  # This will create 'medical_services.db' in your working directory
cursor = connection.cursor()

# check if table already exists
if not os.path.exists("data/medical_data.db"):
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
    cursor.executemany("INSERT INTO medical_data VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", data)

    # Save and close
    connection.commit()
    print("SQLite database created and populated!")

else:
    print("SQLite database already exists")
connection.close()

