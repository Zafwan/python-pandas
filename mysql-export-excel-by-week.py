import mysql.connector
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os

# MySQL Database Configuration
db_config = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'your_database_name'
}

# Connect to the MySQL database
try:
    conn = mysql.connector.connect(**db_config)
    cursor = conn.cursor()
except mysql.connector.Error as err:
    print(f"Error: {err}")
    if err.errno == mysql.connector.errorcode.ER_ACCESS_DENIED_ERROR:
        print("Access denied: Check your username and password.")
    elif err.errno == mysql.connector.errorcode.ER_BAD_DB_ERROR:
        print("Database does not exist.")
    else:
        print("MySQL Error:", err)
    exit()

# SQL query to retrieve data for the year 2023 and group by week
query = """
    SELECT *
    FROM your_table_name
    WHERE YEAR(action_time) = 2023
    ORDER BY action_time
"""

# Execute the query
try:
    cursor.execute(query)
    data = cursor.fetchall()
except mysql.connector.Error as err:
    print(f"Error: {err}")
    conn.close()
    exit()

# Create a Pandas DataFrame from the query result
df = pd.DataFrame(data, columns=[col[0] for col in cursor.description])

# Create a directory to store the Excel files if it doesn't exist
output_dir = "excel_exports"
if not os.path.exists(output_dir):
    os.mkdir(output_dir)

# Group the data by week and export to separate Excel files
for week_number, group in df.groupby(df['action_time'].dt.strftime('%U')):
    # Create a new Excel workbook
    wb = Workbook()
    ws = wb.active

    # Convert the DataFrame to rows and add them to the worksheet
    for r_idx, row in enumerate(dataframe_to_rows(group, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # Save the workbook to an Excel file
    file_name = f"{output_dir}/week_{week_number}_data.xlsx"
    wb.save(file_name)
    print(f"Excel file '{file_name}' exported for week {week_number}")

# Close the MySQL database connection
conn.close()
