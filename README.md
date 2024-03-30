# Data Transfer: Excel to Database

This C# application provides functionality to transfer data between an Excel file and a SQL Server database. It allows users to export data from a database to an Excel spreadsheet and import data from an Excel spreadsheet to a database table.

## Features

- **Export Data to Excel:** Retrieve data from a SQL Server database and export it to an Excel file.
- **Import Data from Excel:** Read data from an Excel file and insert it into a SQL Server database table.

## Usage

1. **Export Data to Excel:**

   - Click on the "Database to Excel" button to export data from the database to an Excel spreadsheet.
   - The application will create a new Excel file and populate it with data from the specified database table.
   - Each row in the Excel file represents a record from the database table.
   - The Excel file will be opened automatically after it is generated.

2. **Import Data from Excel:**
   - Click on the "Excel to Database" button to import data from an Excel spreadsheet to the database.
   - Select an Excel file containing the data to be imported.
   - The application will read the data from the Excel file and insert it into the specified database table.
   - Each row in the Excel file should correspond to a record in the database table.
   - After the import process is completed, a success message will be displayed.

## Error Handling

- If an error occurs during the data transfer process, an error message will be displayed to the user.
- Error messages include details about the type of error and suggestions for resolution.

## Credits

This project created as part of [C# Projeleri : Uygulama Geliştirerek C# Öğrenin](https://www.udemy.com/course/csharprojeleri "https://www.udemy.com/course/csharprojeleri") Udemy Course

---
