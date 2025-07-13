Certainly! Below is a sample README file for your project that describes the functionality of the Excel SQL Query Tool without the recent CSV support. You can customize it further based on your preferences.

---

# GTS Excel SQL Query Tool

## Overview

The GTS Excel SQL Query Tool is a graphical user interface (GUI) application built using Python's Tkinter library. This tool allows users to load Excel files, execute SQL queries on the data, and export the results back to Excel format. It provides a user-friendly way to interact with Excel data using SQL syntax.

## Features

- **Load Excel Files**: Browse and load multiple Excel files (.xlsx, .xls) into an in-memory SQLite database.
- **View Available Tables**: Display a hierarchical view of the loaded tables and their respective sheets.
- **Execute SQL Queries**: Write and execute SQL queries against the loaded data.
- **View Query Results**: Display the results of executed queries in a treeview format.
- **Export Results**: Export query results to an Excel file.
- **Query History**: Keep track of previously executed queries for easy access.

## Requirements

- Python 3.x
- Tkinter (comes pre-installed with Python)
- Pandas
- Openpyxl
- SQLite3 (comes pre-installed with Python)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/GTS_Excel_SQL_Query_Tool.git
   cd GTS_Excel_SQL_Query_Tool
   ```

2. Install the required packages:
   ```bash
   pip install pandas openpyxl
   ```

## Usage

1. Run the application:
   ```bash
   python Excel_SQL_Developer.py
   ```

2. Use the "Browse Excel Files" button to select a directory containing your Excel files.

3. Once the files are loaded, you can view the available tables in the left panel.

4. Write your SQL queries in the query input area and click "Execute" to run the query.

5. The results will be displayed in the results panel. You can export the results to an Excel file using the "Export to Excel" button.

6. You can also view the history of executed queries and clear the results as needed.

## Example Queries

- To select all data from a specific sheet:
  ```sql
  SELECT * FROM "Sheet1"
  ```

- To filter data based on a condition:
  ```sql
  SELECT * FROM "Sheet1" WHERE "Column1" = 'Value'
  ```

## Limitations

- The tool currently supports only Excel files (.xlsx, .xls).
- It does not support CSV files or other data formats.

## Acknowledgments

- Thanks to the developers of the Pandas and Tkinter libraries for making this project possible.
