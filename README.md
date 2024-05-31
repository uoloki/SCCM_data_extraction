# SCCM Data Extraction and Excel File Filtering Script Documentation

## Overview
This documentation provides a comprehensive guide on how to use two Python scripts: one for extracting data from Microsoft System Center Configuration Manager (SCCM) and another for filtering the extracted data based on specific criteria. The scripts are designed to be user-friendly and include detailed instructions on setup, configuration, and execution. This guide is intended for users who may not have extensive experience with scripting or database management.

## Prerequisites

### Software Requirements
- **Python 3.6 or higher**
- **Necessary Python libraries:**
  - `pyodbc`
  - `pandas`
  - `openpyxl`
  - `logging`

You can install the required libraries using the following command in your terminal or command prompt:


```sh
pip install pyodbc pandas openpyxl
```

### System Requirements
- Access to a machine with Python installed.
- Access to the SCCM SQL Server database.
- Necessary permissions to read from the SCCM database.

## Initial Setup

### Creating the `credentials.txt` File
The extraction script requires a `credentials.txt` file that contains the database connection credentials. This file should be placed in the same directory as the script.

1. **Open a text editor** (such as Notepad on Windows or any other plain text editor).
2. **Enter your database credentials** in the following format:
   
   ```makefile
   driver={your-driver}
   server={your-server}
   database={your-database}
   username={your-username}
   password={your-password}
   ```

   Replace the placeholders (e.g., `{your-driver}`) with your actual database connection details. For example:

  ```makefile
  driver=ODBC Driver 17 for SQL Server
  server=192.168.1.100
  database=SCCM_DB
  username=admin
  password=yourpassword
  ```


3. **Save the file** as `credentials.txt` in the same directory as your script.

### Finding Your Database Credentials
- **Driver:** This is typically the type of database you are connecting to. For SCCM, it is usually `SQL Server`.
- **Server:** This is the IP address or hostname of your SQL Server. You can find this in your SCCM management console or by asking your network administrator.
- **Database:** This is the name of the database where SCCM stores its data. The default name is usually something like `CM_<SiteCode>` (e.g., `CM_ABC`).
- **Username and Password:** These are the credentials you use to access the SQL Server database. Ensure you have read permissions on the SCCM database.

## Script Configuration

### Customizing SQL Queries
The extraction script includes predefined SQL queries to extract specific data from the SCCM database. These queries are defined in the `queries` dictionary within the script. You can modify these queries or add new ones based on your requirements.

#### Example Queries

1. **Hardware Inventory Query:**

   ```sql
   SELECT
       v_GS_COMPUTER_SYSTEM.Name0 AS ComputerName,
       v_GS_PROCESSOR.Name0 AS ProcessorName,
       v_GS_PROCESSOR.NumberOfCores0 AS NumberOfCores,
       v_GS_X86_PC_MEMORY.TotalPhysicalMemory0 AS TotalPhysicalMemory
   FROM
       v_GS_COMPUTER_SYSTEM
   JOIN
       v_GS_PROCESSOR ON v_GS_COMPUTER_SYSTEM.ResourceID = v_GS_PROCESSOR.ResourceID
   JOIN
       v_GS_X86_PC_MEMORY ON v_GS_COMPUTER_SYSTEM.ResourceID = v_GS_X86_PC_MEMORY.ResourceID
   ```

2. **Software Inventory Query:**

   ```sql
   SELECT
       v_GS_ADD_REMOVE_PROGRAMS.DisplayName0 AS SoftwareName,
       v_GS_ADD_REMOVE_PROGRAMS.Version0 AS Version,
       v_GS_ADD_REMOVE_PROGRAMS.Publisher0 AS Publisher,
       v_GS_COMPUTER_SYSTEM.Name0 AS ComputerName
   FROM
       v_GS_ADD_REMOVE_PROGRAMS
   JOIN
       v_GS_COMPUTER_SYSTEM ON v_GS_ADD_REMOVE_PROGRAMS.ResourceID = v_GS_COMPUTER_SYSTEM.ResourceID
   ```

3. **Backup Status Query:**

   ```sql
   SELECT
       v_GS_BACKUPSTATUS.BackupDateTime0 AS BackupDateTime,
       v_GS_BACKUPSTATUS.BackupStatus0 AS BackupStatus,
       v_GS_COMPUTER_SYSTEM.Name0 AS ComputerName
   FROM
       v_GS_BACKUPSTATUS
   JOIN
       v_GS_COMPUTER_SYSTEM ON v_GS_BACKUPSTATUS.ResourceID = v_GS_COMPUTER_SYSTEM.ResourceID
   ```

### Adding or Modifying Queries
To add or modify queries, update the `queries` dictionary in the script. Use the following format:

```python
queries = {
    'QueryName': """
        YOUR SQL QUERY HERE
    """
}
```

## Running the Extraction Script

### Execution Steps
1. **Ensure that the `credentials.txt` file is in the same directory as the script.**
2. **Open a terminal or command prompt.**
3. **Navigate to the directory containing the script:**

   ```sh
   cd path/to/your/script
   ```

4. **Run the script:**

   ```sh
   python sccm_data_extraction.py
   ```

### Understanding the Output
The script will execute the specified SQL queries and save the results to an Excel file named `sccm_data.xlsx`. Each query's results will be saved in a separate sheet within the Excel file. The column widths will be adjusted automatically to fit the content.

### Log Files
The script generates log messages to provide detailed information about the execution process, including successful operations and any errors encountered. 

## Filtering the Excel File

### Preparing the Input File
Ensure you have an Excel file (`sccm_data.xlsx`) that contains the data you want to filter. This file should have columns with the `_Y` suffix indicating which rows to keep based on the value 'Y'.

### Running the Filtering Script

1. **Place the filtering script in the same directory as your input Excel file.**

2. **Open a terminal or command prompt.**

3. **Navigate to the directory containing the script and the input file:**

   ```sh
   cd path/to/your/script
   ```

4. **Run the script:**

   ```sh
   python excel_filtering.py
   ```

### Understanding the Output
The script will read the specified input Excel file, filter the data based on the criteria in the `_Y` columns, and save the filtered data to a new Excel file named `filtered_sccm_data.xlsx`. The column widths in the new file will be adjusted automatically to fit the content.

## Troubleshooting

### Common Issues

1. **Invalid Credentials:**
   - Ensure that the `credentials.txt` file contains the correct database connection details.
   - Verify that you have the necessary permissions to access the SCCM database.

2. **File Not Found:**
   - Ensure the input file (`sccm_data.xlsx`) is in the same directory as the script.
   - Verify the file name and extension are correct.

3. **SQL Query Errors:**
   - Check the syntax of your SQL queries.
   - Ensure that the SCCM database tables and columns referenced in the queries exist.

4. **Library Installation Issues:**
   - Ensure that you have installed all the required Python libraries using pip.
   - If you encounter issues with `pyodbc`, make sure you have the appropriate ODBC driver installed for your database.

## Additional Resources

- **SCCM Documentation:** Refer to the official Microsoft documentation for more information on SCCM.
- **SQL Query Reference:** Use SQL reference guides to help formulate and debug your queries.
- **Python Library Documentation:** Refer to the documentation for `pyodbc`, `pandas`, and `openpyxl` for more details on using these libraries.

