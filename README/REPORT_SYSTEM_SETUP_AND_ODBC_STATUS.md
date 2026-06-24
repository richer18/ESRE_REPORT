# ESRE Report System Setup and ODBC Status

Date updated: 2026-06-24

This note records the changes made to `E:\ESRE_REPORT` for easier setup on this PC and future client PCs.

## Main Script Updated

File updated:

```text
E:\ESRE_REPORT\run_collection_query.py
```

Added connection status messages so the script clearly shows whether it connected or failed.

Example command:

```powershell
python run_collection_query.py --test-connection
```

Example status output:

```text
STATUS: Connecting to Firebird database: main-server:i_tax046zamboanguita
STATUS: Firebird client: C:\Program Files\Firebird\Firebird_2_5\bin\fbclient.dll
STATUS: Database connection failed: ...
STATUS: Connecting to Firebird database: E:\ZAMBOANGUITA.FDB
STATUS: Database connected successfully.
STATUS: Connection test finished.
```

## ODBC Mode Added

The script now supports this option:

```powershell
python run_collection_query.py --connection odbc --test-connection
```

ODBC settings used:

```text
DSN: itaxzamboanguita
Firebird client override: C:\Program Files\Firebird\Firebird_2_5\bin\fbclient.dll
```

Reason for the client override:

```text
The DSN was trying to load C:\iTAX\utilities\fbclient.dll.
That DLL is likely 32-bit and fails with 64-bit Python.
The script now overrides ODBC to use the 64-bit Firebird client DLL.
```

Verified working ODBC status:

```text
STATUS: Connecting through ODBC DSN: itaxzamboanguita
STATUS: ODBC Firebird client: C:\Program Files\Firebird\Firebird_2_5\bin\fbclient.dll
STATUS: ODBC connected successfully.
STATUS: Connection test finished.
```

Also verified ODBC SQL test:

```sql
SELECT 1 FROM RDB$DATABASE
```

Result:

```text
1
```

## Connection Modes

Native Firebird mode:

```powershell
python run_collection_query.py --connection native --test-connection
```

ODBC mode:

```powershell
python run_collection_query.py --connection odbc --test-connection
```

Auto mode:

```powershell
python run_collection_query.py --connection auto --test-connection
```

Report example using ODBC:

```powershell
python run_collection_query.py --connection odbc 25 2026-01-01 2026-01-31
```

## Auto Python Package Install Added

The script now checks and installs missing Python packages automatically:

```text
fdb
openpyxl
pyodbc
```

If one package is missing, the script runs:

```powershell
python -m pip install PACKAGE_NAME
```

Important:

```text
This auto-installs Python packages only.
It does not create the ODBC DSN.
It does not install the Firebird ODBC driver.
It does not install the Firebird client DLL.
```

## Installer / Checker Added

New file:

```text
E:\ESRE_REPORT\install_report_system.bat
```

New requirements file:

```text
E:\ESRE_REPORT\requirements.txt
```

Requirements:

```text
fdb
openpyxl
pyodbc==5.3.0
```

Run this on a PC to prepare/check the report system:

```powershell
E:\ESRE_REPORT\install_report_system.bat
```

The installer/checker does this:

```text
1. Checks if Python is installed.
2. If Python is missing, tries to install Python using winget.
3. Installs Python packages from requirements.txt.
4. Checks the 64-bit Firebird client DLL.
5. Checks for ODBC DSN itaxzamboanguita.
6. Runs ODBC connection test.
```

Verified result on this PC:

```text
SUCCESS: ESRE report system dependencies and ODBC connection are ready.
```

## What Must Exist On Each PC

For ODBC reports to work on a client PC:

```text
1. Python installed and available in PATH.
2. Python packages installed: fdb, openpyxl, pyodbc==5.3.0.
3. 64-bit ODBC DSN named itaxzamboanguita.
4. Firebird/InterBase ODBC driver installed.
5. 64-bit Firebird client DLL at:
   C:\Program Files\Firebird\Firebird_2_5\bin\fbclient.dll
6. Network access to main-server.
```

## Recommended Future Web App Setup

For the future web app, install these database dependencies only on the backend/server PC:

```text
React frontend users should not connect directly to Firebird or ODBC.
Browser users should connect to the Laravel/Python API.
The backend/server should connect to Firebird/ODBC.
```

Recommended flow:

```text
User Browser
  -> React Frontend
    -> Backend API
      -> ODBC DSN itaxzamboanguita or native Firebird
        -> main-server / Firebird database
```

## Default Connection Mode

As of 2026-06-24, the default connection mode is now ODBC.

This means normal commands now use the working ODBC DSN automatically:

```powershell
python run_collection_query.py 34 2026-02-01 2026-02-28
```

Equivalent explicit command:

```powershell
python run_collection_query.py --connection odbc 34 2026-02-01 2026-02-28
```

For report 34, if no collector is provided, the script lists collectors and asks for a collector number.

Tested report 34 through ODBC:

```text
python run_collection_query.py 34 2026-02-01 2026-02-28 1
Selected collector: agnes
Rows exported: 1734
Output file: E:\ESRE_REPORT\firebird_metadata\output\query_34_agnes_2026-02-01_to_2026-02-28.xlsx
```
