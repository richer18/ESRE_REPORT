# ESRE Report Project Handoff

This file is the single handoff document for the LGU Treasurer / ESRE reporting work done on the Zamboanguita Firebird database.

## Project Purpose

Build a reporting foundation for LGU Treasurer collections using the legacy Firebird database:

```text
C:\ZAMBOANGUITA_DB\ZAMBOANGUITA.FDB
```

The work focuses on:

- Real Property Tax collection
- Community Tax Certificate / Cedula collection
- Other Fees and Charges
- Daily, monthly, quarterly, yearly collections
- Summary of Collection templates
- RPT record templates
- Advance RPT payment reporting
- RPT sharing summary
- Future web reporting app using Laravel, Python, ReactJS, MySQL, and REST API

The Firebird database should be treated as the official legacy source and should be read-only for reporting work unless a future production posting process is fully validated.

## Important Project Paths

Main project folder:

```text
C:\Users\LIFT-LAPTOP\OneDrive\Desktop\ESRE_REPORT
```

Firebird database:

```text
C:\ZAMBOANGUITA_DB\ZAMBOANGUITA.FDB
```

Firebird 2.5 client DLL expected by the script:

```text
C:\Program Files\Firebird\Firebird_2_5\bin\fbclient.dll
```

Main runner:

```text
run_collection_query.py
```

SQL query pack:

```text
firebird_metadata\collection_analysis_queries.sql
```

Metadata and analysis:

```text
firebird_metadata\summary.md
firebird_metadata\collection_process_analysis.md
firebird_metadata\report_list_revision.md
firebird_metadata\columns.csv
firebird_metadata\tables.csv
firebird_metadata\indexes.csv
firebird_metadata\foreign_keys.csv
firebird_metadata\triggers.csv
```

Excel templates:

```text
report_template\summary_of_collection_template.xlsx
report_template\summary_of_collection_template_no_rpt.xlsx
report_template\summary_of_collection_template_rpt.xlsx
report_template\RECORD OF REAL PROPERTY TAX COLLECTION.xlsx
report_template\RECORD OF REAL PROPERTY TAX COLLECTION - ADVANCE PAYMENT REPORT.xlsx
report_template\SUMMARY_REPORT_SHARING_TEMPLATE.xlsx
```

Generated report output folder:

```text
firebird_metadata\output
```

## How To Run

Open PowerShell or CMD in:

```text
C:\Users\LIFT-LAPTOP\OneDrive\Desktop\ESRE_REPORT
```

List available reports:

```powershell
python .\run_collection_query.py --list
```

Run a report:

```powershell
python .\run_collection_query.py 25 2026-01-01 2026-01-31
```

Run template reports:

```powershell
python .\run_collection_query.py 21 2024-12-01 2024-12-31
python .\run_collection_query.py 22 2024-12-01 2024-12-31
python .\run_collection_query.py 23 2024-12-01 2024-12-31
python .\run_collection_query.py 25 2026-01-01 2026-01-31
python .\run_collection_query.py 26 2026-01-01 2026-01-31
python .\run_collection_query.py 27 2026-01-01 2026-01-31
```

If Excel has the output file open, the script saves a timestamped copy instead of overwriting.

## Report List

Current report list:

```text
 1. Unified collection detail: RPT, CTC, and Other Fees/Charges
 2. Real Property Tax payment detail with RPT posting/account context
 3. RPT totals by Basic/SEF/other RPT line classification
 4. Community Tax Certificate / Cedula payment detail
 5. CTC totals; compares payment line totals to CTC computed/recorded totals
 6. Other Fees and Charges payment detail
 7. Other Fees and Charges totals by revenue code
 8. Daily collections for the three processes
 9. Monthly collections for the three processes
10. Quarterly collections for the three processes
11. Yearly collections for the three processes
12. Receipt range and header-vs-detail reconciliation
13. Other Fees and Charges tax/rate list
14. Other Fees and Charges parent-child hierarchy with rates
15. Total collection per collector: daily, monthly, and yearly
16. Fees collected by selected collector
17. Sources of Collections summary
18. Process flow: Real Property Tax payment
19. Process flow: Community Tax Certificate / Cedula payment
20. Process flow: Other Fees and Charges
21. Summary of Collection
22. Summary of Collection no rpt
23. Summary of Collection rpt
24. Summary in RPT based on the SUMMARY sheet layout
25. Record of Real Property Tax Collection
26. Record of Real Property Tax Collection - Advance Payment Report
27. Summary Report Sharing
```

Reports 1 to 20 mostly export CSV from SELECT-only SQL.

Reports 21 to 23 generate Excel files using uploaded Summary of Collection templates.

Report 24 exports CSV RPT detail summary.

Reports 25 to 27 generate Excel files using uploaded RPT templates.

## Firebird Tables Learned

Main shared cashiering tables:

```text
PAYMENT
PAYMENTDETAIL
PAYMENTCLASSDETAIL
PAYMENTCHEQUE
TAXPAYER
T_ITAXTYPE
T_FUNDTYPE
T_PAYMODE
T_STATUS
```

Important RPT tables:

```text
RPTASSESSMENT
RPTASSESSMENTDETAIL
PROPERTY
PROPERTYOWNER
PROPERTYOWNERDETAIL
RPTLIABILITY
TPACCOUNT
POSTINGJOURNAL
MANUALDEBIT
T_PROPERTYKIND
T_CLASSIFICATION
T_BARANGAY
```

Important CTC / Cedula tables:

```text
COMMUNITYTAXCERTIFICATE
CTCOTHERLGU
T_CTCRATE
TAXPAYER
```

Important Other Fees and Charges tables:

```text
PAYMENT
PAYMENTDETAIL
T_OTHERPAYMENTRATE
BUSINESSPERMIT
BUSINESSPERMITDETAIL
BUSINESSPERMITJOURNAL
BURIALPERMIT
MARRIAGELICENSE
CASHTICKET
WATERBILL
WATERBILLDETAIL
EEBILL
EECLEARANCE
MCHCHARGES
```

## Core Payment Logic

Official receipt header:

```text
PAYMENT
```

Official receipt number:

```text
PAYMENT.RECEIPTNO
```

Payment date:

```text
PAYMENT.PAYMENTDATE
```

Receipt amount:

```text
PAYMENT.AMOUNT
```

Receipt line amounts:

```text
PAYMENTDETAIL.AMOUNTPAID
```

RPT class / property / tax-year detail:

```text
PAYMENTCLASSDETAIL
```

Defensive filters used:

```sql
COALESCE(PAYMENT.VOID_BV, 0) = 0
COALESCE(PAYMENTCLASSDETAIL.CANCELLED_BV, 0) = 0
```

## Process Identification

RPT:

```sql
PAYMENT.PAYGROUP_CT = 'RPT'
```

CTC / Cedula:

```sql
PAYMENTDETAIL.SOURCE_CT IN ('CTCI', 'CTCC')
OR PAYMENTDETAIL.ITAXTYPE_CT = 'CTC'
```

Other Fees and Charges:

```sql
Not RPT
Not CTC
```

## RPT Basic / SEF Logic

Basic tax:

```text
PAYMENTCLASSDETAIL.ITAXTYPE_CT = 'BSC'
```

SEF:

```text
PAYMENTCLASSDETAIL.ITAXTYPE_CT = 'SEF'
```

Discount:

```text
PAYMENTCLASSDETAIL.CASETYPE_CT = 'DED'
```

Important: discounts are stored as negative amounts in Firebird. For Excel templates, write discount as positive:

```text
ABS(DED amount)
```

Penalty:

```text
PAYMENTCLASSDETAIL.CASETYPE_CT = 'PEN'
```

Regular tax/base amount:

```text
Other CASETYPE_CT values, usually REG
```

## Property Kind Mapping

From `T_PROPERTYKIND`:

```text
B = BUILDING
L = LAND
M = MACHINERIES
P = IMPRVMNTS
```

## Property Classification Mapping

From `T_CLASSIFICATION`:

```text
A    = AGRICULTURAL
C    = COMMERCIAL
I    = INDUSTRIAL
R    = RESIDENTIAL
S*   = SPECIAL
SS   = SCIENTIFIC, grouped under SPECIAL
SC   = CULTURAL, grouped under SPECIAL
SED  = SPECIAL EDUCATION, grouped under SPECIAL
SGOV = SPECIAL GOVERNMENT, grouped under SPECIAL
```

Important fix:

```text
Any CLASSCODE_CT starting with S is normalized to SPECIAL.
```

## Important RPT Bug Found And Fixed

Problem:

One `PAYMENT_ID` and one `TAXTRANS_ID` can contain multiple `PAYMENTCLASSDETAIL.CLASSCODE_CT` values.

Old grouping:

```text
PAYMENT_ID + TAXTRANS_ID
```

This could hide Scientific/SPECIAL amounts under the first classification found for the same TD/ARP.

Fixed grouping:

```text
PAYMENT_ID + TAXTRANS_ID + normalized property classification
```

This affects reports:

```text
25. Record of Real Property Tax Collection
26. Record of Real Property Tax Collection - Advance Payment Report
27. Summary Report Sharing
```

January 2026 example:

```text
Property Kind: BUILDING
Classification: SCIENTIFIC / SPECIAL
BSC Gross: 520.40
BSC Discount: 104.08
BSC Net: 416.32
SEF Gross: 520.40
SEF Discount: 104.08
SEF Net: 416.32
Grand Net: 832.64
```

After fix, January 2026 report 25 reconciles:

```text
Report total:   4,820,216.46
Firebird total: 4,820,216.46
Difference:     0.00
```

## Report 21: Summary of Collection

Uses template:

```text
summary_of_collection_template.xlsx
```

Combines:

```text
No-RPT collections
RPT collections
```

Generated as Excel.

## Report 22: Summary of Collection no RPT

Uses template:

```text
summary_of_collection_template_no_rpt.xlsx
```

Includes:

```text
Business taxes
Permits
Fees and charges
CTC
Other local revenue sources
```

Generated as Excel.

## Report 23: Summary of Collection RPT

Uses template:

```text
summary_of_collection_template_rpt.xlsx
```

Includes RPT Basic and SEF summary buckets.

Generated as Excel.

## Report 24: RPT Summary Layout Detail

Exports CSV.

Shows how RPT summary values are allocated by:

```text
Property group
Tax type
Line type
Current tax year used
Provincial share
Municipal share
Barangay share
```

## Report 25: Record of Real Property Tax Collection

Uses template:

```text
RECORD OF REAL PROPERTY TAX COLLECTION.xlsx
```

Data starts at row 11.

One output row is:

```text
PAYMENT_ID + TAXTRANS_ID + PROPERTY_CLASSIFICATION
```

Important columns:

```text
Date
Paid By
Name of Taxpayer
Period Covered
PIN
O.R. No.
TD/ARP No.
Barangay
Basic Current
Basic Discount
Basic Prior
Basic Penalty
Basic Gross
Basic Net
SEF Current
SEF Discount
SEF Prior
SEF Penalty
SEF Gross
SEF Net
Grand Gross
Grand Net
25% Share
Property Classification
Property Kind
Collector
```

## Report 26: Advance RPT Payment Report

Uses template:

```text
RECORD OF REAL PROPERTY TAX COLLECTION - ADVANCE PAYMENT REPORT.xlsx
```

Advance rule:

```text
PAYMENTCLASSDETAIL.TAXYEAR > report year
```

For December 2024:

```text
Advance year found: 2025
Advance rows: 804
Advance total: 960,046.92
```

For January 2026:

```text
Rows: 0
Reason: no 2027+ RPT advance rows found
```

## Report 27: Summary Report Sharing

Uses template:

```text
SUMMARY_REPORT_SHARING_TEMPLATE.xlsx
```

Fills BSC and SEF input cells by:

```text
Property kind
Property classification
Current/prior/penalty/discount bucket
```

The uploaded template had formula issues in SEF sharing sections. The script corrects these formulas in the generated workbook only. The original template is not modified.

Because the template has no separate building-special row, the generated workbook labels the building catch-all row as:

```text
BLDG-INDUS/SPECIAL
```

January 2026 report 27 verified:

```text
BLDG-INDUS/SPECIAL BSC Gross: 520.40
BLDG-INDUS/SPECIAL BSC Discount: 104.08
BLDG-INDUS/SPECIAL SEF Gross: 520.40
BLDG-INDUS/SPECIAL SEF Discount: 104.08
BSC net from inputs: 2,410,108.23
SEF net from inputs: 2,410,108.23
Grand net from inputs: 4,820,216.46
```

## Sources Of Collections Used In Report 17

The requested source list includes:

```text
Manufacturing
Distributor
Retailing
Banks & Other Financial Int.
Other Business Tax
Sand & Gravel
Fines & Penalties
Mayor's Permit
Weight & Measure
Tricycle Permit Fee
Occupation Tax
Cert. of Ownership
Cert. of Transfer
Cockpit Share
Docking and Mooring Fee
Sultadas
Miscellaneous
Registration of Birth
Marriage Fee
Burial Fee
Correction of Entry
Fishing Permit Fee
Sale of Agri. Prod.
Sale of Acct. Forms
Water Fee
Market Stall Fee
Cash Tickets
SlaughterHouse Fee
Rental of Equipment
Doc Stamp Tax
Police Clearance
Secretaries Fees
Med./Lab. Fees
Garbage Fees
Cutting Tree
Com Tax Cert.
Building Permit Fee
Electrical Permit Fee
Zoning Fee
Livestock
Diving Fee
```

## Parent / Child Revenue Code Understanding

Other Fees and Charges can have a hierarchy:

```text
Parent
  Child 1
    Child 2
```

Report 14 investigates this through `T_OTHERPAYMENTRATE` and related rate/source code relationships.

## Collector Reports

Reports 15 and 16 focus on collectors:

```text
15. Total collection per collector: daily, monthly, yearly
16. Fees collected by selected collector
```

Run selected collector report like:

```powershell
python .\run_collection_query.py 16 2026-01-01 2026-01-31 --collector GTZ
```

## Future Web App Architecture

Recommended future architecture:

```text
ReactJS Frontend
   ↓ REST API
Laravel Backend
   ↓
MySQL reporting database
   ↓
Python Firebird reader / importer
   ↓
ZAMBOANGUITA.FDB
```

Recommended first principle:

```text
Firebird .FDB = official legacy source
MySQL = reporting/import database
```

Start read-only:

```text
Firebird .FDB -> Python importer -> MySQL -> Laravel API -> React reports
```

Do not write back to `.FDB` in early phases.

## Future Import To MySQL

Imported data should be saved in MySQL.

Suggested MySQL tables:

```text
import_batches
report_payments
report_payment_lines
rpt_collection_records
rpt_advance_payment_records
rpt_sharing_summaries
ctc_collections
other_fee_collections
collector_summaries
report_snapshots
sync_logs
```

Always preserve Firebird source IDs:

```text
PAYMENT_ID
PAYMENTDETAIL_ID
PAYCLASSDETAIL_ID
TAXTRANS_ID
RECEIPTNO
```

These allow MySQL rows to trace back to the original Firebird records.

## Suggested Future Web Reporting Modules

```text
Dashboard
Daily Collection Report
Monthly Collection Report
Quarterly Collection Report
Yearly Collection Report
Summary of Collection
Summary of Collection no RPT
Summary of Collection RPT
Record of Real Property Tax Collection
Advance RPT Payment Report
Summary Report Sharing
CTC / Cedula Report
Other Fees and Charges Report
Collector Report
Excel Template Export
Import Batch Monitor
Sync Logs
Audit Logs
```

## GitHub Repository

GitHub repository:

```text
https://github.com/richer18/ESRE_REPORT
```

Repository description:

```text
LGU Treasurer ESRE reporting toolkit for Zamboanguita collections, integrating Firebird .FDB analysis, RPT/CTC/Other Fees reports, Excel template exports, and future MySQL/Laravel/React reporting migration.
```

Important Git finding:

```text
C:\Users\LIFT-LAPTOP has a .git folder.
```

That means normal `git status` from inside `ESRE_REPORT` can accidentally see the whole Windows user directory. To avoid uploading unrelated personal files, the GitHub setup script creates/uses a separate `.git` inside:

```text
C:\Users\LIFT-LAPTOP\OneDrive\Desktop\ESRE_REPORT
```

GitHub helper files:

```text
git_hub_runner\00_setup_github_repo.bat
git_hub_runner\github_menu.bat
git_hub_runner\git_status_local.bat
git_hub_runner\update_from_github.bat
git_hub_runner\update_github.bat
.gitignore
```

Recommended first-time setup:

```powershell
cd C:\Users\LIFT-LAPTOP\OneDrive\Desktop\ESRE_REPORT
.\git_hub_runner\00_setup_github_repo.bat
```

Menu runner:

```powershell
.\git_hub_runner\github_menu.bat
```

Update local files from GitHub:

```powershell
.\git_hub_runner\update_from_github.bat
```

Push local changes to GitHub:

```powershell
.\git_hub_runner\update_github.bat
```

Push with a custom commit message:

```powershell
.\git_hub_runner\update_github.bat "Update RPT report templates"
```

Check local repo status:

```powershell
.\git_hub_runner\git_status_local.bat
```

Files intentionally ignored by `.gitignore`:

```text
*.FDB / *.fdb
firebird_metadata/output/
google_sheet_exports/
__pycache__/
.env files
```

Reason: generated reports, Google exports, and Firebird databases can contain taxpayer/payment data and should not be uploaded unless intentionally reviewed.

Current local folder organization:

```text
README\PROJECT_HANDOFF_README.md
git_hub_runner\*.bat
work_flow\source/reference files
run_collection_query.py
firebird_metadata\
report_template\
```

## Move To Another PC Checklist

Copy this folder:

```text
C:\Users\LIFT-LAPTOP\OneDrive\Desktop\ESRE_REPORT
```

Copy or make accessible:

```text
C:\ZAMBOANGUITA_DB\ZAMBOANGUITA.FDB
```

Install Firebird 2.5 client or make sure this exists:

```text
C:\Program Files\Firebird\Firebird_2_5\bin\fbclient.dll
```

Install Python and required package:

```powershell
pip install fdb openpyxl
```

Test:

```powershell
cd C:\Users\YOUR_USER\Desktop\ESRE_REPORT
python .\run_collection_query.py --list
python .\run_collection_query.py 25 2026-01-01 2026-01-31
```

If the database path changes, update `DB_PATH` in:

```text
run_collection_query.py
```

If Firebird client path changes, update `FB_CLIENT` in:

```text
run_collection_query.py
```

## Safety Rule

Current scripts are intended for reporting.

Use read-only logic against Firebird.

Do not modify:

```text
PAYMENT
PAYMENTDETAIL
PAYMENTCLASSDETAIL
RPTASSESSMENT
PROPERTY
```

unless a future official posting design is fully validated and approved.
