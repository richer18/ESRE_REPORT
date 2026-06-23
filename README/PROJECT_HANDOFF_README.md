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
28. Provincial RPT Coding / Province Remittance Report
29. Abstract of General Collections
30. Abstract of Trust Funds Collections
31. Full Report Collections
32. CMCI Annex A-B Business Permit Registration Report
33. Tax on Business Summary from BPLS Business Tax
34. Generate Receipt Collector
```

Reports 1 to 20 mostly export CSV from SELECT-only SQL.

Reports 21 to 23 generate Excel files using uploaded Summary of Collection templates.

Report 24 exports CSV RPT detail summary.

Reports 25 to 27 generate Excel files using uploaded RPT templates.

Reports 28 to 30 generate Excel files using uploaded templates.

Report 31 generates an Excel file using the uploaded Full Report Collections template. `DUE FROM` is manual and remains editable in the generated workbook.

Report 32 generates an Excel file using the uploaded CMCI Annex A-B template.

Report 33 generates a Tax on Business Excel summary from BPLS Business Tax and Surcharge.

Report 34 generates an Excel collector receipt audit list from Firebird `PAYMENT`, including RPT, Community Tax Certificate, and Other Fees/Charges receipt headers. It intentionally includes paid, cancelled, and void payments for collector review.

Example:

```powershell
python .\run_collection_query.py 34 2026-01-01 2026-01-31 angelique
```

Collector selection can also be number-based. Run report 34 without a collector and the script will print collectors for the selected period, then ask for a number:

```powershell
python .\run_collection_query.py 34 2026-01-01 2026-01-31
```

You can also pass the collector number directly:

```powershell
python .\run_collection_query.py 34 2026-01-01 2026-01-31 1
```

Report 34 output columns:

```text
DATE | Collector | Receipt Type | Receipt No | Taxpayer name | Status | Total
```

`python .\run_collection_query.py --list` includes reports 1 to 34.

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

RCD / Report of Collections and Deposits finding:

- The database has RCD-related structures:
  - `RCDCTRLNUMBER`
  - `RCDACCOUNTABLEFORM`
  - `RCDCASHBREAKDOWN`
  - `RCDFUNDTYPEBREAKDOWN`
  - `RCDDEPOSIT`
  - `RCDORBOOKLET`
  - `RCDORBOOKLETDEFECTIVE`
  - `RCDORBOOKLETINVENTORY`
  - `REMITTANCE`
  - `SECURITY_USERS_RCD`
- Live read-only check found `0` rows in the main RCD transaction/control tables listed above, except `SECURITY_USERS_RCD` which has 3 users: `rowena`, `ricardo`, and `joy`.
- `PAYMENT` has an `RCDNUMBER` column, but direct read-only scan found `0` nonblank `PAYMENT.RCDNUMBER` values in the current database.
- Meaning: the schema supports RCD / accountable form / cash breakdown / deposit workflows, but current collection records are not linked to generated RCD numbers in the `.FDB`.
- Remittance/deposit date fields exist in schema, but live read-only check found no usable remittance rows:
  - `REMITTANCE.TRANSDATE`, `AMOUNTREMITTED`, `CASHREMITTED`, `CHEQUEREMITTED`, `COLLECTOR`, `REMITTE`, `REFRCDNUM`
  - `RCDDEPOSIT.DEPOSITDATE`, `AMOUNT`, `DEPOSIT_SLIP`, `BANKCODE_CT`, `RCDCTRL_ID`
  - `RCDCTRLNUMBER.AOAPPROVEDDATE`, `LOAPPROVEDDATE`, `COAPPROVEDDATE`, `REMITEDAMOUNT`
  - `BANKSTATEMENT` and `BANKSTATEMENTLINE` tables exist but are empty in the current database
- Live row counts: `REMITTANCE = 0`, `RCDDEPOSIT = 0`, `RCDCTRLNUMBER = 0`, `BANKSTATEMENT = 0`, `BANKSTATEMENTLINE = 0`.
- Meaning: current `.FDB` cannot tell when a specific receipt was remitted/deposited from actual remittance tables. It can only show payment date/transaction date from `PAYMENT` unless another external system/file records remittance.

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

## Report 28: Provincial RPT Coding / Province Remittance Report

Analysis file:

```text
firebird_metadata\provincial_rpt_coding_template_analysis.md
```

Uploaded template:

```text
report_template\PROVINCIAL_RPT_CODING_TEMPLATE.xlsx
```

Workbook sheets:

```text
GF
SEF
```

Purpose understood:

```text
Monthly report on the collection of Real Property Tax by property classification.
Likely used as support for remittance/coding to the Provincial Government.
```

Template layout:

```text
Rows 9-21  Property classification lines
C/E/G/I    Amount columns
B/D/F/H    Account/coding columns
Row 23     Subtotals
Row 25     Total remittance
```

Likely Firebird source:

```text
PAYMENT
PAYMENTCLASSDETAIL
RPTASSESSMENT
PROPERTY
T_CLASSIFICATION
T_ACTUALUSE
T_PROPERTYKIND
```

Likely split:

```text
GF sheet  = PAYMENTCLASSDETAIL.ITAXTYPE_CT = 'BSC'
SEF sheet = PAYMENTCLASSDETAIL.ITAXTYPE_CT = 'SEF'
```

Report 28 must be based on Report 27 sharing logic:

```text
Basic/GF province share = 35 percent
SEF province share      = 50 percent
```

Reason:

```text
27. Summary Report Sharing already contains the sharing formulas.
Report 28 should use the same Firebird buckets and sharing rules, then place the Provincial share into the province coding/remittance template.
```

Required RPT reconciliation chain:

```text
21. Summary of Collection
23. Summary of Collection RPT
25. Record of Real Property Tax Collection
27. Summary Report Sharing
28. Provincial RPT Coding / Province Remittance Report
```

Control meaning:

```text
Report 21 RPT lines should agree with Report 23 RPT totals.
Report 23 RPT totals should agree with Report 25 detailed RPT totals.
Report 25 Basic/SEF totals should agree with the raw BSC/SEF buckets used in Report 27.
Report 27 Provincial Share totals should agree with Report 28 GF and SEF remittance/coding totals.
```

Report 28 is implemented in `run_collection_query.py`, but these mapping issues should still be confirmed:

```text
1. Where SPECIAL/SCIENTIFIC/S* classifications should be placed.
2. Whether "Machinery" in the template means class M/Mineral or property kind M/Machineries.
3. How to handle property kind P/Improvements.
4. The Land-TIMBER row has no account/coding values in the uploaded template.
5. Report 28 should reconcile to Report 27 Provincial Share totals before printing/submission.
```

January 2026 implementation test:

```text
Command: python .\run_collection_query.py 28 2026-01-01 2026-01-31
Output: firebird_metadata\output\query_28_2026-01-01_to_2026-01-31.xlsx
GF/BSC provincial share total: 843,537.880500
SEF provincial share total: 1,205,054.11500
These match the expected provincial shares from Report 27/Report 25 basis:
- BSC net 2,410,108.23 x 35 percent
- SEF net 2,410,108.23 x 50 percent
```

## Reports 29-30: Abstract Collection Templates

Analysis file:

```text
firebird_metadata\abstract_collection_templates_analysis.md
```

Uploaded templates:

```text
report_template\ABSTRACT_OF_GENERAL_COLLECTIONS.xlsx
report_template\ABSTRACT_OF_TRUST_FUNDS_COLLECTIONS.xlsx
```

Both templates contain:

```text
data
daily_collection
```

Report 29: Abstract of General Collections

```text
Receipt-level source columns for General Fund/non-RPT collections.
Uses PAYMENT + PAYMENTDETAIL + T_ITAXTYPE + T_FUNDTYPE.
Should reuse classify_summary_source() mapping from run_collection_query.py.
```

Report 30: Abstract of Trust Funds Collections

```text
Receipt-level trust/shared columns for Building, Electrical, Zoning, Livestock, and Diving.
Uses PAYMENT + PAYMENTDETAIL + T_ITAXTYPE + T_FUNDTYPE.
Should reuse split_summary_amount() rules where applicable.
```

Trust split rules found in template:

```text
Building Fee: 80 percent Local, 15 percent Trust Fund, 5 percent National
Livestock: 80 percent Local, 20 percent National
Diving Fee: 40 percent GF, 30 percent Fishers, 30 percent Barangay
Electrical Fee: one amount column
Zoning Fee: one amount column
```

Reconciliation chain:

```text
21. Summary of Collection
22. Summary of Collection no rpt
17. Sources of Collections summary
29. Abstract of General Collections
30. Abstract of Trust Funds Collections
```

Open mapping items:

```text
1. Confirm Cash Tickets source/code rule.
2. Confirm whether CTC is excluded from the General Collections abstract.
3. Confirm whether Electrical and Zoning should remain unsplit.
4. Confirm whether Building Fee should include PFB, BUF, and INS.
5. Confirm if daily_collection should be formula-driven from data or generated directly from Firebird.
```

January 2026 implementation test:

```text
Report 29 command: python .\run_collection_query.py 29 2026-01-01 2026-01-31
Output: firebird_metadata\output\query_29_2026-01-01_to_2026-01-31.xlsx
Receipt-level rows: 3,882
Daily rows: 23
Data total: 4,898,687.21
Daily total: 4,898,687.21
Difference: 0.00

Report 30 command: python .\run_collection_query.py 30 2026-01-01 2026-01-31
Output: firebird_metadata\output\query_30_2026-01-01_to_2026-01-31.xlsx
Receipt-level rows: 836
Daily rows: 23
Data total: 632,496.20
Daily total: 632,496.20
Difference: 0.00
```

## Report 31: Full Report Collections

Analysis file:

```text
firebird_metadata\full_report_collections_template_analysis.md
```

Uploaded template:

```text
report_template\FULL_REPORT_COLLECTIONS.xlsx
```

Workbook structure:

```text
Sheet1 only
Rows 8-31 are daily rows
Row 32 contains monthly totals
Rows 36-38 compute:
RCD TOTAL - DUE FROM = TOTAL COLLECTIONS
```

Observed columns:

```text
A = DATE
B = CTC
C = RPT
D = GF AND TF
E = DUE FROM
F = RCD TOTAL
```

Recommended reconciliation chain:

```text
CTC column should reconcile to Community Tax / Cedula daily totals.
RPT column should reconcile to Reports 23, 25, and 27.
GF AND TF column should reconcile to Reports 22, 29, and 30.
RCD TOTAL should reconcile to cashier/RCD totals.
TOTAL COLLECTIONS should equal RCD TOTAL less DUE FROM.
```

Implementation rule:

```text
DUE FROM is manual.
Do not compute it from Firebird.
Generated workbook should leave daily DUE FROM cells blank or zero for manual entry.
TOTAL COLLECTIONS should update from the template formula after manual entry.
```

January 2026 implementation test:

```text
Command: python .\run_collection_query.py 31 2026-01-01 2026-01-31
Output: firebird_metadata\output\query_31_2026-01-01_to_2026-01-31.xlsx
Rows exported: 24
CTC total: 198,579.28
RPT total: 4,820,216.46
GF AND TF total: 5,531,183.41
RCD TOTAL before manual DUE FROM: 10,549,979.15
DUE FROM rows blank/manual: 24
GF AND TF reconciles to Report 29 + Report 30:
4,898,687.21 + 632,496.20 = 5,531,183.41
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

## SRE / LIFT Report Understanding

New reference files reviewed:

```text
C:\Users\LIFT-LAPTOP\Downloads\SRE_Zamboanguita_Q1_2026.xls
C:\Users\LIFT-LAPTOP\Downloads\SRS_Zamboanguita_Q1_2026.xls
C:\Users\LIFT-LAPTOP\Downloads\SOE_Zamboanguita_Q4_2025 (3).xls
D:\eSRE Retooling (Roll-out of LIFT System V4)\1. eSRE v4_Training Manual_031623 - Final.pptx
D:\eSRE Retooling (Roll-out of LIFT System V4)\2. JMC of DepEd, DBM, and DILG on the Revised Guideline  on the Use of the SEF.pdf
D:\eSRE Retooling (Roll-out of LIFT System V4)\BLGF - MC No. 001.2023 - Extension of Deadline for the Submission of the Q1 FY2023 LIFT Reports.pdf
D:\eSRE Retooling (Roll-out of LIFT System V4)\LIFT - Special Health Fund PPT Presentation.pptx
```

The file `SRE_Zamboanguita_Q1_2026.xls` is a LIFT/eSRE Statement of Receipts and Expenditures report:

```text
Report: Statement of Receipts and Expenditures
LGU: Zamboanguita, Negros Oriental
Period Covered: Q1, 2026
Sheet: sre_report
Rows: 75
Columns: 11
Generated: 21/05/2026 11:33 AM
```

Key Q1 2026 SRE values found:

```text
Local Sources total: 12,653,372.74
Tax Revenue total: 7,356,727.51
Real Property Tax total: 2,881,141.00
Real Property Tax - General Fund: 1,280,507.11
Real Property Tax - SEF: 1,600,633.89
Tax on Business total: 4,110,483.16
Other Taxes total: 365,103.35
Non-Tax Revenue total: 5,296,645.23
External Sources / National Tax Allotment: 47,498,793.00
Total Current Operating Income: 60,152,165.74
Current Operating Expenditures: 33,394,643.52
Net Operating Income from Current Operations: 26,757,522.22
Fund/Cash Balance End: 53,793,414.30
```

Important relationship to the Firebird reports:

```text
Report 25 shows gross RPT collection detail.
SRE shows LGU fund reporting amounts.
```

For a municipality, Basic RPT is shared among province, municipality, and barangay. SEF is shared between province and municipality. Therefore, SRE RPT values should be compared to the municipal share/fund columns, not directly to gross RPT detail totals.

Training manual findings:

```text
LIFT LGU System lets treasurers, budget officers, and assessors encode data for SRE and QRRPA reports.
Treasurer modules include income targets, RPT receipts, general collection receipts, trust fund receipts, expenditures, fund/cash balance, financial operations, and SRE-NGAS reconciliation.
RPT receipt encoding includes RPT classification, barangay, PIN, tax declaration number, Basic Tax, SEF, and special levy items.
The system computes disposition of proceeds for RPT after saving.
The training manual shows sharing percentages:
Province: Basic 35 percent, SEF 50 percent
Municipality: Basic 40 percent, SEF 50 percent
Barangay: Basic 25 percent
City: Basic 70 percent, SEF 100 percent, Barangay 30 percent
```

SEF guideline finding:

```text
SEF comes from the additional 1 percent real property tax.
For provinces/municipalities, the additional 1 percent collected in the province is shared equally by the province and municipality.
SEF proceeds support supplementary annual budgetary needs for public schools through the Local School Board.
```

BLGF memo finding:

```text
BLGF MC No. 001.2023 extended the Q1 FY2023 LIFT Reports deadline to 15 July 2023 due to LIFT enhancement and roll-out.
Sanctions for late submission were temporarily suspended only up to that deadline.
```

Special Health Fund finding:

```text
LIFT v4 adds SHF-related items and columns in SRE/SRS/SOE reports.
SHF is a separate special fund under Universal Health Care, intended for health services, health system operating costs, capital investments, and health worker remuneration/incentives.
```

Implication for future web reporting:

```text
The Firebird collection reports can supply receipt-level and collection-detail data.
The SRE report also needs budget targets, appropriations, expenditures, fund/cash balance, debt service, non-income receipts, continuing appropriations, and SRE-NGAS reconciliation.
Not all SRE fields can come from cashier collection tables alone.
```

The file `SRS_Zamboanguita_Q1_2026.xls` is the LIFT/eSRE Statement of Receipts Sources report:

```text
Report: Statement of Receipts Sources
LGU: Zamboanguita, Negros Oriental
Period Covered: Q1, 2026
Sheet: srs_report
Rows: 188
Columns: 11
```

SRS is the detailed source schedule behind the receipts side of SRE. It shows:

```text
Particulars
Account Code
Income Target / Approved Budget
Actual Receipts
Excess of Actual vs Target
Percent over/under target
```

Key Q1 2026 SRS values found:

```text
Tax Revenues actual: 5,756,093.62
Real Property Tax - Basic actual: 1,280,507.11
  Current Year: 1,075,939.17
  Current Year Fines/Penalties: 0.00
  Prior Years: 181,166.56
  Prior Years Fines/Penalties: 23,401.38
Tax on Business actual: 4,110,483.16
Other Taxes actual: 365,103.35
Community Tax - Individual actual: 365,103.35
Non-Tax Revenues actual: 5,296,645.23
Regulatory Fees actual: 2,938,113.20
Service/User Charges actual: 311,418.50
Economic Enterprises actual: 2,041,016.10
Other Income/Receipts actual: 6,097.43
Total Income - Local Sources actual: 11,052,738.85
National Tax Allotment actual: 47,498,793.00
Total General Fund actual: 58,551,531.85
Special Education Fund actual: 1,600,633.89
Grand Total GF + SEF actual: 60,152,165.74
Advance Payment for RPT: 0.00
```

Relationship between SRS and SRE:

```text
SRS explains the receipt/source breakdown.
SRE summarizes receipts plus expenditures and fund/cash balances.
SRS Grand Total actual 60,152,165.74 matches SRE Total Current Operating Income 60,152,165.74.
```

Relationship between SRS and Firebird reports:

```text
Firebird collection reports can supply many SRS actual receipt lines.
Report 23/25/27 can support RPT Basic and SEF lines.
CTC reports support Community Tax lines.
Other Fees and Charges reports support business taxes, permits, service/user charges, economic enterprise, and miscellaneous local-source lines.
Income targets/approved budget values are not from cashier collections and must come from LIFT budget/target modules or be imported/encoded separately.
```

The file `SOE_Zamboanguita_Q4_2025 (3).xls` is the LIFT/eSRE Statement of Expenditures report:

```text
Report: Statement of Expenditures
LGU: Zamboanguita, Negros Oriental
Period Covered: Q4, 2025
Sheet: soe_report
Rows: 224
Columns: 21
```

SOE is the expenditure-side detail schedule. It shows:

```text
Particulars
NGAS Code
Budget Appropriation by PS / MOOE / FE / CO / Total
Actual Expenditures by PS / MOOE / FE / CO / Total
```

SOE allotment class meanings:

```text
PS   = Personal Services
MOOE = Maintenance and Other Operating Expenses
FE   = Financial Expenses / Debt Service
CO   = Capital Outlay
```

Key Q4 2025 SOE values found:

```text
General Fund - General Public Services actual: 118,025,305.30
General Fund - Social Services actual: 31,481,155.52
  Health, Nutrition & Population Control actual: 14,413,569.75
  Social Services and Social Welfare actual: 17,067,585.77
General Fund - Economic Services actual: 14,262,940.26
Debt Service actual: 11,435,581.94
  Principal actual: 7,159,308.94
  Interest and Other Charges actual: 4,276,273.00
Total General Fund actual expenditures: 175,204,983.02
Total SEF actual expenditures: 1,047,349.56
Total Expenditures actual: 176,252,332.58
Payment of Prior Year Accounts Payable - GF: 4,827,195.39
Payment of Prior Year Accounts Payable - SEF: 75,170.83
Continuing Appropriation actual: 95,065,071.82
```

Relationship between SOE, SRS, and SRE:

```text
SRS = detailed receipts/source side.
SOE = detailed expenditures/use side.
SRE = summary that combines receipts, expenditures, non-operating items, and fund/cash balance.
```

Relationship between SOE and Firebird cashier reports:

```text
SOE is mostly not from cashier collection tables.
It needs budget/expenditure/appropriation/disbursement data, such as PS, MOOE, FE, CO, debt service, accounts payable, and continuing appropriations.
Firebird collection reports can support SRS receipt lines, but SOE needs a separate expenditure data source or LIFT import.
```

## Troubleshooting Negative SOE / Fund Cash Balance

Negative SOE or Fund/Cash Balance values are usually caused by mismatch between budget-side data, accounting-side actual expenditures, and treasurer-side cash/fund balance data.

Common causes:

```text
1. Supplemental Budget was approved but not encoded in LIFT.
2. Supplemental Budget was encoded under the wrong office, function, PPA, NGAS code, fund, or allotment class.
3. Actual expenditures are higher than encoded budget appropriation.
4. Actual expenditures were encoded in the wrong allotment class: PS, MOOE, FE, or CO.
5. Actual expenditures were encoded under the wrong fund: GF vs SEF.
6. Accounts Payable or prior-year payment was duplicated or encoded in the wrong section.
7. Continuing Appropriation budget/expenditure was not encoded correctly.
8. Fund/Cash Beginning Balance is wrong or not reconciled with prior year ending balance.
9. Cash in bank / fund balance composition is incomplete.
10. Debt service principal/interest was encoded under the wrong line.
```

LIFT areas to check:

```text
Authorized Budget > Budget Appropriation > Expenditures
Authorized Budget > Budget Appropriation > Debt Services
Authorized Budget > Budget Appropriation > Unappropriated Surplus
Actual Transaction > Expenditures > Expenditures
Actual Transaction > Expenditures > Accounts Payable
Actual Transaction > Expenditures > Debt Services
Actual Transaction > Others > Fund/Cash Balance
Actual Transaction > Others > SRE-NGAS Reconciliation
```

Office responsibility:

```text
Budget Office:
- Original appropriation
- Supplemental Budget
- Realignment/augmentation
- Continuing appropriation budget

Accounting Office:
- Actual expenditures
- Disbursements
- Accounts payable
- Prior-year AP payments
- NGAS classification

Treasurer's Office:
- Receipts
- Cash beginning and ending balance
- Debt service payments
- Fund/cash balance reporting
- SRE/SRS/SOE report checking/submission
```

Fixing sequence:

```text
1. Identify the exact negative SOE/SRE row.
2. Check if actual expenditures are greater than budget appropriation.
3. If yes, ask Budget Office if there is a Supplemental Budget, realignment, or augmentation.
4. Encode/update the Supplemental Budget in the correct LIFT budget appropriation screen.
5. Check fund: GF or SEF.
6. Check allotment class: PS, MOOE, FE, or CO.
7. Check office/function/PPA/NGAS code.
8. Ask Accounting to verify actual expenditures and accounts payable.
9. Check Fund/Cash Balance beginning amount against prior year's ending balance.
10. Check continuing appropriation tab if prior-year appropriations/expenditures are involved.
11. Recompute/regenerate SOE and SRE.
12. Recheck eSRE alerts: Actual Expenditures > Budget Appropriation, Fund Balance End < 0, Amount Available for Appropriations/Operations < 0.
```

For next-quarter preparation, use this checklist:

```text
README\NEXT_QUARTER_LIFT_REPORT_CHECKLIST.md
README\NEXT_QUARTER_LIFT_REPORT_CHECKLIST.docx
```

The `.docx` version is the printable Word copy. It is generated by:

```text
README\build_next_quarter_checklist_docx.py
```

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

## Future Web App Repository Notes

Repository:

```text
C:\Users\LIFT-LAPTOP\OneDrive\Desktop\ESRE_REPORT\LGU_TreasuryReportingSystem
https://github.com/richer18/LGU_TreasuryReportingSystem
```

Planned stack:

```text
Backend: Laravel REST API
Frontend: React/Vite
Runner: Python reporting/import jobs
Server runner: local startup helpers for Laravel/React development servers
Docs: documentation, workflow notes, table mapping, formulas, and report rules
Database/reporting source: Firebird .FDB first, future MySQL import/cache layer
Rule: all Python scripts for the future web app must be saved in LGU_TreasuryReportingSystem\runner.
```

Initial hygiene/API setup completed:

```text
Root .gitignore added for Laravel, React, local databases, generated reports, node_modules, vendor, and environment files.
Frontend axios dependency installed.
Laravel API routing enabled in backend/bootstrap/app.php.
Root README.md added to explain folder responsibilities.
docs/ARCHITECTURE.md added for system architecture.
runner/README.md added for Python runner purpose and safety rules.
server_runner/README.md added for local server helper purpose.
server_runner batch files added:
- server_menu.bat
- start_backend.bat: starts Laravel API on 127.0.0.1:8000
- start_frontend.bat: starts React/Vite on 127.0.0.1:5173
- start_all.bat: starts both servers
- check_ports.bat: checks backend/frontend ports
Initial backend/routes/api.php added with:
- GET /api/health
- GET /api/reports
- GET /api/reports/{number}
Report catalog module started:
- app/Reports/ReportCatalog.php
- app/Http/Controllers/Api/ReportCatalogController.php
git_hub_updater scripts improved:
- GIT_OPTIONAL_LOCKS=0 added to reduce OneDrive Git lock issues.
- Push/pull scripts now require the current branch to be main.
- Push script fetches GitHub first and blocks if remote main has changes missing locally.
- Pull script blocks when local changes exist.
- git_hub_updater/README.md added.
Firebird connection bridge started:
- PHP installation does not currently include pdo_firebird/interbase.
- Laravel backend connects to Firebird through Python runner first.
- runner/firebird_probe.py uses read-only Firebird settings: READ_COMMITTED_RO, no_db_triggers, no_gc.
- All future Python scripts for this app must be placed under runner/.
- backend/config/firebird.php stores connection settings.
- backend/app/Services/FirebirdProbeService.php calls the Python probe.
- backend/app/Http/Controllers/Api/FirebirdStatusController.php exposes GET /api/firebird/status.
- backend/routes/console.php exposes php artisan firebird:status.
- Verified against C:\ZAMBOANGUITA_DB\ZAMBOANGUITA.FDB: ok=true, 237 user tables, 1 view.
Frontend Firebird status display added:
- frontend/src/App.jsx now calls GET /api/firebird/status through the axios instance.
- frontend/src/App.css and frontend/src/index.css now show a simple database connection dashboard.
- Frontend displays Connected/Disconnected, user table count, view count, bridge mode, database path, client library, and sample tables.
- Verified npm run build and confirmed frontend http://127.0.0.1:5173 is running.
- API CORS response allows frontend origin.
- If Vite shows a stale missing App.css import overlay, restart the frontend with server_runner/start_frontend.bat; it now uses --force to rebuild Vite cache.
Frontend UI/UX shell added:
- Dashboard page added with Firebird connection status and database metrics.
- Sidebar navigation added for Dashboard, General Fund, Trust Fund, Community Tax, Real Property Tax, and Settings.
- Fund report pages added with date range controls and grouped report lists.
- Temporary Reports master button/page added to the sidebar with report list 1-31.
- Settings page added with API/frontend connection paths and Python runner rule.
- lucide-react installed for sidebar and page icons.
- Verified npm run lint and npm run build.
Laravel/React authentication implemented:
- laravel/sanctum installed for REST API bearer-token authentication.
- backend/app/Http/Controllers/Api/AuthController.php added.
- backend/routes/api.php now includes POST /api/login, GET /api/user, and POST /api/logout.
- backend/app/Models/User.php now uses Sanctum HasApiTokens.
- backend/config/permissions.php added for role permission lists.
- User auth profile migration added: role and account_status.
- Sanctum personal_access_tokens migration added.
- Default local admin seed added: admin@zamboanguita.local / admin123.
- Frontend AuthProvider/useAuth/authStorage added under frontend/src/auth.
- Frontend page code was split by folder to keep files maintainable:
  - frontend/src/pages/Login/LoginPage.jsx
  - frontend/src/pages/Dashboard/DashboardPage.jsx
  - frontend/src/pages/Reports/ReportsPage.jsx
  - frontend/src/pages/Settings/SettingsPage.jsx
  - frontend/src/data/reportCatalog.js
  - frontend/src/utils/firebirdStatus.js
- frontend/src/axiosinstance/axiosInstance.js now attaches Authorization: Bearer token and clears token on 401.
- frontend/src/App.jsx login now calls Laravel /api/login, persists token in localStorage, validates saved session through /api/user, and logs out through /api/logout.
- Laravel auth SQLite DB uses C:\Users\LIFT-LAPTOP\AppData\Local\Temp\lgu_treasury_auth.sqlite because SQLite schema writes under the OneDrive project folder caused disk I/O/journal errors.
- Firebird remains the reporting source and is still accessed read-only through the Python runner bridge.
- Verified API login/user/logout through HTTP, php artisan test, npm run lint, and npm run build.
```

The initial report registry mirrors the current ESRE report list 1-31. Reports 21-31 are marked as implemented in the Python/FDB reporting script, while 1-17 remain source-query/planning items and 18-20 are process documentation.

Current git caution: if .git/index.lock appears again, close Git/GitHub Desktop/VS Code Git operations first. Delete the lock file only when no git.exe process remains.

## Existing ETMS Reference Review

Reviewed old system in read-only mode:

```text
C:\Users\LIFT-LAPTOP\OneDrive\Desktop\ElectronicTreasurerManagementSystem\frontend
C:\Users\LIFT-LAPTOP\OneDrive\Desktop\ElectronicTreasurerManagementSystem\backend
```

Important: do not run MySQL on this laptop for the old ETMS because it may conflict with LIFT server ports.

The old system is useful as a workflow/reference source:

```text
Laravel 12 backend with Sanctum login
React CRA frontend with Material UI/Toolpad dashboard
Role and module permission middleware
General Fund, Trust Fund, Community Tax, Real Property Tax modules
Full Report / RCD / ESRE / Summary reports
RCD accountable forms and receipt range workflows
Dashboard charts and collector summaries
Python scraper/tools folders for old MySQL import workflow
```

New-system decision:

```text
Use old ETMS concepts only.
Do not copy MySQL dependency or old scraper folder layout.
All Python in new system still belongs in LGU_TreasuryReportingSystem\runner.
New system remains Firebird-read-only first through Laravel REST API + Python runner.
```

Detailed note saved:

```text
firebird_metadata\existing_etms_system_review.md
```

## Old ETMS General Fund Reference Review

Reviewed old General Fund module in read-only mode:

```text
C:\Users\LIFT-LAPTOP\OneDrive\Desktop\ElectronicTreasurerManagementSystem\frontend\src\FRONTEND\components\ABSTRACT\GF
```

Main useful patterns:

- GF page has month/year filters, search, add/edit/view/delete actions, daily report, financial report, receipt/collector report, dashboard total cards, and Excel export.
- Payment entry uses header fields: date, taxpayer/name, receipt number, receipt type, cashier, local TIN.
- Payment details are source/rate rows: source_id, description, amount.
- Total is sum of detail amounts.
- Cash Tickets has an old workflow rule: auto receipt number `00YYYYMMDD`.
- GF report categories include Total Revenue, Tax on Business, Regulatory Fees, Receipts from Economic Enterprises, and Service/User Charges.
- SOC/financial report uses source fields matching ESRE source names, including Cockpit_Prov_Share and Cockpit_Local_Share.
- Old formula separates cockpit provincial share: municipal total = total - cockpit provincial share.
- Collector receipt report supports date range, month/year, collector, report type, and receipt range filters.

Detailed note saved:

```text
firebird_metadata\old_etms_general_fund_review.md
```

## New LGU Treasury System - General Fund UI/API Implemented

Implemented a component-based General Fund page in the new React/Laravel app:

```text
LGU_TreasuryReportingSystem\frontend\src\pages\GeneralFund\GeneralFundPage.jsx
LGU_TreasuryReportingSystem\frontend\src\pages\GeneralFund\components\
LGU_TreasuryReportingSystem\frontend\src\pages\GeneralFund\hooks\useGeneralFundData.js
LGU_TreasuryReportingSystem\frontend\src\pages\GeneralFund\utils\generalFundFormat.js
```

Frontend components:

- `GeneralFundFilters` - date range, collector, receipt range, refresh.
- `GeneralFundSummaryCards` - total GF, receipt count, collector count, receipt reference.
- `GeneralFundActionStrip` - opens secondary report views in dialogs.
- `GeneralFundDialog` - reusable dialog shell for General Fund report views.
- `GeneralFundCategoryBreakdown` - dialog body for Tax on Business, Regulatory Fees, Receipts from Economic Enterprises, Service/User Charges, Miscellaneous.
- `GeneralFundDailyTable` - dialog body for daily totals.
- `GeneralFundSourceBreakdown` - dialog body for top source/revenue code totals.
- `GeneralFundReceiptReport` - dialog body for collector/receipt filter report.
- `GeneralFundCollectionsTable` - receipt-level collection table.

UI decision:

- The main General Fund page should stay clean: filters, summary cards, action buttons, and receipt-level collections table.
- Default General Fund date filter should start on the first day of the current month and end on today.
- Main General Fund collections table should have client-side pagination with 10/25/50 rows per page.
- Category Breakdown, Collector Receipt Report, Daily Collection, and Source Breakdown should only appear after clicking their action button, inside a dialog.
- Dialog contents should use scrollable table layouts with sticky table headers, not card/list blocks, because LGU report review requires scanning many rows.
- All current General Fund tables under `frontend/src/pages/GeneralFund` should use MUI table components, not raw HTML `<table>` or custom div-grid tables.
- General Fund main page performance note:
  - Avoid loading every secondary report on initial page open.
  - Initial load should fetch only summary, main collections, and collectors.
  - Heavy secondary reports such as Source Breakdown and Daily Collection should load when their dialog/report view is opened.
  - Reason: every General Fund endpoint currently starts a Python Firebird runner process and scans/classifies `PAYMENT` + `PAYMENTDETAIL`; calling many endpoints at once makes the main table slower.
- General Fund report dialogs should use a wider reusable dialog width up to 1280px so MUI tables fit better before horizontal scrolling.
- Main General Fund Collections table should use MUI `Table` components and show columns in this order:
  - Date
  - Taxpayer
  - Receipt
  - Collector
  - Total
  - Action
- General Fund collection results should be ordered by collection date ascending, then receipt number ascending.
- Main General Fund Collections action column should show `View` and `Update`.
- `View` and `Update` dialogs should visually follow the old ETMS `GF\GeneralFund.jsx` main-table popup style:
  - wide MUI dialog
  - gradient title/header
  - payment header summary
  - detail/table section
  - clean action buttons
- `View` opens a read-only collection detail dialog.
- `Update` opens a form-style update dialog only; save remains disabled and no Firebird write should happen until an official update workflow is designed.
- Collector Receipt Report should follow the old ETMS `GenerateReport.jsx` / `ReceiptCollectionReportDialog.jsx` workflow:
  - filter by daily/date range or by month/year
  - choose collector/cashier
  - optionally filter receipt number from/to
  - show Date, Collector, Receipt Type, Receipt No., Taxpayer, Lines, and Total in a paginated table
  - show total collection for the result set
  - allow CSV download of the result
  - treat AMABELLA General Fund rows as Cash Tickets in the displayed receipt type
- The new app translates the old POST `generate-report` behavior into a read-only REST GET call to `/api/general-fund/receipt-report`.
- Daily Collection should follow the old ETMS `GF\TableData\DailyTable.jsx` workflow:
  - choose month and year
  - load daily General Fund summary rows
  - show daily category columns for Tax on Business, Regulatory Fees, Receipts From Economic Enterprise, Service/User Charges, Miscellaneous, receipt count, and total
  - show a grand total row and summary totals
  - use MUI `Table`, `TableContainer`, `TablePagination`, and related MUI table components
  - paginate the main daily table with 5/10/25 rows per page
  - provide CSV download
  - provide a read-only `View` action per date that opens receipt-level transaction details with pagination
  - do not implement old ETMS comment save/edit actions in the new app yet because the current Firebird reporting bridge is read-only
- The daily Firebird runner output now includes category totals in addition to receipt count, receipt range, and day total.
- Frontend auth now dispatches an `auth:unauthorized` event on API 401 responses so the app state returns to login instead of leaving stale pages that show raw `Unauthenticated` API text.

Implemented read-only Firebird runner:

```text
LGU_TreasuryReportingSystem\runner\general_fund_readonly.py
```

Runner logic:

- Connects using the existing Firebird read-only connection helper.
- Uses `PAYMENT` + `PAYMENTDETAIL` + `T_ITAXTYPE`.
- General Fund classification now follows Report 29, Abstract of General Collections:
  - pull non-void, non-RPT payment details from `PAYMENT` + `PAYMENTDETAIL`
  - classify lines with the same `classify_summary_source()` rules used by Report 29
  - exclude Community Tax lines
  - exclude Trust Abstract names: Building Permit Fee, Electrical Permit Fee, Zoning Fee, Livestock, Diving Fee
  - do not rely only on `PAYMENTDETAIL.FUNDTYPE_CT = 'GF'` because some valid General Fund rows have blank detail fund type
- Returns JSON only; no database write.

Implemented Laravel service/controller/routes:

```text
LGU_TreasuryReportingSystem\backend\app\Services\GeneralFundReportService.php
LGU_TreasuryReportingSystem\backend\app\Http\Controllers\Api\GeneralFundController.php
```

Protected API endpoints under Sanctum auth:

```text
GET /api/general-fund/summary
GET /api/general-fund/collections
GET /api/general-fund/daily
GET /api/general-fund/sources
GET /api/general-fund/collectors
GET /api/general-fund/receipt-report
GET /api/general-fund/payment-details/{paymentId}
```

Latest General Fund collection detail update:

- Main General Fund table keeps `Date | Taxpayer | Receipt | Collector | Total | Action`.
- `View` and `Update` now call a read-only `payment-details` runner report by `PAYMENT.PAYMENT_ID`.
- Main General Fund collections are now grouped by receipt/date/taxpayer/collector instead of raw `PAYMENT_ID` only. This avoids duplicate receipt rows when Firebird has sibling payment headers for the same OR number.
- Verified duplicate examples:
  - receipt `9808707` on `2026-01-05` had two Firebird `PAYMENT_ID` rows (`AF51` and `AF56`) and now appears once in the main table with `line_count = 2`
  - receipt `9808708` on `2026-01-05` also had two Firebird `PAYMENT_ID` rows and now appears once
- For `2026-01-01` to `2026-01-31`, exact audit found 8 true duplicate receipt groups before grouping and 0 duplicate receipt rows after grouping:
  - `9808707`, `9808708`, `9811795`, `9810739`, `9813000`, `9814308`, `0034551`, `9814099`
- The View/Update detail endpoint receives receipt/date/taxpayer/collector filters so it can still show all child payment lines under the grouped receipt.
- The detail dialog now displays the actual paid source line from `PAYMENTDETAIL` classification, for example `Water Fees`, instead of the old generic `General Fund payment total`.
- Description is now treated as parent/child:
  - parent = report source bucket from Report 29 classification, for example `Med./Lab. Fees`
  - child = `PAYMENTDETAIL.SOURCEID` lookup description from `T_OTHERPAYMENTRATE.OPRATE_ID`, for example `Source 918 - URINALYSIS FEE`
  - sample verified receipt `9808709` on `2026-01-05`: parent `Med./Lab. Fees`, children `URINALYSIS FEE` and `STOOL LABORATORY/FECALYSIS FEE`
- The detail breakdown table no longer includes a `Collector` column; collector remains available in the main table/header context.
- The General Fund View dialog was reduced from a wide 1100px layout to a more compact 820px layout.
- `Update` remains a read-only prepared layout; saving is still disabled until an official edit/posting process is approved.

Verified for `2026-01-01` to `2026-01-31`:

```text
General Fund total: 4,898,687.21
Receipt count: 3,882
Collector count: 5
Category count: 5
Daily rows: 23
Source rows: 30
Collectors: 5
```

Notes:

- Some receipt numbers in Firebird are text-like values such as `NOV. 2025`, so receipt references must be displayed as text, not treated as purely numeric ranges.
- This is read-only reporting. Add/edit/delete General Fund payment workflows from the old ETMS were analyzed but intentionally not implemented in the new app.
- Verified `npm run lint`, `npm run build`, direct Python runner, protected API smoke tests, and browser UI checks for Collector Receipt Report / Daily Collection.

General Fund UI update:

- Removed the `Receipt Reference` summary card.
- Category Breakdown now has its own date filter and a `Total Overall` footer row.
- `Miscellaneous` category is treated under `Regulatory Fees` for General Fund category reporting.
- General Fund View/Update detail descriptions now show the child/source fee description only, for example `WATER PAYMENT FEE`; the parent label `Water Fees`, `Source 815`, and detail `Receipt Type` column were removed from the detail table.

- Removed the General Fund summary cards row (`Total General Fund`, `Receipts`, and `Collectors`) from the General Fund page.

- Added a `Total Overall` footer row to the General Fund Source Breakdown dialog.

- Added a Month / Year filter to the General Fund Source Breakdown dialog.

- Reports page now has a Month and Year filter plus per-report Download buttons. Current download output is a CSV metadata/manifest file for the selected month while backend Excel/PDF report generation is still pending.

- Reports page Generate workflow updated: report rows 21 to 31 now use `Generate`; generated preview appears below using an official report-template style with Download and Print buttons.

- Reports page simplified: removed the visible master table/date range text. UI now shows Report title, Month and Year, a report dropdown under Generate Report, and a single Generate Report button. Dropdown includes reports 21-31 plus an Other Reports group for 1-20.

- Generated report preview now uses template-specific Excel-style layouts based on files in `LGU_TreasuryReportingSystem/template`, including Summary of Collections, RPT records, Abstracts, Sharing, Provincial Coding, and Full Report shapes.

- Reports page Generate workflow is now connected to real Firebird read-only preview data for Summary reports:
  - New runner: `LGU_TreasuryReportingSystem\runner\report_preview_readonly.py`
  - New Laravel service: `LGU_TreasuryReportingSystem\backend\app\Services\ReportPreviewService.php`
  - New Laravel controller: `LGU_TreasuryReportingSystem\backend\app\Http\Controllers\Api\GeneratedReportController.php`
  - New protected endpoint: `GET /api/generated-reports/{number}/preview?date_from=YYYY-MM-DD&date_to=YYYY-MM-DD`
  - Frontend file updated: `LGU_TreasuryReportingSystem\frontend\src\pages\Reports\ReportsPage.jsx`
  - Config key added: `firebird.report_preview_script`
- Currently connected reports:
  - Report 21: Summary of Collection
  - Report 22: Summary of Collection no RPT
  - Report 23: Summary of Collection RPT
  - Report 24 uses the same RPT summary runner shape as Report 23 while the more detailed SUMMARY-sheet mapping is still pending.
- The Generate button now calls the Laravel API, Laravel calls the Python runner, and the Python runner queries `.FDB` in read-only mode using `firebird_probe.connect()`.
- The preview table no longer depends on placeholder `RESULT` values for reports 21-24 when the API returns data.
- January 2026 verification:
  - Report 21 total collections from runner: `10,549,979.15`
  - Report 23 RPT total from runner: `4,820,216.46`
  - These match earlier handoff totals.
- Build/route verification:
  - `php -l app\Services\ReportPreviewService.php` passed
  - `php -l app\Http\Controllers\Api\GeneratedReportController.php` passed
  - `php artisan route:list --path=generated-reports` shows `GET api/generated-reports/{number}/preview`
  - `npm run build` passed

- Summary of Collection preview UI header was revised to match the Excel template:
  - `Sources of Collections`, `Total Collections`, and `National` now use two-row vertical header cells.
  - `Provincial` now spans `General Fund`, `Special Educ. Fund`, and `Total`.
  - `Municipal` now spans `General Fund`, `Special Educ. Fund`, `Trust Fund`, and `Total`.
  - `Barangay Share` and `Fisheries` now use two-row vertical header cells.
  - Verified again with `npm run build`.

- Reports page Download button now downloads a generated `.xlsx` file instead of a CSV placeholder for Summary reports:
  - New runner: `LGU_TreasuryReportingSystem\runner\report_excel_export_readonly.py`
  - New config key: `firebird.report_excel_script`
  - New protected endpoint: `GET /api/generated-reports/{number}/download?date_from=YYYY-MM-DD&date_to=YYYY-MM-DD`
  - Frontend `Download` now requests the endpoint as a blob and saves the Excel file locally.
  - Current Excel download support covers reports 21, 22, 23, and 24.
  - The exporter uses the uploaded templates in `LGU_TreasuryReportingSystem\template` and fills real `.FDB` data.
  - January 2026 test export verified:
    - workbook period cell: `Month of January 2026`
    - Retailing row: `2,800`
    - TOTAL row: `10,549,979.15`
  - Verified with Python workbook readback, PHP syntax checks, `php artisan route:list --path=generated-reports`, and `npm run build`.

- Summary of Collection preview table was improved again for Excel-like column balance:
  - Added fixed `colgroup` widths to the Summary table.
  - Widened the Summary preview sheet to a landscape-style `1380px`.
  - Balanced Provincial and Municipal child columns so `General Fund` no longer stretches too wide.
  - Right-side columns (`Trust Fund`, `Total`, `Barangay Share`, `Fisheries`) are kept visible in the preview area.
  - Verified with `npm run build`.

- Fixed Report Download 500 error:
  - Symptom: clicking `Download` returned `Request failed with status code 500`.
  - Root cause: Laravel/Herd launched `C:\Python314\python.exe`, but the PHP child process did not include the Python user-site package directory where `openpyxl` is installed:
    `C:\Users\LIFT-LAPTOP\AppData\Roaming\Python\Python314\site-packages`
  - Fixes applied:
    - `backend\config\firebird.php` now hard-prefers `C:\Python314\python.exe` when present.
    - `backend\app\Services\ReportPreviewService.php` now injects `PYTHONPATH` for the Python user-site packages.
    - `runner\report_excel_export_readonly.py` now also appends the user-site package path before importing `openpyxl`.
    - `frontend\src\pages\Reports\ReportsPage.jsx` now parses blob error responses so future backend JSON errors display readable messages.
  - Verified actual authenticated HTTP download:
    - Endpoint: `GET /api/generated-reports/21/download?date_from=2026-01-01&date_to=2026-01-31`
    - Result: `.xlsx` file downloaded successfully, size about 9.3 KB.
    - Workbook readback still confirmed `Month of January 2026`, `Retailing = 2,800`, `TOTAL = 10,549,979.15`.
  - Backend was restarted and left running on `http://127.0.0.1:8000`.
  - Verified again with PHP syntax checks, direct Python runner export, API health check, authenticated HTTP download, and `npm run build`.

- Business Permit / BPLS Excel exports analyzed:
  - Folder: `BUSINESS_PERMIT_REPORT`
  - Files reviewed:
    - `ABSTRACT_OF_GENERAL_COLLECTION-BPLS-ALL-2026_06_19_20_17_06.xlsx`
    - `BUSINESS_ESTABLISHMENT-BPLS-ALL-2026_06_19_20_21_45.xlsx`
    - `REGISTERED_BUSINESSES-BPLS-2026_06_19_20_18_01.xlsx`
  - All three files are flat one-sheet exports with no formulas.
  - Abstract of General Collection:
    - Header row: 7
    - Transaction rows: 852 plus 1 summary row
    - Period in file: `2026-01-05` to `2026-03-31`; actual OR date range found: `2026-01-06` to `2026-03-30`
    - Unique business IDs: 850
    - Unique OR numbers: 851
    - Transaction mix: 764 Renewal, 86 New, 1 Quarterly, 2 blank
    - Transaction-only Amount Paid total: `6,582,603.21`
    - Transaction-only Business Tax total: `4,076,164.66`
    - One duplicate business ID: `0704625-2024-0000711`
  - Business Establishment:
    - Rows: 881
    - Unique business IDs: 858
    - Unique OR numbers: 857
    - Application date range: `2026-01-06` to `2026-03-31`
    - Total Amount Paid: `6,618,461.96`
    - Missing permit rows: 25
    - Missing OR rows: 24 with amount total `10,642.50`
    - 12 duplicate business IDs; examples include `0704625-2024-0000021`, `0704625-2024-0001042`, `0704625-2024-0000393`
  - Registered Businesses:
    - Rows: 1,123
    - Unique business IDs: 1,123
    - Status of application: 849 For Pick-Up, 177 blank, 56 Issued, 41 Paid
    - Registration status: 988 Renewal, 135 New
    - Missing permit count: 177
  - Cross-file reconciliation:
    - All Abstract business IDs are present in Business Establishment and Registered Businesses.
    - Business Establishment has 1 business ID not in Registered Businesses: `0704625-2024-0000467`
    - Registered Businesses has 266 business IDs not in Business Establishment, likely unpaid/not-yet-permitted/for pick-up records.
    - Abstract ORs not in Establishment: 0
    - Establishment ORs not in Abstract: 6, totaling `25,216.25`
    - Establishment total difference vs Abstract: `35,858.75`, explained by `25,216.25` extra ORs plus `10,642.50` missing-OR amount rows.

- CMCI Annex A-B Business Permit Registration template analyzed:
  - File reviewed: `BUSINESS_PERMIT_REPORT\2025-2026_ANNEX-A-B_cmci_report.xlsx`
  - Recommended report list item: `32. CMCI Annex A-B Business Permit Registration Report`
  - Added to `run_collection_query.py` report list as Report 32.
  - Running Report 32 now generates an `.xlsx` output using the CMCI Annex template.
  - Source workbooks used:
    - `BUSINESS_PERMIT_REPORT\REGISTERED_BUSINESSES-BPLS-*.xlsx`
    - `BUSINESS_PERMIT_REPORT\BUSINESS_ESTABLISHMENT-BPLS-*.xlsx`
  - Workbook has 3 sheets:
    - `Annex A (Jan. to Dec. 2025)`
    - `Annex B (Jan. to Mar. 2026`
    - `PSIC`
  - The Annex sheets are blank CMCI/DTI templates with formatting and dropdown validations, not exported data.
  - No formulas were found.
  - Effective data-entry columns are A to P:
    - A `LGU`
    - B `Province`
    - C `Region`
    - D `Classification`
    - E `LGU Type`
    - F `Business Name`
    - G-I `Business Address` parts
    - J `Owner's Name`
    - K `Industry / Nature of Business`
    - L `Business Type`
    - M `Capitalization Size`
    - N `New / Renewal`
    - O `Year of Registration`
    - P `Permit No.`
  - The `PSIC` sheet contains the allowed major industry categories, including Agriculture, Mining, Manufacturing, Wholesale/Retail, Accommodation/Food, Real Estate, Education, Health, and other major categories.
  - Important implementation note:
    - Columns K to N are text-sensitive and must use dropdown-approved values.
    - `Business Type` must be normalized from BPLS values, for example `SOLE PROPRIETORSHIP` -> `Single Proprietor`.
    - `Capitalization Size` must be computed from capital/gross thresholds: Micro, Small, Medium, Large.
    - `Industry / Nature of Business` requires a mapping from BPLS `Line of Business` or business line code to PSIC major category.
  - Suggested source data:
    - Prefer `REGISTERED_BUSINESSES-BPLS` for official registration/masterlist fields.
    - Use `BUSINESS_ESTABLISHMENT-BPLS` to enrich permit number, OR/payment, gross sales, capitalization, and application type.
  - January to March 2026 implementation test:
    - Command: `python .\run_collection_query.py 32 2026-01-01 2026-03-31`
    - Output: `firebird_metadata\output\query_32_2026-01-01_to_2026-03-31.xlsx`
    - Rows exported: `856`
    - Target sheet filled: `Annex B (Jan. to Mar. 2026`
  - CMCI static values confirmed/updated:
    - Region: `REGION VII (CENTRAL VISAYAS)`
    - Classification: `Third Class Municipality`
  - Capitalization Size now writes the threshold text expected by the dropdown:
    - `Micro (less than P3000000)`
    - `Small ( P3000001 - P15000000)`
    - `Medium (P15000001 - P100000000)`
    - `Large (more than P100000000)`
  - The first implementation uses heuristic PSIC mapping from BPLS line-of-business text. This should still be reviewed before official CMCI portal submission.

- Report 33 Tax on Business Summary from BPLS Business Tax implemented:
  - Command: `python .\run_collection_query.py 33 2026-01-01 2026-03-31`
  - Output: `firebird_metadata\output\query_33_2026-01-01_to_2026-03-31.xlsx`
  - Source files:
    - `BUSINESS_PERMIT_REPORT\ABSTRACT_OF_GENERAL_COLLECTION-BPLS-*.xlsx`
    - `BUSINESS_PERMIT_REPORT\BUSINESS_ESTABLISHMENT-BPLS-*.xlsx`
  - Workbook sheets:
    - `Summary` visible
    - `Detail` hidden for audit
    - `Notes` hidden for audit
  - Visible Summary columns only:
    - `Category`
    - `Business Tax`
    - `Fines & Penalties / Surcharge`
    - `Total`
  - Core rule:
    - Business tax value comes from Abstract of General Collection column `Business Tax`.
    - Fines & Penalties value comes from Abstract of General Collection column `Surcharge`.
    - Classification uses Business Establishment `Business Nature` and `Business Line`, matched by OR Number first and Business ID fallback.
  - Categories:
    - Manufacturing
    - Distributor
    - Retailing
    - Banks & Other Financial Int.
    - Other Business Tax
    - Fines & Penalties
  - January to March 2026 verification:
    - Visible summary rows: `6`
    - Hidden detail rows: `773`
    - Business Tax total: `4,076,164.66`
    - Surcharge / Fines & Penalties total: `29,006.00`
    - Grand total: `4,105,170.66`
    - Category allocation:
      - Manufacturing: `364,734.49`
      - Distributor: `67,354.00`
      - Retailing: `1,616,323.83`
      - Banks & Other Financial Int.: `124,295.34`
      - Other Business Tax: `1,903,457.00`
      - Fines & Penalties: `29,006.00`
  - Manual review note:
    - The user said these categories may still be manually based on the other business tax app/database.
    - Current implementation provides a generated starting point from BPLS Business Nature/Business Line.
  - Reports 21 and 22 now connect to Report 33 for Tax on Business rows:
    - Manufacturing
    - Distributor
    - Retailing
    - Banks & Other Financial Int.
    - Other Business Tax
    - Fines & Penalties
  - When Report 33 has BPLS data for the selected period, Reports 21 and 22 override those Tax on Business rows with BPLS Business Tax/Surcharge amounts. If no BPLS data exists for the selected period, the reports keep the original Firebird summary values.

## LGU Treasury Web UI Direction - General Fund and Reports

Reviewed the user's preferred UI references:

```text
C:\Users\LIFT-LAPTOP\Downloads\general_fund_ui.tsx
C:\Users\LIFT-LAPTOP\Downloads\reports_catalog_ui.tsx
```

Adopted design direction:

- Use a quiet official treasury reporting style: white panels, light slate backgrounds, deep navy headers, teal primary actions, compact typography, and structured report controls.
- Keep the system operational and report-focused, not a marketing-style dashboard.
- Preserve the existing Laravel API / Python runner behavior; UI polish should not change Firebird read-only reporting logic.

Implemented UI updates in the new LGU Treasury frontend:

- General Fund filters now use a polished toolbar panel with field icons for date and collector filters.
- General Fund action buttons now sit beside a compact "General Fund analytics" context block.
- General Fund dialogs now use an official dark header with icon/title grouping and a cleaner report body.
- General Fund main collections table now has a more refined toolbar and MUI action buttons while keeping columns:
  - Date
  - Taxpayer
  - Receipt
  - Collector
  - Total
  - Action
- Reports generator panel now uses the preferred report-catalog layout:
  - Office label
  - Month and Year selector
  - Generate Report dropdown
  - Generate Report button
  - Short helper note explaining reports 21-31 preview/export behavior
- Generated report preview toolbar now uses a spreadsheet-style Download button and Print button.

Frontend files updated:

```text
LGU_TreasuryReportingSystem\frontend\src\pages\GeneralFund\components\GeneralFundFilters.jsx
LGU_TreasuryReportingSystem\frontend\src\pages\GeneralFund\components\GeneralFundActionStrip.jsx
LGU_TreasuryReportingSystem\frontend\src\pages\GeneralFund\components\GeneralFundDialog.jsx
LGU_TreasuryReportingSystem\frontend\src\pages\GeneralFund\components\GeneralFundCollectionsTable.jsx
LGU_TreasuryReportingSystem\frontend\src\pages\Reports\ReportsPage.jsx
LGU_TreasuryReportingSystem\frontend\src\App.css
```

Verification:

```text
cd LGU_TreasuryReportingSystem\frontend
npm run build
```

Result: build passed. Vite still reports the existing large bundle warning, but no compile errors were found.

## General Fund UI Update - Collector Workflow Split

General Fund page design was revised toward the preferred collection-monitor layout:

- Top General Fund area now shows:
  - `Collection Monitor`
  - `General Fund`
  - `Generate Receipt` button
- The filter card now follows the cleaner layout:
  - Date From
  - Date To
  - Collector
  - Refresh
  - Receipt Range Optional
- The old `Collector Receipt Report` action was split into two clearer workflows:
  - `Collection per Collector` now shows collector-level totals for the active filter period.
  - `Generate Receipt` opens the old collector receipt/date range/receipt range generator function.
- Added a new MUI table component for collector totals:

```text
LGU_TreasuryReportingSystem\frontend\src\pages\GeneralFund\components\GeneralFundCollectorCollections.jsx
```

Collector totals table columns:

```text
Collector | Receipts | Share | Total Collection
```

The table includes a `Total Overall` footer.

Frontend files updated:

```text
LGU_TreasuryReportingSystem\frontend\src\pages\GeneralFund\GeneralFundPage.jsx
LGU_TreasuryReportingSystem\frontend\src\pages\GeneralFund\components\GeneralFundActionStrip.jsx
LGU_TreasuryReportingSystem\frontend\src\pages\GeneralFund\components\GeneralFundFilters.jsx
LGU_TreasuryReportingSystem\frontend\src\pages\GeneralFund\components\GeneralFundCollectorCollections.jsx
LGU_TreasuryReportingSystem\frontend\src\App.css
```

Verification:

```text
cd LGU_TreasuryReportingSystem\frontend
npm run build
```

Result: build passed. Existing Vite large chunk warning remains only a performance warning.

Follow-up update:

- `Collection per Collector` now has its own date filter inside the dialog:
  - Date From
  - Date To
  - Apply Filter
- The filter reloads collector totals from the read-only Laravel endpoint:

```text
GET /api/general-fund/collectors?date_from=YYYY-MM-DD&date_to=YYYY-MM-DD
```

- The collector totals heading now shows the selected date range.
- Verified again with `npm run build`.

Generate Receipt status update:

- Added `Status` column to the Generate Receipt result table.
- Status source fields:
  - `PAYMENT.VOID_BV`
  - `PAYMENT.STATUS_CT`
  - `T_STATUS.DESCRIPTION`
- Status mapping:
  - `VOID_BV = 1` or status description containing `VOID` -> `Void`
  - `STATUS_CT = CNL` or status description containing `CANCEL` -> `Cancelled`
  - otherwise -> `Paid`
- Receipt report now includes void/cancelled receipts for audit/search visibility, but the displayed `Total Collection` sums only rows with `Status = Paid`.
- CSV download now includes the `Status` column.
- January 5, 2026 smoke test found a real cancelled row:
  - receipt `2019857`
  - `STATUS_CT = CNL`
  - `T_STATUS.DESCRIPTION = CANCELLED`
- Reporting rule: official collection reports and totals must be `Paid` only. Exclude cancelled and void rows from reports such as Summary of Collection, Abstracts, Full Report Collections, General Fund main collections, daily collection, source breakdown, category breakdown, and collector totals.
- Firebird paid-only predicate for `PAYMENT`:
  - `COALESCE(PAYMENT.VOID_BV, 0) = 0`
  - `COALESCE(TRIM(PAYMENT.STATUS_CT), '') NOT IN ('CNL', 'CAN', 'CNC', 'CANCEL', 'CANCELLED', 'VOID', 'VOI')`
- For RPT class/detail reporting, also keep `COALESCE(PAYMENTCLASSDETAIL.CANCELLED_BV, 0) = 0`.
- Search Receipt and Generate Receipt audit views may still display `Cancelled` and `Void` receipts so the Treasurer staff can investigate OR status, but those rows must not be counted in official report totals.
- Verified direct runner and frontend build:

```text
python .\LGU_TreasuryReportingSystem\runner\general_fund_readonly.py receipt-report --date-from 2026-01-05 --date-to 2026-01-05 --limit 3
cd LGU_TreasuryReportingSystem\frontend
npm run build
```

## Search Receipt Module

Added sidebar page:

```text
Search Receipt
```

Purpose:

- Search by OR / receipt number.
- View receipt header and line details.
- Prepare restricted update workflow for only:
  - Assigned Collector
  - OR Receipt No.

Frontend:

```text
LGU_TreasuryReportingSystem\frontend\src\pages\SearchReceipt\SearchReceiptPage.jsx
LGU_TreasuryReportingSystem\frontend\src\App.jsx
LGU_TreasuryReportingSystem\frontend\src\App.css
```

Backend:

```text
LGU_TreasuryReportingSystem\backend\app\Services\SearchReceiptService.php
LGU_TreasuryReportingSystem\backend\app\Http\Controllers\Api\SearchReceiptController.php
LGU_TreasuryReportingSystem\backend\routes\api.php
```

Python runner:

```text
LGU_TreasuryReportingSystem\runner\search_receipt.py
```

API endpoints:

```text
GET   /api/search-receipts?receipt_no=ORNO&limit=100
GET   /api/search-receipts/{paymentId}
PATCH /api/search-receipts/{paymentId}
```

Search result fields include:

```text
Date | OR Receipt | Taxpayer | Assigned Collector | Status | Total | Action
```

Status mapping follows the Generate Receipt logic:

```text
PAYMENT.VOID_BV = 1 -> Void
PAYMENT.STATUS_CT = CNL or T_STATUS.DESCRIPTION contains CANCEL -> Cancelled
otherwise -> Paid
```

Write safety:

- Search and View are read-only.
- PATCH endpoint is restricted to `PAYMENT.COLLECTOR` and `PAYMENT.RECEIPTNO`.
- Actual Firebird update is disabled by default.
- To enable only after Treasurer approval and backup/testing:

```text
FIREBIRD_ALLOW_RECEIPT_UPDATE=1
```

Verification:

```text
python .\LGU_TreasuryReportingSystem\runner\search_receipt.py search --receipt-no 9808706 --limit 3
python .\LGU_TreasuryReportingSystem\runner\search_receipt.py update --payment-id 8BFD6D95-2B35-403A-A55E-A5B50D9FB345 --assigned-collector ricardo --new-receipt-no 9808706
php -l app\Services\SearchReceiptService.php
php -l app\Http\Controllers\Api\SearchReceiptController.php
php artisan route:list --path=search-receipts
npm run lint
npm run build
```

Observed search test:

```text
Receipt 9808706
Taxpayer: CADELINA PAZ
Assigned collector: ricardo
Status: Paid
Total: 30.00
```

Observed update guard test:

```text
updated: false
write_enabled: false
message: Receipt updates are disabled. Set FIREBIRD_ALLOW_RECEIPT_UPDATE=1 only after Treasurer approval.
```

Search Receipt performance update:

- Problem found: OR search was slow because the first implementation used:

```text
UPPER(TRIM(PAYMENT.RECEIPTNO)) CONTAINING UPPER(?)
```

and joined/grouped `PAYMENTDETAIL` immediately.

- Optimized runner:
  - exact `PAYMENT.RECEIPTNO = ?` lookup first so Firebird can use `IDX_PAYMENT_RECEIPTNO`
  - trimmed exact lookup second
  - `STARTING WITH` lookup third
  - `CONTAINING` only as final fallback
  - payment detail totals are fetched only for matched `PAYMENT_ID` rows
- Frontend now trims OR input and requests only 25 rows.
- Example OR `2019857` direct runner timing improved:

```text
Before: about 1.92 seconds
After: about 0.24 seconds
```

- Verified OR `2019857`:

```text
Taxpayer: MAQUILING, AQUILLA CAFINO
Assigned collector: flora
Status: Cancelled
Amount: 99.00
Detail: WATER PAYMENT FEE
```

## Income Target Module

Added sidebar page:

```text
Income Target
```

Source workbook:

```text
LGU_TreasuryReportingSystem\IncomeTarget\2026_Income_Target.xlsx
```

Workbook structure found:

```text
Sheet: Sheet1
Rows: 178
Columns: 3
Header: Particulars | Income Target ( Approved Budget)
```

Implemented read-only Excel reader:

```text
LGU_TreasuryReportingSystem\runner\income_target_readonly.py
```

Reader dependency note:

- The runner first uses `openpyxl` when that package is available.
- If Laravel launches a Python environment without `openpyxl`, the runner now falls back to a built-in read-only `.xlsx` parser using Python standard library `zipfile` + XML parsing.
- This fixes the backend error: `ModuleNotFoundError: No module named 'openpyxl'`.

The reader preserves the workbook hierarchy by converting leading spaces in the `Particulars` text into levels:

```text
level 0 = parent/major line
level 1 = child line
level 2 = deeper child line
```

Implemented backend:

```text
LGU_TreasuryReportingSystem\backend\app\Services\IncomeTargetService.php
LGU_TreasuryReportingSystem\backend\app\Http\Controllers\Api\IncomeTargetController.php
GET /api/income-target?year=2026
```

Implemented frontend:

```text
LGU_TreasuryReportingSystem\frontend\src\pages\IncomeTarget\IncomeTargetPage.jsx
LGU_TreasuryReportingSystem\frontend\src\App.jsx
LGU_TreasuryReportingSystem\frontend\src\App.css
```

Income Target page shows:

- Year selector
- Section filter
- Search by particulars
- Summary cards
- Paginated target table

Summary values parsed from the 2026 workbook:

```text
Tax Revenues: 8,798,637.00
Non-Tax Revenues: 17,432,735.00
Local Sources: 26,231,372.00
External Sources: 189,995,172.00
Total General Fund: 216,226,544.00
Total Special Education Fund: 1,962,402.00
Grand Total (GF + SEF): 218,188,946.00
```

Verification:

```text
python .\LGU_TreasuryReportingSystem\runner\income_target_readonly.py --year 2026
php -l app\Services\IncomeTargetService.php
php -l app\Http\Controllers\Api\IncomeTargetController.php
php artisan route:list --path=income-target
npm run lint
npm run build
```

Result:

```text
Python reader ok=true
Route available: GET api/income-target
Frontend lint passed
Frontend build passed
```

Income Target annual increase rule:

- The Treasurer workflow uses a 10% increase every year.
- If the selected year has its own workbook, the app reads the actual workbook values.
- If the selected year has no workbook, the app projects from the nearest earlier available workbook using compound growth:

```text
Projected Target = Source Workbook Target x (1.10 ^ years_from_source)
```

Example verified:

```text
2026 Grand Total actual: 218,188,946.00
2027 Grand Total projected: 240,007,840.60
```

Frontend now shows a note identifying whether the displayed target is actual workbook data or projected data.

## Dashboard - Collections vs Income Target

Dashboard was revised from a simple database-status screen into a Treasurer collection monitor.

Purpose:

- Show whether paid collections are hitting the Income Target.
- Use existing APIs only; no new backend endpoint was required.
- Compare paid local collections against the `Local Sources` income target, not the full grand target, because Report 21 treasury collections are local collection data while the grand target includes external sources such as NTA.

Data sources:

```text
GET /api/income-target?year=YYYY
GET /api/generated-reports/21/preview?date_from=YYYY-01-01&date_to=YYYY-MM-DD
GET /api/generated-reports/21/preview?date_from=YYYY-MM-01&date_to=YYYY-MM-DD
```

Dashboard metrics added:

- Local Collection Achievement percentage
- YTD Paid Collections
- Local Sources Target
- Current-month collections
- Monthly Target Pace
- Expected-to-date local target
- Variance against expected-to-date target
- Remaining local target
- Grand Income Target as context
- Collection share breakdown from Report 21 total row:
  - Municipal General Fund
  - Municipal SEF
  - Provincial Share
  - Barangay / Fisheries Share

Frontend files updated:

```text
LGU_TreasuryReportingSystem\frontend\src\pages\Dashboard\DashboardPage.jsx
LGU_TreasuryReportingSystem\frontend\src\App.css
```

Verification:

```text
cd LGU_TreasuryReportingSystem\frontend
npm run lint
npm run build
```

Result: lint passed and build passed. Existing Vite chunk-size warning may still appear but is not a compile error.

Dashboard follow-up update:

- Removed technical database cards from the Dashboard:
  - Firebird .FDB Status
  - User Tables
  - Views
  - Bridge Mode
- Added Treasurer-facing collection panels:
  - Collector Collection, showing top collectors for the selected year-to-date period
  - Dive Tickets Monthly Collection
  - Top 3 Dive Ticket Buyers
  - Collections vs Income Target by source group
- Source-group target comparison now covers:
  - Tax on Business
  - Regulatory Fees and Charges
  - Receipt from Economic Enterprise
  - Service/User Charges
  - Other Taxes
  - RPT
- Dashboard compares Report 21 paid YTD collection rows against matching Income Target rows.
- RPT dashboard actual is computed from Report 27 Summary Report Sharing, not gross RPT total.
- RPT dashboard rule:
  - Use only `BSC` / Basic RPT sharing.
  - Use Land Sharing + Building Sharing.
  - Current = current amount minus discount.
  - Prior = prior amount.
  - Penalties = current-year penalties + prior-year penalties.
  - Dashboard RPT value = 40% Municipal Share from Land + Building Sharing.
  - This follows the Treasurer rule shown in Report 27 that Basic RPT is split 35% Province, 40% Municipality, and 25% Barangay.
- Dive tickets needed a new read-only runner/report mode because existing General Fund source reports intentionally exclude `Diving Fee` as a trust/abstract-side item.

New read-only runner mode:

```text
python .\LGU_TreasuryReportingSystem\runner\general_fund_readonly.py dive-tickets --date-from YYYY-MM-DD --date-to YYYY-MM-DD
```

New API endpoint:

```text
GET /api/general-fund/dive-tickets?date_from=YYYY-MM-DD&date_to=YYYY-MM-DD
```

Files updated:

```text
LGU_TreasuryReportingSystem\runner\general_fund_readonly.py
LGU_TreasuryReportingSystem\backend\app\Http\Controllers\Api\GeneralFundController.php
LGU_TreasuryReportingSystem\backend\routes\api.php
LGU_TreasuryReportingSystem\frontend\src\pages\Dashboard\DashboardPage.jsx
LGU_TreasuryReportingSystem\frontend\src\App.css
```

January 2026 dive-ticket smoke test:

```text
Total: 109,500.00
Receipts: 19
Buyers: 17
Top buyers:
- THE BEACH HOUSES OF DAUIN: 15,000.00
- ATMOSPHERE RESORTS: 15,000.00
- SEA EXPLORER: 15,000.00
```

Verification:

```text
php -l LGU_TreasuryReportingSystem\backend\app\Http\Controllers\Api\GeneralFundController.php
php -l LGU_TreasuryReportingSystem\backend\routes\api.php
cd LGU_TreasuryReportingSystem\frontend
npm run lint
npm run build
```

Result: PHP syntax passed, frontend lint passed, and frontend build passed.

Dashboard chart/UI update:

- Added MUI-style dashboard visualization panels without adding a new chart dependency.
- Used existing `@mui/material` components plus custom SVG/CSS charts:
  - MUI `Paper`
  - MUI `Chip`
  - MUI `LinearProgress`
  - Custom SVG semi-gauge for Local Collection Achievement
  - Custom donut chart for Collection Share
  - Custom horizontal bar charts for Collector Collection and Dive Ticket Buyers
  - Source Group Target Gauge cards for category achievement
- Dashboard chart panels now include:
  - Annual target gauge
  - Collection Share donut chart
  - Collector Collection bar chart
  - Dive Ticket Buyers bar chart
  - Source Group Target Gauge grid
- Dive Tickets behavior:
  - `Dive Tickets Monthly Collection` uses the selected/current month.
  - `Top 3 Dive Ticket Buyers` uses the whole selected year, January 1 to December 31.
  - This lets the monthly panel show current activity while the buyer ranking shows the strongest annual dive-ticket customers.
- Dashboard `Collections vs Income Target by Source Group` table alignment:
  - Source column is left-aligned.
  - Actual, Target, Rate, and Variance headers and values are right-aligned.
  - Numeric values use tabular number styling so currency/rate columns line up cleanly.

Files updated:

```text
LGU_TreasuryReportingSystem\frontend\src\pages\Dashboard\DashboardPage.jsx
LGU_TreasuryReportingSystem\frontend\src\App.css
```

Verification:

```text
cd LGU_TreasuryReportingSystem\frontend
npm run lint
npm run build
```

Result: lint passed and build passed. Existing Vite chunk-size warning may still appear, but it is not a compile error.

## Portable Login Database For Cloned PCs

Problem found:

- Login failed after cloning the app to another PC because the backend auth SQLite database path was tied to one laptop:

```text
C:\Users\LIFT-LAPTOP\AppData\Local\Temp\lgu_treasury_auth.sqlite
```

Fix implemented:

- `backend\.env.example` now uses a portable SQLite setting:

```text
DB_CONNECTION=sqlite
DB_DATABASE=database.sqlite
SESSION_DRIVER=file
QUEUE_CONNECTION=sync
CACHE_STORE=file
```

- `backend\config\database.php` now resolves relative SQLite database names into:

```text
backend\database\database.sqlite
```

- New setup runner added:

```text
LGU_TreasuryReportingSystem\server_runner\setup_backend_auth_db.bat
```

Use on a newly cloned PC:

```text
server_runner\server_menu.bat
Option 1. Setup backend auth database
```

The setup creates/copies:

```text
backend\.env
backend\database\database.sqlite
```

Then runs:

```text
composer install   only if backend/vendor is missing
php artisan key:generate --force
php artisan config:clear
php artisan migrate --force
php artisan db:seed --force
```

Default test login seeded:

```text
Email: admin@zamboanguita.local
Password: admin123
```
