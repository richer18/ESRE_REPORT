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
```

Reports 1 to 20 mostly export CSV from SELECT-only SQL.

Reports 21 to 23 generate Excel files using uploaded Summary of Collection templates.

Report 24 exports CSV RPT detail summary.

Reports 25 to 27 generate Excel files using uploaded RPT templates.

Reports 28 to 30 are analyzed/proposed template reports and are not yet implemented in the runner.

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

## Proposed Report 28: Provincial RPT Coding / Province Remittance Report

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

Do not automate this report until these mapping issues are confirmed:

```text
1. Where SPECIAL/SCIENTIFIC/S* classifications should be placed.
2. Whether "Machinery" in the template means class M/Mineral or property kind M/Machineries.
3. How to handle property kind P/Improvements.
4. The Land-TIMBER row has no account/coding values in the uploaded template.
5. Report 28 should reconcile to Report 27 Provincial Share totals before printing/submission.
```

## Proposed Reports 29-30: Abstract Collection Templates

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
