# Abstract Collection Templates Analysis

Date analyzed: 2026-06-18

Files analyzed:

```text
report_template\ABSTRACT_OF_GENERAL_COLLECTIONS.xlsx
report_template\ABSTRACT_OF_TRUST_FUNDS_COLLECTIONS.xlsx
```

No workbook data or template formulas were modified during this analysis.

## Summary

Two new reporting templates were reviewed:

```text
29. Abstract of General Collections
30. Abstract of Trust Funds Collections
```

Both templates contain:

```text
data
daily_collection
```

The likely workflow is:

```text
1. Fill the data sheet with receipt-level collection rows.
2. Summarize the same values by date in the daily_collection sheet.
3. Reconcile totals to Summary of Collection / no-RPT reports.
4. Print or submit the abstract.
```

## Abstract Of General Collections

Workbook:

```text
ABSTRACT_OF_GENERAL_COLLECTIONS.xlsx
```

Sheets:

```text
data
daily_collection
```

The workbook has no formulas in the uploaded version.

### data Sheet

Observed headers:

```text
A   Date
B   Receipt Number
C   Names
D   Manufacturing
E   Distributor
F   Retailing
G   Financial
H   Other
I   Sand & Gravel
J   Fines & Penalties
K   Mayor's Permit
L   W. & M.
M   Trirycle Operators
N   Occu.
O   Cert. of Ownership
P   Cert. of Transfer
Q   Cockpit Prov. Share
R   Cockpit Local Share
S   Docking and Mooring Fee
T   Sultadas
U   MISCS.
V   Reg. of
W   Marriage Fees
X   Burial Fees
Y   Correction of Entry
Z   Fishing Permit Fee
AA  Sale of Agri. Prod.
AB  Sale of Acct. Form
AC  Water Fees
AD  Stall Fees
AE  Cash Tickets
AF  Slaughter House Fee
AG  Rental of Equipment
AH  Doc. Stamp
AI  Police Clearance
AJ  Cert.
AK  Med./Dent. & Lab. Fees
AL  Garbage Fees
AM  CASHIER
AN  TOTAL
AO  TYPE OF RECIEPT
```

Notes:

```text
The label "Reg. of" appears truncated and likely means Registration of Birth.
The label "Cert." likely means Secretary's Fees / certification fees.
The template spelling has "RECIEPT"; keep template spelling unless the office wants it corrected.
```

### daily_collection Sheet

Observed headers:

```text
A   Date
B-AJ Same collection source columns as data sheet, excluding receipt/name/cashier/type fields
AK  TOTAL
```

Purpose:

```text
Daily totals by source of collection.
```

### Recommended Firebird Source

Use:

```text
PAYMENT
PAYMENTDETAIL
T_ITAXTYPE
T_FUNDTYPE
```

Recommended base filters:

```sql
WHERE p.PAYMENTDATE >= CAST(:date_from AS DATE)
  AND p.PAYMENTDATE < DATEADD(1 DAY TO CAST(:date_to AS DATE))
  AND COALESCE(p.VOID_BV, 0) = 0
  AND NOT (pd.ITAXTYPE_CT IN ('BSC', 'SEF') OR p.PAYGROUP_CT = 'RPT')
  AND NOT (pd.SOURCE_CT IN ('CTCI', 'CTCC') OR pd.ITAXTYPE_CT = 'CTC')
```

Recommended fields:

```text
PAYMENT.PAYMENTDATE      -> Date
PAYMENT.RECEIPTNO        -> Receipt Number
PAYMENT.PAIDBY           -> Names
PAYMENT.COLLECTOR/USERID -> CASHIER
PAYMENT.PAYGROUP_CT      -> TYPE OF RECEIPT fallback
PAYMENTDETAIL.ITAXTYPE_CT
PAYMENTDETAIL.SOURCE_CT
PAYMENTDETAIL.SOURCEID
PAYMENTDETAIL.AMOUNTPAID
```

### General Collections Mapping

Recommended mapping follows the existing `classify_summary_source()` logic in `run_collection_query.py`.

```text
Manufacturing              MAS
Distributor                WHO
Retailing                  RET
Financial                  BFI
Other                      OBT, CIC, PED, EMD, excluding source IDs 807/808
Sand & Gravel              TSG, TSB
Fines & Penalties          FPT and related fines/penalty codes
Mayor's Permit             MP
W. & M.                    FWM
Trirycle Operators         TOP, MTO, FLF
Occu.                      OCC
Cert. of Ownership         COO
Cert. of Transfer          COT
Cockpit Prov. Share        CS split 50 percent
Cockpit Local Share        CS split 50 percent
Docking and Mooring Fee    FRF with source IDs 580/639
Sultadas                   ST, ATM
MISCS.                     IM, OPF, SBF, or source IDs 807/808
Reg. of                    RB
Marriage Fees              RM
Burial Fees                BF
Correction of Entry        CE
Fishing Permit Fee         FRF, IF
Sale of Agri. Prod.        IAP
Sale of Acct. Form         IAF
Water Fees                 WTR, IWO
Stall Fees                 RFM, MSF
Cash Tickets               Needs confirmation; no stable code observed yet
Slaughter House Fee        SPF, RFS
Rental of Equipment        RFR, IPG, ICO
Doc. Stamp                 SF with source ID 810
Police Clearance           PCL
Cert.                      SF, HEC, OCL, excluding Doc. Stamp source ID rule
Med./Dent. & Lab. Fees     MDL
Garbage Fees               GCF
```

Open items:

```text
1. Confirm Cash Tickets code/source rule.
2. Confirm whether Community Tax Certificate is intentionally excluded from this abstract.
3. Confirm whether Trust Fund-coded items such as Building, Electrical, Zoning, Livestock, and Diving are intentionally excluded and handled in the Trust Fund abstract.
4. Confirm if daily_collection should be formula-driven from data or generated directly from Firebird.
```

## Abstract Of Trust Funds Collections

Workbook:

```text
ABSTRACT_OF_TRUST_FUNDS_COLLECTIONS.xlsx
```

Sheets:

```text
data
daily_collection
```

The workbook has no formulas in the uploaded version.

### data Sheet

Observed headers:

```text
A   Date
B   Receipt Number
C   Names
D   Building Fee - 80% Local
E   Building Fee - 15% T.F.
F   Building Fee - 5% Nat'L.
G   Electrical Fee
H   Zoning Fee
I   Livestock Dev. Fund - 80% Local
J   Livestock Dev. Fund - 20% Nat'l
K   Diving Fee - 40% GF
L   Diving Fee - 30% Fishers
M   Diving Fee - 30% Brgy
N   CASHIER
O   Total
P   TYPE OF RECIEPT
```

### daily_collection Sheet

Observed headers:

```text
A   DATE
B-D Building split: 80% Local, 15% T.F., 5% Nat'L.
E   Electrical fee
F   Zoning Fee
G-H Livestock split: 80% Local, 20% Nat'l
I-K Diving split: 40% GF, 30% Fishers, 30% Brgy
L   Total
```

Purpose:

```text
Daily totals for trust fund / shared collections.
```

### Recommended Firebird Source

Use:

```text
PAYMENT
PAYMENTDETAIL
T_ITAXTYPE
T_FUNDTYPE
```

Relevant codes found in Firebird:

```text
PFB = Permit Fees under Bldg. Code
BUF = Building Permit
INS = Inspection Fees
EP  = Electrical Permit Fees
ZLC = Zonal/Location Clearance Fees
IFL = Income from Livestock
IFD = Income from Diving
```

Observed fund type:

```text
PFB/BUF/INS = TF
EP          = TF
ZLC         = TF
IFL         = TF
IFD         = TF
```

### Trust Fund Split Rules

The template matches the split rules already used in `split_summary_amount()`:

```text
Building Permit Fee:
  80 percent Local
  15 percent Trust Fund
  5 percent National

Livestock:
  80 percent Local
  20 percent National

Diving Fee:
  40 percent General Fund
  30 percent Fishers
  30 percent Barangay
```

Electrical and Zoning:

```text
Electrical Fee appears as one amount column.
Zoning Fee appears as one amount column.
No split was shown in the uploaded template.
```

### Trust Fund Mapping

Recommended column mapping:

```text
Building Fee gross = PFB + BUF + INS
  D = gross x 80 percent
  E = gross x 15 percent
  F = gross x 5 percent

Electrical Fee = EP
  G = gross

Zoning Fee = ZLC
  H = gross

Livestock Dev. Fund = IFL
  I = gross x 80 percent
  J = gross x 20 percent

Diving Fee = IFD
  K = gross x 40 percent
  L = gross x 30 percent
  M = gross x 30 percent
```

Recommended total:

```text
O = SUM(D:M)
```

Open items:

```text
1. Confirm whether Electrical Fee and Zoning Fee should remain unsplit.
2. Confirm whether Building Fee should include all PFB, BUF, and INS codes.
3. Confirm if TYPE OF RECIEPT should use PAYMENT.PAYGROUP_CT, T_FUNDTYPE, or a custom label.
4. Confirm if Trust Fund daily_collection should be formula-driven from data or generated directly from Firebird.
```

## Reconciliation Chain

These abstracts should reconcile with existing reports:

```text
21. Summary of Collection
22. Summary of Collection no rpt
17. Sources of Collections summary
29. Abstract of General Collections
30. Abstract of Trust Funds Collections
```

Recommended control checks:

```text
Report 29 general collection total
  should reconcile to the matching non-RPT/non-CTC General Fund source totals in Reports 21/22/17.

Report 30 trust fund collection total
  should reconcile to the trust fund/source totals in Reports 21/22/17 after applying the same split rules.

Receipt-level totals in each data sheet
  should equal the sum of source columns for the same receipt.

Daily totals in daily_collection
  should equal grouped totals from the data sheet by date.
```

## Recommended Implementation Process

Recommended future automation:

```text
1. Add Report 29 for Abstract of General Collections.
2. Add Report 30 for Abstract of Trust Funds Collections.
3. Reuse `classify_summary_source()` for source-to-column mapping.
4. Reuse `split_summary_amount()` for split/shared collection rules.
5. Build receipt-level rows first in `data`.
6. Build date-level totals in `daily_collection`.
7. Reconcile generated totals to Reports 21/22/17 before printing.
```

Do not automate until the open mapping items above are confirmed.
