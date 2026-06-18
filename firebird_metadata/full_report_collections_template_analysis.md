# Full Report Collections Template Analysis

Date analyzed: 2026-06-18

File analyzed:

```text
report_template\FULL_REPORT_COLLECTIONS.xlsx
```

No workbook data or formulas were modified during this analysis.

## Summary

This template appears to be a monthly top-level collection control report.

Implemented report number:

```text
31. Full Report Collections
```

The workbook has one sheet:

```text
Sheet1
```

The template is compact:

```text
Rows 1-2   Title
Rows 4-5   Month and year labels
Row 7      Column headers
Rows 8-31  Daily rows
Row 32     Monthly totals
Rows 36-38 RCD total, less due from, total collections
```

## Template Layout

Observed columns:

```text
A = DATE
B = CTC
C = RPT
D = GF AND TF
E = DUE FROM
F = RCD TOTAL
```

Observed formulas:

```text
B32 = SUM(B8:B31)
C32 = SUM(C8:C31)
D32 = SUM(D8:D31)
E32 = SUM(E8:E31)
F32 = SUM(F8:F31)
F36 = F32
F37 = E32
F38 = F36 - F37
```

Interpretation:

```text
RCD TOTAL          = total daily collections recorded in RCD/payment records.
LESS: DUE FROM    = collections included in RCD total but deducted from LGU actual total.
TOTAL COLLECTIONS = RCD total less due from.
```

The `DUE FROM` column is manual. The generated workbook should leave it editable/blank unless the user enters values.

## Relationship To Existing Reports

This report should reconcile with existing reports:

```text
19. Process flow: Community Tax Certificate / Cedula payment
21. Summary of Collection
22. Summary of Collection no rpt
23. Summary of Collection rpt
25. Record of Real Property Tax Collection
29. Abstract of General Collections
30. Abstract of Trust Funds Collections
31. Full Report Collections
```

Recommended control chain:

```text
CTC column
  should agree with Community Tax / Cedula collections for each day.

RPT column
  should agree with RPT collection totals for each day.

GF AND TF column
  should agree with non-RPT/non-CTC General Fund and Trust Fund collections.

RCD TOTAL column
  should agree with CTC + RPT + GF AND TF + DUE FROM, if DUE FROM is part of gross RCD.

TOTAL COLLECTIONS
  should agree with the LGU-owned net collection total after deducting DUE FROM.
```

## Recommended Firebird Source

Likely source tables:

```text
PAYMENT
PAYMENTDETAIL
PAYMENTCLASSDETAIL
COMMUNITYTAXCERTIFICATE
T_ITAXTYPE
T_FUNDTYPE
RCDCTRLNUMBER
RCDFUNDTYPEBREAKDOWN
```

Main fields:

```text
PAYMENT.PAYMENTDATE
PAYMENT.PAYMENT_ID
PAYMENT.PAYGROUP_CT
PAYMENT.RECEIPTNO
PAYMENT.RCDNUMBER
PAYMENT.AMOUNT
PAYMENT.VOID_BV

PAYMENTDETAIL.ITAXTYPE_CT
PAYMENTDETAIL.FUNDTYPE_CT
PAYMENTDETAIL.SOURCE_CT
PAYMENTDETAIL.SOURCEID
PAYMENTDETAIL.AMOUNTPAID

PAYMENTCLASSDETAIL.ITAXTYPE_CT
PAYMENTCLASSDETAIL.AMOUNT
PAYMENTCLASSDETAIL.CANCELLED_BV
```

Recommended base filters:

```sql
WHERE p.PAYMENTDATE >= CAST(:date_from AS DATE)
  AND p.PAYMENTDATE < DATEADD(1 DAY TO CAST(:date_to AS DATE))
  AND COALESCE(p.VOID_BV, 0) = 0
```

## Proposed Column Logic

### CTC

Use Community Tax / Cedula payment lines:

```text
PAYMENTDETAIL.SOURCE_CT IN ('CTCI', 'CTCC')
OR PAYMENTDETAIL.ITAXTYPE_CT = 'CTC'
```

Daily total:

```text
SUM(PAYMENTDETAIL.AMOUNTPAID)
```

### RPT

Use RPT payment group:

```text
PAYMENT.PAYGROUP_CT = 'RPT'
```

Recommended source:

```text
PAYMENTCLASSDETAIL for Basic/SEF detail, excluding cancelled class rows.
```

Daily total should reconcile with:

```text
23. Summary of Collection RPT
25. Record of Real Property Tax Collection
27. Summary Report Sharing
```

### GF AND TF

Use non-RPT and non-CTC payment detail rows:

```text
PAYMENT.PAYGROUP_CT <> 'RPT'
AND NOT (PAYMENTDETAIL.SOURCE_CT IN ('CTCI', 'CTCC') OR PAYMENTDETAIL.ITAXTYPE_CT = 'CTC')
```

This should include:

```text
General Fund collections
Trust Fund collections
Other fees and charges
Business tax
Permits
Service/user charges
Economic enterprise
```

It should reconcile with:

```text
22. Summary of Collection no rpt
29. Abstract of General Collections
30. Abstract of Trust Funds Collections
```

### DUE FROM

This column is manually encoded.

Implementation rule:

```text
Do not compute DUE FROM from Firebird.
Leave daily DUE FROM cells blank or zero for manual entry.
Keep the template formulas so TOTAL COLLECTIONS updates after manual entry.
```

### RCD TOTAL

Recommended logic:

```text
RCD TOTAL = CTC + RPT + GF AND TF
```

Then the template computes:

```text
TOTAL COLLECTIONS = RCD TOTAL - manual DUE FROM
```

## Remaining Open Questions Before Coding

```text
1. Should RCD TOTAL use PAYMENT.AMOUNT, PAYMENTDETAIL.AMOUNTPAID, or RCD fund breakdown tables?
2. Should CTC be gross receipt amount or LGU share only?
3. Should RPT be gross RPT, municipal share, or total Basic+SEF collection?
4. Should GF AND TF include Trust Fund gross amounts or only LGU-owned portions after splits?
5. Should rows 8-31 represent calendar days 1-24/31, or only days with collections?
```

## Implemented Process

Implemented in `run_collection_query.py`:

```text
1. Aggregate daily CTC from PAYMENTDETAIL source markers CTCI/CTCC or ITAXTYPE_CT = CTC.
2. Aggregate daily RPT from PAYMENTCLASSDETAIL, excluding cancelled class rows.
3. Aggregate daily GF AND TF from non-RPT and non-CTC PAYMENTDETAIL rows.
4. Leave DUE FROM blank/manual.
5. Write RCD TOTAL as the daily sum of CTC + RPT + GF AND TF.
6. Preserve monthly total formulas and TOTAL COLLECTIONS = RCD TOTAL - DUE FROM.
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
```
