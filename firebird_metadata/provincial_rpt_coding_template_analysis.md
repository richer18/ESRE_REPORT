# Provincial RPT Coding Template Analysis

Date analyzed: 2026-06-18

File analyzed:

```text
report_template\PROVINCIAL_RPT_CODING_TEMPLATE.xlsx
```

Purpose understood:

```text
Monthly report on the collection of Real Property Tax by property classification.
Likely used as support for remittance/coding to the Provincial Government.
```

No template data was modified during this analysis.

## Workbook Summary

The workbook has two sheets:

```text
SEF
GF
```

Both sheets have the same structure:

```text
Rows 1-3   Report title and municipality
Row 5      For the month of
Row 7      Fund label
Row 8      Amount bucket headers
Rows 9-21  Property classification lines
Row 23     Subtotal formulas
Row 25     Total remittance formula
```

Formulas found in each sheet:

```text
C23 = SUM(C9:C21)
E23 = SUM(E9:E21)
G23 = SUM(G9:G21)
I23 = SUM(I9:I21)
I25 = SUM(C23:I23)
```

Interpretation of columns:

```text
A = Property classification label
B = Account/coding for current-year regular amount
C = Current-year regular amount
D = Account/coding for prior-year regular amount
E = Prior-year regular amount
F = Account/coding for current-year penalty
G = Current-year penalty amount
H = Account/coding for prior-year penalty
I = Prior-year penalty amount
```

The total formula in `I25` includes only numeric amount columns because the code columns are blank/non-numeric in the subtotal row.

## Sheet Purpose

### GF Sheet

The GF sheet appears to represent the Basic Real Property Tax portion.

Observed account/coding pattern:

```text
40102040-101-xx-xx = Basic RPT regular/current or prior year
40102040-102-xx-xx = Basic RPT penalty
```

Recommended Firebird tax type:

```text
PAYMENTCLASSDETAIL.ITAXTYPE_CT = 'BSC'
```

For provincial remittance/coding, this should follow the same sharing basis already used in Report 27:

```text
Report 27: Summary Report Sharing
Basic RPT / BSC Provincial Share = 35 percent
```

This is not a new assumption for Report 28. It comes from the existing Report 27 sharing template/formulas.

### SEF Sheet

The SEF sheet appears to represent the Special Education Fund portion.

Observed account/coding pattern:

```text
40102050-401-xx-xx = SEF regular/current or prior year
40102050-402-xx-xx = SEF penalty
```

Recommended Firebird tax type:

```text
PAYMENTCLASSDETAIL.ITAXTYPE_CT = 'SEF'
```

For provincial remittance/coding, this should follow the same sharing basis already used in Report 27:

```text
Report 27: Summary Report Sharing
SEF Provincial Share = 50 percent
```

This is not a new assumption for Report 28. It comes from the existing Report 27 sharing template/formulas.

## Relationship To Report 27

Report 28 should be based on:

```text
27. Summary Report Sharing
```

Reason:

```text
Report 27 already contains the validated sharing logic:
- BSC / Basic RPT: 35 percent Provincial, 40 percent Municipal, 25 percent Barangay
- SEF: 50 percent Provincial, 50 percent Municipal
```

Therefore, Report 28 should not compute sharing independently. It should reuse the same source buckets and sharing rules used by Report 27.

Practical relationship:

```text
Report 27 answers: How is RPT shared among Province, Municipality, Barangay?
Report 28 answers: How should the Provincial share be coded/remitted by account and property classification?
```

Recommended dependency:

```text
Report 28 should use the same Firebird extraction and row mapping foundation as Report 27:
- PAYMENT
- PAYMENTCLASSDETAIL
- RPTASSESSMENT
- PROPERTY
- PROPERTYKIND_CT
- CLASSCODE_CT
- TAXYEAR
- CASETYPE_CT
- ITAXTYPE_CT
```

Report 28 should then take only the Provincial share side:

```text
GF sheet:
  BSC provincial share from Report 27 logic.

SEF sheet:
  SEF provincial share from Report 27 logic.
```

## Required Reconciliation Chain

This report must not stand alone. The RPT reports should agree in this chain:

```text
21. Summary of Collection
23. Summary of Collection RPT
25. Record of Real Property Tax Collection
27. Summary Report Sharing
28. Provincial RPT Coding / Province Remittance Report
```

Meaning:

```text
Report 21 = overall collection control report.
Report 23 = RPT portion of collections.
Report 25 = detailed RPT receipt/property/tax-year/classification listing.
Report 27 = sharing computation for BSC and SEF.
Report 28 = provincial coding/remittance view of the Report 27 provincial share.
```

Recommended control relationship:

```text
Report 21 RPT lines
  should agree with Report 23 RPT totals.

Report 23 RPT totals
  should agree with Report 25 total RPT collection, after using the same gross/net treatment.

Report 25 Basic/BSC and SEF totals
  should agree with the raw BSC/SEF buckets used in Report 27.

Report 27 Provincial Share
  should agree with Report 28 GF and SEF remittance/coding totals.
```

If these reports do not match, investigate before preparing remittance:

```text
1. Date range mismatch
2. Voided payment included/excluded differently
3. Cancelled PAYMENTCLASSDETAIL row included/excluded differently
4. Gross versus net of discount mismatch
5. Current year versus prior year bucket mismatch
6. Penalty current/prior bucket mismatch
7. Property classification mapping mismatch
8. Special/Scientific/S* mapping mismatch
```

## Firebird Tables Involved

Primary source tables:

```text
PAYMENT
PAYMENTCLASSDETAIL
```

Supporting RPT/property tables:

```text
RPTASSESSMENT
PROPERTY
T_CLASSIFICATION
T_ACTUALUSE
T_PROPERTYKIND
T_BARANGAY
```

Important fields:

```text
PAYMENT.PAYMENT_ID
PAYMENT.PAYMENTDATE
PAYMENT.PAYGROUP_CT
PAYMENT.RECEIPTNO
PAYMENT.VOID_BV

PAYMENTCLASSDETAIL.PAYMENT_ID
PAYMENTCLASSDETAIL.TAXTRANS_ID
PAYMENTCLASSDETAIL.ITAXTYPE_CT
PAYMENTCLASSDETAIL.CASETYPE_CT
PAYMENTCLASSDETAIL.TAXYEAR
PAYMENTCLASSDETAIL.AMOUNT
PAYMENTCLASSDETAIL.CLASSCODE_CT
PAYMENTCLASSDETAIL.PROPERTYKIND_CT
PAYMENTCLASSDETAIL.ACTUALUSE_CT
PAYMENTCLASSDETAIL.CANCELLED_BV

RPTASSESSMENT.TAXTRANS_ID
RPTASSESSMENT.PREDOMCLASSCODE_CT
RPTASSESSMENT.PROP_ID

PROPERTY.PROP_ID
PROPERTY.PROPERTYKIND_CT
```

Recommended base filters:

```sql
WHERE p.PAYMENTDATE >= CAST(:date_from AS DATE)
  AND p.PAYMENTDATE < DATEADD(1 DAY TO CAST(:date_to AS DATE))
  AND p.PAYGROUP_CT = 'RPT'
  AND COALESCE(p.VOID_BV, 0) = 0
  AND COALESCE(pcd.CANCELLED_BV, 0) = 0
```

## RPT Amount Logic

Use `PAYMENTCLASSDETAIL.AMOUNT`, not only `PAYMENTDETAIL.AMOUNT`.

Reason:

```text
The province template needs amounts by property classification.
PAYMENTCLASSDETAIL contains the classification, property kind, actual use, tax year, and amount breakdown.
```

Recommended amount buckets:

```text
Current Year Regular:
  TAXYEAR = report year
  CASETYPE_CT not in ('DED', 'PEN')

Prior Year Regular:
  TAXYEAR <> report year
  CASETYPE_CT not in ('DED', 'PEN')

Discount:
  CASETYPE_CT = 'DED'
  Usually negative in the database.
  It should reduce the regular/current or prior-year bucket.

Current Year Penalty:
  TAXYEAR = report year
  CASETYPE_CT = 'PEN'

Prior Year Penalty:
  TAXYEAR <> report year
  CASETYPE_CT = 'PEN'
```

Recommended net calculation:

```text
Regular amount bucket = REG amount + DED amount
Penalty amount bucket = PEN amount
```

The existing RPT report logic treats discount as a reduction from gross collection. For this province remittance template, discount should not be remitted as a positive amount unless the Province explicitly requires gross-before-discount reporting.

## Property Classification Codes Found

Firebird `T_CLASSIFICATION` values found:

```text
A    = AGRICULTURAL
C    = COMMERCIAL
I    = INDUSTRIAL
M    = MINERAL
R    = RESIDENTIAL
T    = TIMBERLAND/FOREST
S*   = Special classifications
SS   = SCIENTIFIC
SC   = CULTURAL
SED  = SPECIAL EDUCATION
SGOV = SPECIAL GOVERNMENT
SH   = HOSPITAL
SRD  = SPECIAL ROADLOT
SRE  = SPECIAL RELIGIOUS
SW   = LOCAL WATER DISTRICT
```

Firebird `T_PROPERTYKIND` values found:

```text
B = BUILDING
L = LAND
M = MACHINERIES
P = IMPROVEMENTS
```

Firebird `T_ACTUALUSE` includes:

```text
AA   = AGRICULTURAL
AC   = COMMERCIAL
AI   = INDUSTRIAL
AM   = MINERAL
AR   = RESIDENTIAL
ARC  = RECREATION
ATF  = TIMBERLAND/FOREST
ASS  = SCIENTIFIC
ASED = EDUCATIONAL
```

Important issue:

```text
The template says "Machinery", but Firebird classification code M means MINERAL.
Firebird property kind M means MACHINERIES.
This must be clarified before automating the report.
```

## Template Row Mapping

Observed template rows:

```text
9   Land Residential
10  Land Commercial
11  Land Industrial
12  Land Machinery
13  Land Agricultural
14  Land Recreational
15  Land-TIMBER
16  Building Residential
17  Building Commercial
18  Building Industrial
19  Building Machinery
20  Building Agricultural
21  Building Recreational
```

Suggested first-pass mapping:

```text
Land Residential     = PROPERTYKIND_CT = L and actual use/class Residential
Land Commercial      = PROPERTYKIND_CT = L and actual use/class Commercial
Land Industrial      = PROPERTYKIND_CT = L and actual use/class Industrial
Land Machinery       = needs confirmation: may mean Mineral, not machinery
Land Agricultural    = PROPERTYKIND_CT = L and actual use/class Agricultural
Land Recreational    = PROPERTYKIND_CT = L and ACTUALUSE_CT = ARC/Recreation
Land-TIMBER          = PROPERTYKIND_CT = L and class/use Timberland/Forest

Building Residential  = PROPERTYKIND_CT = B and actual use/class Residential
Building Commercial   = PROPERTYKIND_CT = B and actual use/class Commercial
Building Industrial   = PROPERTYKIND_CT = B and actual use/class Industrial
Building Machinery    = needs confirmation: may mean machinery or mineral
Building Agricultural = PROPERTYKIND_CT = B and actual use/class Agricultural
Building Recreational = PROPERTYKIND_CT = B and ACTUALUSE_CT = ARC/Recreation
```

Open mapping questions:

```text
1. Where should SPECIAL/SCIENTIFIC amounts go?
2. Should S* classifications be grouped under Recreational, Special, Industrial/Special, or a new row?
3. Should property kind M/MACHINERIES be mapped to Building Machinery, Land Machinery, or a separate machinery row?
4. Should property kind P/IMPROVEMENTS be included with Building rows or excluded?
5. The Land-TIMBER row has no account codes in the uploaded template. Ask Province for the proper codes before using it.
```

## January 2026 Test Findings

Read-only aggregate checks against the Firebird database for 2026-01-01 to 2026-01-31 found:

```text
SPECIAL/SCIENTIFIC exists in January 2026 RPT data.
Example class code: SS
Property kind: B
Gross BSC: 520.40
BSC discount: 104.08
Gross SEF: 520.40
SEF discount: 104.08
```

This confirms the template needs a clear rule for Special/Scientific before totals can be trusted.

The same January 2026 data also includes:

```text
Property kind M = MACHINERIES
Property kind P = IMPROVEMENTS
```

The template does not clearly provide separate rows for these groups.

## Recommended Process

Recommended monthly process for this provincial remittance report:

```text
1. Select report month/date range.
2. Pull all non-void RPT payments from PAYMENT.
3. Join PAYMENTCLASSDETAIL to get tax type, case type, tax year, class, actual use, property kind, and amount.
4. Exclude cancelled class detail rows.
5. Split tax type:
   - BSC goes to GF sheet.
   - SEF goes to SEF sheet.
6. Split amount bucket:
   - current year regular
   - prior year regular
   - current year penalty
   - prior year penalty
7. Apply discount as reduction to the matching regular bucket.
8. Apply province share using Report 27 sharing rules:
   - BSC/GF provincial share = 35 percent.
   - SEF provincial share = 50 percent.
9. Map each row to the template property classification line.
10. Fill only amount cells C/E/G/I.
11. Keep account/coding columns B/D/F/H from the template.
12. Recompute row 23 and row 25.
13. Reconcile:
   - GF report total should match Report 27 BSC Provincial Share.
   - SEF report total should match Report 27 SEF Provincial Share.
14. Investigate any unmapped property kind/class/use before submission.
```

## Recommended Reconciliation Checks

Before using the report for remittance:

```text
1. Report 21 RPT lines should reconcile to Report 23 RPT totals.
2. Report 23 RPT totals should reconcile to Report 25 RPT totals.
3. Total BSC gross/net from PAYMENTCLASSDETAIL should reconcile to Report 25 Basic totals.
4. Total SEF gross/net from PAYMENTCLASSDETAIL should reconcile to Report 25 SEF totals.
5. Report 25 Basic/SEF totals should reconcile to the raw BSC/SEF buckets used in Report 27.
6. Province share calculation should reconcile to Report 27:
   - Basic/BSC Provincial Share
   - SEF Provincial Share
7. Sum of Report 28 template amount cells should equal the Report 27 Provincial Share totals.
8. Unmapped classes should be listed separately for review.
```

## Draft SELECT-Only Query Logic

This is draft logic for future automation. It is not yet implemented as a report runner.

```sql
SELECT
    CASE
        WHEN pcd.ITAXTYPE_CT = 'BSC' THEN 'GF'
        WHEN pcd.ITAXTYPE_CT = 'SEF' THEN 'SEF'
        ELSE pcd.ITAXTYPE_CT
    END AS SHEET_NAME,
    COALESCE(pcd.PROPERTYKIND_CT, prop.PROPERTYKIND_CT) AS PROPERTY_KIND,
    COALESCE(pcd.CLASSCODE_CT, ra.PREDOMCLASSCODE_CT) AS CLASS_CODE,
    pcd.ACTUALUSE_CT,
    pcd.ITAXTYPE_CT,
    pcd.CASETYPE_CT,
    pcd.TAXYEAR,
    SUM(pcd.AMOUNT) AS AMOUNT
FROM PAYMENT p
JOIN PAYMENTCLASSDETAIL pcd
  ON pcd.PAYMENT_ID = p.PAYMENT_ID
LEFT JOIN RPTASSESSMENT ra
  ON ra.TAXTRANS_ID = pcd.TAXTRANS_ID
LEFT JOIN PROPERTY prop
  ON prop.PROP_ID = ra.PROP_ID
WHERE p.PAYMENTDATE >= CAST(:date_from AS DATE)
  AND p.PAYMENTDATE < DATEADD(1 DAY TO CAST(:date_to AS DATE))
  AND p.PAYGROUP_CT = 'RPT'
  AND COALESCE(p.VOID_BV, 0) = 0
  AND COALESCE(pcd.CANCELLED_BV, 0) = 0
  AND pcd.ITAXTYPE_CT IN ('BSC', 'SEF')
GROUP BY
    CASE
        WHEN pcd.ITAXTYPE_CT = 'BSC' THEN 'GF'
        WHEN pcd.ITAXTYPE_CT = 'SEF' THEN 'SEF'
        ELSE pcd.ITAXTYPE_CT
    END,
    COALESCE(pcd.PROPERTYKIND_CT, prop.PROPERTYKIND_CT),
    COALESCE(pcd.CLASSCODE_CT, ra.PREDOMCLASSCODE_CT),
    pcd.ACTUALUSE_CT,
    pcd.ITAXTYPE_CT,
    pcd.CASETYPE_CT,
    pcd.TAXYEAR
ORDER BY 1, 2, 3, 4, 5, 6, 7;
```

## Recommended Next Step

Before coding the exporter, use Report 27 as the sharing authority and ask Province/Treasurer/Accounting to confirm the remaining mapping rules:

```text
1. Where should SPECIAL/SCIENTIFIC/S* classifications be placed?
2. What should "Machinery" mean in this template: class M/Mineral, property kind M/Machineries, or both?
3. What account codes should be used for Land-TIMBER?
4. How should property kind P/Improvements be handled?
5. Should the Provincial RPT Coding report reconcile directly to Report 27 provincial-share totals before printing/submission?
```

Once confirmed, this can become:

```text
28. Provincial RPT Coding / Province Remittance Report
```
