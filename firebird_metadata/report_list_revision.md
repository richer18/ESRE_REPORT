# Report List Revision Notes

Scope: Reports 1 to 27 in `run_collection_query.py --list`.

Revision date: 2026-06-17

## New Finding

Some RPT receipts can contain one `PAYMENT_ID` and one `TAXTRANS_ID` but multiple property classifications in `PAYMENTCLASSDETAIL.CLASSCODE_CT`.

Observed example for January 2026:

- Property kind: `B` / Building
- Classification code: `SS` / Scientific
- Report grouping: `SPECIAL`
- BSC gross: `520.40`
- BSC discount: `104.08`
- BSC net: `416.32`
- SEF gross: `520.40`
- SEF discount: `104.08`
- SEF net: `416.32`
- Grand net: `832.64`

The previous report 25 grouping used only `PAYMENT_ID + TAXTRANS_ID`, so the Scientific/Special amount could be hidden under the first classification found for the same TD/ARP.

## Applied Fix

Reports with property classification detail now use:

```text
PAYMENT_ID + TAXTRANS_ID + normalized property classification
```

Special classification normalization:

```text
CLASSCODE_CT starting with S => SPECIAL
Examples: S, SC, SED, SGOV, SH, SRD, SRE, SS, SW
```

This keeps mixed-class TD/ARP payments visible as separate report rows while keeping the total collection reconciled to the `.FDB`.

## Report Impact Review

1. Unified collection detail: unaffected. Uses `PAYMENT` + `PAYMENTDETAIL`; no property classification output.
2. RPT payment detail with posting/account context: informational SQL only. It exposes `PAYMENTCLASSDETAIL.CLASSCODE_CT`; users should group by class code if doing class analysis.
3. RPT totals by Basic/SEF/other RPT line classification: unaffected. Totals by tax type/fund, not property classification.
4. CTC payment detail: unaffected.
5. CTC totals: unaffected.
6. Other Fees and Charges payment detail: unaffected.
7. Other Fees and Charges totals by revenue code: unaffected.
8. Daily collections: unaffected.
9. Monthly collections: unaffected.
10. Quarterly collections: unaffected.
11. Yearly collections: unaffected.
12. Receipt range and header-vs-detail reconciliation: unaffected.
13. Other Fees and Charges tax/rate list: unaffected.
14. Other Fees and Charges parent-child hierarchy with rates: unaffected.
15. Total collection per collector: unaffected.
16. Fees collected by selected collector: unaffected.
17. Sources of Collections summary: unaffected.
18. RPT payment process flow: revised understanding; class-level RPT reports must use `PAYMENTCLASSDETAIL`.
19. CTC payment process flow: unaffected.
20. Other Fees and Charges process flow: unaffected.
21. Summary of Collection: unaffected by property classification display; it uses RPT summary buckets.
22. Summary of Collection no RPT: unaffected.
23. Summary of Collection RPT: unaffected by property classification display; it uses Basic/SEF and property kind summary buckets.
24. Summary in RPT based on SUMMARY layout: unaffected by property classification display; it uses RPT summary buckets.
25. Record of Real Property Tax Collection: fixed. Now splits rows by normalized property classification.
26. Record of Real Property Tax Collection - Advance Payment Report: fixed. Now splits rows by normalized property classification.
27. Summary Report Sharing: reviewed. `S*` class codes are grouped into `SPECIAL`; building-side special/industrial values are placed in `BLDG-INDUS/SPECIAL` because the uploaded template has no separate building-special row.

## Verification

January 2026 report 25:

- Rows exported after fix: `4707`
- `SPECIAL` row found: yes
- Special BSC gross/discount: `520.40 / 104.08`
- Special SEF gross/discount: `520.40 / 104.08`
- Report total: `4,820,216.46`
- Firebird total: `4,820,216.46`
- Difference: `0.00`

January 2026 report 27:

- `BLDG-INDUS/SPECIAL` BSC gross/discount: `520.40 / 104.08`
- `BLDG-INDUS/SPECIAL` SEF gross/discount: `520.40 / 104.08`
- BSC net from inputs: `2,410,108.23`
- SEF net from inputs: `2,410,108.23`
- Grand net from inputs: `4,820,216.46`

January 2026 report 26:

- Rows exported: `0`
- Reason: advance-payment rule is `TAXYEAR > report year`; January 2026 had no 2027+ RPT advance-payment rows found.
