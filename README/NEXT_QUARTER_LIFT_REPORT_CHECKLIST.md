# Next Quarter LIFT / eSRE Reporting Checklist

Use this checklist before encoding, while encoding, and before final submission of SRE / SRS / SOE reports.

## 1. Before Encoding

Confirm the reporting period:

```text
Year:
Quarter:
Date prepared:
Prepared by:
```

Check if the report period is open in LIFT:

```text
Report Period Maintenance
Status must be Open
```

Prepare source documents from each office.

## 2. Documents From Budget Office

Ask Budget Office for:

```text
[ ] Annual Budget / Appropriation Ordinance
[ ] Supplemental Budget Ordinance, if any
[ ] Realignment / Augmentation document, if any
[ ] Continuing Appropriation list
[ ] Budget Appropriation per office
[ ] Budget Appropriation per function
[ ] Budget Appropriation per PPA
[ ] Budget Appropriation per allotment class
```

Breakdown must show:

```text
[ ] Fund: General Fund / SEF / Trust Fund / SHF, if applicable
[ ] Office
[ ] Function
[ ] Program / Project / Activity
[ ] NGAS code, if available
[ ] PS
[ ] MOOE
[ ] FE
[ ] CO
```

Check if there is Supplemental Budget:

```text
[ ] Supplemental Budget encoded in LIFT
[ ] Correct office
[ ] Correct fund
[ ] Correct function
[ ] Correct PPA
[ ] Correct PS / MOOE / FE / CO
```

## 3. Documents From Accounting Office

Ask Accounting Office for:

```text
[ ] Actual Expenditures per quarter
[ ] Disbursement summary
[ ] DV summary
[ ] Check disbursement summary
[ ] Accounts Payable list
[ ] Prior-year AP payments
[ ] Debt service actual payments
[ ] Continuing Appropriation actual expenditures
```

Breakdown must show:

```text
[ ] Fund: General Fund / SEF / Trust Fund / SHF, if applicable
[ ] Office
[ ] Function
[ ] Program / Project / Activity
[ ] NGAS code
[ ] PS
[ ] MOOE
[ ] FE
[ ] CO
```

## 4. Documents From Treasurer's Office

Prepare Treasurer data:

```text
[ ] Collections / receipts summary
[ ] RPT collection
[ ] CTC / Cedula collection
[ ] Business tax collection
[ ] Other fees and charges
[ ] Trust fund receipts
[ ] Debt service payment confirmation
[ ] Cash beginning balance
[ ] Cash ending balance
[ ] Bank balance / cash in bank
[ ] SRE-NGAS reconciliation data
```

From the Firebird reporting tool, generate as needed:

```powershell
python .\run_collection_query.py --list
python .\run_collection_query.py 21 YYYY-MM-DD YYYY-MM-DD
python .\run_collection_query.py 22 YYYY-MM-DD YYYY-MM-DD
python .\run_collection_query.py 23 YYYY-MM-DD YYYY-MM-DD
python .\run_collection_query.py 25 YYYY-MM-DD YYYY-MM-DD
python .\run_collection_query.py 26 YYYY-MM-DD YYYY-MM-DD
python .\run_collection_query.py 27 YYYY-MM-DD YYYY-MM-DD
```

## 5. LIFT Encoding Areas

Budget Office related:

```text
[ ] Authorized Budget > Budget Appropriation > Expenditures
[ ] Authorized Budget > Budget Appropriation > Debt Services
[ ] Authorized Budget > Budget Appropriation > Unappropriated Surplus
```

Treasurer receipts:

```text
[ ] Actual Transaction > Receipts > Real Property Tax
[ ] Actual Transaction > Receipts > General Collections
[ ] Actual Transaction > Receipts > Trust Fund Receipts
```

Expenditure / accounting related:

```text
[ ] Actual Transaction > Expenditures > Expenditures
[ ] Actual Transaction > Expenditures > Accounts Payable
[ ] Actual Transaction > Expenditures > Debt Services
[ ] Actual Transaction > Expenditures > Trust Fund Expenditures
```

Other Treasurer checks:

```text
[ ] Actual Transaction > Others > Fund/Cash Balance
[ ] Actual Transaction > Others > SRE-NGAS Reconciliation
[ ] Actual Transaction > Others > Financial Operations
```

## 6. SRS Checks

Check Statement of Receipts Sources:

```text
[ ] RPT Basic agrees with municipal GF share
[ ] SEF agrees with municipal SEF share
[ ] Business Tax agrees with collection summary
[ ] Community Tax agrees with CTC report
[ ] Permits and Licenses agree with Other Fees report
[ ] Service/User Charges agree with Other Fees report
[ ] Economic Enterprises agree with Other Fees report
[ ] National Tax Allotment agrees with official NTA release
[ ] Grand Total GF + SEF agrees with SRE receipts
```

Remember:

```text
SRS = detailed receipts/source report
```

## 7. SOE Checks

Check Statement of Expenditures:

```text
[ ] Budget Appropriation encoded
[ ] Supplemental Budget encoded, if any
[ ] Actual Expenditures encoded
[ ] Debt Service encoded
[ ] Accounts Payable encoded
[ ] Continuing Appropriation encoded
[ ] PS/MOOE/FE/CO classification correct
[ ] Office/function/PPA correct
[ ] GF/SEF fund classification correct
```

Watch for:

```text
[ ] Actual Expenditures > Budget Appropriation
[ ] Negative balances
[ ] Wrong fund
[ ] Wrong allotment class
[ ] Missing supplemental budget
[ ] Missing continuing appropriation
[ ] Duplicate AP payment
```

Remember:

```text
SOE = detailed expenditure/use report
```

## 8. SRE Checks

Check Statement of Receipts and Expenditures:

```text
[ ] SRE receipts agree with SRS grand total
[ ] SRE expenditures agree with SOE total expenditures
[ ] Beginning cash balance agrees with previous ending balance
[ ] Ending fund/cash balance is not negative
[ ] Fund/Cash Balance End agrees with Fund/Cash Balance module
[ ] National Tax Allotment agrees with official source
[ ] Non-income receipts checked
[ ] Non-operating expenditures checked
[ ] AP payments checked
[ ] Continuing appropriations checked
```

Remember:

```text
SRE = summary of receipts, expenditures, and fund/cash balance
```

## 9. If SOE Or Fund/Cash Balance Is Negative

Do not guess. Check in this order:

```text
1. Exact row with negative
2. Budget appropriation amount
3. Actual expenditure amount
4. Supplemental Budget
5. Realignment / augmentation
6. Fund: GF or SEF
7. Allotment class: PS / MOOE / FE / CO
8. Office / function / PPA / NGAS code
9. Accounts Payable
10. Continuing Appropriation
11. Beginning Cash Balance
12. Ending Cash Balance / bank balance
```

Ask offices:

```text
Budget Office:
- Please verify appropriation, supplemental budget, and continuing appropriation.

Accounting Office:
- Please verify actual expenditures, disbursements, AP, and debt service classification.

Treasurer's Office:
- Verify receipts, cash balance, bank balance, and SRE/SRS/SOE reconciliation.
```

## 10. Before Submission

Final checks:

```text
[ ] SRS generated and reviewed
[ ] SOE generated and reviewed
[ ] SRE generated and reviewed
[ ] Alerts checked
[ ] Negative values explained or corrected
[ ] Budget Office confirmed budget figures
[ ] Accounting Office confirmed expenditure figures
[ ] Treasurer confirmed receipts and cash balance
[ ] Certified correct by authorized officer
[ ] Backup/export saved
```

Save copies:

```text
[ ] SRE Excel/PDF
[ ] SRS Excel/PDF
[ ] SOE Excel/PDF
[ ] Supporting collection reports
[ ] Budget support files
[ ] Accounting support files
[ ] Screenshots of resolved alerts, if needed
```

## 11. Simple Reminder

```text
Treasurer = receipts, cash, fund balance, report checking
Budget = appropriations, supplemental budget, continuing appropriation
Accounting = expenditures, AP, disbursements, debt service classification
```

If a negative appears:

```text
It is usually a reconciliation issue, not automatically a Treasurer mistake.
```
