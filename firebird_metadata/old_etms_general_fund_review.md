# Old ETMS General Fund Module Review

Source reviewed read-only:

```text
C:\Users\LIFT-LAPTOP\OneDrive\Desktop\ElectronicTreasurerManagementSystem\frontend\src\FRONTEND\components\ABSTRACT\GF
```

Related components reviewed:

```text
src\components\MD-Components\FillupForm\AbstractGF.jsx
src\components\MD-Components\FillupForm\GeneralFundPaymentEditForm.jsx
src\components\MD-Components\Popup\GeneralFundDialog.jsx
src\FRONTEND\components\ABSTRACT\shared\ReceiptCollectionReportDialog.jsx
```

## General Fund UI Structure

The old General Fund module is organized around one large `GeneralFund.jsx` page plus smaller table/report components under `TableData`.

Main page responsibilities:

- Month/year filtering.
- Search.
- Add payment dialog.
- Edit payment dialog.
- Delete payment action.
- View payment detail dialog.
- Daily report view.
- Financial report view.
- Receipt/collector report dialog.
- Dashboard total cards.
- Excel download.

Important subcomponents:

- `TableData\DailyTable.jsx` - daily grouped General Fund collection view.
- `TableData\ReportTable.jsx` - monthly financial/SOC-style General Fund summary.
- `TableData\GenerateReport.jsx` - collector receipt report wrapper for GF.
- `TableData\GeneralFundAllTable.jsx` - payment detail breakdown table.
- `TableData\components\Table\TotalReport.jsx` - all GF source totals.
- `TableData\components\Table\TaxOnBusiness.jsx` - Tax on Business breakdown.
- `TableData\components\Table\RegulatoryFees.jsx` - Regulatory Fees breakdown.
- `TableData\components\Table\ServiceUserCharges.jsx` - Service/User Charges breakdown.
- `TableData\components\Table\ReceiptsFromEconomicEnterprise.jsx` - Economic Enterprise breakdown.

## Important Old API Endpoints

The old frontend depends on these MySQL-oriented endpoints:

```text
generalFundDataAll
general-fund-dashboard-summary
generalFundDataReport
allDataGeneralFund
viewalldataGeneralFundTableView
generalFundPaymentRates
form-types
taxpayers
generalFundPayment
generalFundPaymentEdit/{id}
generalFundPaymentView/{id}
deleteGF/{id}
generate-report
getGFComments/{date}
commentGFCounts
insertGFComment
updateGFComment
general-fund-total-tax-report
general-fund-tax-on-business-report
general-fund-regulatory-fees-report
general-fund-service-user-charges
general-fund-receipts-from-economic-enterprise-report
```

For the new Firebird-first system, do not copy these endpoints directly. Use them as workflow names only, then rebuild Laravel API endpoints against Firebird/Python runner data.

## Payment Entry Workflow

`AbstractGF.jsx` shows the old GF payment process:

1. User selects payment date.
2. User selects or types taxpayer name.
3. Taxpayer lookup can return `ownerName` and `localTin`.
4. User enters receipt number.
5. User selects receipt type from `form-types`.
6. User selects cashier/collector.
7. User selects one or more GF source/rate descriptions from `generalFundPaymentRates`.
8. User enters amount per selected source.
9. Total is computed client-side as sum of selected source amounts.
10. Payload is posted to `generalFundPayment`.

Payload shape:

```json
{
  "date": "YYYY-MM-DD",
  "name": "taxpayer name",
  "receipt_no": "receipt number",
  "type_receipt": "receipt type code",
  "cashier": "collector/cashier",
  "local_tin": "optional local tin",
  "details": [
    {
      "source_id": "rate/source id",
      "description": "source description",
      "amount": 0
    }
  ]
}
```

Cash Tickets behavior:

- If selected source description is `Cash Tickets`, receipt number is auto-generated as `00YYYYMMDD`.
- This is a local workflow rule and must be verified against the Firebird/LIFT source before adopting in the new system.

## Payment Edit Workflow

`GeneralFundPaymentEditForm.jsx` shows:

- Load existing payment header from `generalFundPaymentEdit/{id}`.
- Header fields: `PAYMENTDATE`, `PAIDBY`, `RECEIPTNO`, `AFTYPE`, `COLLECTOR` or `USERID`.
- Detail fields: `PAYMENTDETAIL_ID`, `SOURCEID`, `description`, `amount`.
- Total is recomputed client-side from detail amounts.
- User can change receipt info, cashier, source description, add/remove source rows, and update amounts.

This should not be implemented as write-back to Firebird in the new app yet. For reporting app phase, use the pattern only for viewing/validation.

## Report Workflow

Main General Fund report categories:

- Total Revenue
- Tax on Business
- Regulatory Fees
- Receipts from Economic Enterprises
- Service/User Charges

The monthly financial report maps backend fields into report rows:

```text
Manufacturing
Distributor
Retailing
Financial
Other_Business_Tax
Sand_Gravel
Fines_Penalties
Mayors_Permit
Weighs_Measure
Tricycle_Operators
Occupation_Tax
Cert_of_Ownership
Cert_of_Transfer
Cockpit_Prov_Share
Cockpit_Local_Share
Docking_Mooring_Fee
Sultadas
Miscellaneous_Fee
Reg_of_Birth
Marriage_Fees
Burial_Fees
Correction_of_Entry
Fishing_Permit_Fee
Sale_of_Agri_Prod
Sale_of_Acct_Form
Water_Fees
Stall_Fees
Cash_Tickets
Slaughter_House_Fee
Rental_of_Equipment
Doc_Stamp
Police_Report_Clearance
Secretaries_Fee
Med_Dent_Lab_Fees
Garbage_Fees
Cutting_Tree
total
```

Notable calculation:

```text
reportMunicipalTotal = total - cockpitProvShare
```

This matters for LGU reporting because cockpit provincial share is separated from municipal GF total.

## Daily Collection Workflow

`DailyTable.jsx` uses:

- Month/year filter.
- `allDataGeneralFund` for daily grouped rows.
- `viewalldataGeneralFundTableView` to show details for a selected date.
- `Overall Total` for daily total calculation.
- Comment system per date: `getGFComments/{date}`, `commentGFCounts`, `insertGFComment`, `updateGFComment`.

Useful idea for new system:

- A daily GF page should show date, receipt range/count, collector/source breakdown, daily total, and optional notes.

## Collector Receipt Report Workflow

`GenerateReport.jsx` wraps shared `ReceiptCollectionReportDialog`.

It supports:

- Date range filter.
- Month/year filter.
- Collector filter.
- Receipt number range filter.
- Report type filter fixed to `GF`.
- Total collection computed from returned row totals.

Payload to old endpoint:

```json
{
  "dateType": "dateRange or monthYear",
  "dateFrom": "YYYY-MM-DD or YYYY-MM",
  "dateTo": "YYYY-MM-DD or YYYY",
  "reportType": "GF",
  "cashier": "collector code",
  "orFrom": "optional receipt from",
  "orTo": "optional receipt to"
}
```

Collector options used in GF:

```text
FLORA MY
IRIS
AGNES
RICARDO
AMABELLA
```

Special display rule:

```text
If cashier is AMABELLA and report_type is GF, display receipt type as Cash Tickets.
Otherwise GF is displayed as General Fund.
```

## Recommendation For New System

For `LGU_TreasuryReportingSystem`, build the new General Fund module by folder:

```text
frontend/src/pages/GeneralFund/
frontend/src/pages/GeneralFund/components/
frontend/src/pages/GeneralFund/data/
frontend/src/pages/GeneralFund/hooks/
```

Recommended frontend pages/components:

- `GeneralFundPage.jsx` - page shell and filters.
- `components/GeneralFundSummaryCards.jsx`.
- `components/GeneralFundCollectionTable.jsx`.
- `components/GeneralFundDailyTable.jsx`.
- `components/GeneralFundSourceBreakdown.jsx`.
- `components/GeneralFundReceiptReportDialog.jsx`.
- `components/GeneralFundPaymentDetailDialog.jsx`.
- `data/generalFundSources.js`.
- `hooks/useGeneralFundCollections.js`.

Recommended new Laravel API endpoints:

```text
GET /api/general-fund/summary
GET /api/general-fund/collections
GET /api/general-fund/daily
GET /api/general-fund/sources
GET /api/general-fund/payment/{id}
GET /api/general-fund/receipt-report
```

Backend should be read-only against Firebird/Python runner first. Do not implement add/edit/delete until the reporting layer is stable and write ownership is approved.
