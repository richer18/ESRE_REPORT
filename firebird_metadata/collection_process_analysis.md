# Collection Process Analysis

Scope: Real Property Tax Payment, Community Tax / Cedula Payment, and Other Fees and Charges.

Inspection basis: Firebird metadata plus read-only checks of classification fields. No schema or data was modified.

## Shared Cashiering Layer

These three processes converge in the same cashiering tables:

- `PAYMENT`: payment header / official receipt header.
- `PAYMENTDETAIL`: receipt line items by income/tax type and fund type.
- `PAYMENTCLASSDETAIL`: extra RPT classification breakdown by tax declaration/class/use/year.
- `PAYMENTCHEQUE`: cheque/check details when payment mode requires it.
- `TAXPAYER`: taxpayer/payor master, joined through `LOCAL_TIN`.
- `T_ITAXTYPE`: revenue/tax type lookup, joined through `PAYMENTDETAIL.ITAXTYPE_CT`.
- `T_FUNDTYPE`: fund lookup, joined through `PAYMENTDETAIL.FUNDTYPE_CT`.
- `T_PAYMODE`: payment mode lookup, joined through `PAYMENT.PAYMODE_CT`.
- `T_STATUS`: payment status lookup, joined through `PAYMENT.STATUS_CT`.
- `RCDCTRLNUMBER`, `RCDORBOOKLET`, `RCDACCOUNTABLEFORM`, `RCDCASHBREAKDOWN`, `RCDFUNDTYPEBREAKDOWN`, `RCDDEPOSIT`: collection reporting, OR booklet, accountable form, cash breakdown, fund breakdown, and deposit controls.

Important shared fields:

- `PAYMENT.PAYMENT_ID`: primary key of the payment header.
- `PAYMENT.RECEIPTNO`: official receipt number.
- `PAYMENT.PAYMENTDATE`: collection date used for daily/monthly/quarterly/yearly reporting.
- `PAYMENT.AMOUNT`: receipt/header amount.
- `PAYMENT.PAYGROUP_CT`: broad payment group. Observed groups include `RPT`, `OTH`, `TOB`, and `EE`.
- `PAYMENT.LOCAL_TIN`, `PAYMENT.PAIDBY`, `PAYMENT.COLLECTOR`, `PAYMENT.USERID`.
- `PAYMENT.RCDNUMBER`, `PAYMENT.ORB_ID`, `PAYMENT.AFTYPE`: RCD / OR booklet / accountable form references.
- `PAYMENT.VOID_BV`: likely void flag. Queries defensively filter `COALESCE(VOID_BV, 0) = 0`.
- `PAYMENTDETAIL.AMOUNTPAID`: line-level amount; safest basis for revenue totals.
- `PAYMENTDETAIL.ITAXTYPE_CT`: revenue/tax code.
- `PAYMENTDETAIL.FUNDTYPE_CT`: fund code.
- `PAYMENTDETAIL.SOURCE_CT`, `PAYMENTDETAIL.SOURCEID`: source module/source record pointer.
- `PAYMENTDETAIL.DEBITPOSTINGID`: link to debit/account posting, important for RPT.

## Real Property Tax Payment

Likely involved tables:

- Core payment: `PAYMENT`, `PAYMENTDETAIL`, `PAYMENTCLASSDETAIL`, `PAYMENTCHEQUE`.
- RPT account/posting: `TPACCOUNT`, `POSTINGJOURNAL`, `MANUALDEBIT`.
- Property and assessment: `PROPERTY`, `PROPERTYOWNER`, `PROPERTYOWNERDETAIL`, `RPTASSESSMENT`, `RPTASSESSMENTDETAIL`, `RPTLIABILITY`, `RPTTRANSACTION`.
- RPT supporting appraisal/status tables: `RPTLANDAPPRAISAL`, `RPTBLDGINFO`, `RPTBLDGFLOOR`, `RPTMACHAPPRAISAL`, `RPTPLANTTREEAPPRAISAL`, `RPTCANCELLED`, `RPTCOMPROMISE`.

Important fields:

- `PAYMENT.PAYGROUP_CT = 'RPT'`: primary RPT filter.
- `PAYMENT.RECEIPTNO`, `PAYMENT.PAYMENTDATE`, `PAYMENT.AMOUNT`.
- `PAYMENTDETAIL.ITAXTYPE_CT`: observed RPT-related codes include `BSC` for RPT Basic and `SEF` for Special Education Fund; `RPT` also exists as a parent/group code.
- `PAYMENTDETAIL.AMOUNTPAID`: paid amount by RPT component.
- `PAYMENTDETAIL.DEBITPOSTINGID`: expected link to `TPACCOUNT.POSTING_ID`.
- `TPACCOUNT.TAXTRANS_ID`, `PROP_ID`, `TAXYEAR`, `TAXPERIOD_CT`, `CASETYPE_CT`, `EARMARK_CT`.
- `PAYMENTCLASSDETAIL.TAXTRANS_ID`, `CLASSCODE_CT`, `PROPERTYKIND_CT`, `ACTUALUSE_CT`, `TAXYEAR`, `AMOUNT`.
- `POSTINGJOURNAL.RPTTAXDUE`, `POSTINGJOURNAL.SEFTAXDUE`, `POSTINGJOURNAL.POSTED_BV`.

Important RPT classification note:

- A single `PAYMENT_ID` and `TAXTRANS_ID` can contain multiple `PAYMENTCLASSDETAIL.CLASSCODE_CT` values.
- Detail reports that show property classification must group by `PAYMENT_ID + TAXTRANS_ID + CLASSCODE_CT` or a normalized classification bucket.
- Special classifications are normalized as `CLASSCODE_CT` starting with `S` = `SPECIAL`, including codes such as `SS` / Scientific.
- Without this classification-level grouping, Special/Scientific amounts can be hidden under another classification for the same TD/ARP.

How payment is recorded:

- One `PAYMENT` row stores the receipt header.
- One or more `PAYMENTDETAIL` rows store paid amounts by income/tax/fund code.
- RPT payments are identified from `PAYMENT.PAYGROUP_CT = 'RPT'`.
- `PAYMENTDETAIL.DEBITPOSTINGID` points toward `TPACCOUNT.POSTING_ID`, connecting payment lines to tax-year/property/account postings.
- `PAYMENTCLASSDETAIL` provides the class/use/tax-year breakdown when the same receipt covers multiple assessed property classes or years.

Receipt numbers:

- Stored in `PAYMENT.RECEIPTNO`.
- `PAYMENT.ORB_ID` links the receipt to OR booklet controls.
- `PAYMENT.RCDNUMBER` links the payment to RCD reporting/control.

Totals:

- Use `SUM(PAYMENTDETAIL.AMOUNTPAID)` for collection reports.
- Use `PAYMENT.AMOUNT` for header reconciliation.
- Use `PAYMENTCLASSDETAIL.AMOUNT` for RPT class/use/tax-year allocation.
- Use `TPACCOUNT` or `POSTINGJOURNAL` only when analyzing liabilities/postings, not as the primary cash collection total.

Triggers/procedures:

- `PAYMENT_LASTEDITED`, `PAYMENTDETAIL_LASTEDITED`, `PAYMENTCLASSDETAIL_LASTEDITED`: set `DATALASTEDITED` for non-`SYSDBA` users.
- `DELETE_PAYMENT`, `DELETE_PAYMENTDETAIL`, `DELETE_PAYMENTCLASSDETAIL`: write delete audit rows to `DELETE_AUDITTRAIL`.
- `POSTINGJOURNAL_LASTEDITED`, `MANUALDEBIT_LASTEDITED`, RPT table `*_LASTEDITED` triggers: audit/update stamping.
- `PROC_RPTDEBITPOSTING`: materially affects RPT posting. It reads unposted `POSTINGJOURNAL` rows, inserts `TPACCOUNT` debit rows for `BSC` and `SEF`, and marks `POSTINGJOURNAL.POSTED_BV = 1`.

## Community Tax / Cedula Payment

Likely involved tables:

- Core payment: `PAYMENT`, `PAYMENTDETAIL`.
- Cedula certificate: `COMMUNITYTAXCERTIFICATE`.
- Other LGU CTC: `CTCOTHERLGU`.
- Rate table: `T_CTCRATE`.
- Taxpayer: `TAXPAYER`.

Important fields:

- `PAYMENTDETAIL.SOURCE_CT IN ('CTCI', 'CTCC')`: observed CTC source markers.
- `PAYMENTDETAIL.ITAXTYPE_CT = 'CTC'`: CTC revenue code fallback.
- `PAYMENTDETAIL.SOURCEID`: likely points to `COMMUNITYTAXCERTIFICATE.CTC_ID` for CTC detail rows.
- `COMMUNITYTAXCERTIFICATE.CTC_ID`, `CTCNO`, `DATEISSUED`, `CTCTYPE`, `CTCYEAR`, `LOCAL_TIN`.
- `COMMUNITYTAXCERTIFICATE.BASICTAXDUE`, `BUSTAXDUE`, `SALTAXDUE`, `RPTAXDUE`, `INTEREST`, `TOTALAMOUNTPAID`.

How payment is recorded:

- Receipt header is in `PAYMENT`.
- Paid line amount is in `PAYMENTDETAIL.AMOUNTPAID`.
- Certificate detail is in `COMMUNITYTAXCERTIFICATE`.
- The relationship appears to be `PAYMENTDETAIL.SOURCEID = COMMUNITYTAXCERTIFICATE.CTC_ID` when `SOURCE_CT` is `CTCI` or `CTCC`.

Receipt numbers:

- Official receipt number remains `PAYMENT.RECEIPTNO`.
- Cedula certificate number is separate: `COMMUNITYTAXCERTIFICATE.CTCNO`.

Totals:

- Cash collection total: `SUM(PAYMENTDETAIL.AMOUNTPAID)`.
- Certificate recorded total: `SUM(COMMUNITYTAXCERTIFICATE.TOTALAMOUNTPAID)`.
- Component recomputation: `BASICTAXDUE + BUSTAXDUE + SALTAXDUE + RPTAXDUE + INTEREST`.

Triggers/procedures:

- `TD_CTC`: stamps `COMMUNITYTAXCERTIFICATE.TRANSDATE = CURRENT_TIMESTAMP`.
- No CTC-specific stored procedure was found.
- Shared payment/payment-detail audit triggers also apply.

## Other Fees and Charges

Likely involved tables:

- Core payment: `PAYMENT`, `PAYMENTDETAIL`, `PAYMENTCHEQUE`.
- Business and permits: `BUSINESSPERMIT`, `BUSINESSPERMITDETAIL`, `BUSINESSPERMITJOURNAL`, `BUSINESSPERMITCLEARANCE`, `BUSINESSPERMITPROMISSORY`.
- Civil/local fees: `BURIALPERMIT`, `MARRIAGELICENSE`, `CASHTICKET`, `CATTLEOWNERSHIP`, `CATTLETRANSFER`.
- Economic enterprise / enforcement: `EEBILL`, `EECLEARANCE`, `EECITATION`, `EECITATIONTICKETDETAIL`, `EECONTRACT`.
- MCH/tricycle/franchise-related: `MCHCHARGES`, `MCHDRIVER`, `MCHOPERATOR`, `MCHREQUIREMENT`, `MCHOFFENSE`.
- Water-related fees: `WATERBILL`, `WATERBILLDETAIL`, `WATERCONTRACT`, `WM_READING`, `WM_WC`.
- Rates/lookups: `T_OTHERPAYMENTRATE`, `T_ITAXTYPE`, `T_FUNDTYPE`, `T_BPADDONCHARGES`, `T_BPCATEGORY`, `T_EERATE`, `T_MCHRATE`, `T_WATERITEM`.

Important fields:

- Use everything not classified as RPT or CTC:
  - `COALESCE(PAYMENT.PAYGROUP_CT, '') <> 'RPT'`
  - not `PAYMENTDETAIL.SOURCE_CT IN ('CTCI','CTCC')`
  - not `PAYMENTDETAIL.ITAXTYPE_CT = 'CTC'`
- `PAYMENT.PAYGROUP_CT`: observed `OTH`, `TOB`, `EE`, and null.
- `PAYMENTDETAIL.ITAXTYPE_CT`: revenue classification.
- `PAYMENTDETAIL.FUNDTYPE_CT`: fund classification.
- `PAYMENTDETAIL.SOURCE_CT`, `SOURCEID`: module/source record pointer.
- `PAYMENTDETAIL.UNIT`, `AMOUNTPAID`: quantity/rate-derived amount when applicable.

How payment is recorded:

- One `PAYMENT` header per receipt.
- One or more `PAYMENTDETAIL` rows by revenue code.
- Some source modules maintain their own permit/bill/certificate rows, but collection reporting should start from `PAYMENT` plus `PAYMENTDETAIL`.

Receipt numbers:

- Official receipt number is `PAYMENT.RECEIPTNO`.
- Booklet/RCD references are `PAYMENT.ORB_ID` and `PAYMENT.RCDNUMBER`.

Totals:

- Use `SUM(PAYMENTDETAIL.AMOUNTPAID)` grouped by `ITAXTYPE_CT`, `FUNDTYPE_CT`, `PAYGROUP_CT`, or source module.
- Use `PAYMENT.AMOUNT` only for receipt-level reconciliation.
- For business permit computations, source tables such as `BUSINESSPERMITDETAIL.TAXDUE`, `BUSINESSPERMITJOURNAL.TAXAMOUNT`, `SURCHARGE`, and `INTEREST` are assessment/billing inputs, not necessarily the cash collection total.

Triggers/procedures:

- Business permit tables have `TD_*` timestamp triggers and several `*_BI` generator triggers.
- `BURIALPERMIT_BI` generates `BUR_ID`.
- `TD_BUSINESSPERMIT`, `TD_BPDETAIL`, `TD_BPJOURNAL`, `TD_EECLEARANCE`, `TD_WATERBILL`, `TD_WATERBILLDETAIL`: timestamp updates.
- Shared `PAYMENT*` triggers apply for payment audit/delete behavior.
- No stored procedure specifically for other-fee collection posting was found in the five user procedures.

## Query Pack

Use `collection_analysis_queries.sql` in this folder. It contains SELECT-only queries for:

- Unified collection detail.
- RPT details and RPT totals.
- CTC details and CTC totals.
- Other fees and charges details and totals.
- Daily, monthly, quarterly, yearly summaries.
- Receipt/header vs detail reconciliation.
