import argparse
import json
import os
import re
import sys
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path


DEFAULT_DB_PATHS = [
    r"main-server:i_tax046zamboanguita",
    r"C:\ZAMBOANGUITA_DB\ZAMBOANGUITA.FDB",
]
DEFAULT_FB_CLIENT_PATHS = [
    r"C:\Program Files\Firebird\Firebird_2_5\bin\fbclient.dll",
]
DEFAULT_ODBC_DSN = "itaxzamboanguita"
OUTPUT_DIR = Path(__file__).resolve().parent / "firebird_metadata" / "output"

COLLECTOR_ALIASES = {
    "iris": "angelique",
    "iris arbolado": "angelique",
    "angelique iris": "angelique",
    "flora my": "flora",
}


def parse_date(value):
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value or ""):
        raise argparse.ArgumentTypeError("Date must use YYYY-MM-DD format.")
    return value


def normalize_collector(value):
    value = (value or "").strip()
    return COLLECTOR_ALIASES.get(value.lower(), value)


def is_firebird_server_path(value):
    value = (value or "").strip()
    if not value:
        return False
    if re.match(r"^[A-Za-z]:[\\/]", value):
        return False
    return ":" in value


def db_path_candidates():
    env_path = os.environ.get("ESRE_FIREBIRD_DB")
    if env_path:
        return [env_path]
    return list(DEFAULT_DB_PATHS)


def resolve_fb_client():
    env_path = os.environ.get("ESRE_FIREBIRD_CLIENT")
    candidates = [env_path] if env_path else []
    candidates.extend(DEFAULT_FB_CLIENT_PATHS)
    for candidate in candidates:
        candidate = (candidate or "").strip()
        if candidate and Path(candidate).exists():
            return candidate
    return candidates[0] if candidates else DEFAULT_FB_CLIENT_PATHS[0]


def resolve_odbc_dsn():
    return os.environ.get("ESRE_ODBC_DSN", DEFAULT_ODBC_DSN).strip()


def open_firebird_native(user, password):
    import fdb

    fb_client = resolve_fb_client()
    last_error = None
    for db_path in db_path_candidates():
        db_path = (db_path or "").strip()
        if not db_path:
            continue
        if not is_firebird_server_path(db_path) and not Path(db_path).exists():
            continue
        try:
            return fdb.connect(
                dsn=db_path,
                user=user,
                password=password,
                charset="WIN1252",
                fb_library_name=fb_client,
                isolation_level=fdb.ISOLATION_LEVEL_READ_COMMITED_RO,
                no_db_triggers=True,
                no_gc=True,
            )
        except Exception as exc:
            last_error = exc
    if last_error is not None:
        raise last_error
    raise RuntimeError("No usable Firebird database path was found.")


def open_firebird_odbc(user, password):
    import pyodbc

    dsn = resolve_odbc_dsn()
    fb_client = resolve_fb_client()
    connection_string = f"DSN={dsn};UID={user};PWD={password};CLIENT={fb_client};"
    return pyodbc.connect(connection_string, autocommit=False)


def open_connection(mode, user, password):
    mode = (mode or "odbc").strip().lower()
    if mode == "native":
        return open_firebird_native(user, password)
    if mode == "odbc":
        return open_firebird_odbc(user, password)
    if mode == "auto":
        try:
            return open_firebird_native(user, password)
        except Exception:
            return open_firebird_odbc(user, password)
    raise RuntimeError(f"Unsupported connection mode: {mode}")


def to_json_value(value):
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def money(value):
    if value is None:
        return Decimal("0")
    if isinstance(value, Decimal):
        return value
    return Decimal(str(value))


def numeric_or(value):
    text = str(value or "").strip()
    digits = re.sub(r"\D", "", text)
    if not digits:
        return None
    return int(digits)


def receipt_in_range(receipt_no, or_from, or_to):
    if not or_from and not or_to:
        return True
    receipt_value = numeric_or(receipt_no)
    from_value = numeric_or(or_from) if or_from else None
    to_value = numeric_or(or_to) if or_to else None
    if receipt_value is None:
        return False
    if from_value is not None and receipt_value < from_value:
        return False
    if to_value is not None and receipt_value > to_value:
        return False
    return True


def payment_status_label(status_ct, void_bv):
    status = (status_ct or "").strip().upper()
    if void_bv:
        return "VOID"
    if status in ("VOID", "VOI"):
        return "VOID"
    if status in ("CNL", "CAN", "CNC", "CANCEL", "CANCELLED"):
        return "CANCELLED"
    return "PAID"


def classify_summary_source(itaxtype, source_id, source_ct):
    code = (itaxtype or "").strip()
    source_ct = (source_ct or "").strip()
    try:
        source_id = int(source_id) if source_id is not None else None
    except (TypeError, ValueError):
        source_id = None
    if source_ct in ("CTCI", "CTCC") or code == "CTC":
        return "Community Tax"
    if code == "MAS":
        return "Manufacturing"
    if code == "WHO":
        return "Distributor"
    if code == "RET":
        return "Retailing"
    if code == "BFI":
        return "Banks & Other Financial Int."
    if code in ("OBT", "CIC", "PED", "EMD") and source_id not in (807, 808):
        return "Other Business Tax"
    if code in ("TSG", "TSB"):
        return "Sand & Gravel"
    if code == "FPT":
        return "Fines & Penalties"
    if code == "MP":
        return "Mayor's Permit"
    if code == "FWM":
        return "Weights & Measures"
    if code in ("TOP", "MTO", "FLF"):
        return "Tricycle Permit Fee"
    if code == "OCC":
        return "Occupation Tax"
    if code == "COO":
        return "Cert. of Ownership"
    if code == "COT":
        return "Cert. of Transfer"
    if code == "CS":
        return "Cockpit Share"
    if code == "FRF" and source_id in (580, 639):
        return "Docking and Mooring Fee"
    if code in ("ST", "ATM"):
        return "Sultadas"
    if code in ("IM", "OPF", "SBF") or source_id in (807, 808):
        return "Miscellaneous"
    if code == "RB":
        return "Registration of Birth"
    if code == "RM":
        return "Marriage Fees"
    if code == "BF":
        return "Burial Fees"
    if code == "CE":
        return "Correction of Entry"
    if code in ("FRF", "IF"):
        return "Fishing Permit Fee"
    if code == "IAP":
        return "Sale of Agri. Prod."
    if code == "IAF":
        return "Sale of Acc. Forms"
    if code in ("WTR", "IWO"):
        return "Water Fees"
    if code in ("RFM", "MSF"):
        return "Market Stall Fee"
    if code == "SPF" or code == "RFS":
        return "Slaughterhouse Fee"
    if code in ("RFR", "IPG", "ICO"):
        return "Rent of Equipment"
    if code == "SF" and source_id == 810:
        return "Doc Stamp Tax"
    if code in ("SF", "PCL", "HEC", "OCL"):
        return "Secretary Fees"
    if code == "MDL":
        return "Med./Lab. Fees"
    if code == "GCF":
        return "Garbage Fees"
    if code in ("PFB", "BUF", "INS"):
        return "Building Permit Fee"
    if code == "EP":
        return "Electrical Permit Fee"
    if code == "ZLC":
        return "Zoning Fee"
    if code == "IFL":
        return "Livestock"
    if code == "IFD":
        return "Diving Fee"
    return None


def fetch_payments(con, args):
    collector = normalize_collector(args.collector)
    collector_clause = ""
    params = [args.date_from, args.date_to]
    if collector:
        collector_clause = """
          AND UPPER(COALESCE(NULLIF(TRIM(p.COLLECTOR), ''), NULLIF(TRIM(p.USERID), ''), '')) = ?
        """
        params.append(collector.upper())
    sql = f"""
        SELECT
            p.PAYMENT_ID,
            p.PAYMENTDATE,
            p.TRANSDATE,
            p.RECEIPTNO,
            p.PAIDBY,
            p.COLLECTOR,
            p.USERID,
            p.AFTYPE,
            p.PAYGROUP_CT,
            p.STATUS_CT,
            p.VOID_BV,
            p.RCDNUMBER,
            p.AMOUNT
        FROM PAYMENT p
        WHERE p.PAYMENTDATE >= CAST(? AS DATE)
          AND p.PAYMENTDATE < DATEADD(1 DAY TO CAST(? AS DATE))
          {collector_clause}
        ORDER BY p.PAYMENTDATE, p.RECEIPTNO, p.PAYMENT_ID
    """
    cur = con.cursor()
    try:
        cur.execute(sql, tuple(params))
        rows = []
        for row in cur.fetchall():
            receipt_no = str(row[3] or "").strip()
            if receipt_in_range(receipt_no, args.or_from, args.or_to):
                rows.append(row)
        return rows
    finally:
        try:
            cur.close()
        except Exception:
            pass


def fetch_payment_details(con, payment_ids):
    if not payment_ids:
        return {}, {}
    detail_rows = {}
    class_rows = {}
    for payment_id in payment_ids:
        cur = con.cursor()
        try:
            cur.execute(
                """
                SELECT
                    pd.PAYMENTDETAIL_ID,
                    pd.ITAXTYPE_CT,
                    it.DESCRIPTION,
                    pd.CASETYPE_CT,
                    pd.SOURCE_CT,
                    pd.SOURCEID,
                    pd.FUNDTYPE_CT,
                    pd.AMOUNTPAID,
                    pd.UNIT,
                    pd.STATUS_CT,
                    pd.RECEIPTITEMORDER
                FROM PAYMENTDETAIL pd
                LEFT JOIN T_ITAXTYPE it ON it.CODE = pd.ITAXTYPE_CT
                WHERE pd.PAYMENT_ID = ?
                ORDER BY pd.RECEIPTITEMORDER, pd.PAYMENTDETAIL_ID
                """,
                (payment_id,),
            )
            detail_rows[payment_id] = cur.fetchall()
        finally:
            try:
                cur.close()
            except Exception:
                pass

        cur = con.cursor()
        try:
            cur.execute(
                """
                SELECT
                    pcd.PAYCLASSDETAIL_ID,
                    pcd.ITAXTYPE_CT,
                    it.DESCRIPTION,
                    pcd.CASETYPE_CT,
                    pcd.TAXYEAR,
                    pcd.CLASSCODE_CT,
                    pcd.PROPERTYKIND_CT,
                    pcd.ACTUALUSE_CT,
                    pcd.AMOUNT,
                    pcd.CANCELLED_BV
                FROM PAYMENTCLASSDETAIL pcd
                LEFT JOIN T_ITAXTYPE it ON it.CODE = pcd.ITAXTYPE_CT
                WHERE pcd.PAYMENT_ID = ?
                ORDER BY pcd.ITAXTYPE_CT, pcd.TAXYEAR, pcd.CASETYPE_CT, pcd.PAYCLASSDETAIL_ID
                """,
                (payment_id,),
            )
            class_rows[payment_id] = cur.fetchall()
        finally:
            try:
                cur.close()
            except Exception:
                pass
    return detail_rows, class_rows


def case_label(case_type):
    value = (case_type or "").strip().upper()
    if value == "PEN":
        return "Penalty"
    if value == "DED":
        return "Discount"
    if value == "ADV":
        return "Advance"
    if value == "CUR":
        return "Current"
    return value or "Collection"


def build_breakdown(payment_id, paygroup, detail_rows, class_rows, fund_type):
    paygroup = (paygroup or "").strip().upper()
    fund_type = (fund_type or "").strip().upper()
    if paygroup == "RPT" or class_rows.get(payment_id):
        wanted_tax_type = None
        if fund_type == "100_GF":
            wanted_tax_type = "BSC"
        elif fund_type == "200_SEF":
            wanted_tax_type = "SEF"
        breakdown = []
        for (
            detail_id,
            tax_type,
            tax_description,
            case_type,
            tax_year,
            class_code,
            property_kind,
            actual_use,
            amount,
            cancelled_bv,
        ) in class_rows.get(payment_id, []):
            tax_type = (tax_type or "").strip()
            if wanted_tax_type and tax_type != wanted_tax_type:
                continue
            amount = money(amount)
            if cancelled_bv:
                net_amount = Decimal("0")
            else:
                net_amount = amount
            breakdown.append(
                {
                    "DESCRIPTION": f"Real Property Tax - {tax_description or tax_type} - {case_label(case_type)}",
                    "ACCOUNT_CODE": tax_type,
                    "TAX_TYPE": tax_type,
                    "SOURCE_ID": "",
                    "SOURCE_NAME": tax_description or ("Basic Real Property Tax" if tax_type == "BSC" else "Special Education Fund"),
                    "CATEGORY": "Real Property Tax",
                    "PERIOD_COVERED": str(tax_year or ""),
                    "CLASS_CODE": (class_code or "").strip(),
                    "PROPERTY_KIND": (property_kind or "").strip(),
                    "ACTUAL_USE": (actual_use or "").strip(),
                    "AMOUNT": to_json_value(amount),
                    "PENALTY": to_json_value(amount if (case_type or "").strip().upper() == "PEN" else Decimal("0")),
                    "DISCOUNT": to_json_value(abs(amount) if (case_type or "").strip().upper() == "DED" else Decimal("0")),
                    "NET_AMOUNT": to_json_value(net_amount),
                    "FIREBIRD_DETAIL_ID": detail_id,
                    "CANCELLED_BV": bool(cancelled_bv),
                }
            )
        return breakdown

    breakdown = []
    for (
        detail_id,
        tax_type,
        tax_description,
        case_type,
        source_ct,
        source_id,
        fund,
        amount,
        unit,
        status_ct,
        item_order,
    ) in detail_rows.get(payment_id, []):
        tax_type = (tax_type or "").strip()
        source_ct = (source_ct or "").strip()
        source_name = classify_summary_source(tax_type, source_id, source_ct)
        if not source_name:
            source_name = tax_description or tax_type or source_ct or "Payment Detail"
        amount = money(amount)
        breakdown.append(
            {
                "DESCRIPTION": source_name,
                "ACCOUNT_CODE": tax_type,
                "TAX_TYPE": tax_type,
                "SOURCE_ID": source_id,
                "SOURCE_CT": source_ct,
                "SOURCE_NAME": source_name,
                "CATEGORY": source_name,
                "PERIOD_COVERED": "",
                "AMOUNT": to_json_value(amount),
                "PENALTY": to_json_value(amount if tax_type == "FPT" else Decimal("0")),
                "DISCOUNT": 0.0,
                "NET_AMOUNT": to_json_value(amount),
                "FIREBIRD_DETAIL_ID": detail_id,
                "FUNDTYPE_CT": (fund or "").strip(),
                "DETAIL_STATUS_CT": (status_ct or "").strip(),
                "RECEIPTITEMORDER": item_order,
            }
        )
    return breakdown


def detect_form_type(aftype, paygroup, breakdown):
    value = (aftype or "").strip()
    if value:
        return value
    paygroup = (paygroup or "").strip().upper()
    if paygroup == "RPT":
        return "AF 56"
    for detail in breakdown:
        if detail.get("TAX_TYPE") == "CTC" or detail.get("SOURCE_CT") in ("CTCI", "CTCC"):
            return "Comm Tax."
    return "AF 51"


def detect_payment_type(paygroup, breakdown):
    paygroup = (paygroup or "").strip().upper()
    if paygroup == "RPT":
        return "Real Property Tax"
    for detail in breakdown:
        if detail.get("TAX_TYPE") == "CTC" or detail.get("SOURCE_CT") in ("CTCI", "CTCC"):
            return "Community Tax Certificate"
    return "General Fund / Other Fees and Charges"


def validation_for_item(receipt_status, rcd_number, amount, is_duplicate, breakdown):
    messages = []
    status = "VALID"
    if receipt_status in ("VOID", "CANCELLED"):
        messages.append(f"Receipt is {receipt_status.lower()}.")
        status = "WARNING"
    if rcd_number:
        messages.append(f"Receipt already has Firebird RCD reference: {rcd_number}.")
        status = "WARNING"
    if amount <= 0:
        messages.append("Amount is zero or null.")
        status = "ERROR"
    if is_duplicate:
        messages.append("Duplicate OR number found in this JSON batch.")
        status = "ERROR"
    if not breakdown:
        messages.append("No payment breakdown lines found for this receipt/fund.")
        status = "ERROR"
    if not messages:
        messages.append("OR found, receipt is paid, and payment breakdown was extracted.")
    return status, " ".join(messages)


def expected_missing_ors(or_from, or_to, found_receipts):
    start = numeric_or(or_from)
    end = numeric_or(or_to)
    if start is None or end is None or end < start or end - start > 5000:
        return []
    width = max(len(re.sub(r"\D", "", str(or_from))), len(re.sub(r"\D", "", str(or_to))))
    found_numbers = {numeric_or(receipt) for receipt in found_receipts}
    return [str(number).zfill(width) for number in range(start, end + 1) if number not in found_numbers]


def build_json(args):
    extracted_at = datetime.now().astimezone().isoformat(timespec="seconds")
    collector = normalize_collector(args.collector)
    con = open_connection(args.connection, args.user, args.password)
    try:
        payment_rows = fetch_payments(con, args)
        payment_ids = [row[0] for row in payment_rows]
        detail_rows, class_rows = fetch_payment_details(con, payment_ids)
        con.rollback()
    finally:
        con.close()

    seen_receipts = {}
    items = []
    validation_issues = []

    for row in payment_rows:
        (
            payment_id,
            payment_date,
            transaction_date,
            receipt_no,
            paid_by,
            payment_collector,
            user_id,
            aftype,
            paygroup,
            status_ct,
            void_bv,
            rcd_number,
            header_amount,
        ) = row
        receipt_no = str(receipt_no or "").strip()
        breakdown = build_breakdown(payment_id, paygroup, detail_rows, class_rows, args.fund)
        line_total = sum(money(detail.get("NET_AMOUNT")) for detail in breakdown)
        amount = line_total if breakdown else money(header_amount)
        receipt_status = payment_status_label(status_ct, void_bv)
        seen_receipts[receipt_no] = seen_receipts.get(receipt_no, 0) + 1
        is_duplicate = seen_receipts[receipt_no] > 1
        validation_status, validation_message = validation_for_item(
            receipt_status, rcd_number, amount, is_duplicate, breakdown
        )
        if validation_status != "VALID":
            validation_issues.append(
                {
                    "OR_NO": receipt_no,
                    "ISSUE_TYPE": validation_status,
                    "SEVERITY": "ERROR" if validation_status == "ERROR" else "WARNING",
                    "MESSAGE": validation_message,
                }
            )

        payment_type = detect_payment_type(paygroup, breakdown)
        form_type = detect_form_type(aftype, paygroup, breakdown)
        tax_type = ""
        account_code = ""
        if breakdown:
            tax_type = breakdown[0].get("TAX_TYPE") or ""
            account_code = breakdown[0].get("ACCOUNT_CODE") or ""

        items.append(
            {
                "OR_DATE": to_json_value(payment_date),
                "OR_NO": receipt_no,
                "OR_FROM": args.or_from or receipt_no,
                "OR_TO": args.or_to or receipt_no,
                "TAXPAYER_NAME": (paid_by or "").strip() or "-",
                "AMOUNT": to_json_value(amount),
                "LINE_TOTAL": to_json_value(line_total),
                "TRANSACTION_DATE": to_json_value(transaction_date or payment_date),
                "RCD_NO": rcd_number or "",
                "RCD_DATE": args.rcd_date or args.date_to,
                "RCD_TRANSACTION_DATE": extracted_at,
                "RCD_STATUS": "DRAFT",
                "RCD_COLLECTOR": collector or (payment_collector or user_id or ""),
                "COLLECTOR_ID": collector or (payment_collector or user_id or ""),
                "CASHIER_NAME": args.cashier_name or "",
                "FUND_TYPE": args.fund,
                "FORM_TYPE": form_type,
                "PAYMENT_TYPE": payment_type,
                "TAX_TYPE": tax_type,
                "ACCOUNT_CODE": account_code,
                "RECEIPT_STATUS": receipt_status,
                "SOURCE_TABLE": "PAYMENT/PAYMENTDETAIL/PAYMENTCLASSDETAIL",
                "FIREBIRD_PAYMENT_ID": payment_id,
                "FIREBIRD_DETAIL_ID": breakdown[0].get("FIREBIRD_DETAIL_ID") if breakdown else "",
                "IS_DUPLICATE": is_duplicate,
                "VALIDATION_STATUS": validation_status,
                "VALIDATION_MESSAGE": validation_message,
                "CREATED_AT": extracted_at,
                "UPDATED_AT": extracted_at,
                "CONFIRMED_AT": None,
                "CONFIRMED_BY": None,
                "REMARKS": "",
                "payment_breakdown": breakdown,
            }
        )

    missing = expected_missing_ors(args.or_from, args.or_to, [item["OR_NO"] for item in items])
    for missing_or in missing:
        validation_issues.append(
            {
                "OR_NO": missing_or,
                "ISSUE_TYPE": "MISSING_OR",
                "SEVERITY": "WARNING",
                "MESSAGE": "OR number is inside the requested range but was not found in Firebird for the selected filters.",
            }
        )

    total_amount = sum(money(item["AMOUNT"]) for item in items)
    paid_amount = sum(money(item["AMOUNT"]) for item in items if item["RECEIPT_STATUS"] == "PAID")
    group_totals = {}
    for item in items:
        key = (item["FUND_TYPE"], item["FORM_TYPE"], item["PAYMENT_TYPE"])
        group = group_totals.setdefault(
            key,
            {
                "FUND_TYPE": item["FUND_TYPE"],
                "FORM_TYPE": item["FORM_TYPE"],
                "PAYMENT_TYPE": item["PAYMENT_TYPE"],
                "TOTAL_RECEIPTS": 0,
                "TOTAL_AMOUNT": Decimal("0"),
            },
        )
        group["TOTAL_RECEIPTS"] += 1
        group["TOTAL_AMOUNT"] += money(item["AMOUNT"])

    json_batch_id = f"RCDTMP-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    return {
        "JSON_BATCH_ID": json_batch_id,
        "EXTRACTED_AT": extracted_at,
        "EXTRACTED_BY": args.extracted_by,
        "SOURCE_SYSTEM": "Firebird .FDB read-only",
        "SOURCE_DATABASE": os.environ.get("ESRE_FIREBIRD_DB", "ZAMBOANGUITA.FDB"),
        "SCHEMA_VERSION": "1.0",
        "batch": {
            "RCD_NO": args.rcd_no or "",
            "RCD_DATE": args.rcd_date or args.date_to,
            "RCD_TRANSACTION_DATE": extracted_at,
            "RCD_STATUS": "DRAFT",
            "RCD_COLLECTOR": collector,
            "COLLECTOR_ID": collector,
            "CASHIER_NAME": args.cashier_name or "",
            "CONFIRMED_AT": None,
            "CONFIRMED_BY": None,
            "REMARKS": "Preview JSON only. No Firebird data was inserted, updated, or deleted.",
        },
        "filters": {
            "FUND_TYPE": args.fund,
            "FORM_TYPE": args.form_type or "",
            "OR_FROM": args.or_from or "",
            "OR_TO": args.or_to or "",
            "TRANSACTION_DATE_FROM": args.date_from,
            "TRANSACTION_DATE_TO": args.date_to,
            "INCLUDE_CANCELLED_OR_VOID": True,
            "READ_ONLY": True,
        },
        "items": items,
        "summary": {
            "TOTAL_RECEIPTS": len(items),
            "TOTAL_AMOUNT": to_json_value(total_amount),
            "PAID_TOTAL_AMOUNT": to_json_value(paid_amount),
            "VALID_COUNT": sum(1 for item in items if item["VALIDATION_STATUS"] == "VALID"),
            "MISSING_COUNT": len(missing),
            "VOID_COUNT": sum(1 for item in items if item["RECEIPT_STATUS"] == "VOID"),
            "CANCELLED_COUNT": sum(1 for item in items if item["RECEIPT_STATUS"] == "CANCELLED"),
            "DUPLICATE_COUNT": sum(1 for item in items if item["IS_DUPLICATE"]),
            "INVALID_COUNT": sum(1 for item in items if item["VALIDATION_STATUS"] == "ERROR"),
            "WARNING_COUNT": sum(1 for item in items if item["VALIDATION_STATUS"] == "WARNING") + len(missing),
            "GROUP_TOTALS": [
                {
                    **{key: value for key, value in group.items() if key != "TOTAL_AMOUNT"},
                    "TOTAL_AMOUNT": to_json_value(group["TOTAL_AMOUNT"]),
                }
                for group in group_totals.values()
            ],
        },
        "validation_issues": validation_issues,
        "next_step": {
            "STATUS": "FOR_USER_REVIEW",
            "MESSAGE": "Review this JSON first. MySQL save should only happen after confirmation.",
        },
    }


def write_output(payload, output_path=None):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if output_path:
        output_path = Path(output_path)
    else:
        collector = re.sub(r"[^A-Za-z0-9_-]+", "_", payload["batch"]["RCD_COLLECTOR"] or "collector")
        output_path = OUTPUT_DIR / f"rcd_remittance_preview_{collector}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with output_path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, ensure_ascii=False, default=to_json_value)
    return output_path


def compact_payload(payload):
    compact_items = []
    for item in payload.get("items", []):
        compact_items.append(
            {
                "OR_NO": item.get("OR_NO", ""),
                "OR_DATE": item.get("OR_DATE", ""),
                "TAXPAYER_NAME": item.get("TAXPAYER_NAME", ""),
                "AMOUNT": item.get("AMOUNT", 0),
                "RECEIPT_STATUS": item.get("RECEIPT_STATUS", ""),
                "IS_REMITTED": bool(item.get("RCD_NO")),
                "RCD_NO": item.get("RCD_NO", ""),
                "RCD_DATE": item.get("RCD_DATE", ""),
                "RCD_STATUS": item.get("RCD_STATUS", ""),
                "RCD_COLLECTOR": item.get("RCD_COLLECTOR", ""),
                "PAYMENT_TYPE": item.get("PAYMENT_TYPE", ""),
                "FORM_TYPE": item.get("FORM_TYPE", ""),
                "FUND_TYPE": item.get("FUND_TYPE", ""),
                "VALIDATION_STATUS": item.get("VALIDATION_STATUS", ""),
                "VALIDATION_MESSAGE": item.get("VALIDATION_MESSAGE", ""),
            }
        )

    return {
        "JSON_BATCH_ID": payload.get("JSON_BATCH_ID"),
        "EXTRACTED_AT": payload.get("EXTRACTED_AT"),
        "EXTRACTED_BY": payload.get("EXTRACTED_BY"),
        "SOURCE_SYSTEM": payload.get("SOURCE_SYSTEM"),
        "SCHEMA_VERSION": payload.get("SCHEMA_VERSION"),
        "batch": payload.get("batch", {}),
        "filters": payload.get("filters", {}),
        "items": compact_items,
        "summary": payload.get("summary", {}),
        "validation_issues": payload.get("validation_issues", []),
        "next_step": payload.get("next_step", {}),
    }


def main():
    parser = argparse.ArgumentParser(
        description="Preview-only RCD remittance scraper. Reads Firebird in read-only mode and writes JSON."
    )
    parser.add_argument("date_from", type=parse_date, help="Start collection date, YYYY-MM-DD.")
    parser.add_argument("date_to", type=parse_date, help="End collection date, YYYY-MM-DD.")
    parser.add_argument("collector", help="Collector name/code. Example: RICARDO, IRIS, angelique.")
    parser.add_argument("--or-from", help="Official receipt number from.")
    parser.add_argument("--or-to", help="Official receipt number to.")
    parser.add_argument("--fund", default="100_GF", choices=("100_GF", "200_SEF", "ALL"), help="RCD fund/template.")
    parser.add_argument("--form-type", default="", help="Optional display/input form type.")
    parser.add_argument("--rcd-no", default="", help="Optional manual RCD number.")
    parser.add_argument("--rcd-date", type=parse_date, help="Optional RCD date. Defaults to date_to.")
    parser.add_argument("--cashier-name", default="", help="Optional cashier/remittance receiver.")
    parser.add_argument("--extracted-by", default="LGU Treasury Admin", help="User who generated the preview.")
    parser.add_argument("--connection", choices=("odbc", "native", "auto"), default=os.environ.get("ESRE_CONNECTION", "odbc"))
    parser.add_argument("--user", default="SYSDBA")
    parser.add_argument("--password", default="masterkey")
    parser.add_argument("--output", help="Optional JSON output path.")
    parser.add_argument(
        "--compact",
        action="store_true",
        help="Write shorter JSON for remittance status preview. Omits payment_breakdown.",
    )
    args = parser.parse_args()

    payload = build_json(args)
    if args.compact:
        payload = compact_payload(payload)
    output_path = write_output(payload, args.output)
    print(f"JSON_BATCH_ID: {payload['JSON_BATCH_ID']}")
    print(f"Receipts: {payload['summary']['TOTAL_RECEIPTS']}")
    print(f"Total amount: {payload['summary']['TOTAL_AMOUNT']:,.2f}")
    print(f"Paid total amount: {payload['summary']['PAID_TOTAL_AMOUNT']:,.2f}")
    print(f"Validation issues: {len(payload['validation_issues'])}")
    print(f"Output file: {output_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
