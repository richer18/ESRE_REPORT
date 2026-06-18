from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any
from uuid import uuid4

import fdb
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field


DB_PATH = r"E:\ZAMBOANGUITA.FDB"
FB_CLIENT = r"C:\Program Files\Firebird\Firebird_2_5\bin\fbclient.dll"


app = FastAPI(
    title="Citizen Payment Kiosk API Example",
    version="0.1.0",
    description="Example API layer between a kiosk app and the LGU Firebird database.",
)


class PaymentIntentRequest(BaseModel):
    process: str = Field(..., examples=["RPT", "CTC", "OTHER"])
    local_tin: str | None = Field(None, examples=["1234567890123"])
    payer_name: str = Field(..., examples=["JUAN DELA CRUZ"])
    amount: Decimal = Field(..., gt=0, examples=["687.50"])
    reference_note: str | None = Field(None, examples=["Water Fee January 2026"])


class PaymentIntentResponse(BaseModel):
    payment_reference: str
    status: str
    process: str
    payer_name: str
    amount: Decimal
    created_at: datetime
    next_step: str


def connect_readonly():
    return fdb.connect(
        dsn=DB_PATH,
        user="SYSDBA",
        password="masterkey",
        charset="UTF8",
        fb_library_name=FB_CLIENT,
        isolation_level=fdb.ISOLATION_LEVEL_READ_COMMITED_RO,
        no_db_triggers=True,
        no_gc=True,
    )


def rows_to_dicts(cursor) -> list[dict[str, Any]]:
    columns = [description[0].strip().lower() for description in cursor.description]
    return [
        {
            column: value.strip() if isinstance(value, str) else value
            for column, value in zip(columns, row)
        }
        for row in cursor.fetchall()
    ]


@app.get("/health")
def health():
    return {"status": "ok", "service": "citizen-payment-kiosk-api-example"}


@app.get("/taxpayer/{local_tin}")
def get_taxpayer(local_tin: str):
    sql = """
        SELECT
            LOCAL_TIN,
            OWNERNAME,
            OWNERADDRESS,
            FIRSTNAME,
            LASTNAME,
            MI,
            BARANGAY_CT,
            MUNICIPAL_ID,
            ACTIVE_BV
        FROM TAXPAYER
        WHERE LOCAL_TIN = ?
    """
    con = connect_readonly()
    try:
        cur = con.cursor()
        cur.execute(sql, (local_tin,))
        rows = rows_to_dicts(cur)
        con.rollback()
    finally:
        con.close()

    if not rows:
        raise HTTPException(status_code=404, detail="Taxpayer not found.")
    return rows[0]


@app.get("/rpt/payables/{local_tin}")
def get_rpt_payables(local_tin: str):
    sql = """
        SELECT
            tpa.LOCAL_TIN,
            tpa.PROP_ID,
            tpa.TAXTRANS_ID,
            tpa.TAXYEAR,
            tpa.ITAXTYPE_CT,
            it.DESCRIPTION AS ITAXTYPE_DESCRIPTION,
            SUM(COALESCE(tpa.DEBITAMOUNT, 0)) AS TOTAL_DEBIT,
            SUM(COALESCE(tpa.CREDITAMOUNT, 0)) AS TOTAL_CREDIT,
            SUM(COALESCE(tpa.DEBITAMOUNT, 0) - COALESCE(tpa.CREDITAMOUNT, 0)) AS BALANCE
        FROM TPACCOUNT tpa
        LEFT JOIN T_ITAXTYPE it ON it.CODE = tpa.ITAXTYPE_CT
        WHERE tpa.LOCAL_TIN = ?
          AND COALESCE(tpa.CANCELLED_BV, 0) = 0
        GROUP BY
            tpa.LOCAL_TIN,
            tpa.PROP_ID,
            tpa.TAXTRANS_ID,
            tpa.TAXYEAR,
            tpa.ITAXTYPE_CT,
            it.DESCRIPTION
        HAVING SUM(COALESCE(tpa.DEBITAMOUNT, 0) - COALESCE(tpa.CREDITAMOUNT, 0)) > 0
        ORDER BY tpa.TAXYEAR, tpa.PROP_ID, tpa.ITAXTYPE_CT
    """
    con = connect_readonly()
    try:
        cur = con.cursor()
        cur.execute(sql, (local_tin,))
        rows = rows_to_dicts(cur)
        con.rollback()
    finally:
        con.close()
    return {"local_tin": local_tin, "payables": rows}


@app.get("/ctc/rates")
def get_ctc_rates():
    sql = "SELECT * FROM T_CTCRATE ORDER BY YEARSTART"
    con = connect_readonly()
    try:
        cur = con.cursor()
        cur.execute(sql)
        rows = rows_to_dicts(cur)
        con.rollback()
    finally:
        con.close()
    return {"rates": rows}


@app.get("/other-fees/rates")
def get_other_fee_rates():
    sql = """
        SELECT
            opr.OPRATE_ID,
            opr.ITAXTYPE_CT,
            it.DESCRIPTION AS ITAXTYPE_DESCRIPTION,
            opr.OPGROUP,
            opr.OPSUBGROUP,
            opr.DESCRIPTION AS RATE_DESCRIPTION,
            opr.RATE,
            opr.YEARSTART,
            opr.VALIDFROM,
            opr.ORDINANCENO,
            it.FUNDTYPE_CT
        FROM T_OTHERPAYMENTRATE opr
        LEFT JOIN T_ITAXTYPE it ON it.CODE = opr.ITAXTYPE_CT
        ORDER BY opr.ITAXTYPE_CT, opr.OPGROUP, opr.OPSUBGROUP, opr.DESCRIPTION
    """
    con = connect_readonly()
    try:
        cur = con.cursor()
        cur.execute(sql)
        rows = rows_to_dicts(cur)
        con.rollback()
    finally:
        con.close()
    return {"rates": rows}


@app.post("/payment-intents", response_model=PaymentIntentResponse)
def create_payment_intent(request: PaymentIntentRequest):
    # This intentionally does not write to PAYMENT or PAYMENTDETAIL.
    # In production, store payment intent records in a separate kiosk database/table,
    # then post official receipts only after verified payment gateway confirmation.
    reference = f"KIOSK-{datetime.now():%Y%m%d}-{uuid4().hex[:10].upper()}"
    return PaymentIntentResponse(
        payment_reference=reference,
        status="PENDING_PAYMENT",
        process=request.process.upper(),
        payer_name=request.payer_name,
        amount=request.amount,
        created_at=datetime.now(),
        next_step="Send this reference to payment gateway or cashier validation queue.",
    )

