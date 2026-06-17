# Citizen Payment Kiosk API Example

This is a starter API-layer design for a future Citizen Payment Kiosk.

Important principle:

```text
Kiosk screen -> API layer -> Firebird database
```

The kiosk should not connect directly to `ZAMBOANGUITA.FDB`.

## What This Example Does

- Shows how to connect to Firebird through Python.
- Provides read-only inquiry endpoints.
- Provides a safe `payment-intent` endpoint that only prepares a reference, not an official receipt.
- Keeps official receipt posting separate from online/kiosk payment confirmation.

## Install Dependencies

```powershell
pip install fastapi uvicorn fdb
```

## Run API

From the main project folder:

```powershell
python -m uvicorn kiosk_api_example.main:app --host 127.0.0.1 --port 8000
```

Open:

```text
http://127.0.0.1:8000/docs
```

## Example Endpoints

```text
GET /health
GET /taxpayer/{local_tin}
GET /rpt/payables/{local_tin}
GET /ctc/rates
GET /other-fees/rates
POST /payment-intents
```

## Production Notes

For a real kiosk, add:

- authentication for kiosk terminals
- HTTPS
- audit logs
- payment gateway callback verification
- duplicate-payment prevention
- official receipt number locking/allocation
- cashier/RCD/accountable form rules
- strict read/write transaction separation

