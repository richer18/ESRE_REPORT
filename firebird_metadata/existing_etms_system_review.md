# Existing Electronic Treasurer Management System Review

Reviewed source:

```text
C:\Users\LIFT-LAPTOP\OneDrive\Desktop\ElectronicTreasurerManagementSystem\frontend
C:\Users\LIFT-LAPTOP\OneDrive\Desktop\ElectronicTreasurerManagementSystem\backend
```

Review mode: read-only. No MySQL server was started.

## Stack

```text
Backend: Laravel 12 + Sanctum + MySQL-oriented APIs
Frontend: React CRA + Material UI + Toolpad DashboardLayout
Database assumption: MySQL, mostly configured around 3307 in backend config but .env.example says 3306
```

## Major Modules Found

```text
Authentication and role permissions
Dashboard and chart summaries
General Fund abstract collections
Trust Fund abstract collections
Community Tax Certificate / Cedula
Real Property Tax
Summary / Full Report / ESRE reports
RCD / accountable forms / receipt ranges
BPLO / business registration / MCH / e-bike trisikad
Waterworks
Procurement document forms
Calendar and user management
```

## Useful Patterns To Reuse

1. Authentication pattern:
   - Sanctum session authentication
   - CSRF cookie flow
   - `AuthContext` in React
   - axios interceptor for `401`

2. Permission pattern:
   - Role-to-permission matrix
   - Module permission middleware
   - Frontend `hasAnyPermission` route guard

3. Sidebar/navigation pattern:
   - Dashboard
   - Daily Operations
   - Collections
   - Business and Utilities
   - Administration and Setup
   - Reports and Analytics

4. Treasury report grouping:
   - Real Property Tax
   - General Fund
   - Trust Fund
   - Community Tax Certificate
   - Full Report
   - ESRE
   - RCD / Summary of Collection

5. Reporting helper logic:
   - Active/cancelled filters
   - month/year/day filters
   - collector/search filters
   - General Fund source-ID bucket mapping
   - Trust Fund fee-code mapping
   - RPT land/building classification mapping, including special classification handling

6. Full Report / RCD calculation concept:
   - CTC total
   - RPT total
   - General Fund total
   - Trust Fund total
   - GF + TF
   - Due From manual/adjustment fields
   - under/over fields

## What Not To Copy Directly

```text
Do not copy MySQL triggers/events into the new Firebird-first system.
Do not start MySQL on this laptop because of LIFT server port conflict risk.
Do not copy old run_file/tools layout; Python belongs in LGU_TreasuryReportingSystem\runner.
Do not copy the whole controller-per-small-endpoint structure blindly.
Do not migrate to React CRA; the new system is Vite + React.
Do not trust the old repo git status until it is separately cleaned; it shows a very noisy deleted-file state.
```

## Recommended Use In New System

```text
Use old system as workflow/reference only.
Keep new backend as Laravel REST API.
Keep Firebird as read-only source through Python runner.
Use old permission/auth ideas later when real login is implemented.
Use old sidebar/report grouping to refine the new UI.
Use old report helper formulas as reference, but rewrite against Firebird/FDB query outputs.
Use old RCD/accountable form design as a future module reference.
```

## Priority Ideas To Port

1. Real Laravel Sanctum login and logout.
2. Role/permission matrix for modules.
3. Sidebar grouping from the old dashboard.
4. Report preview API pattern per module.
5. Active/cancelled/void filters in all reports.
6. Full Report / RCD reconciliation model.
7. Collector summary/dashboard widgets.
