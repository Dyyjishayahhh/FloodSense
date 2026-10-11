# Provisional feasibility-study materials

## Schedule feasibility

The title page identifies a September 2026–May 2027 study period. The repository shows three broad development checkpoints (2026-09-20, 2026-10-09, and 2026-10-11) and staged implementation guides. The following is a planning schedule, not a record that every listed task has been completed.

| Activity | Sep–Oct 2026 | Nov–Dec 2026 | Jan–Feb 2027 | Mar–Apr 2027 | May 2027 | Status/evidence |
|---|---|---|---|---|---|---|
| Requirements and consultation review | ███ | ░ |  |  |  | Ongoing; survey still open |
| Core backend/mobile/Admin development | ███ | █ | ░ |  |  | Major integrated functions present |
| Local data acquisition and validation | ░ | ███ | ███ | ░ |  | Blocked/awaiting authorized data and expert validation |
| Specific SDLC documentation | ░ | █ |  |  |  | Pending group decision/evidence |
| Functional/integration testing | ███ | ███ | █ | ░ |  | Mobile current; backend current rerun pending |
| Needs-survey closure and analysis | ░ | ███ |  |  |  | `[Insert cutoff date and valid N.]` |
| ISO/IEC 25010 evaluation |  | ░ | ███ | █ |  | Planned; evaluator/instrument details pending |
| Manuscript consolidation and defense preparation | ░ | █ | ███ | ███ | ███ | Working revision prepared |

Legend: █ planned/active concentration; ░ preparatory or contingency period. Confirm dates with the approved school calendar.

## Cost feasibility

No reliable actual-expense ledger was found. Older feasibility material refers to superseded Firebase/AI concepts and must not be used as the current budget. Enter documented receipts or approved estimates only.

| Line item | Basis | Quantity/unit | Actual or approved estimate (PHP) | Evidence required |
|---|---|---|---|---|
| Flutter/Dart, Python/Django, PostgreSQL/PostGIS | Open-source development stack | — | 0 licensing fee; exclude hosting/usage costs | Official license/project records |
| Map provider usage | Mapbox/OSM provider-aware configuration | `[Enter tier/usage]` | `[Enter verified amount.]` | Account billing/tier and attribution terms |
| Hosting/database/deployment | Not yet verified as production deployed | `[Enter plan.]` | `[Enter verified amount.]` | Quotation or invoice |
| Internet/data | Survey, development, testing | `[Enter months/GB.]` | `[Enter actual amount.]` | Receipts/plan statement |
| Transportation | BDRRMO, CPDCO, school, survey work | `[Enter trips.]` | `[Enter actual amount.]` | Trip log/receipts |
| Printing/reproduction | Letters, consent, instruments, manuscript | `[Enter pages/copies.]` | `[Enter actual amount.]` | Receipts |
| Device testing | Existing or borrowed Android devices | `[List devices.]` | `[Enter incremental cost.]` | Ownership/borrow record |
| Contingency | Only if adviser approves | `[Enter basis.]` | `[Enter approved amount.]` | Approved budget rule |
| Total | Sum only verified entries | — | `[Compute after entries are supplied.]` | Complete ledger |

## Operational feasibility

The workflow is understandable as a resident-initiated process: authenticate/complete onboarding, confirm a hypothetical rainfall scenario and location, request an assessment, review an explanation and map, and optionally use structured preparedness or center information. Authorized administrators maintain governed records through a separate portal. Operational adoption is not yet established because the needs survey and system evaluation are incomplete, public center and susceptibility data remain unapproved, and production support arrangements are absent.

Operational risks include users mistaking a scenario result for a forecast, outdated center records, inaccessible network service, unclear responsibility for publication approval, and incorrect extrapolation from provisional maps. Mitigations include prominent limitation/source notices, publication and review states, audit history, one-time foreground location, no raw-rule editing in the ordinary Admin portal, and an escalation statement directing users to official warnings and authorities.

## Technical feasibility

The integrated stack demonstrates technical feasibility for a prototype: Flutter/Dart mobile UI; Django/DRF/GeoDjango services; PostgreSQL/PostGIS records; deterministic inference; structured DSS; provider-aware Mapbox/OSM mapping; and custom Admin workflows. Current mobile verification passed 389 tests and static analysis. A current backend rerun, production-like environment test, device/browser matrix, load/security testing, backup/restore procedure, monitoring, and deployment evidence remain pending.

The largest technical constraint is not the ability to compute or display results; it is the absence of an approved operational Bacoor susceptibility dataset, validated rule parameters, and authorized current evacuation-center data. Software readiness must not be substituted for scientific or institutional validity.
