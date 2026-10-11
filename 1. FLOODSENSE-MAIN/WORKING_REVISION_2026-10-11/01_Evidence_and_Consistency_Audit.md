# Evidence and consistency audit

## Repository state

- Current development commit: `9d32d815e0c54e081e4b2ca617bbe939d1df6b7c`
- Commit date/message: 2026-10-11 — `Replace Project FIles`
- Branch: `main`
- Working tree: not clean. Several tracked interview/questionnaire-path copies are deleted while equivalent or newer interview files are untracked under `interviews/`; `tmp/` contains audit working files. These user changes were preserved.
- Current mobile verification: `flutter test` passed 389 tests; `flutter analyze` reported no issues.
- Backend verification: not rerun in the available bundled Python because Django dependencies were unavailable. Repository documentation records a historical 2026-10-05 run of 891 backend tests, 327 Flutter tests, Flutter analysis, and a debug APK build. This is historical evidence, not a current rerun.

## Evidence hierarchy applied

Integrated code, routes, services, models, migrations, interfaces, and tests were used together. A model, proposal, mockup, script, or screenshot alone was not treated as a complete function. The 47-barangay layer was treated as an administrative reference only. Provisional MGB and other research data were not described as official or operational.

## Current development status

| Feature or module | Status | Current evidence and interface | Research/data qualification | Manuscript action |
|---|---|---|---|---|
| User-confirmed rainfall intensity and duration | IMPLEMENTED AND CURRENTLY VERIFIED | Mobile assessment flow, API options/evaluate routes, current mobile tests | Scenario input; not live weather | Retain and state explicitly |
| Scenario confirmation and explicit assessment request | IMPLEMENTED AND CURRENTLY VERIFIED | Mobile interaction and tests | No automatic polling | Retain |
| Manual area selection and map pin | IMPLEMENTED AND CURRENTLY VERIFIED | Resident map and assessment widgets | Limited to supported data | Retain |
| One-time foreground GPS and barangay confirmation/correction | IMPLEMENTED AND CURRENTLY VERIFIED | Location flow, geographic endpoints, current mobile tests | Not background tracking | Retain with privacy limit |
| 47-barangay reference display | IMPLEMENTED AND CURRENTLY VERIFIED | Boundary endpoint and map display | Derived/technically checked; not City-issued and not susceptibility data | Keep qualification |
| Deterministic rule matching, priorities, conflicts, and explanation | IMPLEMENTED AND CURRENTLY VERIFIED | Expert-system services/endpoints and current mobile integration tests | Operational local validity blocked by data/rule approval | Describe software function, not official result |
| Susceptibility output (Low–Very High) | IMPLEMENTED AND CURRENTLY VERIFIED | Assessment response and result UI | Demonstration/provisional inputs only | Retain qualification |
| Structured multi-question DSS | IMPLEMENTED AND CURRENTLY VERIFIED | Structured-flow start/answer endpoints, resident Prepare UI, current mobile tests | Content still requires source/approval governance | Correct manuscript’s “unverified/incomplete” wording |
| Published DSS guidance separate from assessment | IMPLEMENTED AND CURRENTLY VERIFIED | Guidance endpoints and UI | Cannot change susceptibility class or issue orders | Retain separation |
| Resident registration, login, verification, legal acceptance, onboarding, preferences, password and deletion controls | IMPLEMENTED AND CURRENTLY VERIFIED | Auth/account routes, resident shell, screens, current mobile tests | Production identity/deployment not established | Correct manuscript’s “unimplemented” wording |
| Assessment history | NOT IMPLEMENTED | No persistent resident assessment history; DSS history flag is false | — | Keep as limitation |
| Notifications/live alerts | OUTSIDE THE CURRENT SCOPE | Mobile copy explicitly states live alerts are not enabled | — | Do not claim |
| Provider-aware Mapbox/OSM map | IMPLEMENTED AND CURRENTLY VERIFIED | Mapbox and flutter_map/OSM paths, provider configuration, tests | Token/provider configuration required | Update technology description |
| Evacuation-center Admin workflow | IMPLEMENTED AND CURRENTLY VERIFIED | Models, routes, custom Admin forms/import review | Data approval/currentness separate | Retain |
| Resident nearest-center lookup and display | IMPLEMENTED AND CURRENTLY VERIFIED | Map/nearest endpoints, resident UI, current mobile tests | Public usefulness blocked by authorized current center data; no live capacity/safe routing | Correct manuscript’s “not exposed” wording |
| Evacuation-center CSV/XLSX staged import | IMPLEMENTED AND CURRENTLY VERIFIED | Staging/review/import workflow and tests in code | Applies to center records, not Expert System parameters | Distinguish from parameter CSV requirement |
| Fixed-schema Expert System parameter CSV export/import | CONFIRMED REQUIREMENT — NOT YET IMPLEMENTED | No complete export/import routes, schema validation, ordinary Admin controls, and end-to-end tests | Confirmed governance requirement | State as unimplemented |
| Custom Admin authentication/dashboard | IMPLEMENTED AND CURRENTLY VERIFIED | Authenticated custom portal routes/templates/tests | Technical Django Admin is separate | Retain distinction |
| Map/rainfall reference review, parameter governance, DSS content, sources, audit history | IMPLEMENTED AND CURRENTLY VERIFIED | Custom Admin routes, services, interfaces | Publication/approval status still governs content | Update scope |
| Raw rule editing by ordinary administrators | OUTSIDE THE CURRENT SCOPE | Not exposed in custom Admin | Technical Django Admin reserved for research/developer maintenance | Preserve prohibition |
| MGB susceptibility import/data | IMPLEMENTED BUT NOT EXPOSED TO THE INTENDED USER | Models/import processing exist | Data are provisional/unapproved; no proof of operational integration | Do not present as official/current public data |
| Approved operational Bacoor susceptibility dataset | BLOCKED OR AWAITING DATA | No approved dataset in `research_data/approved` | Requires provenance, permission, method, and institutional/expert validation | Keep as primary limitation |
| Real-time rainfall/water-level monitoring, official forecast/warning | OUTSIDE THE CURRENT SCOPE | No operational sensors/feed/polling; explicitly excluded | — | Exclude claims |
| Offline operation | NOT IMPLEMENTED | Integrated functions depend on server/network | Error/offline states are not offline capability | Keep limitation |
| Production deployment | NOT IMPLEMENTED | No verified production release/operations evidence | — | Do not claim |
| Needs survey | ACTIVE WORK IN PROGRESS | Exact approved questionnaire exists and is currently being answered | Final valid N, percentages, and response rate unavailable | Use placeholders only |
| ISO/IEC 25010 evaluation | PLANNED | Quality model/reference present; no completed validated evaluation dataset | Evaluator groups/instrument/scoring pending | Do not report results |

## Contradictions and corrections

1. The manuscript said resident accounts and onboarding were conceptual or unavailable; the current mobile/backend integration exposes them. The working copy corrects this.
2. The manuscript said the structured multi-question DSS was not verified; it is now integrated and covered by current mobile tests. The working copy corrects this while preserving content-validation limits.
3. The manuscript said residents could not discover nearby centers; current endpoints and resident interfaces provide nearest-center lookup. The correction does not imply official center data, live capacity, or safe routing.
4. The manuscript described `flutter_map`/Leaflet as the sole map approach; the repository now includes a provider-aware Mapbox path with OSM/flutter_map fallback and Mapbox-enabled Admin maps.
5. The manuscript named Design Science Research and a generic iterative/incremental SDLC. Current records prove artifact development, surveys, interviews, testing, and revisions, but do not prove that the group deliberately followed the Peffers DSR method or a specific Scrum/Kanban/RAD process. The working copy uses a provisional developmental-descriptive research classification and marks specific SDLC selection pending.
6. Older feasibility material describes Firebase, an AI assistant, offline Firestore behavior, and other superseded concepts. These were not carried into the working copy because current integrated code uses Django/PostgreSQL/PostGIS and deterministic/structured DSS flows.
7. The BDRRMO consultation suggested a Low-yellow, Moderate-violet, High-orange, Very High-red scheme, whereas the manuscript/code use green/yellow/orange/red. The interview suggestion is not recorded as a formally approved standard; the matter remains pending.

## Evidence ready now versus provisional

Ready now: problem context from official literature and interviews; system scope boundaries; current implemented software functions; current mobile test count; research instruments; planned statistical procedures; provisional feasibility analysis; data-governance limitations.

Provisional: exact SOP wording affected by final survey evidence; final respondent profile and needs findings; final sampling/reporting; selected specific SDLC; approved susceptibility parameters and data; BDRRMO-approved DSS content and center dataset; final diagrams/appendix letters; ISO/IEC 25010 evaluation and conclusions; production claims.
