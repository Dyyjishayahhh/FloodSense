# FloodSense Development–Manuscript Alignment Tracker

## Audit record

- Audit date: 2026-09-20 (Asia/Manila)
- Development repository: `2. FLOODSENSE-DEVELOPMENT TEAM`
- Current commit: `fc62e47d808ec8762abfb25c23e338e2843e4a3f`
- Commit date: 2026-09-20 15:49:46 +08:00
- Commit message: `Updates`
- Previous audited commit: None; this is the baseline tracker audit.
- Working tree before documentation authoring: Clean; branch `main` matched `origin/main`.
- Working tree after documentation authoring: Expected untracked manuscript, tracker, and temporary authoring script; no source-code changes were made.
- Latest-commit change summary: The current commit introduced the primary development-team tree and its backend, mobile, tests, documents, research-data structure, diagrams, mockups, screenshots, instruments, and processing scripts while removing the superseded `2. FLOODSENSE-CODE AND FILES` tree. The full Git name-status output was reviewed during the baseline audit.
- Manuscript baseline: `1. FLOODSENSE-MAIN/MAIN FILESS/FloodSense_Chapter_1_MAIN.docx`
- Development-aligned manuscript: `1. FLOODSENSE-MAIN/MAIN FILESS/FloodSense_Chapters_1_to_3_Development_Aligned_2026-09-20.docx`
- Chapter 2 enhanced manuscript: `1. FLOODSENSE-MAIN/MAIN FILESS/FloodSense_Chapters_1_to_3_Chapter_2_Enhanced_2026-09-21.docx`
- Chapter 2 application-and-citation manuscript: `1. FLOODSENSE-MAIN/MAIN FILESS/FloodSense_Chapters_1_to_3_Chapter_2_Applications_and_Narrative_Citations_2026-09-21.docx`

## Chapter 2 revision record

- Revision date: 2026-09-21 (Asia/Manila).
- Development-status basis: Baseline audit at commit `fc62e47d808ec8762abfb25c23e338e2843e4a3f`; no later commit was present when the Chapter 2 revision began.
- Editing boundary: Only Chapter 2 and directly corresponding reference entries were revised in the new manuscript copy.
- Preservation check: The Chapter 1 and Chapter 3 XML-content hashes in the source and revised manuscripts match. The source manuscript and `FloodSense_Chapter_1_MAIN.docx` were not overwritten.
- Chapter 2 evidence added: Recent foreign and Philippine literature on susceptibility, rule-based systems, explainability, DSS, GIS, PostGIS, location privacy, flood-risk communication, provenance, validation, software quality, national disaster responsibilities, Bacoor administrative geography, and comparable flood studies or systems.
- Local studies added: National Philippine exposure modeling; Marikina-Pasig inundation modeling; Matina data-poor flood modeling; Santa Fe GIS/FAHP mapping; Davao vulnerability and risk perception; Metro Manila social-vulnerability assessment; Metro Manila flood-governance barriers; a Leyte flood-warning prototype; and a Philippine ISO/IEC 25010 application study.
- Scope controls retained: External weights, formulas, sensor functions, monitoring, forecasts, warnings, alerts, and study findings were not transferred to FloodSense. The 47-barangay layer remains an administrative reference, and no approved operational susceptibility dataset is claimed.
- Excluded source: The 2023 Imus River Basin Research Square item was excluded from the formal chapter because it is a non-peer-reviewed preprint.
- Tests rerun for this documentation revision: No Django or Flutter tests; this revision changed no source code.
- Document QA: Microsoft Word rendering completed on 2026-09-21. The 58-page manuscript was rendered, Chapter 2 pages 25-46 and the updated reference pages were visually inspected, and the tables, headings, margins, page breaks, citations, and reference layout were confirmed readable and within page boundaries.

## Chapter 2 application-and-citation revision record

- Revision date: 2026-09-21 (Asia/Manila).
- Source manuscript preserved: `FloodSense_Chapters_1_to_3_Chapter_2_Enhanced_2026-09-21.docx` (SHA-256 `81AE98E2A51DCA50613A3D0E03D7FA20AB67EA933D901B2F7AA5BE5EEC648AAD`).
- New output: `FloodSense_Chapters_1_to_3_Chapter_2_Applications_and_Narrative_Citations_2026-09-21.docx`.
- Editing boundary: Chapter 2 and its required reference entries only. Chapter 1 and Chapter 3 WordprocessingML content hashes match the source, all four document sections were retained, and neither the source manuscript nor `FloodSense_Chapter_1_MAIN.docx` was overwritten.
- Structural change: Related Foreign Studies and Related Local Studies were reorganized around ten named, verifiable applications or systems, with the responsible institution or authors, users, geographic scope, functions, limitations, relevance, similarities, differences, and non-transferable elements discussed for each.
- Foreign applications reviewed: FloodAdapt, GloFAS, Iowa Flood Information System, FEMA Resilience Analysis and Planning Tool, and England's Check the Long Term Flood Risk for an Area service.
- Philippine applications reviewed: HazardHunterPH, GeoMapperPH, PlanSmartPH Ready to Rebuild, UP NOAH Know Your Hazards/NOAH Studio, and the Bentoso et al. Leyte flood-warning decision-support prototype.
- Citation treatment: Narrative or author-prominent APA citations were made the dominant pattern in the application discussions. Official and primary sources were used for institutional systems; the full Bentoso et al. conference paper was inspected for the Leyte prototype.
- Scope protection: None of the comparators' forecasting, monitoring, sensor, alerting, simulation, recovery-planning, national-dataset, or official-authority functions were transferred to FloodSense. Fixed-schema CSV exchange remains a confirmed but unimplemented requirement, and no approved operational Bacoor susceptibility dataset is claimed.
- Document QA: Microsoft Word rendering completed after two layout corrections. The final manuscript contains 59 physical pages; Chapter 2 occupies physical pages 25-46 and references occupy physical pages 56-59. Every affected page was visually inspected at full-page resolution. The three comparison tables each begin on a separate page, remain within the margins, contain no split data rows, and use readable repeated visual conventions consistent with the SURECUT model.

| Feature | Development status | Code evidence | Document evidence | Commit | Manuscript section | Documentation status | Required action | Permission |
| ------- | ------------------ | ------------- | ----------------- | ------ | ------------------ | -------------------- | --------------- | ---------- |
| Foreign application comparisons | OUTSIDE THE CURRENT SCOPE | None; the applications are literature comparators, not FloodSense functions | Official Deltares, Copernicus/ECMWF, Iowa Flood Center, FEMA, and Environment Agency sources | `fc62e47` | Ch. 2 Related Foreign Studies; Tables 1 and 3 | Added with narrative institutional citations and explicit non-transferability limits | Retain as comparators; do not infer FloodSense forecasting, sensors, simulation, or official authority | None for current literature treatment |
| Philippine application comparisons | OUTSIDE THE CURRENT SCOPE | None; the applications are literature comparators, not FloodSense functions | GeoRisk Philippines, DOST-PHIVOLCS, UP NOAH, and Bentoso et al. sources | `fc62e47` | Ch. 2 Related Local Studies; Tables 2 and 3 | Added with distinct developer, user, scope, function, limitation, and FloodSense comparison discussions | Retain provenance and authority distinctions; do not import datasets or features without separate approval | Data, methodology, and institutional approval required before any future integration |
| Chapter 2 citation and reference alignment | IMPLEMENTED AND CURRENTLY VERIFIED | Not applicable to source code | Verified sources and rendered manuscript | `fc62e47` | Ch. 2 and corresponding reference entries | Narrative citation pattern applied; ten application sources aligned with reference entries; 44 unique references retained | Reverify URLs and publication details if the chapter is updated after a later repository or source change | None for current documentation revision |

## Test and evidence record

- Tests rerun: Full Django and Flutter suites were not rerun.
- Historical test evidence: Backend and mobile test modules are present for assessment, rule behavior, geographic resolution, map behavior, permissions, DSS workflows, Admin workflows, provenance, and models.
- Tests completed during the audit: All 117 Python source files passed abstract-syntax-tree parsing.
- Unavailable test tools: Django, Django REST Framework, pytest, psycopg, Flutter, and Dart were unavailable in the documentation runtime.
- Approved data: No approved operational susceptibility dataset is present; `research_data/approved/README.md` is instructional only.
- Provisional data: Contents of `research_data/provisional/` are unvalidated, restricted, fictional, incomplete, or awaiting approval. The provisional MGB material is not established as operational system data.
- Administrative reference data: The derived 47-barangay boundary layer passed documented technical geometry checks but is not established as City-issued or City-verified and contains no flood-susceptibility facts.
- Survey evidence: The initial 50-resident requirements survey and key-informant records are historical requirements evidence. The current survey package is an instrument and contains no validated response set or evaluation result.
- Missing evidence: Approved susceptibility values, authoritative rainfall parameters for operational use, current official center data for resident discovery, completed CSV workflow, full current test execution, validated survey responses, ISO/IEC 25010 evaluation results, and production-deployment evidence.

## Alignment table

| Feature | Development status | Code evidence | Document evidence | Commit | Manuscript section | Documentation status | Required action | Permission |
| ------- | ------------------ | ------------- | ----------------- | ------ | ------------------ | -------------------- | --------------- | ---------- |
| Rainfall-intensity options | IMPLEMENTED AND CURRENTLY VERIFIED | `mobile/lib/features/assessment/widgets/scenario_selector.dart`; backend scenario models/serializers | Day 5 mobile guide; current decision documents | `fc62e47` | Ch. 1 Scope; Ch. 3 Mobile Application | Added as five hypothetical options | Confirm authoritative thresholds before operational use | Data/source approval required |
| Rainfall-duration options | IMPLEMENTED AND CURRENTLY VERIFIED | Assessment widgets and request models; backend serializers | Day 5 guide | `fc62e47` | Ch. 1 Scope; Ch. 3 Mobile Application | Added as 1, 3, 6, 12, and 24 hours | Verify combinations against approved parameters | Data/source approval required |
| Scenario confirmation and explicit assessment request | IMPLEMENTED AND CURRENTLY VERIFIED | `assessment_screen.dart`; `assessment_controller.dart` | Mobile implementation guides | `fc62e47` | Ch. 1 Introduction/Scope; Ch. 3 Mobile Application | Added | Retain pre-event, user-initiated boundary | None for current description |
| Manual geographic-area selection | IMPLEMENTED AND CURRENTLY VERIFIED | `zone_selector.dart`; geographic models/API client | Mobile guides | `fc62e47` | Ch. 1 Scope; Ch. 3 Mobile Application | Added | Retest after geographic-data changes | None |
| Manual map pin | IMPLEMENTED AND CURRENTLY VERIFIED | `dynamic_map_card.dart`; location controller | GPS/location guides | `fc62e47` | Ch. 1 Scope; Ch. 3 Mobile Application | Added | Maintain confirmation before assessment | None |
| One-time foreground GPS | IMPLEMENTED AND CURRENTLY VERIFIED | `location_service.dart`; Android manifest; location controller | GPS and safe-improvements plan; handoff notes | `fc62e47` | Ch. 1 Scope/Privacy; Ch. 2 Foreground Location; Ch. 3 Mobile Application | Chapter 2 expanded with current Android foreground/one-time permission guidance and privacy limitations | Verify device-level permission behavior | None |
| Barangay detection | IMPLEMENTED AND CURRENTLY VERIFIED | `server/geography/services.py`; resolve endpoint; mobile location flow | `RESOLVE_BARANGAY_CONTRACT.md` | `fc62e47` | Ch. 1 Scope; Ch. 3 Backend/Mobile | Added | Retest against approved boundaries | Boundary approval required for operational use |
| Barangay confirmation or correction | IMPLEMENTED AND CURRENTLY VERIFIED | Mobile location state/controller/UI | Location handoff documents | `fc62e47` | Ch. 1 Scope; Ch. 3 Mobile Application | Added | Preserve user correction and explicit confirmation | None |
| Neutral 47-barangay reference display | IMPLEMENTED AND CURRENTLY VERIFIED | Reference boundary map widget; geography endpoint; processed GeoJSON | Administrative-boundary README and mapping file | `fc62e47` | Ch. 1 Scope/Definitions; Ch. 3 Mobile/Data Governance | Added with technical/institutional distinction | Obtain City verification before describing as official | Institutional approval required |
| Demonstration dynamic map coloring | IMPLEMENTED AND CURRENTLY VERIFIED | `dynamic_map_card.dart`; map assessment API | Day 6 map guide; screenshots | `fc62e47` | Ch. 1 Introduction/Scope; Ch. 3 Mobile Application | Added as demonstration behavior | Replace demonstration inputs only after approved data exist | Data approval required |
| Deterministic inference | IMPLEMENTED AND CURRENTLY VERIFIED | `server/expert/services.py`; expert views/serializers/models | Day 3 RBES guide; architecture decision | `fc62e47` | Ch. 1 Introduction/Scope; Ch. 2 Rule-Based Expert Systems/Explainability; Ch. 3 Expert System | Chapter 2 expanded with verified literature on knowledge bases, rules engines, deterministic processing, and explainability | Preserve deterministic boundary | None |
| Rule matching, priorities, and conflict handling | IMPLEMENTED AND CURRENTLY VERIFIED | Expert service and tests | RBES guide | `fc62e47` | Ch. 3 Expert System/Testing | Added | Rerun full tests in backend environment | None |
| Susceptibility result and explanation | IMPLEMENTED AND CURRENTLY VERIFIED | Assessment result model/card; expert API | Day 5/6 guides | `fc62e47` | Ch. 1; Ch. 3 | Added with demonstration/data limit | Validate explanation with approved knowledge | Expert/data approval required |
| DSS published preparedness guidance | IMPLEMENTED AND CURRENTLY VERIFIED | `server/dss/services.py`; guidance API; mobile guidance section | Day 4 DSS guide | `fc62e47` | Ch. 1; Ch. 2 DSS/Preparedness Guidance; Ch. 3 DSS | Chapter 2 expanded and retains the separation between assessment and preparedness guidance | Validate and publish approved guidance | Content approval required |
| Resident structured multi-question DSS flow | ACTIVE WORK IN PROGRESS | DSS models/services exist; complete resident interaction not verified | Designs and specifications show intended flow | `fc62e47` | Ch. 3 DSS | Identified as incomplete | Complete interface, integration, and tests | Approval required if scope changes |
| Custom Admin authentication | IMPLEMENTED AND CURRENTLY VERIFIED | Admin portal views, URLs, login template, permissions | Admin frontend guides | `fc62e47` | Ch. 1 Scope; Ch. 3 Admin | Added | Rerun permission tests | None |
| Custom Admin dashboard | IMPLEMENTED AND CURRENTLY VERIFIED | Dashboard template/service/tests | Admin guides | `fc62e47` | Ch. 1 Scope; Ch. 3 Admin | Added | Maintain evidence-based status cards | None |
| Map-data review | IMPLEMENTED AND CURRENTLY VERIFIED | Admin map-data service/view/template/tests | Admin Day guides | `fc62e47` | Ch. 1 Scope; Ch. 3 Admin | Added as read-only review | Do not describe as unrestricted map editor | None |
| Rainfall-reference review | IMPLEMENTED AND CURRENTLY VERIFIED | Admin settings/rainfall templates and services | Parameter governance proposal/guide | `fc62e47` | Ch. 1 Scope; Ch. 3 Admin | Added as review | Link to approved source when available | Source approval required |
| Parameter governance | IMPLEMENTED AND CURRENTLY VERIFIED | Admin settings views/services/tests | Parameter governance proposal and current decisions | `fc62e47` | Ch. 1 Scope; Ch. 3 Admin | Added with limits | Keep raw rule editing excluded | None |
| Ordinary Admin raw rule editing | OUTSIDE THE CURRENT SCOPE | No ordinary custom-portal rule editor; rule models appear in technical Django Admin | Confirmed TA instruction/current decision | `fc62e47` | Ch. 1 Introduction/Scope; Ch. 3 Admin | Corrected outside Theoretical Framework | Retain distinction from technical Django Admin | Confirmed correction already authorized |
| Technical Django Admin rule maintenance | IMPLEMENTED BUT NOT EXPOSED TO THE INTENDED USER | `server/expert/admin.py`; Django Admin registration | Current decision documents | `fc62e47` | Ch. 3 Admin | Added as developer/research maintenance only | Apply strict staff permissions and audit procedure | Institutional/developer authorization required |
| Fixed-schema CSV export | CONFIRMED REQUIREMENT — NOT YET IMPLEMENTED | No verified route, service, validation, UI, integration, or test | Confirmed TA instruction | `fc62e47` | Ch. 1 Scope/Definitions; Ch. 3 CSV Requirement | Added as future requirement | Implement canonical export schema and permissions | Implementation and validation approval required |
| Fixed-schema CSV import and validation | CONFIRMED REQUIREMENT — NOT YET IMPLEMENTED | No verified route, service, schema/value validation, UI, integration, or test | Confirmed TA instruction | `fc62e47` | Ch. 1 Scope/Definitions; Ch. 3 CSV Requirement | Added as future requirement | Implement rejection rules, audit, rollback, and tests | Implementation and validation approval required |
| DSS-content workflow | IMPLEMENTED AND CURRENTLY VERIFIED | Admin guidance forms/views/workflow/tests; DSS migrations | Admin and DSS guides | `fc62e47` | Ch. 1 Scope; Ch. 3 Admin/DSS | Added | Validate content authority and publication states | Content approval required |
| Evacuation-center Admin workflow | IMPLEMENTED AND CURRENTLY VERIFIED | Evacuation models/workflow; Admin forms/views/templates | Admin guides | `fc62e47` | Ch. 1 Scope; Ch. 3 Admin | Added | Keep separate from resident discovery claim | Record approval required |
| Resident nearest-center discovery | NOT IMPLEMENTED | No complete resident interface/API integration found | Mockups and specifications only | `fc62e47` | Ch. 1 Scope; Ch. 3 Status | Corrected as unavailable | Implement and verify before documenting as operational | Scope/data approval required |
| Data-source workflow and provenance | IMPLEMENTED AND CURRENTLY VERIFIED | Provenance models/policies/workflow; Admin data-source views | Architecture/current decisions | `fc62e47` | Ch. 1; Ch. 2 Provenance/Geographic Validation; Ch. 3 Data Governance | Chapter 2 expanded with GIS provenance, risk-informed data, and Philippine hazard-information limitations | Populate only supported sources and statuses | Record-specific approval required |
| Audit history | IMPLEMENTED AND CURRENTLY VERIFIED | Admin audit-history template/view and workflow records | Admin guides | `fc62e47` | Ch. 1 Scope; Ch. 3 Admin | Added | Confirm retention and access policy | Institutional policy required |
| Resident accounts and login | NOT IMPLEMENTED | Account model exists; no current resident-facing login/registration flow | Mockups/specifications show intention | `fc62e47` | Ch. 1 Scope/Definitions; Ch. 3 Status | Corrected as unavailable | Implement end-to-end flow before claiming availability | Scope/privacy approval required |
| Assessment history | NOT IMPLEMENTED | No complete resident history interface/integration verified | Designs/specifications only | `fc62e47` | Ch. 3 Status/Limitations | Added as missing | Implement only if retained in scope | Permission required for scope addition |
| Notifications or autonomous alerts | OUTSIDE THE CURRENT SCOPE | No background alert implementation | Some mockups/plans may imply future behavior | `fc62e47` | Ch. 1 Scope; Ch. 3 Limitations | Explicitly excluded | Do not add without scope approval | Explicit permission required |
| Offline use | OUTSIDE THE CURRENT SCOPE | No complete offline data/cache flow | Offline-state mockup only | `fc62e47` | Ch. 1 Scope; Ch. 3 Limitations | Explicitly excluded | Do not claim from mockup | Explicit permission required |
| Official susceptibility-data integration | BLOCKED OR AWAITING DATA | Approved directory contains no dataset; seed/demo data only | Approved/provisional READMEs | `fc62e47` | Ch. 1 Limitations; Ch. 3 Data Governance | Added as blocker | Obtain provenance, approval, methodology, import, and validation evidence | Institutional and expert approval required |
| Provisional MGB processing | PROPOSED — AWAITING APPROVAL | Extraction script and provisional README; no operational integration proven | Provisional data documentation | `fc62e47` | Ch. 3 Data Governance | Not described as operational | Confirm license, attribution, method, approval, and import | Explicit institutional/data approval required |
| Initial requirements survey | IMPLEMENTED — HISTORICAL TEST EVIDENCE ONLY | Research records rather than source code | Initial survey summary; Chapter 1 | `fc62e47` | Ch. 1 Problem; Ch. 3 Participants | Preserved as historical requirements evidence | Verify source package for final appendices | Research approval as applicable |
| New survey and user evaluation | ACTIVE WORK IN PROGRESS | No result data | Survey instrument/package only | `fc62e47` | Ch. 3 Participants/Evaluation | Author note inserted; no results fabricated | Complete collection and validated analysis | Research/TA approval required |
| ISO/IEC 25010 system evaluation | NOT IMPLEMENTED | No completed evaluation dataset or scores | Planned objective/instrument | `fc62e47` | Ch. 2 Software Quality; Ch. 3 Evaluation Plan | Chapter 2 literature expanded; evaluation remains future method only and no scores are reported | Finalize approved characteristics, instrument, sample, and analysis | TA/institutional approval required |
| Production deployment | NOT IMPLEMENTED | No verified production release/deployment evidence | Setup guides only | `fc62e47` | Ch. 3 Limitations | Added as unavailable | Complete security, data, test, and deployment readiness | Institutional authorization required |

## Unresolved contradictions

1. The original Theoretical Framework and System Architecture text describes ordinary administrator rule maintenance and several resident functions as though they are available. It was deliberately preserved unchanged because the user requires separate explicit permission before any Theoretical Framework change.
2. Original objectives and research questions retain intended functions that exceed the current prototype. The development-aligned Scope, Chapter 3 status table, and limitations distinguish the current state without converting each implementation change into a new objective.
3. Mockups and diagrams include resident accounts, nearest-center discovery, notifications, offline states, map editing, and rule testing that are not proven as integrated current functions.

## Pending permissions

- Explicit permission to revise the Theoretical Framework/System Architecture passage so it matches the confirmed Admin boundary and current implementation state.
- Institutional approval and provenance for operational geographic, susceptibility, rainfall, preparedness, and evacuation-center data.
- Approval of the fixed-schema CSV specification before implementation and operational documentation.
- Approval of the final survey/evaluation instrument, sampling procedure, statistical treatment, and completed analysis.
- Separate authorization for any source-code, diagram, dataset, approved-document, or original-manuscript modification.
