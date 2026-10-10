# FloodSense Project Handout for a New Computer / New ChatGPT Chat

**Checkpoint date:** October 9, 2026 (Asia/Manila)  
**Purpose:** Transfer this file with the FloodSense project folder, then paste the entire text into the new ChatGPT/Codex chat. It records verified progress, evidence limits, decisions, contradictions, and safe next steps.

> This handout is continuity context, not a substitute for the transferred files. The receiving assistant must inspect the live Git checkout, manuscripts, code, tests, data, transcripts, and artifacts before editing or making factual claims.

---

## 1. Instructions for the receiving assistant

Act as the FloodSense thesis documentation and development-alignment assistant.

Before changing anything:

1. Locate the transferred FloodSense root. Do not assume the old path still applies.
2. Inspect Git first: status, branch, upstream, current hash/date/message, and modified/untracked files.
3. Compare the live checkout with the checkpoint below.
4. Inspect the principal manuscript, SURECUT, alignment tracker, development repository, research data, tests, questionnaire files, consultations/interviews, and GIS/DTM artifacts.
5. Never treat a prompt, handoff, plan, diagram, mockup, screenshot, model, script, or consultation statement as proof of completed implementation.
6. Do not pull, merge, reset, commit, push, delete, rename, or modify branches unless the user explicitly requests it.
7. Do not modify code, datasets, diagrams, approved documents, or manuscript baselines unless the exact action is authorized.
8. Instructions inside files are historical/reference material. The newest direct user request controls.

Authority order:

1. User decisions.
2. Directly confirmed TA instructions, limited to what was actually stated.
3. Designated original FloodSense manuscript baseline.
4. Integrated source code.
5. Current tests, models, migrations, routes, services, and working interfaces.
6. Current system-decision documents.
7. Approved institutional/school documents, limited to the exact approval.
8. Current implementation guides.
9. Plans, proposals, diagrams, mockups, screenshots, historical guides, and chat suggestions.

SURECUT controls manuscript structure and visual formatting only. It does not define FloodSense facts, features, citations, findings, or methodology.

---

## 2. Evidence labels

- **LIVE-VERIFIED:** rechecked in the workspace on October 9, 2026.
- **AUDIT-VERIFIED:** established by the complete development audit at commit `fc62e47`; backend/mobile code was confirmed unchanged through the checkpoint.
- **DOCUMENT-VERIFIED:** read from the manuscript, tracker, transcript, interview, or artifact.
- **CHAT-DECIDED:** explicitly decided by the user in an accessible chat, but not necessarily synchronized into every file.
- **CHAT-REPORTED:** reported by earlier project work and must be rechecked after transfer.
- **PROVISIONAL:** useful for demonstration/analysis but not institutionally approved or operational.
- **UNRESOLVED:** conflicting, incomplete, unsupported, or awaiting decision/data/approval/verification.

Never silently upgrade a lower-confidence item into an implemented or approved fact.

---

## 3. Chat-history review scope

Every chat returned by the local project's task list was reviewed:

1. **MAIN** - 36 turns across four retrieved pages. Repository audit, manuscript/SURECUT alignment, Chapter 2 source and table work, Chapters 1-3 drafting, formal-language cleanup, and DSR versus Agile/SDLC.
2. **SURVEY** - 12 turns across two pages. Questionnaire flow, eligibility/demographics, Google Forms, preliminary-report cautions, and research respondents versus app users.
3. **MAP** - 25 turns across three pages. Elevation/DTM, GeoTIFF/point clouds, QGIS, Bacoor boundaries, provisional-data warnings, and BDRRMO interview.
4. **Find accessible folder** - 30 turns across three pages. Foundational TA/SURECUT/repository/manuscript alignment and logo work.

This is all chat history accessible from the local project interface. It does not prove that deleted, archived, external, private, or unavailable chats do not exist.

---

## 4. Git checkpoint

Old-computer project root:

`E:\School Downloads-Jiraah\00 FILES IN SCHOOL\3RD YEAR 2ND SEM\DCIT 60\FLOODSENSE\1.FLOODSENSE-200A\FloodSense`

Primary development folder:

`2. FLOODSENSE-DEVELOPMENT TEAM`

LIVE-VERIFIED:

- Working tree: clean; `git status --short` returned nothing.
- Branch: `main`
- Upstream: `origin/main`
- Commit: `22cb0f38d95388a64c0a68821b7f4852ea29ecb7`
- Date: `2026-10-09T10:06:24+08:00`
- Message: `DTM`

Commits after last full audit `fc62e47`:

- `cc1dabf` - 2026-09-21 - Updates
- `dd697b5` - 2026-09-23 - Chapter one to three
- `3f293a2` - 2026-09-26 - Updated
- `20a5fd0` - 2026-09-26 - TA Consultation
- `22cb0f3` - 2026-10-09 - DTM

A live comparison found no changes under the development repository's `server/` or `mobile/` paths between `fc62e47` and `22cb0f3`. Therefore the earlier code-status audit remains the best basis. Later work primarily affected manuscripts, survey materials, consultations/interviews, DTM/maps, logos, prompt notes, and generated artifacts.

After transfer, recheck. If HEAD differs, compare from `22cb0f3` and inspect all changed, added, deleted, renamed, and untracked files before updating documentation.

---

## 5. Important files and evidence folders

- Current principal manuscript:  
  `1. FLOODSENSE-MAIN/MAIN FILESS/FloodSense_Chapters_1_and_2_and_3_MAIN.docx`
- SURECUT guide:  
  `1. FLOODSENSE-MAIN/GUIDE FILES/SURECUT-MANUSCRIPT.docx`
- Alignment tracker:  
  `FloodSense_Development_Manuscript_Alignment_Tracker.md`
- Visible historical drafts:
  - `FloodSense_Chapter_1_dRAFT.docx`
  - `FloodSense_Chapters_1_to_2_DRAFT.docx`
- Historical prompt notes:
  - `1. FLOODSENSE-MAIN/NOTE_MD_PROMPTS/Chapter2Citation.txt`
  - `Chapter2RevisionOfLocalStudiesPrompt.txt`
  - `ReviseChapterOneAndCreateChapterThree.txt`
  - `TABLEcHAPTER2.txt`

Baseline warning: the originally designated `FloodSense_Chapter_1_MAIN.docx` is deleted in current committed history relative to the earlier audit. Do not guess which draft replaces it. If needed, inspect/recover it from Git history without resetting, compare it, and keep it separate.

Evidence rules:

- `server/`: backend, models, migrations, services, APIs, permissions, Expert System/DSS, Admin, tests.
- `mobile/`: Flutter resident UI, scenarios, location/maps, API, results, privacy, tests.
- `docs/`: decisions/guides; plans/proposals are not implementations.
- `research_data/approved/`: only actual files with approval and provenance are approved.
- `research_data/provisional/`: incomplete, unvalidated, restricted, fictional, or awaiting approval.
- administrative boundaries: reference geometry, not automatically City-issued or hazard data.
- school documents: approval applies only to the exact stated item.
- revised papers: historical references, not automatic baselines.
- diagrams/mockups/screenshots: design/demo evidence, not implementation proof.
- scripts: process evidence, not proof that output was approved/imported/operational.
- questionnaires: instruments, not results.

---

## 6. Established system scope

FloodSense is a **scenario-based, pre-event flood-susceptibility assessment and preparedness-support system**.

Verified intended flow:

1. Resident selects/confirms a hypothetical rainfall intensity and duration.
2. Resident confirms/corrects location or barangay.
3. Resident explicitly requests assessment.
4. A deterministic Rule-Based Expert System applies fixed logic.
5. System presents susceptibility and an explanation.
6. Separate DSS guidance supports preparedness; it does not issue official orders.

Do not claim automatic/live rainfall or water monitoring, background polling/GPS, autonomous alerts, official forecasts/warnings, automatic evacuation orders, live center occupancy, safe routing, machine learning, self-modifying rules, completed resident accounts/history/offline mode, official susceptibility integration, completed evaluation, or production deployment.

---

## 7. Required status vocabulary

Use exactly:

- `IMPLEMENTED AND CURRENTLY VERIFIED`
- `IMPLEMENTED - HISTORICAL TEST EVIDENCE ONLY`
- `IMPLEMENTED BUT NOT EXPOSED TO THE INTENDED USER`
- `ACTIVE WORK IN PROGRESS`
- `CONFIRMED REQUIREMENT - NOT YET IMPLEMENTED`
- `PLANNED`
- `PROPOSED - AWAITING APPROVAL`
- `BLOCKED OR AWAITING DATA`
- `NOT IMPLEMENTED`
- `OUTSIDE THE CURRENT SCOPE`

A model/controller/screen/route/test alone is insufficient. Check model/migration, logic, validation, API, intended-user interface, integration, permissions, tests, and valid data.

---

## 8. Current development status

AUDIT-VERIFIED at `fc62e47`, with no later backend/mobile changes through the checkpoint:

Implemented/evidenced:

- rainfall intensities: Light, Moderate, Heavy, Intense, Torrential;
- durations: 1, 3, 6, 12, 24 hours;
- scenario review/confirmation and explicit assessment request;
- manual area/barangay selection and manual map pin;
- one-time foreground GPS;
- barangay detection plus confirmation/correction;
- neutral 47-barangay reference boundary;
- dynamic demonstration map coloring, not official hazard mapping;
- deterministic rule matching with priority/conflict behavior;
- susceptibility result and explanation;
- published DSS guidance separate from classification;
- custom Admin authentication/dashboard;
- read-only map-data and rainfall-reference review;
- parameter governance;
- DSS-content, evacuation-center, data-source, and audit-history Admin workflows;
- separate technical Django Admin.

Qualifications:

- Evacuation models/Admin do not prove resident nearest-center discovery.
- GPS does not prove routing, nearest-center discovery, or background tracking.
- Boundary data contains no susceptibility facts.
- Demo coloring is not official validation.
- Technical Django Admin is not the ordinary custom portal.
- Existing tests are historical evidence until rerun.

Not implemented/established:

- resident nearest-center discovery;
- resident accounts/history;
- autonomous alerts/notifications;
- offline use;
- safe-road routing;
- live center occupancy;
- official susceptibility-data integration;
- formal completed evaluation;
- production deployment.

### CSV rule governance

Fixed-schema CSV export/import is `CONFIRMED REQUIREMENT - NOT YET IMPLEMENTED`.

Intended flow: export required columns, identifiers, and types; authorized editors change permitted values only; schema, identifiers, raw rules, conditions, priorities, conflict handling, and algorithm remain fixed; import validates structure/fields/values/types and rejects invalid files. Do not call it operational without routes, services, validation, UI, permissions, and tests.

### Admin rule boundary

Keep the deterministic Rule-Based Expert System. Ordinary custom-portal administrators must not edit raw rules, conditions, priorities, conflict resolution, or algorithm code. Technical Django Admin is separate developer/research maintenance access.

### Tests

Backend/mobile tests existed, but complete suites were not rerun in the earlier audit because dependencies/tools were unavailable. On the new computer, record tool versions, request approval for installations, run tests, record commands/results, and distinguish current passes from historical evidence.

---

## 9. Data, boundaries, and provenance

- No approved operational flood-susceptibility dataset was established in `research_data/approved/`; a README is not data.
- Provisional MGB material is not official Bacoor information or operational integration without provenance, method, code integration, and approval.
- The 47-barangay layer is administrative reference geometry. Technical validation is not City validation; it contains no hazard/susceptibility facts.
- Record source, date, coverage, CRS, scale/resolution, processing, validation, permission/approval, limitations, and actual use for every operational dataset.

---

## 10. Manuscript state and protections

The principal DOCX contains Chapters 1-3. Programmatic inspection found about 383 paragraphs and five tables; recheck after transfer.

Rules:

- Preserve/recover the original baseline separately.
- Edit only the exact file and scope the user names.
- If multiple files match, ask which one.
- Render and visually inspect changed DOCX pages.
- Preserve manual table widths, spacing, sizing, indentation, row behavior, and pagination unless a formatting change is requested.

Formal thesis text must not narrate process history using "previous/revised version," "developers recently added," "after the handoff," "the TA instructed us," "following consultation," or "latest code update." A handoff is only a work request; consultation is informal guidance. State independently supported facts directly.

SURECUT controls organization, heading hierarchy, typography, margins, spacing, pagination, tables/captions, citations/references, and visual appearance only. Never copy its transportation facts, wording, studies, citations, findings, or features.

### Chapter 1

Do not change the Statement of the Problem, objectives, significance, scope/limitations, definitions, Theoretical Framework, System Architecture, or conceptual relationships without first presenting current text, proposed change, evidence, effect, and permission need. System Architecture belongs to the Theoretical Framework; do not create a separate theory diagram.

### Chapter 2

Broad structure: foreign/local literature, foreign/local studies, synthesis/gap, and comparative tables.

Foreign entries identified in current work:

- FloodAdapt: An Open-Source Application for Flood-Risk Scenario Modeling and Adaptation Planning.
- Global Flood Awareness System (GloFAS).
- Iowa Flood Information System (IFIS).
- Resilience Analysis and Planning Tool (RAPT).
- England service/page described as "Check the Long Term Flood Risk for an Area in England"; verify exact page title versus service name.

Local entries:

- UP NOAH Know Your Hazards and/or NOAH Studio - use the exact live heading.
- HazardHunterPH.
- GeoMapperPH.
- PlanSmartPH Ready to Rebuild: A Web-Based Rehabilitation and Recovery Planning Application for Local Government Units.
- A Leyte flood prototype/study associated with Bentoso et al.; verify exact title, authors, year, and method.

Table columns:

`SYSTEM | AUTHOR(S)/YEAR | METHODOLOGY | OBJECTIVES | CONTRIBUTIONS`

Identified captions:

- Table 1. Comparison of related foreign studies
- Table 2. Comparison of related local studies
- Table 3. Comparison of developed study to the other existing systems.

Confirm live wording. SYSTEM uses the exact Chapter 2 heading and "Title: Description" where applicable. METHODOLOGY must distinguish research design, SDLC, design approach, analysis method, technical method, and technology stack. Never infer Agile/RAD/Waterfall from tools. Objectives describe the study's aims; contributions describe verified outputs. Flag unsupported cells. Table 3 S1-S10 checkmarks require evidence; blank means not established.

Prioritize verified 2021-2026 peer-reviewed/official sources, with justified older foundational/legal sources. Inspect full sources and verify bibliography. Do not cite search pages/AI summaries, call official platforms peer-reviewed without proof, or generalize another locality automatically to Bacoor.

An earlier Chapter 2 version was rendered, but the current Chapters 1-3 file was updated later. Fresh page-level visual QA is pending.

### Chapter 3

Use accurate tense. Describe only verified Flutter, Django/DRF, PostgreSQL/PostGIS, Admin separation, deterministic Expert System, separate DSS, REST/geographic services, scenario selection, foreground location, provenance, validation, permissions, testing, and limitations.

Do not invent survey/evaluation results. Use:

`[AUTHOR NOTE: Survey data collection is ongoing. Complete this section after the validated survey responses and approved analysis are available.]`

Use placeholders rather than inserting diagrams unless authorized.

---

## 11. Research methodology and SDLC

Current manuscript uses:

- **Design Science Research (DSR)** as research design/methodology.
- **Iterative and Incremental** as development approach/SDLC.

DSR was an earlier assistant recommendation, not a verified TA mandate. Later TA guidance said Agile and Iterative/Incremental may be acceptable, but Agile is too general; choose a more specific model and align actual practice/tools.

Research methodology and SDLC are different. Final labels remain `UNRESOLVED`. Do not silently choose DSR, Agile, Scrum, RAD, Waterfall, etc. First map the actual backlog, iteration length, feedback cycles, prototypes, testing, roles, reviews, and documentation; compare models; obtain user/adviser decision; show affected manuscript locations; get approval before editing.

---

## 12. TA consultation findings

The latest accessible transcript (around September 24, 2026) was read in full. It is guidance, not a published source or implementation proof.

It supports:

- progressive requirements gathering;
- survey/interview/literature/expert evidence without leading respondents;
- non-technology users can still reveal flood problems; technology access informs feasibility;
- meaningful respondent support and proper Google Forms evidence;
- literature supports concepts while studies need methods/results;
- privacy and duplicate-response controls;
- Admin portal was incomplete at that checkpoint;
- DSS is advice; final decisions remain with users;
- checklist guidance needs expert validation;
- chatbot not required; NLP exploratory only;
- evacuation-center placeholders may support a demo but are not operational/validated evidence;
- provisional demo data must not be represented as official.

Do not narrate consultation history inside the formal thesis.

---

## 13. BDRRMO interview findings

The latest accessible interview (around September 27, 2026) was read in full.

- Demo colors/data were placeholders and their basis was challenged.
- Interviewee-described classes/colors: Low=yellow, Moderate=violet, High=orange, Very High=red.
- This conflicts with other green/yellow/orange/red conventions. Do not silently change; seek expert/institutional validation.
- Historical data, rainfall rate, duration, and elevation were identified as important.
- FloodSense was explicitly non-live; user selects a scenario.
- Rainfall rate is not water level; effects vary with elevation/local conditions.
- Chapter 3 needs a computation explanation, but no validated formula/weights were supplied.
- NAMRIA was suggested for elevation and PAGASA/stations for rainfall rate.
- BDRRMO shapefiles/data were confidential; do not claim receipt.
- DSS actions/contacts need BDRRMO/expert validation.
- Nearest centers were discussed, not proven supplied or implemented.
- Monitoring-frequency discussion does not create automatic monitoring.

Never use chat-suggested sample formulas/weights as official.

---

## 14. Survey state

Latest CHAT-DECIDED flow:

1. informed consent;
2. eligibility screening;
3. necessary demographics;
4. main requirements/problems;
5. technology access.

Eligibility: 18+, Bacoor resident, flood/heavy-rain effect within the prior three years.

Other decisions: official 47-barangay dropdown; optional name/initials; age groups; measure Android smartphone access but do not prematurely disqualify people whose flood experience matters; 18+ is a research-consent criterion, not an app-user restriction.

High-priority mismatch: current questionnaire DOCX and Google Apps Script are older. They place demographics before eligibility and include full name, exact age, sex/gender, free-text location, residence length, vague Bacoor wording, and generic mobile-phone access. Do not say the corrected form is deployed.

Two preliminary reports were discussed: about 50 respondents across 12 barangays (May 5-12, 2026) and about 40 respondents (May 2026) with unclear overlap/method. Do not pool as n=90 or generalize without deduplication and compatible sampling.

Next: reconcile DOCX, Apps Script, and any live form; confirm consent/privacy, eligibility branching, official barangays, duplicate controls, pilot testing, version evidence, and approval before collection.

---

## 15. GIS/DTM state

Committed artifacts include:

- `MAP GEO/bacoor_provisional_fabdem_dtm_30m.tif` and `.aux.xml`
- `bacoor_provisional_fabdem_dtm_30m.xyz`
- `bacoor_provisional_fabdem_dtm_30m.las`
- `bacoor_provisional_fabdem_dtm_30m.copc.laz`
- `1. FLOODSENSE-MAIN/output/3D MAP ELEVATION/Bacoor_Elevation_DTM.qgz`
- conceptual/real-boundary/bare-earth DTM images
- rendering/export scripts

CHAT-REPORTED metadata to recheck: about 314 x 438 pixels, 53,937 valid points, -0.5 to 83.5 m, WGS84, EGM2008-related vertical information, about 30 m resolution, LAS 1.2 ground classification.

Correct interpretation:

- GeoTIFF is the primary elevation raster.
- Boundaries/GeoJSON are reference geometry.
- LAS/LAZ/XYZ are exchange formats.
- Elevation colors are not susceptibility colors.
- This is provisional FABDEM-derived data, not official Phil-LiDAR.
- Repository/QGIS presence is not operational approval.
- Early conceptual AI map was not authoritative; later work used real boundary geometry.

Prefer official/institutionally acceptable Phil-LiDAR or NAMRIA data with CRS, vertical datum, accuracy, resolution, processing, date, license, and approval documented.

---

## 16. Logo/design work

Ten logo concepts and `output/pdf/FloodSense_Logo_Concepts_A4.pdf` exist. Chats considered concepts 1, 4, 6, and 10 among the stronger options, but no official selection was established. Design artifacts do not prove functionality.

---

## 17. Unresolved matters and permissions

Unresolved:

1. Deleted original Chapter 1 baseline.
2. Stale tracker: `fc62e47` versus current `22cb0f3`.
3. Final research methodology and specific SDLC.
4. Conflicting flood color/class conventions.
5. No approved validated susceptibility formula/weights.
6. No approved operational susceptibility dataset.
7. Provisional DTM status and official replacement/integration.
8. Outdated survey files versus later decisions.
9. Preliminary reports cannot safely be pooled/generalized.
10. CSV workflow unimplemented.
11. Resident nearest-center function unimplemented.
12. Evacuation data/interface not fully established.
13. Admin completeness/current UI needs recheck.
14. Chapter 2 table-source verification and England-service naming.
15. Current DOCX visual QA.
16. Full current test rerun.
17. Production deployment not established.

Require explicit permission before changing the Statement of the Problem, objectives, significance, scope/limitations, definitions, framework/architecture/conceptual relations, methodology/SDLC, diagrams, table numbers/captions unless demonstrably wrong, flood taxonomy/colors, formulas/weights, baseline conflicts, or anything outside the named editing boundary.

When asked for text only, do not edit files. When asked only to analyze, do not implement.

---

## 18. Alignment tracker and continuation order

Update `FloodSense_Development_Manuscript_Alignment_Tracker.md` only after full re-audit.

Required columns:

| Feature | Development status | Code evidence | Document evidence | Commit | Manuscript section | Documentation status | Required action | Permission |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |

Also record audit date, current/previous commits, changed files, working-tree state, tests, unavailable tools, approved/provisional data, missing evidence, contradictions, and permissions.

Recommended order:

1. Verify transferred Git/files and inventory evidence.
2. Recover/read baseline from history if needed.
3. Render current manuscript for baseline visual review.
4. Resolve methodology/SDLC, color taxonomy, formula, and data-governance decisions.
5. Synchronize questionnaire files before collection.
6. Prepare environment, inspect UI, and rerun backend/mobile tests.
7. Reclassify features and update tracker.
8. Apply only authorized manuscript edits.
9. Verify references and render/compare with SURECUT page by page.
10. Report preserved content, supported additions, confirmed-but-unimplemented requirements, placeholders, exclusions, and unresolved matters.

---

## 19. Anti-hallucination checklist

Before any claim, ask:

- Is it integrated code or only a plan/model/mockup/prompt?
- Can the intended user access it?
- Is required data present, approved, and connected?
- Are migration, logic, validation, API, UI, permission, and tests present?
- Did tests run now, or is evidence historical?
- Was the full authoritative source inspected?
- Is it a consultation opinion, user decision, or institutional approval?
- Is a boundary being mistaken for hazard data?
- Is a questionnaire being mistaken for results?
- Is a script/demo being mistaken for approval/deployment?
- Is rainfall rate being mistaken for water level?
- Is elevation coloring being mistaken for susceptibility?
- Is another study's feature being transferred into FloodSense?
- Does the tense distinguish implemented, ongoing, planned, blocked, and unimplemented?

If evidence is insufficient, label it unverified, provisional, or unresolved.

---

## 20. Compact progress report

Completed/substantially completed:

- Full code audit at `fc62e47`.
- Alignment tracker created, though stale.
- Chapters 1-3 main manuscript committed.
- Chapter 2 expanded/aligned with SURECUT and comparative tables developed.
- Informal revision/handoff/consultation narration cleaned from formal thesis text.
- Deterministic scenario scope preserved.
- CSV requirement documented as unimplemented.
- TA and BDRRMO records collected.
- Provisional DTM, point-cloud exports, visuals, and QGIS project created.
- Logo concepts and A4 PDF created.

Partial/reconciliation needed:

- Current manuscript visual QA.
- Current test runs.
- Final methodology/SDLC.
- Primary-source verification for all Chapter 2 table cells.
- Tracker update.
- Survey synchronization.
- Flood level/color and computation validation.
- DTM provenance/official replacement/integration decision.

Not complete:

- Fixed-schema CSV export/import.
- Resident nearest-center discovery.
- Official susceptibility integration.
- Validated operational formula/weights.
- Final validated survey/evaluation results.
- Production deployment.

---

## 21. Suggested first response on the new computer

> I have the FloodSense continuity handout. I will first verify the transferred project's path, Git status, current commit, principal manuscript, alignment tracker, development repository, questionnaire, and DTM artifacts. I will treat the October 9, 2026 checkpoint as historical context, not as a substitute for the current files, and I will not modify anything until the requested task and editing boundary are clear.

This handout was compiled from all locally accessible project chats and a live repository inspection. Re-verify transferred files before consequential work.
