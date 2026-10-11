# Survey data dictionary and analysis plan

Source instrument: `revised-questionnaire-the-approved-survey.docx`. This is the active version and was not modified. Close the survey and archive a dated export before cleaning or analysis.

## Eligibility and record handling

- Include only respondents who consented, answered Q1 as 18 or older, and answered Q2 as living in an included Bacoor community.
- Q2 “I am not sure” requires a documented inclusion decision; do not silently count it as Yes.
- Retain “Prefer not to answer,” “I am not sure,” “I cannot recall,” “not applicable,” and skipped items as distinct codes where analytically useful; do not convert them to favorable or unfavorable responses.
- Report the denominator for every item. Multiple-response items may exceed 100% when percentages are based on valid respondents selecting each option.
- Remove names from the analysis dataset and keep any linkage file separately with restricted access.

| Item | Variable | Response/coding | Valid/missing treatment | Planned analysis/output | SOP/requirement link |
|---|---|---|---|---|---|
| Consent | consent | 1 agree; 0 decline | Decline = terminate/exclude | Flow count only | Ethics |
| D1 | name_raw | Text; direct identifier | Never tabulate publicly; blank allowed | Remove from analysis copy | Privacy |
| D2 | age | Integer; special code for prefer not to answer | Range-check; report missing | Median/range or approved age groups after collection | Participant description |
| D3 | sex_gender | 1 male; 2 female; 3 self-describe; 9 prefer not | Self-description protected; suppress small cells if needed | Frequency/percentage | Participant description |
| D4 | residence_raw; barangay_std | Free text normalized to approved barangay list; retain original separately | Invalid/non-Bacoor flagged; ambiguous entries reviewed | Frequency/map only when confidentiality permits; `[Insert final distribution by barangay.]` | Local coverage/eligibility |
| D5 | residence_length | 1 <1 year; 2 1–5; 3 6–10; 4 >10; 9 prefer not | 9 separate from missing | Frequency/percentage | Context |
| Q1 | age_eligible | 1 yes; 0 no; 9 prefer not | 0/9 terminate/exclude | Eligibility flow | Sampling |
| Q2 | bacoor_eligible | 1 yes; 0 no; 8 unsure | 0 exclude; 8 pending documented rule | Eligibility flow | Sampling |
| Q3 | flood_effect_frequency | 5 very often; 4 often; 3 sometimes; 2 rarely; 1 never; 8 cannot recall | 8 separate/not in ordered summaries | Frequency/percentage; optional median, no unsupported mean | Problem context |
| Q4 | first_information_source | Categories as printed (official social media; barangay PA; SMS; TV/radio; government/weather site; family/community; personal observation; none) | Single response; blank missing | Bar chart/table | Current practices |
| Q5 | obtain_needed_information | 5 always to 1 never; 8 usually no need | 8 separate | Frequency/percentage; collapse only with predeclared rule | Access problem |
| Q6 | information_timing | Early enough; shortly before action; after action started; after affected; none; cannot recall | Keep all categories | Frequency/percentage | Timeliness problem |
| Q7 | understandability | Very easy; easy; neutral; difficult; very difficult; receives none | “Receives none” separate | Frequency/percentage; favorable/neutral/unfavorable only if declared before analysis | Clarity problem |
| Q8 | location_specificity | All details; some; too general; not important; receives none; unsure | Keep separate | Frequency/percentage | Location-specific information |
| Q9 | conflicting_information | Very often; often; sometimes; rarely; never; one source; unsure | One-source/unsure separate | Frequency/percentage | Consistency/provenance |
| Q10 | decision_effect | Helps in time; helps but late; available/not useful; cannot obtain; no need; unsure | Keep separate | Frequency/percentage | Decision support need |
| Q11 | information_problems | One binary variable per option | Multi-select; check contradictions with “no problem” | Count and % selecting each; ranked bar chart | Problem evidence |
| Q12 | useful_lead_time | One binary variable per option | Multi-select; “depends/no need” separate | Count and % selecting each | Preparedness timing; not evidence of live forecast capability |
| Q13 | helpful_information_types | One binary variable per option; maximum 2 selections | Flag >2 selections for review; do not auto-delete without rule | Count and % selecting each; feature-retention evidence with interview/literature triangulation | Requirements |
| Q14 | needs_met | 5 completely to 1 not at all; 8 no need | 8 separate | Frequency/percentage; optional median | Main information gap |
| Q15 | mobile_access | Own smartphone; shared smartphone; basic phone; no regular phone; prefer not | Keep separate | Frequency/percentage; `[Insert final device-ownership result.]` | Android operational feasibility |
| Q16 | helpful_delivery_methods | One binary variable per option | Multi-select; contradictory selections reviewed | Count and % selecting each; `[Insert final connectivity result.]` only if connectivity was separately measured—this item measures method preference, not internet quality | Delivery feasibility |

## Open-text and validation procedures

The active instrument contains direct identifiers and free-text residence/self-description fields, but no general open-comment question. Two researchers should independently normalize barangay text using an approved codebook, compare disagreements, retain an audit column, and resolve uncertain locations without guessing. Any later open comments should be coded through an explicit codebook, double-coding sample, discrepancy resolution, and frequency reporting with de-identified excerpts only when consent permits.

## Statistical treatment

Use counts and valid percentages for nominal and multiple-response items. Report ordered response distributions and medians where useful. Do not use a weighted mean merely because choices are numbered; it is defensible only if the team predefines a genuinely ordered equal-interval interpretation and the adviser approves it. No inferential test is currently required by the SOP. If subgroup comparisons are later proposed, document hypotheses, sample adequacy, cell-size protection, and multiplicity before running them.
