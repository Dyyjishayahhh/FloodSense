# FloodSense Thesis Revision Guide

## Latest TA Consultation Changes and Required Alterations

**Consultation date:** October 7, 2026  
**Source transcript:** `SirDhan_Oct7.docx`  
**Coverage:** Required revisions for Chapters 1 to 3 and the appendices

## Executive Summary

The TA wants the manuscript reorganized so that the research problem is established before FloodSense is presented as the solution. Every major problem must be supported by evidence, converted into an aligned objective, investigated through an appropriate method, and answered either by a research finding or a validated system function.

The latest consultation concentrates on five major corrections:

1. Rewrite Chapter 1 so that it presents the context and evidence-based problems before discussing the proposed system.
2. Make the Statement of the Problem, objectives, framework, architecture, scope, and limitations directly consistent with one another.
3. Rework the related studies in Chapter 2 so that every study contains identifiable Introduction, Methodology, Results, and Discussion information.
4. Expand Chapter 3 beyond software development by including research materials, a specific SDLC method, feasibility analysis, requirements gathering, testing, and evaluation.
5. Move detailed supporting diagrams and feasibility records into the appendices and reference them from the main chapters.

## Governing Revision Principle

The manuscript should follow this evidence chain:

> Local context and evidence -> main problem -> specific problems -> objectives -> research methods -> findings and system requirements -> validated system functions

FloodSense should not be introduced as the answer before the manuscript has demonstrated what problems actually exist and why those problems require a technological response.

# Chapter 1 Required Changes

## 1 Introduction

### TA's concern

The current opening reportedly presents the problem and proposed solution too quickly. It reads like a proposal that is already selling FloodSense rather than a completed or developing research study that first establishes its context.

### Required alteration

- Begin with an engaging description of the broader flood and preparedness context.
- Give the reader a clear reason to continue reading.
- Establish the importance of the subject before discussing the detailed research problem.
- Avoid opening with wording such as `This study aims to develop...`.
- Do not immediately enumerate the application, modules, technologies, or intended solution.
- Let the Introduction lead naturally into the Statement of the Problem.

### Expected result

The opening should introduce the topic and importance of the study. The detailed problem analysis should then be developed in the opening narrative of the Statement of the Problem.

## 2 Statement of the Problem

### Role of the SOP in the FloodSense manuscript

The FloodSense manuscript does not have a separate Background of the Study section. The TA's comparison means that the opening paragraphs of the SOP should perform the background function: they should explain the context, evidence, research gap, and main problem before presenting the numbered specific questions.

### TA's concern

The current SOP reportedly begins with solution-oriented wording such as `This study aims to develop...`. The TA wants the SOP to establish the background and problems without prematurely presenting FloodSense as the solution.

### Required alteration

- Remove the proposed application, technologies, and development language from the beginning of the SOP.
- Begin with a background-style problem narrative covering the flood context, local conditions, existing practices, and evidence from residents and stakeholders.
- Explain the gap between current conditions and the needs revealed by the evidence.
- Lead from that narrative into one evidence-based main problem.
- Derive the specific problems from that main problem.
- Use actual survey and interview findings to demonstrate that the problems exist.
- Write each specific problem as a researchable question.
- Ensure that every specific problem will be answered by either a research finding or a major validated system function.

### Recommended structure

1. Background-style narrative describing the flood context and local situation.
2. Survey, interview, BDRRMO, observation, and literature evidence establishing the gap.
3. Main problem derived from that evidence.
4. Specific problem concerning the community's flood experience or information difficulty.
5. Specific problem concerning preparedness or decision-making needs.
6. Specific problem concerning access, visualization, evacuation, or another validated concern.
7. Specific problem concerning system development, testing, or evaluation when required by the school's format.

### Important rule

Survey results may be used in the SOP. The TA described this as beneficial because the results provide direct evidence that the problem exists.

## 3 Objectives of the Study

### Alignment requirement

Each objective must correspond directly to one SOP question:

- SOP Question 1 -> Objective 1
- SOP Question 2 -> Objective 2
- SOP Question 3 -> Objective 3
- SOP Question 4 -> Objective 4

The objectives must not introduce a problem, variable, population, or feature that is absent from the SOP.

### Development objective

The TA wants the development objective to identify:

- The application being developed.
- Android as the intended platform if that remains the final scope.
- Flutter if it is the actual development framework.
- The specific development methodology used.
- The major functionalities of the system.

If several functionalities are named, present them as bullet points rather than one long sentence.

### Testing objective

- Keep the testing objective concise.
- State the testing approach without placing the full procedure inside the objective.
- Move the complete testing process to Chapter 3.

### Evaluation objective

- Keep the evaluation objective concise.
- If ISO/IEC 25010 is used, state it clearly but discuss the characteristics, instrument, participants, and computation in Chapter 3.

## 4 Theoretical or Conceptual Framework

### TA's concern

The current framework reportedly contains too many modules and names several modules according to their implementation technologies rather than their conceptual functions.

### Required alteration

- Name modules according to what they process or accomplish.
- Consolidate small functions under a limited number of meaningful main modules.
- Explain the inputs, processing, and outputs of each main module.
- Explain which research problem and user need each module addresses.
- Remove the database as a major conceptual module because a database is an expected implementation component.

### Examples of names to avoid

- Django REST Processing Module
- GeoDjango Processing Module
- PostgreSQL and PostGIS Database Module

### Better naming direction

- Flood Risk Assessment Module
- Map and Area Checking Module
- Preparedness and Evacuation Guidance Module
- Information Management Module

These names are examples only. The final names must match the validated and implemented system.

## 5 System Architecture

### Required alteration

- Position the administrator on the right side, following the TA's requested layout.
- Show only actual system users or actors around the system.
- Consolidate the internal modules.
- Make the direction of interaction between users and the system understandable.
- Remove reference descriptions that are incorrectly shown as users.

### Items to remove or relocate

- PAGASA-approved rainfall references when presented as an actor.
- Flood, spatial, preparedness, and evacuation-center reference descriptions when presented as users.
- Expert or reviewer labels when those persons do not directly operate the system.

Reference data and validation sources should instead be explained in the framework, data-source discussion, methodology, or technical implementation section.

## 6 Time and Place of the Study

### TA's concern

The section is too long and includes research procedures that belong in Chapter 3.

### Required alteration

Keep the section short and specific. Include only:

- The period during which the research was conducted.
- The place where the study and system development occurred.
- The locations of the surveys and interviews.
- The geographic coverage of the study.

### Move to Chapter 3

- How the survey was conducted.
- How interviews were performed.
- How the application was developed.
- How testing and evaluation were completed.

## 7 Scope and Limitations

### TA's concern

The scope and limitations are mixed together, making it difficult to determine where one ends and the other begins.

### Required alteration

The two may remain under one heading, but they must be separated into distinct paragraphs.

### Scope paragraph

State:

- What FloodSense covers.
- Who can use it.
- Which locations are included.
- Which platforms are included.
- Which main functions are included.
- Which data and processes are covered.

### Limitations paragraph

Begin with a clear transition such as `However, the study is limited to...` and state:

- Geographic limitations.
- Data limitations.
- Connectivity or device limitations.
- Missing official information.
- Platform restrictions.
- Features not included in the current version.
- Dependence on BDRRMO or other expert validation.

# Chapter 2 Required Changes

## 8 Related Local and Foreign Studies

### TA's main requirement

Every related study should contain identifiable IMRaD information:

- **Introduction:** the problem, purpose, and context of the study.
- **Methodology:** how the research was conducted and how the system was developed.
- **Results:** the findings, performance, evaluation, or effectiveness of the study.
- **Discussion:** what the results mean and how they relate to the study's objectives.

### Required alteration

- Replace summaries that only describe a related application's features.
- Select studies that report an actual methodology and results.
- Identify the research methodology used by each study.
- Identify the SDLC model separately when the source reports it.
- Explain whether the related system or intervention was effective.
- Explain the relevance, difference, and research gap in relation to FloodSense.

### Important distinction

Research methodology and SDLC methodology are not the same:

- Research methodology explains how the investigation was conducted.
- SDLC methodology explains how the software was developed.

### Interpretation of the IMRaD instruction

The most defensible interpretation is that each related-study discussion should extract the IMRaD elements. It does not necessarily require the entire Chapter 2 to be reorganized under four headings called Introduction, Methodology, Results, and Discussion unless the school's template specifically requires that structure.

# Chapter 3 Required Changes

## 9 Materials

### TA's concern

The current Materials section concentrates mainly on software-development technologies and hardware.

### Required alteration

Describe materials used across the entire research process, including:

- Requirements gathering.
- Surveys and interviews.
- Communication and coordination.
- Task management.
- Prototyping and design.
- Application development.
- Testing and evaluation.
- Data analysis.

### Possible materials

- Google Forms.
- Interview guides and recording or documentation tools.
- The communication platform actually used by the group.
- The task board or project-management tool actually used.
- Flutter and the actual development tools.
- Database and mapping technologies.
- Testing and evaluation instruments.

Do not claim the use of Slack, Trello, Kanban, or another tool unless the group actually used it. Detailed hardware specifications are unnecessary unless they directly affect the study.

## 10 Research Method and SDLC Methodology

### TA's concern

`Iterative and Incremental` remains too general, and the current methodology diagram does not sufficiently represent the team's actual process.

### Required alteration

- Select a specific Agile method, such as Scrum or Kanban, based on the team's real workflow.
- Explain why the selected method accurately represents the project.
- Replace the current diagram with a diagram of the selected SDLC method.
- Explain how FloodSense followed every relevant phase.
- Connect requirements gathering, analysis, design, development, testing, evaluation, and revision to the chosen method.

### Selection evidence

The method should be supported by the team's actual:

- Task-management tools.
- Sprint or workflow structure.
- Backlog or task-board records.
- Meetings and reviews.
- Iterations and prototype versions.
- Feedback and revision process.
- Version-control practices.

The team should not claim Scrum or Kanban merely because it was recommended by a generated response or sounds more specific.

## 11 Feasibility Study

### Status

This is one of the largest new manuscript requirements introduced in the latest consultation.

The feasibility study should appear before the detailed requirements-gathering discussion.

### 11.1 Schedule Feasibility

Explain whether the project can be completed within the available period.

Include:

- Major research and development activities.
- Planned and completed periods.
- Milestones.
- Evidence that the project is achievable within the research timeline.
- A Gantt chart in the appendices.

### 11.2 Cost Feasibility

Explain whether the project can be completed within an achievable budget.

Include:

- Development costs.
- Research and data-gathering expenses.
- Hosting, connectivity, transportation, printing, or other relevant costs.
- Resources already owned or available for free.
- A line-item budget table in the appendices.

### 11.3 Operational Feasibility

Explain whether FloodSense can fit into the intended stakeholder's actual operations.

Include:

- Existing BDRRMO or client operations.
- Current operational problems.
- How FloodSense may improve the process.
- Stakeholder acceptance or willingness to use the system.
- Required changes, training, or responsibilities.

### 11.4 Technical Feasibility

Explain whether the group can build and maintain the intended system.

Include:

- Required technologies.
- Available devices, tools, and services.
- Team knowledge and skills.
- Technologies that still need to be learned.
- Known technical risks.
- Evidence from the existing prototype that the project is technically achievable.

### Presentation rule

Summarize the feasibility conclusions in Chapter 3. Place the detailed Gantt chart, budget table, and other supporting records in the appendices and cite them from the main discussion.

## 12 Requirements Gathering

### Required alteration

Present requirements gathering as a complete research process, not merely a statement that the system has an administrator and a user.

Explain:

- Who supplied the requirements.
- How participants were selected.
- Which survey and interview instruments were used.
- How information was collected.
- How responses were analyzed.
- How requirements were derived from the results.
- How resident and administrator requirements differed.
- How BDRRMO knowledge was incorporated and validated.
- How later consultations changed the initial requirements.

## 13 Required Analysis Diagrams

The TA requested or recommended the following diagrams:

### Use Case Diagram

Show:

- System actors.
- Functions available to every actor.
- Boundaries between resident, administrator, BDRRMO, and other roles.

### Data Flow Diagram

Show:

- Sources of data.
- Major system processes.
- Data stores.
- Movement of data among users, processes, and storage.
- Outputs returned to users.

### HIPO Diagram

Show:

- Hierarchy of major functions.
- Subfunctions contained in every major module.
- Inputs, processes, and outputs.
- Logical organization of the system's functionality.

### Placement

Place the full diagrams in the appendices. Discuss and cross-reference them in the relevant Chapter 3 sections.

# Appendix Requirements

The revised manuscript should include, when applicable:

- Gantt chart for schedule feasibility.
- Line-item budget for cost feasibility.
- Use Case Diagram.
- Data Flow Diagram.
- HIPO Diagram.
- Survey questionnaire.
- Interview guide.
- Complete interview reports or approved summaries.
- Consent and data-privacy documents.
- Survey results or supporting tables.
- Expert-validation records.
- Other documentation supporting the research procedures.

Use explicit cross-references in the manuscript, such as:

> The detailed project schedule is presented in Appendix A.

> The cost-feasibility table is provided in Appendix B.

> The Use Case, Data Flow, and HIPO diagrams are presented in Appendices C to E.

The final appendix letters must follow the actual order of the completed manuscript.

# Latest Changes Introduced or Made More Explicit

The October 7 consultation introduced or substantially clarified the following requirements:

1. Rewrite the Introduction as a reader-oriented opening rather than an immediate problem-and-solution statement.
2. Remove `This study aims to develop...` and other solution language from the SOP.
3. Use survey results as direct support for the research problems.
4. Ensure that specific problems are answered by research findings or major system functions.
5. Align every objective with an SOP question.
6. Present multiple functionalities as bullet points under the development objective.
7. Shorten the testing and evaluation objectives and move their details to Chapter 3.
8. Rename modules according to purpose rather than technologies.
9. Consolidate the framework into fewer major modules.
10. Remove the database as a major conceptual module.
11. Revise the position and contents of the system-architecture actors.
12. Shorten the Time and Place of the Study.
13. Separate Scope and Limitations into clearly identifiable paragraphs.
14. Require IMRaD information in every useful related study.
15. Require research methodology and results in the related-study summaries.
16. Expand Materials to cover requirements gathering and the full research process.
17. Replace the current SDLC diagram.
18. Select a specific Agile method based on the team's real workflow.
19. Add schedule, cost, operational, and technical feasibility analyses.
20. Prepare a Gantt chart and line-item budget.
21. Add Use Case, Data Flow, and HIPO diagrams.
22. Move detailed diagrams and feasibility records into the appendices.
23. Cross-reference all relevant appendix materials from the main chapters.

# Earlier Guidance Reinforced by This Consultation

The latest consultation reinforces the following earlier directions:

- `Agile` and `Iterative and Incremental` are too general without a specific working method.
- The selected methodology must match the tools and process actually used by the team.
- Requirements gathering is part of the research, not only a software-development activity.
- Problems and features must be supported by evidence.
- The manuscript must agree with the system that was actually implemented.
- Requirements should be derived progressively from surveys, interviews, stakeholder input, and later validation.

The October 7 transcript does not cancel the earlier directions concerning the hybrid Expert System and Decision Support System, BDRRMO validation, scenario-based recommendations, optional NLP, technology access, or progressive requirements gathering. Those instructions remain relevant unless a later consultation explicitly changes them.

# Matters Still Requiring a Group Decision

The TA did not finalize the following matters in this consultation:

- Whether Scrum or Kanban will be used.
- Whether Android remains the sole user platform.
- The final names and number of main modules.
- The final actors shown in the system architecture.
- The final set of validated system functionalities.
- The exact structure required by the school's manuscript template.
- The final appendix numbering.
- The exact testing methods.
- The ISO/IEC 25010 characteristics to be evaluated.
- The final participant and sampling methodology.

These must be confirmed using the group's actual records, school requirements, and later adviser instructions.

# Recommended Revision Sequence

Revise the manuscript in this order:

1. Consolidate all survey, interview, and stakeholder evidence.
2. Define the evidence-based main problem and specific problems.
3. Rewrite the Introduction and the background-style opening narrative of the SOP.
4. Rewrite the Statement of the Problem without solution language.
5. Align the objectives with the SOP.
6. Confirm the final system scope and limitations.
7. Simplify the theoretical or conceptual framework.
8. Revise the system architecture.
9. Shorten the Time and Place section.
10. Rework the related studies using IMRaD information.
11. Select the specific Agile methodology using actual workflow evidence.
12. Replace and explain the SDLC diagram.
13. Add the four feasibility analyses.
14. Expand the Materials and requirements-gathering discussions.
15. Create the Use Case, DFD, HIPO, Gantt chart, and budget table.
16. Arrange and cross-reference the appendices.
17. Conduct a final alignment audit across Chapters 1 to 3.

# Working Revision Checklist

## Chapter 1

- [ ] Rewrite the Introduction to hook the reader.
- [ ] Establish the local context before introducing FloodSense.
- [ ] Support the SOP's background-style problem narrative with survey, interview, stakeholder, and literature evidence.
- [ ] Remove solution language from the SOP.
- [ ] Define one main problem and evidence-based specific problems.
- [ ] Align every objective with an SOP question.
- [ ] Format the major functionalities as bullet points under the development objective.
- [ ] Shorten the testing and evaluation objectives.
- [ ] Rename framework modules according to their functions.
- [ ] Consolidate the number of main modules.
- [ ] Remove the database as a conceptual module.
- [ ] Revise the system architecture actors and layout.
- [ ] Shorten Time and Place of the Study.
- [ ] Separate the Scope and Limitations paragraphs.

## Chapter 2

- [ ] Retain only sources relevant to the finalized problems and system functions.
- [ ] Add the methodology used in every related study.
- [ ] Add the reported results and discussion of every related study.
- [ ] Distinguish research methodology from SDLC methodology.
- [ ] Explain each study's relevance, difference, and research gap.
- [ ] Complete the synthesis of the literature and studies.

## Chapter 3

- [ ] Expand Materials beyond development hardware and software.
- [ ] Document the tools used for surveys, interviews, communication, and task management.
- [ ] Select and justify the specific Agile method.
- [ ] Replace the current SDLC diagram.
- [ ] Explain the actual development process phase by phase.
- [ ] Add schedule feasibility.
- [ ] Add cost feasibility.
- [ ] Add operational feasibility.
- [ ] Add technical feasibility.
- [ ] Expand the requirements-gathering procedure.
- [ ] Explain how gathered data became system requirements.
- [ ] Document testing and evaluation procedures.

## Appendices

- [ ] Add the Gantt chart.
- [ ] Add the line-item budget table.
- [ ] Add the Use Case Diagram.
- [ ] Add the Data Flow Diagram.
- [ ] Add the HIPO Diagram.
- [ ] Add the questionnaire and interview guide.
- [ ] Add consent and data-privacy documents.
- [ ] Add relevant survey and expert-validation records.
- [ ] Insert correct appendix cross-references in Chapters 1 to 3.

# Final Alignment Audit

Before submission, confirm that:

- Every problem is supported by evidence.
- Every SOP question has a corresponding objective.
- Every objective has a corresponding research method.
- Every claimed feature answers a validated problem or requirement.
- Every module shown in the framework exists in the architecture and actual system.
- Every actor shown in the architecture has a real system role.
- The selected SDLC method matches the team's actual workflow.
- The Materials section lists only tools actually used.
- Feasibility conclusions are supported by schedules, costs, stakeholder information, and technical evidence.
- Chapter 2 sources directly support the study's problems, methods, and system functions.
- All detailed supporting records are placed in and cited from the appendices.

