from pathlib import Path

from docx import Document


DOCX = Path(r"E:\School Downloads-Jiraah\00 FILES IN SCHOOL\3RD YEAR 2ND SEM\DCIT 60\FLOODSENSE\1.FLOODSENSE-200A\FloodSense\1. FLOODSENSE-MAIN\WORKING_REVISION_2026-10-11\FloodSense_Chapters_1_to_3_Survey_Ongoing_Working_Revision_2026-10-11.docx")

replacements = {
    "Flood Susceptibility. The relative potential of a supported geographic area to experience rainfall-related flooding under a user-confirmed scenario. The current prototype demonstrates Low, Moderate, High, or Very High outputs but does not yet contain an approved operational Bacoor City susceptibility dataset and is not a real-time prediction.":
        "Flood Susceptibility. The relative potential of a supported geographic area to experience rainfall-related flooding under a user-confirmed scenario. FloodSense expresses the assessment as Low, Moderate, High, or Very High and identifies it as a research-system output rather than an official Bacoor City classification or real-time prediction.",
    "Research materials included the bilingual resident questionnaire and consent notice, Google Forms and the exported response dataset, interview guides and recorded institutional interviews, request letters and approved school documents, scholarly and official references, the development-manuscript alignment records, requirements documents, diagrams, interface designs, test cases, and verification records.":
        "Research materials included the bilingual resident questionnaire and consent notice, Google Forms and the exported response dataset, interview guides and recorded institutional interviews, request letters and approved school documents, scholarly and official references, requirements and traceability records, diagrams, interface designs, test cases, and verification records.",
    "FloodSense followed an Agile Development Method. This method was selected because requirements, interfaces, backend services, geographic functions, and tests were refined through repeated cycles of planning, design, implementation, verification, review, and correction. The study does not claim Scrum or Kanban ceremonies that are not established by the project records. Agile activities were organized into the following recurring phases:":
        "FloodSense followed an Agile Development Method because requirements, interfaces, backend services, geographic functions, and tests were refined through repeated cycles of planning, design, implementation, verification, review, and correction. The available workflow evidence supports iterative Agile development but does not document the formal roles and ceremonies required to identify the process as Scrum or Kanban. The following phases recurred throughout development:",
}

doc = Document(DOCX)
found = set()
for p in doc.paragraphs:
    if p.text in replacements:
        p.text = replacements[p.text]
        found.add(p.text)

if len(found) != len(replacements):
    raise RuntimeError(f"Expected {len(replacements)} replacements; completed {len(found)}")

doc.save(DOCX)
print(DOCX)
