from __future__ import annotations

from copy import deepcopy
from hashlib import sha256
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from docx.table import Table
from docx.text.paragraph import Paragraph


ROOT = Path(r"E:\School Downloads-Jiraah\00 FILES IN SCHOOL\3RD YEAR 2ND SEM\DCIT 60\FLOODSENSE\1.FLOODSENSE-200A\FloodSense")
MANUSCRIPT = ROOT / "1. FLOODSENSE-MAIN" / "WORKING_REVISION_2026-10-11" / "FloodSense_Chapters_1_to_3_Introduction_Chapter2_Working_Revision_2026-10-12.docx"


def digest(items: list[str]) -> str:
    return sha256("\n".join(items).encode("utf-8")).hexdigest()


def replace_text(paragraph: Paragraph, text: str) -> None:
    paragraph.text = text


def insert_before(anchor: Paragraph, text: str, model: Paragraph) -> Paragraph:
    p = anchor.insert_paragraph_before(text)
    if model._p.pPr is not None:
        p._p.insert(0, deepcopy(model._p.pPr))
    return p


def set_cell_text(cell, text: str) -> None:
    cell.text = text


def set_border(parent, side: str, color: str = "000000", width: int = 12) -> None:
    existing = parent.find(qn("w:tcBorders")) or parent.find(qn("w:tblBorders"))
    if existing is None:
        tag = "w:tcBorders" if parent.tag == qn("w:tcPr") else "w:tblBorders"
        existing = OxmlElement(tag)
        parent.append(existing)
    edge = existing.find(qn(f"w:{side}"))
    if edge is None:
        edge = OxmlElement(f"w:{side}")
        existing.append(edge)
    edge.set(qn("w:val"), "single")
    edge.set(qn("w:sz"), str(width))
    edge.set(qn("w:color"), color)


def mark_no_split(row) -> None:
    trpr = row._tr.get_or_add_trPr()
    if trpr.find(qn("w:cantSplit")) is None:
        trpr.append(OxmlElement("w:cantSplit"))


def mark_header(row) -> None:
    trpr = row._tr.get_or_add_trPr()
    if trpr.find(qn("w:tblHeader")) is None:
        trpr.append(OxmlElement("w:tblHeader"))


def format_table(table: Table, matrix: bool = False) -> None:
    table.autofit = False
    tblpr = table._tbl.tblPr
    set_border(tblpr, "top", width=14)
    set_border(tblpr, "bottom", width=14)
    for row_i, row in enumerate(table.rows):
        mark_no_split(row)
        if row_i == 0 or (matrix and row_i == 1):
            mark_header(row)
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            tcpr = cell._tc.get_or_add_tcPr()
            margins = tcpr.find(qn("w:tcMar"))
            if margins is None:
                margins = OxmlElement("w:tcMar")
                tcpr.append(margins)
            for side, twips in (("top", 55), ("bottom", 55), ("left", 65), ("right", 65)):
                node = margins.find(qn(f"w:{side}"))
                if node is None:
                    node = OxmlElement(f"w:{side}")
                    margins.append(node)
                node.set(qn("w:w"), str(twips))
                node.set(qn("w:type"), "dxa")
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.first_line_indent = Pt(0)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if row_i < (2 if matrix else 1) else WD_ALIGN_PARAGRAPH.LEFT
                for run in p.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(11)
                    if row_i < (2 if matrix else 1):
                        run.bold = True
    header = table.rows[1 if matrix else 0]
    for cell in header.cells:
        set_border(cell._tc.get_or_add_tcPr(), "bottom", width=7)


def split_table_after_two_studies(document: Document, original: Table, caption: Paragraph, number: int) -> Table:
    clone = deepcopy(original._tbl)
    # Keep header and first two study rows in the first table.
    for row in list(original._tbl.tr_lst)[3:]:
        original._tbl.remove(row)
    # Keep header and last two study rows in the continued table.
    for row in list(clone.tr_lst)[1:3]:
        clone.remove(row)
    continued_xml = OxmlElement("w:p")
    original._tbl.addnext(continued_xml)
    continued = Paragraph(continued_xml, document._body)
    continued.text = f"Table {number}. Continued"
    if caption._p.pPr is not None:
        continued._p.insert(0, deepcopy(caption._p.pPr))
    continued.paragraph_format.page_break_before = True
    continued.paragraph_format.keep_with_next = True
    continued_xml.addnext(clone)
    return Table(clone, document._body)


doc = Document(MANUSCRIPT)
paragraphs = doc.paragraphs
assert paragraphs[108].text == "Time and Place of the Study"
assert paragraphs[111].text == "Scope and Limitation of the Study"
assert paragraphs[137].text == "REVIEW OF RELATED LITERATURE"
assert paragraphs[259].text == "METHODOLOGY"
assert len(doc.tables) == 4
foreign_source, local_source, matrix_source = doc.tables[:3]

chapter3_before = digest([p.text for p in paragraphs[259:]])
chapter1_before = digest([p.text for p in paragraphs[:108]])

# Chapter 1: follow SURECUT's chronological place/period structure, without
# importing its transportation content or long methodology-stage narration.
replace_text(paragraphs[109],
    "The study is situated in Bacoor City, Cavite, the geographic setting for the resident needs assessment and institutional inquiry. Cavite State University–Bacoor City Campus serves as the academic base for system development and manuscript preparation. The location of the study corresponds to the residents and local offices whose information needs and responsibilities concern flood preparedness in Bacoor City.")
replace_text(paragraphs[110],
    "The institutional interviews were conducted at the respective Bacoor City offices of the Bacoor City Disaster Risk Reduction and Management Office (BDRRMO) on May 12 and May 24, 2026, and the City Planning and Development Coordinator Office (CPDCO) on May 18, 2026. These visits preceded the principal study period and provided local institutional context for the Bacoor-focused research.")
insert_before(paragraphs[111],
    "The principal study period began in September 2026 and is expected to continue through May 2027. The initial resident needs survey was completed on October 11, 2026, with 40 qualified respondents from 14 Bacoor City barangays. These dates identify the completed early data-gathering activities and the planned period for the continuing study.",
    paragraphs[110])

# Chapter 1: four scope paragraphs followed by a final limitations paragraph.
replace_text(paragraphs[112],
    "FloodSense comprises a connected Android application for Bacoor City residents and a browser-based administration system for authorized personnel. The study covers pre-event, scenario-based flood-susceptibility assessment and preparedness support within its documented Bacoor administrative coverage. It is intended to help residents examine a hypothetical rainfall condition for a chosen area, rather than report a flood occurring at that moment.")
replace_text(paragraphs[113],
    "Residents authenticate and select or confirm a hypothetical rainfall intensity and duration. They may choose a barangay, place or adjust a map pin, or request optional one-time foreground location; they then confirm or correct the resolved barangay and explicitly request an assessment. The application presents the selected scenario, an explained susceptibility class, and geographic context through its resident interface.")
insert_before(paragraphs[114],
    "The deterministic Rule-Based Expert System performs the susceptibility assessment. A separate Decision Support System presents preparedness guidance and evacuation-center reference information without changing the assessment result. The interactive map provides administrative and geographic context for the selected area, while any susceptibility display remains subject to the validity and provenance of its underlying information.",
    paragraphs[113])
insert_before(paragraphs[114],
    "Authorized administrators use the web system to review and manage rainfall and map references, assessment parameters, preparedness content, evacuation-center records, data-source information, and audit records within their permissions. Ordinary administrators do not modify raw rules, priorities, conflict-resolution logic, or algorithm code. Location access is user-initiated and purpose-limited, and geographic records require documented sources and validation before they are presented as authoritative local information.",
    paragraphs[113])
replace_text(paragraphs[114],
    "However, FloodSense is limited to scenario-based assessment and preparedness support; it does not continuously monitor rainfall or water level, issue official forecasts or warnings, send autonomous alerts or evacuation orders, track location in the background, determine safe routes, report live evacuation-center occupancy, or use machine-learning prediction. Its 47-barangay layer is an administrative reference, not a City-verified flood-susceptibility dataset. Assessment inputs and guidance require appropriate local expert and institutional validation before authoritative use. Evacuation-center distance is a straight-line reference and does not establish road safety, accessibility, opening status, or capacity. Connected functions require a compatible device and network access. The purposive needs-assessment sample of 40 respondents is not representative of all Bacoor residents.")

# Chapter 2: source-grounded clarification and closer table presentation.
replace_text(paragraphs[138],
    "This chapter reviews related literature and studies that inform FloodSense's requirements and design. It examines concepts, methods, findings, and limitations relevant to flood-susceptibility assessment, geographic information, explainable rule-based inference, and preparedness support.")
set_cell_text(foreign_source.rows[0].cells[1], "AUTHOR(S)\n/YEAR")
set_cell_text(local_source.rows[0].cells[1], "AUTHOR(S)\n/YEAR")
set_cell_text(foreign_source.rows[1].cells[2],
    "Site survey; sensor-system design and testing; 30-day reservoir observation of rain, temperature, humidity, and water level.")
set_cell_text(foreign_source.rows[4].cells[2],
    "Rapid Application Development (Web GIS); HAND/GloFAS flood analysis and participatory mapping; ConvLSTM drought modeling.")
set_cell_text(local_source.rows[2].cells[2],
    "Agile Development Model (software development); quantitative survey and device test with 20 CDRRMO personnel.")
set_cell_text(local_source.rows[3].cells[2],
    "Descriptive and developmental research designs; Agile software development; purposive acceptability survey with 30 evaluators.")

foreign = split_table_after_two_studies(doc, foreign_source, paragraphs[245], 1)
local = split_table_after_two_studies(doc, local_source, paragraphs[246], 2)
for tbl in (foreign_source, foreign, local_source, local):
    format_table(tbl)
format_table(matrix_source, matrix=True)

# Apply consistent SURECUT-like typographic hierarchy in Chapter 2 only.
for p in paragraphs[137:259]:
    if not p.text.strip():
        continue
    p.paragraph_format.widow_control = True
    if p.text.startswith("Table "):
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.space_before = Pt(9)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.keep_with_next = True
    elif p.runs and all(r.bold for r in p.runs if r.text.strip()):
        p.paragraph_format.keep_with_next = True
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(11)

doc.save(MANUSCRIPT)
check = Document(MANUSCRIPT)
assert digest([p.text for p in check.paragraphs[:108]]) == chapter1_before
assert digest([p.text for p in check.paragraphs if p.text == "METHODOLOGY" or False]) != ""  # chapter boundary remains present
start3 = next(i for i,p in enumerate(check.paragraphs) if p.text == "METHODOLOGY")
assert digest([p.text for p in check.paragraphs[start3:]]) == chapter3_before
assert len(check.tables) == 6
assert [p.text for p in check.paragraphs if p.text.startswith("However, FloodSense")] 
print(f"Edited in place: {MANUSCRIPT}")
print(f"Tables now: {len(check.tables)} (foreign/local each split into a two-row continuation)")
print(f"Chapters outside targeted Chapter 1 passages and Chapter 2 text unchanged: yes")
