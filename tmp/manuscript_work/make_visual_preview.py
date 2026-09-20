from html import escape
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


DOCX = Path(r"1. FLOODSENSE-MAIN/MAIN FILESS/FloodSense_Manuscript_Chapters_1_to_3_Final.docx")
HTML = Path(r"tmp/manuscript_work/final_visual_preview.html")


def blocks(document):
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield Table(child, document)


chapter_names = {"INTRODUCTION", "REVIEW OF RELATED LITERATURE", "METHODOLOGY", "REFERENCES"}
section_names = {
    "Statement of the Problem", "Objectives of the Study", "Theoretical Framework",
    "System Architecture", "Significance of the Study", "Time and Place of the Study",
    "Scope and Limitation of the Study", "Definition of Terms", "Related Foreign Literature",
    "Related Local Literature", "Related Foreign Studies", "Related Local Studies", "Synthesis",
    "Materials", "Method", "Participants of the Study", "Statistical Treatment of Data",
    "Articles", "Journals", "Research Records",
}
sub_names = {
    "Flood Susceptibility and Spatial Factors", "Interactive Mapping and Risk Communication",
    "Decision Support and Explainability", "Flood Conditions and Institutional Responsibility",
    "Scenario-Based Hazard Information", "Local Requirements and Data Governance",
    "Requirements Analysis", "System Design", "CSV Parameter-Table Exchange",
    "Implementation and Verification",
}

parts = []
paragraph_index = 0
for block in blocks(Document(DOCX)):
    if isinstance(block, Paragraph):
        text = block.text.strip()
        if not text:
            continue
        cls = "body"
        if paragraph_index < 5:
            cls = f"cover cover-{paragraph_index + 1}"
        elif text.startswith("FloodSense: An Android-Based Expert System"):
            cls = "identity-title new-page"
        elif text.startswith("Reymart V. Goc-ong"):
            cls = "identity-authors"
        elif text.startswith("An undergraduate thesis manuscript"):
            cls = "identity-note"
        elif text in chapter_names:
            cls = "chapter" + (" new-page" if text != "INTRODUCTION" else "")
        elif text in section_names:
            cls = "section"
        elif text in sub_names:
            cls = "subsection"
        elif text.startswith("["):
            cls = "pending"
        elif text.startswith("Table 1."):
            cls = "table-title"
        elif text[:2].strip(". ").isdigit():
            cls = "numbered"
        elif text.startswith("FloodSense Research Team.") or "https://" in text or text.startswith(("Alabbad,", "Debnath,", "Galanti,", "Johnson,", "Kostopoulos,", "Lin,", "Nagumo,", "Nguyen,", "Oubennaceur,", "GeoRisk", "International Organization", "Japan International", "Philippine Statistics")):
            cls = "reference"
        parts.append(f'<p class="{cls}">{escape(text).replace(chr(10), "<br>")}</p>')
        paragraph_index += 1
    else:
        rows = []
        for row in block.rows:
            cells = "".join(f"<td>{escape(cell.text.strip())}</td>" for cell in row.cells)
            rows.append(f"<tr>{cells}</tr>")
        parts.append("<table>" + "".join(rows) + "</table>")

css = r"""
@page { size: A4; margin: 1in 1in 1in 1.5in; }
* { box-sizing: border-box; }
body { margin: 0; font-family: Arial, sans-serif; font-size: 11pt; color: #000; }
p { margin: 0; }
.new-page { break-before: page; page-break-before: always; }
.cover { text-align: center; line-height: 1.15; }
.cover-1 { font-weight: bold; margin-top: .1in; margin-bottom: 2.0in; }
.cover-2 { margin-bottom: 1.7in; }
.cover-3 { margin-bottom: 1.55in; }
.cover-4 { font-weight: bold; }
.cover-5 { break-after: page; page-break-after: always; }
.identity-title, .identity-authors { text-align: center; line-height: 1.15; font-weight: bold; margin-bottom: .38in; }
.identity-note { text-align: justify; line-height: 1.15; border-top: 1.5px solid #000; border-bottom: 1.5px solid #000; padding: 4px 0; margin-bottom: .3in; }
.chapter { text-align: center; font-weight: bold; line-height: 2; margin-bottom: .35in; }
.section { font-weight: bold; line-height: 2; margin-top: .35in; }
.subsection { font-weight: bold; line-height: 2; }
.body { text-align: justify; text-indent: .5in; line-height: 2; }
.numbered { text-align: justify; margin-left: .75in; text-indent: -.25in; line-height: 2; }
.pending { text-align: center; font-style: italic; line-height: 2; margin: .18in 0; }
.table-title { font-style: italic; line-height: 1.15; margin: .18in 0 .08in; }
.reference { line-height: 1.15; margin: 0 0 .18in .5in; text-indent: -.5in; }
table { width: 100%; border-collapse: collapse; font-size: 9pt; margin: 0; }
td { border: 1px solid #000; padding: 5px; vertical-align: middle; }
tr:first-child td { font-weight: bold; text-align: center; }
"""
HTML.write_text("<!doctype html><html><head><meta charset='utf-8'><style>" + css + "</style></head><body>" + "\n".join(parts) + "</body></html>", encoding="utf-8")
print(HTML)
