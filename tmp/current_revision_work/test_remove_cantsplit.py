from pathlib import Path
from docx import Document
from docx.oxml.ns import qn

root = Path(r"E:\School Downloads-Jiraah\00 FILES IN SCHOOL\3RD YEAR 2ND SEM\DCIT 60\FLOODSENSE\1.FLOODSENSE-200A\FloodSense")
path = root / r"1. FLOODSENSE-MAIN\WORKING_REVISION_2026-10-11\FloodSense_Chapters_1_to_3_Survey_Ongoing_Working_Revision_2026-10-11.docx"
doc = Document(path)
for table in doc.tables[:2]:
    for row in table.rows:
        trpr = row._tr.get_or_add_trPr()
        for child in list(trpr):
            if child.tag == qn("w:cantSplit"):
                trpr.remove(child)
doc.save(path)
print(path)
