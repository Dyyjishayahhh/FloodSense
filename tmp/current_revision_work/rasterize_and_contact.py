from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageDraw, ImageFont


root = Path(r"E:\School Downloads-Jiraah\00 FILES IN SCHOOL\3RD YEAR 2ND SEM\DCIT 60\FLOODSENSE\1.FLOODSENSE-200A\FloodSense")
pdf_path = root / r"tmp\current_revision_work\render_final_word\FloodSense_Chapters_1_to_3_Survey_Ongoing_Working_Revision_2026-10-11.pdf"
pages_dir = root / r"tmp\current_revision_work\render_final_word\pages"
contact_dir = root / r"tmp\current_revision_work\render_final_word\contacts"
pages_dir.mkdir(parents=True, exist_ok=True)
contact_dir.mkdir(parents=True, exist_ok=True)

pdf = pdfium.PdfDocument(pdf_path)
page_paths = []
for index in range(len(pdf)):
    image = pdf[index].render(scale=1.45).to_pil().convert("RGB")
    path = pages_dir / f"page-{index + 1:03d}.png"
    image.save(path)
    page_paths.append(path)

font = ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf", 26)
for start in range(0, len(page_paths), 8):
    group = page_paths[start : start + 8]
    sheet = Image.new("RGB", (2000, 2840), "#dddddd")
    draw = ImageDraw.Draw(sheet)
    for i, path in enumerate(group):
        page = Image.open(path).convert("RGB")
        page.thumbnail((475, 1320))
        x = 15 + (i % 4) * 495
        y = 55 + (i // 4) * 1390
        sheet.paste(page, (x, y))
        draw.text((x, 15 + (i // 4) * 1390), f"Page {start + i + 1}", font=font, fill="black")
    sheet.save(contact_dir / f"contact_{start + 1:03d}_{start + len(group):03d}.png")

print(f"pages={len(page_paths)} contacts={len(list(contact_dir.glob('contact_*.png')))}")
