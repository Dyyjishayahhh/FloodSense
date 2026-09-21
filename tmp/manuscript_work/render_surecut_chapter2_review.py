from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageDraw


pdf_path = Path(r"1. FLOODSENSE-MAIN/GUIDE FILES/SURECUT-MANUSCRIPT.pdf")
out_dir = Path(r"tmp/manuscript_work/surecut_chapter2_review")
out_dir.mkdir(parents=True, exist_ok=True)

doc = pdfium.PdfDocument(pdf_path)
pages = []
for pdf_page in range(36, 63):
    output = out_dir / f"surecut-ch2-pdf-{pdf_page}.png"
    doc[pdf_page - 1].render(scale=1.5).to_pil().save(output)
    pages.append(output)

cell_w, cell_h, label_h = 700, 990, 34
for sheet_start in range(0, len(pages), 4):
    group = pages[sheet_start:sheet_start + 4]
    canvas = Image.new("RGB", (cell_w * 2 + 30, (cell_h + label_h) * 2 + 30), "#777777")
    draw = ImageDraw.Draw(canvas)
    for offset, page_path in enumerate(group):
        page = Image.open(page_path).convert("RGB")
        page.thumbnail((cell_w, cell_h))
        x = 10 + (offset % 2) * (cell_w + 10)
        y = 10 + (offset // 2) * (cell_h + label_h + 10)
        draw.rectangle([x, y, x + cell_w, y + label_h - 2], fill="white")
        draw.text((x + 8, y + 7), page_path.stem, fill="black")
        canvas.paste(page, (x, y + label_h))
    canvas.save(out_dir / f"contact-{sheet_start // 4 + 1}.png")

print(f"pages={len(pages)} contacts={(len(pages)+3)//4}")
