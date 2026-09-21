from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageDraw


folder = Path(r"tmp/manuscript_work/rendered_development_aligned")
pdf_path = folder / "FloodSense_Chapters_1_to_3_Development_Aligned_2026-09-20.pdf"
for pattern in ("page-*.png", "contact-*.png"):
    for old in folder.glob(pattern):
        old.unlink()
doc = pdfium.PdfDocument(pdf_path)
pages = []
for index in range(len(doc)):
    out = folder / f"page-{index + 1}.png"
    doc[index].render(scale=1.5).to_pil().save(out)
    pages.append(out)

cell_w, cell_h, label_h = 700, 990, 34
for sheet_index in range(0, len(pages), 4):
    group = pages[sheet_index:sheet_index + 4]
    canvas = Image.new("RGB", (cell_w * 2 + 30, (cell_h + label_h) * 2 + 30), "#777777")
    draw = ImageDraw.Draw(canvas)
    for offset, page_path in enumerate(group):
        image = Image.open(page_path).convert("RGB")
        image.thumbnail((cell_w, cell_h))
        x = 10 + (offset % 2) * (cell_w + 10)
        y = 10 + (offset // 2) * (cell_h + label_h + 10)
        draw.rectangle([x, y, x + cell_w, y + label_h - 2], fill="white")
        draw.text((x + 8, y + 7), page_path.stem, fill="black")
        canvas.paste(image, (x, y + label_h))
    canvas.save(folder / f"contact-{sheet_index // 4 + 1}.png")

print(f"pages={len(pages)} contact_sheets={(len(pages) + 3) // 4}")
