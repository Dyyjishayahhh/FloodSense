from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageDraw


ROOT = Path(r"E:\School Downloads-Jiraah\00 FILES IN SCHOOL\3RD YEAR 2ND SEM\DCIT 60\FLOODSENSE\1.FLOODSENSE-200A\FloodSense")
folder = ROOT / "tmp" / "ch2_citation_revision_2026-09-21" / "word-render"

for name in ("baseline", "final"):
    pdf_path = folder / f"{name}.pdf"
    out_dir = folder / name
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob("*.png"):
        old.unlink()
    doc = pdfium.PdfDocument(pdf_path)
    pages = []
    for index in range(len(doc)):
        out = out_dir / f"page-{index + 1:03d}.png"
        doc[index].render(scale=1.8).to_pil().save(out)
        pages.append(out)

    cell_w, cell_h, label_h = 620, 875, 30
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
            draw.text((x + 8, y + 6), page_path.stem, fill="black")
            canvas.paste(image, (x, y + label_h))
        canvas.save(out_dir / f"contact-{sheet_index // 4 + 1:02d}.png")
    print(f"{name}: pages={len(pages)}")
