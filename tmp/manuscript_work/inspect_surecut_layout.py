from pathlib import Path
import json
import pdfplumber
import pypdfium2 as pdfium
from PIL import Image, ImageDraw

pdf_path = Path(r"1. FLOODSENSE-MAIN/GUIDE FILES/SURECUT-MANUSCRIPT.pdf")
out = Path(r"tmp/manuscript_work/surecut_layout")
out.mkdir(parents=True, exist_ok=True)
physical_pages = [22, 23, 24, 25, 26, 27, 28, 36, 37, 38, 39, 63, 64, 65, 66, 67, 116, 117, 118, 119, 120]

pdf = pdfium.PdfDocument(pdf_path)
for number in physical_pages:
    image = pdf[number - 1].render(scale=1.7).to_pil()
    image.save(out / f"surecut-{number}.png")

with pdfplumber.open(pdf_path) as source:
    metrics = []
    for number in physical_pages:
        page = source.pages[number - 1]
        words = page.extract_words(extra_attrs=["fontname", "size"])
        metrics.append({
            "page": number,
            "width": page.width,
            "height": page.height,
            "first_words": words[:35],
            "font_sizes": sorted({round(float(w["size"]), 2) for w in words}),
            "font_names": sorted({w["fontname"] for w in words}),
        })
    (out / "metrics.json").write_text(json.dumps(metrics, indent=2), encoding="utf-8")

paths = [out / f"surecut-{n}.png" for n in physical_pages]
cell_w, cell_h, label_h = 700, 990, 34
for start in range(0, len(paths), 4):
    group = paths[start:start + 4]
    canvas = Image.new("RGB", (cell_w * 2 + 30, (cell_h + label_h) * 2 + 30), "#777777")
    draw = ImageDraw.Draw(canvas)
    for offset, path in enumerate(group):
        image = Image.open(path).convert("RGB")
        image.thumbnail((cell_w, cell_h))
        x = 10 + (offset % 2) * (cell_w + 10)
        y = 10 + (offset // 2) * (cell_h + label_h + 10)
        draw.rectangle([x, y, x + cell_w, y + label_h - 2], fill="white")
        draw.text((x + 8, y + 7), path.stem, fill="black")
        canvas.paste(image, (x, y + label_h))
    canvas.save(out / f"contact-{start // 4 + 1}.png")
print(len(paths), len(paths + []))
