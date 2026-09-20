from pathlib import Path
from PIL import Image, ImageDraw

folder = Path(r"tmp/manuscript_work/visual_preview_pages")
pages = sorted(folder.glob("page-*.png"), key=lambda p: int(p.stem.split("-")[1]))
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
    out = folder / f"contact-{sheet_index // 4 + 1}.png"
    canvas.save(out)
print((len(pages) + 3) // 4)
