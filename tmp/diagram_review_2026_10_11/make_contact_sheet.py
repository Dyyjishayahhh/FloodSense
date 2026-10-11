from pathlib import Path
from PIL import Image, ImageOps, ImageDraw


root = Path(__file__).parent / "pages"
files = sorted(root.glob("*.png"))
thumbs = []
for path in files:
    img = Image.open(path).convert("RGB")
    img.thumbnail((430, 560))
    canvas = Image.new("RGB", (460, 610), "white")
    x = (460 - img.width) // 2
    canvas.paste(img, (x, 28))
    ImageDraw.Draw(canvas).text((10, 5), path.stem, fill="black")
    thumbs.append(canvas)

for batch_idx in range(0, len(thumbs), 6):
    batch = thumbs[batch_idx:batch_idx+6]
    sheet = Image.new("RGB", (460 * 3, 610 * 2), "#cccccc")
    for i, thumb in enumerate(batch):
        sheet.paste(thumb, ((i % 3) * 460, (i // 3) * 610))
    sheet.save(root.parent / f"contact_{batch_idx//6 + 1}.png")
