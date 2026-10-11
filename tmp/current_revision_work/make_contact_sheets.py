from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

root = Path(r"E:\School Downloads-Jiraah\00 FILES IN SCHOOL\3RD YEAR 2ND SEM\DCIT 60\FLOODSENSE\1.FLOODSENSE-200A\FloodSense")
pages = sorted((root / "tmp" / "docx" / "phase1_pages_v4").glob("page-*.png"))
out = root / "tmp" / "docx" / "phase1_contact_v4"
out.mkdir(parents=True, exist_ok=True)
font = ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf", 28)
for start in range(0, len(pages), 8):
    group = pages[start:start+8]
    sheet = Image.new("RGB", (2000, 2840), "#DDDDDD")
    draw = ImageDraw.Draw(sheet)
    for i, path in enumerate(group):
        im = Image.open(path).convert("RGB")
        im.thumbnail((475, 1320))
        x = 15 + (i % 4) * 495
        y = 55 + (i // 4) * 1390
        sheet.paste(im, (x, y))
        draw.text((x, 15 + (i // 4) * 1390), f"Page {start+i+1}", font=font, fill="black")
    sheet.save(out / f"contact_{start+1:02d}_{start+len(group):02d}.png")
print(f"{len(pages)} pages -> {len(list(out.glob('contact_*.png')))} contact sheets")
