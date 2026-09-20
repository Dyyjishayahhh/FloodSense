from collections import Counter
from pathlib import Path

import pdfplumber


source = Path("1. FLOODSENSE-MAIN/GUIDE FILES/SURECUT-MANUSCRIPT.pdf")
output = Path("tmp/manuscript_work/surecut_extracted.txt")

with pdfplumber.open(source) as pdf:
    print(f"pages={len(pdf.pages)}")
    for index, page in enumerate(pdf.pages, start=1):
        chars = page.chars
        fonts = Counter((c.get("fontname"), round(float(c.get("size", 0)), 1)) for c in chars)
        common = fonts.most_common(5)
        if chars:
            x0 = min(float(c["x0"]) for c in chars)
            x1 = max(float(c["x1"]) for c in chars)
            top = min(float(c["top"]) for c in chars)
            bottom = max(float(c["bottom"]) for c in chars)
        else:
            x0 = x1 = top = bottom = 0
        print(
            f"page={index} size={page.width:.1f}x{page.height:.1f} "
            f"bounds=({x0:.1f},{top:.1f})-({x1:.1f},{bottom:.1f}) fonts={common}"
        )

    with output.open("w", encoding="utf-8") as stream:
        for index, page in enumerate(pdf.pages, start=1):
            stream.write(f"\n===== PAGE {index} =====\n")
            stream.write(page.extract_text(layout=True) or "")
            stream.write("\n")

print(output)
