from pathlib import Path
import pypdfium2 as pdfium


ROOT = Path(__file__).parent
OUT = ROOT / "pages"
OUT.mkdir(exist_ok=True)
for pdf_path in sorted(ROOT.glob("*.pdf")):
    pdf = pdfium.PdfDocument(pdf_path)
    for index in range(len(pdf)):
        image = pdf[index].render(scale=2.0).to_pil()
        image.save(OUT / f"{pdf_path.stem}_p{index+1}.png")
    print(pdf_path.name, len(pdf))
