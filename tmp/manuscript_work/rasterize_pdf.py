from pathlib import Path
import pypdfium2 as pdfium

pdf_path = Path(r"tmp/manuscript_work/final_visual_preview.pdf")
out_dir = Path(r"tmp/manuscript_work/visual_preview_pages")
out_dir.mkdir(parents=True, exist_ok=True)
doc = pdfium.PdfDocument(pdf_path)
for index in range(len(doc)):
    image = doc[index].render(scale=1.5).to_pil()
    image.save(out_dir / f"page-{index + 1}.png")
print(len(doc))
