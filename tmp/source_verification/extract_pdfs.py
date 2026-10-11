from pathlib import Path
from pypdf import PdfReader


for pdf in Path(__file__).parent.glob("*.pdf"):
    if not pdf.stat().st_size:
        continue
    try:
        reader = PdfReader(pdf)
        text = "\n\n".join((p.extract_text() or "") for p in reader.pages)
        pdf.with_suffix(".txt").write_text(text, encoding="utf-8")
        print(pdf.name, len(reader.pages), len(text))
    except Exception as exc:
        print(pdf.name, "ERROR", exc)
