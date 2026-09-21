# FloodSense Chapter 2 revision contract

## Reference

- Retained manuscript: `E:\School Downloads-Jiraah\00 FILES IN SCHOOL\3RD YEAR 2ND SEM\DCIT 60\FLOODSENSE\1.FLOODSENSE-200A\FloodSense\1. FLOODSENSE-MAIN\MAIN FILESS\FloodSense_Chapters_1_to_3_Development_Aligned_2026-09-20.docx`
- SHA-256: `40CFE0F4BC0B09B38B7F5BBE88C9080D23A328CF73D3052802FE2D55AFD688B8`
- Baseline page count: 41 pages in the verified Microsoft Word render.
- Section count: 4.
- Style evidence: `tmp/ch2_revision_2026-09-21/template-style-evidence.json`.
- Baseline render evidence: `tmp/manuscript_work/rendered_development_aligned/`.
- Visual structure authority: `1. FLOODSENSE-MAIN/GUIDE FILES/SURECUT-MANUSCRIPT.pdf`, especially Chapter 2 on PDF pages 36-62.

## Page system

- A4 portrait, 8.27 by 11.69 inches.
- Main manuscript margins: left 1.50 inches; right, top, and bottom 1.00 inch.
- Four existing manuscript sections and their section breaks are preserve-only.
- Existing preliminary-page, Chapter 1, Chapter 3, and reference-list headers, footers, page-number fields, and section relationships are preserve-only.
- The existing narrow-margin section belongs to an embedded Chapter 1 figure and must not be changed.

## Typography and paragraph roles

- Body and heading family: Arial, black.
- Chapter title: centered, uppercase, bold, approximately 12 points; page break before.
- Major section heading: left aligned, bold, approximately 11 points.
- Topic or study lead: left aligned, bold, approximately 11 points.
- Body: 11 points, justified, first-line indent approximately 0.50 inch, double spaced, no decorative rules.
- Table caption: left aligned and italic, above the table.
- Reference entry: 11 points, hanging indent, single spaced within entries with separation between entries.

## Editable slots

- `word/document.xml`: replace only the content beginning with the paragraph `REVIEW OF RELATED LITERATURE` and ending immediately before `METHODOLOGY`.
- `word/document.xml`: replace the existing Chapter 2 comparison table with the three approved comparative-synthesis tables.
- `word/document.xml`: update only reference-list paragraphs that directly support Chapter 2; preserve all unrelated references.
- `FloodSense_Development_Manuscript_Alignment_Tracker.md`: update only Chapter 2 documentation-alignment entries and audit notes authorized by the user.

## Preserve-only content

- Every paragraph, table, drawing, section property, header, footer, and relationship before `REVIEW OF RELATED LITERATURE`, except page-number movement caused by the expanded Chapter 2.
- Every paragraph, table, drawing, section property, header, footer, and relationship from `METHODOLOGY` through the paragraph immediately before `REFERENCES`.
- Unrelated reference entries.
- Original `FloodSense_Chapter_1_MAIN.docx` and the retained development-aligned manuscript.

## Chapter 2 content plan

- Related Foreign Literature.
- Related Local Literature.
- Related Foreign Studies.
- Related Local Studies using exact verified study or application titles as bold lead-ins.
- Synthesis of the Reviewed Literature.
- Research Gap.
- Comparative Synthesis of Related Studies with three concise tables.
- Do not include the unreviewed Imus River Basin preprint in the formal manuscript.
- Do include the verified Leyte flood-warning conference system as a clearly bounded contrast; do not transfer its real-time monitoring or alert functions to FloodSense.

## Fidelity gates

- Canonical LibreOffice rendering was attempted but unavailable because the packaged environment contains no `soffice.exe`; Microsoft Word PDF export is the required fallback.
- Extract and compare Chapter 1 and Chapter 3 text before and after editing; canonical text hashes must match.
- Confirm all original non-Chapter-2 tables and drawings remain present.
- Render every final page, inspect all Chapter 2 pages at full size, and compare its hierarchy and table treatment with SURECUT Chapter 2.
- Confirm no code paths, commit hashes, repository narration, survey results, or unsupported implementation claims appear in formal Chapter 2.
