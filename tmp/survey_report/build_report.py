from __future__ import annotations

from collections import Counter
from datetime import datetime
from pathlib import Path
import json
import math
import re
import textwrap

import openpyxl
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(r"E:\School Downloads-Jiraah\00 FILES IN SCHOOL\3RD YEAR 2ND SEM\DCIT 60\FLOODSENSE\1.FLOODSENSE-200A\FloodSense")
OUT_DIR = ROOT / "1. FLOODSENSE-MAIN" / "QUESTIONNAIRES_SURVEY" / "FLOODSENSE_INITIAL_SURVEY_RESULTS_AND_SUMMARY"
SOURCE = OUT_DIR / "FloodSense Resident Survey Questionnaire (Responses).xlsx"
DOCX_OUT = OUT_DIR / "FloodSense_Initial_Survey_Results_and_Summary.docx"
TMP_DIR = ROOT / "tmp" / "survey_report"
CHART_DIR = TMP_DIR / "charts"
CHART_DIR.mkdir(parents=True, exist_ok=True)

BLUE = "17365D"
MID_BLUE = "2F75B5"
LIGHT_BLUE = "D9EAF7"
PALE_BLUE = "EEF5FA"
GRAY = "666666"
LIGHT_GRAY = "F2F2F2"
WHITE = "FFFFFF"
BLACK = "000000"
ORANGE = "ED7D31"


def clean(value):
    if value is None:
        return ""
    normalized = (
        str(value)
        .replace("\ufffd", "-")
        .replace("\u2011", "-")
        .replace("\u2012", "-")
        .replace("\u2013", "-")
        .replace("\u2014", "-")
    )
    return " ".join(normalized.strip().split())


def english(value):
    return clean(value).split(" / ")[0].strip()


def pct(count, denominator):
    return 0.0 if denominator == 0 else round(count * 100.0 / denominator, 1)


wb = openpyxl.load_workbook(SOURCE, data_only=True)
ws = wb[wb.sheetnames[0]]
headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
raw_rows = [
    [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
    for r in range(2, ws.max_row + 1)
    if any(ws.cell(r, c).value not in (None, "") for c in range(1, ws.max_column + 1))
]
N = len(raw_rows)


def values(col):
    return [clean(row[col]) for row in raw_rows if clean(row[col])]


def ordered_counts(col, ordered_labels=None):
    counter = Counter(values(col))
    if ordered_labels:
        rows = []
        used = set()
        for label in ordered_labels:
            match = next((v for v in counter if english(v) == label or v == label), None)
            if match is not None:
                rows.append((label, counter[match]))
                used.add(match)
            else:
                rows.append((label, 0))
        for label, count in counter.most_common():
            if label not in used:
                rows.append((english(label), count))
        return rows
    return [(english(label), count) for label, count in counter.most_common()]


def multi_counts(col, options):
    result = []
    cells = values(col)
    for label, needle in options:
        count = sum(1 for cell in cells if needle.lower() in cell.lower())
        result.append((label, count))
    return result


def table_rows(counts, denominator=N):
    return [[label, str(count), f"{pct(count, denominator):.1f}%"] for label, count in counts]


consent_count = sum(1 for v in values(1) if v.startswith("Yes"))
adult_count = sum(1 for v in values(2) if v.startswith("Yes"))
resident_count = sum(1 for v in values(3) if v.startswith("Yes"))
experience_count = sum(1 for v in values(4) if v.startswith("Yes"))
name_count = len(values(5))

barangay_counts = ordered_counts(6)
age_counts = ordered_counts(
    7,
    [
        "18-24 years old",
        "25-34 years old",
        "35-44 years old",
        "45-54 years old",
        "55-64 years old",
        "65 years old or older",
    ],
)

impact_options = [
    ("Floodwater entered the home", "Floodwater entered our house"),
    ("Street or community flooding", "Floodwater affected our street or community"),
    ("Usual travel route difficult or impassable", "My usual travel route became difficult or impassable"),
    ("Work or studies disrupted", "My work or studies were disrupted"),
    ("Evacuated or temporarily relocated", "We evacuated or temporarily relocated"),
    ("Property, utilities, or essential services affected", "Property, utilities, or essential services were affected"),
]
impact_counts = multi_counts(8, impact_options)

event_counts = ordered_counts(
    9,
    [
        "1 event",
        "2-3 events",
        "4-5 events",
        "More than 5 events",
        "Not sure",
    ],
)

source_options = [
    ("Barangay announcement or public-address system", "Barangay announcements or public-address system"),
    ("Bacoor City or BDRRMO official social-media page", "Bacoor City or BDRRMO official social-media page"),
    ("PAGASA or another government agency", "PAGASA or another government agency"),
    ("Television or radio news", "Television or radio news"),
    ("Online news outlet", "Online news outlet"),
    ("SMS or text message", "SMS or text message"),
    ("Facebook groups or other social-media posts", "Facebook groups or other social-media posts"),
    ("Family, friends, or neighbors", "Family, friends, or neighbors"),
    ("Community organization or homeowners' association", "Community organization or homeowners"),
    ("Government website or online bulletin", "Government website or online bulletin"),
    ("Usually receives no flood information", "I usually receive no flood information"),
]
source_counts = multi_counts(10, source_options)

access_counts = ordered_counts(
    11,
    [
        "Very easy",
        "Easy",
        "Neither easy nor difficult",
        "Difficult",
        "Very difficult",
        "I usually cannot obtain the information",
        "Not sure",
    ],
)

problem_labels = [
    "Information arrives too late or is outdated",
    "Information is difficult to understand",
    "Information is too general or not location-specific",
    "Different sources provide conflicting information",
    "Source or date is unclear",
    "Information is difficult to access when internet or electricity is unreliable",
    "Preparedness instructions are too general",
    "Verified evacuation-center information is difficult to find",
]
frequency_order = ["Never", "Rarely", "Sometimes", "Often", "Always", "Not sure"]
frequency_weights = {"Never": 1, "Rarely": 2, "Sometimes": 3, "Often": 4, "Always": 5}


def frequency_stats(col, label):
    counts = {key: 0 for key in frequency_order}
    score_sum = 0
    score_n = 0
    for value in values(col):
        key = english(value)
        if key in counts:
            counts[key] += 1
        if key in frequency_weights:
            score_sum += frequency_weights[key]
            score_n += 1
    mean = score_sum / score_n if score_n else math.nan
    often_always = counts["Often"] + counts["Always"]
    return {
        "label": label,
        "counts": counts,
        "valid_n": score_n,
        "mean": mean,
        "often_always": often_always,
    }


problem_stats = [frequency_stats(12 + i, problem_labels[i]) for i in range(8)]
problem_stats_ranked = sorted(problem_stats, key=lambda item: item["mean"], reverse=True)
decision_frequency = frequency_stats(20, "Decision-making affected by inadequate information")

decision_options = [
    ("Preparing emergency supplies", "Preparing emergency supplies"),
    ("Moving belongings to a safer place", "Moving important belongings to a safer place"),
    ("Deciding whether it is safe to travel", "Deciding whether it is safe to travel"),
    ("Choosing a road or area to avoid", "Choosing a road or area to avoid"),
    ("Deciding whether or when to evacuate", "Deciding whether or when to evacuate"),
    ("Identifying an evacuation center", "Identifying an evacuation center"),
    ("Deciding whether to attend work or school", "Deciding whether to attend work or school"),
    ("Helping household members who may need assistance", "Helping children, older persons"),
    ("None of the above", "None of the above"),
]
decision_counts = multi_counts(21, decision_options)

susceptibility_counts = ordered_counts(
    22,
    [
        "Very easy",
        "Easy",
        "Neither easy nor difficult",
        "Difficult",
        "Very difficult",
        "I do not know how to determine it",
        "Not sure",
    ],
)

evac_find_counts = ordered_counts(
    23,
    [
        "Very easy",
        "Easy",
        "Neither easy nor difficult",
        "Difficult",
        "Very difficult",
        "I do not know where to find it",
        "I have not needed this information",
        "Not sure",
    ],
)

priority_options = [
    ("Flood susceptibility of the barangay or selected location", "Flood susceptibility of my barangay or selected location"),
    ("Possible effects of rainfall intensity and duration", "Possible effects of rainfall intensity and duration"),
    ("Current road or travel conditions", "Current road or travel conditions"),
    ("Pre-event household preparedness actions", "Pre-event household preparedness actions"),
    ("Evacuation-center references", "Evacuation-center references"),
    ("Map visualization of the selected location and susceptibility result", "Map-based visualization of the selected location"),
    ("Verified evacuation-center names and map locations", "Verified evacuation-center names and map locations"),
]
priority_counts = multi_counts(24, priority_options)

guidance_options = [
    ("Guidance is too general", "The guidance is too general"),
    ("Guidance is not specific to location or situation", "It is not specific to my location or situation"),
    ("Recommended actions are difficult to understand", "The recommended actions are difficult to understand"),
    ("Unclear which action should be done first", "It is unclear which action should be done first"),
    ("Source of guidance is unclear", "The source of the guidance is unclear"),
    ("Guidance is difficult to access", "The guidance is difficult to access"),
    ("Usually receives no preparedness guidance", "I do not usually receive preparedness guidance"),
    ("No problem experienced with available guidance", "I have not experienced a problem with available guidance"),
    ("Not sure", "Not sure"),
]
guidance_counts = multi_counts(25, guidance_options)

evac_detail_options = [
    ("Center name", "Name of the center"),
    ("Address or map location", "Address or map location"),
    ("Approximate straight-line distance", "Approximate straight-line distance"),
    ("Contact information", "Contact information"),
    ("Information source or responsible office", "Information source or responsible office"),
    ("Date last verified", "Date the information was last verified"),
    ("Does not need evacuation-center information", "I do not need evacuation-center information"),
    ("Not sure", "Not sure"),
]
evac_detail_counts = multi_counts(26, evac_detail_options)

android_counts = ordered_counts(
    27,
    [
        "Yes, my own device",
        "Yes, a shared device",
        "No",
        "Not sure whether it is Android",
    ],
)

internet_counts = ordered_counts(
    28,
    [
        "Always reliable",
        "Usually reliable",
        "Sometimes reliable",
        "Rarely reliable",
        "No internet access",
        "Not sure",
    ],
)

location_options = [
    ("Manual barangay selection", "Manual barangay selection"),
    ("Temporary map pin without saved location history", "Temporary map pin that is not saved"),
    ("One-time GPS after explanation and permission", "One-time GPS access after an explanation"),
    ("Not comfortable providing or selecting a location", "I am not comfortable providing or selecting a location"),
    ("Not sure", "Not sure"),
]
location_counts = multi_counts(29, location_options)

language_counts = ordered_counts(
    30,
    ["English", "Filipino", "Both English and Filipino"],
)


def contradiction_count(col, negative_needle, positive_needles):
    count = 0
    for cell in values(col):
        low = cell.lower()
        if negative_needle.lower() in low and any(n.lower() in low for n in positive_needles):
            count += 1
    return count


contradictions = {
    "no information plus information sources": contradiction_count(
        10,
        "I usually receive no flood information",
        [x[1] for x in source_options[:-1]],
    ),
    "none plus affected decisions": contradiction_count(
        21,
        "None of the above",
        [x[1] for x in decision_options[:-1]],
    ),
    "no guidance problem plus another guidance problem": contradiction_count(
        25,
        "I have not experienced a problem with available guidance",
        [x[1] for x in guidance_options[:7]],
    ),
    "no evacuation information need plus requested details": contradiction_count(
        26,
        "I do not need evacuation-center information",
        [x[1] for x in evac_detail_options[:6]],
    ),
    "location discomfort plus accepted method": contradiction_count(
        29,
        "I am not comfortable providing or selecting a location",
        [x[1] for x in location_options[:3]],
    ),
}

signatures = [
    tuple(clean(value).lower() for idx, value in enumerate(row) if idx not in {0, 5})
    for row in raw_rows
]
duplicate_groups = sorted(
    [count for count in Counter(signatures).values() if count > 1], reverse=True
)

timestamps = [row[0] for row in raw_rows if isinstance(row[0], datetime)]
rapid_pairs = []
for index in range(1, len(timestamps)):
    delta = (timestamps[index] - timestamps[index - 1]).total_seconds()
    if delta < 10:
        rapid_pairs.append(delta)

comments = [clean(row[31]) for row in raw_rows if clean(row[31])]
comment_norms = [comment.lower() for comment in comments]
comment_repeats = {text: count for text, count in Counter(comment_norms).items() if count > 1}


def has_any(text, terms):
    low = text.lower()
    return any(term in low for term in terms)


open_themes = [
    (
        "Map, GPS, barangay, or location information",
        ["gps", "map", "specific location", "baranggay", "brggy", "brgy", "barangay", "location ng"],
    ),
    (
        "Real-time conditions, flood timing, or early warning",
        ["real-time", "real time", "kailan babaha", "if baha na", "early warning", "alerts"],
    ),
    (
        "Flood depth or severity",
        ["kataas", "water levels", "severity"],
    ),
    (
        "Evacuation-center location or contact information",
        ["evacuation", "evacuatan"],
    ),
    (
        "Detailed or clearer announcements",
        ["detalyadong", "details para", "detailed"],
    ),
    (
        "Hazard-map details such as drainage canals or open manholes",
        ["drainage canals", "open manholes", "hidden dangers"],
    ),
    (
        "No additional information identified",
        ["none"],
    ),
]
theme_counts = []
for label, terms in open_themes:
    count = sum(1 for comment in comments if has_any(comment, terms))
    theme_counts.append((label, count))


def single_count(counts, label):
    return next((count for item, count in counts if item == label), 0)


hard_to_obtain = (
    single_count(access_counts, "Difficult")
    + single_count(access_counts, "Very difficult")
    + single_count(access_counts, "I usually cannot obtain the information")
)
decision_often_always = decision_frequency["often_always"]
susceptibility_challenge = (
    single_count(susceptibility_counts, "Difficult")
    + single_count(susceptibility_counts, "Very difficult")
    + single_count(susceptibility_counts, "I do not know how to determine it")
)
evac_challenge = (
    single_count(evac_find_counts, "Difficult")
    + single_count(evac_find_counts, "Very difficult")
    + single_count(evac_find_counts, "I do not know where to find it")
)
android_access = (
    single_count(android_counts, "Yes, my own device")
    + single_count(android_counts, "Yes, a shared device")
)
reliable_internet = (
    single_count(internet_counts, "Always reliable")
    + single_count(internet_counts, "Usually reliable")
)


def _font(size, bold=False):
    font_name = "arialbd.ttf" if bold else "arial.ttf"
    try:
        return ImageFont.truetype(str(Path(r"C:\Windows\Fonts") / font_name), size)
    except OSError:
        return ImageFont.load_default()


def _wrapped_label(label, width=42):
    return textwrap.wrap(label, width=width, break_long_words=False, break_on_hyphens=False) or [""]


def plot_horizontal(counts, title, filename, denominator=N, color=MID_BLUE, max_items=None):
    items = [(label, count) for label, count in counts if count > 0]
    if max_items:
        items = sorted(items, key=lambda pair: pair[1], reverse=True)[:max_items]
    max_value = max(N, max((count for _, count in items), default=0), 1)
    wrapped = [(_wrapped_label(label), count) for label, count in items]
    row_heights = [max(84, 29 * len(lines) + 30) for lines, _ in wrapped]
    width = 1800
    top = 135
    bottom = 100
    height = top + sum(row_heights) + bottom
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    title_font = _font(34, bold=True)
    label_font = _font(24)
    value_font = _font(22, bold=True)
    axis_font = _font(20)
    draw.text((55, 40), title, fill="#17365D", font=title_font)
    x0, x1 = 820, 1550
    plot_width = x1 - x0
    for tick in range(0, max_value + 1, max(1, max_value // 4)):
        x = x0 + plot_width * tick / max_value
        draw.line((x, top - 8, x, height - bottom + 10), fill="#E0E0E0", width=2)
        tick_text = str(tick)
        box = draw.textbbox((0, 0), tick_text, font=axis_font)
        draw.text((x - (box[2] - box[0]) / 2, height - bottom + 25), tick_text, fill="#555555", font=axis_font)
    y = top
    for (lines, value), row_h in zip(wrapped, row_heights):
        text_h = len(lines) * 29
        label_y = y + (row_h - text_h) / 2
        for idx, line in enumerate(lines):
            draw.text((55, label_y + idx * 29), line, fill="#222222", font=label_font)
        bar_y0 = y + row_h * 0.25
        bar_y1 = y + row_h * 0.75
        bar_x1 = x0 + plot_width * value / max_value
        draw.rounded_rectangle((x0, bar_y0, max(x0 + 2, bar_x1), bar_y1), radius=8, fill=f"#{color}")
        value_text = f"{value} ({pct(value, denominator):.1f}%)"
        draw.text((min(bar_x1 + 14, 1590), y + (row_h - 24) / 2), value_text, fill="#222222", font=value_font)
        y += row_h
    draw.text((x0 + 235, height - 42), "Number of respondents", fill="#555555", font=axis_font)
    path = CHART_DIR / filename
    image.save(path, dpi=(220, 220))
    return path


def plot_problem_means(stats, filename):
    ordered = sorted(stats, key=lambda item: item["mean"], reverse=True)
    wrapped = [(_wrapped_label(item["label"], 48), item["mean"]) for item in ordered]
    row_heights = [max(94, 29 * len(lines) + 34) for lines, _ in wrapped]
    width = 1800
    top = 150
    bottom = 125
    height = top + sum(row_heights) + bottom
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    title_font = _font(34, bold=True)
    label_font = _font(23)
    value_font = _font(23, bold=True)
    axis_font = _font(19)
    draw.text((55, 38), "Mean frequency of problems with existing flood-information sources", fill="#17365D", font=title_font)
    x0, x1 = 850, 1540
    plot_width = x1 - x0
    for tick in range(1, 6):
        x = x0 + plot_width * (tick - 1) / 4
        draw.line((x, top - 8, x, height - bottom + 10), fill="#E0E0E0", width=2)
        tick_labels = {1: "Never\n1", 2: "Rarely\n2", 3: "Sometimes\n3", 4: "Often\n4", 5: "Always\n5"}
        parts = tick_labels[tick].split("\n")
        for idx, part in enumerate(parts):
            box = draw.textbbox((0, 0), part, font=axis_font)
            draw.text((x - (box[2] - box[0]) / 2, height - bottom + 22 + idx * 23), part, fill="#555555", font=axis_font)
    y = top
    for (lines, value), row_h in zip(wrapped, row_heights):
        text_h = len(lines) * 29
        label_y = y + (row_h - text_h) / 2
        for idx, line in enumerate(lines):
            draw.text((55, label_y + idx * 29), line, fill="#222222", font=label_font)
        bar_y0 = y + row_h * 0.25
        bar_y1 = y + row_h * 0.75
        bar_x1 = x0 + plot_width * (value - 1) / 4
        draw.rounded_rectangle((x0, bar_y0, max(x0 + 2, bar_x1), bar_y1), radius=8, fill=f"#{ORANGE}")
        draw.text((min(bar_x1 + 14, 1600), y + (row_h - 24) / 2), f"{value:.2f}", fill="#222222", font=value_font)
        y += row_h
    path = CHART_DIR / filename
    image.save(path, dpi=(220, 220))
    return path


chart_barangay = plot_horizontal(barangay_counts, "Respondents by barangay", "barangays.png")
chart_impacts = plot_horizontal(impact_counts, "Effects of flooding or heavy rainfall", "impacts.png")
chart_sources = plot_horizontal(source_counts, "Current flood-information sources", "sources.png")
chart_problems = plot_problem_means(problem_stats, "problem_means.png")
chart_decisions = plot_horizontal(decision_counts[:-1], "Decisions affected by insufficient flood information", "decisions.png")
chart_priorities = plot_horizontal(priority_counts, "Information considered most important before possible flooding", "priorities.png")
chart_guidance = plot_horizontal(guidance_counts[:-2], "Problems with existing preparedness guidance", "guidance.png")
chart_evac = plot_horizontal(evac_detail_counts[:-2], "Useful evacuation-center details", "evac_details.png")


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=90, start=110, bottom=90, end=110):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_borders(table, color="D9D9D9", size="6"):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = borders.find(qn(f"w:{edge}"))
        if tag is None:
            tag = OxmlElement(f"w:{edge}")
            borders.append(tag)
        tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), size)
        tag.set(qn("w:color"), color)


def repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def keep_paragraph(paragraph, next_paragraph=False):
    p_pr = paragraph._p.get_or_add_pPr()
    keep = OxmlElement("w:keepNext" if next_paragraph else "w:keepLines")
    p_pr.append(keep)


def set_repeat_no_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def set_column_widths(table, widths):
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = Inches(width)


def style_run(run, size=None, bold=None, color=None, italic=None, font="Arial"):
    run.font.name = font
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), font)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), font)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def add_table(doc, headers_row, body_rows, widths=None, numeric_cols=None, font_size=9.0):
    numeric_cols = set(numeric_cols or [])
    table = doc.add_table(rows=1, cols=len(headers_row))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)
    hdr = table.rows[0]
    repeat_table_header(hdr)
    for idx, text in enumerate(headers_row):
        cell = hdr.cells[idx]
        set_cell_shading(cell, BLUE)
        set_cell_margins(cell)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(str(text))
        style_run(run, size=9.0, bold=True, color=WHITE)
    for row_index, row_data in enumerate(body_rows):
        cells = table.add_row().cells
        if row_index % 2 == 1:
            for cell in cells:
                set_cell_shading(cell, PALE_BLUE)
        set_repeat_no_split(table.rows[-1])
        for col_index, text in enumerate(row_data):
            cell = cells[col_index]
            set_cell_margins(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_index in numeric_cols else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(str(text))
            style_run(run, size=font_size)
    if widths:
        set_column_widths(table, widths)
    return table


def add_heading(doc, text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.space_before = Pt(12 if level == 1 else 9)
    p.paragraph_format.space_after = Pt(5)
    keep_paragraph(p, next_paragraph=True)
    run = p.add_run(text)
    style_run(run, size=15 if level == 1 else 12, bold=True, color=BLACK)
    return p


def add_body(doc, text, bold_lead=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.12
    if bold_lead and text.startswith(bold_lead):
        r1 = p.add_run(bold_lead)
        style_run(r1, size=10.5, bold=True)
        r2 = p.add_run(text[len(bold_lead):])
        style_run(r2, size=10.5)
    else:
        run = p.add_run(text)
        style_run(run, size=10.5)
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.18)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(text)
    style_run(run, size=10.3)
    return p


def add_chart(doc, path, caption, width=6.7):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(3)
    p.add_run().add_picture(str(path), width=Inches(width))
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(8)
    run = cap.add_run(caption)
    style_run(run, size=9, italic=True, color=GRAY)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, text, end])
    style_run(run, size=9, color=GRAY)


doc = Document()
section = doc.sections[0]
section.page_width = Inches(8.5)
section.page_height = Inches(11)
section.top_margin = Inches(0.75)
section.bottom_margin = Inches(0.7)
section.left_margin = Inches(0.8)
section.right_margin = Inches(0.8)

styles = doc.styles
normal = styles["Normal"]
normal.font.name = "Arial"
normal._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
normal._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
normal.font.size = Pt(10.5)
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.12

title_style = styles["Title"]
title_style.font.name = "Arial"
title_style._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
title_style._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
title_style.font.size = Pt(24)
title_style.font.bold = True
title_style.font.color.rgb = RGBColor.from_string(BLACK)

for sec in doc.sections:
    footer = sec.footer
    footer_p = footer.paragraphs[0]
    add_page_number(footer_p)

# Cover page
doc.add_paragraph().paragraph_format.space_after = Pt(55)
title = doc.add_paragraph(style="Title")
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.paragraph_format.space_after = Pt(16)
title.add_run("FloodSense Initial Survey Results and Summary")

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
subtitle.paragraph_format.space_after = Pt(12)
r = subtitle.add_run("Flood Information and Pre-Event Preparedness Needs of Bacoor City Residents")
style_run(r, size=14, bold=True, color=BLUE)

meta = doc.add_paragraph()
meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
meta.paragraph_format.space_before = Pt(25)
meta.paragraph_format.space_after = Pt(6)
r = meta.add_run("Descriptive analysis of 40 qualified respondents")
style_run(r, size=11.5)

meta2 = doc.add_paragraph()
meta2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = meta2.add_run("Survey responses collected on 11 October 2026")
style_run(r, size=10.5, color=GRAY)

doc.add_paragraph().paragraph_format.space_after = Pt(70)
scope = doc.add_paragraph()
scope.alignment = WD_ALIGN_PARAGRAPH.CENTER
scope.paragraph_format.left_indent = Inches(0.75)
scope.paragraph_format.right_indent = Inches(0.75)
r = scope.add_run(
    "This report presents the complete descriptive results, integrated interpretation, research implications, and limitations of the FloodSense initial resident survey."
)
style_run(r, size=10.5, italic=True, color=GRAY)

doc.add_page_break()

add_heading(doc, "Executive Summary", 1)
add_body(
    doc,
    f"All {N} submissions met the survey's consent and eligibility requirements: every respondent consented, was at least 18 years old, lived in Bacoor City, and reported a flood or heavy-rainfall experience within the survey reference period. The results provide strong descriptive evidence of information-access and preparedness difficulties among this respondent group, but they do not establish citywide prevalence or prove the effectiveness of FloodSense before post-use evaluation."
)
add_body(
    doc,
    f"The clearest problem signal is the difficulty of obtaining adequate flood information. A total of {hard_to_obtain} respondents ({pct(hard_to_obtain, N):.1f}%) described obtaining needed information as difficult, very difficult, or usually impossible. In addition, {decision_often_always} respondents ({pct(decision_often_always, N):.1f}%) said late, unclear, incomplete, or conflicting information often or always affected their decisions. All eight assessed information problems received mean frequency scores within the 'Often' range."
)
add_body(
    doc,
    f"Existing channels are widely used, especially official Bacoor City or BDRRMO social-media pages, Facebook groups, news outlets, SMS, PAGASA or other agencies, and personal networks. Their continued use alongside high levels of reported difficulty supports the conclusion that the issue is not simply the absence of information. Respondents reported gaps in timeliness, clarity, consistency, accessibility, source transparency, location specificity, preparedness guidance, and verified evacuation-center information."
)
add_body(
    doc,
    f"The sample also indicates practical feasibility for an Android-based supplementary tool: {android_access} respondents ({pct(android_access, N):.1f}%) reported regular access to their own or a shared Android device, while {reliable_internet} ({pct(reliable_internet, N):.1f}%) described their internet access during heavy rainfall as always or usually reliable. Most respondents preferred bilingual English and Filipino presentation. Manual barangay selection, a temporary map pin, and one-time GPS access all received substantial acceptance."
)
add_body(
    doc,
    "The evidence supports a focused role for FloodSense as a supplementary, scenario-based pre-event decision-support application. The findings can justify location-specific susceptibility assessment, explainable results, map-based location selection, scenario-relevant preparedness guidance, and verified evacuation-center references. They do not justify claims of live flood forecasting, automatic warnings, real-time water-level monitoring, road-safety guidance, safe-route recommendation, live center availability, or official evacuation orders."
)

add_heading(doc, "Key Findings", 2)
for item in [
    f"{30} respondents ({pct(30, N):.1f}%) reported more than five flood or heavy-rainfall events within the three-year reference period.",
    f"{hard_to_obtain} respondents ({pct(hard_to_obtain, N):.1f}%) reported difficulty obtaining needed flood information or usually could not obtain it.",
    f"{decision_often_always} respondents ({pct(decision_often_always, N):.1f}%) said inadequate information often or always affected decisions.",
    f"{susceptibility_challenge} respondents ({pct(susceptibility_challenge, N):.1f}%) found it difficult, very difficult, or did not know how to determine whether their location might be susceptible to flooding.",
    f"{evac_challenge} respondents ({pct(evac_challenge, N):.1f}%) found verified evacuation-center information difficult to find or did not know where to find it.",
    f"{android_access} respondents ({pct(android_access, N):.1f}%) had regular access to an Android smartphone through their own or a shared device.",
    f"The sample covered {len(barangay_counts)} barangays, but {sum(count for label, count in barangay_counts if label == 'Niog' or label.startswith('Zapote'))} respondents ({pct(sum(count for label, count in barangay_counts if label == 'Niog' or label.startswith('Zapote')), N):.1f}%) came from Niog or the three Zapote barangays.",
]:
    add_bullet(doc, item)

doc.add_page_break()

add_heading(doc, "1 Survey Scope and Method", 1)
add_heading(doc, "1.1 Purpose", 2)
add_body(
    doc,
    "The initial resident survey was designed to identify current flood-information problems, their effects on household preparedness and decision-making, the information residents find difficult to obtain, and the feasibility requirements for an Android-based supplementary information tool. It was not a post-use evaluation of FloodSense."
)
add_heading(doc, "1.2 Respondent Eligibility and Data Collection", 2)
add_body(
    doc,
    f"The workbook contained {N} completed submissions collected on 11 October 2026. Eligibility required informed consent, an age of at least 18 years, current residence in Bacoor City, and experience of flooding or heavy rainfall that affected the household, community, travel, work, or studies during the stated three-year reference period. All {N} submissions met these conditions."
)
add_table(
    doc,
    ["Eligibility requirement", "Qualified respondents", "Percentage"],
    [
        ["Provided informed consent", consent_count, f"{pct(consent_count, N):.1f}%"],
        ["At least 18 years old", adult_count, f"{pct(adult_count, N):.1f}%"],
        ["Current Bacoor City resident", resident_count, f"{pct(resident_count, N):.1f}%"],
        ["Experienced flood or heavy-rainfall effects", experience_count, f"{pct(experience_count, N):.1f}%"],
    ],
    widths=[4.35, 1.15, 1.15],
    numeric_cols=[1, 2],
)

add_heading(doc, "1.3 Analytical Method", 2)
add_body(
    doc,
    "The report uses descriptive statistics. Single-response questions are summarized with frequencies and percentages. Checkbox questions use the number of respondents selecting each option, divided by 40; because a respondent could select more than one option, percentages for a checkbox question can total more than 100%."
)
add_body(
    doc,
    "For the eight frequency statements about existing information problems, responses were weighted as Never = 1, Rarely = 2, Sometimes = 3, Often = 4, and Always = 5. 'Not sure' responses were excluded only from the mean calculation and remain visible in the frequency tables. Mean ranges were interpreted as follows: 1.00-1.80 Never, 1.81-2.60 Rarely, 2.61-3.40 Sometimes, 3.41-4.20 Often, and 4.21-5.00 Always."
)
add_body(
    doc,
    f"The optional name or initials field was completed by {name_count} respondents ({pct(name_count, N):.1f}%). Names and initials were excluded from every analysis, table, chart, and narrative in this report."
)

add_heading(doc, "1.4 Data-Quality Review and Limitations", 2)
add_body(
    doc,
    "The results should be interpreted as evidence about the surveyed group, not as population estimates for all Bacoor residents. The study used a small purposive sample and did not use probability sampling or inferential statistical tests."
)
for item in [
    f"Geographic coverage included {len(barangay_counts)} of Bacoor's 47 barangays. Niog and Zapote 1, Zapote 2, and Zapote 3 accounted for 70.0% of the sample.",
    f"Respondents aged 18-34 accounted for {single_count(age_counts, '18-24 years old') + single_count(age_counts, '25-34 years old')} of {N} responses ({pct(single_count(age_counts, '18-24 years old') + single_count(age_counts, '25-34 years old'), N):.1f}%).",
    "The three-year recall period increases coverage of flood experience but may introduce recall error about the timing and details of older events.",
    f"One pair of submissions had identical answers across all substantive fields after excluding timestamp and optional name. The submissions were retained because independence could not be determined from the workbook alone.",
    f"Five consecutive submissions occurred less than ten seconds apart. This timing cluster does not prove invalidity, but it should be considered when assessing response independence.",
    f"Six optional comments repeated the same long recommendation word for word. Closed-ended responses were counted as submitted, while qualitative interpretation reports the repetition explicitly rather than treating it as six independent ideas.",
    "A few checkbox responses combined a mutually exclusive option with substantive choices. These responses were preserved as recorded, and affected categories were counted separately.",
]:
    add_bullet(doc, item)

contradiction_rows = [[label.capitalize(), count] for label, count in contradictions.items() if count]
if contradiction_rows:
    add_table(
        doc,
        ["Checkbox inconsistency observed", "Responses"],
        contradiction_rows,
        widths=[5.6, 1.05],
        numeric_cols=[1],
    )

doc.add_page_break()

add_heading(doc, "2 Respondent Profile", 1)
add_heading(doc, "2.1 Barangay Distribution", 2)
add_body(
    doc,
    f"The {N} respondents represented {len(barangay_counts)} barangays. Niog had the largest representation with {single_count(barangay_counts, 'Niog')} respondents ({pct(single_count(barangay_counts, 'Niog'), N):.1f}%). Zapote 1, Zapote 2, and Zapote 3 collectively contributed {sum(count for label, count in barangay_counts if label.startswith('Zapote'))} respondents ({pct(sum(count for label, count in barangay_counts if label.startswith('Zapote')), N):.1f}%)."
)
add_chart(doc, chart_barangay, "Figure 1. Distribution of respondents by barangay (n = 40).")
add_table(doc, ["Barangay", "Frequency", "Percentage"], table_rows(barangay_counts), widths=[4.25, 1.2, 1.2], numeric_cols=[1, 2])

add_heading(doc, "2.2 Age Distribution", 2)
add_body(
    doc,
    f"The sample was concentrated among younger adults. The 25-34 group accounted for {single_count(age_counts, '25-34 years old')} respondents ({pct(single_count(age_counts, '25-34 years old'), N):.1f}%), followed by the 18-24 group with {single_count(age_counts, '18-24 years old')} ({pct(single_count(age_counts, '18-24 years old'), N):.1f}%)."
)
add_table(doc, ["Age group", "Frequency", "Percentage"], table_rows(age_counts), widths=[4.25, 1.2, 1.2], numeric_cols=[1, 2])

doc.add_page_break()

add_heading(doc, "3 Flood and Heavy-Rainfall Experience", 1)
add_heading(doc, "3.1 Reported Effects", 2)
add_body(
    doc,
    "Respondents reported effects across the household, community, mobility, work or study, evacuation, property, and essential-service dimensions. Because this was a multiple-response question, each percentage represents the share of all 40 respondents selecting the effect."
)
add_chart(doc, chart_impacts, "Figure 2. Reported effects of flooding or heavy rainfall (multiple responses, n = 40).")
add_table(doc, ["Reported effect", "Frequency", "Percentage"], table_rows(impact_counts), widths=[4.55, 1.05, 1.05], numeric_cols=[1, 2])

add_heading(doc, "3.2 Number of Experienced Events", 2)
add_body(
    doc,
    f"Thirty respondents ({pct(30, N):.1f}%) reported more than five separate flood or heavy-rainfall events during the three-year reference period. The remaining ten ({pct(10, N):.1f}%) reported two to three events. No respondent selected one event, four to five events, or not sure."
)
add_table(doc, ["Reported number of events", "Frequency", "Percentage"], table_rows(event_counts), widths=[4.25, 1.2, 1.2], numeric_cols=[1, 2])

doc.add_page_break()

add_heading(doc, "4 Existing Flood-Information Sources", 1)
add_body(
    doc,
    "Respondents relied on multiple formal and informal channels. The continued use of many sources is important to interpretation: the reported problem is not that residents have no information channel at all, but that available channels do not consistently meet their needs."
)
add_chart(doc, chart_sources, "Figure 3. Sources currently used for flood or heavy-rainfall information (multiple responses, n = 40).")
add_table(doc, ["Information source", "Frequency", "Percentage"], table_rows(source_counts), widths=[4.55, 1.05, 1.05], numeric_cols=[1, 2])
add_body(
    doc,
    "One response selected both 'usually receives no flood information' and other sources. This contradiction was retained in the category counts but should not be interpreted as a consistent source-use pattern."
)

add_heading(doc, "5 Difficulty Obtaining Needed Information", 1)
add_body(
    doc,
    f"A combined {hard_to_obtain} respondents ({pct(hard_to_obtain, N):.1f}%) selected difficult, very difficult, or usually cannot obtain the information. Only {single_count(access_counts, 'Easy') + single_count(access_counts, 'Very easy')} ({pct(single_count(access_counts, 'Easy') + single_count(access_counts, 'Very easy'), N):.1f}%) selected easy or very easy."
)
add_table(doc, ["Ease of obtaining needed information", "Frequency", "Percentage"], table_rows(access_counts), widths=[4.25, 1.2, 1.2], numeric_cols=[1, 2])

doc.add_page_break()

add_heading(doc, "6 Problems with Existing Flood-Information Sources", 1)
add_body(
    doc,
    "Every assessed problem had a mean score in the 'Often' range. The highest mean concerned information that was too general or not location-specific. Other high-ranking problems involved difficulty accessing information when internet or electricity was unreliable, overly general preparedness instructions, unclear sources or dates, and difficulty finding verified evacuation-center information."
)
add_chart(doc, chart_problems, "Figure 4. Mean frequency of eight problems with existing information sources. Not-sure responses were excluded from means.")

problem_table_rows = []
for rank, item in enumerate(problem_stats_ranked, 1):
    problem_table_rows.append([
        rank,
        item["label"],
        f"{item['mean']:.2f}",
        "Often" if item["mean"] < 4.21 else "Always",
        f"{item['often_always']} ({pct(item['often_always'], N):.1f}%)",
        item["valid_n"],
    ])
add_table(
    doc,
    ["Rank", "Problem", "Mean", "Interpretation", "Often or always", "Valid n"],
    problem_table_rows,
    widths=[0.45, 3.25, 0.62, 0.88, 1.05, 0.55],
    numeric_cols=[0, 2, 4, 5],
    font_size=8.4,
)

add_heading(doc, "6.1 Complete Frequency Distribution", 2)
freq_rows = []
for item in problem_stats:
    c = item["counts"]
    freq_rows.append([
        item["label"], c["Never"], c["Rarely"], c["Sometimes"], c["Often"], c["Always"], c["Not sure"]
    ])
add_table(
    doc,
    ["Problem", "Never", "Rarely", "Sometimes", "Often", "Always", "Not sure"],
    freq_rows,
    widths=[3.25, 0.55, 0.55, 0.65, 0.55, 0.55, 0.65],
    numeric_cols=[1, 2, 3, 4, 5, 6],
    font_size=8.1,
)

doc.add_page_break()

add_heading(doc, "7 Effects on Preparedness and Decision-Making", 1)
add_heading(doc, "7.1 Frequency of Decision Effects", 2)
add_body(
    doc,
    f"Thirty respondents ({pct(decision_often_always, N):.1f}%) said late, unclear, incomplete, or conflicting information affected their decisions often or always. Only two respondents ({pct(2, N):.1f}%) selected never."
)
add_table(
    doc,
    ["Frequency", "Count", "Percentage"],
    [[label, decision_frequency["counts"][label], f"{pct(decision_frequency['counts'][label], N):.1f}%"] for label in frequency_order],
    widths=[4.25, 1.2, 1.2],
    numeric_cols=[1, 2],
)

add_heading(doc, "7.2 Types of Decisions Affected", 2)
add_chart(doc, chart_decisions, "Figure 5. Decisions made difficult, uncertain, or delayed by insufficient flood information (multiple responses, n = 40).", width=4.8)
add_table(doc, ["Decision area", "Frequency", "Percentage"], table_rows(decision_counts), widths=[4.55, 1.05, 1.05], numeric_cols=[1, 2])
add_body(
    doc,
    f"A total of {contradictions['none plus affected decisions']} responses selected 'None of the above' together with substantive decision areas. The responses were retained, but the combinations are internally inconsistent."
)

h = add_heading(doc, "8 Location-Specific and Evacuation Information", 1)
h.paragraph_format.page_break_before = True
add_heading(doc, "8.1 Determining Location Susceptibility", 2)
add_body(
    doc,
    f"A total of {susceptibility_challenge} respondents ({pct(susceptibility_challenge, N):.1f}%) found it difficult or very difficult to determine whether their barangay or location might be susceptible to flooding, or stated that they did not know how to determine it. This directly supports the need for accessible location-specific susceptibility information."
)
add_table(doc, ["Difficulty determining susceptibility", "Frequency", "Percentage"], table_rows(susceptibility_counts), widths=[4.25, 1.2, 1.2], numeric_cols=[1, 2])

add_heading(doc, "8.2 Finding Verified Evacuation-Center Information", 2)
add_body(
    doc,
    f"A total of {evac_challenge} respondents ({pct(evac_challenge, N):.1f}%) found verified evacuation-center information difficult or very difficult to locate, or did not know where to find it."
)
add_table(doc, ["Ease of finding verified center information", "Frequency", "Percentage"], table_rows(evac_find_counts), widths=[4.25, 1.2, 1.2], numeric_cols=[1, 2])

add_heading(doc, "8.3 Evacuation-Center Details Considered Useful", 2)
add_chart(doc, chart_evac, "Figure 6. Evacuation-center details considered useful (multiple responses, n = 40).")
add_table(doc, ["Evacuation-center detail", "Frequency", "Percentage"], table_rows(evac_detail_counts), widths=[4.55, 1.05, 1.05], numeric_cols=[1, 2])
add_body(
    doc,
    "Two respondents selected 'does not need evacuation-center information' together with one or more requested details. These inconsistent selections were retained and disclosed."
)

doc.add_page_break()

add_heading(doc, "9 Priority Information Needs", 1)
add_body(
    doc,
    "Respondents selected up to three information types considered most important before possible flooding. The distribution identifies both needs that fit the current FloodSense scope and needs that remain outside it. Current road or travel conditions and real-time or forecast-related information should not be presented as current FloodSense functions unless the research scope and validated data sources are formally expanded."
)
add_chart(doc, chart_priorities, "Figure 7. Information considered most important before possible flooding (up to three responses per participant, n = 40).")
add_table(doc, ["Information need", "Frequency", "Percentage"], table_rows(priority_counts), widths=[4.55, 1.05, 1.05], numeric_cols=[1, 2])

add_heading(doc, "10 Problems with Existing Preparedness Guidance", 1)
add_body(
    doc,
    "The results indicate that respondents' concerns extend beyond obtaining warnings. Many also reported that existing advice is too general, not specific to their location or situation, difficult to understand or access, unclear in action sequence, or unclear in source. These findings support a need for clearer, sourced, and situation-relevant pre-event guidance."
)
add_chart(doc, chart_guidance, "Figure 8. Problems experienced with preparedness guidance from existing sources (multiple responses, n = 40).")
add_table(doc, ["Preparedness-guidance response", "Frequency", "Percentage"], table_rows(guidance_counts), widths=[4.55, 1.05, 1.05], numeric_cols=[1, 2])
add_body(
    doc,
    "One response selected both 'usually receives no preparedness guidance' and 'no problem experienced.' The contradictory combination was retained in both category counts."
)

doc.add_page_break()

add_heading(doc, "11 Technology, Connectivity, and Presentation Requirements", 1)
add_heading(doc, "11.1 Android Access", 2)
add_body(
    doc,
    f"Regular Android access was reported by {android_access} respondents ({pct(android_access, N):.1f}%): {single_count(android_counts, 'Yes, my own device')} had their own device and {single_count(android_counts, 'Yes, a shared device')} used a shared device. One respondent was unsure whether the available device used Android."
)
add_table(doc, ["Android access", "Frequency", "Percentage"], table_rows(android_counts), widths=[4.25, 1.2, 1.2], numeric_cols=[1, 2])

add_heading(doc, "11.2 Internet Reliability During Heavy Rainfall", 2)
add_body(
    doc,
    f"Internet access was described as always or usually reliable by {reliable_internet} respondents ({pct(reliable_internet, N):.1f}%). However, {single_count(internet_counts, 'Sometimes reliable') + single_count(internet_counts, 'Rarely reliable')} ({pct(single_count(internet_counts, 'Sometimes reliable') + single_count(internet_counts, 'Rarely reliable'), N):.1f}%) reported only sometimes or rarely reliable access, indicating a need for efficient data use and clear handling of unavailable information."
)
add_table(doc, ["Internet reliability", "Frequency", "Percentage"], table_rows(internet_counts), widths=[4.25, 1.2, 1.2], numeric_cols=[1, 2])

add_heading(doc, "11.3 Accepted Location-Selection Methods", 2)
add_body(
    doc,
    "Respondents could select more than one method. One-time GPS, a temporary map pin, and manual barangay selection were all accepted by substantial shares of the sample. The results support retaining multiple methods rather than making precise GPS the only way to use location-specific functions."
)
add_table(doc, ["Location-selection response", "Frequency", "Percentage"], table_rows(location_counts), widths=[4.55, 1.05, 1.05], numeric_cols=[1, 2])
add_body(
    doc,
    "One response combined location discomfort with accepted location methods. The system should continue to provide a manual option, a clear purpose explanation, and explicit permission before one-time GPS access."
)

add_heading(doc, "11.4 Language Preference", 2)
add_body(
    doc,
    f"Bilingual English and Filipino presentation was preferred by {single_count(language_counts, 'Both English and Filipino')} respondents ({pct(single_count(language_counts, 'Both English and Filipino'), N):.1f}%)."
)
add_table(doc, ["Preferred presentation", "Frequency", "Percentage"], table_rows(language_counts), widths=[4.25, 1.2, 1.2], numeric_cols=[1, 2])

doc.add_page_break()

add_heading(doc, "12 Optional Open-Ended Responses", 1)
add_body(
    doc,
    f"Twenty-one respondents ({pct(len(comments), N):.1f}%) entered an optional comment. The comments were coded into overlapping themes, so theme percentages can total more than 100%. Names were not used in coding."
)
add_table(doc, ["Theme", "Comments", "Percentage of 21 comments"], [[label, count, f"{pct(count, len(comments)):.1f}%"] for label, count in theme_counts], widths=[4.55, 1.05, 1.05], numeric_cols=[1, 2])
add_body(
    doc,
    "The most common ideas concerned maps or location information, real-time conditions or early warning, water depth or severity, and evacuation-center information. Some of these requests exceed the approved scenario-based scope. They should be recorded as resident needs without being represented as existing FloodSense capabilities."
)
add_body(
    doc,
    "Qualitative caution: six entries used the same long recommendation word for word, two entries consisted of 'none,' and two consisted of 'map.' The repeated long text was counted six times in the frequency table because it appeared in six submitted records, but it represents only one unique wording pattern."
)

add_heading(doc, "13 Integrated Answers to the Research Concerns", 1)
add_heading(doc, "13.1 Is There a Real Flood-Information Problem", 2)
add_body(
    doc,
    f"Yes, among the surveyed respondents. The evidence includes {hard_to_obtain} respondents reporting difficulty or inability to obtain needed information, all eight information-problem indicators falling in the 'Often' mean range, and {decision_often_always} respondents saying inadequate information often or always affected decisions. This supports the existence of an information-access and decision-support problem within the surveyed group."
)

add_heading(doc, "13.2 Is an Additional Tool Needed Despite Existing Sources", 2)
add_body(
    doc,
    "The data support the need for a supplementary tool, not the replacement of official sources. Respondents already use Bacoor City or BDRRMO pages, PAGASA or other government agencies, news, SMS, Facebook groups, personal networks, and barangay channels. Their simultaneous reports of persistent problems indicate that existing channels do not always provide information in the location-specific, explainable, accessible, and actionable form they need."
)

add_heading(doc, "13.3 How FloodSense Can Address the Identified Problems", 2)
mapping_rows = [
    ["General or non-location-specific information", "Barangay selection, temporary map pin, and optional one-time GPS for supported Bacoor locations"],
    ["Difficulty determining local susceptibility", "Scenario-based susceptibility assessment for a selected location"],
    ["Unclear basis, source, or date", "Explainable results with source, status, date, and limitation information"],
    ["Preparedness instructions too general", "Separately sourced pre-event guidance relevant to the assessed scenario"],
    ["Verified evacuation-center information difficult to find", "Verified center references with map location and approximate straight-line distance"],
    ["Internet reliability varies", "Efficient requests, clear unavailable states, and no fabricated result when data cannot be retrieved"],
    ["Different language preferences", "Clear bilingual English and Filipino presentation where feasible"],
]
add_table(doc, ["Survey-supported problem or requirement", "FloodSense response within current scope"], mapping_rows, widths=[2.6, 4.05], font_size=8.8)
add_body(
    doc,
    "This mapping establishes design relevance. It does not prove that the implemented application is already effective. Actual usefulness, usability, clarity, reliability, and decision-support value must be measured after respondents use the working prototype."
)

add_heading(doc, "13.4 What the Initial Survey Does Not Prove", 2)
for item in [
    "It does not prove the scientific accuracy of the susceptibility method or expert rules.",
    "It does not validate BDRRMO, CPDCO, PAGASA, geographic, rainfall, water-level, or evacuation-center datasets.",
    "It does not prove that FloodSense improves preparedness or decision quality after use.",
    "It does not evaluate actual usability, performance, reliability, security, or ISO 25010 characteristics.",
    "It does not authorize claims of live forecasting, automatic warnings, road safety, safe routes, center availability, or evacuation orders.",
]:
    add_bullet(doc, item)

doc.add_page_break()

add_heading(doc, "14 Implications for the Statement of the Problem", 1)
add_body(
    doc,
    "The survey can support the resident-focused part of the Statement of the Problem by establishing current information challenges, their decision effects, and priority requirements. The final Statement of the Problem should retain separate questions for expert validation, system development and functional testing, and post-use quality evaluation because those matters cannot be answered by this initial survey alone."
)
add_heading(doc, "14.1 Evidence-Supported Problem Areas", 2)
for item in [
    "Access to timely, understandable, consistent, and source-transparent flood information",
    "Availability of location-specific flood-susceptibility information",
    "Effects of inadequate information on household preparation, travel, evacuation, work, school, and assistance to vulnerable household members",
    "Availability of clear and situation-relevant pre-event preparedness guidance",
    "Availability of verified evacuation-center references and useful facility details",
    "Feasibility and privacy requirements for Android, connectivity, language, and location selection",
]:
    add_bullet(doc, item)

add_heading(doc, "14.2 Recommended Evidence Chain", 2)
evidence_rows = [
    ["Initial resident survey", "Establishes resident problems, decision effects, needs, and access constraints"],
    ["BDRRMO and CPDCO interviews and records", "Validate institutional knowledge, sources, workflows, geographic data, and operational boundaries"],
    ["Expert and methodology validation", "Validate inference rules, parameters, interpretations, and limitations"],
    ["Functional and technical testing", "Confirm that approved functions work as specified"],
    ["Post-use evaluation", "Measure actual usefulness, usability, clarity, acceptability, and relevant ISO 25010 characteristics"],
]
add_table(doc, ["Evidence source", "Research contribution"], evidence_rows, widths=[2.1, 4.55], font_size=9.0)

h = add_heading(doc, "15 Conclusions", 1)
h.paragraph_format.page_break_before = True
add_body(
    doc,
    "The 40-response initial survey provides coherent descriptive evidence that the surveyed Bacoor residents experience recurring flood and heavy-rainfall impacts and face substantial problems obtaining information that is timely, clear, consistent, accessible, source-transparent, and specific to their location. These problems frequently affect preparedness and everyday decisions."
)
add_body(
    doc,
    "The findings support the need for an additional, resident-facing information tool even though respondents already use official pages, government agencies, news, social media, SMS, barangay channels, and personal networks. The defensible justification for FloodSense is therefore its role as a supplementary, scenario-based, pre-event decision-support tool that organizes approved information into location-specific susceptibility results, explanations, preparedness guidance, and verified evacuation-center references."
)
add_body(
    doc,
    "The survey does not demonstrate that FloodSense has already improved preparedness. That conclusion requires respondents to use the implemented system and complete a separate post-use evaluation. The present findings should be used to define and justify the problem, derive requirements within scope, and guide the later evaluation design."
)

add_heading(doc, "16 Recommendations", 1)
recommendations = [
    "Use the findings to revise the resident-focused SOP and objectives, but phrase claims as applying to the surveyed respondents rather than all Bacoor residents.",
    "Retain FloodSense as a complement to official information channels and maintain explicit limitations that it is not a live forecast, official warning, evacuation order, route-safety service, or live center-availability system.",
    "Prioritize location-specific assessment, explainable results, scenario-relevant preparedness guidance, verified evacuation-center references, privacy-preserving location choices, and bilingual presentation.",
    "Validate scientific inputs, geographic data, institutional workflows, and center records through BDRRMO, CPDCO, relevant source documents, and expert review.",
    "Conduct a separate post-use evaluation with structured tasks and measures for actual usefulness, clarity, usability, reliability, and relevant ISO 25010 characteristics.",
    "For future survey rounds, replace mutually exclusive checkbox combinations with single-choice questions or branching, use a shorter recall period or ask respondents to focus on their most recent event, and use respondent codes if matching initial and post-use responses.",
    "Expand geographic coverage if broader Bacoor conclusions are required, especially in barangays absent or underrepresented in this sample.",
    "Review the timing cluster, exact duplicate response pair, and repeated open-ended wording before treating every submission as an independent observation in formal inferential analysis.",
]
for number, item in enumerate(recommendations, 1):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.25)
    p.paragraph_format.space_after = Pt(4)
    r1 = p.add_run(f"{number}. ")
    style_run(r1, size=10.3, bold=True, color=BLUE)
    r2 = p.add_run(item)
    style_run(r2, size=10.3)

doc.add_page_break()

add_heading(doc, "Appendix A Complete Question-Level Results", 1)
add_body(
    doc,
    "This appendix consolidates the frequency and percentage results used throughout the report. Checkbox percentages are based on all 40 respondents and can total more than 100%."
)

appendix_sections = [
    ("A1 Barangay", barangay_counts),
    ("A2 Age group", age_counts),
    ("A3 Flood and heavy-rainfall effects", impact_counts),
    ("A4 Number of experienced events", event_counts),
    ("A5 Current information sources", source_counts),
    ("A6 Ease of obtaining needed information", access_counts),
    ("A7 Decisions affected", decision_counts),
    ("A8 Difficulty determining location susceptibility", susceptibility_counts),
    ("A9 Ease of finding verified evacuation-center information", evac_find_counts),
    ("A10 Priority information needs", priority_counts),
    ("A11 Problems with preparedness guidance", guidance_counts),
    ("A12 Useful evacuation-center details", evac_detail_counts),
    ("A13 Android access", android_counts),
    ("A14 Internet reliability", internet_counts),
    ("A15 Accepted location-selection methods", location_counts),
    ("A16 Language preference", language_counts),
]
for heading, counts in appendix_sections:
    add_heading(doc, heading, 2)
    add_table(doc, ["Response", "Frequency", "Percentage"], table_rows(counts), widths=[4.55, 1.05, 1.05], numeric_cols=[1, 2], font_size=8.7)

add_heading(doc, "Appendix B Complete Information-Problem Matrix", 1)
for item in problem_stats:
    add_heading(doc, item["label"], 2)
    rows = [[label, item["counts"][label], f"{pct(item['counts'][label], N):.1f}%"] for label in frequency_order]
    rows.append(["Weighted mean excluding Not sure", f"{item['mean']:.2f}", "Often"])
    add_table(doc, ["Response", "Frequency", "Percentage or interpretation"], rows, widths=[4.25, 1.2, 1.2], numeric_cols=[1, 2], font_size=8.7)

add_heading(doc, "Appendix C Source and Analysis Notes", 1)
add_body(
    doc,
    "Source: FloodSense Resident Survey Questionnaire response workbook, worksheet 'Form responses 2,' containing 40 submissions and 32 columns. The workbook was analyzed without modifying the source file."
)
add_body(
    doc,
    "Percentages were calculated as frequency divided by 40 unless another denominator is stated. Values may differ from 100.0% by 0.1 percentage point because of rounding. Open-ended themes overlap and use the 21 nonblank comments as their denominator."
)
add_body(
    doc,
    "No respondent name or initials are reproduced. Optional comments were used only for theme coding and data-quality review."
)

doc.save(DOCX_OUT)

summary = {
    "source": str(SOURCE),
    "responses": N,
    "consented_and_eligible": {
        "consent": consent_count,
        "adult": adult_count,
        "bacoor_resident": resident_count,
        "flood_experience": experience_count,
    },
    "headline": {
        "hard_to_obtain": [hard_to_obtain, pct(hard_to_obtain, N)],
        "decision_often_always": [decision_often_always, pct(decision_often_always, N)],
        "susceptibility_challenge": [susceptibility_challenge, pct(susceptibility_challenge, N)],
        "evacuation_information_challenge": [evac_challenge, pct(evac_challenge, N)],
        "android_access": [android_access, pct(android_access, N)],
        "reliable_internet": [reliable_internet, pct(reliable_internet, N)],
    },
    "barangays": barangay_counts,
    "problem_means": [(item["label"], round(item["mean"], 3), item["often_always"]) for item in problem_stats_ranked],
    "data_quality": {
        "duplicate_group_sizes": duplicate_groups,
        "rapid_intervals_under_10_seconds": rapid_pairs,
        "repeated_comment_counts": sorted(comment_repeats.values(), reverse=True),
        "checkbox_contradictions": contradictions,
    },
    "docx": str(DOCX_OUT),
}
(TMP_DIR / "analysis_summary.json").write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
print(json.dumps(summary, indent=2, ensure_ascii=False))
