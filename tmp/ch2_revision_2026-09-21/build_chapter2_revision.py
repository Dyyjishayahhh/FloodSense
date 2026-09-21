from __future__ import annotations

import hashlib
import json
import shutil
from pathlib import Path

from lxml import etree
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(r"E:\School Downloads-Jiraah\00 FILES IN SCHOOL\3RD YEAR 2ND SEM\DCIT 60\FLOODSENSE\1.FLOODSENSE-200A\FloodSense")
SOURCE = ROOT / "1. FLOODSENSE-MAIN" / "MAIN FILESS" / "FloodSense_Chapters_1_to_3_Chapter_2_Enhanced_2026-09-21.docx"
OUTPUT = ROOT / "1. FLOODSENSE-MAIN" / "MAIN FILESS" / "FloodSense_Chapters_1_to_3_Chapter_2_Applications_and_Narrative_Citations_2026-09-21.docx"
REPORT = ROOT / "tmp" / "ch2_revision_2026-09-21" / "citation_revision_preservation_report.json"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def set_font(run, size=11, bold=False, italic=False):
    run.font.name = "Arial"
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), "Arial")
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), "Arial")
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Arial")
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)


def set_keep_with_next(paragraph, value=True):
    paragraph.paragraph_format.keep_with_next = value


def chapter_heading(anchor, text):
    p = anchor.insert_paragraph_before()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.page_break_before = True
    p.paragraph_format.first_line_indent = Inches(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(12)
    set_keep_with_next(p)
    set_font(p.add_run(text), size=12, bold=True)
    return p


def section_heading(anchor, text):
    p = anchor.insert_paragraph_before()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Inches(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    set_keep_with_next(p)
    set_font(p.add_run(text), size=11, bold=True)
    return p


def topic_heading(anchor, text):
    p = anchor.insert_paragraph_before()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Inches(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(0)
    set_keep_with_next(p)
    set_font(p.add_run(text), size=11, bold=True)
    return p


def body(anchor, text):
    p = anchor.insert_paragraph_before()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Inches(0.5)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    set_font(p.add_run(text), size=11)
    return p


def lead_body(anchor, lead, text):
    p = anchor.insert_paragraph_before()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Inches(0.5)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    set_keep_with_next(p)
    set_font(p.add_run(lead), size=11, bold=True)
    set_font(p.add_run(text), size=11)
    return p


def caption(anchor, text):
    p = anchor.insert_paragraph_before()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Inches(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    set_keep_with_next(p)
    set_font(p.add_run(text), size=10, italic=True)
    return p


def set_cell_margins(cell, top=70, start=80, bottom=70, end=80):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for edge, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        tag = tc_mar.find(qn(f"w:{edge}"))
        if tag is None:
            tag = OxmlElement(f"w:{edge}")
            tc_mar.append(tag)
        tag.set(qn("w:w"), str(value))
        tag.set(qn("w:type"), "dxa")


def shade_cell(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_table_borders(table, color="D9D9D9", size="4"):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        item = borders.find(qn(f"w:{edge}"))
        if item is None:
            item = OxmlElement(f"w:{edge}")
            borders.append(item)
        item.set(qn("w:val"), "single")
        item.set(qn("w:sz"), size)
        item.set(qn("w:space"), "0")
        item.set(qn("w:color"), color)


def add_comparison_table(doc, anchor, headers, widths, rows):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.style = "Table Grid"
    anchor._p.addprevious(table._tbl)
    set_table_borders(table)

    header = table.rows[0]
    tr_pr = header._tr.get_or_add_trPr()
    repeat = OxmlElement("w:tblHeader")
    repeat.set(qn("w:val"), "true")
    tr_pr.append(repeat)
    no_split = OxmlElement("w:cantSplit")
    tr_pr.append(no_split)

    for index, (cell, text) in enumerate(zip(header.cells, headers)):
        cell.width = widths[index]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_margins(cell)
        shade_cell(cell, "E7E6E6")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Inches(0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.paragraph_format.space_after = Pt(0)
        set_font(p.add_run(text), size=8, bold=True)

    for row_index, values in enumerate(rows):
        row = table.add_row()
        tr_pr = row._tr.get_or_add_trPr()
        no_split = OxmlElement("w:cantSplit")
        tr_pr.append(no_split)
        for index, (cell, text) in enumerate(zip(row.cells, values)):
            cell.width = widths[index]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            if row_index % 2:
                shade_cell(cell, "F7F7F7")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Inches(0)
            p.paragraph_format.line_spacing = 1.05
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            set_font(p.add_run(text), size=8)
    return table


def reference_format(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.left_indent = Inches(0.5)
    paragraph.paragraph_format.first_line_indent = Inches(-0.5)
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(6)
    for run in paragraph.runs:
        set_font(run, size=11)


def add_reference_before(anchor, text):
    p = anchor.insert_paragraph_before()
    set_font(p.add_run(text), size=11)
    reference_format(p)
    return p


def c14n_slice(document, start_text=None, end_text=None):
    body_el = document._body._element
    children = list(body_el)
    start_index = 0
    end_index = len(children)
    if start_text is not None:
        for idx, child in enumerate(children):
            text = "".join(node.text or "" for node in child.iter(qn("w:t")))
            if child.tag == qn("w:p") and text.strip() == start_text:
                start_index = idx
                break
    if end_text is not None:
        for idx, child in enumerate(children):
            text = "".join(node.text or "" for node in child.iter(qn("w:t")))
            if idx >= start_index and child.tag == qn("w:p") and text.strip() == end_text:
                end_index = idx
                break
    payload = b"".join(etree.tostring(child, method="c14n") for child in children[start_index:end_index])
    return hashlib.sha256(payload).hexdigest().upper()


if not SOURCE.exists():
    raise FileNotFoundError(SOURCE)

source_doc = Document(SOURCE)
before_ch1 = c14n_slice(source_doc, end_text="REVIEW OF RELATED LITERATURE")
before_ch3 = c14n_slice(source_doc, start_text="METHODOLOGY", end_text="REFERENCES")
before_sections = len(source_doc.sections)
before_tables = len(source_doc.tables)
before_shapes = len(source_doc.inline_shapes)

shutil.copy2(SOURCE, OUTPUT)
doc = Document(OUTPUT)

start_p = next(p for p in doc.paragraphs if p.text.strip() == "REVIEW OF RELATED LITERATURE")
method_p = next(p for p in doc.paragraphs if p.text.strip() == "METHODOLOGY")
body_el = doc._body._element
children = list(body_el)
start_index = children.index(start_p._p)
end_index = children.index(method_p._p)
for child in children[start_index:end_index]:
    body_el.remove(child)

chapter_heading(method_p, "REVIEW OF RELATED LITERATURE")
body(method_p, "This chapter reviews literature and studies that clarify the concepts, technologies, responsibilities, and limitations relevant to FloodSense. The discussion is organized into foreign and local literature followed by foreign and local studies. It examines flood susceptibility, scenario-based assessment, deterministic Rule-Based Expert Systems, explainable inference, Decision Support Systems, geographic information, spatial databases, location privacy, risk communication, data provenance, administrative governance, software quality, and Philippine flood conditions.")
body(method_p, "The reviewed materials are used to explain the research problem and the design boundaries of FloodSense. Features described in another application or study are not treated as FloodSense functions. FloodSense remains a scenario-based, pre-event susceptibility-assessment and preparedness-support system in which the user confirms a hypothetical rainfall scenario and explicitly requests an assessment. It is distinct from operational forecasting, real-time monitoring, official warning, and automatic evacuation systems.")

section_heading(method_p, "Related Foreign Literature")

topic_heading(method_p, "Flood Susceptibility and Scenario-Based Assessment")
body(method_p, "Flood susceptibility expresses the relative tendency of a location to experience flooding when relevant conditions are present. It differs from flood hazard, which concerns the probability and physical characteristics of an event; exposure, which concerns people or assets in affected places; vulnerability, which concerns the capacity to be harmed; and risk, which combines hazard and potential consequences. Membele, Naidu, and Mutanga (2022) found that flood-vulnerability mapping in developing countries varies substantially in terminology, indicators, spatial scale, data availability, and validation practice. Their review shows why a public system must name the type of assessment it performs and avoid presenting a general vulnerability or hazard map as a locally validated susceptibility conclusion.")
body(method_p, "Scenario-based assessment supports planning by holding selected conditions explicit instead of implying that the system is observing an event as it occurs. A hypothetical rainfall intensity and duration can be treated as assessment inputs only when their meanings, permitted combinations, and geographic applicability are documented. This supports the FloodSense interaction in which the resident selects or confirms a scenario before requesting assessment. It does not establish that the scenario is current, predicted, or officially issued, and it does not justify importing weights or thresholds derived for another place.")

topic_heading(method_p, "Deterministic Rule-Based Expert Systems and Knowledge Bases")
body(method_p, "A Rule-Based Expert System represents domain knowledge through facts, conditions, and production rules and applies an inference procedure to derive a conclusion. Buchanan and Shortliffe (1984) provide the foundational distinction between the knowledge base and the inference mechanism. More recently, Papadopoulos et al. (2022) reviewed technologies and standards used in rule-based decision-support systems and described how encoded knowledge and rules engines evaluate case-specific facts. Although their review concerns clinical systems, the architectural distinction is relevant to FloodSense: flood knowledge, scenario facts, geographic facts, and explanations belong to controlled knowledge structures, while the inference procedure determines which eligible rule applies.")
body(method_p, "Deterministic processing means that the same approved facts and the same published rule version should produce the same result. It does not imply that the rules are scientifically valid merely because the software executes them consistently. Knowledge acquisition, conflict handling, parameter definitions, version control, and expert or institutional validation remain necessary. Accordingly, ordinary FloodSense administrators may manage approved records and permitted values but do not rewrite raw rule conditions, priorities, conflict-resolution logic, or algorithm code. The confirmed fixed-schema CSV exchange is a future controlled-value workflow and is not evidence that such exchange is currently operational.")

topic_heading(method_p, "Explainable Inference")
body(method_p, "Explainability is necessary when a system produces a classification that may influence preparedness decisions. Kostopoulos, Davrazos, and Kotsiantis (2024) reviewed explainable decision-support approaches and emphasized transparency, interpretability, and human-understandable justification. Their review primarily addresses artificial-intelligence-based systems, but its communication principles apply to deterministic inference: a user should be able to understand the relevant inputs, the basis of the result, and the limitations of the evidence without being required to inspect source code or raw rule definitions.")
body(method_p, "For FloodSense, an explanation should identify the confirmed rainfall scenario, resolved geographic context, susceptibility class, applicable basis, source or version information, and any limitation state. This supports accountability while preserving the distinction between explanation and rule administration. It also prevents a color or label from appearing authoritative without disclosing what it represents. Explainability does not remove the need for validated data, and a clear explanation cannot turn provisional information into an official local assessment.")

topic_heading(method_p, "Decision Support and Preparedness Guidance")
body(method_p, "A Decision Support System organizes information, alternatives, or recommended actions to assist human judgment. In flood preparedness, support may include understandable mitigation choices, household preparation, communication planning, protection of documents and valuables, assistance for vulnerable household members, attention to official advisories, and evacuation readiness. Mostafiz et al. (2022) observed that many web-based flood tools provide technical information but do not consistently provide locally actionable information for individuals and communities. Their review supports combining understandable risk communication with practical guidance while clearly stating the tool's authority and limitations.")
body(method_p, "FloodSense separates the susceptibility assessment from preparedness guidance. The deterministic Expert System produces the assessment, while the DSS presents source-backed pre-event guidance without changing the susceptibility class. Information about evacuation facilities may support readiness when its source, verification date, location, and operational limitations are known, but distance alone does not prove route safety, current opening, accessibility, or available capacity. The DSS therefore assists preparation and does not issue an official warning, evacuation order, or guarantee of safety.")

topic_heading(method_p, "GIS, Interactive Maps, and Spatial Databases")
body(method_p, "Geographic Information Systems organize, analyze, and present information associated with locations. Their value in flood research depends on the suitability of the layers, coordinate reference systems, spatial resolution, classification methods, and validation procedures. An administrative boundary can resolve the name of a barangay or provide map context, but it contains no flood-susceptibility fact unless appropriate hazard or susceptibility evidence has been associated with it. This distinction is important when users may interpret any colored polygon as an official hazard layer.")
body(method_p, "Spatial databases support the controlled storage and querying of geographic records. The PostGIS Development Group (2026) explains that PostGIS extends PostgreSQL with geometry and geography types, spatial indexes, spatial relationships, coordinate transformation, validity operations, and spatial processing. These capabilities are relevant to point-in-polygon queries, administrative-area lookup, map-feature storage, and provenance-linked geographic records. They provide technical infrastructure, not scientific validation: a valid geometry and a successful spatial query do not establish that a flood conclusion is institutionally approved.")

topic_heading(method_p, "Flood-Risk Communication")
body(method_p, "Flood-risk information must be understandable, actionable, and explicit about uncertainty. Mostafiz et al. (2022) identified gaps in web-based flood communication, particularly where tools provide information without sufficiently connecting it to decisions that residents can make. Clear terms, legends, labels, explanation text, and source notices help prevent non-specialists from treating model outputs as certain predictions. Communication should also distinguish long-term or scenario-based information from time-sensitive official warnings.")
body(method_p, "These principles support a guided FloodSense flow rather than an unexplained technical map. Susceptibility colors must be paired with text, the selected scenario, and limitations. Preparedness information must direct users toward official channels during actual events. The interface must not imply that a neutral boundary layer, a demonstration color, or a provisional record represents a current flood condition.")

topic_heading(method_p, "Data Provenance and Geographic Validation")
body(method_p, "Provenance records where information originated, how it was transformed, which version was used, and what review or approval applies. Tullis and Kar (2021) connected provenance to ethical replicability and reproducibility in GIScience, including critical applications in which geographic data affect public decisions. Fakhruddin et al. (2022) similarly emphasized risk-informed data practices and the importance of understanding how disaster and climate data are produced, shared, interpreted, and governed.")
body(method_p, "For FloodSense, provenance should accompany susceptibility parameters, geographic layers, preparedness items, and resource information. Technical geometry validation may confirm coordinate-system consistency, valid shapes, or feature counts, but institutional validation concerns authority, currency, meaning, and permitted public use. The two forms of validation are related but not interchangeable. This distinction prevents a technically processed dataset from being represented as official Bacoor City flood information without documented approval.")

topic_heading(method_p, "Foreground Location and Privacy")
body(method_p, "Mobile location can reveal a person's position and may contribute to identification or tracking. Android Developers (2026) distinguishes foreground access, one-time access, and background access and recommends requesting only the permission needed for the active function. This supports a FloodSense flow in which location is obtained only after a user initiates location selection or assessment, while manual selection remains available when the resident declines permission or prefers another area.")
body(method_p, "Foreground acquisition does not imply continuous monitoring. FloodSense does not require background GPS tracking, automatic polling, or autonomous location-based alerts. Location handling should remain purpose-limited to the active assessment, use only the precision necessary for the supported geographic operation, and avoid retaining unnecessary movement information. The presentation must also allow the user to confirm or correct the detected barangay because device accuracy and boundary proximity can affect geographic resolution.")

topic_heading(method_p, "Software Testing and Quality Evaluation")
body(method_p, "Software testing establishes whether specified behavior is implemented under defined conditions, while quality evaluation examines broader product characteristics. ISO/IEC 25010:2023 defines a product-quality model with nine characteristics for specifying and evaluating information and communication technology products (International Organization for Standardization [ISO], 2023). The model can guide the selection of criteria for FloodSense, but it does not provide a ready-made acceptance result or replace an approved evaluation instrument.")
body(method_p, "A defensible evaluation must identify the tested build, participants or evaluators, tasks, criteria, scale, procedure, valid responses, and interpretation method. Repository tests and interface demonstrations can support implementation verification, but they are not resident-acceptance findings. Because the final FloodSense survey and evaluation are not complete, Chapter 2 uses ISO/IEC 25010 only as literature supporting future assessment and does not report quality scores.")

section_heading(method_p, "Related Local Literature")

topic_heading(method_p, "Philippine Flood Conditions and Institutional Preparedness")
body(method_p, "Official Philippine statistics continue to record substantial effects from extreme events and disasters. The Philippine Statistics Authority (PSA, 2026) summarizes national information on disaster occurrence, affected populations, deaths, and economic losses. These national figures establish the continuing relevance of disaster preparedness but do not identify the susceptibility of a particular Bacoor barangay or validate the inputs of a local application.")
body(method_p, "The National Disaster Risk Reduction and Management Council's National Disaster Response Plan identifies institutional responsibilities and coordinated actions for hydro-meteorological hazards (NDRRMC, 2024). The plan reinforces that warnings, response decisions, and emergency coordination belong to authorized government mechanisms. FloodSense may support pre-event understanding and preparedness, but residents must continue to follow PAGASA, local disaster offices, barangay authorities, and emergency responders during an actual event.")

topic_heading(method_p, "Official Hazard Information and Warning Authority")
body(method_p, "PAGASA operates rainfall and flood-information processes that use observations, forecasts, water-level information, and coordinated warning protocols. Its 2023 Annual Climate Bulletin explains the role of rainfall warning systems and the use of official warning levels for public action (PAGASA, 2024). FloodSense rainfall selections are hypothetical assessment inputs; they are not automatic updates from PAGASA, an interpretation of current rainfall, or an official warning product.")
body(method_p, "GeoRisk Philippines presents HazardHunterPH as a national platform for indicative hazard assessment while advising users that results depend on available government data and do not replace official hazard reports from mandated agencies (GeoRisk Philippines, n.d.). Its source notices and limitations demonstrate responsible hazard communication. The same principle applies to FloodSense: sources, dates, limitations, and validation status should remain visible, and an indicative or provisional record should not be represented as an absolute prediction.")

topic_heading(method_p, "Cavite, Bacoor City, and the Imus River Basin")
body(method_p, "The Japan International Cooperation Agency reported the inauguration of flood-mitigation facilities associated with the Imus River Basin and low-lying areas of Imus and Bacoor (JICA, 2021). This institutional record confirms the relevance of watershed and infrastructure conditions to the local context. It does not provide barangay susceptibility classifications, rainfall-rule thresholds, or validation for FloodSense, and it should not be treated as a substitute for hydrologic or hazard data.")
body(method_p, "Administrative geography in Bacoor has also changed. The PSA documented the 2023 merger and renaming of barangays and currently lists 47 barangays in the City of Bacoor (PSA, 2023, 2025). These records support the current identities used in a reference layer. They do not establish the accuracy of every derived polygon, City verification of a third-party boundary, or flood conditions within any barangay. Administrative reference and flood-susceptibility evidence must therefore remain separate.")

topic_heading(method_p, "Administrative Data Governance and Responsible Use")
body(method_p, "Government publications, local records, and research datasets differ in authority and permitted use. An official list of administrative units may verify names and codes, while an authorized hazard agency may provide hazard information, and a local office may validate operational records. Combining these sources requires documentation of ownership, date, scale, processing, approval, and restrictions. Provenance is particularly important when a dataset has been transformed, because technical processing can change geometry or attributes without changing the institution responsible for the original information.")
body(method_p, "FloodSense consequently distinguishes approved, provisional, technically validated, and unsupported information. The 47-barangay layer may serve as a current administrative reference after documented derivation and geometry checks, but it is not a flood-susceptibility layer and is not automatically City-issued or City-verified. No operational susceptibility dataset should be described as established until the necessary source, method, expert review, institutional authority, and database integration are documented.")

topic_heading(method_p, "Philippine Data Privacy and Software-Quality Context")
body(method_p, "Republic Act No. 10173, the Data Privacy Act of 2012, requires personal-information processing to observe transparency, legitimate purpose, and proportionality. For a resident-facing location function, these principles support clear notice, user initiation, limited collection, secure handling, and avoidance of continuous background tracking. They also support role-controlled administrative access. The Act is legally authoritative despite falling outside the preferred publication period for recent literature.")
body(method_p, "Philippine software studies have also used ISO/IEC 25010 as an evaluation framework, but the criteria and edition must be identified accurately. An evaluation of another application cannot be transferred to FloodSense, and the presence of a questionnaire does not establish completed results. FloodSense quality conclusions will require its own approved instrument, respondents or evaluators, documented procedure, and completed analysis.")

section_heading(method_p, "Related Foreign Studies")

topic_heading(method_p, "FloodAdapt")
body(method_p, "Deltares (2025) developed FloodAdapt as an open-source decision-support application for exploring flood risk under alternative present and future conditions. The application links the SFINCS flood model with Delft-FIAT impact assessment and allows users to define what-if events, future scenarios, and adaptation measures. Its outputs include mapped flooding and impacts, summary metrics, and cost-benefit information intended for planners, analysts, and decision makers rather than for automatic public warning.")
body(method_p, "FloodAdapt shows how explicit scenarios can be compared without presenting them as observations of a current event. Its physics-based modeling, impact calculations, and adaptation analysis require specialized data and expertise, and the downloadable release does not supply Bacoor-specific inputs or institutional approval. The model configuration, damage functions, and scenario assumptions therefore cannot be transferred directly to Bacoor City.")
body(method_p, "FloodAdapt is similar to FloodSense in asking users to consider a defined scenario and in presenting mapped decision information. It differs in audience, scale, scientific method, and output: FloodSense uses a resident-confirmed rainfall scenario and deterministic rules to produce a bounded susceptibility class followed by separate preparedness guidance. FloodSense does not perform hydrodynamic simulation, damage estimation, adaptation comparison, or cost-benefit analysis.")

topic_heading(method_p, "Global Flood Awareness System")
body(method_p, "The Copernicus Emergency Management Service (2026) operates the Global Flood Awareness System, or GloFAS, as a global flood-forecasting and monitoring service. Its web interface provides map layers, forecast products, hydrological information, and data access intended to support preparedness and emergency-management activity across transnational river basins. The service draws on operational forecasting infrastructure and is maintained within the Copernicus Emergency Management Service and the European Centre for Medium-Range Weather Forecasts.")
body(method_p, "GloFAS demonstrates the institutional, computational, and data requirements of an operational forecast service. Global model resolution, forecast uncertainty, local river behavior, and the need for national or local interpretation limit direct use for household-level decisions. Its products also do not constitute Bacoor City validation and cannot be converted into FloodSense rules without an approved methodology, local evidence, and appropriate institutional authority.")
body(method_p, "Both GloFAS and FloodSense use maps and flood-related categories to support understanding, but their functions are fundamentally different. GloFAS is an operational monitoring and forecasting service; FloodSense is a user-requested, scenario-based pre-event assessment. FloodSense does not ingest live GloFAS forecasts, update rainfall automatically, monitor rivers, or issue official warnings.")

topic_heading(method_p, "Iowa Flood Information System")
body(method_p, "The Iowa Flood Center (2024) describes the Iowa Flood Information System, or IFIS, as a publicly accessible platform that integrates a statewide sensor network, weather and hydrologic information, flood-inundation maps, and real-time forecasts. The system has evolved since its 2011 launch and supports residents, emergency managers, researchers, and other users through an interactive geographic interface. Its operation depends on Iowa-specific sensors, models, data streams, and institutional maintenance.")
body(method_p, "The Iowa Flood Center (n.d.) also states that the platform is under active development and that its maps and information are provisional and are not a substitute for official regulatory determinations. This qualification is important because technically sophisticated public maps can still have geographic, predictive, and legal limitations. IFIS does not provide an approved susceptibility dataset, rule base, or evaluation result for Bacoor City.")
body(method_p, "IFIS resembles FloodSense in offering location-oriented map information and understandable public presentation. It differs because IFIS uses real-time sensor and forecasting infrastructure, whereas FloodSense requires a resident to confirm a hypothetical scenario and explicitly request deterministic assessment. FloodSense has no statewide sensor network, continuous monitoring, autonomous alerting, or real-time forecast function.")

topic_heading(method_p, "Resilience Analysis and Planning Tool")
body(method_p, "The Federal Emergency Management Agency (FEMA, 2025) maintains the Resilience Analysis and Planning Tool, or RAPT, as a public geographic application for community resilience planning. The tool brings together more than one hundred layers concerning population, infrastructure, hazards, and community conditions and supports visualization, filtering, comparison, and export for planning, mitigation, response, and recovery activities in the United States.")
body(method_p, "RAPT illustrates the value of bringing contextual data and provenance into one map, but it is not limited to floods and does not convert every displayed layer into a location-specific hazard conclusion. Its United States datasets, administrative units, definitions, and legal context are not applicable to Bacoor without local substitution and validation. Layer availability also does not guarantee suitability for a particular decision.")
body(method_p, "RAPT and FloodSense both use geographic interfaces to organize information for decision support. FloodSense differs by applying a protected deterministic inference process to a resident-confirmed rainfall scenario and then presenting separate preparedness guidance. FloodSense does not reproduce RAPT's national data catalog, demographic analysis, resilience indicators, or export capabilities.")

topic_heading(method_p, "Check the Long Term Flood Risk for an Area in England")
body(method_p, "The Environment Agency (2024) supports England's national assessment of flood and coastal erosion risk, while the public Check the Long Term Flood Risk for an Area service allows a user to search a place and review long-term risk from rivers and the sea, surface water, reservoirs, and groundwater. The service also provides contextual information on climate change and links users to planning and preparedness resources through an official government channel.")
body(method_p, "The service explicitly distinguishes area-level information from the risk of an individual property. That limitation reflects differences in spatial scale, source datasets, update cycles, and the uncertainty of local conditions. Its English hazard models, categories, responsibilities, and official status cannot be generalized to Bacoor City or used to validate FloodSense.")
body(method_p, "The location-search interaction and careful limitation statements are relevant to FloodSense's map and communication design. FloodSense, however, is not a national long-term-risk authority and does not cover multiple hazard mechanisms or property-level risk. It presents a bounded hypothetical-scenario susceptibility assessment and directs users to official agencies for current forecasts, warnings, and emergency instructions.")

section_heading(method_p, "Related Local Studies")

topic_heading(method_p, "HazardHunterPH")
body(method_p, "GeoRisk Philippines (n.d.) presents HazardHunterPH as a public web application that generates indicative multi-hazard assessment reports for a user-selected Philippine location. The platform combines hazard information from mandated government agencies and communicates the available layers, source institutions, and assessment limitations. It is intended to improve public access to location-based hazard information rather than to replace the official reports issued by responsible agencies.")
body(method_p, "HazardHunterPH's own guidance states that the report is indicative and should not be used as an official or legal document. Available results depend on the datasets supplied to the platform, and a location report does not eliminate the need for agency confirmation. These limitations are directly relevant to Bacoor because neither a national platform nor an administrative reference layer automatically establishes FloodSense's local susceptibility parameters.")
body(method_p, "HazardHunterPH and FloodSense both use a selected location and a map to communicate hazard-related information. HazardHunterPH retrieves available government hazard layers across multiple hazards, whereas FloodSense applies deterministic rules to a user-confirmed hypothetical rainfall scenario and presents separate preparedness guidance. FloodSense must preserve HazardHunterPH's lesson on visible sources and limitations without implying that the two systems have the same data, authority, or output.")

topic_heading(method_p, "GeoMapperPH")
body(method_p, "The Department of Science and Technology-Philippine Institute of Volcanology and Seismology (DOST-PHIVOLCS, 2022a) describes GeoMapperPH as a GeoRisk Philippines application used by trained local-government and agency personnel to collect information on hazards, exposure, vulnerability, and coping capacity. The platform supports standardized field data collection and contributes records to exposure databases that can be used in risk assessment and planning. Its documented training context shows that data collection is governed by authorized users, defined procedures, and institutional coordination.")
body(method_p, "GeoMapperPH is not presented as an unrestricted resident assessment application. The usefulness of its records depends on collection quality, classification consistency, source responsibility, access controls, and continued updating. Training materials or the existence of a collection application do not prove that its datasets have been approved for FloodSense or imported into the FloodSense database.")
body(method_p, "The application is relevant to FloodSense because it illustrates provenance, standardized schemas, permissions, and administrative data governance. FloodSense differs in serving a resident scenario-assessment flow and in protecting its deterministic rule logic from ordinary administrators. The confirmed fixed-schema CSV workflow would govern permitted values rather than reproduce GeoMapperPH's field-data collection functions, and that workflow is not yet operational.")

topic_heading(method_p, "PlanSmartPH Ready to Rebuild")
body(method_p, "DOST-PHIVOLCS (2022b) describes PlanSmartPH Ready to Rebuild as a web application developed with national agencies and development partners to assist local governments in preparing rehabilitation and recovery plans. The platform organizes baseline data, damage and loss information, recovery needs, programs, and planning outputs so that local officials can develop structured post-disaster plans more efficiently.")
body(method_p, "PlanSmartPH operates in an institutional planning context and depends on the completeness and validation of local-government data. Its recovery-planning workflow, participating agencies, datasets, and user roles differ from a resident-facing pre-event susceptibility tool. It does not provide approved Bacoor rainfall-rule parameters or establish that FloodSense can make official recovery, response, or evacuation decisions.")
body(method_p, "Both applications demonstrate the value of structured data and decision support, but they support different phases of disaster management. PlanSmartPH assists post-disaster rehabilitation and recovery planning, while FloodSense supports pre-event scenario understanding and household preparedness. FloodSense therefore does not inherit PlanSmartPH's planning authority, damage assessment, financial estimates, or institutional outputs.")

topic_heading(method_p, "UP NOAH Know Your Hazards and NOAH Studio")
body(method_p, "The UP NOAH Center (n.d.) maintains public location-based interfaces through Know Your Hazards and NOAH Studio. These platforms allow users to search places, view hazard layers, and examine flood scenarios such as the five-, twenty-five-, and one-hundred-year layers derived from simulations, satellite information, and historical or geographic datasets. Their public maps support awareness and exploration across Philippine locations.")
body(method_p, "UP NOAH also publishes disclaimers on the interpretation and authority of its layers, including the limits of administrative boundaries and the need to consult appropriate agencies. Return-period flood scenarios are not the same as the rainfall-intensity and duration selections used by FloodSense. The platform's layers, simulations, and classifications cannot be treated as imported or approved FloodSense data unless provenance, permission, methodology, validation, and integration are separately established.")
body(method_p, "The UP NOAH interfaces are similar to FloodSense in their use of interactive maps and location-based flood communication. They differ in scenario definition, data production, geographic coverage, and institutional purpose. FloodSense does not claim to reproduce NOAH simulations, predict inundation extent, or provide the same nationwide hazard products.")

topic_heading(method_p, "Web-Based Solution for Flood Warning Decision Support in Leyte")
body(method_p, "Bentoso et al. (2021) developed and tested a web-based flood-warning decision-support prototype for the Province of Leyte. The researchers used rainfall and water-level sensors, data communication, a web dashboard, graphical readings, warning thresholds, and notification functions to help responsible users monitor conditions in flood-prone areas. The conference paper describes a prototype and its functional testing; it does not establish province-wide operational deployment.")
body(method_p, "The system's sensor locations, communication infrastructure, rainfall and water-level thresholds, and warning procedures were designed for its own context. Those elements cannot be transferred to Bacoor without a new scientific and institutional basis. A working prototype also does not prove the accuracy or authority of warnings outside the study setting.")
body(method_p, "The Leyte application and FloodSense both organize flood-related information through a digital interface and aim to support preparedness decisions. The difference is decisive: the Leyte prototype monitors sensors and generates warning information, while FloodSense has no continuous monitoring, background polling, autonomous alerts, or official warning authority. FloodSense remains a user-requested hypothetical scenario assessment with separate guidance.")

section_heading(method_p, "Synthesis of the Reviewed Literature")
body(method_p, "The literature agrees that flood-related classifications depend on the meaning, quality, scale, and validation of their inputs. Susceptibility maps, hydraulic models, exposure estimates, and social-vulnerability indices answer different questions. They cannot be treated as interchangeable merely because all use maps or ordered categories. For FloodSense, this requires a precise description of scenario-based susceptibility and prevents administrative boundaries, national exposure products, social indices, or another basin's model from being represented as local flood truth.")
body(method_p, "The reviewed Rule-Based Expert System and explainability literature supports the separation of a controlled knowledge base from its inference procedure. Deterministic rules can produce traceable results when facts, conditions, priorities, versions, and explanations are documented. However, deterministic execution is not a substitute for expert and institutional validation. Protecting raw rule logic from ordinary administration and limiting future CSV exchange to permitted values preserve the difference between governed data maintenance and rewriting the assessment algorithm.")
body(method_p, "Decision-support and risk-communication research further shows that a classification alone is insufficient for public use. Users require understandable labels, limitations, sources, and actionable preparedness information. FloodSense therefore separates susceptibility assessment from DSS guidance. The guidance may assist household preparation and encourage attention to official channels, but it does not change the assessment, predict current conditions, issue a warning, or order evacuation.")
body(method_p, "GIS and spatial-database technologies support location selection, containment queries, storage, map visualization, and geographic processing. The Philippine studies demonstrate that similar technologies can support national exposure assessment, basin modeling, municipal risk zoning, community research, and warning prototypes. Their diversity also demonstrates why external formulas, weights, thresholds, and system functions cannot be transferred automatically. Each method reflects a different geographic scale, dataset, user group, and institutional purpose.")
body(method_p, "The local literature and application evidence emphasize the authority of PAGASA, NDRRMC, DOST-PHIVOLCS, GeoRisk Philippines, PSA, UP NOAH, local disaster offices, and other custodians. HazardHunterPH, GeoMapperPH, PlanSmartPH, and UP NOAH demonstrate distinct functions: public indicative assessment, governed data collection, institutional recovery planning, and hazard visualization. Their differences show why an official institution's participation in one platform does not authorize FloodSense to claim its data, functions, or authority. Bacoor's current 47-barangay identity structure remains an administrative matter, not a flood-susceptibility dataset.")
body(method_p, "Finally, the software-quality literature supports systematic testing and evaluation but does not permit results to be claimed before data collection and analysis are complete. Automated tests can demonstrate specified implementation behavior, while an approved ISO/IEC 25010-based evaluation can examine broader product quality. FloodSense must report each type of evidence according to what was actually tested and must not treat an empty survey package as completed evaluation.")

section_heading(method_p, "Research Gap")
body(method_p, "Existing applications already provide scenario-based flood modeling, global forecasting, sensor-supported monitoring, national hazard reports, geographic resilience planning, long-term risk lookup, exposure-data collection, recovery planning, and operational warning prototypes. These contributions establish mature methods and design principles that FloodSense can learn from. They also demonstrate that real-time warning and forecasting require observations, calibrated models, thresholds, communications infrastructure, and institutional authority that are outside the current FloodSense scope.")
body(method_p, "The reviewed literature provides limited evidence for a Bacoor-specific public interaction that combines a user-confirmed hypothetical rainfall scenario, a supported location, deterministic and explainable rule matching, and separate pre-event preparedness guidance while making provenance and authority limits visible. This is not a claim that such a combination has never existed. It identifies an insufficiently documented intersection within the sources reviewed, especially for Bacoor City and its current barangay structure.")
body(method_p, "A more significant gap concerns local validation. The reviewed Philippine studies were conducted at national scale or in Metro Manila, Davao, Romblon, Leyte, and other basins. Even the Cavite institutional material does not provide approved FloodSense susceptibility parameters. Their findings clarify methods, communication needs, and governance problems but cannot establish Bacoor classifications. Local data, expert review, source permission, institutional validation, and transparent versioning remain necessary before FloodSense can be represented as an operational public assessment.")
body(method_p, "FloodSense addresses the documented design gap as a bounded research prototype: the resident confirms a hypothetical scenario and requests an assessment; the deterministic Expert System produces an explainable susceptibility result from approved inputs; and the separate DSS presents source-backed preparedness guidance. FloodSense remains different from PAGASA products, official hazard reports, real-time warning platforms, monitoring networks, and evacuation authorities. Its value depends on maintaining that distinction and reporting unsupported data or functions as limitations rather than capabilities.")

section_heading(method_p, "Comparative Synthesis of Related Studies")
body(method_p, "The following tables summarize the principal relationships among the reviewed studies. The comparisons identify method, contribution, limitation, and relevance without treating any external dataset, formula, or feature as part of FloodSense.")

table1_caption = caption(method_p, "Table 1. Comparison of related foreign applications and systems")
table1_caption.paragraph_format.page_break_before = True
add_comparison_table(
    doc,
    method_p,
    ["APPLICATION", "DEVELOPER\nYEAR", "USERS\nSCOPE", "MAIN\nFUNCTIONS", "LIMITATION", "FLOODSENSE\nCOMPARISON"],
    [Inches(1.20), Inches(0.90), Inches(0.90), Inches(0.90), Inches(0.90), Inches(0.95)],
    [
        ("FloodAdapt", "Deltares, 2025", "Planners; scenario analysis", "Flood and impact modeling; adaptation comparison", "Needs expert, place-specific inputs", "Both use scenarios; FloodSense has no hydrodynamic or cost-benefit model"),
        ("GloFAS", "Copernicus EMS, 2026", "Global forecasting and preparedness", "Operational forecasts, map layers, data access", "Global scale needs local interpretation", "FloodSense is not monitoring or forecasting"),
        ("IFIS", "Iowa Flood Center, 2024", "Iowa public and emergency users", "Sensors, inundation maps, real-time forecasts", "Provisional; Iowa-specific infrastructure", "Both use public maps; FloodSense has no live sensors or alerts"),
        ("RAPT", "FEMA, 2025", "United States resilience planning", "Over 100 contextual and hazard layers", "Not flood-only; U.S.-specific data", "Both support map-based decisions; FloodSense adds bounded rule inference"),
        ("Check Long Term Flood Risk", "Environment Agency, 2024", "Public area-level risk lookup", "Multiple flood sources and preparedness links", "Not individual-property risk", "Both use location search; FloodSense uses a hypothetical rainfall scenario"),
    ],
)

table2_caption = caption(method_p, "Table 2. Comparison of related local applications and systems")
table2_caption.paragraph_format.page_break_before = True
local_comparison_rows = [
        ("HazardHunterPH", "GeoRisk Philippines, n.d.", "Public Philippine location reports", "Indicative multi-hazard assessment", "Not an official or legal report", "Both use location; different data, authority, and inference"),
        ("GeoMapperPH", "DOST-PHIVOLCS, 2022", "Trained LGU and agency personnel", "Governed exposure and capacity data collection", "Quality depends on authorized collection", "Supports provenance and schemas; not resident assessment"),
        ("PlanSmartPH", "DOST-PHIVOLCS et al., 2022", "LGU recovery planning", "Baseline, loss, needs, and recovery plans", "Institutional post-disaster scope", "Both support decisions; different phase and authority"),
        ("UP NOAH platforms", "UP NOAH Center, n.d.", "Public nationwide hazard viewing", "Location search and simulated flood layers", "Layer and boundary disclaimers", "Both use maps; FloodSense does not reproduce NOAH models"),
        ("Leyte warning prototype", "Bentoso et al., 2021", "Monitoring users in Leyte", "Sensors, dashboard, thresholds, alerts", "Operational deployment not established", "FloodSense has no sensors, live monitoring, or alerts"),
]
add_comparison_table(
    doc,
    method_p,
    ["APPLICATION", "DEVELOPER\nYEAR", "USERS\nSCOPE", "MAIN\nFUNCTIONS", "LIMITATION", "FLOODSENSE\nCOMPARISON"],
    [Inches(1.20), Inches(0.90), Inches(0.90), Inches(0.90), Inches(0.90), Inches(0.95)],
    local_comparison_rows,
)

table3_caption = caption(method_p, "Table 3. Comparison of FloodSense with selected existing systems and approaches")
table3_caption.paragraph_format.page_break_before = True
add_comparison_table(
    doc,
    method_p,
    ["SYSTEM", "BASIS", "OUTPUT", "TIME/SCOPE", "SIMILARITY", "DIFFERENCE"],
    [Inches(1.15), Inches(1.00), Inches(0.95), Inches(0.95), Inches(0.85), Inches(0.85)],
    [
        ("FloodAdapt", "User-configured scenarios and models", "Flood, impact, and adaptation results", "Planning and research", "Scenario-based mapped support", "FloodSense uses rules, not physics or cost-benefit models"),
        ("GloFAS / IFIS", "Operational forecasts, models, and sensors", "Monitoring and forecast maps", "Operational services", "Public geographic information", "FloodSense has no live feeds, forecasts, or autonomous alerts"),
        ("HazardHunterPH / UP NOAH", "Government or research hazard layers", "Location-based hazard maps or reports", "National public reference", "Local hazard messages", "FloodSense has separate scenario inference and no claim to their datasets"),
        ("GeoMapperPH / PlanSmartPH", "Governed institutional data", "Collection or recovery-planning outputs", "Authorized LGU and agency workflows", "Structured decision information", "FloodSense serves a different user flow and disaster phase"),
        ("FloodSense", "User-confirmed hypothetical rainfall and supported location", "Explainable susceptibility plus separate guidance", "Pre-event research prototype", "GIS and decision support", "No official forecast, warning, monitoring, or automatic order"),
    ],
)

# Update only references used or corrected for Chapter 2.
references_heading = next(p for p in doc.paragraphs if p.text.strip() == "REFERENCES")
replacement_map = {
    "International Organization for Standardization. (2023). ISO/IEC 25010:2023 systems and software engineering - Systems and software Quality Requirements and Evaluation (SQuaRE) - Product quality model. https://www.iso.org/standard/78176.html":
        "International Organization for Standardization. (2023). ISO/IEC 25010:2023 systems and software engineering - Systems and software Quality Requirements and Evaluation (SQuaRE) - Product quality model. https://www.iso.org/standard/78176.html",
    "Philippine Atmospheric, Geophysical and Astronomical Services Administration. (2019, January 16). Official statement of DOST-PAGASA on the heavy rainfall event associated with Tropical Depression Usman. https://www.pagasa.dost.gov.ph/article/38":
        "Philippine Atmospheric, Geophysical and Astronomical Services Administration. (2024). 2023 annual climate bulletin. https://pubfiles.pagasa.dost.gov.ph/pagasaweb/files/cad/Climate%20Bulletin%202023_v5.pdf",
    "Philippine Statistics Authority. (2026). Component 4: Extreme events and disasters. https://psa.gov.ph/statistics/environment-statistics/highlights/component-4-extreme-events-and-disaster":
        "Philippine Statistics Authority. (2026). Highlights of the Compendium of Philippine Environment Statistics 2016-2025, Component 4: Extreme events and disasters. https://psa.gov.ph/statistics/environment-statistics/highlights/component-4-extreme-events-and-disaster",
}
for p in doc.paragraphs:
    if p.text in replacement_map:
        old_text = p.text
        p.clear()
        set_font(p.add_run(replacement_map[old_text]), size=11)
        reference_format(p)

new_references = [
    "Bentoso, L. D., Juan, E. O., Brosas, D. G., Paragas, J. R., Nuevas, L. K., & Velarde, M. W. C. (2021). Web-based solution for flood warning decision support in the Province of Leyte, Philippines. In 2021 3rd International Conference on Research and Academic Community Services (ICRACOS) (pp. 185-190). IEEE. https://doi.org/10.1109/ICRACOS53680.2021.9701990",
    "Copernicus Emergency Management Service. (2026). GloFAS homepage. European Centre for Medium-Range Weather Forecasts. https://confluence.ecmwf.int/spaces/CEMS/pages/315561239/3.1+GloFAS+homepage",
    "Cayamanda, K. J. G., Paunlagui, M. M., Baconguis, R. D. T., & Quimbo, M. A. T. (2021). Vulnerability profile and risk perception towards an inclusive disaster risk reduction for the flood vulnerable communities of Davao City. International Review of Social Sciences Research, 1(1), 25-53. https://doi.org/10.53378/346474",
    "Dasallas, L., An, H., & Lee, S. (2022). Developing an integrated multiscale rainfall-runoff and inundation model: Application to an extreme rainfall event in Marikina-Pasig River Basin, Philippines. Journal of Hydrology: Regional Studies, 39, 100995. https://doi.org/10.1016/j.ejrh.2022.100995",
    "Dulawan, J. M. T., Imamura, Y., Konishi, T., Amaguchi, H., & Ohara, M. (2024). A systematic framework for assessing social vulnerability to flood for integrated flood risk management: A case study in Metro Manila, Philippines. International Journal of Disaster Risk Reduction, 112, 104778. https://doi.org/10.1016/j.ijdrr.2024.104778",
    "Deltares. (2025, December 2). FloodAdapt. https://www.deltares.nl/en/software-and-data/products/floodadapt",
    "Department of Science and Technology-Philippine Institute of Volcanology and Seismology. (2022a, December 6). DOST-PHIVOLCS conducts training on the use of GeoRiskPH platforms for the Municipality of General Luna, Quezon. GeoRisk Philippines Initiative. https://www.georisk.gov.ph/articles/2022/12/dost-phivolcs-conducts-training-on-the-use-of-georiskph-platforms-for-the-municipality-of-general-luna-quezon",
    "Department of Science and Technology-Philippine Institute of Volcanology and Seismology. (2022b, September 28). DOST-PHIVOLCS conducts PlanSmart Ready to Rebuild training of trainers. GeoRisk Philippines Initiative. https://www.georisk.gov.ph/articles/2022/9/dost-phivolcs-conducts-plansmart-ready-to-rebuild-training-of-trainers",
    "Environment Agency. (2024). National assessment of flood and coastal erosion risk in England 2024. GOV.UK. https://www.gov.uk/government/publications/national-assessment-of-flood-and-coastal-erosion-risk-in-england-2024/national-assessment-of-flood-and-coastal-erosion-risk-in-england-2024",
    "Federal Emergency Management Agency. (2025). Resilience Analysis and Planning Tool user guide. https://www.fema.gov/sites/default/files/documents/fema_rapt-user-guide_2025.pdf",
    "Gacul, L., Ferrancullo, D., Gallano, R., Fadriquela, K. J., Mendez, K. J., Morada, J. R., Morgado, J. K., & Gacu, J. (2024). GIS-based identification of flood risk zone in a rural municipality using fuzzy analytical hierarchy process (FAHP). Revue Internationale de Geomatique, 33(1), 295-320. https://doi.org/10.32604/rig.2024.055085",
    "GeoRisk Philippines. (n.d.). HazardHunterPH: Hazard assessment at your fingertips. Retrieved September 21, 2026, from https://hazardhunter.georisk.gov.ph/map",
    "Gumba, G., Brosas, D. G., & Paragas, J. R. (2021). Assessment of SIAS application using software quality model. In 2021 3rd International Conference on Research and Academic Community Services (ICRACOS) (pp. 197-202). IEEE. https://doi.org/10.1109/ICRACOS53680.2021.9701982",
    "Johnson, B. A., Estoque, R. C., Li, X., Kumar, P., Dasgupta, R., Avtar, R., & Magcale-Macandog, D. B. (2021). High-resolution urban change modeling and flood exposure estimation at a national scale using open geospatial data: A case study of the Philippines. Computers, Environment and Urban Systems, 90, 101704. https://doi.org/10.1016/j.compenvurbsys.2021.101704",
    "Iowa Flood Center. (2024). A strong purpose. IIHR Currents 2023-24. University of Iowa. https://iihr.uiowa.edu/news/iihr-currents/iihr-currents-2023-24/strong-purpose",
    "Iowa Flood Center. (n.d.). Iowa Flood Information System terms and conditions. University of Iowa. Retrieved September 21, 2026, from https://s-iihr72.iihr.uiowa.edu/beta.ifis.iowawis.org/terms.php",
    "Macalalad, R. V., Xu, S., Badilla, R. A., Paat, S. F., Tajones, B. C., Chen, Y., & Bagtasa, G. (2021). Flash flood modeling in the data-poor basin: A case study in Matina River Basin. Tropical Cyclone Research and Review, 10(2), 87-95. https://doi.org/10.1016/j.tcrr.2021.06.003",
    "Mercado, J. M., Kawamura, A., & Medina, R. (2023). An expanded interpretive structural modeling analysis of the barriers to integrated flood risk management adaptation in Metro Manila. Water, 15(6), 1029. https://doi.org/10.3390/w15061029",
    "Mostafiz, R. B., Rohli, R. V., Friedland, C. J., & Lee, Y.-C. (2022). Actionable information in flood risk communications and the potential for new web-based tools for long-term planning for individuals and community. Frontiers in Earth Science, 10, 840250. https://doi.org/10.3389/feart.2022.840250",
    "National Disaster Risk Reduction and Management Council. (2024). National disaster response plan 2024: Hydro-meteorological hazards. https://ndrrmc.gov.ph/attachments/article/4125/National_Disaster_Response_Plan_NDRP_2024.pdf",
    "Papadopoulos, P., Soflano, M., Chaudy, Y., Adejo, W., & Connolly, T. M. (2022). A systematic review of technologies and standards used in the development of rule-based clinical decision support systems. Health and Technology, 12, 713-727. https://doi.org/10.1007/s12553-022-00672-9",
    "Philippine Statistics Authority. (2023). Third quarter 2023 PSGC updates. https://psa.gov.ph/classification/psgc/node/1684061390",
    "Philippine Statistics Authority. (2025). City of Bacoor: Barangays in the City of Bacoor. https://psa.gov.ph/classification/psgc/barangays/0402103000",
    "PostGIS Development Group. (2026). PostGIS manual: Introduction. https://postgis.net/docs/manual-3.7/en/postgis_introduction.html",
    "UP NOAH Center. (n.d.). Know your hazards. University of the Philippines Resilience Institute. Retrieved September 21, 2026, from https://noah.up.edu.ph/know-your-hazards",
    "Williams, S., Griffiths, J., Miville, B., Romeo, E., Leiofi, M., O'Driscoll, M., Iakopo, M., Mulitalo, S., Ting, J. C., Paulik, R., & Elley, G. (2021). An impacts-based flood decision support system for a tropical Pacific island catchment with short warnings lead time. Water, 13(23), 3371. https://doi.org/10.3390/w13233371",
]

existing_texts = {p.text for p in doc.paragraphs}
for ref in new_references:
    if ref in existing_texts:
        continue
    references_element_index = body_el.index(references_heading._p)
    anchors = [
        p
        for p in doc.paragraphs
        if p.text
        and p.text != "REFERENCES"
        and p.text > ref
        and body_el.index(p._p) > references_element_index
    ]
    anchor = anchors[0] if anchors else None
    if anchor is None:
        p = doc.add_paragraph()
        set_font(p.add_run(ref), size=11)
        reference_format(p)
    else:
        add_reference_before(anchor, ref)
    existing_texts.add(ref)

# Ask Word to refresh page and page-count fields when opened.
settings = doc.settings._element
update_fields = settings.find(qn("w:updateFields"))
if update_fields is None:
    update_fields = OxmlElement("w:updateFields")
    settings.append(update_fields)
update_fields.set(qn("w:val"), "true")

doc.save(OUTPUT)

final_doc = Document(OUTPUT)
after_ch1 = c14n_slice(final_doc, end_text="REVIEW OF RELATED LITERATURE")
after_ch3 = c14n_slice(final_doc, start_text="METHODOLOGY", end_text="REFERENCES")

report = {
    "source": str(SOURCE),
    "source_sha256": sha256(SOURCE),
    "output": str(OUTPUT),
    "output_sha256": sha256(OUTPUT),
    "chapter1_xml_hash_before": before_ch1,
    "chapter1_xml_hash_after": after_ch1,
    "chapter1_unchanged": before_ch1 == after_ch1,
    "chapter3_xml_hash_before": before_ch3,
    "chapter3_xml_hash_after": after_ch3,
    "chapter3_unchanged": before_ch3 == after_ch3,
    "sections_before": before_sections,
    "sections_after": len(final_doc.sections),
    "tables_before": before_tables,
    "tables_after": len(final_doc.tables),
    "inline_shapes_before": before_shapes,
    "inline_shapes_after": len(final_doc.inline_shapes),
    "chapter2_paragraph_count": sum(
        1
        for p in final_doc.paragraphs[
            next(i for i, p in enumerate(final_doc.paragraphs) if p.text.strip() == "REVIEW OF RELATED LITERATURE") :
            next(i for i, p in enumerate(final_doc.paragraphs) if p.text.strip() == "METHODOLOGY")
        ]
    ),
}
REPORT.write_text(json.dumps(report, indent=2), encoding="utf-8")
print(json.dumps(report, indent=2))
