from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


SOURCE = Path(r"1. FLOODSENSE-MAIN/MAIN FILESS/FloodSense_Chapter_1_REvision_Version_1.0.docx")
OUT = Path(r"1. FLOODSENSE-MAIN/MAIN FILESS/FloodSense_Manuscript_Chapters_1_to_3_Final.docx")
FONT = "Arial"
SIZE = 11


def set_run(run, size=SIZE, bold=False, italic=False):
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    for key in ("ascii", "hAnsi", "eastAsia"):
        rpr.rFonts.set(qn(f"w:{key}"), FONT)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)


def clear_paragraph(paragraph):
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)


def replace_text(paragraph, text, lead=None):
    clear_paragraph(paragraph)
    if lead and text.startswith(lead):
        set_run(paragraph.add_run(lead), bold=True)
        set_run(paragraph.add_run(text[len(lead):]))
    else:
        set_run(paragraph.add_run(text))


def body(doc, text, first_indent=True, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    if first_indent:
        pf.first_line_indent = Inches(0.5)
    set_run(p.add_run(text), italic=italic)
    return p


def body_parts(doc, parts, first_indent=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    if first_indent:
        pf.first_line_indent = Inches(0.5)
    for text, bold, italic in parts:
        set_run(p.add_run(text), bold=bold, italic=italic)
    return p


def lead(doc, heading, text):
    return body_parts(doc, [(heading, True, False), (text, False, False)])


def chapter(doc, text, page_break=True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.page_break_before = page_break
    pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(25.3)
    pf.keep_with_next = True
    set_run(p.add_run(text.upper()), bold=True)
    return p


def heading(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    pf.space_before = Pt(25.3)
    pf.space_after = Pt(0)
    pf.keep_with_next = True
    set_run(p.add_run(text), bold=True)
    return p


def subheading(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.keep_with_next = True
    set_run(p.add_run(text), bold=True)
    return p


def note(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    pf.space_before = Pt(12.65)
    pf.space_after = Pt(12.65)
    pf.keep_together = True
    set_run(p.add_run(f"[{text}]"), italic=True)
    return p


def reference(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.left_indent = Inches(0.5)
    pf.first_line_indent = Inches(-0.5)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(12.65)
    set_run(p.add_run(text))
    return p


def configure_page(section):
    section.page_width = Cm(21.03)
    section.page_height = Cm(29.70)
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.5)
    section.right_margin = Inches(1.0)
    section.header_distance = Inches(0.45)
    section.footer_distance = Inches(0.5)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, end])
    set_run(run)


def set_page_start(section, number):
    node = section._sectPr.find(qn("w:pgNumType"))
    if node is None:
        node = OxmlElement("w:pgNumType")
        section._sectPr.append(node)
    node.set(qn("w:start"), str(number))


def set_paragraph_border(paragraph):
    ppr = paragraph._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    ppr.append(borders)
    for edge in ("top", "bottom"):
        item = OxmlElement(f"w:{edge}")
        item.set(qn("w:val"), "single")
        item.set(qn("w:sz"), "12")
        item.set(qn("w:space"), "3")
        item.set(qn("w:color"), "000000")
        borders.append(item)


source_doc = Document(SOURCE)
doc = Document()
configure_page(doc.sections[0])

# Keep the approved Chapter 1 text and SURECUT page design, replacing only the
# specifically confirmed rule-governance, CSV exchange, and currency-of-source passages.
normal = doc.styles["Normal"]
normal.font.name = FONT
normal._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), FONT)
normal._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), FONT)
normal.font.size = Pt(SIZE)

replacements = {
    "The Android application presents the assessment through a dynamic GIS-style map": (
        "The Android application presents the assessment through a dynamic GIS-style map that users can pan and zoom. Selectable geographic polygons display susceptibility through a consistent color scheme, while Check My Area allows a user to position a pin and obtain the assessment associated with that selected location. The map also presents verified evacuation-center locations and relevant details. Users select rainfall intensity from Light, Moderate, Heavy, Intense, or Torrential and select a duration of 1, 3, 6, 12, or 24 hours; the Expert System evaluates the supported scenario against approved geographic and historical information. A separate Web Administration System enables authorized administrators to maintain DSS preparedness content, permitted fixed-table values through canonical CSV export and import, map polygons and points, geographic records, evacuation centers, provenance information, and user accounts. Expert rules and the inference algorithm remain protected from ordinary administrator access. Through this arrangement, FloodSense translates institutional knowledge into an accessible susceptibility assessment, spatial visualization, and practical pre-event preparedness support without claiming current-condition prediction or substituting for official BDRRMO and Philippine Atmospheric, Geophysical and Astronomical Services Administration (PAGASA) warnings.",
        None,
    ),
    "The problem also concerns the accessibility and operational maintenance of institutional knowledge.": (
        "The problem also concerns the accessibility and operational maintenance of institutional knowledge. The BDRRMO maintains historical flood records, exposure information, preparedness protocols, and evacuation-center information, while the CPDCO maintains geographic and planning data. Current dissemination practices, including social media, email, and barangay coordination, do not fully address the public's need for on-demand, situation-specific guidance. A community-facing platform must therefore communicate approved information while ensuring that map objects, preparedness content, permitted fixed-table values, data sources, and administrative accounts can be maintained by authorized personnel while expert rules and inference logic remain protected. How can a Rule-Based Expert System, dynamic GIS-style map, Decision Support System, and Web Administration System organize this institutional knowledge into an explainable assessment and appropriate pre-event preparedness guidance for Bacoor City residents?",
        None,
    ),
    "develop an Android application using Flutter and Dart": (
        "develop an Android application using Flutter and Dart, supported by Django, Django REST Framework, GeoDjango, and PostgreSQL/PostGIS, that integrates a Rule-Based Expert System, a protected knowledge base and fixed inference process, PAGASA-aligned rainfall input, a dynamic GIS-style map using flutter_map, selectable color-coded susceptibility polygons, Check My Area, evacuation-center locations, a Decision Support System for pre-event preparedness, user account and onboarding functions, and a separate Web Administration System with a Leaflet-based administrative map for authorized maintenance of DSS content, geographic records, map polygons and points, permitted fixed-table values through canonical CSV export and import, provenance, evacuation centers, and users;",
        None,
    ),
    "FloodSense addresses this gap through an Android application": (
        "FloodSense addresses this gap through an Android application that connects expert-provided and expert-validated knowledge with the information and preparedness needs of Bacoor City residents. A Rule-Based Expert System organizes approved facts and conditional rules in a knowledge base and applies an inference process to the user's selected rainfall intensity, rainfall duration, and geographic area. Rule-based decision support can present a traceable basis for a result through knowledge-based and human-understandable explanations (Kostopoulos et al., 2024). In FloodSense, the conclusion is an explainable flood susceptibility assessment expressed as Low, Moderate, High, or Very High. The Decision Support System (DSS) then uses the assessment together with the user's situational inputs to present suitable pre-event preparedness guidance. This function supports, but does not make, the user's final decision; decision-support frameworks organize relevant information and structured alternatives for human decision-making (Alabbad et al., 2022).",
        None,
    ),
    "Web Administration and Knowledge Management Module.": (
        "Web Administration and Knowledge Management Module. This module is the controlled maintenance workspace used by authorized administrators. It manages permitted geographic zones and map features, Decision Support System questions and branches, evacuation-center records, public explanations and disclaimers, source and provenance information, user-support actions, and the values used by the fixed assessment table. The Expert System rules, rule conditions, priorities, conflict-resolution procedure, and inference algorithm are protected and are not displayed, created, edited, deleted, or replaced through the ordinary administrator interface. For permitted table-value maintenance, the administrator exports the canonical comma-separated values (CSV) file generated by FloodSense. The exported file retains the exact table structure required by the algorithm. An authorized editor may change only the values stored within the existing fields; column names, column order, parameter identifiers, required keys, and the table structure must remain unchanged. The administrator may then import the updated CSV file. FloodSense validates the fixed schema before accepting it, and the unchanged algorithm reads the updated values through that same schema. The process does not import expert rules, alter the inference logic, or train a machine-learning model.",
        "Web Administration and Knowledge Management Module. ",
    ),
    "Rule-Based Expert System Assessment Module.": (
        "Rule-Based Expert System Assessment Module. This module implements a deterministic Rule-Based Expert System. Its knowledge base stores domain knowledge, while its inference procedure applies eligible rules to known facts to reach a conclusion. Knowledge-based and rule-oriented explanations can make decision-support outputs more understandable and traceable (Kostopoulos et al., 2024). FloodSense represents approved knowledge through geographic-zone facts, rainfall categories and durations, decision tables, structured IF-THEN rules, exceptions, priorities, sources, versions, and explanation text. Forward chaining begins with the resolved zone, rainfall scenario, active zone profile, and published rule set; it tests conditions and fires the applicable rules. The primary outputs are Low, Moderate, High, and Very High. Conflicting rules may return Uncertain, missing coverage may return Insufficient Data, and a point beyond approved boundaries returns Outside Supported Area.",
        "Rule-Based Expert System Assessment Module. ",
    ),
    "Decision Support System Module.": (
        "Decision Support System Module. This module provides pre-event preparedness guidance through a deterministic, database-driven decision tree. Decision-support frameworks organize relevant information, scenarios, and structured alternatives to assist human decision-making (Alabbad et al., 2022). FloodSense begins with the latest valid Expert System result as read-only context and combines it with the resident's answers to approved situational questions. The configured path may provide guidance concerning go-bag preparation, family communication, protection of documents and valuables, assistance for vulnerable household members, monitoring of official advisories, and evacuation readiness. The DSS is separate from the Expert System: it does not change the susceptibility class, predict current conditions, or issue an independent evacuation order.",
        "Decision Support System Module. ",
    ),
    "For the Bacoor Disaster Risk Reduction and Management Office (BDRRMO).": (
        "For the Bacoor Disaster Risk Reduction and Management Office (BDRRMO). The application provides a complementary channel for communicating approved, location-specific flood susceptibility information and preparedness content beyond social media, email, and barangay coordination. Through the Web Administration System, authorized personnel can maintain DSS guidance, permitted table values through the fixed-schema CSV export-and-import process, map features, geographic information, evacuation-center records, provenance, and user accounts. Expert rules and the inference algorithm remain protected from ordinary administrator access. This supports continuity, accountability, and public accessibility while preserving the authority of official advisories and emergency directives.",
        "For the Bacoor Disaster Risk Reduction and Management Office (BDRRMO). ",
    ),
    "A separate Web Administration System is included for authorized administrators.": (
        "A separate Web Administration System is included for authorized administrators. It provides role-controlled functions for managing DSS questions and guidance, permitted rainfall and geographic values, map polygons and points, geographic attributes, evacuation-center records, provenance and validation information, and user accounts. Expert rules, rule conditions, priorities, conflict-resolution logic, and the inference algorithm are fixed and protected from ordinary administrator access. For permitted value updates, the administrator exports the canonical CSV table, an authorized editor changes only the values within its existing fields, and the administrator imports the file with the same column names, order, identifiers, required keys, and overall structure. Files that do not match the required schema are rejected. The CSV process does not change or import expert rules. A Leaflet-based administrative map supports the maintenance of approved points and polygons, while susceptibility colors presented to residents are produced from Expert System results rather than manual painting of assessment areas. Administrative actions are subject to authentication, authorization, validation, and audit-related information appropriate to the prototype.",
        None,
    ),
    "The primary users of the Android application are Bacoor City residents": (
        "The primary users of the Android application are Bacoor City residents, particularly those from communities that have experienced flooding. Users can register and sign in, review onboarding information and a disclaimer, select a rainfall intensity of Light, Moderate, Heavy, Intense, or Torrential, and select a duration of 1, 3, 6, 12, or 24 hours. The rainfall options represent hypothetical scenarios and are mapped internally to values stored under the fixed fields required by the algorithm. Permitted values may be updated only through the canonical CSV export-and-import process without changing the table structure or inference logic. The application does not require users to enter numerical rainfall measurements, elevation, land use, historical flood frequency, or technical geographic attributes. These facts are maintained in the system through authorized records and are applied by the Expert System during assessment.",
        None,
    ),
    "The feasible implementation stack consists of Flutter and Dart": (
        "The feasible implementation stack consists of Flutter and Dart for the Android application; flutter_map for the resident-facing map; Python Django and Django REST Framework for backend services and application programming interfaces; GeoDjango for geographic operations; PostgreSQL with PostGIS for structured and spatial records; and Leaflet for the administrative map. The application requires a compatible Android device and an internet connection to retrieve protected knowledge, geographic information, assessment results, preparedness content, accounts, and evacuation-center data. An iOS application and a fully offline mode are outside the study. The system also excludes live sensors, automated meteorological feeds, flood-depth measurement, emergency dispatch, direct control of public warning infrastructure, and continuous background tracking.",
        None,
    ),
    "The system architecture uses a vertical client-server presentation": (
        "The System Architecture serves as the visual representation of the FloodSense theoretical framework. It uses a vertical client-server presentation of the approved FloodSense design. Resident requests enter through authentication, rainfall and location input, and the secured Django service layer. Institutional data and expert review enter through the Web Administration and Knowledge Management Module. The Rule-Based Expert System produces the susceptibility result, which is used by the map and explanation modules and passed as read-only context to the separate DSS. The database stores the controlled relational and spatial records, while basemap and navigation services are used only for display and external directions.",
        None,
    ),
    "Data privacy and security are addressed in accordance with Republic Act No. 10173": (
        "Data privacy and security are addressed through applicable Philippine privacy requirements. The prototype limits collection to information necessary for accounts, administration, testing, and requested application functions. Selected map points and situational answers are processed only to produce the requested assessment and preparedness support according to the implemented design; continuous location history is not part of the study. Access to administrative functions is limited to authorized accounts. The evaluation uses a pilot group and the resulting findings may not be generalized beyond Bacoor City without testing among larger, more diverse populations and validation against the knowledge, geography, and institutional processes of another locality.",
        None,
    ),
    "Knowledge Base.": (
        "Knowledge Base. The structured collection of expert-validated facts, protected conditional rules, geographic attributes, fixed table fields, approved values, susceptibility categories, and related provenance used by the Expert System. Ordinary administrators cannot view or modify the protected rules or inference logic.",
        "Knowledge Base. ",
    ),
    "PAGASA-Aligned Rainfall Input.": (
        "PAGASA-Aligned Rainfall Input. The user interface for selecting Light, Moderate, Heavy, Intense, or Torrential rainfall and a duration of 1, 3, 6, 12, or 24 hours. These selections represent the rainfall scenarios supported by FloodSense and do not claim to report current weather conditions.",
        "PAGASA-Aligned Rainfall Input. ",
    ),
    "Rainfall Parameters.": (
        "Rainfall Parameters. The fixed fields used to represent supported rainfall intensity and duration scenarios. Through the CSV workflow, only the values stored within the existing fields may be updated; the field names, order, identifiers, required keys, and structure remain unchanged.",
        "Rainfall Parameters. ",
    ),
    "Preparedness Provenance.": (
        "Preparedness Provenance. Information identifying the source, validation status, responsible custodian, and relevant date associated with a protected rule, geographic record, preparedness item, permitted table value, or evacuation-center entry.",
        "Preparedness Provenance. ",
    ),
    "Web Administration System.": (
        "Web Administration System. The browser-based interface used by authorized administrators to maintain permitted data, DSS content, geographic records, map polygons and points, evacuation centers, provenance information, user accounts, and fixed-table values through canonical CSV export and import. It does not expose or permit changes to expert rules, rule conditions, priorities, conflict-resolution logic, or the inference algorithm.",
        "Web Administration System. ",
    ),
}

for paragraph in doc.paragraphs:
    text = paragraph.text.strip()
    for prefix, (new_text, lead_text) in replacements.items():
        if text.startswith(prefix):
            replace_text(paragraph, new_text, lead_text)
            break

# SURECUT-style title page.
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
p.paragraph_format.space_after = Pt(135)
set_run(p.add_run("FLOODSENSE: AN ANDROID-BASED EXPERT SYSTEM FOR FLOOD\nSUSCEPTIBILITY ASSESSMENT AND PRE-EVENT PREPAREDNESS IN\nBACOOR CITY"), bold=True)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
p.paragraph_format.space_after = Pt(132)
set_run(p.add_run("Undergraduate Thesis\nSubmitted to the Faculty of the\nDepartment of Computer Studies\nCavite State University - Bacoor City Campus\nCity of Bacoor, Cavite"))

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
p.paragraph_format.space_after = Pt(125)
set_run(p.add_run("In partial fulfillment of the\nrequirements for the degree\nBachelor of Science in Computer Science"))

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
set_run(p.add_run("REYMART V. GOC-ONG\nMARTIN LORENZ D. JUANITES\nROMAR T. PULAO"), bold=True)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_run(p.add_run("May 2027"))

# First manuscript page. Page 1 is counted but hidden; subsequent page numbers
# appear at the upper-right, matching SURECUT.
main = doc.add_section(WD_SECTION.NEW_PAGE)
configure_page(main)
main.header.is_linked_to_previous = False
main.first_page_header.is_linked_to_previous = False
main.different_first_page_header_footer = True
add_page_number(main.header.paragraphs[0])
main.first_page_header.paragraphs[0].clear()
set_page_start(main, 1)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
p.paragraph_format.space_after = Pt(28)
set_run(p.add_run("FloodSense: An Android-Based Expert System for Flood Susceptibility\nAssessment and Pre-Event Preparedness in Bacoor City"), bold=True)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
p.paragraph_format.space_after = Pt(28)
set_run(p.add_run("Reymart V. Goc-ong\nMartin Lorenz D. Juanites\nRomar T. Pulao"), bold=True)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
p.paragraph_format.space_after = Pt(24)
set_paragraph_border(p)
set_run(p.add_run("An undergraduate thesis manuscript submitted to the faculty of the Department of Computer Studies, Cavite State University - Bacoor City Campus, City of Bacoor, Cavite, in partial fulfillment of the requirements for the degree of Bachelor of Science in Computer Science with Contribution No. __________________. Prepared under the supervision of ________________________________."))

# Recreate the original Chapter 1 text in a clean document container. This
# avoids carrying forward the source file's malformed drawing objects while
# preserving the chapter's content and order.
section_titles = {
    "Statement of the Problem",
    "Objectives of the Study",
    "Theoretical Framework",
    "System Architecture",
    "Significance of the Study",
    "Time and Place of the Study",
    "Scope and Limitation of the Study",
    "Definition of Terms",
}
lead_prefixes = (
    "Authentication and Onboarding Module. ",
    "PAGASA-Aligned Rainfall and Location Input Module. ",
    "Web Administration and Knowledge Management Module. ",
    "Django REST and GeoDjango Processing Module. ",
    "Rule-Based Expert System Assessment Module. ",
    "Dynamic GIS Map and Check My Area Module. ",
    "Results Summary and Explainable Result Module. ",
    "Decision Support System Module. ",
    "Verified Evacuation Center Resource Module. ",
    "User Account and Preferences Module. ",
    "PostgreSQL and PostGIS Database Module. ",
    "For the Bacoor Disaster Risk Reduction and Management Office (BDRRMO). ",
    "For the City Planning and Development Coordinator Office - Bacoor (CPDCO). ",
    "For residents of historically flood-affected barangays. ",
    "For future researchers. ",
    "Android Application. ",
    "Check My Area. ",
    "Color-Coded Susceptibility Levels. ",
    "CPDCO (City Planning and Development Coordinator Office). ",
    "Decision Support System (DSS). ",
    "Dynamic Map. ",
    "Expert-Provided and Expert-Validated Knowledge. ",
    "Expert System. ",
    "Flood Susceptibility. ",
    "Geographic Information System (GIS). ",
    "GeoDjango. ",
    "Inference Process. ",
    "Knowledge Base. ",
    "PAGASA-Aligned Rainfall Input. ",
    "PostGIS. ",
    "Preparedness Provenance. ",
    "Rainfall Parameters. ",
    "Rule-Based Expert System. ",
    "Web Administration System. ",
)
copying = False
objective_number = 0
for source_paragraph in source_doc.paragraphs:
    text = source_paragraph.text.strip()
    if text == "INTRODUCTION":
        copying = True
        chapter(doc, text, page_break=False)
        continue
    if not copying or not text:
        continue
    if text == "REFERENCES":
        break
    for prefix, (new_text, _lead_text) in replacements.items():
        if text.startswith(prefix):
            text = new_text
            break
    if text in section_titles:
        heading(doc, text)
        if text == "Theoretical Framework":
            objective_number = 0
        if text == "System Architecture":
            note(doc, "Insert the FloodSense System Architecture diagram, which represents the theoretical framework, in this location.")
        continue
    if text == "Specifically, it aimed to:":
        body(doc, text)
        objective_number = 1
        continue
    if objective_number and objective_number <= 5:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = p.paragraph_format
        pf.left_indent = Inches(0.75)
        pf.first_line_indent = Inches(-0.25)
        pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        set_run(p.add_run(f"{objective_number}.  {text}"))
        objective_number += 1
        continue
    matched_lead = next((prefix for prefix in lead_prefixes if text.startswith(prefix)), None)
    if matched_lead:
        lead(doc, matched_lead, text[len(matched_lead):])
    else:
        body(doc, text)


# CHAPTER 2
chapter(doc, "REVIEW OF RELATED LITERATURE")
body(doc, "This chapter presents literature and studies that provide the basis for the flood-susceptibility, geographic-information, decision-support, public-communication, and software-quality requirements of FloodSense. The sources are examined according to their relevance to a scenario-based Android application for Bacoor City. Methods, factors, formulas, and weights from other locations are not treated as FloodSense parameters unless they are supported by local evidence and expert validation.")

heading(doc, "Related Foreign Literature")
subheading(doc, "Flood Susceptibility and Spatial Factors")
body(doc, "Flood susceptibility expresses the relative tendency of an area to experience flooding under defined conditions. Debnath et al. (2024) demonstrated that geospatial and expert-informed approaches can combine multiple flood-conditioning factors and classify areas into ordered susceptibility zones. The study also identified data scarcity, spatial resolution, validation, and the availability of long-term hydrological records as important limitations. These concerns support the use of explicit limitation states in FloodSense when the available Bacoor data cannot justify a classification.")
body(doc, "Nguyen et al. (2024) combined a Geographic Information System and an analytical hierarchy process to assess urban flood susceptibility using multiple factors, expert consultation, historical flood points, and receiver operating characteristic validation. The study supports clear factor definitions, documented expert judgment, spatial preparation, and independent validation. Its factors and weights are not transferred to FloodSense because their applicability to Bacoor City has not been established.")

subheading(doc, "Interactive Mapping and Risk Communication")
body(doc, "Oubennaceur et al. (2021) found that interactive story maps can combine hazard maps, explanatory text, and user-oriented interaction to communicate flood risk to non-specialists. Clear legends, understandable categories, careful color use, and limited technical language are relevant to the FloodSense map and explanation screens. A map color must be accompanied by text and supporting information so that the user is not required to interpret a technical layer without context.")
body(doc, "Lin (2023) examined relationships among flood-risk information seeking, application use, coping beliefs, collective efficacy, and community action. The findings support connecting a susceptibility result to understandable and source-backed preparedness information. They do not support presenting a software output as an official order or a guarantee of safety.")

subheading(doc, "Decision Support and Explainability")
body(doc, "Alabbad et al. (2022) developed a flood-mitigation decision-support framework that combined maps, scenarios, guidance, and decision-tree analysis. The study illustrates how a Decision Support System can organize relevant alternatives and contextual information without making an unexplained final decision for the user. FloodSense applies this principle to pre-event household preparedness and keeps the Decision Support System separate from the susceptibility classification performed by the Expert System.")
body(doc, "Kostopoulos et al. (2024) reviewed explainable decision-support approaches, including visual, rule-based, case-based, natural-language, and knowledge-based explanations. Their review emphasizes the importance of domain knowledge and explanations that people can understand. Galanti et al. (2023) likewise showed that explanation quality must be evaluated with intended users. These findings support displaying the basis and limitations of a FloodSense result while protecting the rule set and inference procedure from ordinary administrator modification.")

heading(doc, "Related Local Literature")
subheading(doc, "Flood Conditions and Institutional Responsibility")
body(doc, "The Japan International Cooperation Agency (JICA, 2021) reported the inauguration of flood-mitigation facilities in the Imus River Basin, including structures intended to reduce flooding in low-lying portions of Bacoor and Imus. The project confirms the relevance of basin-level flood management to Bacoor City and shows that a resident-facing application operates alongside government planning, engineering measures, and official emergency functions.")

subheading(doc, "Scenario-Based Hazard Information")
body(doc, "GeoRisk Philippines (2026) presents HazardHunterPH as a reference platform for examining hazard information and cautions users that map outputs are not absolute predictions. This boundary supports source labeling, explicit limitations, and the requirement that FloodSense must not treat a basemap or barangay boundary as sufficient evidence for one susceptibility conclusion.")

subheading(doc, "Local Requirements and Data Governance")
body(doc, "The Bacoor Disaster Risk Reduction and Management Office interviews identified the need to preserve authoritative records, avoid unsupported barangay-wide generalization, validate evacuation-center information, and maintain preparedness content from authorized sources (FloodSense Research Team, 2026c; 2026d). The City Planning and Development Coordinator Office interview identified geographic materials relevant to the study but did not establish that every available layer was already suitable for algorithm use (FloodSense Research Team, 2026e). These records support documented provenance, validation, limitations, and controlled maintenance.")

heading(doc, "Related Foreign Studies")
body(doc, "Recent foreign studies demonstrate several approaches relevant to the FloodSense design. Oubennaceur et al. (2021) focused on interactive communication of mapped flood risk; Alabbad et al. (2022) developed a decision-support framework for mitigation alternatives; Galanti et al. (2023) examined understandable explanations in decision-support systems; Debnath et al. (2024) compared geospatial and expert-informed susceptibility models; and Nguyen et al. (2024) developed and validated an urban susceptibility map. Collectively, these studies show that public-facing flood applications require defensible data, a clearly bounded method, validation, understandable maps, and communication appropriate to the system's evidence and authority.")
body(doc, "The methods of the reviewed studies cannot be transferred automatically to Bacoor City. Susceptibility research depends on the selected factors, measurements, geographic resolution, weights, and validation observations. FloodSense therefore relies only on locally supported and approved values within its fixed table structure and returns a limitation state when the available knowledge cannot support a conclusion.")

heading(doc, "Related Local Studies")
body(doc, "Johnson et al. (2021) modeled urban expansion and flood exposure across the Philippines using open geospatial data. The study indicates that urban growth can increase the number of people and developed areas exposed to flooding, supporting the need for spatially explicit preparedness information. Its national-scale results cannot, however, be treated as a location-level Bacoor City assessment without local validation.")
body(doc, "Nagumo et al. (2022) created a three-dimensional flood-hazard map for a flood-prone Philippine area and emphasized that visualization can help residents understand local flood characteristics. The study supports accessible map presentation but also demonstrates that a visualization must be connected to a documented method and appropriate local data. Neither local study provides a validated set of fixed values for the FloodSense algorithm.")

heading(doc, "Synthesis")
body(doc, "The reviewed sources establish five requirements for FloodSense. First, susceptibility must be tied to defined conditions and supported geographic resolution. Second, an interactive map requires understandable legends, labels, explanations, and limitations. Third, preparedness guidance must remain separate from susceptibility classification and must leave final decisions to residents and authorized institutions. Fourth, rule-based results should be explainable, while the rule set and inference logic remain protected from ordinary administrative modification. Fifth, values obtained from local sources must retain their provenance and be accepted only through a controlled structure that the fixed algorithm can read.")
body(doc, "FloodSense addresses a local information gap by connecting a user-selected rainfall scenario and supported Bacoor City location to a deterministic susceptibility assessment and separately sourced pre-event preparedness guidance. Its contribution is the coordinated use of protected rule-based reasoning, geographic presentation, understandable results, controlled data maintenance, and explicit limitations. It does not adopt another study's formula, factor weights, or conclusions as Bacoor City parameters without local evidence and validation.")

p = doc.add_paragraph()
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
p.paragraph_format.space_before = Pt(12.65)
p.paragraph_format.space_after = Pt(6)
set_run(p.add_run("Table 1. Comparative Synthesis of Selected Related Studies and Systems"), italic=True)

table = doc.add_table(rows=1, cols=4)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.style = "Table Grid"
table.autofit = False
widths = [Inches(1.25), Inches(1.15), Inches(1.55), Inches(1.82)]
headers = ["SYSTEM OR STUDY", "AUTHOR(S)", "METHOD", "RELEVANCE AND BOUNDARY"]
for i, (cell, text) in enumerate(zip(table.rows[0].cells, headers)):
    cell.width = widths[i]
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell.paragraphs[0].paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    set_run(cell.paragraphs[0].add_run(text), size=9, bold=True)

rows = [
    ("Flood Risk StoryMaps", "Oubennaceur et al. (2021)", "Interactive risk communication", "Supports understandable spatial communication; does not provide Bacoor parameters."),
    ("Flood Mitigation DSS", "Alabbad et al. (2022)", "Web decision-support framework", "Supports structured guidance while leaving the final decision to people."),
    ("Flood Susceptibility Model", "Debnath et al. (2024)", "Geospatial and expert-informed models", "Supports multiple factors and validation; methods require local evidence."),
    ("Urban Susceptibility Map", "Nguyen et al. (2024)", "GIS and analytical hierarchy process", "Supports factor documentation and validation; weights are not transferred."),
    ("HazardHunterPH", "GeoRisk Philippines (2026)", "National hazard-reference platform", "Supports source and limitation notices; outputs are not absolute predictions."),
    ("FloodSense", "Present Study", "Fixed rule-based scenario assessment", "Provides explainable Bacoor-focused assessment using protected rules and controlled values."),
]
for row in rows:
    cells = table.add_row().cells
    for i, (cell, text) in enumerate(zip(cells, row)):
        cell.width = widths[i]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.paragraph_format.space_after = Pt(0)
        set_run(p.add_run(text), size=9)

tr_pr = table.rows[0]._tr.get_or_add_trPr()
tbl_header = OxmlElement("w:tblHeader")
tbl_header.set(qn("w:val"), "true")
tr_pr.append(tbl_header)


# CHAPTER 3
chapter(doc, "METHODOLOGY")
body(doc, "This chapter presents the materials and documented activities used in developing FloodSense, the participants involved in completed requirements gathering, and the treatment of research information that remains incomplete. No result from the ongoing survey or final system evaluation is reported until the corresponding responses, sampling information, and analysis have been completed and validated.")

heading(doc, "Materials")
body(doc, "The FloodSense implementation uses Flutter and Dart for the Android application; Python Django and Django REST Framework for backend services; GeoDjango for geographic operations; PostgreSQL with PostGIS for relational and spatial records; flutter_map for the resident-facing map; and Leaflet for the Web Administration System map. Git supports source management, while interface-design and document-processing software support design and research documentation. The final manuscript will identify only the development and testing tools actually used by the researchers.")
note(doc, "The verified specifications of the development computer and Android test device are not yet available; this information will be inserted after confirmation.")

heading(doc, "Method")
note(doc, "The software development life-cycle model and its corresponding diagram are not included because the development process and the model that accurately represents it have not yet been confirmed.")
body(doc, "The documented development activities begin with requirements gathering and evidence review. Survey summaries, interview transcripts, approved concept documents, geographic-data records, system specifications, source materials, and implementation evidence are examined to identify supported requirements and unresolved information. A feature that depends on unavailable data or an unverified relationship is not represented as completed system behavior.")

subheading(doc, "Requirements Analysis")
body(doc, "The researchers define the intended users, rainfall-scenario inputs, supported geographic areas, susceptibility outputs, preparedness functions, data sources, limitations, and administrative roles. Resident requirements and institutional information are compared with the implemented system so that a manuscript statement is retained only when it can be supported by a research record or the application.")

subheading(doc, "System Design")
body(doc, "The Android application, backend services, spatial database, Rule-Based Expert System, Decision Support System, and Web Administration System are organized as separate but connected components. The fixed inference procedure receives supported scenario and geographic facts, produces a susceptibility result, and passes that result as read-only context to the preparedness component. The administrative interface maintains only the permitted records and values required by its role.")
note(doc, "The system architecture, process, data, and interface diagrams will be inserted in their designated locations after the design documentation is finalized.")

subheading(doc, "CSV Parameter-Table Exchange")
body(doc, "The Web Administration System provides a controlled export-and-import process for values used by the fixed algorithm. The administrator first exports the canonical CSV table produced by FloodSense. The exported table contains the fixed fields and identifiers required by the system. An authorized editor may update only the values inside those existing fields. The editor must not rename, reorder, add, or remove columns; alter parameter identifiers or required keys; or change the structure of the table.")
body(doc, "After the value update, the administrator imports the CSV file into FloodSense. The system verifies that the file uses the expected schema before accepting the values. A file with missing, additional, renamed, or reordered required fields is rejected. The imported values are then read by the existing algorithm through the unchanged table structure. This procedure does not expose or import expert rules and does not modify, retrain, or replace the inference algorithm.")

subheading(doc, "Implementation and Verification")
body(doc, "Implementation connects the Android interface to the backend and spatial database through validated requests. Verification covers input validation, authentication and permissions, deterministic assessment behavior, geographic lookup, result consistency, map presentation, guidance separation, CSV schema validation, and safe handling of incomplete or unsupported data. The final list of tests and results will be reported only after execution records are complete.")

heading(doc, "Participants of the Study")
body(doc, "The completed requirements-gathering activities include an initial survey of 50 Bacoor City residents from 12 barangays and a separate supplementary survey of 40 respondents from 14 barangays. The two surveys represent different respondent groups and are therefore reported separately. Key informants include representatives of the Bacoor Disaster Risk Reduction and Management Office and the City Planning and Development Coordinator Office whose responsibilities are relevant to flood information, preparedness, institutional records, and geographic data.")
note(doc, "The new survey and final system evaluation are still being conducted. The approved sampling procedure, participant classifications, final number of valid respondents, response rate, and evaluation results will be inserted after completion and validation.")

heading(doc, "Statistical Treatment of Data")
body(doc, "Frequencies and percentages are used only for the completed descriptive requirements-survey items for which valid response counts are available. Percentage is computed as the frequency of a response divided by the number of valid responses for the item, multiplied by one hundred. The two requirements surveys are analyzed separately and are not combined into a single sample.")
note(doc, "The statistical treatment for the ongoing survey and final system evaluation has not yet been finalized. Scale scoring, verbal interpretations, sample-size computation, reliability analysis, and evaluation tables will be inserted only after the instrument, sampling procedure, completed responses, and analysis plan have been approved and validated.")


# REFERENCES
chapter(doc, "REFERENCES")
heading(doc, "Articles")
for item in [
    "GeoRisk Philippines. (2026). HazardHunterPH: Hazard assessment at your fingertips. Retrieved from https://hazardhunter.georisk.gov.ph/map",
    "International Organization for Standardization. (2023). ISO/IEC 25010:2023 systems and software engineering - Systems and software Quality Requirements and Evaluation (SQuaRE) - Product quality model. Retrieved from https://www.iso.org/standard/78176.html",
    "Japan International Cooperation Agency. (2021, October 1). Flood risk protection project in Cavite inaugurated. Retrieved from https://www.jica.go.jp/english/overseas/philippine/information/press/2021/211001.html",
    "Philippine Statistics Authority. (2026). Component 4: Extreme events and disasters. Retrieved from https://psa.gov.ph/statistics/environment-statistics/highlights/component-4-extreme-events-and-disaster",
]:
    reference(doc, item)

heading(doc, "Journals")
for item in [
    "Alabbad, Y., Yildirim, E., & Demir, I. (2022). Flood mitigation data analytics and decision support framework: Iowa Middle Cedar Watershed case study. Science of the Total Environment, 814, 152768. https://doi.org/10.1016/j.scitotenv.2021.152768",
    "Debnath, J., Sahariah, D., Nath, N., Saikia, A., Lahon, D., Islam, M. N., Hashimoto, S., Meraj, G., Kumar, P., Singh, S. K., Kanga, S., & Chand, K. (2024). Modelling on assessment of flood risk susceptibility at the Jia Bharali River basin in Eastern Himalayas by integrating multicollinearity tests and geospatial techniques. Modeling Earth Systems and Environment, 10, 2393-2419. https://doi.org/10.1007/s40808-023-01912-1",
    "Galanti, R., Coma-Puig, B., de Leoni, M., Carmona, J., & Navarin, N. (2023). An explainable decision support system for predictive process analytics. Engineering Applications of Artificial Intelligence, 120, 105904. https://doi.org/10.1016/j.engappai.2023.105904",
    "Johnson, B. A., Estoque, R. C., Li, X., Kumar, P., Dasgupta, R., Avtar, R., & Magcale-Macandog, D. B. (2021). High-resolution urban change modeling and flood exposure estimation at a national scale using open geospatial data: A case study of the Philippines. Computers, Environment and Urban Systems, 90, 101704. https://doi.org/10.1016/j.compenvurbsys.2021.101704",
    "Kostopoulos, G., Davrazos, G., & Kotsiantis, S. (2024). Explainable artificial intelligence-based decision support systems: A recent review. Electronics, 13(14), 2842. https://doi.org/10.3390/electronics13142842",
    "Lin, C. A. (2023). Flood risk management via risk communication, cognitive appraisal, collective efficacy, and community action. Sustainability, 15(19), 14191. https://doi.org/10.3390/su151914191",
    "Nagumo, N., Ohara, M., Fujikane, M., Inoue, T., Hiramatsu, Y., & Jaranilla-Sanchez, P. A. (2022). Creation of a 3D flood hazard map for a flood-prone area in the Republic of the Philippines and dissemination of the mapping technology. E-journal GEO, 17(1), 123-136. https://doi.org/10.4157/ejgeo.17.123",
    "Nguyen, H. N., Fukuda, H., & Nguyen, M. N. (2024). Assessment of the susceptibility of urban flooding using GIS with an analytical hierarchy process in Hanoi, Vietnam. Sustainability, 16(10), 3934. https://doi.org/10.3390/su16103934",
    "Oubennaceur, K., Chokmani, K., El Alem, A., & Gauthier, Y. (2021). Flood risk communication using ArcGIS StoryMaps. Hydrology, 8(4), 152. https://doi.org/10.3390/hydrology8040152",
]:
    reference(doc, item)

heading(doc, "Research Records")
for item in [
    "FloodSense Research Team. (2026a). FloodSense first resident requirements survey summary [Unpublished research record]. Cavite State University - Bacoor City Campus.",
    "FloodSense Research Team. (2026b). FloodSense supplementary resident requirements survey analysis [Unpublished research record]. Cavite State University - Bacoor City Campus.",
    "FloodSense Research Team. (2026c). First key informant interview transcript with the Bacoor Disaster Risk Reduction and Management Office [Unpublished research record]. Cavite State University - Bacoor City Campus.",
    "FloodSense Research Team. (2026d). Second key informant interview transcript with the Bacoor Disaster Risk Reduction and Management Office [Unpublished research record]. Cavite State University - Bacoor City Campus.",
    "FloodSense Research Team. (2026e). Key informant interview transcript with the Bacoor City Planning and Development Coordinator Office [Unpublished research record]. Cavite State University - Bacoor City Campus.",
]:
    reference(doc, item)

# Compact table cell margins and apply Arial to all generated table content.
for tbl in doc.tables:
    for row in tbl.rows:
        for cell in row.cells:
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_mar = tc_pr.first_child_found_in("w:tcMar")
            if tc_mar is None:
                tc_mar = OxmlElement("w:tcMar")
                tc_pr.append(tc_mar)
            for edge in ("top", "left", "bottom", "right"):
                node = tc_mar.find(qn(f"w:{edge}"))
                if node is None:
                    node = OxmlElement(f"w:{edge}")
                    tc_mar.append(node)
                node.set(qn("w:w"), "80")
                node.set(qn("w:type"), "dxa")

OUT.parent.mkdir(parents=True, exist_ok=True)
doc.save(OUT)
print(OUT)
