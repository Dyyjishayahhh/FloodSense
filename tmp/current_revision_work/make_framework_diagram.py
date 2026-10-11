from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


OUT = Path(r"tmp\current_revision_work\FloodSense_Theoretical_Framework.png")
OUT.parent.mkdir(parents=True, exist_ok=True)

W, H = 1800, 1420
img = Image.new("RGB", (W, H), "white")
draw = ImageDraw.Draw(img)

regular = ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf", 25)
small = ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf", 21)
bold = ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf", 29)
title_font = ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf", 43)


def wrapped(text: str, font, max_width: int) -> list[str]:
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        trial = word if not current else current + " " + word
        if draw.textbbox((0, 0), trial, font=font)[2] <= max_width:
            current = trial
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def box(x1, y1, x2, y2, heading: str, body: str, fill="#f5f5f5"):
    draw.rounded_rectangle((x1, y1, x2, y2), radius=18, fill=fill, outline="black", width=3)
    h_lines = wrapped(heading, bold, x2 - x1 - 36)
    y = y1 + 20
    for line in h_lines:
        bbox = draw.textbbox((0, 0), line, font=bold)
        draw.text(((x1 + x2 - (bbox[2] - bbox[0])) / 2, y), line, font=bold, fill="black")
        y += 36
    y += 7
    for line in wrapped(body, small, x2 - x1 - 42):
        bbox = draw.textbbox((0, 0), line, font=small)
        draw.text(((x1 + x2 - (bbox[2] - bbox[0])) / 2, y), line, font=small, fill="black")
        y += 28


def arrow(x1, y1, x2, y2):
    draw.line((x1, y1, x2, y2), fill="black", width=5)
    import math
    angle = math.atan2(y2 - y1, x2 - x1)
    size = 20
    for offset in (2.55, -2.55):
        ax = x2 + size * math.cos(angle + offset)
        ay = y2 + size * math.sin(angle + offset)
        draw.line((x2, y2, ax, ay), fill="black", width=5)


def poly_arrow(points):
    draw.line(points, fill="black", width=5, joint="curve")
    x1, y1 = points[-2]
    x2, y2 = points[-1]
    import math
    angle = math.atan2(y2 - y1, x2 - x1)
    size = 20
    for offset in (2.55, -2.55):
        ax = x2 + size * math.cos(angle + offset)
        ay = y2 + size * math.sin(angle + offset)
        draw.line((x2, y2, ax, ay), fill="black", width=5)


title = "FloodSense Theoretical Framework and System Architecture"
tb = draw.textbbox((0, 0), title, font=title_font)
draw.text(((W - (tb[2] - tb[0])) / 2, 28), title, font=title_font, fill="black")

box(45, 180, 315, 390, "Resident", "Uses the Android application and explicitly requests an assessment.", "#e8eef7")
box(1485, 180, 1755, 390, "Authorized Administrator", "Uses the custom Web Administration System under assigned permissions.", "#e8eef7")

modules = [
    (365, 125, 780, 350, "1. User Access and Authentication Module", "Registration, sign-in, verification, recovery, privacy acceptance, profiles, roles, and permissions."),
    (1020, 125, 1435, 350, "2. Scenario and Location Input Module", "Rainfall intensity and duration, manual barangay or pin, optional one-time GPS, and confirmation."),
    (365, 430, 780, 700, "3. Susceptibility Assessment and Explanation Module", "Input validation, deterministic rule matching, conflict handling, susceptibility class, explanation, and map presentation."),
    (1020, 430, 1435, 700, "4. Preparedness Decision Support Module", "Source-backed preparedness guidance, structured DSS flow, and evacuation-center information kept separate from assessment."),
    (690, 795, 1110, 1070, "5. Administration and Data Governance Module", "Controlled management of references, parameters, DSS content, centers, sources, review records, permissions, and audit history."),
]
for coords in modules:
    box(*coords)

box(360, 1190, 1440, 1365, "Supporting Knowledge and Geographic Foundation", "Controlled rules, rainfall references, administrative geography, provenance records, preparedness content, evacuation-center records, and PostgreSQL/PostGIS storage.", "#ededed")

arrow(315, 285, 365, 230)
poly_arrow([(1485, 300), (1460, 760), (1110, 890)])
arrow(780, 230, 1020, 230)
arrow(570, 350, 570, 430)
arrow(1230, 350, 1230, 430)
arrow(780, 565, 1020, 565)
arrow(900, 1070, 900, 1190)

note = "Outputs: explained susceptibility result and map context | separate preparedness guidance and center information"
nb = draw.textbbox((0, 0), note, font=regular)
draw.text(((W - (nb[2] - nb[0])) / 2, 1115), note, font=regular, fill="black")

img.save(OUT, dpi=(300, 300))
print(OUT.resolve())
