"""Render a provisional Bacoor bare-earth DTM from checked-in boundaries and FABDEM.

This is a presentation reference only. It deliberately keeps the administrative
boundary provenance and provisional terrain provenance visible in the image.
"""

from __future__ import annotations

import io
import json
import math
import urllib.request
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
BOUNDARY_DIR = ROOT / "research_data" / "administrative_boundaries" / "processed"
CITY_PATH = BOUNDARY_DIR / "bacoor_city_boundary.geojson"
BARANGAY_PATH = BOUNDARY_DIR / "bacoor_barangay_boundaries.geojson"
OUTPUT_PATH = ROOT / "frontend-design-output" / "bacoor-provisional-bare-earth-dtm.png"

FABDEM_TILE_SIZE = 3600
FABDEM_URL = (
    "https://huggingface.co/buckets/links-ads/fabdem/resolve/tiles/"
    "N10E120-N20E130_FABDEM_V1-2/N14E{longitude:03d}_FABDEM_V1-2.tif"
)
CANVAS_W, CANVAS_H = 2048, 1536


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def iter_polygons(geometry: dict):
    if geometry["type"] == "Polygon":
        yield geometry["coordinates"]
    elif geometry["type"] == "MultiPolygon":
        yield from geometry["coordinates"]
    else:
        raise ValueError(f"Unsupported geometry: {geometry['type']}")


def iter_points(geometry: dict):
    for polygon in iter_polygons(geometry):
        for ring in polygon:
            yield from ring


def fetch_fabdem_tile(longitude: int) -> np.ndarray:
    url = FABDEM_URL.format(longitude=longitude)
    request = urllib.request.Request(url, headers={"User-Agent": "FloodSense academic DTM renderer"})
    with urllib.request.urlopen(request, timeout=45) as response:
        tile = np.asarray(Image.open(io.BytesIO(response.read())), dtype=np.float32).copy()
    tile[tile <= -9990] = np.nan
    return tile


def gaussian_blur(array: np.ndarray, sigma: float = 1.35) -> np.ndarray:
    radius = max(1, int(math.ceil(sigma * 3)))
    x = np.arange(-radius, radius + 1, dtype=np.float32)
    kernel = np.exp(-(x * x) / (2 * sigma * sigma))
    kernel /= kernel.sum()
    filled = np.where(np.isfinite(array), array, np.nanmedian(array))
    horizontal = np.apply_along_axis(lambda row: np.convolve(row, kernel, mode="same"), 1, filled)
    return np.apply_along_axis(lambda col: np.convolve(col, kernel, mode="same"), 0, horizontal)


def colorize_elevation(elevation: np.ndarray) -> np.ndarray:
    stops = np.array([-5, 0, 5, 10, 20, 30, 40, 50, 70, 100], dtype=np.float32)
    colors = np.array(
        [
            [18, 86, 151],
            [22, 137, 177],
            [19, 181, 184],
            [26, 198, 154],
            [72, 203, 103],
            [142, 212, 76],
            [213, 220, 85],
            [232, 199, 100],
            [205, 151, 80],
            [153, 107, 67],
        ],
        dtype=np.float32,
    )
    out = np.empty((*elevation.shape, 3), dtype=np.float32)
    for channel in range(3):
        out[..., channel] = np.interp(elevation, stops, colors[:, channel])

    # A strong multidirectional-looking hillshade matches the supplied DTM reference style.
    dy, dx = np.gradient(elevation, 30.0, 30.0)
    slope = np.arctan(np.hypot(dx, dy))
    aspect = np.arctan2(-dx, dy)
    altitude = math.radians(43)
    azimuth = math.radians(315)
    shade = (
        math.sin(altitude) * np.cos(slope)
        + math.cos(altitude) * np.sin(slope) * np.cos(azimuth - aspect)
    )
    shade = np.clip((shade + 1.0) / 2.0, 0.0, 1.0)
    out *= (0.66 + 0.58 * shade[..., None])
    return np.clip(out, 0, 255).astype(np.uint8)


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = [
        Path("C:/Windows/Fonts/seguisb.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf"),
        Path("C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size=size)
    return ImageFont.load_default()


def edge_point(edge: int, c: int, r: int, vals: tuple[float, ...], level: float):
    v0, v1, v2, v3 = vals
    endpoints = {
        0: ((c, r, v0), (c + 1, r, v1)),
        1: ((c + 1, r, v1), (c + 1, r + 1, v2)),
        2: ((c + 1, r + 1, v2), (c, r + 1, v3)),
        3: ((c, r + 1, v3), (c, r, v0)),
    }
    a, b = endpoints[edge]
    denominator = b[2] - a[2]
    t = 0.5 if abs(denominator) < 1e-9 else (level - a[2]) / denominator
    t = min(max(t, 0.0), 1.0)
    return a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t


def marching_segments(grid: np.ndarray, level: float, stride: int = 2):
    sampled = grid[::stride, ::stride]
    table = {
        1: [(3, 0)], 2: [(0, 1)], 3: [(3, 1)], 4: [(1, 2)],
        5: [(3, 0), (1, 2)], 6: [(0, 2)], 7: [(3, 2)],
        8: [(2, 3)], 9: [(0, 2)], 10: [(0, 1), (2, 3)],
        11: [(1, 2)], 12: [(3, 1)], 13: [(0, 1)], 14: [(3, 0)],
    }
    segments = []
    rows, cols = sampled.shape
    for r in range(rows - 1):
        for c in range(cols - 1):
            vals = (
                float(sampled[r, c]),
                float(sampled[r, c + 1]),
                float(sampled[r + 1, c + 1]),
                float(sampled[r + 1, c]),
            )
            case = sum((value >= level) << bit for bit, value in enumerate(vals))
            for edge_a, edge_b in table.get(case, []):
                a = edge_point(edge_a, c, r, vals, level)
                b = edge_point(edge_b, c, r, vals, level)
                segments.append(
                    ((a[0] * stride, a[1] * stride), (b[0] * stride, b[1] * stride))
                )
    return segments


def draw_geometry(
    draw: ImageDraw.ImageDraw,
    geometry: dict,
    project,
    *,
    fill=None,
    outline=None,
    width: int = 1,
):
    for polygon in iter_polygons(geometry):
        outer = [project(lon, lat) for lon, lat in polygon[0]]
        if fill is not None:
            draw.polygon(outer, fill=fill)
        if outline is not None:
            draw.line(outer + [outer[0]], fill=outline, width=width, joint="curve")
        if fill is not None:
            for hole in polygon[1:]:
                draw.polygon([project(lon, lat) for lon, lat in hole], fill=0)


def text_with_halo(draw, xy, text, text_font, fill, anchor="mm", halo=3):
    draw.text(xy, text, font=text_font, fill=fill, anchor=anchor,
              stroke_width=halo, stroke_fill=(255, 255, 255, 235))


def main() -> None:
    city_fc = load_json(CITY_PATH)
    barangay_fc = load_json(BARANGAY_PATH)
    city_geometry = city_fc["features"][0]["geometry"]
    points = list(iter_points(city_geometry))
    longitudes = [p[0] for p in points]
    latitudes = [p[1] for p in points]
    min_lon, max_lon = min(longitudes), max(longitudes)
    min_lat, max_lat = min(latitudes), max(latitudes)

    # Bacoor crosses 121°E, so join the two adjacent one-degree FABDEM tiles.
    tile_lon0 = int(math.floor(min_lon))
    tile_lon1 = int(math.floor(max_lon))
    tiles = [fetch_fabdem_tile(tile_lon) for tile_lon in range(tile_lon0, tile_lon1 + 1)]
    mosaic = np.concatenate(tiles, axis=1)
    mosaic_west = float(tile_lon0)
    mosaic_north = 15.0

    pad = 26
    crop_x0 = max(0, int(math.floor((min_lon - mosaic_west) * FABDEM_TILE_SIZE)) - pad)
    crop_x1 = min(mosaic.shape[1], int(math.ceil((max_lon - mosaic_west) * FABDEM_TILE_SIZE)) + pad)
    crop_y0 = max(0, int(math.floor((mosaic_north - max_lat) * FABDEM_TILE_SIZE)) - pad)
    crop_y1 = min(mosaic.shape[0], int(math.ceil((mosaic_north - min_lat) * FABDEM_TILE_SIZE)) + pad)
    raw_elevation = mosaic[crop_y0:crop_y1, crop_x0:crop_x1]
    elevation = gaussian_blur(raw_elevation, sigma=1.35)

    canvas = Image.new("RGB", (CANVAS_W, CANVAS_H), (242, 246, 248))
    draw = ImageDraw.Draw(canvas, "RGBA")

    # Header
    draw.rectangle((0, 0, CANVAS_W, 166), fill=(17, 53, 88, 255))
    draw.text((72, 35), "BACOOR CITY — PROVISIONAL BARE-EARTH DTM",
              font=font(50, True), fill="white")
    draw.text((74, 104),
              "FABDEM 30 m terrain • Buildings and forest-height biases removed • Awaiting Phil-LiDAR DTM",
              font=font(25), fill=(210, 228, 241))

    # Map placement corrects longitude spacing by latitude while preserving the real outline.
    map_region = (50, 190, 1550, 1360)
    max_width = map_region[2] - map_region[0]
    max_height = map_region[3] - map_region[1]
    center_lat = float(city_fc["features"][0]["properties"]["center_lat"])
    longitude_factor = math.cos(math.radians(center_lat))
    scale_y = min(
        max_height / elevation.shape[0],
        max_width / (elevation.shape[1] * longitude_factor),
    )
    scale_x = scale_y * longitude_factor
    display_w = int(round(elevation.shape[1] * scale_x))
    display_h = int(round(elevation.shape[0] * scale_y))
    map_x = map_region[0] + (max_width - display_w) // 2
    map_y = map_region[1] + (max_height - display_h) // 2

    def project(lon: float, lat: float):
        source_x = (lon - mosaic_west) * FABDEM_TILE_SIZE - crop_x0
        source_y = (mosaic_north - lat) * FABDEM_TILE_SIZE - crop_y0
        return source_x * scale_x, source_y * scale_y

    terrain_rgb = colorize_elevation(elevation)
    terrain = Image.fromarray(terrain_rgb, mode="RGB").resize(
        (display_w, display_h), Image.Resampling.BICUBIC
    )
    mask = Image.new("L", (display_w, display_h), 0)
    mask_draw = ImageDraw.Draw(mask)
    draw_geometry(mask_draw, city_geometry, project, fill=255)
    canvas.paste(terrain, (map_x, map_y), mask)

    # Convert local map coordinates to canvas coordinates for overlays.
    def project_canvas(lon: float, lat: float):
        x, y = project(lon, lat)
        return x + map_x, y + map_y

    # Contours calculated from the real provisional DEM, clipped to the city.
    contour_layer = Image.new("RGBA", (display_w, display_h), (0, 0, 0, 0))
    contour_draw = ImageDraw.Draw(contour_layer, "RGBA")
    contour_font = font(21, True)
    min_level = max(0, int(math.ceil(float(np.nanmin(elevation)) / 10.0) * 10))
    max_level = int(math.floor(float(np.nanmax(elevation)) / 10.0) * 10)
    label_positions = []
    for level in range(min_level, max_level + 1, 10):
        segments = marching_segments(elevation, level, stride=2)
        candidates = []
        for (ax, ay), (bx, by) in segments:
            line = (ax * scale_x, ay * scale_y, bx * scale_x, by * scale_y)
            contour_draw.line(line, fill=(42, 59, 53, 165), width=2)
            mx, my = (line[0] + line[2]) / 2, (line[1] + line[3]) / 2
            if 20 < mx < display_w - 20 and 20 < my < display_h - 20:
                safe_box = (
                    max(0, int(mx) - 58), max(0, int(my) - 24),
                    min(display_w, int(mx) + 58), min(display_h, int(my) + 24),
                )
                # Keep the full label inside Bacoor so the city clip cannot cut off digits.
                if mask.crop(safe_box).getextrema()[0] > 0:
                    score = abs(mx - display_w * 0.53) + abs(my - display_h * 0.50)
                    candidates.append((score, mx, my))
        candidates.sort()
        for _, mx, my in candidates:
            if all((mx - ox) ** 2 + (my - oy) ** 2 > 120**2 for ox, oy in label_positions):
                text_with_halo(contour_draw, (mx, my), f"{level} m", contour_font,
                               (55, 43, 31, 255), halo=4)
                label_positions.append((mx, my))
                break
    contour_layer.putalpha(Image.composite(contour_layer.getchannel("A"), Image.new("L", mask.size, 0), mask))
    canvas.alpha_composite(contour_layer, (map_x, map_y)) if canvas.mode == "RGBA" else canvas.paste(
        contour_layer, (map_x, map_y), contour_layer
    )

    # Real 47-barangay reference boundaries, kept light so bare-earth relief remains dominant.
    boundary_layer = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    boundary_draw = ImageDraw.Draw(boundary_layer, "RGBA")
    for feature in barangay_fc["features"]:
        draw_geometry(boundary_draw, feature["geometry"], project_canvas,
                      outline=(7, 38, 39, 120), width=3)
    for feature in barangay_fc["features"]:
        draw_geometry(boundary_draw, feature["geometry"], project_canvas,
                      outline=(255, 255, 255, 205), width=2)
    draw_geometry(boundary_draw, city_geometry, project_canvas,
                  outline=(10, 30, 42, 255), width=7)
    canvas.paste(boundary_layer, (0, 0), boundary_layer)
    draw = ImageDraw.Draw(canvas, "RGBA")

    # Two true feature-center labels relevant to the team's scenario discussion.
    selected_names = {"Niog", "Molino III"}
    label_font = font(24, True)
    for feature in barangay_fc["features"]:
        props = feature["properties"]
        if props.get("adm4_name") not in selected_names:
            continue
        x, y = project_canvas(float(props["center_lon"]), float(props["center_lat"]))
        label = props["adm4_name"]
        box = draw.textbbox((x, y), label, font=label_font, anchor="mm", stroke_width=3)
        draw.rounded_rectangle((box[0] - 9, box[1] - 5, box[2] + 9, box[3] + 5),
                               radius=8, fill=(255, 255, 255, 220), outline=(17, 53, 88, 180), width=2)
        draw.text((x, y), label, font=label_font, anchor="mm", fill=(17, 45, 64),
                  stroke_width=1, stroke_fill="white")

    # North arrow.
    nx, ny = 1482, 236
    draw.text((nx, ny - 26), "N", font=font(28, True), anchor="ms", fill=(17, 45, 64))
    draw.polygon([(nx, ny), (nx - 18, ny + 66), (nx, ny + 51), (nx + 18, ny + 66)],
                 fill=(17, 45, 64), outline="white")

    # Scale bar calculated at the city center.
    meters_per_native_pixel = 111_320.0 * longitude_factor / FABDEM_TILE_SIZE
    meters_per_display_pixel = meters_per_native_pixel / scale_x
    segment_px = int(round(1000 / meters_per_display_pixel))
    sx, sy = 92, 1321
    for index in range(2):
        fill = (17, 45, 64) if index % 2 == 0 else (255, 255, 255)
        draw.rectangle((sx + index * segment_px, sy, sx + (index + 1) * segment_px, sy + 16),
                       fill=fill, outline=(17, 45, 64), width=2)
    for index, label in enumerate(("0", "1", "2 km")):
        draw.text((sx + index * segment_px, sy - 7), label, font=font(18, True),
                  anchor="ms", fill=(17, 45, 64))

    # Legend card.
    legend = (1590, 220, 1992, 1085)
    draw.rounded_rectangle(legend, radius=22, fill=(255, 255, 255, 244),
                           outline=(183, 197, 205), width=2)
    draw.text((1624, 260), "BARE-EARTH DTM", font=font(30, True), fill=(17, 53, 88))
    draw.text((1624, 306), "Elevation above mean\nsea level (metres)",
              font=font(22), fill=(54, 72, 82), spacing=6)

    legend_min = 0
    legend_max = max(10, int(math.ceil(float(np.nanmax(elevation)) / 10.0) * 10))
    bar_x0, bar_y0, bar_x1, bar_y1 = 1640, 390, 1730, 955
    sample_values = np.linspace(legend_max, legend_min, bar_y1 - bar_y0).reshape(-1, 1)
    # Two identical columns keep the shared colorizer's gradient calculation valid.
    legend_rgb = colorize_elevation(np.repeat(sample_values, 2, axis=1))
    legend_img = Image.fromarray(legend_rgb, mode="RGB").resize((bar_x1 - bar_x0, bar_y1 - bar_y0))
    canvas.paste(legend_img, (bar_x0, bar_y0))
    draw.rectangle((bar_x0, bar_y0, bar_x1, bar_y1), outline=(55, 69, 76), width=2)
    for value in range(legend_min, legend_max + 1, 10):
        y = bar_y1 - (value - legend_min) / max(legend_max - legend_min, 1) * (bar_y1 - bar_y0)
        draw.line((bar_x1, y, bar_x1 + 14, y), fill=(55, 69, 76), width=2)
        draw.text((bar_x1 + 24, y), f"{value} m", font=font(20), anchor="lm", fill=(37, 54, 63))

    draw.text((1624, 1003), "47 barangay reference polygons",
              font=font(19, True), fill=(17, 53, 88))
    draw.text((1624, 1035), "White lines: barangay boundaries\nDark lines: 10 m terrain contours",
              font=font(18), fill=(67, 82, 90), spacing=5)

    # Source and warning footer.
    draw.rounded_rectangle((70, 1390, 1978, 1503), radius=18,
                           fill=(255, 246, 230), outline=(214, 139, 40), width=2)
    draw.text((96, 1412), "TEMPORARY REFERENCE — NOT FOR ENGINEERING OR OFFICIAL HAZARD DECISIONS",
              font=font(25, True), fill=(143, 68, 16))
    draw.text((96, 1454),
              "Boundaries: OCHA COD-AB / NAMRIA / PSA-derived reference, not City-verified.  "
              "Terrain: FABDEM v1.2 (forest and building height biases removed), 30 m, CC BY-NC-SA 4.0.",
              font=font(16), fill=(84, 72, 61))

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(OUTPUT_PATH, format="PNG", optimize=True)
    inside_values = []
    terrain_mask = np.asarray(mask.resize((elevation.shape[1], elevation.shape[0]), Image.Resampling.NEAREST)) > 0
    if terrain_mask.shape == elevation.shape:
        inside_values = elevation[terrain_mask]
    values = inside_values if len(inside_values) else elevation.ravel()
    print(f"Saved: {OUTPUT_PATH}")
    print(f"Boundary features: {len(barangay_fc['features'])}")
    print(f"Provisional elevation range in city: {float(np.min(values)):.1f}–{float(np.max(values)):.1f} m")


if __name__ == "__main__":
    main()
