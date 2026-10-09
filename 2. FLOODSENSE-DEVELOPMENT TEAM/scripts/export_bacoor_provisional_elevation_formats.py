"""Export the provisional Bacoor FABDEM terrain to GeoTIFF, XYZ, and LAS.

The exports use the checked-in Bacoor city boundary and FABDEM v1.2 bare-earth
elevations. They are temporary research files, not Phil-LiDAR deliverables.
"""

from __future__ import annotations

import datetime as dt
import math
import struct
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, TiffImagePlugin, TiffTags

from render_bacoor_provisional_dtm import (
    CITY_PATH,
    FABDEM_TILE_SIZE,
    ROOT,
    fetch_fabdem_tile,
    iter_points,
    iter_polygons,
    load_json,
)


OUTPUT_DIR = ROOT / "frontend-design-output" / "MAP GEO"
BASE_NAME = "bacoor_provisional_fabdem_dtm_30m"
GEOTIFF_PATH = OUTPUT_DIR / f"{BASE_NAME}.tif"
XYZ_PATH = OUTPUT_DIR / f"{BASE_NAME}.xyz"
LAS_PATH = OUTPUT_DIR / f"{BASE_NAME}.las"

NODATA = -9999.0
PIXEL_SIZE_DEGREES = 1.0 / FABDEM_TILE_SIZE
HORIZONTAL_EPSG = 4326  # WGS 84
VERTICAL_EPSG = 3855  # EGM2008 height


def load_clipped_dtm():
    city_fc = load_json(CITY_PATH)
    geometry = city_fc["features"][0]["geometry"]
    points = list(iter_points(geometry))
    min_lon = min(point[0] for point in points)
    max_lon = max(point[0] for point in points)
    min_lat = min(point[1] for point in points)
    max_lat = max(point[1] for point in points)

    tile_lon0 = int(math.floor(min_lon))
    tile_lon1 = int(math.floor(max_lon))
    tiles = [fetch_fabdem_tile(tile_lon) for tile_lon in range(tile_lon0, tile_lon1 + 1)]
    mosaic = np.concatenate(tiles, axis=1)
    mosaic_west = float(tile_lon0)
    mosaic_north = 15.0

    col0 = max(0, int(math.floor((min_lon - mosaic_west) * FABDEM_TILE_SIZE)))
    col1 = min(mosaic.shape[1], int(math.ceil((max_lon - mosaic_west) * FABDEM_TILE_SIZE)))
    row0 = max(0, int(math.floor((mosaic_north - max_lat) * FABDEM_TILE_SIZE)))
    row1 = min(mosaic.shape[0], int(math.ceil((mosaic_north - min_lat) * FABDEM_TILE_SIZE)))

    elevation = mosaic[row0:row1, col0:col1].astype(np.float32, copy=True)
    west = mosaic_west + col0 * PIXEL_SIZE_DEGREES
    north = mosaic_north - row0 * PIXEL_SIZE_DEGREES

    mask_image = Image.new("1", (elevation.shape[1], elevation.shape[0]), 0)
    mask_draw = ImageDraw.Draw(mask_image)

    def project(lon: float, lat: float):
        return (
            (lon - west) / PIXEL_SIZE_DEGREES,
            (north - lat) / PIXEL_SIZE_DEGREES,
        )

    for polygon in iter_polygons(geometry):
        outer = [project(lon, lat) for lon, lat in polygon[0]]
        mask_draw.polygon(outer, fill=1)
        for hole in polygon[1:]:
            mask_draw.polygon([project(lon, lat) for lon, lat in hole], fill=0)

    mask = np.asarray(mask_image, dtype=bool)
    valid = mask & np.isfinite(elevation) & (elevation > -9990)
    clipped = np.where(valid, elevation, NODATA).astype(np.float32)
    return clipped, valid, west, north


def write_geotiff(path: Path, elevation: np.ndarray, west: float, north: float) -> None:
    # GeoTIFF keys: geographic WGS 84, pixel-is-area, angular degrees, EGM2008 height.
    ascii_params = "WGS 84|EGM2008 height|"
    wgs_count = len("WGS 84|")
    vertical_offset = wgs_count
    vertical_count = len("EGM2008 height|")
    keys = (
        1, 1, 0, 7,
        1024, 0, 1, 2,
        1025, 0, 1, 1,
        2048, 0, 1, HORIZONTAL_EPSG,
        2049, 34737, wgs_count, 0,
        2054, 0, 1, 9102,
        4096, 0, 1, VERTICAL_EPSG,
        4097, 34737, vertical_count, vertical_offset,
    )

    ifd = TiffImagePlugin.ImageFileDirectory_v2()
    ifd.tagtype[33550] = TiffTags.DOUBLE
    ifd.tagtype[33922] = TiffTags.DOUBLE
    ifd.tagtype[34735] = TiffTags.SHORT
    ifd.tagtype[34737] = TiffTags.ASCII
    ifd.tagtype[42113] = TiffTags.ASCII
    ifd[33550] = (PIXEL_SIZE_DEGREES, PIXEL_SIZE_DEGREES, 0.0)
    ifd[33922] = (0.0, 0.0, 0.0, west, north, 0.0)
    ifd[34735] = keys
    ifd[34737] = ascii_params
    ifd[42113] = str(int(NODATA))

    image = Image.fromarray(elevation, mode="F")
    image.save(path, format="TIFF", compression="tiff_deflate", tiffinfo=ifd)


def valid_coordinates(elevation: np.ndarray, valid: np.ndarray, west: float, north: float):
    rows, cols = np.nonzero(valid)
    longitudes = west + (cols.astype(np.float64) + 0.5) * PIXEL_SIZE_DEGREES
    latitudes = north - (rows.astype(np.float64) + 0.5) * PIXEL_SIZE_DEGREES
    heights = elevation[rows, cols].astype(np.float64)
    return longitudes, latitudes, heights


def write_xyz(path: Path, longitudes, latitudes, heights) -> None:
    with path.open("w", encoding="ascii", newline="\n") as output:
        for lon, lat, height in zip(longitudes, latitudes, heights):
            output.write(f"{lon:.8f} {lat:.8f} {height:.3f}\n")


def fixed_ascii(value: str, length: int) -> bytes:
    encoded = value.encode("ascii", errors="replace")[:length]
    return encoded.ljust(length, b"\0")


def write_las(path: Path, longitudes, latitudes, heights) -> None:
    point_count = len(heights)
    if not point_count:
        raise ValueError("Cannot write an empty LAS file")

    x_scale, y_scale, z_scale = 1e-7, 1e-7, 0.001
    x_offset = math.floor(float(np.min(longitudes)))
    y_offset = math.floor(float(np.min(latitudes)))
    z_offset = 0.0

    # LAS 1.2 GeoTIFF projection VLR for WGS 84 + EGM2008 height.
    geokeys = (
        1, 1, 0, 4,
        1024, 0, 1, 2,
        2048, 0, 1, HORIZONTAL_EPSG,
        2054, 0, 1, 9102,
        4096, 0, 1, VERTICAL_EPSG,
    )
    geokey_payload = struct.pack("<" + "H" * len(geokeys), *geokeys)
    vlr_header = struct.pack(
        "<H16sHH32s",
        0,
        fixed_ascii("LASF_Projection", 16),
        34735,
        len(geokey_payload),
        fixed_ascii("GeoKeyDirectoryTag", 32),
    )
    vlr = vlr_header + geokey_payload

    header_size = 227
    point_offset = header_size + len(vlr)
    header = bytearray(header_size)
    header[0:4] = b"LASF"
    struct.pack_into("<H", header, 4, 0)  # File Source ID
    struct.pack_into("<H", header, 6, 0)  # Global Encoding
    struct.pack_into("<IHH8s", header, 8, 0, 0, 0, b"\0" * 8)
    struct.pack_into("<BB", header, 24, 1, 2)
    header[26:58] = fixed_ascii("FABDEM V1.2", 32)
    header[58:90] = fixed_ascii("FloodSense provisional export", 32)
    today = dt.date.today()
    struct.pack_into("<HH", header, 90, today.timetuple().tm_yday, today.year)
    struct.pack_into("<H", header, 94, header_size)
    struct.pack_into("<I", header, 96, point_offset)
    struct.pack_into("<I", header, 100, 1)  # One VLR
    struct.pack_into("<B", header, 104, 0)  # Point Data Record Format 0
    struct.pack_into("<H", header, 105, 20)
    struct.pack_into("<I", header, 107, point_count)
    struct.pack_into("<5I", header, 111, point_count, 0, 0, 0, 0)
    struct.pack_into("<3d", header, 131, x_scale, y_scale, z_scale)
    struct.pack_into("<3d", header, 155, x_offset, y_offset, z_offset)
    struct.pack_into(
        "<6d",
        header,
        179,
        float(np.max(longitudes)), float(np.min(longitudes)),
        float(np.max(latitudes)), float(np.min(latitudes)),
        float(np.max(heights)), float(np.min(heights)),
    )

    with path.open("wb") as output:
        output.write(header)
        output.write(vlr)
        for lon, lat, height in zip(longitudes, latitudes, heights):
            x = int(round((float(lon) - x_offset) / x_scale))
            y = int(round((float(lat) - y_offset) / y_scale))
            z = int(round((float(height) - z_offset) / z_scale))
            # Return 1 of 1; ASPRS classification 2 = ground.
            output.write(struct.pack("<iiiHBBbBH", x, y, z, 0, 9, 2, 0, 0, 0))


def validate_outputs(expected_points: int) -> None:
    with Image.open(GEOTIFF_PATH) as image:
        if image.mode != "F" or image.size[0] <= 0 or image.size[1] <= 0:
            raise ValueError("Invalid floating-point GeoTIFF")
        required_tags = (33550, 33922, 34735, 42113)
        missing = [tag for tag in required_tags if tag not in image.tag_v2]
        if missing:
            raise ValueError(f"GeoTIFF is missing tags: {missing}")

    with XYZ_PATH.open("r", encoding="ascii") as source:
        xyz_count = sum(1 for _ in source)
    if xyz_count != expected_points:
        raise ValueError(f"XYZ count mismatch: {xyz_count} != {expected_points}")

    with LAS_PATH.open("rb") as source:
        header = source.read(227)
    if header[:4] != b"LASF":
        raise ValueError("Invalid LAS signature")
    las_count = struct.unpack_from("<I", header, 107)[0]
    if las_count != expected_points:
        raise ValueError(f"LAS count mismatch: {las_count} != {expected_points}")


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    elevation, valid, west, north = load_clipped_dtm()
    longitudes, latitudes, heights = valid_coordinates(elevation, valid, west, north)

    write_geotiff(GEOTIFF_PATH, elevation, west, north)
    write_xyz(XYZ_PATH, longitudes, latitudes, heights)
    write_las(LAS_PATH, longitudes, latitudes, heights)
    validate_outputs(len(heights))

    print(f"GeoTIFF: {GEOTIFF_PATH}")
    print(f"XYZ: {XYZ_PATH}")
    print(f"LAS: {LAS_PATH}")
    print(f"Raster: {elevation.shape[1]} x {elevation.shape[0]} cells")
    print(f"Valid ground points: {len(heights)}")
    print(f"Elevation range: {float(np.min(heights)):.3f} to {float(np.max(heights)):.3f} m")


if __name__ == "__main__":
    main()
