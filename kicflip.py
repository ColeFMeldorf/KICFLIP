#!/usr/bin/env python3
"""
heic_grid_to_pptx.py

Take a folder of .heic photos and lay them out in a grid on a single
PowerPoint slide.

Usage:
    python heic_grid_to_pptx.py /path/to/folder [-o output.pptx]

What this script needs installed (one-time setup):
    pip install pillow-heif python-pptx pillow

"""

import argparse
import io
import math
from pathlib import Path

from PIL import Image
from pillow_heif import register_heif_opener
from pptx import Presentation
from pptx.util import Inches

register_heif_opener()


def find_heic_files(folder: Path) -> list[Path]:
    """
    Return every .heic file directly inside `folder`, sorted by name.

    We check both lowercase ".heic" and uppercase ".HEIC" because file
    extensions on some systems (notably ones synced from an iPhone) are
    not consistently lowercase, and Python's glob matching is
    case-sensitive on Linux/Mac.
    """
    matches = list(folder.glob("*.heic")) + list(folder.glob("*.HEIC"))
    # De-duplicate in case a filesystem is case-insensitive and a file
    # would otherwise be matched twice, then sort for a stable order.
    unique_matches = sorted(set(matches), key=lambda p: p.name.lower())
    return unique_matches


def choose_grid_shape(n: int) -> tuple[int, int]:
    """
    Given `n` images, decide how many rows and columns to use so the
    grid is as close to square as possible, while never leaving an
    entirely empty row.

    Returns (rows, cols).
    """
    if n <= 0:
        return (0, 0)
    cols = math.ceil(math.sqrt(n))
    rows = math.ceil(n / cols)
    return (rows, cols)


def add_grid_slide(
    image_paths: list[Path],
    output_path: Path,
    slide_width_in: float = 13.333,
    slide_height_in: float = 7.5,
    margin_in: float = 0.4,
    cell_gap_in: float = 0.15,
) -> None:
    """
    Build a single-slide .pptx with all `image_paths` arranged in a grid.

    Each image is scaled to fit inside its grid cell while keeping its
    original aspect ratio (so photos don't look stretched or squashed),
    and centered within that cell.
    """
    rows, cols = choose_grid_shape(len(image_paths))

    prs = Presentation()
    # A "blank" layout is normally index 6 in the default template --
    # it has no title or body placeholder boxes we'd have to remove.
    blank_layout = prs.slide_layouts[6]
    prs.slide_width = Inches(slide_width_in)
    prs.slide_height = Inches(slide_height_in)
    slide = prs.slides.add_slide(blank_layout)

    # Usable area = whole slide minus a margin on every side.
    usable_width_in = slide_width_in - 2 * margin_in
    usable_height_in = slide_height_in - 2 * margin_in

    # Size of one grid cell, including its own gap allowance.
    cell_width_in = (usable_width_in - (cols - 1) * cell_gap_in) / cols
    cell_height_in = (usable_height_in - (rows - 1) * cell_gap_in) / rows

    for index, image_path in enumerate(image_paths):
        row = index // cols
        col = index % cols

        # Top-left corner of this cell.
        cell_left_in = margin_in + col * (cell_width_in + cell_gap_in)
        cell_top_in = margin_in + row * (cell_height_in + cell_gap_in)

        # Open the image with Pillow (via the pillow-heif plugin) so we
        # can read its pixel dimensions -- needed to scale it into the
        # cell without distorting it.
        #
        # We also convert it to PNG bytes in memory here, rather than
        # pointing python-pptx at the original .heic file. That's
        # because python-pptx only recognizes a fixed list of image
        # formats when it writes the .pptx package (JPEG, PNG, GIF,
        # BMP, TIFF, WMF) and does not know what to do with HEIC --
        # it would raise a "ValueError: unsupported image format"
        # if we handed it the .heic path directly. Since Pillow can
        # already read HEIC (thanks to register_heif_opener above) and
        # write PNG, converting in memory sidesteps that limitation
        # without ever writing a temporary file to disk.
        with Image.open(image_path) as img:
            img_width_px, img_height_px = img.size
            # HEIC photos can be in color modes (e.g. CMYK) that PNG
            # doesn't support; converting to RGB keeps this reliable.
            rgb_img = img.convert("RGB")
            image_bytes = io.BytesIO()
            rgb_img.save(image_bytes, format="PNG")
            image_bytes.seek(0)  # rewind so add_picture reads from the start

        img_aspect = img_width_px / img_height_px
        cell_aspect = cell_width_in / cell_height_in

        if img_aspect > cell_aspect:
            # Image is relatively wider than the cell -> width is the
            # limiting dimension.
            draw_width_in = cell_width_in
            draw_height_in = cell_width_in / img_aspect
        else:
            # Image is relatively taller than the cell -> height is the
            # limiting dimension.
            draw_height_in = cell_height_in
            draw_width_in = cell_height_in * img_aspect

        # Center the (possibly smaller) image within its cell.
        left_in = cell_left_in + (cell_width_in - draw_width_in) / 2
        top_in = cell_top_in + (cell_height_in - draw_height_in) / 2

        # add_picture accepts either a file path or a file-like object
        # (our in-memory PNG bytes both work) and embeds it into the
        # .pptx package at the given position and size.
        slide.shapes.add_picture(
            image_bytes,
            Inches(left_in),
            Inches(top_in),
            width=Inches(draw_width_in),
            height=Inches(draw_height_in),
        )

    prs.save(str(output_path))


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Arrange all .heic images in a folder into a grid on one PowerPoint slide."
    )
    parser.add_argument("folder", type=str, help="Path to the folder containing .heic images")
    parser.add_argument(
        "-o",
        "--output",
        type=str,
        default="heic_grid.pptx",
        help="Path for the output .pptx file (default: heic_grid.pptx)",
    )
    args = parser.parse_args()

    folder = Path(args.folder).expanduser().resolve()
    if not folder.is_dir():
        raise SystemExit(f"Error: '{folder}' is not a folder that exists.")

    image_paths = find_heic_files(folder)
    if not image_paths:
        raise SystemExit(f"No .heic files found in '{folder}'.")

    output_path = Path(args.output).expanduser().resolve()
    add_grid_slide(image_paths, output_path)

    rows, cols = choose_grid_shape(len(image_paths))
    print(f"Found {len(image_paths)} HEIC image(s). Arranged as a {rows}x{cols} grid.")
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    main()