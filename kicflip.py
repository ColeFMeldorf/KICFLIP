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
import sys

from PIL import Image
from pillow_heif import register_heif_opener
from pptx import Presentation
from pptx.util import Inches
import glob
from plotting import _load_xls, make_plot


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


import copy
def copy_slide_to_presentation(src_prs, dest_prs, slide_index_to_copy):
    # 2. Get the slide you want to copy
    source_slide = src_prs.slides[slide_index_to_copy]

    # 3. Create a blank slide in the destination presentation using its own blank layout
    # Typically layout index 6 is a completely blank slide layout
    blank_layout = dest_prs.slide_layouts[6]
    new_slide = dest_prs.slides.add_slide(blank_layout)

    # 4. Copy all shapes/contents from the source slide to the new destination slide
    for shape in source_slide.shapes:
        # Deep copy the XML element of the shape
        el = shape.element
        new_el = copy.deepcopy(el)

        # Append the copied shape into the destination slide's shape tree
        new_slide.shapes._spTree.append(new_el)

    # 5. Handle slide relationships (important for copying images/charts correctly)
    # for rel in source_slide.part.rels.values():
    #     # Prevent copying slide notes or other unnecessary sub-structures
    #     if "notesSlide" not in rel.reltype:
    #         new_slide.part.rels.get_or_add(
    #             rel.reltype,
    #             rel._target,
    #             rel.rId
    #         )

    return dest_prs

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
    prs: Presentation = Presentation(),
    index_to_replace=None,
) -> None:
    """
    Build a single-slide .pptx with all `image_paths` arranged in a grid.

    Each image is scaled to fit inside its grid cell while keeping its
    original aspect ratio (so photos don't look stretched or squashed),
    and centered within that cell.
    """
    rows, cols = choose_grid_shape(len(image_paths))


    # A "blank" layout is normally index 6 in the default template --
    # it has no title or body placeholder boxes we'd have to remove.
    blank_layout = prs.slide_layouts[6]
    prs.slide_width = Inches(slide_width_in)
    prs.slide_height = Inches(slide_height_in)
    slide = prs.slides.add_slide(blank_layout)

    replace_slide(index_to_replace, prs) if index_to_replace is not None else None

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

def populate_grid_slide(
    slide,
    image_paths: list[Path],
    slide_width_in: float,
    slide_height_in: float,
    margin_in: float = 0.4,
    cell_gap_in: float = 0.15,
) -> None:
    """
    Fill an *already-created, empty* slide with all `image_paths`
    arranged in a grid.

    This is split out from `add_grid_slide` (below) so the same
    grid-drawing logic can be reused whether the slide lives in a
    brand-new presentation or was just added to an existing template
    (see `replace_slide_in_template`). This function doesn't know or
    care which -- it just draws pictures onto whatever `slide` object
    it's handed.

    Each image is scaled to fit inside its grid cell while keeping its
    original aspect ratio (so photos don't look stretched or squashed),
    and centered within that cell.
    """
    rows, cols = choose_grid_shape(len(image_paths))

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

# def replace_slide(index_to_replace, prs):
#     # Determine the index of the slide you want to replace (e.g., index 2)
#     index_to_replace = 2
#     old_slide = prs.slides[index_to_replace]

#     # Create a new slide with a specific layout (e.g., blank or title layout)
#     blank_slide_layout = prs.slide_layouts[6]
#     new_slide = prs.slides.add_slide(blank_slide_layout)

#     # Move the new slide's XML element to the old slide's position
#     # and remove the old slide's XML element
#     old_rId = prs.slides._sldIdLst[index_to_replace]
#     new_rId = prs.slides._sldIdLst[-1]

#     prs.slides._sldIdLst.insert(index_to_replace, new_rId)
#     prs.slides._sldIdLst.remove(old_rId)

#     return prs
#     # Save the updated presentation
#     # prs.save("updated_presentation.pptx")


def replace_slide_in_template(
    template_path: Path,
    image_paths: list[Path],
    index: int,
    output_path: Path,
    margin_in: float = 0.4,
    cell_gap_in: float = 0.15,
) -> None:
    """
    Open an existing template .pptx, replace the slide currently at
    position `index` (0-based -- the first slide is index 0) with a
    new slide showing `image_paths` arranged in a grid, and save the
    result to `output_path`. The original template file is left
    untouched; this only writes to `output_path`.

    Why this doesn't just take a ready-made "slide" object: in
    python-pptx, a Slide isn't a standalone, portable thing -- it's
    permanently tied to the specific Presentation (the specific .pptx
    package) it was created inside, the way a page is bound into one
    particular book. There's no supported way to lift a Slide out of
    one Presentation object and staple it into a different one.
    So instead, this function builds the new slide's picture grid
    directly inside the template's own Presentation object, then
    reshuffles the slide order so it lands at `index`, then removes
    whatever slide used to be there.

    Some vocabulary that shows up below:
      - "Presentation object" (`prs`): python-pptx's in-memory
        representation of the whole .pptx file -- every slide,
        layout, and embedded image.
      - "Slide layout": a template/blueprint a slide is built from
        (e.g. "Title Slide", "Blank"). Every slide points at one.
      - The slide *order* in a .pptx is controlled by a single list,
        internally called the "slide ID list". Rearranging that list
        is how PowerPoint (and we) reorder slides -- the slides
        themselves don't store their own position.
    """
    prs = Presentation(str(template_path))

    slide_count = len(prs.slides)
    if not (0 <= index < slide_count):
        raise IndexError(
            f"index {index} is out of range for a template with {slide_count} slide(s) "
            f"(valid indices: 0 to {slide_count - 1})"
        )

    # Pick a layout for the new slide. We still want it to inherit the
    # template's theme (colors, fonts, background), but we'd rather NOT
    # reuse a layout like "Title and Content" as-is, since that leaves
    # empty, unused title/body placeholder boxes sitting on the slide
    # (harmless when presenting, but shows as "Click to add title"
    # prompts if someone opens the file to edit it). So: look for a
    # layout that produces zero placeholder shapes when used -- most
    # PowerPoint templates ship one named "Blank" for exactly this
    # purpose -- and fall back to the replaced slide's own layout only
    # if nothing like that exists in this template.
    layout_to_reuse = None
    for layout in prs.slide_layouts:
        probe_slide = prs.slides.add_slide(layout)
        is_placeholder_free = len(probe_slide.placeholders) == 0
        # Remove the probe slide immediately -- we're only using it to
        # test the layout, not to keep it.
        prs.part.drop_rel(prs.slides._sldIdLst[-1].rId)
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[-1])
        if is_placeholder_free:
            layout_to_reuse = layout
            break
    if layout_to_reuse is None:
        layout_to_reuse = prs.slides[index].slide_layout

    # python-pptx can only ever ADD a new slide onto the end of the
    # deck -- there's no "insert at position" operation. So we add it
    # at the end first, and move it into place as a second step below.
    new_slide = prs.slides.add_slide(layout_to_reuse)

    # Read the template's own slide dimensions (they're stored in EMUs,
    # a PowerPoint measurement unit; 914400 EMU = 1 inch) so our grid
    # math matches this template's actual slide size, rather than
    # assuming the 13.333x7.5in size add_grid_slide defaults to.
    slide_width_in = prs.slide_width / 914400
    slide_height_in = prs.slide_height / 914400

    populate_grid_slide(
        new_slide,
        image_paths,
        slide_width_in=slide_width_in,
        slide_height_in=slide_height_in,
        margin_in=margin_in,
        cell_gap_in=cell_gap_in,
    )

    # `sldIdLst` (slide ID list) is the actual ordered list PowerPoint
    # reads to decide slide order -- think of it as the book's table of
    # contents, separate from the pages themselves. Right now it looks
    # like: [old_slide_0, ..., old_slide_at_index, ..., old_slide_last, new_slide]
    # because add_slide() always appends. We need to pull out both the
    # entry for our new slide and the entry for the old slide at
    # `index`, then reinsert just the new one at `index`.
    slide_id_list = prs.slides._sldIdLst
    entries = list(slide_id_list)

    new_slide_entry = entries[-1]        # just-appended slide is always last
    old_slide_entry = entries[index]     # the slide we're replacing

    # Removing the old slide's entry here only removes it from the
    # ordering list, not the slide's actual content -- that's handled
    # separately just below with drop_rel.
    slide_id_list.remove(new_slide_entry)
    slide_id_list.remove(old_slide_entry)
    slide_id_list.insert(index, new_slide_entry)

    # The ordering list only controls sequence -- the old slide's
    # actual content is a separate "part" of the .pptx package that
    # still exists until we explicitly drop it. `drop_rel` removes the
    # relationship between the presentation and that slide part; when
    # nothing references a part anymore, python-pptx removes the part
    # itself from the package on save.
    #
    # Note: any images that were embedded *only* in the removed slide
    # can be left behind in the package as unused data (harmless to
    # open, just a little extra file size) -- python-pptx's part
    # clean-up doesn't chase that deep. For a handful of replacements
    # this is negligible; skip past it unless file size becomes a
    # real problem.
    prs.part.drop_rel(old_slide_entry.rId)

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
    parser.add_argument(
        "--dont_overwrite",
        action="store_true",
        help="If specified, the script will not overwrite an existing output file.",

    )


    args = parser.parse_args()

    if not args.dont_overwrite:
        # remove the output file if it already exists, so we can overwrite it.
        output_path = Path(args.output).expanduser().resolve()
        if output_path.exists():
            print(f"Removing existing output file: {output_path}")
            try:
                output_path.unlink()
            except PermissionError:
                raise PermissionError("You probably have the output file open in PowerPoint. Please close it and try again.")



    prs = Presentation("template.pptx")

    folder = Path(args.folder).expanduser().resolve()
    if not folder.is_dir():
        raise SystemExit(f"Error: '{folder}' is not a folder that exists.")

    image_paths = find_heic_files(folder)
    if not image_paths:
        raise SystemExit(f"No .heic files found in '{folder}'.")

    output_path = Path(args.output).expanduser().resolve()
    #prs = Presentation()

    # Add the Title Slide
    print("Adding title slide...")
    #prs = copy_slide_to_presentation(template_prs, prs, 0)

    # slide_layout = prs.slide_layouts[1]
    # slide = prs.slides.add_slide(slide_layout)
    # # Add a title to the slide
    # slide.shapes.title.text = "LXX Perfusion"
    # #title_shape = slide.shapes.title
    # #title_shape.text = "LXX Perfusion"
    # # Add a subtitle to the slide
    # subtitle_shape = slide.placeholders[0]
    # subtitle_shape.text = "Description that I enter of perfusion."

    print("Adding an image grid slide...")
    add_grid_slide(image_paths, output_path, prs=prs, index_to_replace=3)

    rows, cols = choose_grid_shape(len(image_paths))
    print(f"Found {len(image_paths)} HEIC image(s). Arranged as a {rows}x{cols} grid.")

    folder_path = Path(args.folder)
    glob_path = str(folder_path / "*.xlsx")
    print(f"Looking for .xlsx files in: {glob_path}")

    xlsx_files = glob.glob(glob_path)

    df = _load_xls(Path(args.folder) / Path("blood-sample.xls"))
    print("Making plot for Lactate...")
    make_plot(df, "Lactate", args.folder)
    df = _load_xls(Path(args.folder) / Path("trends.xls"))
    print("Making plot for HCT...")
    make_plot(df, "HCT", args.folder)

    # Fetch those plots and add them to the presentation
    plot_files = glob.glob(str(folder_path / "*.png"))
    add_grid_slide(plot_files, output_path, prs=prs)

    print(f"Saved: {output_path}")


if __name__ == "__main__":
    main()