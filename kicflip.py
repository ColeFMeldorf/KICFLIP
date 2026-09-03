#!/usr/bin/env python3
"""
heic_grid_to_pptx.py

Take a folder of .heic photos (and some blood-work plots) and drop
them into specific slides of an existing PowerPoint template.

Usage:
    python heic_grid_to_pptx.py /path/to/folder --template template.pptx -o output.pptx

What this script needs installed (one-time setup):
    pip install pillow-heif python-pptx pillow
"""

import argparse
import functools
import glob
import io
import math
from pathlib import Path

from PIL import Image
from pillow_heif import register_heif_opener
from pptx import Presentation
from pptx.opc.packuri import PackURI
from pptx.util import Inches

from plotting import _load_xls, make_plot

# This line is what actually "teaches" Pillow to understand .heic files.
# It has to run once, before any Image.open() call on a HEIC file.
register_heif_opener()


def find_heic_files(folder: Path) -> list[Path]:
    """
    Return every .heic file directly inside `folder`, sorted by name.
    """
    matches = list(folder.glob("*.heic")) + list(folder.glob("*.HEIC"))
    unique_matches = sorted(set(matches), key=lambda p: p.name.lower())
    return unique_matches


def choose_grid_shape(n: int) -> tuple[int, int]:
    """
    Given `n` images, decide how many rows and columns to use so the
    grid is as close to square as possible.
    """
    if n <= 0:
        return (0, 0)
    cols = math.ceil(math.sqrt(n))
    rows = math.ceil(n / cols)
    return (rows, cols)


def populate_grid_slide(
    slide,
    slide_width_in: float,
    slide_height_in: float,
    image_paths: list[Path],
    margin_in: float = 0.4,
    cell_gap_in: float = 0.15,
) -> None:
    """
    Fill an already-created, empty slide with all `image_paths`
    arranged in a grid, each scaled to fit its cell without distortion
    and centered within it.
    """
    rows, cols = choose_grid_shape(len(image_paths))

    usable_width_in = slide_width_in - 2 * margin_in
    usable_height_in = slide_height_in - 2 * margin_in
    cell_width_in = (usable_width_in - (cols - 1) * cell_gap_in) / cols
    cell_height_in = (usable_height_in - (rows - 1) * cell_gap_in) / rows

    for index, image_path in enumerate(image_paths):
        row = index // cols
        col = index % cols
        cell_left_in = margin_in + col * (cell_width_in + cell_gap_in)
        cell_top_in = margin_in + row * (cell_height_in + cell_gap_in)

        with Image.open(image_path) as img:
            img_width_px, img_height_px = img.size
            rgb_img = img.convert("RGB")
            image_bytes = io.BytesIO()
            rgb_img.save(image_bytes, format="PNG")
            image_bytes.seek(0)

        img_aspect = img_width_px / img_height_px
        cell_aspect = cell_width_in / cell_height_in

        if img_aspect > cell_aspect:
            draw_width_in = cell_width_in
            draw_height_in = cell_width_in / img_aspect
        else:
            draw_height_in = cell_height_in
            draw_width_in = cell_height_in * img_aspect

        left_in = cell_left_in + (cell_width_in - draw_width_in) / 2
        top_in = cell_top_in + (cell_height_in - draw_height_in) / 2

        slide.shapes.add_picture(
            image_bytes,
            Inches(left_in),
            Inches(top_in),
            width=Inches(draw_width_in),
            height=Inches(draw_height_in),
        )


def add_grid_slide(
    image_paths: list[Path],
    output_path: Path,
    slide_width_in: float = 13.333,
    slide_height_in: float = 7.5,
    margin_in: float = 0.4,
    cell_gap_in: float = 0.15,
) -> None:
    """
    Build a brand-new, single-slide .pptx with all `image_paths`
    arranged in a grid, and save it to `output_path`.
    """
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    prs.slide_width = Inches(slide_width_in)
    prs.slide_height = Inches(slide_height_in)
    slide = prs.slides.add_slide(blank_layout)

    populate_grid_slide(
        slide,
        slide_width_in,
        slide_height_in,
        image_paths,
        margin_in=margin_in,
        cell_gap_in=cell_gap_in,
    )

    prs.save(str(output_path))


def find_blank_layout_index(template_path: Path):
    """
    Figure out which layout in this template (by position in
    prs.slide_layouts) produces slides with zero placeholder shapes --
    most templates ship one meant for exactly this ("Blank") -- and
    return its index, or None if every layout leaves at least one
    placeholder behind.

    IMPORTANT: this loads its OWN separate, throwaway copy of the
    template purely to test this, and discards that copy immediately
    -- it never touches whatever live `Presentation` object you're
    actually assembling. See `replace_slide_in_template`'s docstring
    for why that separation matters: call this ONCE per template and
    pass the result into every `replace_slide_in_template` call you
    make against that template, rather than re-detecting it against
    your live, in-progress presentation.
    """
    scratch_prs = Presentation(str(template_path))
    for i, layout in enumerate(scratch_prs.slide_layouts):
        probe_slide = scratch_prs.slides.add_slide(layout)
        if len(probe_slide.placeholders) == 0:
            return i
    return None


def _assign_unique_slide_partname(prs: Presentation, slide_part) -> None:
    """
    Force `slide_part` (a just-added slide's underlying package part)
    onto a slide filename ("partname", e.g. "/ppt/slides/slide7.xml")
    that's not already used by anything else currently in `prs`.

    Why this is necessary: python-pptx decides a new slide's filename
    purely from HOW MANY slides currently exist in the presentation --
    not by scanning for a name that's actually free. That's fine for
    a single, one-shot deck. But if you replace more than one slide in
    the same live `Presentation` object (as this script does -- once
    for the lab-value plots, once for the iPhone photos), the slide
    count returns to its original number after each replacement
    (one added, one removed nets to zero change). Since the next
    slide's filename is computed from that same count both times,
    TWO DIFFERENT replacement calls can independently compute the
    IDENTICAL "next" filename for two DIFFERENT slides that are both
    still kept in the deck. The result: two slides silently claim the
    same internal filename. python-pptx itself doesn't notice or
    complain when this happens, but when the file is saved, Python's
    zip writer prints "Duplicate name" warnings, and PowerPoint
    detects the corruption and asks to "repair" the file when opened.

    The fix is to explicitly walk every part actually reachable in the
    presentation (`prs.part.package.iter_parts()` -- "reachable"
    meaning something in the file still points to it, which is also
    how python-pptx decides what to include when saving) and assign a
    number that's genuinely not in use by any of them, rather than
    trusting python-pptx's built-in count-based guess.
    """
    used_numbers = set()
    prefix = "/ppt/slides/slide"
    suffix = ".xml"
    for part in prs.part.package.iter_parts():
        partname = str(part.partname)
        if partname.startswith(prefix) and partname.endswith(suffix):
            digits = partname[len(prefix): -len(suffix)]
            if digits.isdigit():
                used_numbers.add(int(digits))

    n = 1
    while n in used_numbers:
        n += 1
    slide_part.partname = PackURI(f"{prefix}{n}{suffix}")


def replace_slide_in_template(
    prs: Presentation,
    index: int,
    populate_slide,
    blank_layout_index: int | None = None,
) -> None:
    """
    Replace the slide currently at position `index` (0-based -- the
    first slide is index 0) in `prs`, mutating `prs` in place. Does
    NOT save -- call `prs.save(...)` yourself once you're done making
    all the replacements you want to make on this `prs`.

    `populate_slide` is a function *you* provide with the signature
    `populate_slide(new_slide, slide_width_in, slide_height_in)`,
    called once to draw content onto the freshly-added slide -- a grid
    of pictures (`populate_grid_slide`, above), a title and bullets, a
    chart, whatever. This function itself never looks at what got
    drawn; it only handles the mechanics of putting a new slide at the
    right position and removing the old one.

    `blank_layout_index`: pass the result of a single call to
    `find_blank_layout_index(template_path)` here. If you're calling
    this function more than once against the same `prs` (as this
    script does), compute it once, before your first call, and pass
    the SAME value into every call -- see that function's docstring
    for why re-detecting it against the live `prs` on a later call is
    unsafe, not just slower. If omitted, this falls back to reusing
    the replaced slide's own current layout (which may leave empty
    "click to add text" placeholder boxes on the new slide -- harmless
    to present, just a little untidy to edit).

    Why this doesn't just take a ready-made "slide" object: in
    python-pptx, a Slide is permanently tied to the specific
    Presentation package it was created inside -- like a page bound
    into one particular book -- so there's no way to lift one out of
    a different Presentation object and staple it in here. Instead,
    this creates a genuinely new slide directly inside `prs`.
    """
    slide_count = len(prs.slides)
    if not (0 <= index < slide_count):
        raise IndexError(
            f"index {index} is out of range for a presentation with {slide_count} slide(s) "
            f"(valid indices: 0 to {slide_count - 1})"
        )

    if blank_layout_index is not None:
        layout_to_reuse = prs.slide_layouts[blank_layout_index]
    else:
        layout_to_reuse = prs.slides[index].slide_layout

    # python-pptx can only ever ADD a new slide onto the end of the
    # deck -- there's no "insert at position" operation. So we add it
    # at the end first, and move it into place as a second step below.
    new_slide = prs.slides.add_slide(layout_to_reuse)

    # See _assign_unique_slide_partname's docstring: this is the fix
    # for the "Duplicate name" / "needs repair" bug.
    _assign_unique_slide_partname(prs, new_slide.part)

    slide_width_in = prs.slide_width / 914400
    slide_height_in = prs.slide_height / 914400

    populate_slide(new_slide, slide_width_in, slide_height_in)

    # sldIdLst (slide ID list) is the ordered list controlling slide
    # order -- the book's table of contents, separate from the pages
    # themselves. Pull out the entry for our new slide (always last,
    # since add_slide appends) and the entry for the old slide at
    # `index`, then reinsert just the new one at `index`.
    slide_id_list = prs.slides._sldIdLst
    entries = list(slide_id_list)
    new_slide_entry = entries[-1]
    old_slide_entry = entries[index]

    slide_id_list.remove(new_slide_entry)
    slide_id_list.remove(old_slide_entry)
    slide_id_list.insert(index, new_slide_entry)

    # Drops the relationship to the old slide's part; once nothing
    # references it, python-pptx excludes it from the saved package.
    # (Any images embedded ONLY in that removed slide can be left
    # behind as unused data in the package -- harmless, just a little
    # extra file size.)
    prs.part.drop_rel(old_slide_entry.rId)


def grid_slide_populate_and_replace(
    prs: Presentation,
    image_paths: list[Path],
    index: int,
    blank_layout_index: int | None,
    label: str,
) -> None:
    """
    Convenience wrapper: build a grid-of-images populate function bound
    to `image_paths`, and use it to replace slide `index` in `prs`.
    Mutates `prs` in place; does not save.
    """
    rows, cols = choose_grid_shape(len(image_paths))
    populate_fn = functools.partial(populate_grid_slide, image_paths=image_paths)
    replace_slide_in_template(prs, index, populate_fn, blank_layout_index=blank_layout_index)
    print(f"{label}: arranged {len(image_paths)} image(s) as a {rows}x{cols} grid, replacing slide {index}.")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Insert grids of images into specific slides of a PowerPoint template."
    )
    parser.add_argument("folder", type=str, help="Path to the folder containing .heic images and .xls data")
    parser.add_argument(
        "-o", "--output", type=str, default="heic_grid.pptx",
        help="Path for the output .pptx file (default: heic_grid.pptx)",
    )
    parser.add_argument(
        "--template", type=str, default="template.pptx",
        help="Path to an existing .pptx template to insert the grids into.",
    )
    parser.add_argument(
        "--dont_overwrite", action="store_true",
        help="If specified, do not delete an existing output file before writing.",
    )
    args = parser.parse_args()

    folder = Path(args.folder).expanduser().resolve()

    if not folder.endswith("/"):
        folder = Path(str(folder) + "/")

    if not folder.is_dir():
        raise SystemExit(f"Error: '{folder}' is not a folder that exists.")

    image_paths = find_heic_files(folder)
    if not image_paths:
        raise SystemExit(f"No .heic files found in '{folder}'.")

    template_path = Path(args.template).expanduser().resolve()
    if not template_path.is_file():
        template_path = Path(__file__).parent / args.template
        if not template_path.is_file():
            raise SystemExit(f"Error: template '{args.template}' is not a file that exists.")
    print("Using template slide deck:", template_path)

    output_path = Path(args.output).expanduser().resolve()
    if not args.dont_overwrite and output_path.exists():
        print(f"Removing existing output file: {output_path}")
        try:
            output_path.unlink()
        except PermissionError:
            raise PermissionError(
                "You probably have the output file open in PowerPoint. Please close it and try again."
            )

    ##### MAKE PLOTS ########
    df = _load_xls(folder / "blood-sample.xls")
    print("Making plot for Lactate...")
    make_plot(df, "Lactate", args.folder)
    df = _load_xls(folder / "trends.xls")
    print("Making plot for HCT...")
    make_plot(df, "HCT", args.folder)
    print("Making plot for Perfusion Flows...")
    make_plot(df, ["PF", "HAF", "PVF"], args.folder, title="Perfusion Flows")
    print("Making plot for Hepatic Artery Pressures...")
    make_plot(df, ["HAP Mean", "HAP Systolic", "HAP Diastolic"], args.folder, title="Hepatic Artery Pressures")
    plot_files = [Path(p) for p in glob.glob(str(folder / "*.png"))]



    # Compute this ONCE, from a disposable copy of the template, and
    # reuse it for every replacement below -- see
    # find_blank_layout_index's docstring for why re-detecting this
    # per-call against the live, in-progress presentation is unsafe.
    blank_layout_index = find_blank_layout_index(template_path)

    prs = Presentation(str(template_path))

    print("Building plot slide...")
    plot_files_HCT = [Path(p) for p in glob.glob(str(folder / "*HCT*.png"))]
    plot_files_Lactate = [Path(p) for p in glob.glob(str(folder / "*Lactate*.png"))]
    plot_files = plot_files_HCT + plot_files_Lactate
    grid_slide_populate_and_replace(
        prs, plot_files, index=5, blank_layout_index=blank_layout_index, label="Plots"
    )

    plot_files_PF = [Path(p) for p in glob.glob(str(folder / "*Perfusion Flows*.png"))]
    plot_files_HAP = [Path(p) for p in glob.glob(str(folder / "*Hepatic Artery Pressures*.png"))]
    plot_files = plot_files_PF + plot_files_HAP
    grid_slide_populate_and_replace(
        prs, plot_files, index=4, blank_layout_index=blank_layout_index, label="Plots"
    )

    print("Building iPhone picture slide...")
    grid_slide_populate_and_replace(
        prs, image_paths, index=3, blank_layout_index=blank_layout_index, label="Photos"
    )

    prs.save(str(output_path))
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    main()