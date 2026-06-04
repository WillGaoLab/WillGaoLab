"""Streamlit upload/download interface for Excel pixel art generation."""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import streamlit as st

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from excel_pixel_art.analytics import render_clarity_analytics
from excel_pixel_art.canvas import CANVAS_PRESETS, FIT_MODES, ORIENTATIONS
from excel_pixel_art.converter import image_to_excel
from excel_pixel_art.physical import image_to_physical_excel, image_to_physical_masks, image_to_physical_pdf
from excel_pixel_art.poster_print import validate_poster_split


def main() -> None:
    st.set_page_config(
        page_title="Excel Pixel Art Generator",
        page_icon="",
        layout="wide",
    )
    render_clarity_analytics()

    st.title("Excel Pixel Art Generator")
    st.caption("Upload one image, then generate completely separate Digital and Physical outputs.")

    uploaded_file = st.file_uploader(
        "Image",
        type=["png", "jpg", "jpeg", "webp", "bmp"],
        accept_multiple_files=False,
    )
    if uploaded_file is not None:
        st.image(uploaded_file, caption=uploaded_file.name, use_container_width=True)

    st.divider()
    _digital_layer(uploaded_file)
    st.divider()
    _physical_layer(uploaded_file)
    st.divider()
    _use_privacy_and_liability()


def _digital_layer(uploaded_file) -> None:
    st.header("Digital Layer")
    st.markdown("**Excel Mode**")
    st.caption("Excel pixel-art workflow and workbook output.")

    left, right = st.columns([1, 1])
    with left:
        st.subheader("Canvas")
        canvas_keys = [""] + sorted(CANVAS_PRESETS)
        canvas_size = st.selectbox(
            "Paper size",
            options=canvas_keys,
            index=canvas_keys.index("a4"),
            format_func=_canvas_label,
            key="digital_canvas_size",
        )
        orientation = st.selectbox(
            "Orientation",
            options=sorted(ORIENTATIONS),
            index=0,
            key="digital_orientation",
        )
        fit = st.selectbox(
            "Image fit",
            options=sorted(FIT_MODES),
            index=0,
            key="digital_fit",
        )
        background_color = st.color_picker(
            "Background",
            value="#FFFFFF",
            key="digital_background",
        ).lstrip("#")

    with right:
        st.subheader("Detail")
        use_custom_resolution = st.checkbox(
            "Use exact resolution",
            value=True,
            key="digital_exact_resolution",
        )
        if use_custom_resolution:
            width_cells = st.number_input(
                "Width cells",
                min_value=1,
                max_value=2000,
                value=32,
                step=8,
                key="digital_width",
            )
            height_cells = st.number_input(
                "Height cells",
                min_value=1,
                max_value=2000,
                value=32,
                step=8,
                key="digital_height",
            )
            resolution = (int(width_cells), int(height_cells))
            max_size = 32
        else:
            max_size = st.number_input(
                "Max cells",
                min_value=1,
                max_value=512,
                value=32,
                step=8,
                key="digital_max_size",
            )
            resolution = None

        color_count = st.slider(
            "Indexed colors",
            min_value=2,
            max_value=256,
            value=24,
            key="digital_colors",
        )
        cell_size = st.number_input(
            "Excel cell size",
            min_value=0.1,
            max_value=20.0,
            value=3.0,
            step=0.1,
            key="digital_cell_size",
        )

    if st.button(
        "Generate Digital workbook",
        type="primary",
        disabled=uploaded_file is None,
        key="generate_digital",
    ):
        with st.spinner("Generating Digital workbook..."):
            workbook_bytes, download_name = _build_digital_workbook(
                uploaded_file=uploaded_file,
                max_size=int(max_size),
                cell_size=float(cell_size),
                color_count=int(color_count),
                canvas_size=canvas_size or None,
                resolution=resolution,
                orientation=orientation,
                fit=fit,
                background_color=background_color,
            )
        st.download_button(
            "Download Digital workbook",
            data=workbook_bytes,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_digital",
        )


def _physical_layer(uploaded_file) -> None:
    st.header("Physical Layer")
    st.markdown("**Print Mode**")
    st.caption("Independent print canvas, material palette, poster split, masks, and workbook output.")

    canvas_column, detail_column = st.columns([1, 1])
    canvas_keys = [""] + sorted(CANVAS_PRESETS)

    with canvas_column:
        st.subheader("Print Canvas")
        canvas_size = st.selectbox(
            "Physical paper size",
            options=canvas_keys,
            index=canvas_keys.index("a4"),
            format_func=_canvas_label,
            key="physical_canvas_size",
        )
        orientation = st.selectbox(
            "Physical orientation",
            options=sorted(ORIENTATIONS),
            index=0,
            key="physical_orientation",
        )
        fit = st.selectbox(
            "Physical image fit",
            options=sorted(FIT_MODES),
            index=0,
            key="physical_fit",
        )
        background_color = st.color_picker(
            "Physical background",
            value="#FFFFFF",
            key="physical_background",
        ).lstrip("#")

    with detail_column:
        st.subheader("Material Detail")
        use_custom_resolution = st.checkbox(
            "Use exact physical resolution",
            value=True,
            key="physical_exact_resolution",
        )
        if use_custom_resolution:
            width_cells = st.number_input(
                "Physical width cells",
                min_value=1,
                max_value=2000,
                value=32,
                step=8,
                key="physical_width",
            )
            height_cells = st.number_input(
                "Physical height cells",
                min_value=1,
                max_value=2000,
                value=32,
                step=8,
                key="physical_height",
            )
            resolution = (int(width_cells), int(height_cells))
            max_size = 32
        else:
            max_size = st.number_input(
                "Physical max cells",
                min_value=1,
                max_value=512,
                value=32,
                step=8,
                key="physical_max_size",
            )
            resolution = None

        material_color_count = st.slider(
            "Material Palette colors",
            min_value=2,
            max_value=256,
            value=24,
            key="physical_colors",
        )
        cell_size = st.number_input(
            "Print cell size",
            min_value=0.1,
            max_value=20.0,
            value=3.0,
            step=0.1,
            key="physical_cell_size",
        )

    palette_column, poster_column, masks_column, output_column = st.columns([1, 1, 1, 1])
    with palette_column:
        st.subheader("Material Palette")
        palette_label = st.radio(
            "Palette",
            options=["Adaptive", "LEGO", "Liquitex Basics 24"],
            label_visibility="collapsed",
            key="physical_palette_mode",
        )
        palette_mode = {
            "Adaptive": "adaptive",
            "LEGO": "lego",
            "Liquitex Basics 24": "liquitex_basics_24",
        }[palette_label]
        if palette_mode != "adaptive":
            material_color_count = 40 if palette_mode == "lego" else 24

    with poster_column:
        st.subheader("Poster Split")
        split_poster = st.checkbox("Enable", value=False, key="physical_poster_split")
        poster_pages_wide = st.number_input(
            "Pages Wide", min_value=1, max_value=20, value=2, step=1,
            disabled=not split_poster, key="physical_pages_wide",
        )
        poster_pages_tall = st.number_input(
            "Pages Tall", min_value=1, max_value=20, value=2, step=1,
            disabled=not split_poster, key="physical_pages_tall",
        )
        if not split_poster:
            poster_pages_wide = 1
            poster_pages_tall = 1
        poster_split_valid = True
        if resolution is None:
            poster_split_valid = not split_poster
            if split_poster:
                st.warning(
                    "Poster Split Summary\n\n"
                    "Status: Pending\n\n"
                    "Enable exact physical resolution to verify equal tile dimensions."
                )
        else:
            poster_summary = validate_poster_split(
                resolution,
                (int(poster_pages_wide), int(poster_pages_tall)),
            )
            poster_split_valid = poster_summary.valid
            cells_per_page = (
                f"{poster_summary.cells_per_page[0]} x {poster_summary.cells_per_page[1]} cells"
                if poster_summary.cells_per_page is not None
                else "Not available"
            )
            summary_text = (
                "**Poster Split Summary**\n\n"
                f"Master: `{resolution[0]} x {resolution[1]} cells`\n\n"
                f"Split: `{int(poster_pages_wide)} x {int(poster_pages_tall)} pages`\n\n"
                f"Each page: `{cells_per_page}`\n\n"
                f"Status: **{'Valid' if poster_summary.valid else 'Invalid'}**"
            )
            if poster_summary.valid:
                st.success(summary_text)
            else:
                st.error(f"{summary_text}\n\nReason: {poster_summary.reason}")

    with masks_column:
        st.subheader("Color Masks")
        generate_color_masks = st.checkbox("Generate Color Masks", value=False, key="physical_masks")
        max_color_masks = None

    with output_column:
        st.subheader("Output")
        output_workbook = st.checkbox("Workbook", value=True, key="physical_output_workbook")
        output_pdf = st.checkbox("Printable PDF", value=True, key="physical_output_pdf")
        output_masks = st.checkbox("Masks", value=True, disabled=not generate_color_masks, key="physical_output_masks")

    if palette_mode == "lego":
        st.caption("LEGO official color reference: https://www.bricklink.com/catalogColors.asp")
    elif palette_mode == "liquitex_basics_24":
        st.caption(
            "Liquitex official product: https://www.liquitex.com | "
            "24 Color Set: https://www.michaels.com/product/liquitex-basics-acrylic-24-color-paint-set-10268659"
        )
    if st.button(
        "Generate Physical outputs",
        type="primary",
        disabled=(
            uploaded_file is None
            or not poster_split_valid
            or not (output_workbook or output_pdf or (output_masks and generate_color_masks))
        ),
        key="generate_physical",
    ):
        with st.spinner("Generating Physical outputs..."):
            archive_bytes, archive_name = _build_physical_outputs(
                uploaded_file=uploaded_file,
                max_size=int(max_size),
                cell_size=float(cell_size),
                material_color_count=int(material_color_count),
                canvas_size=canvas_size or None,
                resolution=resolution,
                orientation=orientation,
                fit=fit,
                background_color=background_color,
                poster_pages=(int(poster_pages_wide), int(poster_pages_tall)),
                generate_color_masks=generate_color_masks,
                max_color_masks=int(max_color_masks) if max_color_masks is not None else None,
                palette_mode=palette_mode,
                output_workbook=output_workbook,
                output_pdf=output_pdf,
                output_masks=output_masks and generate_color_masks,
            )
        st.download_button(
            "Download Physical outputs ZIP",
            data=archive_bytes,
            file_name=archive_name,
            mime="application/zip",
            key="download_physical_outputs",
        )


def _use_privacy_and_liability() -> None:
    st.subheader("Use, Privacy, and Liability")
    with st.expander("Important information", expanded=False):
        st.markdown(
            """
**Uploaded Images**

Uploaded images are processed solely to generate requested outputs. The application does not
intentionally retain uploaded images after processing. Temporary processing files are deleted when no
longer needed. Third-party hosting and analytics providers, including Streamlit Community Cloud and
Microsoft Clarity, may process usage and technical data under their own terms and privacy policies.

Users are solely responsible for ensuring they have all necessary rights, permissions, and legal
authority to upload, transform, print, share, or otherwise use any content submitted to the application.

WillGaoLab and William (Peidong) Gao do not claim ownership of user-uploaded content and assume no
responsibility for uploaded materials, intellectual property violations, privacy violations, or resulting
claims.

**Generated Outputs**

All outputs, including image transformations, color mappings, material palettes, quantity estimates,
print layouts, masks, poster splits, and production recommendations, are generated automatically and may
contain errors or inaccuracies.

Actual results may vary significantly due to differences in screens, printers, materials, lighting
conditions, manufacturing tolerances, product availability, and production methods.

Users must independently verify all colors, dimensions, quantities, materials, compatibility, safety
requirements, and production settings before use.

**Third-Party Brands**

Any brand, product, retailer, or manufacturer names are provided solely for reference purposes.
WillGaoLab and William (Peidong) Gao are not affiliated with, sponsored by, endorsed by, or associated
with any referenced third party unless explicitly stated.

**Limitation of Liability**

All purchasing, printing, material selection, assembly, fabrication, and production decisions are made
entirely at the user's own risk.

To the maximum extent permitted by law, WillGaoLab and William (Peidong) Gao shall not be liable for any
losses, expenses, damages, failed projects, print errors, material waste, injuries, or other consequences
arising from the use of this application or its outputs.

This application and all outputs are provided **"AS IS"** without warranties of any kind.

This information is not legal advice.
            """
        )


def _build_digital_workbook(
    uploaded_file,
    max_size: int,
    cell_size: float,
    color_count: int,
    canvas_size: str | None,
    resolution: tuple[int, int] | None,
    orientation: str,
    fit: str,
    background_color: str,
) -> tuple[bytes, str]:
    return _build_workbook(
        uploaded_file=uploaded_file,
        output_suffix="digital",
        converter=image_to_excel,
        max_size=max_size,
        cell_size=cell_size,
        color_count=color_count,
        canvas_size=canvas_size,
        resolution=resolution,
        orientation=orientation,
        fit=fit,
        background_color=background_color,
    )


def _build_physical_workbook(
    uploaded_file,
    max_size: int,
    cell_size: float,
    material_color_count: int,
    canvas_size: str | None,
    resolution: tuple[int, int] | None,
    orientation: str,
    fit: str,
    background_color: str,
    poster_pages: tuple[int, int],
    generate_color_masks: bool,
    max_color_masks: int | None,
    palette_mode: str = "adaptive",
) -> tuple[bytes, str]:
    return _build_workbook(
        uploaded_file=uploaded_file,
        output_suffix="physical",
        converter=image_to_physical_excel,
        max_size=max_size,
        cell_size=cell_size,
        material_color_count=material_color_count,
        canvas_size=canvas_size,
        resolution=resolution,
        orientation=orientation,
        fit=fit,
        background_color=background_color,
        poster_pages=poster_pages,
        generate_color_masks=generate_color_masks,
        max_color_masks=max_color_masks,
        palette_mode=palette_mode,
    )


def _build_physical_outputs(
    uploaded_file,
    output_workbook: bool,
    output_pdf: bool,
    output_masks: bool,
    **options,
) -> tuple[bytes, str]:
    import zipfile

    suffix = Path(uploaded_file.name).suffix or ".image"
    stem = _safe_stem(uploaded_file.name)
    with tempfile.TemporaryDirectory() as directory:
        directory_path = Path(directory)
        image_path = directory_path / f"upload{suffix}"
        image_path.write_bytes(uploaded_file.getvalue())
        generated_paths = []
        if output_workbook:
            path = image_to_physical_excel(image_path, directory_path / f"{stem}_physical.xlsx", **options)
            generated_paths.append(path)
        if output_pdf:
            pdf_options = {key: value for key, value in options.items() if key not in {"generate_color_masks", "max_color_masks"}}
            path = image_to_physical_pdf(image_path, directory_path / f"{stem}_printable.pdf", **pdf_options)
            generated_paths.append(path)
        if output_masks:
            mask_options = {key: value for key, value in options.items() if key not in {"generate_color_masks"}}
            path = image_to_physical_masks(image_path, directory_path / f"{stem}_masks.zip", **mask_options)
            generated_paths.append(path)

        archive_path = directory_path / f"{stem}_physical_outputs.zip"
        with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for path in generated_paths:
                archive.write(path, path.name)
        return archive_path.read_bytes(), archive_path.name


def _build_workbook(uploaded_file, output_suffix: str, converter, **options) -> tuple[bytes, str]:
    suffix = Path(uploaded_file.name).suffix or ".image"
    stem = _safe_stem(uploaded_file.name)

    with tempfile.TemporaryDirectory() as directory:
        directory_path = Path(directory)
        image_path = directory_path / f"upload{suffix}"
        output_path = directory_path / f"{stem}_{output_suffix}.xlsx"

        image_path.write_bytes(uploaded_file.getvalue())
        converter(image_path=image_path, output_path=output_path, **options)
        return output_path.read_bytes(), output_path.name


def _canvas_label(key: str) -> str:
    if key == "":
        return "Original image size"
    preset = CANVAS_PRESETS[key]
    return f"{preset.label} ({preset.width_mm:g} x {preset.height_mm:g} mm)"


def _safe_stem(filename: str) -> str:
    stem = Path(filename).stem.lower()
    normalized = "".join(character if character.isalnum() else "-" for character in stem)
    normalized = "-".join(part for part in normalized.split("-") if part)
    return normalized or "image"


if __name__ == "__main__":
    main()
