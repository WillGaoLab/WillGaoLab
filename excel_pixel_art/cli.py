"""Command line interface for Excel pixel art generation."""

from __future__ import annotations

import argparse
from pathlib import Path

from .canvas import CANVAS_PRESETS, FIT_MODES, ORIENTATIONS


def parse_resolution(value: str) -> tuple[int, int]:
    normalized = value.lower().replace(" ", "")
    if "x" not in normalized:
        raise argparse.ArgumentTypeError("resolution must use WIDTHxHEIGHT, for example 160x100")

    width_text, height_text = normalized.split("x", 1)
    try:
        width = int(width_text)
        height = int(height_text)
    except ValueError as error:
        raise argparse.ArgumentTypeError("resolution width and height must be whole numbers") from error

    if width < 1 or height < 1:
        raise argparse.ArgumentTypeError("resolution width and height must be at least 1")
    if width > 2000 or height > 2000:
        raise argparse.ArgumentTypeError("resolution width and height must be at most 2000")
    return width, height


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-pixel-art",
        description="Convert an image into an Excel workbook where cells act as pixels.",
    )
    parser.add_argument("image", type=Path, help="Input image path")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output .xlsx path. Defaults to <image-stem>_pixel_art.xlsx",
    )
    parser.add_argument(
        "--max-size",
        type=int,
        default=96,
        help="Maximum width or height in Excel cells. Default: 96",
    )
    parser.add_argument(
        "--cell-size",
        type=float,
        default=3.0,
        help="Excel cell size used for both row height and column width. Default: 3.0",
    )
    parser.add_argument(
        "--colors",
        type=int,
        default=48,
        help="Maximum number of indexed colors to use in the template. Default: 48",
    )
    parser.add_argument(
        "--canvas-size",
        choices=sorted(CANVAS_PRESETS),
        help="Paper canvas preset. Choices include a4, letter, b5, square, and more.",
    )
    parser.add_argument(
        "--resolution",
        type=parse_resolution,
        help="Exact Excel cell resolution as WIDTHxHEIGHT, for example 160x100. Overrides --max-size.",
    )
    parser.add_argument(
        "--orientation",
        choices=sorted(ORIENTATIONS),
        default="auto",
        help="Canvas orientation when a paper preset is selected. Default: auto",
    )
    parser.add_argument(
        "--fit",
        choices=sorted(FIT_MODES),
        default="contain",
        help="How the image fits the canvas: contain preserves the whole image, cover fills and crops. Default: contain",
    )
    parser.add_argument(
        "--background-color",
        default="FFFFFF",
        help="Background color for paper canvas margins as a 6-digit hex color. Default: FFFFFF",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    output = args.output or args.image.with_name(f"{args.image.stem}_pixel_art.xlsx")

    if not args.image.exists():
        print(f"Error: Input image not found: {args.image}")
        return 1

    try:
        from .converter import image_to_excel

        image_to_excel(
            image_path=args.image,
            output_path=output,
            max_size=args.max_size,
            cell_size=args.cell_size,
            color_count=args.colors,
            canvas_size=args.canvas_size,
            resolution=args.resolution,
            orientation=args.orientation,
            fit=args.fit,
            background_color=args.background_color,
        )
    except ModuleNotFoundError as error:
        dependency = error.name or "required dependency"
        print(f"Missing dependency: {dependency}. Run `python -m pip install -e .`.")
        return 1
    except (FileNotFoundError, ValueError) as error:
        print(f"Error: {error}")
        return 1

    print(f"Wrote {output}")
    return 0
