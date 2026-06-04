"""V3 vector poster-tile rendering for Physical Layer printable PDFs."""

from __future__ import annotations

from dataclasses import dataclass

from pathlib import Path

from PIL import Image
from reportlab.lib.colors import Color, HexColor, black, white
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen.canvas import Canvas

A4_PAGE_SIZE = A4
PAGE_MARGIN = 36
HEADER_HEIGHT = 54
COORDINATE_GUTTER = 26


@dataclass(frozen=True)
class PosterTile:
    """A poster tile and its global canvas bounds."""

    image: Image.Image
    page_column: int
    page_row: int
    left: int
    top: int
    right: int
    bottom: int


@dataclass(frozen=True)
class PosterSplitSummary:
    """Validation result for splitting a master cell canvas into equal tiles."""

    master_resolution: tuple[int, int]
    poster_pages: tuple[int, int]
    cells_per_page: tuple[int, int] | None
    valid: bool
    reason: str | None


def validate_poster_split(
    master_resolution: tuple[int, int],
    poster_pages: tuple[int, int],
) -> PosterSplitSummary:
    """Verify that a master canvas divides into equal-sized poster tiles."""
    width, height = master_resolution
    pages_wide, pages_tall = poster_pages
    if pages_wide < 1 or pages_tall < 1:
        return PosterSplitSummary(master_resolution, poster_pages, None, False, "Page dimensions must be at least 1.")

    reasons = []
    if width % pages_wide:
        reasons.append(f"Width {width} is not divisible by {pages_wide}.")
    if height % pages_tall:
        reasons.append(f"Height {height} is not divisible by {pages_tall}.")
    if reasons:
        return PosterSplitSummary(master_resolution, poster_pages, None, False, " ".join(reasons))

    return PosterSplitSummary(
        master_resolution,
        poster_pages,
        (width // pages_wide, height // pages_tall),
        True,
        None,
    )


def write_vector_poster_pdf(
    image: Image.Image,
    output_path: str | Path,
    poster_pages: tuple[int, int],
    page_size: tuple[float, float] = A4_PAGE_SIZE,
) -> Path:
    """Write three sharp vector poster page sets to a true A4 PDF."""
    output_path = Path(output_path)
    rgb_image = image.convert("RGB")
    palette = _build_palette(rgb_image)
    tiles = split_poster_tiles(rgb_image, poster_pages)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    pdf = Canvas(str(output_path), pagesize=page_size, pageCompression=1)
    pdf.setTitle("Physical Layer Poster Split")

    for tile in tiles:
        _draw_vector_tile_page(pdf, tile, palette, poster_pages, page_size, fill_colors=True)
        pdf.showPage()
    for tile in tiles:
        _draw_vector_reference_page(pdf, tile, page_size)
        pdf.showPage()
    for tile in tiles:
        _draw_vector_tile_page(pdf, tile, palette, poster_pages, page_size, fill_colors=False)
        pdf.showPage()

    pdf.save()
    return output_path


def split_poster_tiles(image: Image.Image, poster_pages: tuple[int, int]) -> list[PosterTile]:
    """Split an image into row-major poster tiles while preserving global bounds."""
    summary = validate_poster_split(image.size, poster_pages)
    if not summary.valid:
        raise ValueError(summary.reason)
    pages_wide, pages_tall = poster_pages

    tiles = []
    for page_row in range(pages_tall):
        for page_column in range(pages_wide):
            left = round(image.width * page_column / pages_wide)
            right = round(image.width * (page_column + 1) / pages_wide)
            top = round(image.height * page_row / pages_tall)
            bottom = round(image.height * (page_row + 1) / pages_tall)
            tiles.append(
                PosterTile(
                    image=image.crop((left, top, right, bottom)),
                    page_column=page_column,
                    page_row=page_row,
                    left=left,
                    top=top,
                    right=right,
                    bottom=bottom,
                )
            )
    return tiles


def _draw_vector_tile_page(
    pdf: Canvas,
    tile: PosterTile,
    palette: dict[tuple[int, int, int], int],
    poster_pages: tuple[int, int],
    page_size: tuple[float, float],
    fill_colors: bool,
) -> None:
    page_width, page_height = page_size
    _draw_page_background(pdf, page_size)
    tile_width, tile_height = tile.image.size

    available_width = page_width - PAGE_MARGIN * 2 - COORDINATE_GUTTER
    available_height = page_height - PAGE_MARGIN * 2 - HEADER_HEIGHT - COORDINATE_GUTTER
    cell_size = min(available_width / tile_width, available_height / tile_height)
    grid_width = cell_size * tile_width
    grid_height = cell_size * tile_height
    grid_left = (page_width - grid_width + COORDINATE_GUTTER) / 2
    grid_bottom = PAGE_MARGIN + (available_height - grid_height) / 2

    title_size = 12
    subtitle_size = 8
    coordinate_size = max(4, min(8, cell_size * 0.48))
    number_size = max(3, min(8, cell_size * 0.52))

    tile_name = f"{_column_name(tile.page_column + 1)}{tile.page_row + 1}"
    pages_wide, pages_tall = poster_pages
    title = f"Tile {tile_name}  |  Poster {pages_wide} x {pages_tall}"
    subtitle = f"Columns {tile.left + 1}-{tile.right}  |  Rows {tile.top + 1}-{tile.bottom}"
    pdf.setFillColor(black)
    pdf.setFont("Helvetica-Bold", title_size)
    pdf.drawString(PAGE_MARGIN, page_height - PAGE_MARGIN, title)
    pdf.setFillColor(HexColor("#555555"))
    pdf.setFont("Helvetica", subtitle_size)
    pdf.drawString(PAGE_MARGIN, page_height - PAGE_MARGIN - 16, subtitle)

    pixels = tile.image.load()
    for row in range(tile_height):
        for column in range(tile_width):
            color = pixels[column, row]
            x0 = grid_left + column * cell_size
            y0 = grid_bottom + (tile_height - row - 1) * cell_size
            pdf.setFillColor(_pdf_color(color) if fill_colors else white)
            pdf.rect(x0, y0, cell_size, cell_size, stroke=0, fill=1)
            number = str(palette[color])
            pdf.setFillColor(white if fill_colors and _luminance(color) < 120 else black)
            pdf.setFont("Helvetica", number_size)
            pdf.drawCentredString(x0 + cell_size / 2, y0 + (cell_size - number_size) / 2, number)

    pdf.setStrokeColor(HexColor("#333333"))
    pdf.setLineWidth(0.25)
    for column in range(tile_width + 1):
        x = grid_left + column * cell_size
        pdf.line(x, grid_bottom, x, grid_bottom + grid_height)
    for row in range(tile_height + 1):
        y = grid_bottom + row * cell_size
        pdf.line(grid_left, y, grid_left + grid_width, y)
    pdf.setStrokeColor(black)
    pdf.setLineWidth(1)
    pdf.rect(grid_left, grid_bottom, grid_width, grid_height, stroke=1, fill=0)

    pdf.setFillColor(black)
    pdf.setFont("Helvetica", coordinate_size)
    for column in _coordinate_positions(tile_width):
        label = str(tile.left + column + 1)
        x = grid_left + column * cell_size + cell_size / 2
        pdf.drawCentredString(x, grid_bottom + grid_height + 7, label)
    for row in _coordinate_positions(tile_height):
        label = str(tile.top + row + 1)
        y = grid_bottom + (tile_height - row - 0.5) * cell_size - coordinate_size / 3
        pdf.drawRightString(grid_left - 5, y, label)


def _draw_vector_reference_page(
    pdf: Canvas,
    tile: PosterTile,
    page_size: tuple[float, float],
) -> None:
    """Draw only colored vector cell rectangles without production markings."""
    page_width, page_height = page_size
    _draw_page_background(pdf, page_size)
    available_width = page_width - PAGE_MARGIN * 2
    available_height = page_height - PAGE_MARGIN * 2
    scale = min(available_width / tile.image.width, available_height / tile.image.height)
    rendered_width = tile.image.width * scale
    rendered_height = tile.image.height * scale
    left = (page_width - rendered_width) / 2
    bottom = (page_height - rendered_height) / 2
    pixels = tile.image.load()
    for row in range(tile.image.height):
        for column in range(tile.image.width):
            pdf.setFillColor(_pdf_color(pixels[column, row]))
            pdf.rect(
                left + column * scale,
                bottom + (tile.image.height - row - 1) * scale,
                scale,
                scale,
                stroke=0,
                fill=1,
            )


def _build_palette(image: Image.Image) -> dict[tuple[int, int, int], int]:
    counts: dict[tuple[int, int, int], int] = {}
    for color in image.getdata():
        counts[color] = counts.get(color, 0) + 1
    ordered = sorted(counts, key=lambda color: (-counts[color], color))
    return {color: number for number, color in enumerate(ordered, start=1)}


def _coordinate_positions(length: int) -> list[int]:
    if length <= 12:
        return list(range(length))
    step = 5 if length <= 60 else 10
    positions = list(range(0, length, step))
    if length - 1 not in positions:
        positions.append(length - 1)
    return positions


def _pdf_color(color: tuple[int, int, int]) -> Color:
    return Color(color[0] / 255, color[1] / 255, color[2] / 255)


def _draw_page_background(pdf: Canvas, page_size: tuple[float, float]) -> None:
    pdf.setFillColor(white)
    pdf.rect(0, 0, page_size[0], page_size[1], stroke=0, fill=1)


def _luminance(color: tuple[int, int, int]) -> float:
    red, green, blue = color
    return 0.2126 * red + 0.7152 * green + 0.0722 * blue


def _column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name
