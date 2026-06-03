"""Canvas and paper preset definitions."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class CanvasPreset:
    """A paper-like canvas preset used to size the Excel pixel grid."""

    key: str
    label: str
    width_mm: float
    height_mm: float
    excel_paper_size: int | None = None


CANVAS_PRESETS: dict[str, CanvasPreset] = {
    "a0": CanvasPreset("a0", "A0", 841, 1189, None),
    "a1": CanvasPreset("a1", "A1", 594, 841, None),
    "a2": CanvasPreset("a2", "A2", 420, 594, None),
    "a3": CanvasPreset("a3", "A3", 297, 420, 8),
    "a4": CanvasPreset("a4", "A4", 210, 297, 9),
    "a5": CanvasPreset("a5", "A5", 148, 210, 11),
    "a6": CanvasPreset("a6", "A6", 105, 148, None),
    "b0": CanvasPreset("b0", "B0", 1000, 1414, None),
    "b1": CanvasPreset("b1", "B1", 707, 1000, None),
    "b2": CanvasPreset("b2", "B2", 500, 707, None),
    "b3": CanvasPreset("b3", "B3", 353, 500, None),
    "b4": CanvasPreset("b4", "B4", 250, 353, 12),
    "b5": CanvasPreset("b5", "B5", 176, 250, 13),
    "b6": CanvasPreset("b6", "B6", 125, 176, None),
    "letter": CanvasPreset("letter", "Letter", 215.9, 279.4, 1),
    "legal": CanvasPreset("legal", "Legal", 215.9, 355.6, 5),
    "executive": CanvasPreset("executive", "Executive", 190.5, 254.0, 7),
    "folio": CanvasPreset("folio", "Folio", 210.0, 330.0, None),
    "ledger": CanvasPreset("ledger", "Ledger", 432.0, 279.0, 4),
    "tabloid": CanvasPreset("tabloid", "Tabloid", 279.0, 432.0, 3),
    "square": CanvasPreset("square", "Square", 210, 210, None),
}

FIT_MODES = {"contain", "cover"}
ORIENTATIONS = {"auto", "portrait", "landscape"}
