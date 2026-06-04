"""Small local web UI for converting uploaded images."""

from __future__ import annotations

import argparse
import html
import re
import tempfile
from dataclasses import dataclass
from email import policy
from email.parser import BytesParser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from .canvas import CANVAS_PRESETS, FIT_MODES, ORIENTATIONS


MAX_UPLOAD_BYTES = 25 * 1024 * 1024


@dataclass(frozen=True)
class UploadedFile:
    filename: str
    content: bytes


@dataclass(frozen=True)
class MultipartForm:
    fields: dict[str, str]
    files: dict[str, UploadedFile]


class PixelArtRequestHandler(BaseHTTPRequestHandler):
    server_version = "ExcelPixelArt/0.1"

    def do_GET(self) -> None:
        if self.path not in {"/", "/index.html"}:
            self._send_text("Not found", HTTPStatus.NOT_FOUND)
            return

        self._send_html(_render_page())

    def do_POST(self) -> None:
        if self.path != "/convert":
            self._send_text("Not found", HTTPStatus.NOT_FOUND)
            return

        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            self._send_html(_render_page("Choose an image before converting."), HTTPStatus.BAD_REQUEST)
            return
        if content_length > MAX_UPLOAD_BYTES:
            self._send_html(_render_page("Upload is larger than 25 MB."), HTTPStatus.REQUEST_ENTITY_TOO_LARGE)
            return

        content_type = self.headers.get("Content-Type", "")
        if not content_type.startswith("multipart/form-data"):
            self._send_html(_render_page("Upload form data was not recognized."), HTTPStatus.BAD_REQUEST)
            return

        body = self.rfile.read(content_length)
        form = _parse_multipart_form(content_type, body)

        uploaded_file = form.files.get("image")
        if uploaded_file is None or not uploaded_file.filename:
            self._send_html(_render_page("Choose an image before converting."), HTTPStatus.BAD_REQUEST)
            return

        try:
            max_size = _parse_int_field(form, "max_size", default=32)
            cell_size = _parse_float_field(form, "cell_size", default=3.0)
            color_count = _parse_int_field(form, "color_count", default=24, minimum=2, maximum=256)
            canvas_size = _parse_choice_field(form, "canvas_size", choices={"", *CANVAS_PRESETS}, default="")
            resolution = _parse_optional_resolution(form)
            orientation = _parse_choice_field(form, "orientation", choices=ORIENTATIONS, default="auto")
            fit = _parse_choice_field(form, "fit", choices=FIT_MODES, default="contain")
            background_color = form.fields.get("background_color", "FFFFFF")
        except ValueError as error:
            self._send_html(_render_page(str(error)), HTTPStatus.BAD_REQUEST)
            return

        filename = Path(uploaded_file.filename).name
        stem = _safe_stem(filename)
        suffix = Path(filename).suffix or ".image"

        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            image_path = directory_path / f"upload{suffix}"
            output_path = directory_path / f"{stem}_pixel_art.xlsx"

            with image_path.open("wb") as image_file:
                image_file.write(uploaded_file.content)

            try:
                from .converter import image_to_excel

                image_to_excel(
                    image_path,
                    output_path,
                    max_size=max_size,
                    cell_size=cell_size,
                    color_count=color_count,
                    canvas_size=canvas_size or None,
                    resolution=resolution,
                    orientation=orientation,
                    fit=fit,
                    background_color=background_color,
                )
            except ModuleNotFoundError as error:
                dependency = error.name or "required dependency"
                self._send_html(
                    _render_page(f"Missing dependency: {dependency}. Run python -m pip install -e ."),
                    HTTPStatus.INTERNAL_SERVER_ERROR,
                )
                return
            except Exception as error:
                self._send_html(_render_page(f"Could not convert image: {error}"), HTTPStatus.BAD_REQUEST)
                return

            workbook_bytes = output_path.read_bytes()

        download_name = f"{stem}_pixel_art.xlsx"
        self.send_response(HTTPStatus.OK)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header("Content-Length", str(len(workbook_bytes)))
        self.send_header("Content-Disposition", f'attachment; filename="{download_name}"')
        self.end_headers()
        self.wfile.write(workbook_bytes)

    def log_message(self, format: str, *args: object) -> None:
        print(f"{self.address_string()} - {format % args}")

    def _send_html(self, body: str, status: HTTPStatus = HTTPStatus.OK) -> None:
        encoded = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def _send_text(self, body: str, status: HTTPStatus = HTTPStatus.OK) -> None:
        encoded = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)


def _parse_multipart_form(content_type: str, body: bytes) -> MultipartForm:
    header = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8")
    message = BytesParser(policy=policy.default).parsebytes(header + body)
    fields: dict[str, str] = {}
    files: dict[str, UploadedFile] = {}

    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue

        payload = part.get_payload(decode=True) or b""
        filename = part.get_filename()
        if filename:
            files[name] = UploadedFile(filename=filename, content=payload)
        else:
            charset = part.get_content_charset() or "utf-8"
            fields[name] = payload.decode(charset, errors="replace")

    return MultipartForm(fields=fields, files=files)


def _parse_int_field(
    form: MultipartForm,
    name: str,
    default: int,
    minimum: int = 1,
    maximum: int = 512,
) -> int:
    value = form.fields.get(name, str(default))
    try:
        parsed = int(value)
    except (TypeError, ValueError) as error:
        raise ValueError(f"{name.replace('_', ' ').title()} must be a whole number.") from error

    if parsed < minimum or parsed > maximum:
        raise ValueError(f"{name.replace('_', ' ').title()} must be between {minimum} and {maximum}.")
    return parsed


def _parse_float_field(form: MultipartForm, name: str, default: float) -> float:
    value = form.fields.get(name, str(default))
    try:
        parsed = float(value)
    except (TypeError, ValueError) as error:
        raise ValueError(f"{name.replace('_', ' ').title()} must be a number.") from error

    if parsed <= 0 or parsed > 20:
        raise ValueError(f"{name.replace('_', ' ').title()} must be greater than 0 and at most 20.")
    return parsed


def _parse_choice_field(form: MultipartForm, name: str, choices: set[str], default: str) -> str:
    value = form.fields.get(name, default).lower()
    if value not in choices:
        choice_text = ", ".join(sorted(choice or "none" for choice in choices))
        raise ValueError(f"{name.replace('_', ' ').title()} must be one of: {choice_text}.")
    return value


def _parse_optional_resolution(form: MultipartForm) -> tuple[int, int] | None:
    width_text = form.fields.get("resolution_width", "").strip()
    height_text = form.fields.get("resolution_height", "").strip()
    if width_text == "" and height_text == "":
        return None
    if width_text == "" or height_text == "":
        raise ValueError("Custom resolution needs both width cells and height cells.")

    try:
        width = int(width_text)
        height = int(height_text)
    except ValueError as error:
        raise ValueError("Custom resolution width and height must be whole numbers.") from error

    if width < 1 or height < 1:
        raise ValueError("Custom resolution width and height must be at least 1.")
    if width > 2000 or height > 2000:
        raise ValueError("Custom resolution width and height must be at most 2000.")
    return width, height


def _safe_stem(filename: str) -> str:
    stem = Path(filename).stem.lower()
    stem = re.sub(r"[^a-z0-9]+", "-", stem).strip("-")
    return stem or "image"


def _render_page(error: str | None = None) -> str:
    error_html = f'<p class="alert">{html.escape(error)}</p>' if error else ""
    canvas_options = _render_select_options(
        [("", "Original image size")]
        + [(key, f"{preset.label} ({_format_size(preset.width_mm, preset.height_mm)})") for key, preset in CANVAS_PRESETS.items()],
        selected="a4",
    )
    orientation_options = _render_select_options(
        [("auto", "Auto"), ("portrait", "Portrait"), ("landscape", "Landscape")],
        selected="auto",
    )
    fit_options = _render_select_options(
        [("contain", "Contain"), ("cover", "Cover")],
        selected="contain",
    )
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Excel Pixel Art Generator</title>
  <style>
    :root {{
      color-scheme: light;
      --paper: #f7f3ea;
      --ink: #172033;
      --muted: #5f6673;
      --line: #ccd4cf;
      --panel: #ffffff;
      --accent: #0b6b74;
      --accent-dark: #064b52;
      --warn: #a43b31;
      --gold: #c69b35;
    }}

    * {{
      box-sizing: border-box;
    }}

    body {{
      margin: 0;
      min-height: 100vh;
      background: var(--paper);
      color: var(--ink);
      font-family: ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      letter-spacing: 0;
    }}

    main {{
      width: min(960px, calc(100% - 32px));
      margin: 0 auto;
      padding: 48px 0;
    }}

    header {{
      display: grid;
      gap: 10px;
      margin-bottom: 28px;
    }}

    h1 {{
      margin: 0;
      font-size: clamp(2rem, 6vw, 4.5rem);
      line-height: 0.95;
      max-width: 760px;
    }}

    .subhead {{
      margin: 0;
      color: var(--muted);
      font-size: 1rem;
      max-width: 640px;
    }}

    form {{
      display: grid;
      gap: 18px;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 22px;
      box-shadow: 0 14px 40px rgb(23 32 51 / 0.08);
    }}

    .dropzone {{
      display: grid;
      place-items: center;
      min-height: 220px;
      padding: 22px;
      border: 2px dashed var(--accent);
      border-radius: 8px;
      background: #eef8f5;
      text-align: center;
    }}

    .dropzone input {{
      width: min(100%, 360px);
      color: var(--ink);
    }}

    .controls {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 14px;
    }}

    label {{
      display: grid;
      gap: 7px;
      color: var(--muted);
      font-size: 0.88rem;
      font-weight: 700;
    }}

    input[type="number"],
    input[type="text"],
    select {{
      width: 100%;
      min-height: 44px;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 0 12px;
      color: var(--ink);
      font: inherit;
    }}

    button {{
      width: fit-content;
      min-height: 46px;
      border: 0;
      border-radius: 6px;
      padding: 0 18px;
      background: var(--accent);
      color: #fff;
      font: inherit;
      font-weight: 800;
      cursor: pointer;
    }}

    button:hover {{
      background: var(--accent-dark);
    }}

    .alert {{
      margin: 0;
      border-left: 4px solid var(--warn);
      background: #fff3ef;
      padding: 12px 14px;
      color: var(--warn);
      font-weight: 700;
    }}

    .sample {{
      margin-top: 28px;
      border-top: 2px solid var(--gold);
      padding-top: 18px;
      color: var(--muted);
      font-size: 0.95rem;
    }}

    @media (max-width: 680px) {{
      main {{
        width: min(100% - 24px, 960px);
        padding: 28px 0;
      }}

      form {{
        padding: 16px;
      }}

      .controls {{
        grid-template-columns: 1fr;
      }}

      button {{
        width: 100%;
      }}
    }}
  </style>
</head>
<body>
  <main>
    <header>
      <h1>Excel Pixel Art Generator</h1>
      <p class="subhead">Upload an image and download an Excel workbook built from colored cells.</p>
    </header>
    <form action="/convert" method="post" enctype="multipart/form-data">
      {error_html}
      <div class="dropzone">
        <input type="file" name="image" accept="image/*" required>
      </div>
      <div class="controls">
        <label>
          Canvas
          <select name="canvas_size">
            {canvas_options}
          </select>
        </label>
        <label>
          Orientation
          <select name="orientation">
            {orientation_options}
          </select>
        </label>
        <label>
          Fit
          <select name="fit">
            {fit_options}
          </select>
        </label>
        <label>
          Max cells
          <input type="number" name="max_size" min="1" max="512" value="32" required>
        </label>
        <label>
          Width cells
          <input type="number" name="resolution_width" min="1" max="2000" value="32" required>
        </label>
        <label>
          Height cells
          <input type="number" name="resolution_height" min="1" max="2000" value="32" required>
        </label>
        <label>
          Cell size
          <input type="number" name="cell_size" min="0.1" max="20" step="0.1" value="3.0" required>
        </label>
        <label>
          Colors
          <input type="number" name="color_count" min="2" max="256" value="24" required>
        </label>
        <label>
          Background
          <input type="text" name="background_color" value="FFFFFF" maxlength="7" required>
        </label>
      </div>
      <button type="submit">Generate Workbook</button>
    </form>
    <p class="sample">Try the included sample at image/UnderTheWaveOffKanagawa.jpg from the command line.</p>
  </main>
</body>
</html>"""


def _render_select_options(options: list[tuple[str, str]], selected: str) -> str:
    return "\n".join(
        f'<option value="{html.escape(value)}"{_selected(value, selected)}>{html.escape(label)}</option>'
        for value, label in options
    )


def _selected(value: str, selected: str) -> str:
    return " selected" if value == selected else ""


def _format_size(width_mm: float, height_mm: float) -> str:
    return f"{width_mm:g} x {height_mm:g} mm"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run the Excel pixel art upload web app.")
    parser.add_argument("--host", default="127.0.0.1", help="Host to bind. Default: 127.0.0.1")
    parser.add_argument("--port", type=int, default=8000, help="Port to bind. Default: 8000")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    server = ThreadingHTTPServer((args.host, args.port), PixelArtRequestHandler)
    url = f"http://{args.host}:{args.port}"
    print(f"Serving Excel Pixel Art Generator at {url}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping server")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
