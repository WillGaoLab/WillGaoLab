# WillGaoLab Excel Pixel Art Generator

Streamlit app for generating independent Digital and Physical pixel-art production outputs from uploaded images.

Excel Pixel Art Generator is a WillGaoLab product made by William (Peidong) Gao.

## Use Online

Open the live app:

https://willgaolab-dvy5xga3u2xexllw7lei82.streamlit.app/

Upload an image and independently generate:

- Digital Excel paint-by-number workbooks.
- Physical production ZIPs containing selected workbook, printable PDF, and color-mask outputs.

Version 3 keeps the independent Digital and Physical workflows and adds production-ready Poster Split output.

V3 includes:

- Project-wide `32 x 32` and 24-color interactive defaults.
- Poster Split divisibility validation with cells-per-page summaries.
- True A4 printable poster pages.
- Global coordinates and one global color index across poster tiles.
- Three printable PDF sets: colored numbered-grid tiles, clean references, and blank numbered templates.
- Sharp ReportLab vector PDFs with vector cells, grids, and real PDF text.
- Adaptive, LEGO, and Liquitex Basics 24 material palette modes.

## Streamlit Community Cloud

Use this repository in Streamlit Community Cloud with:

- Repository: `WillGaoLab/WillGaoLab`
- Branch: `main`
- Main file path: `streamlit_app.py`

Digital workbooks include:

- `Reference` - finished colored pixel art.
- `Template` - paper-style paint-by-number sheet with light gray numbers.
- `Color Index` - indexed colors with swatches and cell counts.

Physical outputs support an independent print canvas, material palette, validated poster splitting, color masks, workbook output, true-A4 vector printable PDF output, and mask output.

## Print Mode and Material Palette Disclaimer

- Brand and product names are provided only as unaffiliated references. WillGaoLab and William (Peidong) Gao are not sponsored, endorsed, authorized by, or associated with any brand, manufacturer, retailer, or product listed in Print Mode or Material Palette.
- Material palette matches, color names, screen colors, quantities, poster layouts, masks, and other generated recommendations are estimates and may be inaccurate, incomplete, unavailable, or unsuitable for a user's intended project.
- Users must independently verify colors, dimensions, quantities, availability, compatibility, safety, costs, licensing, and all other requirements before purchasing materials, printing, assembling, or producing an output.
- All purchasing, printing, material selection, assembly, and production decisions are made entirely at the user's own risk and judgment.
- To the maximum extent permitted by law, WillGaoLab and William (Peidong) Gao accept no responsibility or liability for purchases, expenses, losses, waste, incorrect colors, inaccurate quantities, print errors, failed projects, injuries, damages, or any other consequences arising from Print Mode, Material Palette, or generated outputs.

## Local Run

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
streamlit run streamlit_app.py
```

Then open:

```text
http://127.0.0.1:8501
```

## Related

- Project repository: https://github.com/PeidongGao/excel-pixel-art-generator
- Personal website page: https://williampeidonggao.com/resources/excel-pixel-art-generator/

## Knowledge Base: Journal Harvester SOP

Journal Harvester SOP provides browser-based workflows for low-volume,
user-initiated export of publicly visible bibliographic metadata from Cell,
Nature, and Science.

- Canonical repository: https://github.com/PeidongGao/journal-harvester-sop
- Website project page: https://williampeidonggao.com/resources/journal-harvester-sop/
- Usage disclaimer: https://github.com/PeidongGao/journal-harvester-sop/blob/main/DISCLAIMER.md

The project does not publish scraped CSV datasets, publisher PDFs, full-text
articles, paywalled content, or copyrighted figures. Users are responsible for
reviewing and complying with applicable publisher terms, institutional
agreements, crawling policies, and law.

## Personal Use and Legal Notice

Excel Pixel Art Generator is a WillGaoLab product made by William (Peidong) Gao. The WillGaoLab name, project identity, documentation, source code organization, and product presentation are maintained by William (Peidong) Gao unless otherwise stated.

The hosted app uses Microsoft Clarity to process usage and technical analytics data. Microsoft Clarity and Streamlit Community Cloud may process data under their own terms and privacy policies.

This app is provided for personal, educational, research, and non-commercial creative use only.

Do not use uploaded images, generated templates, generated workbooks, or derivative outputs for commercial sale, paid products, merchandise, advertising, client work, or other revenue-generating activity unless you have independently secured all required rights and permissions.

Users are solely responsible for ensuring they have the legal right to upload, transform, distribute, print, share, or otherwise use any image processed with this app. The maintainers do not claim ownership of user-uploaded images or generated workbooks, and do not grant rights to third-party images.

Generated outputs may still be subject to copyright, trademark, privacy, publicity, moral rights, license terms, or other legal restrictions. The app is provided as-is, without warranty. The maintainers are not liable for misuse, infringement claims, losses, damages, takedown requests, printing costs, or other consequences arising from use of the app or generated outputs.

This notice is not legal advice. For commercial use, public distribution, uncertain image rights, or jurisdiction-specific questions, consult a qualified legal professional before using the image or output.
