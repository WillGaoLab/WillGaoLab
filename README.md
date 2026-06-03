# WillGaoLab Excel Pixel Art Generator

Streamlit app for generating Excel paint-by-number workbooks from uploaded images.

## Use Online

Open the live app:

https://willgaolab-dvy5xga3u2xexllw7lei82.streamlit.app/

Upload an image, choose paper size, set exact Excel-cell resolution, choose indexed color count, and download an `.xlsx` workbook.

## Streamlit Community Cloud

Use this repository in Streamlit Community Cloud with:

- Repository: `WillGaoLab/WillGaoLab`
- Branch: `main`
- Main file path: `streamlit_app.py`

Each workbook includes:

- `Reference` - finished colored pixel art.
- `Template` - paper-style paint-by-number sheet with light gray numbers.
- `Color Index` - indexed colors with swatches and cell counts.

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

## Personal Use and Legal Notice

This app is provided for personal, educational, research, and non-commercial creative use only.

Do not use uploaded images, generated templates, generated workbooks, or derivative outputs for commercial sale, paid products, merchandise, advertising, client work, or other revenue-generating activity unless you have independently secured all required rights and permissions.

Users are solely responsible for ensuring they have the legal right to upload, transform, distribute, print, share, or otherwise use any image processed with this app. The maintainers do not claim ownership of user-uploaded images or generated workbooks, and do not grant rights to third-party images.

Generated outputs may still be subject to copyright, trademark, privacy, publicity, moral rights, license terms, or other legal restrictions. The app is provided as-is, without warranty. The maintainers are not liable for misuse, infringement claims, losses, damages, takedown requests, printing costs, or other consequences arising from use of the app or generated outputs.

This notice is not legal advice. For commercial use, public distribution, uncertain image rights, or jurisdiction-specific questions, consult a qualified legal professional before using the image or output.
