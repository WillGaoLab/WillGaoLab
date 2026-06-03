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
