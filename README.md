# Thermo Fisher | URL Tools

Streamlit application with two independent tabs:

- **URL Converter:** Converts marketing URLs, builds per-language product inclusion lists, and exports Excel files or a combined ZIP.
- **AEM URLs:** Builds Local Author, Regional Author, Regional Sites / Rollout, and Global Author URLs from the first worksheet of a translation request workbook, including LATAM destination selection.

## Requirements

- Python 3.10 or newer
- Dependencies listed in `requirements.txt`

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
# Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy with Streamlit Community Cloud

1. Push the contents of this folder to your Git repository (keep `app.py` and `requirements.txt` at the repository root).
2. Create a Streamlit app from the repository and select `app.py` as the entry point.
3. Deploy. No secrets or API keys are required by the current code.

**Important:** This tool processes uploaded workbooks. Only deploy to a hosting environment approved for the sensitivity of your company's data; do not upload confidential workbooks to unapproved public hosting.

## Files

- `app.py` — complete Streamlit app
- `requirements.txt` — Python dependencies
- `.streamlit/config.toml` — optional UI defaults
- `.gitignore` — ignores local Python and Streamlit artifacts
