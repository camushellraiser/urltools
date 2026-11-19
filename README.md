# Thermo Fisher | URL Tools

Streamlit application with three tabs:

1. **URL Converter** — Multi-sheet Marketing URL conversions, Product Inclusion Lists, individual Excel downloads and ZIP download.
2. **AEM URLs** — First-sheet translation workbook processing, Local/Regional/Global Author and Rollout URLs, LATAM destination selection, and formatted Excel export.
3. **Naming Converter** — Paste a Wordbee request, edit its project name and requester, and generate Wordbee, per-language AEM, and Shared Folder names using GTS ID, detected languages and Marketing/Product content types from Tab 1.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy

Push the contents of this folder (including `.streamlit/config.toml`) to the root of your Git repository. Configure `app.py` as the Streamlit entrypoint.

## Notes

- Load an Excel workbook in Tab 1 before using the Naming Converter. The detected languages and Marketing/Product flags come from its processed records, not from free text pasted into Tab 3.
- Wordbee request parsing supports a copied tab-delimited row or Markdown table with title, order number, company and requester. The editable fields can be corrected before conversion.
- Reset in any tab clears all three tabs, including uploads and GTS ID, as with a fresh application session.
- The GTS placeholder `GTS2600XX` must be replaced with a numeric GTS ID to enable Convert.
- No external API keys are required. Deploy only in an approved environment for internal workbooks.
