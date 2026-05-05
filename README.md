# ExcelParser

Small Streamlit app and command-line parser for converting the EMEA Excel matrix into JSON for GPT/source retrieval.

## Run in GitHub Codespaces

1. Open this repo in GitHub.
2. Click **Code** > **Codespaces** > **Create codespace on main**.
3. In the terminal run:

```bash
pip install -r requirements.txt
streamlit run app.py
```

4. Upload your Excel file in the browser UI.
5. The uploaded source file is saved into `cache/` and the parsed JSON is saved into `parsed/`.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Command line

```bash
python -m excel_parser.parser cache/source.xlsx --sheet EMEA --output parsed/emea.json --pretty
```

## Current EMEA parsing logic

Configured in `excel_parser/logic.py` so it is easy to modify later.

- Table range: `B5:CN194`
- Column B is the index/row-title column.
- Row 5 contains suite version headers. Merged cells are expanded automatically.
- Rows 6 and 7 contain suite headers for EMEA. Merged cells are expanded automatically.
- Row 8 can contain merged lower-level headers and is captured as `column_header`.
- Data starts below row 8.
- Empty cells become `NO`.
- Cells with `X` become `YES`.
- Hyperlinks are captured as `{title, url}`.
- Asterisks are preserved in the value and flagged with footnote metadata.
- Group mapping currently configured:
  - `B9` title with child rows `B10:B11`
  - `B12` title with child rows `B13:B16`
  - `B17` title with child rows `B18:B59`

Validation is intentionally non-blocking for the first test run. The parser returns warnings instead of failing hard.
