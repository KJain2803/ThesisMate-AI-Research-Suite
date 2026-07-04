# Repository Guidelines

## Project Structure & Module Organization

This repository is a compact Streamlit application. `app.py` contains the full UI, analytics workflow, file parsing, chart generation, and report export logic. `requirements.txt` pins the Python dependencies used by the app. `sample_data.xlsx` is the local fixture for exercising the spreadsheet analytics flow. `README.md` documents setup, deployment context, and the public demo. There is currently no separate `tests/`, `src/`, or static assets directory; create them only when the code is split into reusable modules.

## Build, Test, and Development Commands

Create and activate a virtual environment before running commands:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Run the app locally with:

```bash
streamlit run app.py
```

Install the Chrome runtime used by Kaleido for Plotly image export:

```bash
choreo_get_chrome
```

Use `sample_data.xlsx` to verify upload, statistics, chart, PDF, and presentation export paths.

## Coding Style & Naming Conventions

Use Python with 4-space indentation and descriptive `snake_case` for functions and variables, matching existing helpers such as `safe_plotly_save` and `extract_text_from_file`. Keep Streamlit UI code readable by grouping related controls, calculations, and output sections together. Prefer `pathlib.Path` for filesystem paths and temporary output locations. Avoid committing generated reports, temporary images, virtual environments, or local cache files.

## Testing Guidelines

No automated test framework is currently configured. For logic changes, add focused tests under `tests/` using `pytest`, especially for file parsing, statistical calculations, and report-generation helpers. Until tests exist, manually smoke test with `streamlit run app.py` and `sample_data.xlsx`; confirm app startup, Excel upload, statistical outputs, chart rendering, and export downloads.

## Commit & Pull Request Guidelines

The current Git history uses short summary commits such as `Update README.md`. Keep future commits concise and imperative, for example `Add pytest coverage for Excel analytics`. Pull requests should include a brief description, manual test notes, linked issues when applicable, and screenshots or screen recordings for visible Streamlit UI changes. Mention any dependency updates and whether `choreo_get_chrome` or deployment settings are affected.

## Security & Configuration Tips

Do not commit secrets, private research data, generated dossiers, or local `.venv` contents. Treat uploaded dissertation files and spreadsheets as sensitive data during development.
