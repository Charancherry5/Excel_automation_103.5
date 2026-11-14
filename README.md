# Fibre test workbook app

This Streamlit app provides an editable interface for the two connected sheets in `../CRINCRINTR.xlsx`:

- **Insertion Loss Sheet** — job metadata, endpoints, instrument details, counts, thresholds, and one fibre's measurements.
- **Line Diagram** — the 17 existing route-event rows, editable in a table. The sheet header is linked to the insertion-loss inputs.

The app does not use the other workbook sheets. Downloading creates a copy that contains only the two sheets above. Workbook formulas and named ranges are retained; formulas are set to recalculate when the downloaded file opens in Excel. The source workbook is not overwritten.

The in-app loss preview recalculates directional/average insertion loss, link loss, IL–LL differences, and calculated maximums from the current interface inputs. It does not replace the workbook's pass/fail formula, which already contains broken `#REF!` references. The current form supports the single populated fibre record present in the source workbook. Additional scenario controls can be added later.

## Run

From this folder, double-click `run_app.bat`, or run:

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m streamlit run app.py
```

The app expects `CRINCRINTR.xlsx` in the parent folder.
