# RFS Generator

Streamlit application that turns three years of **Tally trial balances** into a complete
**Restated Financial Statements (RFS)** workbook under the SEBI ICDR Regulations
(Indian GAAP, Schedule III Division I, Rs. in lakhs).

Output layout (every figure is a formula):

`TB sheets (Tally export + Group / Masters / Net Amount / In Lakhs)` ->
`Notes (SUMIF on Masters)` -> `BS_1 / PL_2 (note totals)` ->
`CFS_3, SAP_4, SOA_5, RPT, Gratuity & contingencies, Tax shelter, Accounting / Analytical ratios, Capitalisation, Additional disclosures`,
with `Adj Entry RFS` journal entries for restatement adjustments and source sheets for shareholders and ageing.

## Run

```
pip install -r requirements.txt
streamlit run app.py
```
Open http://localhost:8501. On Windows you can simply double-click `run_app.bat` (needs Python 3.11+ added to PATH). Optional: install LibreOffice so the app can recalculate and preview key figures; otherwise Excel recalculates on open.

## Using the app

Quickest way: drop all your files into **Upload input files (any format)** in the sidebar. Trial balances go to the right year (from the year in the file name or heading, e.g. `TB FY 2025-26.pdf`), other files go to Supporting documents, and a saved project `.json` re-opens the project. Or use the tabs:

1. **ICDR & DRHP check** - at the time of preparing the RFS:
   - *Check latest ICDR updates* reads the SEBI regulations / circulars listings and shows ICDR-related amendments newest first; anything dated after the date the layout was last reviewed is flagged in red.
   - *Check latest DRHP / RHP filings* lists recent offer documents from **SEBI** (filed with SEBI / RoC), **NSE** (Main board and Emerge) and **BSE** (Main board and BSE SME) - choose the sources above the button. Each row shows where it was found and every source listing the same company. NSE / BSE pages are read with a headless browser (`run_app.bat` installs it; manually: `pip install playwright` then `python -m playwright install chromium`). Page URLs can be changed in *Settings - source URLs*. **NSE direct:** the NSE offer documents page loads its list from an NSE data address; copy it once (Chrome / Edge: F12 > Network > Fetch/XHR > reload the page > the request that returns the companies > Request URL, starting https://www.nseindia.com/api/) into *NSE data address* and press *Save settings* - the app then reads NSE directly, and for a month fills `from_date` / `to_date` (or `{from}` / `{to}`) automatically. Settings are saved on the PC. Pick a month and press *Month-wise filings* to list that month's DRHP / RHP filings (last 24 months). *Find company locations* reads the registered office on each offer document's cover page (state checked against the India Post PIN directory; results remembered on the PC) so the list can be filtered by state, with a state-wise summary; pick one and *Download and compare*, or upload a DRHP / RHP (PDF, Word). The app finds the restated financial information section and shows which annexures it contains vs what the app produces, plus headings the app does not produce.
   - Needs internet access. If SEBI cannot be reached the app says so and shows the source links to open manually; listing URLs can be edited in the sidebar (*Settings - source URLs*).
   - The checks and the reference document are recorded in a "Reference & Documents" sheet in the output.
2. **Company** - name, CIN, address, background, latest year end, authorised capital, face value, tax rate, closing stock (not carried in Tally), auditor, directors.
3. **Trial balances** - upload the latest, previous and earliest year TB in **any format**: Excel (.xlsx / .xls), CSV, Word (.docx / .doc), PowerPoint (.pptx / .ppt), PDF or OpenDocument. A Tally Excel export is used as-is; other formats are read as a table (Particulars / Opening / Debit / Credit / Closing, Dr/Cr) and rebuilt into a Tally-style sheet. The app checks Dr = Cr and that closing balances net to nil (shown in red if not).
4. **Supporting documents** - upload any number of documents (sanction letters, actuarial reports, ageing, registers, audited financials...) in the same formats. Text is extracted, can be categorised and searched, and the list goes into the output.
5. **Mapping review** - every ledger gets a suggested note line ("Masters") from Tally group names and keywords, separately for debit and credit balances. Correct anything, press *Save mapping*, and download the mapping JSON to reuse next time.
6. **Shareholders & ageing** - register of members (promoter flag) and receivables / payables ageing.
7. **Adjustments** - restatement journal entries in lakhs (entry no., date, master, Dr, Cr, narration).
8. **Generate** - choose output formats and download:
   - **Excel (.xlsx)** - the formula-linked RFS workbook
   - **Word (.docx / .doc)** - all annexures and notes as a document
   - **PowerPoint (.pptx)** - summary deck (statements, adjustments, ratios, capitalisation)
   - **HTML** - single-page report with a sheet index

   Save the project JSON to come back later (the sidebar upload is only for re-opening a saved project).

Word / PowerPoint / HTML are rendered from calculated figures, so the app needs **LibreOffice** installed, or **Microsoft Excel** on Windows (via `pywin32`, installed from requirements). `.doc` uses LibreOffice or Microsoft Word.

## Code

| Path | Purpose |
|---|---|
| `app.py` | Streamlit UI |
| `rfs/tally.py` | Reads Tally TB exports (Dr/Cr signs, group/ledger tree, repeated ledgers) and TBs in other formats |
| `rfs/docreader.py` | Text and tables from Excel, CSV, Word, PowerPoint, PDF (legacy formats via LibreOffice / MS Office) |
| `rfs/regulatory.py` | Latest ICDR updates and DRHP / RHP filings from SEBI; DRHP restated-section comparison |
| `rfs/classify.py` | Generic rules: Tally standard groups and ledger-name keywords -> Schedule III note line (works for any company; review in the app) |
| `rfs/mapping.py` | Masters (note lines), suggestions, overrides |
| `rfs/builder.py` | Writes the workbook |
| `rfs/exporters.py` | Recalculates and exports to Word, PowerPoint and HTML |
| `tests/smoke_test.py` | Builds a workbook from three TBs with a minimal config |

## Notes / limits
- `.doc`, `.xls`, `.ppt` and OpenDocument inputs need LibreOffice (or Microsoft Office on Windows). Google Docs / Sheets: download as .docx / .xlsx first. Scanned (image-only) PDFs are not read - use a text PDF or the Excel export.
- A TB that is not a Tally Excel export should keep groups and ledgers distinguishable (indentation, bold / italics or group totals); a flat list where Tally repeats ledgers may not reconcile - the app shows this in red.
- The layout has no separate notes for investments or unsecured loans from non-directors: investments map to other current assets (fixed deposits to bank balances) and unsecured loans are shown lender-wise under loans from directors - change in the mapping review if needed.
- Three financial years; a stub period is not yet supported.
- PPE is presented on net (WDV) movement from the TB; gross block needs the fixed asset register.
- Review the yellow cells (inputs) and grey check cells in the output before issue.
- Client files are not stored by the app beyond the session's temporary folder; `.xlsx` and project files are git-ignored.
