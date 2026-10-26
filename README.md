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
Open http://localhost:8501. Optional: install LibreOffice so the app can recalculate and preview key figures; otherwise Excel recalculates on open.

## Using the app

1. **Company** - name, CIN, address, background, latest year end, authorised capital, face value, tax rate, closing stock (not carried in Tally), auditor, directors.
2. **Trial balances** - upload the latest, previous and earliest year Tally TB exports (Detailed view with opening / debit / credit / closing and Dr/Cr; ledgers in italics). The app checks Dr = Cr and that closing balances net to nil.
3. **Mapping review** - every ledger gets a suggested note line ("Masters") from Tally group names and keywords, separately for debit and credit balances. Correct anything, press *Save mapping*, and download the mapping JSON to reuse next time.
4. **Shareholders & ageing** - register of members (promoter flag) and receivables / payables ageing.
5. **Adjustments** - restatement journal entries in lakhs (entry no., date, master, Dr, Cr, narration).
6. **Generate** - download the RFS workbook; save the project JSON to come back later.

## Code

| Path | Purpose |
|---|---|
| `app.py` | Streamlit UI |
| `rfs/tally.py` | Reads Tally TB exports: Dr/Cr signs, group/ledger tree, repeated ledgers |
| `rfs/classify.py` | Rules: Tally group / ledger name -> Schedule III classification |
| `rfs/mapping.py` | Masters (note lines), suggestions, overrides |
| `rfs/builder.py` | Writes the workbook |
| `tests/smoke_test.py` | Builds a workbook from three TBs with a minimal config |

## Notes / limits
- Three financial years; a stub period is not yet supported.
- PPE is presented on net (WDV) movement from the TB; gross block needs the fixed asset register.
- Review the yellow cells (inputs) and grey check cells in the output before issue.
- Client files are not stored by the app beyond the session's temporary folder; `.xlsx` and project files are git-ignored.
