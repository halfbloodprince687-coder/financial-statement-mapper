# Financial Statement -> MCA Filing Template Mapper

A Streamlit app that takes a company's financial statement workbook (the
"source") and auto-maps its figures into an MCA/ROC e-filing template (the
"destination" - the one with `FieldPLBS01`, `FieldPLBS02`... Field-ID
columns), with a review UI, an optional multiplication factor, basic
validation, and an export button.

## How it works

1. **Destination parsing** (`engine.parse_destination`) - scans every sheet
   for the `FieldPLBSxx` header row, then walks down the sheet recording
   every genuine Field-ID cell (e.g. `dom_manufracture`, `EQL_sharecap`) and
   the value cell immediately to its left. Header/label text like
   "Particulars" or "Total" is filtered out so it's never mistaken for a
   real Field ID.

2. **Source parsing** (`engine.parse_source`) - walks every row of the
   source workbook. It tracks which columns currently mean "2024-25",
   "Non-current portion", "Current maturities", etc. (from the nearest
   header row above), and which row-block it's inside (`Note 5 Long-term
   borrowings`). Every row with a text label and at least one number becomes
   a ledger entry carrying `{column header: value}` - this copes with
   simple 2-column current/previous tables and the wider 4-column note
   tables alike, without assuming a fixed layout.

3. **Matching** (`engine.suggest_matches`) - fuzzy string matching (with a
   keyword-overlap bonus) ranks the source ledger against each destination
   label. The UI shows the top 8 candidates; anything scoring 75%+ is
   pre-selected, everything else is left for you to pick.

4. **Current/Previous pairing** (`engine.build_logical_rows`) - the
   `PLCurrent`/`PLPrevious` and `BSCurrent`/`BSPrevious` sheet pairs share
   an identical row layout, so the app groups them into one logical line
   item: pick a source match once, and the app fills both sheets by sorting
   the matched item's values newest-year-first. `BSTrade`, `Financial` and
   `Capital` don't have that Current/Previous split baked in the same way,
   so their fields are reviewed individually.

5. **Multiplication factor** - applied per row, and only to fields that
   don't look like a ratio/EPS/percentage/share-count (heuristic in
   `_guess_scale_default`); you can flip the "Apply multiplier" checkbox on
   any row.

6. **Validation** (`validator.py`) - the destination template already has
   `=SUM(...)` formulas for its subtotal/total rows; the app never
   overwrites those (it writes only to genuine leaf cells and reports any
   attempted write to a formula cell as a skipped notice). It recomputes
   those formulas from what you've mapped so far and runs a Balance Sheet
   identity check (Total Equity & Liabilities vs Total Assets).

7. **Export** (`writer.py`) - writes into a fresh copy of the destination
   `.xlsx`, leaving every existing formula untouched, and offers it as a
   download.

## Important, honest limitations

- **The auto-suggestions are a starting point, not a final answer.** Real
  financial statements are laid out differently company to company, and
  some destination fields (e.g. a domestic/export turnover split) may not
  exist as separate lines in your source at all. Always check the
  Validation panel's recomputed totals against the source's own reported
  totals before exporting - a Total Revenue or Total Assets figure that's
  roughly double (or half) what you expect is the classic sign that two
  destination sub-lines both got matched to the same source total.
- The destination template must be a `.xlsx` file (Field-ID driven, as
  described above). If your master template is only saved as `.xls`, open
  it once in Excel or LibreOffice and "Save As" `.xlsx` - it's a one-time
  conversion since the same template is reused every filing period.
- The source file can be `.xls` or `.xlsx`.
- Validation here is a sanity check, not an audit. It does not replace a
  qualified accountant's review before an actual MCA filing.

## Running locally

**Easiest way** - from the `app/` folder, run the setup script for your OS.
It creates a `.venv` virtual environment, installs everything in
`requirements.txt` into *that* environment, and launches the app - this
avoids the classic "ModuleNotFoundError: xlrd" problem, which almost always
means `pip install` ran against a different Python than the one `streamlit`
is using.

macOS/Linux:
```bash
cd app
chmod +x setup.sh
./setup.sh
```

Windows: double-click `app\setup.bat`, or run it from a terminal:
```
cd app
setup.bat
```

**Manual way**, if you'd rather do it yourself:
```bash
cd app
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```
Whichever way you use, always run `streamlit` from *inside* the activated
`.venv` - running the global/system `streamlit` after installing packages
into a venv is the #1 cause of "module not found" errors.

### "Permission" error on "Use bundled sample"

This is almost always a Windows thing: files extracted from a `.zip` you
downloaded get marked "Blocked" by Windows, which shows up as a permission
error the first time something tries to read them. Fix: right-click the
downloaded `.zip` file (before extracting, or re-download it) -> Properties
-> tick **Unblock** -> OK -> then extract again. The app also now copies the
bundled sample into a fresh temp folder before reading it, and will show a
clear message (with this same tip) instead of a raw traceback if it still
can't read the files - in that case, just upload your own source/destination
files instead, which isn't affected by this.

## Deploying to streamlit.io (Streamlit Community Cloud)

1. Push this `app/` folder to a GitHub repo (keep `app.py`,
   `engine.py`, `writer.py`, `validator.py`, `requirements.txt`, and
   `sample_data/` all in the same folder).
2. Go to https://share.streamlit.io, click "New app", point it at your
   repo and set the main file path to `app.py`.
3. Deploy. Streamlit Cloud installs `requirements.txt` automatically.

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit UI - upload, review, validate, export |
| `engine.py` | Parsing (source & destination) and fuzzy matching |
| `writer.py` | Writes the mapped values into a template copy |
| `validator.py` | Recomputes template formulas / balance-sheet check |
| `sample_data/` | The HI Diamonds Pvt Ltd sample files, for trying the app out |

## Trying it out

Tick "Use bundled sample" in the sidebar to load the sample source and
destination files without uploading anything, so you can see how the review
UI behaves before pointing it at your own filings.
