import io
import os
import shutil
import tempfile

import streamlit as st

# Friendly dependency check - a raw ModuleNotFoundError traceback in the
# Streamlit UI is confusing; tell people exactly what to run instead.
_missing = []
for _mod, _pip_name in [("openpyxl", "openpyxl"), ("xlrd", "xlrd")]:
    try:
        __import__(_mod)
    except ImportError:
        _missing.append(_pip_name)
if _missing:
    st.error(
        "Missing required package(s): " + ", ".join(_missing) + ".\n\n"
        "This usually means the app isn't running inside the virtual environment "
        "that has `requirements.txt` installed. From the `app/` folder, run:\n\n"
        "```\npython -m venv .venv\n"
        + (".venv\\Scripts\\activate" if os.name == "nt" else "source .venv/bin/activate")
        + "\npip install -r requirements.txt\nstreamlit run app.py\n```"
    )
    st.stop()

import engine
import writer
import validator

st.set_page_config(page_title="Financial Statement Mapper", layout="wide")

APP_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE_SOURCE = os.path.join(APP_DIR, "sample_data", "Source_file.xlsx")
SAMPLE_DEST = os.path.join(APP_DIR, "sample_data", "destination_file.xlsx")

SHEET_GROUPS = [
    ("Profit & Loss", ["PLCurrent", "PLPrevious"]),
    ("Balance Sheet", ["BSCurrent", "BSPrevious"]),
    ("Trade Receivables", ["BSTrade"]),
    ("Financial Parameters", ["Financial"]),
    ("Share Capital Raised", ["Capital"]),
]

MULTIPLIER_OPTIONS = [1, 10, 100, 1000, 100000, 10000000]


# --------------------------------------------------------------------------
# Session state helpers
# --------------------------------------------------------------------------

def init_state():
    st.session_state.setdefault("mapping", {})       # row.key -> choice dict
    st.session_state.setdefault("dest_fields", None)
    st.session_state.setdefault("ledger", None)
    st.session_state.setdefault("logical_rows", None)
    st.session_state.setdefault("dest_path", None)
    st.session_state.setdefault("formula_cells", None)


def save_upload(uploaded_file) -> str | None:
    suffix = os.path.splitext(uploaded_file.name)[1]
    try:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(uploaded_file.getbuffer())
        tmp.close()
        return tmp.name
    except (PermissionError, OSError) as e:
        st.sidebar.error(f"Could not save '{uploaded_file.name}' to a temp location: {e}")
        return None


# --------------------------------------------------------------------------
# Sidebar - file inputs & multiplier
# --------------------------------------------------------------------------

def sidebar():
    st.sidebar.title("1. Load files")
    use_sample = st.sidebar.checkbox(
        "Use bundled sample (HI Diamonds Pvt Ltd)", value=False,
        help="Loads the sample source & destination files shipped with this app, "
             "handy for trying the tool out before uploading your own.",
    )

    source_file = st.sidebar.file_uploader("Source file (financial statement) - .xls or .xlsx", type=["xls", "xlsx"])
    dest_file = st.sidebar.file_uploader(
        "Destination template - .xlsx only",
        type=["xlsx"],
        help="The MCA/ROC filing template with FieldPLBS.. Field-ID columns. "
             "If your master template is only saved as .xls, open it once in "
             "Excel/LibreOffice and 'Save As' .xlsx - you only need to do this once, "
             "the same template is reused every filing period.",
    )

    st.sidebar.title("2. Multiplication factor")
    multiplier = st.sidebar.selectbox(
        "Multiply every mapped monetary figure by:", MULTIPLIER_OPTIONS, index=0,
        help="Applies to amount fields only. Ratios, EPS, percentages and share "
             "counts are excluded by default (you can override this per row).",
    )
    custom = st.sidebar.text_input("...or type a custom factor", value="")
    if custom.strip():
        try:
            multiplier = float(custom.strip())
        except ValueError:
            st.sidebar.warning("Custom factor must be a number - ignoring it.")

    source_path = None
    dest_path = None
    if use_sample:
        try:
            tmp_dir = tempfile.mkdtemp(prefix="fsm_sample_")
            source_path = os.path.join(tmp_dir, "Source_file.xlsx")
            dest_path = os.path.join(tmp_dir, "destination_file.xlsx")
            shutil.copyfile(SAMPLE_SOURCE, source_path)
            shutil.copyfile(SAMPLE_DEST, dest_path)
        except (PermissionError, OSError) as e:
            st.sidebar.error(
                f"Could not read the bundled sample files ({e}).\n\n"
                "On Windows, this is usually the zip's files being 'Blocked' after "
                "download: right-click the downloaded .zip -> Properties -> tick "
                "'Unblock' -> OK, then re-extract, or simply upload your own files "
                "below instead."
            )
            source_path, dest_path = None, None
    if source_file is not None:
        source_path = save_upload(source_file)
    if dest_file is not None:
        dest_path = save_upload(dest_file)

    return source_path, dest_path, multiplier


# --------------------------------------------------------------------------
# Parsing (cached per file path)
# --------------------------------------------------------------------------

@st.cache_data(show_spinner="Scanning source workbook...")
def _parse_source_cached(path, mtime_key):
    grids = engine.load_workbook_grids(path)
    return engine.parse_source(grids)


@st.cache_data(show_spinner="Scanning destination template...")
def _parse_dest_cached(path, mtime_key):
    grids = engine.load_workbook_grids(path)
    fields = engine.parse_destination(grids)
    formula_cells = set(writer.load_formula_map(path).keys())
    return fields, formula_cells


def load_everything(source_path, dest_path):
    ledger = _parse_source_cached(source_path, os.path.getmtime(source_path))
    dest_fields, formula_cells = _parse_dest_cached(dest_path, os.path.getmtime(dest_path))
    logical_rows = engine.build_logical_rows(dest_fields)
    return ledger, dest_fields, logical_rows, formula_cells


# --------------------------------------------------------------------------
# Mapping UI
# --------------------------------------------------------------------------

def format_candidate(score, item: engine.SourceItem) -> str:
    vals_preview = ", ".join(f"{k}: {v:,.0f}" for k, v in list(item.values.items())[:2])
    tag = f"[{item.breadcrumb}] " if item.breadcrumb else ""
    return f"{score:.0%} match - {tag}{item.label}  ({vals_preview})"


def render_row(row: engine.LogicalRow, ledger, multiplier):
    mapping = st.session_state["mapping"]
    state = mapping.setdefault(row.key, {"choice": "blank", "manual_values": [None] * len(row.slots),
                                          "scale": row.slots[0].scale_default})

    candidates = engine.suggest_matches(row.label, ledger, top_n=8)
    options = ["(leave blank / zero)", "(enter value manually)"] + [
        format_candidate(s, it) for s, it in candidates
    ]

    cols = st.columns([3, 4, 2, 2])
    with cols[0]:
        st.markdown(f"**{row.label}**")
        st.caption(row.section)

    default_idx = 0
    if state["choice"] == "manual":
        default_idx = 1
    elif isinstance(state["choice"], int):
        default_idx = state["choice"] + 2
    elif candidates and candidates[0][0] >= 0.85:
        default_idx = 2  # auto-pick the best suggestion if it looks decent

    with cols[1]:
        pick = st.selectbox("Source match", options, index=min(default_idx, len(options) - 1),
                             key=f"pick_{row.key}", label_visibility="collapsed")
    pick_idx = options.index(pick)

    values_preview = [None] * len(row.slots)
    if pick_idx == 0:
        state["choice"] = "blank"
    elif pick_idx == 1:
        state["choice"] = "manual"
        with cols[2]:
            for i, slot in enumerate(row.slots):
                lbl = "Current" if i == 0 else ("Previous" if i == 1 else f"Slot {i+1}")
                v = st.number_input(lbl, value=float(state["manual_values"][i] or 0),
                                     key=f"manual_{row.key}_{i}", label_visibility="visible")
                state["manual_values"][i] = v
                values_preview[i] = v
    else:
        item = candidates[pick_idx - 2][1]
        state["choice"] = pick_idx - 2
        state["matched_item_sheet_row"] = (item.sheet, item.row0)
        ordered_keys = engine.sorted_value_keys(item.values)
        for i in range(len(row.slots)):
            if i < len(ordered_keys):
                values_preview[i] = item.values[ordered_keys[i]]
        with cols[2]:
            for i, slot in enumerate(row.slots):
                lbl = "Current" if i == 0 else ("Previous" if i == 1 else f"Slot {i+1}")
                st.metric(lbl, f"{values_preview[i]:,.0f}" if values_preview[i] is not None else "-")

    with cols[3]:
        state["scale"] = st.checkbox("Apply multiplier", value=state["scale"], key=f"scale_{row.key}")
        if state["scale"] and multiplier != 1:
            shown = [v * multiplier if v is not None else None for v in values_preview]
            st.caption("x " + str(multiplier) + " -> " + ", ".join(
                f"{v:,.0f}" for v in shown if v is not None))

    return state, values_preview


def compute_final_values(logical_rows, multiplier):
    """Returns cell_values {(sheet,row0,col0): value} and a values_by_sheet_row0
    lookup used by the validator."""
    cell_values = {}
    by_sheet_row0 = {}
    mapping = st.session_state["mapping"]
    for row in logical_rows:
        state = mapping.get(row.key)
        if not state or state["choice"] == "blank":
            continue
        if state["choice"] == "manual":
            raw_values = state["manual_values"]
        else:
            candidates = engine.suggest_matches(row.label, st.session_state["ledger"], top_n=8)
            idx = state["choice"]
            if idx >= len(candidates):
                continue
            item = candidates[idx][1]
            ordered_keys = engine.sorted_value_keys(item.values)
            raw_values = [item.values[k] if i < len(ordered_keys) else None
                          for i, k in enumerate(ordered_keys)]
            raw_values += [None] * (len(row.slots) - len(raw_values))

        for slot, raw in zip(row.slots, raw_values):
            if raw is None:
                continue
            val = raw * multiplier if state["scale"] else raw
            cell_values[(slot.sheet, slot.row0, slot.value_col0)] = val
            by_sheet_row0[(slot.sheet, slot.row0)] = val
    return cell_values, by_sheet_row0


# --------------------------------------------------------------------------
# Validation panel
# --------------------------------------------------------------------------

def render_validation(dest_fields, formula_cells, by_sheet_row0, dest_path):
    st.subheader("Validation")
    formulas = writer.load_formula_map(dest_path)

    problems = []
    checked = 0
    for (sheet, row0, col0), formula in formulas.items():
        recomputed = validator.recompute_total(formula, col0, {r: v for (s, r), v in by_sheet_row0.items() if s == sheet})
        checked += 1
        # No independent "expected" total to diff against unless a total row
        # itself was mapped - report the recomputed figure either way so the
        # user can eyeball it against the source's own stated total.
        if abs(recomputed) > 0:
            problems.append((sheet, row0 + 1, recomputed))

    with st.expander(f"Recomputed subtotal formulas ({checked} found in template)"):
        if problems:
            for sheet, row1, val in sorted(problems):
                st.write(f"{sheet} row {row1}: recomputes to **{val:,.2f}** from mapped components")
        else:
            st.write("No subtotal formulas produced a non-zero figure yet - map more rows to see totals build up.")

    eql_total = by_sheet_row0.get(("BSCurrent", _row0_of(dest_fields, "BSCurrent", "EQL_tot")))
    ca_total = by_sheet_row0.get(("BSCurrent", _row0_of(dest_fields, "BSCurrent", "CA_tot")))
    if eql_total is not None or ca_total is not None:
        eql_formula = formulas.get(("BSCurrent", _row0_of(dest_fields, "BSCurrent", "EQL_tot"), 2))
        ca_formula = formulas.get(("BSCurrent", _row0_of(dest_fields, "BSCurrent", "CA_tot"), 2))
        eql_val = validator.recompute_total(eql_formula, 2, {r: v for (s, r), v in by_sheet_row0.items() if s == "BSCurrent"}) if eql_formula else None
        ca_val = validator.recompute_total(ca_formula, 2, {r: v for (s, r), v in by_sheet_row0.items() if s == "BSCurrent"}) if ca_formula else None
        result = validator.check_balance_sheet_identity(eql_val, ca_val)
        if result.get("checked"):
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Equity & Liabilities", f"{result['equity_and_liabilities']:,.0f}")
            c2.metric("Total Assets", f"{result['total_assets']:,.0f}")
            c3.metric("Difference", f"{result['difference']:,.2f}",
                      delta="Balanced" if result["balanced"] else "OUT OF BALANCE")
            if not result["balanced"]:
                st.error("Balance Sheet does not balance yet - Total Equity & Liabilities should equal Total Assets. "
                          "Check for unmapped or mis-mapped line items.")
            else:
                st.success("Balance Sheet balances.")


def _row0_of(dest_fields, sheet, field_id):
    for f in dest_fields.get(sheet, []):
        if f.field_id == field_id:
            return f.row0
    return -1


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    init_state()
    st.title("Financial Statement -> MCA Filing Template Mapper")
    st.caption(
        "Scans a company's financial statement workbook, auto-suggests matches against "
        "your MCA/ROC filing template's Field-ID cells, lets you review and adjust the "
        "mapping, applies an optional multiplication factor, validates totals, and "
        "exports a filled copy of the template."
    )

    source_path, dest_path, multiplier = sidebar()

    if not source_path or not dest_path:
        st.info("Upload a source file and a destination .xlsx template in the sidebar "
                "(or tick 'Use bundled sample') to get started.")
        return

    try:
        ledger, dest_fields, logical_rows, formula_cells = load_everything(source_path, dest_path)
    except (PermissionError, OSError) as e:
        st.error(f"Could not read one of the workbook files: {e}\n\n"
                 "Check the file isn't open in Excel elsewhere, and that you have "
                 "read access to it, then try again.")
        return
    except Exception as e:
        st.error(f"Could not parse the workbook(s): {e}")
        return
    st.session_state["ledger"] = ledger
    st.session_state["dest_fields"] = dest_fields
    st.session_state["logical_rows"] = logical_rows
    st.session_state["dest_path"] = dest_path
    st.session_state["formula_cells"] = formula_cells

    st.success(f"Parsed {len(ledger)} source line items and {sum(len(v) for v in dest_fields.values())} "
               f"destination fields across {len(dest_fields)} sheets.")

    tabs = st.tabs([g[0] for g in SHEET_GROUPS])
    rows_by_sheet = {}
    for row in logical_rows:
        rows_by_sheet.setdefault(row.slots[0].sheet, []).append(row)

    for tab, (group_name, sheets) in zip(tabs, SHEET_GROUPS):
        with tab:
            group_rows = []
            for sh in sheets:
                group_rows.extend(rows_by_sheet.get(sh, []))
            st.caption(f"{len(group_rows)} line item(s) in this section.")

            sections = {}
            for row in group_rows:
                sections.setdefault(row.section, []).append(row)

            for section, rows in sections.items():
                with st.expander(f"{section}  ({len(rows)} items)", expanded=len(rows) <= 6):
                    for row in rows:
                        render_row(row, ledger, multiplier)
                        st.divider()

    st.markdown("---")
    cell_values, by_sheet_row0 = compute_final_values(logical_rows, multiplier)
    render_validation(dest_fields, formula_cells, by_sheet_row0, dest_path)

    st.markdown("---")
    st.subheader("Export")
    st.write(f"{len(cell_values)} cell(s) are currently mapped and ready to write into the template.")

    out_path = os.path.join(tempfile.gettempdir(), "filled_destination.xlsx")
    try:
        warnings = writer.write_output(dest_path, out_path, cell_values, formula_cells)
        with open(out_path, "rb") as f:
            data = f.read()
    except Exception as e:
        st.error(f"Could not build the export file: {e}")
        return

    struct = writer.structure_summary(dest_path)
    st.caption(
        f"Export preserves the original template's structure: {struct['sheets']} sheet(s) "
        f"({', '.join(struct['sheet_names'])}), all existing formulas, cell formatting, "
        f"and merged cells - only the mapped value cells are changed."
    )

    st.download_button(
        "Download filled_destination.xlsx",
        data=data,
        file_name="filled_destination.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    if warnings:
        with st.expander(f"{len(warnings)} notice(s) about this export"):
            for w in warnings:
                st.write("- " + w)


if __name__ == "__main__":
    main()
