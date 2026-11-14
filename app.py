from __future__ import annotations

from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

APP_DIR = Path(__file__).resolve().parent
SOURCE_FILE = APP_DIR.parent / "CRINCRINTR.xlsx"
KEEP_SHEETS = ("Insertion Loss Sheet", "Line Diagram")

METADATA_FIELDS: dict[str, tuple[str, str, str]] = {
    "Workbook ID": ("E7", "text", "Test record"),
    "Test date": ("X7", "date", "Test record"),
    "Route": ("E9", "text", "Test record"),
    "Section": ("E10", "text", "Test record"),
    "Service number": ("E11", "text", "Test record"),
    "Separable portion number": ("AJ9", "text", "Test record"),
    "Work order": ("AJ10", "text", "Test record"),
    "Project number": ("AJ11", "text", "Test record"),
    "A-end ESA": ("B14", "text", "Endpoints and instruments"),
    "A-end sheath ID": ("E14", "text", "Endpoints and instruments"),
    "A-end OTDR": ("I14", "text", "Endpoints and instruments"),
    "A-end light source": ("S14", "text", "Endpoints and instruments"),
    "A-end power meter": ("U14", "text", "Endpoints and instruments"),
    "A-end operator": ("W14", "text", "Endpoints and instruments"),
    "B-end ESA": ("B15", "text", "Endpoints and instruments"),
    "B-end sheath ID": ("E15", "text", "Endpoints and instruments"),
    "B-end OTDR": ("I15", "text", "Endpoints and instruments"),
    "B-end light source": ("S15", "text", "Endpoints and instruments"),
    "B-end power meter": ("U15", "text", "Endpoints and instruments"),
    "B-end operator": ("W15", "text", "Endpoints and instruments"),
    "OTDR length (km)": ("E16", "number", "Counts and limits"),
    "Number of fibres": ("P16", "integer", "Counts and limits"),
    "Total splices": ("U16", "integer", "Counts and limits"),
    "Through connectors": ("X16", "integer", "Counts and limits"),
    "Optitips": ("U17", "integer", "Counts and limits"),
    "Optitaps": ("X17", "integer", "Counts and limits"),
    "Splitters": ("AK16", "integer", "Counts and limits"),
    "ODFs": ("AK17", "integer", "Counts and limits"),
    "Line/system ID": ("AK20", "text", "Counts and limits"),
}

MEASUREMENT_FIELDS: tuple[tuple[str, str, str], ...] = (
    ("A20", "Fibre number at A end", "integer"), ("B20", "Fibre number at B end", "integer"),
    ("C20", "1310 nm reference A→B", "number"), ("D20", "1310 nm measured A→B", "number"),
    ("F20", "1310 nm reference B→A", "number"), ("G20", "1310 nm measured B→A", "number"),
    ("J20", "1310 nm link loss A→B", "number"), ("K20", "1310 nm link loss B→A", "number"),
    ("N20", "1550 nm reference A→B", "number"), ("O20", "1550 nm measured A→B", "number"),
    ("Q20", "1550 nm reference B→A", "number"), ("R20", "1550 nm measured B→A", "number"),
    ("U20", "1550 nm link loss A→B", "number"), ("V20", "1550 nm link loss B→A", "number"),
    ("Y20", "1625 nm reference A→B", "number"), ("Z20", "1625 nm measured A→B", "number"),
    ("AB20", "1625 nm reference B→A", "number"), ("AC20", "1625 nm measured B→A", "number"),
    ("AF20", "1625 nm link loss A→B", "number"), ("AG20", "1625 nm link loss B→A", "number"),
)

EVENT_COLUMNS: tuple[tuple[str, str], ...] = (
    ("Event # (B)", "B"), ("Event ID (C)", "C"), ("Splice # (D)", "D"),
    ("Distance from A km (E)", "E"), ("Distance from B km (F)", "F"),
    ("Total distance km (G)", "G"), ("Address (H)", "H"), ("Section (Q)", "Q"),
    ("Splice ID (U)", "U"), ("Cable batch (S)", "S"), ("Cable type (V)", "V"),
    ("Fibres used (X)", "X"), ("Survey length km (Z)", "Z"),
)

HEADER_LINKS: tuple[tuple[str, str, str], ...] = (
    ("Route", "E9", "B5"), ("Section", "E10", "B6"), ("A-end ESA", "B14", "S5"),
    ("A-end sheath ID", "E14", "V5"), ("B-end ESA", "B15", "S6"),
    ("B-end sheath ID", "E15", "V6"), ("Test date", "X7", "Z5"),
    ("Workbook ID", "E7", "Z6"), ("Line/system ID", "AK20", "N6"),
    ("A-end operator", "W14", "S7"), ("B-end operator", "W15", "V7"),
    ("Project number", "AJ11", "Z7"),
)


@st.cache_data
def load_source_values(path: str, modified_ns: int) -> tuple[dict[str, Any], pd.DataFrame]:
    del modified_ns
    wb = load_workbook(path, data_only=True, read_only=False)
    il, ld = wb["Insertion Loss Sheet"], wb["Line Diagram"]
    values = {cell: il[cell].value for _, (cell, _, _) in METADATA_FIELDS.items()}
    values.update({cell: il[cell].value for cell, _, _ in MEASUREMENT_FIELDS})
    events = []
    for row in range(11, 28):
        events.append({label: ld[f"{col}{row}"].value for label, col in EVENT_COLUMNS})
    wb.close()
    return values, pd.DataFrame(events)


def as_number(value: Any, *, integer: bool = False) -> float | int | None:
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    try:
        return int(float(value)) if integer else float(value)
    except (TypeError, ValueError):
        return None


def current_measurements() -> dict[str, float | int | None]:
    return {
        cell: as_number(st.session_state.get(f"cell_{cell}"), integer=kind == "integer")
        for cell, _, kind in MEASUREMENT_FIELDS
    }


def live_metrics(values: dict[str, Any], metadata: dict[str, Any]) -> dict[str, Any]:
    def read(cell: str) -> float | None:
        value = values.get(cell)
        return float(value) if value is not None else None

    def diff(first: str, second: str) -> float | None:
        a, b = read(first), read(second)
        return abs(a - b) if a is not None and b is not None else None

    def avg(a: float | None, b: float | None) -> float | None:
        return (a + b) / 2 if a is not None and b is not None else None

    directional = {
        "1310 nm A→B": diff("C20", "D20"), "1310 nm B→A": diff("F20", "G20"),
        "1550 nm A→B": diff("N20", "O20"), "1550 nm B→A": diff("Q20", "R20"),
        "1625 nm A→B": diff("Y20", "Z20"), "1625 nm B→A": diff("AB20", "AC20"),
    }
    averages = {
        "1310 nm insertion loss": avg(directional["1310 nm A→B"], directional["1310 nm B→A"]),
        "1550 nm insertion loss": avg(directional["1550 nm A→B"], directional["1550 nm B→A"]),
        "1625 nm insertion loss": avg(directional["1625 nm A→B"], directional["1625 nm B→A"]),
        "1310 nm link loss": avg(read("J20"), read("K20")),
        "1550 nm link loss": avg(read("U20"), read("V20")),
        "1625 nm link loss": avg(read("AF20"), read("AG20")),
    }
    for wavelength in ("1310 nm", "1550 nm", "1625 nm"):
        insertion = averages[f"{wavelength} insertion loss"]
        link = averages[f"{wavelength} link loss"]
        averages[f"{wavelength} IL–LL difference"] = insertion - link if insertion is not None and link is not None else None
    length = as_number(metadata.get("E16")) or 0.0
    splices = as_number(metadata.get("U16"), integer=True) or 0
    connectors = as_number(metadata.get("X16"), integer=True) or 0
    optitips = as_number(metadata.get("U17"), integer=True) or 0
    optitaps = as_number(metadata.get("X17"), integer=True) or 0
    odfs = as_number(metadata.get("AK17"), integer=True) or 0
    splitters = as_number(metadata.get("AK16"), integer=True) or 0
    coeffs = (0.35, 0.21, 0.25)
    il_max, ll_max = {}, {}
    for wavelength, coeff in zip(("1310 nm", "1550 nm", "1625 nm"), coeffs):
        extras = optitips * 0.62 + optitaps * 0.40 + odfs * 0.25 + splitters * 16.7
        il_max[wavelength] = coeff * length + 0.1 * splices + 0.3 * connectors + 0.3 + extras
        ll_max[wavelength] = coeff * length + 0.1 * (splices - 2) + 0.3 * (connectors - 2) + 0.3 + extras
    return {"directional": directional, "averages": averages, "max_il": il_max, "max_ll": ll_max}


def clean_value(value: Any) -> Any:
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, str) and not value.strip():
        return None
    return value


def workbook_bytes(metadata: dict[str, Any], measurements: dict[str, Any], events: pd.DataFrame) -> bytes:
    wb = load_workbook(SOURCE_FILE, data_only=False, read_only=False)
    for name in list(wb.sheetnames):
        if name not in KEEP_SHEETS:
            wb.remove(wb[name])
    wb.active = 0
    il, ld = wb["Insertion Loss Sheet"], wb["Line Diagram"]
    for _, (cell, kind, _) in METADATA_FIELDS.items():
        value = metadata.get(cell)
        if kind == "date" and isinstance(value, date) and not isinstance(value, datetime):
            value = datetime.combine(value, datetime.min.time())
        elif kind in ("number", "integer"):
            value = as_number(value, integer=kind == "integer")
        il[cell] = clean_value(value)
    for cell, _, kind in MEASUREMENT_FIELDS:
        il[cell] = as_number(measurements.get(cell), integer=kind == "integer")
    for row_num, (_, record) in enumerate(events.iterrows(), start=11):
        for label, col in EVENT_COLUMNS:
            value = clean_value(record.get(label))
            if label.startswith("Event #"):
                value = as_number(value, integer=True)
            elif label.startswith("Splice #") and value not in (None, "-"):
                value = as_number(value, integer=True)
            elif label.startswith(("Distance", "Total distance", "Survey length")):
                value = as_number(value)
            ld[f"{col}{row_num}"] = value
    ld["C7"] = sum(
        1 for _, record in events.iterrows()
        if clean_value(record.get("Event ID (C)")) is not None
    )
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    output = BytesIO()
    wb.save(output)
    wb.close()
    return output.getvalue()


st.set_page_config(page_title="Fibre test workbook", page_icon=":material/cable:", layout="wide")
st.title("Fibre test workbook")
st.caption("Edit the insertion-loss record and route events. The export contains only the two connected sheets.")

if not SOURCE_FILE.exists():
    st.error(f"Source workbook not found: {SOURCE_FILE}", icon=":material/error:")
    st.stop()
try:
    source_values, source_events = load_source_values(str(SOURCE_FILE), SOURCE_FILE.stat().st_mtime_ns)
except Exception as exc:
    st.exception(exc)
    st.stop()

st.session_state.setdefault("current_values", source_values)
st.session_state.setdefault("event_values", source_events)

left, right = st.columns([1.1, 1], gap="large")
with left:
    st.subheader("Insertion-loss sheet inputs")
    st.caption("Controls write to the listed cells. Existing formula cells are preserved.")
    tabs = st.tabs(["Test record", "Endpoints and instruments", "Counts and limits"])
    for tab, group in zip(tabs, ("Test record", "Endpoints and instruments", "Counts and limits")):
        with tab:
            fields = [(label, cell, kind) for label, (cell, kind, section) in METADATA_FIELDS.items() if section == group]
            for start in range(0, len(fields), 2):
                cols = st.columns(2)
                for column, (label, cell, kind) in zip(cols, fields[start : start + 2]):
                    old_value = st.session_state["current_values"].get(cell)
                    key = f"meta_{cell}"
                    if key not in st.session_state:
                        st.session_state[key] = old_value
                    if kind == "date":
                        default_date = old_value.date() if isinstance(old_value, datetime) else old_value
                        column.date_input(f"{label} ({cell})", value=default_date, key=key)
                    elif kind == "number":
                        column.number_input(f"{label} ({cell})", value=float(old_value or 0), step=0.1, key=key)
                    elif kind == "integer":
                        column.number_input(f"{label} ({cell})", value=int(old_value or 0), step=1, key=key)
                    else:
                        column.text_input(f"{label} ({cell})", value=str(old_value or ""), key=key)

current_metadata = {
    cell: st.session_state.get(f"meta_{cell}", st.session_state["current_values"].get(cell))
    for _, (cell, _, _) in METADATA_FIELDS.items()
}
with right:
    st.subheader("Workbook scope")
    st.write("**Included sheets**")
    st.write("Insertion Loss Sheet · Line Diagram")
    st.write("**Connection**")
    st.write("Insertion-loss metadata feeds the line-diagram header. Route-event rows stay editable in the line diagram table.")
    st.write("**Source file**")
    st.code(SOURCE_FILE.name, language=None)

st.divider()
st.subheader("Line diagram editor")
st.caption("Edit the 17 existing route-event rows. Linked header values update from the insertion-loss controls above.")
edited_events = st.data_editor(
    st.session_state["event_values"], key="line_events_editor", num_rows="fixed",
    hide_index=True, width="stretch",
    column_config={
        "Address (H)": st.column_config.TextColumn("Address (H)", width="large"),
        "Section (Q)": st.column_config.TextColumn("Section (Q)", width="medium"),
        "Splice ID (U)": st.column_config.TextColumn("Splice ID (U)", width="small"),
    },
)
st.session_state["event_values"] = edited_events.copy()

st.subheader("Line diagram header preview")
header_rows = []
for label, source_cell, target_cell in HEADER_LINKS:
    value = current_metadata.get(source_cell)
    if isinstance(value, (date, datetime)):
        value = value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    header_rows.append({"Field": label, "Insertion sheet cell": source_cell, "Line diagram cell": target_cell, "Live value": value})
st.dataframe(pd.DataFrame(header_rows), hide_index=True, alt="Live values linked from insertion-loss inputs to the line diagram")

st.subheader("Fibre measurements")
st.caption("The source workbook has one populated fibre record. Values below map to row 20 on the insertion-loss sheet.")
current_measurements = {
    cell: st.session_state.get(f"cell_{cell}", st.session_state["current_values"].get(cell))
    for cell, _, _ in MEASUREMENT_FIELDS
}
measurement_frame = pd.DataFrame([
    {"Cell": cell, "Measurement": label, "Value": current_measurements[cell]}
    for cell, label, _ in MEASUREMENT_FIELDS
])
edited_measurements = st.data_editor(
    measurement_frame, key="measurements_editor", num_rows="fixed", hide_index=True,
    width="stretch", disabled=["Cell", "Measurement"],
    column_config={"Value": st.column_config.NumberColumn("Value", format="%.3f")},
)
for _, row in edited_measurements.iterrows():
    st.session_state[f"cell_{row['Cell']}"] = row["Value"]
current_measurements = {
    row["Cell"]: as_number(row["Value"], integer=next(kind for cell, _, kind in MEASUREMENT_FIELDS if cell == row["Cell"]) == "integer")
    for _, row in edited_measurements.iterrows()
}
calculated = live_metrics(current_measurements, current_metadata)
st.subheader("Live loss preview")
st.caption("These values recalculate from the current measurements and limits. The Excel formulas remain in the downloadable workbook.")
metric_items = list(calculated["averages"].items())
for start in range(0, len(metric_items), 3):
    cols = st.columns(3)
    for col, (label, value) in zip(cols, metric_items[start : start + 3]):
        col.metric(label, "—" if value is None else f"{value:.3f} dB", border=True)
max_table = pd.DataFrame([
    {"Wavelength": wavelength, "Insertion loss max (dB)": il, "Link loss max (dB)": calculated["max_ll"][wavelength]}
    for wavelength, il in calculated["max_il"].items()
])
st.dataframe(max_table, hide_index=True, alt="Calculated maximum loss by wavelength")
st.info("The original pass/fail formula contains broken references. This preview does not replace its pass/fail result.", title="Formula note")

if st.toggle("Show workbook cell preview", value=False, key="show_workbook_preview"):
    preview = pd.DataFrame(
        [{"Cell": cell, "Value": value} for cell, value in current_metadata.items()]
        + [{"Cell": cell, "Value": value} for cell, value in current_measurements.items()]
    )
    st.dataframe(preview, hide_index=True, alt="Current hard-coded cells to be written into the workbook")

file_data = workbook_bytes(current_metadata, current_measurements, edited_events)
st.download_button(
    "Download edited workbook", data=file_data, file_name="CRINCRINTR_edited.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary", icon=":material/download:",
)
st.caption("Only Insertion Loss Sheet and Line Diagram are included. Existing formulas are retained and recalculate when the file opens in Excel.")
