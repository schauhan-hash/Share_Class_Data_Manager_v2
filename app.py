from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
import json

import streamlit as st

from share_class_manager.core import (
    clean_value,
    diff_records,
    fill_missing_isins,
    normalize_records,
    validate_dataset,
    validate_record,
)
from share_class_manager.schema import (
    DATE_FIELDS,
    DEFAULT_COLUMNS,
    FIELD_GROUPS,
    FIELD_ORDER,
    ROLE_PERMISSIONS,
    SELECT_OPTIONS,
)
from share_class_manager.xlsx_io import read_workbook, write_workbook

ROOT = Path(__file__).parent
SOURCE_FILE = ROOT / "data" / "Genericized_Sample_Data.xlsx"
TEMPLATE_FILE = ROOT / "data" / "Corrected_Excel_File_Format.xlsx"

st.set_page_config(
    page_title="Share Class Data Manager",
    page_icon="H",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def inject_style():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400;500;600;700&display=swap');
        :root {
            --navy: #14284b; --navy-2: #1f3b6e; --maroon: #7a1f2b;
            --page: #f3f5f9; --surface: #ffffff; --ink: #1a2233;
            --muted: #55607a; --line: #d5dbe7; --col-head: #e4eaf4;
        }
        html, body, [class*="css"] { font-family: 'Source Sans 3', 'Segoe UI', sans-serif; }
        .stApp { background: var(--page); color: var(--ink); }
        .block-container { max-width: 1680px; padding-top: 0; padding-bottom: 3rem; }
        header[data-testid="stHeader"] { background: transparent; }
        [data-testid="stSidebar"] { display: none; }
        h1, h2, h3 { color: var(--navy); font-weight: 700; }
        .st-key-app_header {
            position: relative; left: 50%; transform: translateX(-50%);
            width: 100vw; max-width: 100vw; box-sizing: border-box;
            margin: 0; padding: 1.55rem max(2rem, calc((100vw - 1680px) / 2 + 2rem)) 1.15rem;
            background: var(--navy); color: #fff;
        }
        .st-key-app_header [data-testid="stHorizontalBlock"] { align-items: center; }
        .app-hero {
            margin: 0; padding: 0; background: transparent; color: #fff;
            display: flex; align-items: center; gap: 18px;
        }
        .brand-mark {
            width: 56px; height: 56px; flex: none; border-radius: 10px;
            background: var(--maroon); display: grid; place-items: center;
            font-size: 1.65rem; font-weight: 700;
            box-shadow: inset 0 0 0 1px rgba(255,255,255,.2);
        }
        .app-hero h1 { color: #fff; margin: 0 0 .2rem; font-size: 1.75rem; }
        .app-hero p { color: #d3dcee; margin: 0; font-size: 1.02rem; }
        .hero-copy { flex: 1 1 620px; min-width: 0; }
        [data-testid="stPopover"] > button {
            color: var(--navy-2); border: 1px solid var(--line);
            background: var(--surface); border-radius: 6px;
            font-weight: 600; min-height: 38px;
        }
        .st-key-workspace_navigation {
            position: relative; left: 50%; transform: translateX(-50%);
            width: 100vw; max-width: 100vw; box-sizing: border-box;
            margin: 0 0 1.1rem; padding: 0 max(2rem, calc((100vw - 1680px) / 2 + 2rem));
            background: var(--navy); border-bottom: 4px solid var(--maroon);
        }
        .st-key-workspace_navigation [data-testid="stHorizontalBlock"] {
            flex-wrap: nowrap; gap: 4px; overflow-x: auto;
        }
        .st-key-workspace_navigation [data-testid="column"] {
            flex-shrink: 0; min-width: max-content;
        }
        .st-key-workspace_navigation div.stButton > button {
            width: 100%; min-height: 46px; padding: 10px 18px;
            border: 0 !important; border-bottom: 3px solid transparent !important;
            border-radius: 0; background: transparent !important;
            color: #d3dcee !important; font-size: 15px; font-weight: 600;
            text-align: left; white-space: nowrap; box-shadow: none !important;
        }
        .st-key-workspace_navigation div.stButton > button:hover {
            background: rgba(255,255,255,.08) !important; color: #fff !important;
        }
        .st-key-workspace_navigation div.stButton > button[kind="primary"],
        .st-key-workspace_navigation div.stButton > [data-testid="stBaseButton-primary"] {
            color: #fff !important; border-bottom-color: #fff !important;
        }
        .st-key-workspace_navigation div.stButton > button:focus-visible {
            outline: 3px solid #9bbcff; outline-offset: -3px;
        }
        .st-key-app_header [data-testid="stPopover"] > button {
            width: 100%; min-width: 260px; min-height: 68px; padding: 10px 18px 10px 12px;
            justify-content: flex-start; text-align: left;
            border: 1px solid rgba(255,255,255,.32); border-radius: 999px;
            background: rgba(255,255,255,.08); color: #fff;
        }
        .st-key-app_header [data-testid="stPopover"] > button p {
            margin: 0; color: #d3dcee; font-size: .8rem; line-height: 1.15;
        }
        .st-key-app_header [data-testid="stPopover"] > button strong {
            color: #fff; font-size: 1.02rem; line-height: 1.25;
        }
        .section-heading {
            background: var(--surface); border: 1px solid var(--line);
            border-left: 5px solid var(--navy-2); border-radius: 8px;
            padding: 15px 18px; margin: .35rem 0 1rem;
        }
        .section-heading h1 { margin: .15rem 0 .25rem; font-size: 1.55rem; color: var(--navy); }
        .section-description { color: var(--muted); max-width: 1000px; }
        .eyebrow {
            color: var(--maroon); font-weight: 700; text-transform: uppercase;
            letter-spacing: .1em; font-size: .76rem;
        }
        .exception-banner, .notice-banner {
            background: #f9eef0; border: 1px solid #e2c2c8; border-radius: 8px;
            padding: 12px 16px; margin-bottom: 1rem;
        }
        .exception-banner strong { color: var(--maroon); }
        .notice-banner { background: #eef3fb; border-color: #b9c9e6; }
        [data-testid="stTabs"] [data-baseweb="tab-list"] {
            border-bottom: 1px solid var(--line); gap: 0;
        }
        [data-testid="stTabs"] [data-baseweb="tab"] {
            min-height: 44px; padding: 10px 16px; font-weight: 600;
        }
        div.stButton > button, div.stDownloadButton > button {
            border-radius: 6px; border: 1px solid var(--maroon);
            background: var(--maroon); color: white; font-weight: 600; min-height: 40px;
        }
        div.stButton > button:hover, div.stDownloadButton > button:hover {
            background: #651823; border-color: #651823;
        }
        div.stFormSubmitButton > button {
            background: var(--maroon); border-color: var(--maroon);
        }
        [data-testid="stDataFrame"], [data-testid="stExpander"] {
            border: 1px solid var(--line); border-radius: 8px;
        }
        @media (max-width: 640px) {
            .block-container { padding-left: 1rem; padding-right: 1rem; }
            .st-key-app_header, .st-key-workspace_navigation {
                padding-left: 1rem; padding-right: 1rem;
            }
            .brand-mark { width: 46px; height: 46px; }
            .app-hero h1 { font-size: 1.35rem; }
            .st-key-app_header [data-testid="stPopover"] > button { min-width: 0; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def initialize_state():
    if "records" not in st.session_state:
        st.session_state.records = normalize_records(read_workbook(SOURCE_FILE))
        st.session_state.pending = []
        st.session_state.history = []
        st.session_state.role = "Viewer"
        st.session_state.data_version = 0


def permissions():
    return ROLE_PERMISSIONS[st.session_state.role]


def actor():
    role = st.session_state.role
    return {
        "Viewer": "demo.viewer@example.com",
        "Editor": "demo.editor@example.com",
        "Reviewer": "demo.reviewer@example.com",
        "Admin": "demo.admin@example.com",
    }[role]


def set_workspace_page(page):
    st.session_state.workspace_nav = page


def now_utc():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def filtered_records():
    records = st.session_state.records
    search = st.session_state.get("search_filter", "").strip().lower()
    if search:
        records = [
            row
            for row in records
            if any(search in str(row.get(field, "")).lower() for field in FIELD_ORDER)
        ]
    for label, field in [
        ("lifecycle_filter", "Share Class Lifecycle"),
        ("currency_filter", "Share Class Currency"),
        ("type_filter", "Share Class Type"),
        ("umbrella_filter", "Umbrella"),
        ("share_class_filter", "Share Class"),
        ("tpa_filter", "TPA Name"),
    ]:
        selected = st.session_state.get(label, "All")
        if selected and selected != "All":
            records = [row for row in records if row.get(field, "") == selected]
    return records


def page_title(eyebrow, title, description):
    st.markdown(
        f'<section class="section-heading"><div class="eyebrow">{eyebrow}</div>'
        f'<h1>{title}</h1>'
        f'<div class="section-description">{description}</div></section>',
        unsafe_allow_html=True,
    )


def compact_table(records, columns=DEFAULT_COLUMNS):
    selected = [column for column in columns if column in FIELD_ORDER]
    st.dataframe(
        [{column: record.get(column, "") for column in selected} for record in records],
        use_container_width=True,
        hide_index=True,
        height=430,
    )


def dashboard_page():
    page_title(
        "Control centre",
        "Share Class Records",
        "Browse the complete working dataset, apply operational filters, and open record changes from the Create & Edit workspace.",
    )
    records = st.session_state.records
    issues = validate_dataset(records)
    errors = [issue for issue in issues if issue["severity"] == "Error"]
    if issues:
        st.markdown(
            f'<div class="exception-banner"><strong>{len(issues)} data-quality '
            f'exception{"s" if len(issues) != 1 else ""}</strong><br>'
            f'{len(errors)} blocking issue{"s" if len(errors) != 1 else ""} require attention in Data Quality.</div>',
            unsafe_allow_html=True,
        )
    controls = st.columns([2, 1, 1, 1, 1, 1])
    controls[0].text_input(
        "Search",
        key="search_filter",
        placeholder="Name, ISIN, CUSIP, fund or TPA ID",
    )
    filter_fields = [
        ("lifecycle_filter", "Lifecycle", "Share Class Lifecycle"),
        ("umbrella_filter", "Umbrella", "Umbrella"),
        ("share_class_filter", "Share Class", "Share Class"),
        ("tpa_filter", "TPA", "TPA Name"),
        ("currency_filter", "Currency", "Share Class Currency"),
    ]
    for control, (key, label, field) in zip(controls[1:], filter_fields):
        values = ["All"] + sorted({row.get(field, "") for row in records} - {""})
        control.selectbox(label, values, key=key)

    visible = filtered_records()
    st.caption(f"Showing {len(visible)} of {len(records)} records")
    with st.expander("Visible field groups", expanded=False):
        column_groups = st.multiselect(
            "Field groups",
            list(FIELD_GROUPS),
            default=list(FIELD_GROUPS),
            label_visibility="collapsed",
        )
    columns_to_show = []
    for group in column_groups:
        columns_to_show.extend(FIELD_GROUPS[group])
    compact_table(visible, columns=columns_to_show or DEFAULT_COLUMNS)

    with st.expander("Coverage and distribution"):
        left, right = st.columns([1.15, 1])
        with left:
            st.subheader("Classes by currency")
            currency_counts = {}
            for row in records:
                value = row.get("Share Class Currency") or "Missing"
                currency_counts[value] = currency_counts.get(value, 0) + 1
            st.bar_chart(currency_counts, color="#1f3b6e")
        with right:
            st.subheader("Least complete fields")
            field_counts = []
            for field in FIELD_ORDER:
                completed = sum(
                    bool(row.get(field)) and str(row[field]).lower() != "n/a"
                    for row in records
                )
                if completed < len(records):
                    field_counts.append(
                        {
                            "Field": field,
                            "Completion": round(100 * completed / max(len(records), 1), 1),
                        }
                    )
            field_counts.sort(key=lambda row: row["Completion"])
            st.dataframe(field_counts[:10], hide_index=True, use_container_width=True)
    st.markdown(
        '<div class="notice-banner"><strong>Session working copy.</strong> Edits and decisions '
        'remain in this browser session. Download the working workbook or audit log before the '
        'application restarts.</div>',
        unsafe_allow_html=True,
    )


def records_page():
    page_title(
        "Master data",
        "Create & Edit Share Classes",
        "Create a record, update an existing share class, or propose a deletion. Every mutation is routed through the review queue.",
    )
    if "edit" not in permissions():
        st.info(f"{st.session_state.role} can view records. Switch to Editor or Admin to propose changes.")
        return

    editor_tab, delete_tab = st.tabs(["Propose add / edit", "Propose deletion"])
    with editor_tab:
        render_editor()
    with delete_tab:
        render_delete_request()


def field_widget(record, field, prefix, key_suffix):
    value = record.get(field, "")
    label = field + (" *" if field in __import__(
        "share_class_manager.schema", fromlist=["REQUIRED_FIELDS"]
    ).REQUIRED_FIELDS else "")
    if field == "ISIN":
        st.text_input(
            label,
            value=value,
            disabled=False,
            help="Leave blank only when proposing a new record. The next generated ISIN is assigned on submit.",
            key=f"{prefix}_{key_suffix}_{field}",
        )
        return st.session_state[f"{prefix}_{key_suffix}_{field}"]
    if field in SELECT_OPTIONS:
        options = SELECT_OPTIONS[field]
        if value and value not in options:
            options = [value] + options
        selected = st.selectbox(label, options, index=options.index(value) if value in options else 0, key=f"{prefix}_{key_suffix}_{field}")
        return selected
    if field == "Comments" or field == "Initial Management Fee Note":
        return st.text_area(label, value=value, height=90, key=f"{prefix}_{key_suffix}_{field}")
    return st.text_input(label, value=value, key=f"{prefix}_{key_suffix}_{field}")


def render_editor():
    record_options = {
        f"{row.get('ISIN') or 'No ISIN'} — {row.get('TPA Share Class ID') or row.get('Share Class', 'Unnamed')}": index
        for index, row in enumerate(st.session_state.records)
    }
    mode = st.radio("Action", ["Add a record", "Edit an existing record"], horizontal=True)
    if mode == "Add a record":
        before = {}
        record = {field: "" for field in FIELD_ORDER}
        prefix = "new"
    else:
        label = st.selectbox("Select record", list(record_options))
        before = st.session_state.records[record_options[label]]
        record = deepcopy(before)
        prefix = f"edit_{record_options[label]}"

    with st.form(f"{prefix}_form"):
        tabs = st.tabs(list(FIELD_GROUPS))
        edited = deepcopy(record)
        for tab, (group, fields) in zip(tabs, FIELD_GROUPS.items()):
            with tab:
                widget_columns = st.columns(2)
                for position, field in enumerate(fields):
                    with widget_columns[position % 2]:
                        edited[field] = field_widget(edited, field, prefix, mode.replace(" ", "_"))
        reason = st.text_area(
            "Reason for change",
            placeholder="Describe the purpose of this proposed change.",
            height=80,
        )
        submitted = st.form_submit_button("Submit for review", use_container_width=True)

    if submitted:
        if not reason.strip():
            st.error("A reason is required for the audit history.")
            return
        if mode == "Add a record":
            edited = fill_missing_isins([edited], st.session_state.records)[0]
        elif not edited["ISIN"]:
            st.error("ISIN cannot be removed from an existing record.")
            return
        candidate = deepcopy(st.session_state.records)
        if before:
            candidate[next(i for i, row in enumerate(candidate) if row["ISIN"] == before["ISIN"])] = edited
        else:
            candidate.append(edited)
        issues = validate_record(edited)
        dataset_issues = validate_dataset(candidate)
        blockers = [issue for issue in issues if issue["severity"] == "Error"]
        duplicate_blockers = [
            issue
            for issue in dataset_issues
            if issue["severity"] == "Error"
            and "Duplicate value" in issue["message"]
            and clean_value(edited.get(issue["field"], "")) == issue["record"]
        ]
        if blockers or duplicate_blockers:
            st.error("This proposal contains blocking issues and was not submitted.")
            st.dataframe(blockers + duplicate_blockers, hide_index=True, use_container_width=True)
            return
        st.session_state.pending.append(
            {
                "id": f"CR-{len(st.session_state.pending) + len(st.session_state.history) + 1:04d}",
                "kind": "record",
                "title": "Add share class" if not before else "Edit share class",
                "before": deepcopy(before),
                "after": edited,
                "reason": reason,
                "actor": actor(),
                "created": now_utc(),
            }
        )
        st.success(f"Proposal submitted. ISIN: {edited['ISIN']}. A Reviewer can approve it from the review queue.")


def render_delete_request():
    labels = {
        f"{row.get('ISIN')} — {row.get('TPA Share Class ID') or row.get('Share Class', 'Unnamed')}": row
        for row in st.session_state.records
    }
    label = st.selectbox("Record to delete", list(labels), key="delete_record")
    reason = st.text_input("Deletion reason", key="delete_reason")
    confirm = st.checkbox("I confirm this is a proposal, not an immediate deletion.")
    if st.button("Submit deletion for review", type="secondary"):
        if not reason or not confirm:
            st.error("Provide a reason and confirm the proposal.")
            return
        record = labels[label]
        st.session_state.pending.append(
            {
                "id": f"CR-{len(st.session_state.pending) + len(st.session_state.history) + 1:04d}",
                "kind": "delete",
                "title": "Delete share class",
                "before": deepcopy(record),
                "after": {},
                "reason": reason,
                "actor": actor(),
                "created": now_utc(),
            }
        )
        st.success("Deletion proposal submitted.")


def quality_page():
    page_title(
        "Control checks",
        "Data quality",
        "Errors block proposed changes. Warnings highlight conditions that a reviewer should confirm.",
    )
    issues = validate_dataset(st.session_state.records)
    errors = [issue for issue in issues if issue["severity"] == "Error"]
    warnings = [issue for issue in issues if issue["severity"] == "Warning"]
    metric_columns = st.columns(3)
    metric_columns[0].metric("Total issues", len(issues))
    metric_columns[1].metric("Blocking errors", len(errors))
    metric_columns[2].metric("Warnings", len(warnings))
    severity = st.multiselect("Severity", ["Error", "Warning"], default=["Error", "Warning"])
    filtered = [issue for issue in issues if issue["severity"] in severity]
    if filtered:
        st.dataframe(filtered, use_container_width=True, hide_index=True)
    else:
        st.success("No data-quality issues match the selected severity.")


def import_export_page():
    page_title(
        "Workbook exchange",
        "Import and export",
        "Preview a workbook, generate missing ISINs, validate the dataset, and submit a controlled import.",
    )
    left, right = st.columns(2)
    with left:
        st.subheader("Download current data")
        st.download_button(
            "Download working dataset",
            data=write_workbook(st.session_state.records),
            file_name="Share_Class_Master_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with right:
        st.subheader("Download field template")
        st.download_button(
            "Download taxonomy template",
            data=TEMPLATE_FILE.read_bytes(),
            file_name="Share_Class_Field_Taxonomy.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()
    if "import" not in permissions():
        st.info(f"Switch to Editor or Admin to import an Excel workbook.")
        return
    uploaded = st.file_uploader("Upload an .xlsx file with the corrected field headers", type=["xlsx"])
    if not uploaded:
        st.caption("The first sheet is read as the import source. All 37 field headers should be present.")
        return
    try:
        imported, headers = read_workbook(uploaded.getvalue(), with_headers=True)
        missing_headers = [field for field in FIELD_ORDER if field not in headers]
        unexpected_headers = [header for header in headers if header not in FIELD_ORDER]
        if missing_headers or unexpected_headers:
            st.error("The worksheet headers do not match the corrected taxonomy.")
            if missing_headers:
                st.write("Missing:", missing_headers)
            if unexpected_headers:
                st.write("Unexpected:", unexpected_headers)
            return
        if not imported:
            st.error("The worksheet contains no data rows.")
            return
        imported = fill_missing_isins(imported, st.session_state.records)
        issues = validate_dataset(imported)
        errors = [issue for issue in issues if issue["severity"] == "Error"]
        st.success(f"Loaded {len(imported)} records. Blank ISINs were allocated in the preview.")
        compact_table(imported)
        with st.expander(f"Validation results ({len(issues)} issues)"):
            if issues:
                st.dataframe(issues, use_container_width=True, hide_index=True)
            else:
                st.write("No validation issues found.")
        reason = st.text_input("Import reason", placeholder="Why is this dataset being submitted?")
        st.warning("Approval replaces the entire working dataset. This is clear and safe for a demo, but production imports should use a merge strategy.")
        if st.button("Submit validated import for review", disabled=bool(errors)):
            if not reason.strip():
                st.error("An import reason is required.")
            else:
                st.session_state.pending.append(
                    {
                        "id": f"CR-{len(st.session_state.pending) + len(st.session_state.history) + 1:04d}",
                        "kind": "import",
                        "title": "Replace workbook",
                        "before": deepcopy(st.session_state.records),
                        "after": imported,
                        "reason": reason,
                        "actor": actor(),
                        "created": now_utc(),
                    }
                )
                st.success("Import proposal submitted for review.")
    except (ValueError, KeyError, OSError) as exc:
        st.error(f"The workbook could not be read: {exc}")


def review_page():
    page_title(
        "Separation of duties",
        "Review queue",
        "Inspect proposed field-level changes, imports, and deletions before updating the working dataset.",
    )
    if "review" not in permissions():
        st.info("Switch to Reviewer or Admin to inspect and decide pending proposals.")
        st.caption(f"{len(st.session_state.pending)} pending demo proposal(s).")
        return
    if not st.session_state.pending:
        st.success("There are no pending proposals.")
        return
    for request in list(st.session_state.pending):
        with st.expander(f"{request['id']} — {request['title']} — {request['actor']}"):
            st.write(request["reason"])
            if request["kind"] == "record":
                changes = diff_records(request["before"], request["after"])
                st.dataframe(changes or [{"field": "", "before": "", "after": ""}], hide_index=True, use_container_width=True)
            elif request["kind"] == "import":
                st.write(
                    f"Replace {len(request['before'])} current records with "
                    f"{len(request['after'])} imported records."
                )
                compact_table(request["after"])
            else:
                compact_table([request["before"]])
            approve_column, reject_column = st.columns(2)
            reviewer_note = st.text_input("Decision note", key=f"note_{request['id']}")
            if approve_column.button("Approve", key=f"approve_{request['id']}", use_container_width=True):
                apply_request(request, "Approved", reviewer_note)
                st.rerun()
            if reject_column.button("Reject", key=f"reject_{request['id']}", use_container_width=True):
                if not reviewer_note.strip():
                    st.error("A decision note is required to reject a proposal.")
                else:
                    record_decision(request, "Rejected", reviewer_note)
                    st.rerun()


def apply_request(request, status, note):
    if request["kind"] == "import":
        st.session_state.records = deepcopy(request["after"])
    elif request["kind"] == "delete":
        target = request["before"]["ISIN"]
        st.session_state.records = [
            row for row in st.session_state.records if row.get("ISIN") != target
        ]
    else:
        incoming = request["after"]
        replaced = False
        for index, row in enumerate(st.session_state.records):
            if row.get("ISIN") == incoming.get("ISIN"):
                st.session_state.records[index] = deepcopy(incoming)
                replaced = True
                break
        if not replaced:
            st.session_state.records.append(deepcopy(incoming))
    st.session_state.data_version += 1
    record_decision(request, status, note)


def record_decision(request, status, note):
    st.session_state.history.append(
        {
            **deepcopy(request),
            "status": status,
            "reviewer": actor(),
            "decision": note,
            "decided": now_utc(),
        }
    )
    st.session_state.pending = [
        item for item in st.session_state.pending if item["id"] != request["id"]
    ]


def history_page():
    page_title(
        "Audit trail",
        "Change history",
        "Approved and rejected proposals, including actors, reasons, and decision notes.",
    )
    if not st.session_state.history:
        st.info("No decisions have been recorded in this session.")
        return
    summary = [
        {
            "Request": item["id"],
            "Action": item["title"],
            "Status": item["status"],
            "Submitted by": item["actor"],
            "Submitted": item["created"],
            "Reviewed by": item["reviewer"],
            "Reason": item["reason"],
            "Decision note": item["decision"],
        }
        for item in reversed(st.session_state.history)
    ]
    st.dataframe(summary, use_container_width=True, hide_index=True)
    st.download_button(
        "Download audit log",
        json.dumps(st.session_state.history, indent=2),
        file_name="share_class_audit_log.json",
        mime="application/json",
    )


def admin_page():
    page_title(
        "Configuration",
        "Demo administration",
        "Review role capabilities and reset the session to the genericized source workbook.",
    )
    if "admin" not in permissions():
        st.warning("Switch to Admin to view configuration actions.")
        return
    st.subheader("Demo role matrix")
    role_rows = [
        {
            "Role": role,
            "View": "Yes",
            "Export": "Yes",
            "Propose changes": "Yes" if "edit" in granted else "No",
            "Import": "Yes" if "import" in granted else "No",
            "Review": "Yes" if "review" in granted else "No",
            "Admin": "Yes" if "admin" in granted else "No",
        }
        for role, granted in ROLE_PERMISSIONS.items()
    ]
    st.dataframe(role_rows, hide_index=True, use_container_width=True)
    st.subheader("Production identity mapping")
    st.code(
        "Corporate identity: Microsoft Entra ID\n"
        "Allowed domain: configured company domain\n"
        "Example identity: employee@example.com\n"
        "Recommended groups: ShareClass-Viewers, Editors, Reviewers, Admins\n"
        "No assigned group -> no access",
        language="text",
    )
    st.warning(
        "The sidebar role selector demonstrates permissions only. It is not authentication and "
        "must not be used with real or sensitive information."
    )
    confirm = st.checkbox("Reset this browser session to the original demo workbook.")
    if st.button("Reset demo data", disabled=not confirm):
        st.session_state.records = normalize_records(read_workbook(SOURCE_FILE))
        st.session_state.pending = []
        st.session_state.history = []
        st.session_state.data_version = 0
        st.success("Demo data reset.")


initialize_state()

if "browsing_role" in st.session_state:
    st.session_state.role = st.session_state.browsing_role

workspace_pages = {
    "Overview": dashboard_page,
    "Create & Edit Share Classes": records_page,
    "Data Quality": quality_page,
    "Import / Export": import_export_page,
    "Review Queue": review_page,
    "Change History": history_page,
    "Admin": admin_page,
}
workspace_options = list(workspace_pages)
if st.session_state.get("workspace_nav") not in workspace_pages:
    st.session_state.workspace_nav = "Overview"

inject_style()

with st.container(key="app_header"):
    hero_column, role_column = st.columns([4, 1.35])
    with hero_column:
        st.markdown(
            """
            <div class="app-hero">
              <div class="brand-mark" aria-hidden="true">◆</div>
              <div class="hero-copy">
                <h1>Product Ops – Share Class Data Management Tool</h1>
                <p>Create, review and manage master share class records using a controlled approval workflow.</p>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with role_column:
        with st.popover(
            f"Browsing as  \n**Product Ops – {st.session_state.role}**",
            icon=":material/edit:",
            use_container_width=True,
        ):
            st.session_state.role = st.selectbox(
                "Browsing as",
                list(ROLE_PERMISSIONS),
                index=list(ROLE_PERMISSIONS).index(st.session_state.role),
                key="browsing_role",
            )
            st.caption(actor())

with st.container(key="workspace_navigation"):
    nav_columns = st.columns([len(label) + 2 for label in workspace_options], gap="small")
    for index, (column, label) in enumerate(zip(nav_columns, workspace_options)):
        with column:
            st.button(
                label,
                key=f"workspace_tab_{index}",
                type="primary" if st.session_state.workspace_nav == label else "secondary",
                use_container_width=True,
                on_click=set_workspace_page,
                args=(label,),
            )

st.caption(
    f"{len(st.session_state.records)} working records · "
    f"{len(st.session_state.pending)} pending proposals"
)

workspace_pages[st.session_state.workspace_nav]()

