from pathlib import Path
import os
import shutil
import subprocess
import threading
import time
import json

from dotenv import load_dotenv
from flask import Flask, abort, jsonify, render_template, request, send_file, session
import pandas as pd
from werkzeug.utils import secure_filename

from db.nhit_file_process import (
    LC_ETC_FILE_TYPE,
    add_header_keyword,
    delete_header_keyword,
    get_lc_etc_header_keyword_records,
    get_lc_etc_header_keyword_strings,
    update_header_keyword,
)
from Header_Mapping.header_mapping import (
    VALID_INVALID_LOOKUP_HEADER_MAPPING,
    VALID_INVALID_LOOKUP_REQUIRED_COLUMNS,
)
from Scripts.Life_cycle_merge import _detect_header_row_index, merge_files_in_folder


REPO_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(REPO_ROOT / ".env")


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024 * 1024  # 5 GB request limit
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-key-change-for-production")

ALLOWED_EXTENSIONS = {".xlsx", ".csv", ".xls"}
FILE_PROCESS_DIR = Path(__file__).resolve().parent / "File_Process"

# Top-level process groups under File_Process/
MERGE_VALID_LOOKUP_PARENT_FOLDER = "Merge_Valid_Lookup"
EXEMPT_QUERY_PARENT_FOLDER = "Exempt_Query"

# Sub-processes under Merge_Valid_Lookup/
LIFE_CYCLE_SUBPROCESS_FOLDER = "Life_Cycle_Merge"
VALID_INVALID_SUBPROCESS_FOLDER = "Valid_Invalid_Lookup"

# Backward-compatible names used in process state and routing checks
LIFE_CYCLE_PROCESS_FOLDER = LIFE_CYCLE_SUBPROCESS_FOLDER
VALID_INVALID_PROCESS_FOLDER = VALID_INVALID_SUBPROCESS_FOLDER

# Display name for the parent process group (distinct from the Valid/Invalid Lookup sub-process)
MERGE_VALID_LOOKUP_GROUP_LABEL = "Merge + Valid lookup"

VIL_HEADER_MAPPING_FILENAME = "header_mapping.json"
MIN_LIFE_CYCLE_MERGE_FILES = 2

FILE_PROCESS_DIR.mkdir(parents=True, exist_ok=True)
(
    FILE_PROCESS_DIR / MERGE_VALID_LOOKUP_PARENT_FOLDER / LIFE_CYCLE_SUBPROCESS_FOLDER
).mkdir(parents=True, exist_ok=True)
(
    FILE_PROCESS_DIR / MERGE_VALID_LOOKUP_PARENT_FOLDER / VALID_INVALID_SUBPROCESS_FOLDER
).mkdir(parents=True, exist_ok=True)
(FILE_PROCESS_DIR / EXEMPT_QUERY_PARENT_FOLDER).mkdir(parents=True, exist_ok=True)

life_cycle_state_lock = threading.Lock()
life_cycle_state = {
    "running": False,
    "process_type": LIFE_CYCLE_PROCESS_FOLDER,
    "process_name": "",
    "started_at": None,
    "finished_at": None,
    "last_status": "",
}

valid_invalid_state_lock = threading.Lock()
valid_invalid_state = {
    "running": False,
    "process_type": VALID_INVALID_PROCESS_FOLDER,
    "process_name": "",
    "started_at": None,
    "finished_at": None,
    "last_status": "",
}


def allowed_file(filename):
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def get_process_input_dir(parent_folder, subprocess_folder, process_name):
    """Resolve File_Process/<parent>/<subprocess>/<process_name>/input."""
    safe_parent = secure_filename(parent_folder.strip())
    safe_subprocess = secure_filename(subprocess_folder.strip())
    safe_process_name = secure_filename(process_name.strip())
    if not safe_parent or not safe_subprocess or not safe_process_name:
        return None, None
    process_dir = FILE_PROCESS_DIR / safe_parent / safe_subprocess / safe_process_name
    return process_dir, process_dir / "input"


def list_uploaded_files(parent_folder, subprocess_folder, process_name):
    _, input_dir = get_process_input_dir(parent_folder, subprocess_folder, process_name)
    if not input_dir or not input_dir.exists():
        return []
    return sorted([p.name for p in input_dir.iterdir() if p.is_file()], key=str.lower)


def get_valid_invalid_paths(process_name):
    process_dir, input_dir = get_process_input_dir(
        MERGE_VALID_LOOKUP_PARENT_FOLDER,
        VALID_INVALID_SUBPROCESS_FOLDER,
        process_name,
    )
    if not process_dir or not input_dir:
        return None, None, None
    rate_dir = input_dir / "rate"
    return process_dir, input_dir, rate_dir


def get_valid_invalid_stage_dir(process_name):
    process_dir, _, _ = get_valid_invalid_paths(process_name)
    if not process_dir:
        return None
    return process_dir / ".stage"


def list_file_process_directories():
    return sorted([p.name for p in FILE_PROCESS_DIR.iterdir() if p.is_dir()], key=str.lower)


def normalize_lookup_column_name(value):
    return " ".join(str(value).strip().lower().split())


def inspect_valid_invalid_headers(file_path):
    file_path = Path(file_path)
    header_candidates = []
    for aliases in VALID_INVALID_LOOKUP_HEADER_MAPPING.values():
        header_candidates.extend(aliases)

    if file_path.suffix.lower() == ".csv":
        sample = pd.read_csv(file_path, header=None, dtype=str, nrows=VALID_INVALID_HEADER_SCAN_ROWS)
        header_row_index = _detect_header_row_index(sample, header_candidates, min_matches=2)
        dataframe = pd.read_csv(file_path, skiprows=header_row_index, nrows=0)
        source_name = file_path.name
    else:
        excel_file = pd.ExcelFile(file_path)
        dataframe = None
        header_row_index = 0
        source_name = file_path.name
        for sheet_name in excel_file.sheet_names:
            sample = pd.read_excel(
                excel_file,
                sheet_name=sheet_name,
                header=None,
                dtype=str,
                nrows=VALID_INVALID_HEADER_SCAN_ROWS,
            )
            if sample.empty:
                continue
            header_row_index = _detect_header_row_index(sample, header_candidates, min_matches=2)
            dataframe = pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=header_row_index, nrows=0)
            source_name = f"{file_path.name} [{sheet_name}]"
            if len(dataframe.columns) > 0:
                break
        if dataframe is None:
            dataframe = pd.DataFrame()

    available_columns = [str(col).strip() for col in dataframe.columns if str(col).strip()]
    normalized_columns = {
        normalize_lookup_column_name(column): column for column in available_columns
    }

    detected_mapping = {}
    missing_columns = []
    for canonical, aliases in VALID_INVALID_LOOKUP_HEADER_MAPPING.items():
        detected_column = None
        for candidate in aliases:
            normalized_candidate = normalize_lookup_column_name(candidate)
            if normalized_candidate in normalized_columns:
                detected_column = normalized_columns[normalized_candidate]
                break
        if detected_column:
            detected_mapping[canonical] = detected_column
        else:
            missing_columns.append(canonical)

    return {
        "source_name": source_name,
        "header_row_index": header_row_index,
        "available_columns": available_columns,
        "detected_mapping": detected_mapping,
        "missing_columns": missing_columns,
    }


def list_life_cycle_merge_output_files():
    """Files under File_Process/Merge_Valid_Lookup/Life_Cycle_Merge/<process>/output/."""
    root = FILE_PROCESS_DIR / MERGE_VALID_LOOKUP_PARENT_FOLDER / LIFE_CYCLE_SUBPROCESS_FOLDER
    if not root.exists():
        return []
    results = []
    for proc_dir in sorted(root.iterdir(), key=lambda p: p.name.lower()):
        if not proc_dir.is_dir():
            continue
        out_dir = proc_dir / "output"
        if not out_dir.is_dir():
            continue
        for f in sorted(out_dir.iterdir(), key=lambda p: p.name.lower()):
            if f.is_file() and allowed_file(f.name):
                rel = f.relative_to(FILE_PROCESS_DIR.resolve()).as_posix()
                results.append(
                    {
                        "relative_path": rel,
                        "label": f"{proc_dir.name} / output / {f.name}",
                        "process_name": proc_dir.name,
                        "filename": f.name,
                    }
                )
    return results


def resolve_lcm_output_import_file(relative_path):
    """
    Allow only regular files under Life_Cycle_Merge/<any>/output/<file>.
    """
    target = safe_path_from_relative(relative_path)
    if target is None or not target.is_file():
        return None
    try:
        rel_parts = target.relative_to(FILE_PROCESS_DIR.resolve()).parts
    except ValueError:
        return None
    if len(rel_parts) < 5:
        return None
    if (
        rel_parts[0] != MERGE_VALID_LOOKUP_PARENT_FOLDER
        or rel_parts[1] != LIFE_CYCLE_SUBPROCESS_FOLDER
        or rel_parts[3] != "output"
    ):
        return None
    if not allowed_file(target.name):
        return None
    return target


def _vil_pending_lcm_read():
    return session.get("vil_pending_lcm_import") or {}


def _vil_pending_lcm_write():
    m = session.get("vil_pending_lcm_import")
    if m is None:
        m = {}
        session["vil_pending_lcm_import"] = m
    return m


def get_vil_merged_preview_path(process_name):
    """
    Path to inspect for header mapping: staged uploaded file, pending LCM import, or first file in VIL input.
    Returns (path_or_none, source_label) where source_label is 'upload', 'import', or 'uploaded'.
    """
    pending_path, pending_source = get_vil_pending_merged_preview_path(process_name)
    if pending_path:
        return pending_path, pending_source

    _, input_dir, _ = get_valid_invalid_paths(process_name)
    if input_dir and input_dir.exists():
        merged_files = [p for p in input_dir.iterdir() if p.is_file()]
        if merged_files:
            return merged_files[0], "uploaded"
    return None, None


def get_vil_pending_merged_preview_path(process_name):
    """Staged upload or pending LCM import only — not yet confirmed into input/."""
    stage_dir = get_valid_invalid_stage_dir(process_name)
    if stage_dir and stage_dir.exists():
        staged_files = [p for p in stage_dir.iterdir() if p.is_file()]
        if staged_files:
            return staged_files[0], "upload"
    pending_rel = (_vil_pending_lcm_read().get(process_name) or "").strip()
    if pending_rel:
        src = resolve_lcm_output_import_file(pending_rel)
        if src:
            return src, "import"
    return None, None


def get_valid_invalid_header_info_for_ui(process_name):
    path, _src = get_vil_merged_preview_path(process_name)
    if not path:
        return None
    return inspect_valid_invalid_headers(path)


def get_vil_pending_header_info_for_ui(process_name):
    path, _src = get_vil_pending_merged_preview_path(process_name)
    if not path:
        return None
    return inspect_valid_invalid_headers(path)


def _vil_header_mapping_read():
    return session.get("vil_header_mapping") or {}


def get_vil_header_mapping_path(process_name):
    process_dir, _, _ = get_valid_invalid_paths(process_name)
    if not process_dir:
        return None
    return process_dir / VIL_HEADER_MAPPING_FILENAME


def load_vil_header_mapping_from_disk(process_name):
    mapping_path = get_vil_header_mapping_path(process_name)
    if not mapping_path or not mapping_path.is_file():
        return {}
    try:
        data = json.loads(mapping_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(data, dict):
        return {}
    return {
        str(key): str(value)
        for key, value in data.items()
        if isinstance(key, str) and isinstance(value, str) and value
    }


def is_vil_header_mapping_complete(mapping):
    return all(canonical in mapping for canonical in VALID_INVALID_LOOKUP_REQUIRED_COLUMNS)


def vil_header_mapping_is_confirmed(process_name):
    return is_vil_header_mapping_complete(load_vil_header_mapping_from_disk(process_name))


def save_vil_header_mapping_to_disk(process_name, mapping):
    mapping_path = get_vil_header_mapping_path(process_name)
    if not mapping_path:
        return False
    mapping_path.parent.mkdir(parents=True, exist_ok=True)
    mapping_path.write_text(json.dumps(dict(mapping), indent=2), encoding="utf-8")
    return True


def clear_vil_header_mapping_file(process_name):
    mapping_path = get_vil_header_mapping_path(process_name)
    if mapping_path and mapping_path.is_file():
        mapping_path.unlink()


def get_vil_header_mapping(process_name):
    disk_mapping = load_vil_header_mapping_from_disk(process_name)
    if disk_mapping:
        return disk_mapping
    stored = _vil_header_mapping_read().get(process_name) or {}
    return dict(stored)


def save_vil_header_mapping(process_name, mapping):
    save_vil_header_mapping_to_disk(process_name, mapping)
    mapping_store = dict(_vil_header_mapping_read())
    mapping_store[process_name] = dict(mapping)
    session["vil_header_mapping"] = mapping_store
    session.modified = True


def clear_vil_header_mapping(process_name):
    clear_vil_header_mapping_file(process_name)
    mapping_store = session.get("vil_header_mapping")
    if not mapping_store or process_name not in mapping_store:
        return
    mapping_store = dict(mapping_store)
    mapping_store.pop(process_name, None)
    if mapping_store:
        session["vil_header_mapping"] = mapping_store
    else:
        session.pop("vil_header_mapping", None)
    session.modified = True


def get_vil_reconfirm_header_preview(process_name):
    """Confirmed merged file in input/ still needs mapping saved on disk."""
    if vil_header_mapping_is_confirmed(process_name):
        return None, None
    if get_vil_pending_merged_preview_path(process_name)[0]:
        return None, None
    _, input_dir, _ = get_valid_invalid_paths(process_name)
    if not input_dir or not input_dir.exists():
        return None, None
    merged_files = [p for p in input_dir.iterdir() if p.is_file()]
    if not merged_files:
        return None, None
    return merged_files[0], "reconfirm"


def get_vil_reconfirm_header_info_for_ui(process_name):
    path, _source = get_vil_reconfirm_header_preview(process_name)
    if not path:
        return None
    return inspect_valid_invalid_headers(path)


def clear_vil_pending_lcm_for_process(process_name):
    m = session.get("vil_pending_lcm_import")
    if not m or process_name not in m:
        return
    m = dict(m)
    m.pop(process_name, None)
    if m:
        session["vil_pending_lcm_import"] = m
    else:
        session.pop("vil_pending_lcm_import", None)
    session.modified = True


def clear_vil_stage_for_process(process_name):
    stage_dir = get_valid_invalid_stage_dir(process_name)
    if stage_dir and stage_dir.exists():
        shutil.rmtree(stage_dir, ignore_errors=True)


def clear_vil_merge_selection_for_process(process_name):
    clear_vil_pending_lcm_for_process(process_name)
    clear_vil_stage_for_process(process_name)
    clear_vil_header_mapping(process_name)


def clear_vil_confirmed_merged_files(process_name):
    _, input_dir, _ = get_valid_invalid_paths(process_name)
    if not input_dir or not input_dir.exists():
        return
    for existing in list(input_dir.iterdir()):
        if existing.is_file():
            existing.unlink()


def finalize_vil_staged_merged_file(process_name):
    process_dir, input_dir, _ = get_valid_invalid_paths(process_name)
    if not process_dir or not input_dir:
        return False, "Invalid process name."

    input_dir.mkdir(parents=True, exist_ok=True)
    clear_vil_confirmed_merged_files(process_name)

    stage_dir = get_valid_invalid_stage_dir(process_name)
    if stage_dir and stage_dir.exists():
        staged_files = [p for p in stage_dir.iterdir() if p.is_file()]
        if staged_files:
            staged_file = staged_files[0]
            destination = input_dir / secure_filename(staged_file.name)
            shutil.copy2(staged_file, destination)
            clear_vil_stage_for_process(process_name)
            clear_vil_pending_lcm_for_process(process_name)
            return True, destination.name

    pending_rel = (_vil_pending_lcm_read().get(process_name) or "").strip()
    if pending_rel:
        src = resolve_lcm_output_import_file(pending_rel)
        if not src:
            clear_vil_pending_lcm_for_process(process_name)
            return False, "Selected Life Cycle Merge output is no longer valid."
        destination = input_dir / secure_filename(src.name)
        shutil.copy2(src, destination)
        clear_vil_pending_lcm_for_process(process_name)
        clear_vil_stage_for_process(process_name)
        return True, destination.name

    return False, "No staged merged file is available. Upload and map a merged file first."


def safe_path_from_relative(relative_path):
    normalized = (relative_path or "").strip().replace("\\", "/").strip("/")
    target_path = (FILE_PROCESS_DIR / normalized).resolve() if normalized else FILE_PROCESS_DIR.resolve()
    file_process_root = FILE_PROCESS_DIR.resolve()
    if target_path != file_process_root and file_process_root not in target_path.parents:
        return None
    return target_path


def list_directory_entries(relative_path):
    current_path = safe_path_from_relative(relative_path)
    if current_path is None or not current_path.exists() or not current_path.is_dir():
        return None, []

    entries = []
    for item in sorted(current_path.iterdir(), key=lambda p: (not p.is_dir(), p.name.lower())):
        rel = item.relative_to(FILE_PROCESS_DIR.resolve()).as_posix()
        folder_count = None
        file_count = None
        process_type = None
        process_name = None
        parent_folder = None
        output_file_count = 0
        if item.is_dir():
            direct_children = list(item.iterdir())
            folder_count = sum(1 for child in direct_children if child.is_dir())
            file_count = sum(1 for child in direct_children if child.is_file())
            if item.name == "input":
                try:
                    process_name = item.parent.name
                    subprocess_folder = item.parent.parent.name
                    parent_folder = item.parent.parent.parent.name
                    process_type = subprocess_folder
                    output_dir = item.parent / "output"
                    if output_dir.exists() and output_dir.is_dir():
                        output_file_count = sum(
                            1 for child in output_dir.iterdir() if child.is_file() or child.is_dir()
                        )
                except Exception:
                    process_type = None
                    process_name = None
                    parent_folder = None
                    output_file_count = 0
        entries.append(
            {
                "name": item.name,
                "relative_path": rel,
                "is_dir": item.is_dir(),
                "folder_count": folder_count,
                "file_count": file_count,
                "process_type": process_type,
                "process_name": process_name,
                "parent_folder": parent_folder if item.is_dir() and item.name == "input" else None,
                "output_file_count": output_file_count,
            }
        )
    return current_path, entries


def annotate_process_button_state(
    entries,
    life_cycle_running,
    life_cycle_process_name,
    valid_invalid_running,
    valid_invalid_process_name,
):
    annotated_entries = []
    for entry in entries:
        process_disabled = False
        process_disabled_message = ""

        if entry.get("is_dir") and entry.get("name") == "input":
            process_type = entry.get("process_type") or ""
            process_name = entry.get("process_name") or ""
            parent_folder = entry.get("parent_folder") or ""

            if parent_folder and parent_folder != MERGE_VALID_LOOKUP_PARENT_FOLDER:
                process_disabled = True
                process_disabled_message = (
                    "Processing from the directory browser is only available for "
                    f"'{MERGE_VALID_LOOKUP_GROUP_LABEL}' sub-processes."
                )
            elif (
                life_cycle_running
                and process_type == LIFE_CYCLE_PROCESS_FOLDER
                and process_name == life_cycle_process_name
            ):
                process_disabled = True
                process_disabled_message = (
                    f"Process is disabled while Life Cycle Merge is already running for '{process_name}'."
                )
            elif (
                valid_invalid_running
                and process_type == VALID_INVALID_PROCESS_FOLDER
                and process_name == valid_invalid_process_name
            ):
                process_disabled = True
                process_disabled_message = (
                    "Process is disabled while Valid/Invalid Lookup is already running "
                    f"for '{process_name}'."
                )

        annotated_entry = dict(entry)
        annotated_entry["process_disabled"] = process_disabled
        annotated_entry["process_disabled_message"] = process_disabled_message
        annotated_entries.append(annotated_entry)

    return annotated_entries


@app.route("/file-process-download")
def file_process_download():
    """Download a file from under File_Process (path must stay within root)."""
    relative = request.args.get("path", "").strip()
    target = safe_path_from_relative(relative)
    if target is None or not target.exists() or not target.is_file():
        abort(404)
    return send_file(
        target,
        as_attachment=True,
        download_name=secure_filename(target.name) or target.name,
        max_age=0,
    )


def process_life_cycle_files(process_name):
    process_dir, input_dir = get_process_input_dir(
        MERGE_VALID_LOOKUP_PARENT_FOLDER,
        LIFE_CYCLE_SUBPROCESS_FOLDER,
        process_name,
    )
    if not process_dir or not input_dir:
        return False, "Please provide a valid process name."

    if not input_dir.exists():
        return False, "Input folder does not exist. Upload files first."

    files_to_process = [p for p in input_dir.iterdir() if p.is_file()]
    if not files_to_process:
        return False, "Please upload at least one file before final submit."

    try:
        header_keywords = get_lc_etc_header_keyword_strings()
    except Exception as exc:
        return False, f"Could not load LC/ETC header keywords from database: {exc}"

    if len(header_keywords) < 3:
        return (
            False,
            "Configure at least 3 header keywords for LC/ETC in Header Keywords before running the merge.",
        )

    output_dir = process_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / "merged_output.csv"

    try:
        merge_files_in_folder(
            str(input_dir.resolve()),
            str(output_file.resolve()),
            header_keywords,
        )
    except Exception as exc:
        return False, f"Life cycle merge failed: {exc}"

    if not output_file.exists():
        return (
            False,
            "Merge did not create an output file. Ensure inputs are .csv/.xls/.xlsx and "
            "contain headers matching the configured keywords.",
        )

    return (
        True,
        f"Life cycle merge finished for '{process_dir.name}'. "
        f"Output: output/{output_file.name}",
    )


def clear_output_directory(process_dir):
    output_dir = process_dir / "output"
    if not output_dir.exists() or not output_dir.is_dir():
        return 0

    deleted_count = 0
    for child in output_dir.iterdir():
        if child.is_file():
            child.unlink()
            deleted_count += 1
        elif child.is_dir():
            shutil.rmtree(child)
            deleted_count += 1
    return deleted_count


def run_life_cycle_merge_in_background(process_name):
    """Run merge in background and update shared state."""
    with life_cycle_state_lock:
        life_cycle_state["running"] = True
        life_cycle_state["process_type"] = LIFE_CYCLE_PROCESS_FOLDER
        life_cycle_state["process_name"] = process_name
        life_cycle_state["started_at"] = time.time()
        life_cycle_state["finished_at"] = None
        life_cycle_state["last_status"] = "Merge started."

    ok, process_message = process_life_cycle_files(process_name)

    with life_cycle_state_lock:
        life_cycle_state["running"] = False
        life_cycle_state["finished_at"] = time.time()
        life_cycle_state["last_status"] = process_message if ok else f"FAILED: {process_message}"


@app.route("/life-cycle-merge-status")
def life_cycle_merge_status():
    with life_cycle_state_lock:
        state = dict(life_cycle_state)
    if state["started_at"]:
        state["elapsed_seconds"] = int(time.time() - state["started_at"])
    else:
        state["elapsed_seconds"] = 0
    return jsonify(state)


@app.route("/valid-invalid-status")
def valid_invalid_status():
    with valid_invalid_state_lock:
        state = dict(valid_invalid_state)
    if state["started_at"]:
        state["elapsed_seconds"] = int(time.time() - state["started_at"])
    else:
        state["elapsed_seconds"] = 0
    return jsonify(state)


@app.context_processor
def inject_process_labels():
    return {
        "merge_valid_lookup_group_label": MERGE_VALID_LOOKUP_GROUP_LABEL,
        "merge_valid_lookup_parent_folder": MERGE_VALID_LOOKUP_PARENT_FOLDER,
    }


@app.route("/")
def home():
    return render_template("home.html")


@app.route("/merge-valid-lookup")
def merge_valid_lookup_hub():
    return render_template("merge_valid_lookup_hub.html")


@app.route("/exempt-query")
def exempt_query_hub():
    return render_template(
        "exempt_query_hub.html",
        exempt_query_folder=EXEMPT_QUERY_PARENT_FOLDER,
    )


@app.route("/current-process")
def current_process():
    with life_cycle_state_lock:
        current_life_cycle_state = dict(life_cycle_state)
    with valid_invalid_state_lock:
        current_valid_invalid_state = dict(valid_invalid_state)
    return render_template(
        "current_process.html",
        process_directories=list_file_process_directories(),
        life_cycle_state=current_life_cycle_state,
        valid_invalid_state=current_valid_invalid_state,
    )


def process_valid_invalid_files(process_name, header_mapping=None):
    process_dir, input_dir, rate_dir = get_valid_invalid_paths(process_name)
    if not process_dir or not input_dir or not rate_dir:
        return False, "Please provide a valid process name."
    # Ensure folder structure exists, then validate required files explicitly.
    input_dir.mkdir(parents=True, exist_ok=True)
    rate_dir.mkdir(parents=True, exist_ok=True)
    merged_files = [p for p in input_dir.iterdir() if p.is_file()]
    rate_files = [p for p in rate_dir.iterdir() if p.is_file()] if rate_dir.exists() else []
    if not merged_files:
        return False, "Please upload a merged life cycle file before final submit."
    if not rate_files:
        return False, "Please upload a rates file before final submit."
    output_dir = process_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    invalid_output = output_dir / "invalid_table.csv"
    valid_output = output_dir / "valid_table.csv"

    script_dir = Path(__file__).resolve().parent / "Scripts" / "valid_invalid_lookup"
    script_path = script_dir / "Valid_Invalid_Full_Process.py"
    env = os.environ.copy()
    env["VALID_INVALID_LIFECYCLE_PATH"] = str(merged_files[0].resolve())
    env["VALID_INVALID_RATES_PATH"] = str(rate_files[0].resolve())
    env["VALID_INVALID_OUTPUT_DIR"] = str(output_dir.resolve())
    env["VALID_INVALID_HEADER_MAPPING"] = json.dumps(header_mapping or {})

    try:
        result = subprocess.run(
            ["python", str(script_path)],
            cwd=str(script_dir),
            env=env,
            check=False,
        )
    except Exception as exc:
        return False, f"Valid/Invalid Lookup failed to start: {exc}"

    if result.returncode != 0:
        return False, f"Valid/Invalid Lookup failed with exit code {result.returncode}"

    if not invalid_output.exists() or not valid_output.exists():
        return False, "Valid/Invalid Lookup did not create expected output files."

    return (
        True,
        f"Valid/Invalid Lookup finished for '{process_dir.name}'. "
        "Output: output/invalid_table.csv, output/valid_table.csv",
    )


def run_valid_invalid_in_background(process_name, header_mapping=None):
    with valid_invalid_state_lock:
        valid_invalid_state["running"] = True
        valid_invalid_state["process_type"] = VALID_INVALID_PROCESS_FOLDER
        valid_invalid_state["process_name"] = process_name
        valid_invalid_state["started_at"] = time.time()
        valid_invalid_state["finished_at"] = None
        valid_invalid_state["last_status"] = "Valid/Invalid Lookup started."

    ok, process_message = process_valid_invalid_files(process_name, header_mapping=header_mapping)

    with valid_invalid_state_lock:
        valid_invalid_state["running"] = False
        valid_invalid_state["finished_at"] = time.time()
        valid_invalid_state["last_status"] = process_message if ok else f"FAILED: {process_message}"


@app.route("/file-process-directories", methods=["GET", "POST"])
def file_process_directories():
    messages = []
    relative_path = request.values.get("path", "").strip()

    with life_cycle_state_lock:
        life_cycle_running = life_cycle_state["running"]
        life_cycle_process_name = life_cycle_state["process_name"]
    with valid_invalid_state_lock:
        valid_invalid_running = valid_invalid_state["running"]
        valid_invalid_process_name = valid_invalid_state["process_name"]

    delete_disabled = life_cycle_running or valid_invalid_running
    delete_disabled_message = ""
    if life_cycle_running:
        delete_disabled_message = (
            f"Delete is disabled while Life Cycle Merge is running for '{life_cycle_process_name}'."
        )
    elif valid_invalid_running:
        delete_disabled_message = (
            "Delete is disabled while Valid/Invalid Lookup is running"
            f" for '{valid_invalid_process_name}'."
        )

    if request.method == "POST":
        action = request.form.get("action")
        if action == "delete":
            if delete_disabled:
                messages.append(("error", delete_disabled_message))
                current_path, entries = list_directory_entries(relative_path)
                if current_path is None:
                    current_path, entries = list_directory_entries("")
                    messages.append(("error", "Invalid directory path selected."))
                current_relative = (
                    current_path.relative_to(FILE_PROCESS_DIR.resolve()).as_posix()
                    if current_path != FILE_PROCESS_DIR.resolve()
                    else ""
                )
                parent_relative = ""
                if current_path != FILE_PROCESS_DIR.resolve():
                    parent = current_path.parent
                    parent_relative = (
                        parent.relative_to(FILE_PROCESS_DIR.resolve()).as_posix()
                        if parent != FILE_PROCESS_DIR.resolve()
                        else ""
                    )
                return render_template(
                    "file_process_directories.html",
                    messages=messages,
                    current_relative=current_relative,
                    parent_relative=parent_relative,
                    entries=annotate_process_button_state(
                        entries,
                        life_cycle_running,
                        life_cycle_process_name,
                        valid_invalid_running,
                        valid_invalid_process_name,
                    ),
                    delete_disabled=delete_disabled,
                    delete_disabled_message=delete_disabled_message,
                )

            target_relative = request.form.get("target_path", "").strip()
            target_path = safe_path_from_relative(target_relative)
            if target_path is None or not target_path.exists():
                messages.append(("error", "Selected file/folder does not exist."))
            elif target_path == FILE_PROCESS_DIR.resolve():
                messages.append(("error", "Cannot delete File_Process root folder."))
            else:
                with life_cycle_state_lock:
                    running = life_cycle_state["running"]
                    running_process_type = life_cycle_state["process_type"]
                    running_process_name = life_cycle_state["process_name"]

                running_process_dir = (
                    (
                        FILE_PROCESS_DIR
                        / MERGE_VALID_LOOKUP_PARENT_FOLDER
                        / running_process_type
                        / running_process_name
                    ).resolve()
                    if running and running_process_name
                    else None
                )
                if (
                    running_process_dir
                    and (target_path == running_process_dir or running_process_dir in target_path.parents)
                ):
                    messages.append(
                        (
                            "error",
                            f"Cannot delete '{running_process_name}' while its merge process is running.",
                        )
                    )
                else:
                    try:
                        if target_path.is_dir():
                            shutil.rmtree(target_path)
                            messages.append(("success", f"Deleted folder: {target_path.name}"))
                        else:
                            target_path.unlink()
                            messages.append(("success", f"Deleted file: {target_path.name}"))
                    except PermissionError:
                        messages.append(
                            (
                                "error",
                                "Cannot delete because file/folder is currently in use by another process. "
                                "Please close any open file handles and try again.",
                            )
                        )
                    except OSError as exc:
                        messages.append(("error", f"Delete failed: {exc}"))
        elif action == "run_process":
            target_relative = request.form.get("target_path", "").strip()
            target_path = safe_path_from_relative(target_relative)
            overwrite_output = request.form.get("overwrite_output") == "yes"

            if target_path is None or not target_path.exists() or not target_path.is_dir():
                messages.append(("error", "Selected input folder does not exist."))
            elif target_path.name != "input":
                messages.append(("error", "Process can only be started from an input folder."))
            else:
                try:
                    process_name = target_path.parent.name
                    subprocess_folder = target_path.parent.parent.name
                    parent_folder = target_path.parent.parent.parent.name
                    process_type = subprocess_folder
                except Exception:
                    process_name = ""
                    subprocess_folder = ""
                    parent_folder = ""
                    process_type = ""

                process_dir = target_path.parent
                output_dir = process_dir / "output"
                existing_output_items = []
                if output_dir.exists() and output_dir.is_dir():
                    existing_output_items = [p for p in output_dir.iterdir() if p.is_file() or p.is_dir()]

                if existing_output_items and not overwrite_output:
                    messages.append(
                        (
                            "error",
                            f"Output already exists for '{process_name}'. Confirm overwrite to delete old output files and start processing again.",
                        )
                    )
                elif existing_output_items:
                    try:
                        deleted_count = clear_output_directory(process_dir)
                        messages.append(
                            (
                                "success",
                                f"Deleted {deleted_count} existing output item(s) for '{process_name}'.",
                            )
                        )
                    except PermissionError:
                        messages.append(
                            (
                                "error",
                                "Cannot clear the output folder because one or more files are currently in use.",
                            )
                        )
                    except OSError as exc:
                        messages.append(("error", f"Failed to clear output folder: {exc}"))

                if messages and messages[-1][0] == "error" and "Output already exists" in messages[-1][1]:
                    pass
                elif messages and messages[-1][0] == "error" and (
                    "Cannot clear the output folder" in messages[-1][1]
                    or "Failed to clear output folder" in messages[-1][1]
                ):
                    pass
                elif parent_folder == EXEMPT_QUERY_PARENT_FOLDER:
                    messages.append(
                        (
                            "error",
                            "Exempt Query processing is not yet available from the directory browser.",
                        )
                    )
                elif parent_folder != MERGE_VALID_LOOKUP_PARENT_FOLDER:
                    messages.append(("error", "Unsupported process group for this input folder."))
                elif process_type == LIFE_CYCLE_PROCESS_FOLDER:
                    with life_cycle_state_lock:
                        running = life_cycle_state["running"]

                    if running:
                        messages.append(
                            (
                                "error",
                                "A life cycle merge is already running. Please wait for it to finish.",
                            )
                        )
                    else:
                        worker = threading.Thread(
                            target=run_life_cycle_merge_in_background,
                            args=(process_name,),
                            daemon=True,
                        )
                        worker.start()
                        messages.append(
                            (
                                "success",
                                f"Life cycle merge started in background for '{process_name}'.",
                            )
                        )
                elif process_type == VALID_INVALID_PROCESS_FOLDER:
                    with valid_invalid_state_lock:
                        running = valid_invalid_state["running"]
                    if running:
                        messages.append(
                            (
                                "error",
                                "A Valid/Invalid Lookup process is already running. Please wait for it to finish.",
                            )
                        )
                    else:
                        worker = threading.Thread(
                            target=run_valid_invalid_in_background,
                            args=(process_name,),
                            daemon=True,
                        )
                        worker.start()
                        messages.append(
                            (
                                "success",
                                f"Valid/Invalid Lookup started in background for '{process_name}'.",
                            )
                        )
                else:
                    messages.append(("error", "Unsupported process type for this input folder."))

    current_path, entries = list_directory_entries(relative_path)
    if current_path is None:
        current_path, entries = list_directory_entries("")
        messages.append(("error", "Invalid directory path selected."))

    entries = annotate_process_button_state(
        entries,
        life_cycle_running,
        life_cycle_process_name,
        valid_invalid_running,
        valid_invalid_process_name,
    )

    current_relative = (
        current_path.relative_to(FILE_PROCESS_DIR.resolve()).as_posix()
        if current_path != FILE_PROCESS_DIR.resolve()
        else ""
    )
    parent_relative = ""
    if current_path != FILE_PROCESS_DIR.resolve():
        parent = current_path.parent
        parent_relative = (
            parent.relative_to(FILE_PROCESS_DIR.resolve()).as_posix()
            if parent != FILE_PROCESS_DIR.resolve()
            else ""
        )

    return render_template(
        "file_process_directories.html",
        messages=messages,
        current_relative=current_relative,
        parent_relative=parent_relative,
        entries=entries,
        delete_disabled=delete_disabled,
        delete_disabled_message=delete_disabled_message,
    )


@app.route("/api/lc-etc-header-keywords", methods=["GET"])
def api_list_lc_etc_header_keywords():
    try:
        records = get_lc_etc_header_keyword_records()
        return jsonify({"file_type": LC_ETC_FILE_TYPE, "keywords": records})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/lc-etc-header-keywords", methods=["POST"])
def api_add_lc_etc_header_keyword():
    payload = request.get_json(silent=True) or {}
    keyword = (payload.get("header_keywords") or request.form.get("header_keywords") or "").strip()
    if not keyword:
        return jsonify({"error": "Header keyword is required."}), 400
    try:
        record = add_header_keyword(LC_ETC_FILE_TYPE, keyword)
        return jsonify({"keyword": record})
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/lc-etc-header-keywords/<int:keyword_id>", methods=["PUT"])
def api_update_lc_etc_header_keyword(keyword_id):
    payload = request.get_json(silent=True) or {}
    keyword = (payload.get("header_keywords") or "").strip()
    if not keyword:
        return jsonify({"error": "Header keyword is required."}), 400
    try:
        record = update_header_keyword(keyword_id, keyword, file_type=LC_ETC_FILE_TYPE)
        return jsonify({"keyword": record})
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/lc-etc-header-keywords/<int:keyword_id>", methods=["DELETE"])
def api_delete_lc_etc_header_keyword(keyword_id):
    try:
        delete_header_keyword(keyword_id, file_type=LC_ETC_FILE_TYPE)
        return jsonify({"ok": True})
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 404
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/life-cycle-merge", methods=["GET", "POST"])
def life_cycle_merge():
    messages = []
    reset_form_after_submit = False
    current_process_name = request.form.get("process_name", "").strip() if request.method == "POST" else ""

    if request.method == "POST":
        action = request.form.get("action")
        process_dir, input_dir = get_process_input_dir(
            MERGE_VALID_LOOKUP_PARENT_FOLDER,
            LIFE_CYCLE_SUBPROCESS_FOLDER,
            current_process_name,
        )

        if action == "upload":
            if not process_dir or not input_dir:
                messages.append(("error", "Process name is required before upload."))
            else:
                input_dir.mkdir(parents=True, exist_ok=True)
                uploaded_files = request.files.getlist("files")
                if not uploaded_files or all(not file.filename for file in uploaded_files):
                    messages.append(("error", "Please select at least one file to upload."))
                else:
                    existing_files = {
                        name.lower()
                        for name in list_uploaded_files(
                            MERGE_VALID_LOOKUP_PARENT_FOLDER,
                            LIFE_CYCLE_SUBPROCESS_FOLDER,
                            current_process_name,
                        )
                    }
                    uploaded_count = 0
                    skipped_duplicates = []
                    skipped_invalid = []

                    for file in uploaded_files:
                        if not file or not file.filename:
                            continue

                        original_name = secure_filename(file.filename)
                        if not original_name:
                            skipped_invalid.append(file.filename)
                            continue

                        if not allowed_file(original_name):
                            skipped_invalid.append(original_name)
                            continue

                        if original_name.lower() in existing_files:
                            skipped_duplicates.append(original_name)
                            continue

                        file.save(input_dir / original_name)
                        existing_files.add(original_name.lower())
                        uploaded_count += 1

                    if uploaded_count:
                        messages.append(("success", f"Uploaded {uploaded_count} file(s)."))
                    if skipped_duplicates:
                        messages.append(
                            ("error", f"Skipped duplicate file(s): {', '.join(skipped_duplicates)}")
                        )
                    if skipped_invalid:
                        messages.append(
                            (
                                "error",
                                "Skipped unsupported/invalid file(s): "
                                f"{', '.join(skipped_invalid)}",
                            )
                        )

        elif action == "delete":
            raw_filename = request.form.get("filename", "")
            filename = secure_filename(raw_filename)

            if not process_dir or not input_dir:
                messages.append(("error", "Process name is required before delete."))
            elif not filename:
                messages.append(("error", "Invalid filename."))
            else:
                file_path = input_dir / filename
                if file_path.exists() and file_path.is_file():
                    try:
                        file_path.unlink()
                        messages.append(("success", f"Deleted file: {filename}"))
                    except PermissionError:
                        messages.append(
                            (
                                "error",
                                "Cannot delete this file because it is currently in use by another process.",
                            )
                        )
                    except OSError as exc:
                        messages.append(("error", f"Delete failed: {exc}"))
                else:
                    messages.append(("error", f"File not found: {filename}"))

        elif action == "process":
            if not current_process_name:
                messages.append(("error", "Process name is required before final submit."))
            else:
                process_dir, input_dir = get_process_input_dir(
                    MERGE_VALID_LOOKUP_PARENT_FOLDER,
                    LIFE_CYCLE_SUBPROCESS_FOLDER,
                    current_process_name,
                )
                uploaded_count = 0
                if input_dir and input_dir.exists():
                    uploaded_count = sum(1 for p in input_dir.iterdir() if p.is_file())
                if uploaded_count < MIN_LIFE_CYCLE_MERGE_FILES:
                    messages.append(
                        (
                            "error",
                            f"Upload at least {MIN_LIFE_CYCLE_MERGE_FILES} files before starting the merge.",
                        )
                    )
                else:
                    with life_cycle_state_lock:
                        running = life_cycle_state["running"]
                    if running:
                        messages.append(
                            (
                                "error",
                                "A life cycle merge is already running. Please wait for it to finish.",
                            )
                        )
                    else:
                        worker = threading.Thread(
                            target=run_life_cycle_merge_in_background,
                            args=(current_process_name,),
                            daemon=True,
                        )
                        worker.start()
                        messages.append(
                            (
                                "success",
                                f"Merge started in background for '{current_process_name}'. "
                                "You can navigate to other pages.",
                            )
                        )
                        reset_form_after_submit = True
                        current_process_name = ""

    with life_cycle_state_lock:
        current_life_cycle_state = dict(life_cycle_state)

    try:
        lc_etc_header_keywords = get_lc_etc_header_keyword_records()
    except Exception as exc:
        lc_etc_header_keywords = []
        messages.append(("error", f"Could not load LC/ETC header keywords from database: {exc}"))

    return render_template(
        "life_cycle_merge.html",
        current_process_name=current_process_name,
        uploaded_files=list_uploaded_files(
            MERGE_VALID_LOOKUP_PARENT_FOLDER,
            LIFE_CYCLE_SUBPROCESS_FOLDER,
            current_process_name,
        )
        if current_process_name
        else [],
        file_process_directories=list_file_process_directories(),
        messages=messages,
        allowed_extensions=sorted(ALLOWED_EXTENSIONS),
        life_cycle_state=current_life_cycle_state,
        reset_form_after_submit=reset_form_after_submit,
        min_life_cycle_merge_files=MIN_LIFE_CYCLE_MERGE_FILES,
        lc_etc_header_keywords=lc_etc_header_keywords,
        lc_etc_file_type=LC_ETC_FILE_TYPE,
    )


@app.route("/valid-invalid-lookup", methods=["GET", "POST"])
def valid_invalid_lookup():
    messages = []
    current_process_name = request.form.get("process_name", "").strip() if request.method == "POST" else ""
    header_info = None
    selected_header_mapping = {}
    process_dir, input_dir, rate_dir = (None, None, None)

    if request.method == "POST":
        action = request.form.get("action")
        process_dir, input_dir, rate_dir = get_valid_invalid_paths(current_process_name)

        if action == "upload":
            if not process_dir or not input_dir or not rate_dir:
                messages.append(("error", "Process name is required before upload."))
            else:
                input_dir.mkdir(parents=True, exist_ok=True)
                rate_dir.mkdir(parents=True, exist_ok=True)

                merged_source = request.form.get("merged_source", "upload").strip()
                merged_file = request.files.get("merged_life_cycle_file")
                rates_file = request.files.get("rates_file")

                if not rates_file or not rates_file.filename:
                    messages.append(("error", "Please select a Rates file."))

                if merged_source == "import":
                    lcm_rel = request.form.get("lcm_output_import", "").strip()
                    if not lcm_rel:
                        messages.append(("error", "Please select a merged file from a Life Cycle Merge output folder."))
                    else:
                        lcm_path = resolve_lcm_output_import_file(lcm_rel)
                        if not lcm_path:
                            messages.append(("error", "Invalid Life Cycle Merge output file selection."))
                else:
                    if not merged_file or not merged_file.filename:
                        messages.append(("error", "Please select a Merged Life Cycle file to upload."))

                if not messages:
                    rates_name = secure_filename(rates_file.filename)
                    if not rates_name or not allowed_file(rates_name):
                        messages.append(("error", "Unsupported/invalid Rates filename or type."))
                    if (rate_dir / rates_name).exists():
                        messages.append(("error", f"Rates file already exists: {rates_name}"))

                if not messages:
                    clear_vil_merge_selection_for_process(current_process_name)
                    clear_vil_confirmed_merged_files(current_process_name)
                    if merged_source == "import":
                        lcm_rel = request.form.get("lcm_output_import", "").strip()
                        lcm_path = resolve_lcm_output_import_file(lcm_rel)
                        _vil_pending_lcm_write()[current_process_name] = lcm_rel
                        session.modified = True
                        rates_file.save(rate_dir / rates_name)
                        messages.append(
                            (
                                "success",
                                "Rates uploaded. Review the header mapping and click Confirm to add the merged file.",
                            )
                        )
                        try:
                            header_info = inspect_valid_invalid_headers(lcm_path)
                        except Exception as exc:
                            messages.append(("error", f"Could not inspect merged file headers: {exc}"))
                    else:
                        merged_name = secure_filename(merged_file.filename)
                        if not merged_name or not allowed_file(merged_name):
                            messages.append(("error", "Unsupported/invalid Merged Life Cycle filename or type."))
                        stage_dir = get_valid_invalid_stage_dir(current_process_name)
                        staged_path = stage_dir / merged_name if stage_dir else None
                        if staged_path and staged_path.exists():
                            messages.append(("error", f"Merged Life Cycle file already exists: {merged_name}"))
                        if not messages:
                            stage_dir.mkdir(parents=True, exist_ok=True)
                            merged_file.save(stage_dir / merged_name)
                            rates_file.save(rate_dir / rates_name)
                            messages.append(
                                (
                                    "success",
                                    "Rates uploaded. Review the header mapping and click Confirm to add the merged file.",
                                )
                            )
                            try:
                                header_info = inspect_valid_invalid_headers(stage_dir / merged_name)
                            except Exception as exc:
                                messages.append(("error", f"Could not inspect merged file headers: {exc}"))

        elif action == "confirm_mapping":
            if not current_process_name:
                messages.append(("error", "Process name is required before confirming header mapping."))
            else:
                process_dir, input_dir, rate_dir = get_valid_invalid_paths(current_process_name)
                if not process_dir or not input_dir or not rate_dir:
                    messages.append(("error", "Invalid process name."))
                else:
                    try:
                        header_info = get_vil_pending_header_info_for_ui(current_process_name)
                        if header_info is None:
                            header_info = get_vil_reconfirm_header_info_for_ui(current_process_name)
                    except Exception as exc:
                        header_info = None
                        messages.append(("error", f"Could not inspect merged file headers: {exc}"))

                    if header_info:
                        for canonical in VALID_INVALID_LOOKUP_REQUIRED_COLUMNS:
                            selected_value = request.form.get(f"header_mapping__{canonical}", "").strip()
                            if selected_value:
                                selected_header_mapping[canonical] = selected_value
                            elif canonical in header_info["detected_mapping"]:
                                selected_header_mapping[canonical] = header_info["detected_mapping"][canonical]

                        missing_user_mappings = [
                            canonical
                            for canonical in VALID_INVALID_LOOKUP_REQUIRED_COLUMNS
                            if canonical not in selected_header_mapping
                        ]
                        invalid_user_mappings = [
                            canonical
                            for canonical, selected_value in selected_header_mapping.items()
                            if selected_value not in header_info["available_columns"]
                        ]

                        if missing_user_mappings:
                            messages.append(
                                (
                                    "error",
                                    "Please select columns for: " + ", ".join(missing_user_mappings),
                                )
                            )
                        elif invalid_user_mappings:
                            messages.append(
                                (
                                    "error",
                                    "Selected column mapping is invalid for: " + ", ".join(invalid_user_mappings),
                                )
                            )
                        else:
                            pending_path, _pending_source = get_vil_pending_merged_preview_path(current_process_name)
                            if pending_path:
                                try:
                                    ok, merged_name = finalize_vil_staged_merged_file(current_process_name)
                                except PermissionError:
                                    ok, merged_name = False, "Could not finalize merged file because it is in use."
                                except OSError as exc:
                                    ok, merged_name = False, f"Could not finalize merged file: {exc}"
                            else:
                                confirmed_files = (
                                    [p for p in input_dir.iterdir() if p.is_file()]
                                    if input_dir and input_dir.exists()
                                    else []
                                )
                                if confirmed_files:
                                    ok, merged_name = True, confirmed_files[0].name
                                else:
                                    ok, merged_name = False, "No merged file is ready for mapping confirmation."

                            if ok:
                                save_vil_header_mapping(current_process_name, selected_header_mapping)
                                header_info = None
                                messages.append(
                                    (
                                        "success",
                                        f"Header mapping confirmed. Merged file '{merged_name}' is now attached to this process.",
                                    )
                                )
                            else:
                                messages.append(("error", merged_name))
                    else:
                        messages.append(
                            (
                                "error",
                                "No merged file is ready for header mapping. Upload a merged file or choose a merge output first.",
                            )
                        )

        elif action == "delete":
            raw_filename = request.form.get("filename", "")
            filename = secure_filename(raw_filename)
            file_kind = request.form.get("file_kind", "")
            if not process_dir or not input_dir or not rate_dir:
                messages.append(("error", "Process name is required before delete."))
            elif not filename:
                messages.append(("error", "Invalid filename."))
            else:
                target_dir = input_dir if file_kind == "merged_life_cycle" else rate_dir
                file_path = target_dir / filename
                if file_path.exists() and file_path.is_file():
                    try:
                        file_path.unlink()
                        messages.append(("success", f"Deleted file: {filename}"))
                        if file_kind == "merged_life_cycle":
                            header_info = None
                            clear_vil_merge_selection_for_process(current_process_name)
                    except PermissionError:
                        messages.append(
                            ("error", "Cannot delete this file because it is currently in use by another process.")
                        )
                    except OSError as exc:
                        messages.append(("error", f"Delete failed: {exc}"))
                else:
                    messages.append(("error", f"File not found: {filename}"))

        elif action == "process":
            if not current_process_name:
                messages.append(("error", "Process name is required before final submit."))
            else:
                process_dir, input_dir, rate_dir = get_valid_invalid_paths(current_process_name)
                if not process_dir or not input_dir or not rate_dir:
                    messages.append(("error", "Invalid process name."))
                else:
                    input_dir.mkdir(parents=True, exist_ok=True)
                    rate_dir.mkdir(parents=True, exist_ok=True)
                    confirmed_merged_files = [p for p in input_dir.iterdir() if p.is_file()]

                    if not confirmed_merged_files:
                        header_info = None
                        messages.append(
                            (
                                "error",
                                "No confirmed merged life cycle file is ready. Upload, map headers, and click Confirm first.",
                            )
                        )
                    else:
                        stored_mapping = load_vil_header_mapping_from_disk(current_process_name)
                        if not is_vil_header_mapping_complete(stored_mapping):
                            header_info = None
                            messages.append(
                                (
                                    "error",
                                    "Header mapping must be confirmed before Submit. Review the mapping section above and click Confirm.",
                                )
                            )
                        else:
                            selected_header_mapping = stored_mapping

                    rate_files = [p for p in rate_dir.iterdir() if p.is_file()] if rate_dir.exists() else []
                    if not rate_files:
                        messages.append(("error", "Please upload a rates file before Submit."))

                with valid_invalid_state_lock:
                    running = valid_invalid_state["running"]
                if running:
                    messages.append(
                        (
                            "error",
                            "A Valid/Invalid Lookup process is already running. Please wait for it to finish.",
                        )
                    )
                elif any(message_type == "error" for message_type, _ in messages):
                    pass
                else:
                    worker = threading.Thread(
                        target=run_valid_invalid_in_background,
                        args=(current_process_name, selected_header_mapping),
                        daemon=True,
                    )
                    worker.start()
                    messages.append(
                        (
                            "success",
                            f"Valid/Invalid Lookup started in background for '{current_process_name}'. "
                            "You can navigate to other pages.",
                        )
                    )

    if request.method != "POST":
        current_process_name = request.args.get("process_name", "").strip()

    process_dir, input_dir, rate_dir = get_valid_invalid_paths(current_process_name)

    if current_process_name and header_info is None:
        try:
            header_info = get_vil_pending_header_info_for_ui(current_process_name)
            if header_info is None:
                header_info = get_vil_reconfirm_header_info_for_ui(current_process_name)
        except Exception as exc:
            messages.append(("error", f"Could not inspect merged file headers: {exc}"))

    pending_lcm_rel = ""
    merged_preview_source = ""
    if current_process_name:
        pending_lcm_rel = (_vil_pending_lcm_read().get(current_process_name) or "").strip()
        _, pending_source = get_vil_pending_merged_preview_path(current_process_name)
        if pending_source:
            merged_preview_source = pending_source
        else:
            _, merged_preview_source = get_vil_reconfirm_header_preview(current_process_name)

    if header_info and not selected_header_mapping:
        stored_mapping = load_vil_header_mapping_from_disk(current_process_name)
        if is_vil_header_mapping_complete(stored_mapping):
            selected_header_mapping = stored_mapping
        else:
            for canonical in VALID_INVALID_LOOKUP_REQUIRED_COLUMNS:
                if canonical in header_info["detected_mapping"]:
                    selected_header_mapping[canonical] = header_info["detected_mapping"][canonical]

    header_mapping_confirmed = (
        vil_header_mapping_is_confirmed(current_process_name) if current_process_name else False
    )

    return render_template(
        "valid_invalid_lookup.html",
        current_process_name=current_process_name,
        merged_life_cycle_files=[p.name for p in input_dir.iterdir() if p.is_file()] if current_process_name and input_dir and input_dir.exists() else [],
        rates_files=[p.name for p in rate_dir.iterdir() if p.is_file()] if current_process_name and rate_dir and rate_dir.exists() else [],
        messages=messages,
        allowed_extensions=sorted(ALLOWED_EXTENSIONS),
        header_info=header_info,
        required_lookup_columns=VALID_INVALID_LOOKUP_REQUIRED_COLUMNS,
        selected_header_mapping=selected_header_mapping,
        lcm_output_files=list_life_cycle_merge_output_files(),
        pending_lcm_import_rel=pending_lcm_rel,
        merged_preview_source=merged_preview_source or "",
        header_mapping_confirmed=header_mapping_confirmed,
    )


if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000, debug=True)
