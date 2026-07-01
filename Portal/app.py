from pathlib import Path
import os
import shutil
import subprocess
import threading
import time
import json

from flask import Flask, abort, jsonify, render_template, request, send_file, session
import pandas as pd
from werkzeug.utils import secure_filename

from Header_Mapping.header_mapping import (
    LIFE_CYCLE_MERGE_HEADER_KEYWORDS,
    VALID_INVALID_LOOKUP_HEADER_MAPPING,
    VALID_INVALID_LOOKUP_REQUIRED_COLUMNS,
)
from Scripts.Life_cycle_merge import _detect_header_row_index, merge_files_in_folder


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024 * 1024  # 5 GB request limit
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-key-change-for-production")

ALLOWED_EXTENSIONS = {".xlsx", ".csv", ".xls"}
FILE_PROCESS_DIR = Path(__file__).resolve().parent / "File_Process"
LIFE_CYCLE_PROCESS_FOLDER = "Life_Cycle_Merge"
VALID_INVALID_PROCESS_FOLDER = "Valid_Invalid_Lookup"
VALID_INVALID_HEADER_SCAN_ROWS = 50

FILE_PROCESS_DIR.mkdir(parents=True, exist_ok=True)
(FILE_PROCESS_DIR / LIFE_CYCLE_PROCESS_FOLDER).mkdir(parents=True, exist_ok=True)
(FILE_PROCESS_DIR / VALID_INVALID_PROCESS_FOLDER).mkdir(parents=True, exist_ok=True)

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


def get_process_input_dir(process_type, process_name):
    safe_process_type = secure_filename(process_type.strip())
    safe_process_name = secure_filename(process_name.strip())
    if not safe_process_type or not safe_process_name:
        return None, None
    process_dir = FILE_PROCESS_DIR / safe_process_type / safe_process_name
    return process_dir, process_dir / "input"


def list_uploaded_files(process_type, process_name):
    _, input_dir = get_process_input_dir(process_type, process_name)
    if not input_dir or not input_dir.exists():
        return []
    return sorted([p.name for p in input_dir.iterdir() if p.is_file()], key=str.lower)


def get_valid_invalid_paths(process_name):
    process_dir, input_dir = get_process_input_dir(VALID_INVALID_PROCESS_FOLDER, process_name)
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
    """Files under File_Process/Life_Cycle_Merge/<process>/output/ (allowed extensions only)."""
    root = FILE_PROCESS_DIR / LIFE_CYCLE_PROCESS_FOLDER
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
    if len(rel_parts) < 4:
        return None
    if rel_parts[0] != LIFE_CYCLE_PROCESS_FOLDER or rel_parts[2] != "output":
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
    _, input_dir, _ = get_valid_invalid_paths(process_name)
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
    if input_dir and input_dir.exists():
        merged_files = [p for p in input_dir.iterdir() if p.is_file()]
        if merged_files:
            return merged_files[0], "uploaded"
    return None, None


def get_valid_invalid_header_info_for_ui(process_name):
    path, _src = get_vil_merged_preview_path(process_name)
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
        output_file_count = 0
        if item.is_dir():
            direct_children = list(item.iterdir())
            folder_count = sum(1 for child in direct_children if child.is_dir())
            file_count = sum(1 for child in direct_children if child.is_file())
            if item.name == "input":
                try:
                    process_name = item.parent.name
                    process_type = item.parent.parent.name
                    output_dir = item.parent / "output"
                    if output_dir.exists() and output_dir.is_dir():
                        output_file_count = sum(
                            1 for child in output_dir.iterdir() if child.is_file() or child.is_dir()
                        )
                except Exception:
                    process_type = None
                    process_name = None
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

            if (
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
    process_dir, input_dir = get_process_input_dir(LIFE_CYCLE_PROCESS_FOLDER, process_name)
    if not process_dir or not input_dir:
        return False, "Please provide a valid process name."

    if not input_dir.exists():
        return False, "Input folder does not exist. Upload files first."

    files_to_process = [p for p in input_dir.iterdir() if p.is_file()]
    if not files_to_process:
        return False, "Please upload at least one file before final submit."

    output_dir = process_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / "merged_output.csv"

    try:
        merge_files_in_folder(
            str(input_dir.resolve()),
            str(output_file.resolve()),
            LIFE_CYCLE_MERGE_HEADER_KEYWORDS,
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


@app.route("/")
def home():
    return render_template("home.html")


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
                    (FILE_PROCESS_DIR / running_process_type / running_process_name).resolve()
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
                    process_type = target_path.parent.parent.name
                except Exception:
                    process_name = ""
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


@app.route("/life-cycle-merge", methods=["GET", "POST"])
def life_cycle_merge():
    messages = []
    reset_form_after_submit = False
    current_process_name = request.form.get("process_name", "").strip() if request.method == "POST" else ""

    if request.method == "POST":
        action = request.form.get("action")
        process_dir, input_dir = get_process_input_dir(LIFE_CYCLE_PROCESS_FOLDER, current_process_name)

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
                        name.lower() for name in list_uploaded_files(LIFE_CYCLE_PROCESS_FOLDER, current_process_name)
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

    return render_template(
        "life_cycle_merge.html",
        current_process_name=current_process_name,
        uploaded_files=list_uploaded_files(LIFE_CYCLE_PROCESS_FOLDER, current_process_name)
        if current_process_name
        else [],
        file_process_directories=list_file_process_directories(),
        messages=messages,
        allowed_extensions=sorted(ALLOWED_EXTENSIONS),
        life_cycle_state=current_life_cycle_state,
        reset_form_after_submit=reset_form_after_submit,
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
                        header_info = get_valid_invalid_header_info_for_ui(current_process_name)
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
                            try:
                                ok, merged_name = finalize_vil_staged_merged_file(current_process_name)
                            except PermissionError:
                                ok, merged_name = False, "Could not finalize merged file because it is in use."
                            except OSError as exc:
                                ok, merged_name = False, f"Could not finalize merged file: {exc}"

                            if ok:
                                messages.append(
                                    (
                                        "success",
                                        f"Header mapping confirmed. Merged file '{merged_name}' is now attached to this process.",
                                    )
                                )
                                try:
                                    header_info = get_valid_invalid_header_info_for_ui(current_process_name)
                                except Exception as exc:
                                    header_info = None
                                    messages.append(("error", f"Could not refresh merged file headers: {exc}"))
                            else:
                                messages.append(("error", merged_name))
                    else:
                        messages.append(
                            (
                                "error",
                                "No staged merged file is ready. Upload a merged file or choose a merge output first.",
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
                        try:
                            header_info = inspect_valid_invalid_headers(confirmed_merged_files[0])
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
                        pass

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
            header_info = get_valid_invalid_header_info_for_ui(current_process_name)
        except Exception as exc:
            messages.append(("error", f"Could not inspect merged file headers: {exc}"))

    pending_lcm_rel = ""
    if current_process_name:
        pending_lcm_rel = (_vil_pending_lcm_read().get(current_process_name) or "").strip()

    if header_info and not selected_header_mapping:
        for canonical in VALID_INVALID_LOOKUP_REQUIRED_COLUMNS:
            if canonical in header_info["detected_mapping"]:
                selected_header_mapping[canonical] = header_info["detected_mapping"][canonical]

    _, merged_preview_source = (
        get_vil_merged_preview_path(current_process_name) if current_process_name else (None, None)
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
    )


if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000, debug=True)
