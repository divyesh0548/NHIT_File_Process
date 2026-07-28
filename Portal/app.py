from pathlib import Path
import os
import shutil
import subprocess
import sys
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
from header_matching import normalize_header_match
from Scripts.Life_cycle_merge import _detect_header_row_index, merge_files_in_folder


REPO_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(REPO_ROOT / ".env")


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024 * 1024  # 5 GB request limit
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-key-change-for-production")

ALLOWED_EXTENSIONS = {".xlsx", ".csv", ".xls"}
FILE_PROCESS_DIR = Path(__file__).resolve().parent / "File_Process"

# Top-level process groups under File_Process/
VALID_INVALID_PARENT_FOLDER = "Valid_Invalid_Process"
LEGACY_VALID_INVALID_PARENT_FOLDER = "Merge_Valid_Lookup"
EXEMPT_QUERY_PARENT_FOLDER = "Exempt_Query"

# Sub-processes under Valid_Invalid_Process/
LIFE_CYCLE_SUBPROCESS_FOLDER = "Life_Cycle_Merge"
VALID_INVALID_SUBPROCESS_FOLDER = "Valid_Invalid_Lookup"

# Sub-processes under Exempt_Query/
MERGE_NORMALIZE_SUBPROCESS_FOLDER = "Merge_&_Normalize"
MERGE_NORMALIZE_LC_ETC_FOLDER = "LC_ETC"
MERGE_NORMALIZE_VRN_FOLDER = "VRN"
MERGE_NORMALIZE_FILE_KINDS = (MERGE_NORMALIZE_LC_ETC_FOLDER, MERGE_NORMALIZE_VRN_FOLDER)
MERGE_NORMALIZE_FILE_KIND_LABELS = {
    MERGE_NORMALIZE_LC_ETC_FOLDER: "LC/ETC",
    MERGE_NORMALIZE_VRN_FOLDER: "VRN",
}
HEADER_KEYWORDS_LC_ETC_VRN_LABEL = "LC/ETC/VRN"
MERGE_NORMALIZE_GROUP_LABEL = "Merge + Normalize"
MERGE_NORMALIZE_OUTPUT_FILENAME = "normalized_and_merged.csv"

MERGE_REMOVE_DUP_SUBPROCESS_FOLDER = "Merge_&_Remove_Duplicate"
MERGE_REMOVE_DUP_GROUP_LABEL = "Merge + Remove Duplicate"
MERGE_REMOVE_DUP_VRN_SLOT = "vrn"
MERGE_REMOVE_DUP_LC_SLOT = "lc_etc"
SEMI_FINAL_OUTPUT_FILENAME = "semi-final-output.csv"
MERGE_REMOVE_DUP_ALLOWED_EXTENSIONS = {".csv"}

FINAL_EXEMPT_SUBPROCESS_FOLDER = "Final_Exempt_Process"
FINAL_EXEMPT_GROUP_LABEL = "Final Exempt Process"
FINAL_EXEMPT_SEMI_SLOT = "semi"
FINAL_EXEMPT_PASS_SLOT = "pass"
FINAL_EXEMPT_INTERMEDIATE_FOLDER = "Intermediate_Files"
FINAL_EXEMPT_CONFIG_FILENAME = "final_exempt_config.json"
FINAL_EXEMPT_REQUIRED_SEMI_COLUMNS = {
    "Veh Reg No.",
    "Date & Time",
    "MOP",
    "Lane No",
    "Description",
    "TC Class",
}
FINAL_EXEMPT_PASS_ALLOWED_EXTENSIONS = {".xls", ".xlsx"}

# Backward-compatible names used in process state and routing checks
LIFE_CYCLE_PROCESS_FOLDER = LIFE_CYCLE_SUBPROCESS_FOLDER
VALID_INVALID_PROCESS_FOLDER = VALID_INVALID_SUBPROCESS_FOLDER

# Display names for the two parent process groups
VALID_INVALID_GROUP_LABEL = "Valid/Invalid Process"
EXEMPT_QUERY_GROUP_LABEL = "Exempt Query"

# Backward-compatible aliases used in routes and older template variable names
MERGE_VALID_LOOKUP_PARENT_FOLDER = VALID_INVALID_PARENT_FOLDER
MERGE_VALID_LOOKUP_GROUP_LABEL = VALID_INVALID_GROUP_LABEL

VIL_HEADER_MAPPING_FILENAME = "header_mapping.json"
MIN_LIFE_CYCLE_MERGE_FILES = 2
MIN_MERGE_NORMALIZE_FILES = 1
VALID_INVALID_HEADER_SCAN_ROWS = 50
MIN_LC_ETC_HEADER_KEYWORDS = 3
PROCESS_STARTED_STATUS_HINT = "Go to Process Management to view status."

FILE_PROCESS_DIR.mkdir(parents=True, exist_ok=True)


def migrate_valid_invalid_parent_folder():
    """Rename legacy Merge_Valid_Lookup to Valid_Invalid_Process when safe."""
    legacy = FILE_PROCESS_DIR / LEGACY_VALID_INVALID_PARENT_FOLDER
    current = FILE_PROCESS_DIR / VALID_INVALID_PARENT_FOLDER
    if legacy.is_dir() and not current.is_dir():
        try:
            legacy.rename(current)
        except OSError:
            pass
    current.mkdir(parents=True, exist_ok=True)
    for subprocess in (LIFE_CYCLE_SUBPROCESS_FOLDER, VALID_INVALID_SUBPROCESS_FOLDER):
        (current / subprocess).mkdir(parents=True, exist_ok=True)


migrate_valid_invalid_parent_folder()
(
    FILE_PROCESS_DIR
    / EXEMPT_QUERY_PARENT_FOLDER
    / MERGE_NORMALIZE_SUBPROCESS_FOLDER
    / MERGE_NORMALIZE_LC_ETC_FOLDER
).mkdir(parents=True, exist_ok=True)
(
    FILE_PROCESS_DIR
    / EXEMPT_QUERY_PARENT_FOLDER
    / MERGE_NORMALIZE_SUBPROCESS_FOLDER
    / MERGE_NORMALIZE_VRN_FOLDER
).mkdir(parents=True, exist_ok=True)
(
    FILE_PROCESS_DIR
    / EXEMPT_QUERY_PARENT_FOLDER
    / MERGE_REMOVE_DUP_SUBPROCESS_FOLDER
).mkdir(parents=True, exist_ok=True)
(
    FILE_PROCESS_DIR
    / EXEMPT_QUERY_PARENT_FOLDER
    / FINAL_EXEMPT_SUBPROCESS_FOLDER
).mkdir(parents=True, exist_ok=True)

def valid_invalid_parent_dir():
    """Resolved Valid/Invalid Process parent folder (prefers new name, falls back to legacy)."""
    current = FILE_PROCESS_DIR / VALID_INVALID_PARENT_FOLDER
    if current.is_dir():
        return current
    legacy = FILE_PROCESS_DIR / LEGACY_VALID_INVALID_PARENT_FOLDER
    if legacy.is_dir():
        return legacy
    return current


def is_valid_invalid_parent_folder(folder_name):
    return folder_name in (VALID_INVALID_PARENT_FOLDER, LEGACY_VALID_INVALID_PARENT_FOLDER)


def get_valid_invalid_process_paths(subprocess_folder, process_name):
    """Resolve process folders under Valid/Invalid Process (new or legacy parent name)."""
    safe_subprocess = secure_filename((subprocess_folder or "").strip())
    safe_process_name = secure_filename((process_name or "").strip())
    if not safe_subprocess or not safe_process_name:
        return None, None

    for parent_name in (VALID_INVALID_PARENT_FOLDER, LEGACY_VALID_INVALID_PARENT_FOLDER):
        process_dir = FILE_PROCESS_DIR / parent_name / safe_subprocess / safe_process_name
        if process_dir.exists():
            return process_dir, process_dir / "input"

    process_dir = (
        FILE_PROCESS_DIR
        / VALID_INVALID_PARENT_FOLDER
        / safe_subprocess
        / safe_process_name
    )
    return process_dir, process_dir / "input"


def list_available_processes():
    """All user-created process folders with parent group and sub-process labels."""
    items = []
    seen_paths = set()

    def add_item(parent_label, subprocess_label, process_dir):
        rel = process_dir.relative_to(FILE_PROCESS_DIR.resolve()).as_posix()
        if rel in seen_paths:
            return
        seen_paths.add(rel)
        items.append(
            {
                "parent_label": parent_label,
                "subprocess_label": subprocess_label,
                "process_name": process_dir.name,
                "relative_path": rel,
            }
        )

    parent_dir = valid_invalid_parent_dir()
    for subprocess, subprocess_label in (
        (LIFE_CYCLE_SUBPROCESS_FOLDER, "Life Cycle Merge"),
        (VALID_INVALID_SUBPROCESS_FOLDER, "Valid/Invalid Lookup"),
    ):
        subprocess_dir = parent_dir / subprocess
        if subprocess_dir.is_dir():
            for proc_dir in sorted(subprocess_dir.iterdir(), key=lambda p: p.name.lower()):
                if proc_dir.is_dir() and not proc_dir.name.startswith("."):
                    add_item(VALID_INVALID_GROUP_LABEL, subprocess_label, proc_dir)

    exempt_dir = FILE_PROCESS_DIR / EXEMPT_QUERY_PARENT_FOLDER
    if exempt_dir.is_dir():
        merge_normalize_root = exempt_dir / MERGE_NORMALIZE_SUBPROCESS_FOLDER
        for kind in MERGE_NORMALIZE_FILE_KINDS:
            kind_dir = merge_normalize_root / kind
            if kind_dir.is_dir():
                kind_label = merge_normalize_file_kind_label(kind)
                for proc_dir in sorted(kind_dir.iterdir(), key=lambda p: p.name.lower()):
                    if proc_dir.is_dir() and not proc_dir.name.startswith("."):
                        add_item(
                            EXEMPT_QUERY_GROUP_LABEL,
                            f"{MERGE_NORMALIZE_GROUP_LABEL} ({kind_label})",
                            proc_dir,
                        )
        merge_remove_root = exempt_dir / MERGE_REMOVE_DUP_SUBPROCESS_FOLDER
        if merge_remove_root.is_dir():
            for proc_dir in sorted(merge_remove_root.iterdir(), key=lambda p: p.name.lower()):
                if proc_dir.is_dir() and not proc_dir.name.startswith("."):
                    add_item(
                        EXEMPT_QUERY_GROUP_LABEL,
                        MERGE_REMOVE_DUP_GROUP_LABEL,
                        proc_dir,
                    )
        final_exempt_root = exempt_dir / FINAL_EXEMPT_SUBPROCESS_FOLDER
        if final_exempt_root.is_dir():
            for proc_dir in sorted(final_exempt_root.iterdir(), key=lambda p: p.name.lower()):
                if proc_dir.is_dir() and not proc_dir.name.startswith("."):
                    add_item(
                        EXEMPT_QUERY_GROUP_LABEL,
                        FINAL_EXEMPT_GROUP_LABEL,
                        proc_dir,
                    )

    return sorted(
        items,
        key=lambda row: (row["parent_label"], row["subprocess_label"], row["process_name"].lower()),
    )


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

merge_normalize_state_lock = threading.Lock()
merge_normalize_state = {
    "running": False,
    "process_type": MERGE_NORMALIZE_SUBPROCESS_FOLDER,
    "file_kind": "",
    "process_name": "",
    "started_at": None,
    "finished_at": None,
    "last_status": "",
}

merge_remove_dup_state_lock = threading.Lock()
merge_remove_dup_state = {
    "running": False,
    "process_type": MERGE_REMOVE_DUP_SUBPROCESS_FOLDER,
    "process_name": "",
    "started_at": None,
    "finished_at": None,
    "last_status": "",
}

final_exempt_state_lock = threading.Lock()
final_exempt_state = {
    "running": False,
    "process_type": FINAL_EXEMPT_SUBPROCESS_FOLDER,
    "process_name": "",
    "started_at": None,
    "finished_at": None,
    "last_status": "",
}


def normalize_merge_normalize_file_kind(file_kind):
    """Map form / legacy folder values to a supported file kind."""
    kind = (file_kind or "").strip().upper()
    if kind in ("LC", "LC/ETC"):
        return MERGE_NORMALIZE_LC_ETC_FOLDER
    if kind in MERGE_NORMALIZE_FILE_KINDS:
        return kind
    return ""


def merge_normalize_file_kind_label(file_kind):
    kind = normalize_merge_normalize_file_kind(file_kind)
    return MERGE_NORMALIZE_FILE_KIND_LABELS.get(kind, kind or "")


def run_subprocess_streaming(command, cwd=None, env=None):
    """
    Run a subprocess, print stdout/stderr to the portal terminal as they arrive,
    and return (returncode, combined_output_text).
    """
    run_env = os.environ.copy()
    if env:
        run_env.update(env)
    # Force unbuffered child output so logs appear while the job is running.
    run_env["PYTHONUNBUFFERED"] = "1"
    run_env.setdefault("PYTHONIOENCODING", "utf-8")

    collected = []
    try:
        process = subprocess.Popen(
            command,
            cwd=cwd,
            env=run_env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        )
    except Exception:
        raise

    assert process.stdout is not None
    for line in process.stdout:
        collected.append(line)
        print(line, end="", flush=True)

    returncode = process.wait()
    return returncode, "".join(collected)


def process_started_success_message(title, process_name, detail=""):
    label = f"{title} ({detail})" if detail else title
    return (
        f"{label} started successfully for '{process_name}'. "
        f"{PROCESS_STARTED_STATUS_HINT}"
    )


def get_merge_normalize_paths(file_kind, process_name):
    """
    Resolve Exempt Query Merge + Normalize folders:

    File_Process/Exempt_Query/Merge_&_Normalize/<LC_ETC|VRN>/<process_name>/
      input/
      output/
    """
    kind = normalize_merge_normalize_file_kind(file_kind)
    if kind not in MERGE_NORMALIZE_FILE_KINDS:
        return None, None, None
    safe_process_name = secure_filename((process_name or "").strip())
    if not safe_process_name:
        return None, None, None
    process_dir = (
        FILE_PROCESS_DIR
        / EXEMPT_QUERY_PARENT_FOLDER
        / MERGE_NORMALIZE_SUBPROCESS_FOLDER
        / kind
        / safe_process_name
    )
    return process_dir, process_dir / "input", process_dir / "output"


def ensure_merge_normalize_process_dirs(file_kind, process_name):
    process_dir, input_dir, output_dir = get_merge_normalize_paths(file_kind, process_name)
    if not process_dir or not input_dir or not output_dir:
        return None, None, None
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    return process_dir, input_dir, output_dir


def list_merge_normalize_uploaded_files(file_kind, process_name):
    _, input_dir, _ = get_merge_normalize_paths(file_kind, process_name)
    if not input_dir or not input_dir.exists():
        return []
    return sorted([p.name for p in input_dir.iterdir() if p.is_file()], key=str.lower)


def get_merge_remove_dup_paths(process_name):
    """
    File_Process/Exempt_Query/Merge_&_Remove_Duplicate/<process_name>/
      input/vrn/
      input/lc_etc/
      output/
    """
    safe_process_name = secure_filename((process_name or "").strip())
    if not safe_process_name:
        return None, None, None, None
    process_dir = (
        FILE_PROCESS_DIR
        / EXEMPT_QUERY_PARENT_FOLDER
        / MERGE_REMOVE_DUP_SUBPROCESS_FOLDER
        / safe_process_name
    )
    input_dir = process_dir / "input"
    return (
        process_dir,
        input_dir / MERGE_REMOVE_DUP_VRN_SLOT,
        input_dir / MERGE_REMOVE_DUP_LC_SLOT,
        process_dir / "output",
    )


def ensure_merge_remove_dup_dirs(process_name):
    process_dir, vrn_dir, lc_dir, output_dir = get_merge_remove_dup_paths(process_name)
    if not process_dir or not vrn_dir or not lc_dir or not output_dir:
        return None, None, None, None
    vrn_dir.mkdir(parents=True, exist_ok=True)
    lc_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    return process_dir, vrn_dir, lc_dir, output_dir


def get_final_exempt_paths(process_name):
    """
    File_Process/Exempt_Query/Final_Exempt_Process/<process_name>/
      input/semi/  input/pass/  Intermediate_Files/  output/
    """
    safe_process_name = secure_filename((process_name or "").strip())
    if not safe_process_name:
        return None, None, None, None, None
    process_dir = (
        FILE_PROCESS_DIR
        / EXEMPT_QUERY_PARENT_FOLDER
        / FINAL_EXEMPT_SUBPROCESS_FOLDER
        / safe_process_name
    )
    input_dir = process_dir / "input"
    return (
        process_dir,
        input_dir / FINAL_EXEMPT_SEMI_SLOT,
        input_dir / FINAL_EXEMPT_PASS_SLOT,
        process_dir / FINAL_EXEMPT_INTERMEDIATE_FOLDER,
        process_dir / "output",
    )


def ensure_final_exempt_dirs(process_name):
    process_dir, semi_dir, pass_dir, work_dir, output_dir = get_final_exempt_paths(process_name)
    if not process_dir:
        return None, None, None, None, None
    for directory in (semi_dir, pass_dir, work_dir, output_dir):
        directory.mkdir(parents=True, exist_ok=True)
    return process_dir, semi_dir, pass_dir, work_dir, output_dir


def get_final_exempt_semi_file(process_name):
    _, semi_dir, _, _, _ = get_final_exempt_paths(process_name)
    if not semi_dir or not semi_dir.is_dir():
        return None
    files = sorted((p for p in semi_dir.iterdir() if p.is_file() and p.suffix.lower() == ".csv"), key=lambda p: p.name.lower())
    return files[0] if files else None


def list_final_exempt_pass_files(process_name):
    _, _, pass_dir, _, _ = get_final_exempt_paths(process_name)
    if not pass_dir or not pass_dir.is_dir():
        return []
    return sorted((p.name for p in pass_dir.iterdir() if p.is_file()), key=str.lower)


def get_final_exempt_config_path(process_name):
    process_dir, _, _, _, _ = get_final_exempt_paths(process_name)
    return process_dir / FINAL_EXEMPT_CONFIG_FILENAME if process_dir else None


def load_final_exempt_config(process_name):
    config_path = get_final_exempt_config_path(process_name)
    if not config_path or not config_path.is_file():
        return {}
    try:
        value = json.loads(config_path.read_text(encoding="utf-8"))
        return value if isinstance(value, dict) else {}
    except (OSError, ValueError, TypeError):
        return {}


def save_final_exempt_config(process_name, plaza_name, output_file_name):
    config_path = get_final_exempt_config_path(process_name)
    if not config_path:
        return False
    config_path.parent.mkdir(parents=True, exist_ok=True)
    config_path.write_text(
        json.dumps(
            {"plaza_name": plaza_name, "output_file_name": output_file_name},
            indent=2,
        ),
        encoding="utf-8",
    )
    return True


def allowed_final_exempt_pass_file(filename):
    return Path(filename).suffix.lower() in FINAL_EXEMPT_PASS_ALLOWED_EXTENSIONS


def final_exempt_output_file_name(process_name, requested_name=""):
    name = secure_filename((requested_name or "").strip())
    if not name:
        name = f"{secure_filename(process_name)}_final.xlsx"
    if Path(name).suffix.lower() != ".xlsx":
        name = f"{Path(name).stem}.xlsx"
    return name


def list_merge_remove_dup_output_files():
    root = FILE_PROCESS_DIR / EXEMPT_QUERY_PARENT_FOLDER / MERGE_REMOVE_DUP_SUBPROCESS_FOLDER
    if not root.is_dir():
        return []
    results = []
    for process_dir in sorted((p for p in root.iterdir() if p.is_dir()), key=lambda p: p.name.lower()):
        candidate = process_dir / "output" / SEMI_FINAL_OUTPUT_FILENAME
        if candidate.is_file():
            results.append(
                {
                    "relative_path": candidate.relative_to(FILE_PROCESS_DIR.resolve()).as_posix(),
                    "label": f"{process_dir.name} / output / {SEMI_FINAL_OUTPUT_FILENAME}",
                }
            )
    return results


def resolve_merge_remove_dup_output_import(relative_path):
    target = safe_path_from_relative(relative_path)
    if not target or not target.is_file() or target.name != SEMI_FINAL_OUTPUT_FILENAME:
        return None
    try:
        parts = target.relative_to(FILE_PROCESS_DIR.resolve()).parts
    except ValueError:
        return None
    if (
        len(parts) != 5
        or parts[0] != EXEMPT_QUERY_PARENT_FOLDER
        or parts[1] != MERGE_REMOVE_DUP_SUBPROCESS_FOLDER
        or parts[3] != "output"
    ):
        return None
    return target


def get_final_exempt_plazas():
    codes_path = (
        Path(__file__).resolve().parent
        / "Scripts"
        / "Excempy_Query"
        / "Final_7_scripts"
        / "codes_dump.json"
    )
    try:
        codes = json.loads(codes_path.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return []
    return sorted(str(name).strip() for name in codes if str(name).strip())


def allowed_merge_remove_dup_file(filename):
    return Path(filename).suffix.lower() in MERGE_REMOVE_DUP_ALLOWED_EXTENSIONS


def list_merge_remove_dup_slot_files(process_name, slot=MERGE_REMOVE_DUP_VRN_SLOT):
    _, vrn_dir, lc_dir, _ = get_merge_remove_dup_paths(process_name)
    slot_dir = vrn_dir if slot == MERGE_REMOVE_DUP_VRN_SLOT else lc_dir
    if not slot_dir or not slot_dir.exists():
        return []
    return sorted([p.name for p in slot_dir.iterdir() if p.is_file()], key=str.lower)


def get_merge_remove_dup_slot_file(process_name, slot):
    _, vrn_dir, lc_dir, _ = get_merge_remove_dup_paths(process_name)
    slot_dir = vrn_dir if slot == MERGE_REMOVE_DUP_VRN_SLOT else lc_dir
    if not slot_dir or not slot_dir.exists():
        return None
    files = sorted([p for p in slot_dir.iterdir() if p.is_file()], key=lambda p: p.name.lower())
    return files[0] if files else None


def clear_merge_remove_dup_slot(process_name, slot):
    _, vrn_dir, lc_dir, _ = get_merge_remove_dup_paths(process_name)
    slot_dir = vrn_dir if slot == MERGE_REMOVE_DUP_VRN_SLOT else lc_dir
    if not slot_dir or not slot_dir.exists():
        return
    for item in slot_dir.iterdir():
        if item.is_file():
            item.unlink()


def copy_to_merge_remove_dup_slot(process_name, slot, source_path):
    process_dir, vrn_dir, lc_dir, _ = ensure_merge_remove_dup_dirs(process_name)
    if not process_dir:
        return None
    slot_dir = vrn_dir if slot == MERGE_REMOVE_DUP_VRN_SLOT else lc_dir
    clear_merge_remove_dup_slot(process_name, slot)
    destination = slot_dir / secure_filename(Path(source_path).name)
    shutil.copy2(source_path, destination)
    return destination


def list_merge_normalize_output_files(file_kind=None):
    """CSV outputs from Merge + Normalize (VRN and/or LC/ETC)."""
    root = FILE_PROCESS_DIR / EXEMPT_QUERY_PARENT_FOLDER / MERGE_NORMALIZE_SUBPROCESS_FOLDER
    if not root.exists():
        return []
    kinds = [normalize_merge_normalize_file_kind(file_kind)] if file_kind else list(MERGE_NORMALIZE_FILE_KINDS)
    results = []
    for kind in kinds:
        if not kind:
            continue
        kind_dir = root / kind
        if not kind_dir.is_dir():
            continue
        for proc_dir in sorted(kind_dir.iterdir(), key=lambda p: p.name.lower()):
            if not proc_dir.is_dir():
                continue
            out_file = proc_dir / "output" / MERGE_NORMALIZE_OUTPUT_FILENAME
            if not out_file.is_file():
                continue
            rel = out_file.relative_to(FILE_PROCESS_DIR.resolve()).as_posix()
            results.append(
                {
                    "relative_path": rel,
                    "label": (
                        f"{MERGE_NORMALIZE_FILE_KIND_LABELS[kind]} / {proc_dir.name} / "
                        f"output / {out_file.name}"
                    ),
                    "process_name": proc_dir.name,
                    "file_kind": kind,
                    "filename": out_file.name,
                }
            )
    return results


def resolve_merge_normalize_output_import(relative_path, expected_kind):
    """Allow only Merge + Normalize output CSV for the expected VRN or LC/ETC kind."""
    target = safe_path_from_relative(relative_path)
    if target is None or not target.is_file():
        return None
    if target.name != MERGE_NORMALIZE_OUTPUT_FILENAME:
        return None
    if not allowed_merge_remove_dup_file(target.name):
        return None
    try:
        rel_parts = target.relative_to(FILE_PROCESS_DIR.resolve()).parts
    except ValueError:
        return None
    kind = normalize_merge_normalize_file_kind(expected_kind)
    if (
        len(rel_parts) < 6
        or rel_parts[0] != EXEMPT_QUERY_PARENT_FOLDER
        or rel_parts[1] != MERGE_NORMALIZE_SUBPROCESS_FOLDER
        or rel_parts[2] != kind
        or rel_parts[4] != "output"
    ):
        return None
    return target


def parse_input_folder_context(input_path):
    """
    Derive process metadata from an .../input folder.

    Supports:
      <parent>/<subprocess>/<process>/input
      Exempt_Query/Merge_&_Normalize/<LC_ETC|VRN>/<process>/input
    """
    input_path = Path(input_path).resolve()
    if input_path.name != "input":
        return None
    try:
        parts = input_path.relative_to(FILE_PROCESS_DIR.resolve()).parts
    except ValueError:
        return None

    if (
        len(parts) >= 5
        and parts[0] == EXEMPT_QUERY_PARENT_FOLDER
        and parts[1] == MERGE_NORMALIZE_SUBPROCESS_FOLDER
        and parts[-1] == "input"
    ):
        file_kind = normalize_merge_normalize_file_kind(parts[2])
        if not file_kind:
            return None
        return {
            "parent_folder": parts[0],
            "subprocess_folder": parts[1],
            "file_kind": file_kind,
            "process_name": parts[3],
            "process_type": parts[1],
            "process_dir": input_path.parent,
        }

    if (
        len(parts) >= 4
        and parts[0] == EXEMPT_QUERY_PARENT_FOLDER
        and parts[1] == MERGE_REMOVE_DUP_SUBPROCESS_FOLDER
        and parts[-1] == "input"
    ):
        return {
            "parent_folder": parts[0],
            "subprocess_folder": parts[1],
            "file_kind": None,
            "process_name": parts[2],
            "process_type": parts[1],
            "process_dir": input_path.parent,
        }

    if len(parts) >= 4 and parts[-1] == "input":
        return {
            "parent_folder": parts[0],
            "subprocess_folder": parts[1],
            "file_kind": None,
            "process_name": parts[2],
            "process_type": parts[1],
            "process_dir": input_path.parent,
        }
    return None


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
    if is_valid_invalid_parent_folder(parent_folder):
        _, input_dir = get_valid_invalid_process_paths(subprocess_folder, process_name)
    else:
        _, input_dir = get_process_input_dir(parent_folder, subprocess_folder, process_name)
    if not input_dir or not input_dir.exists():
        return []
    return sorted([p.name for p in input_dir.iterdir() if p.is_file()], key=str.lower)


def get_valid_invalid_paths(process_name):
    """
    Resolve Valid/Invalid Lookup process folders:

    File_Process/Valid_Invalid_Process/Valid_Invalid_Lookup/<process_name>/
      input/   — confirmed merged life cycle file
      rate/    — rates file (sibling of input, not nested under it)
      output/  — valid_table.csv / invalid_table.csv
      .stage/  — temporary staged merge before Confirm
    """
    process_dir, input_dir = get_valid_invalid_process_paths(
        VALID_INVALID_SUBPROCESS_FOLDER,
        process_name,
    )
    if not process_dir or not input_dir:
        return None, None, None
    rate_dir = process_dir / "rate"
    return process_dir, input_dir, rate_dir


RESERVED_PROCESS_NAMES = {
    VALID_INVALID_PARENT_FOLDER.lower(),
    LEGACY_VALID_INVALID_PARENT_FOLDER.lower(),
    LIFE_CYCLE_SUBPROCESS_FOLDER.lower(),
    VALID_INVALID_SUBPROCESS_FOLDER.lower(),
    EXEMPT_QUERY_PARENT_FOLDER.lower(),
    MERGE_NORMALIZE_SUBPROCESS_FOLDER.lower(),
    "merge__normalize",
    "merge_and_normalize",
    MERGE_REMOVE_DUP_SUBPROCESS_FOLDER.lower(),
    "merge__remove_duplicate",
    "merge_and_remove_duplicate",
    MERGE_NORMALIZE_LC_ETC_FOLDER.lower(),
    MERGE_REMOVE_DUP_LC_SLOT.lower(),
    "lc",
    "lc/etc",
    MERGE_NORMALIZE_VRN_FOLDER.lower(),
    MERGE_REMOVE_DUP_VRN_SLOT.lower(),
    "input",
    "output",
    "rate",
    "file_process",
}


def is_reserved_process_name(process_name):
    safe_name = secure_filename((process_name or "").strip())
    return not safe_name or safe_name.lower() in RESERVED_PROCESS_NAMES


def ensure_valid_invalid_process_dirs(process_name):
    """Create the standard VIL folder layout; migrate legacy input/rate → rate/."""
    process_dir, input_dir, rate_dir = get_valid_invalid_paths(process_name)
    if not process_dir or not input_dir or not rate_dir:
        return None, None, None

    input_dir.mkdir(parents=True, exist_ok=True)
    rate_dir.mkdir(parents=True, exist_ok=True)
    (process_dir / "output").mkdir(parents=True, exist_ok=True)

    # Older builds nested rates under input/rate — move files up to process/rate.
    legacy_rate_dir = input_dir / "rate"
    if legacy_rate_dir.exists() and legacy_rate_dir.is_dir() and legacy_rate_dir.resolve() != rate_dir.resolve():
        for item in list(legacy_rate_dir.iterdir()):
            if item.is_file():
                destination = rate_dir / item.name
                if destination.exists():
                    item.unlink()
                else:
                    shutil.move(str(item), str(destination))
        try:
            legacy_rate_dir.rmdir()
        except OSError:
            pass

    return process_dir, input_dir, rate_dir


def get_valid_invalid_stage_dir(process_name):
    process_dir, _, _ = get_valid_invalid_paths(process_name)
    if not process_dir:
        return None
    return process_dir / ".stage"


def list_file_process_directories():
    return sorted([p.name for p in FILE_PROCESS_DIR.iterdir() if p.is_dir()], key=str.lower)


def fetch_lc_etc_header_keywords_or_error(min_required=MIN_LC_ETC_HEADER_KEYWORDS):
    """Load LC/ETC/VRN header keywords from DB; return (keywords, error_message)."""
    try:
        keywords = get_lc_etc_header_keyword_strings()
    except Exception as exc:
        return None, (
            f"Could not load {HEADER_KEYWORDS_LC_ETC_VRN_LABEL} header keywords "
            f"from database: {exc}"
        )
    if len(keywords) < min_required:
        return None, (
            f"Configure at least {min_required} header keywords for "
            f"{HEADER_KEYWORDS_LC_ETC_VRN_LABEL} before running."
        )
    return keywords, None


def normalize_lookup_column_name(value):
    return normalize_header_match(value)


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
    """Files under File_Process/Valid_Invalid_Process/Life_Cycle_Merge/<process>/output/."""
    root = valid_invalid_parent_dir() / LIFE_CYCLE_SUBPROCESS_FOLDER
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
        not is_valid_invalid_parent_folder(rel_parts[0])
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
    process_dir, input_dir, _ = ensure_valid_invalid_process_dirs(process_name)
    if not process_dir or not input_dir:
        return False, "Invalid process name."

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
        file_kind = None
        output_file_count = 0
        if item.is_dir():
            direct_children = list(item.iterdir())
            folder_count = sum(1 for child in direct_children if child.is_dir())
            file_count = sum(1 for child in direct_children if child.is_file())
            if item.name == "input":
                context = parse_input_folder_context(item)
                if context:
                    process_name = context["process_name"]
                    process_type = context["process_type"]
                    parent_folder = context["parent_folder"]
                    file_kind = context.get("file_kind")
                    output_dir = item.parent / "output"
                    if output_dir.exists() and output_dir.is_dir():
                        output_file_count = sum(
                            1 for child in output_dir.iterdir() if child.is_file() or child.is_dir()
                        )
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
                "file_kind": file_kind if item.is_dir() and item.name == "input" else None,
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
    merge_normalize_running=False,
    merge_normalize_process_name="",
    merge_normalize_file_kind="",
    merge_remove_dup_running=False,
    merge_remove_dup_process_name="",
    final_exempt_running=False,
    final_exempt_process_name="",
):
    annotated_entries = []
    for entry in entries:
        process_disabled = False
        process_disabled_message = ""

        if entry.get("is_dir") and entry.get("name") == "input":
            process_type = entry.get("process_type") or ""
            process_name = entry.get("process_name") or ""
            parent_folder = entry.get("parent_folder") or ""
            file_kind = entry.get("file_kind") or ""

            supported_groups = {
                VALID_INVALID_PARENT_FOLDER,
                LEGACY_VALID_INVALID_PARENT_FOLDER,
                EXEMPT_QUERY_PARENT_FOLDER,
            }
            exempt_subprocesses = {
                MERGE_NORMALIZE_SUBPROCESS_FOLDER,
                MERGE_REMOVE_DUP_SUBPROCESS_FOLDER,
                FINAL_EXEMPT_SUBPROCESS_FOLDER,
            }
            if parent_folder and parent_folder not in supported_groups:
                process_disabled = True
                process_disabled_message = (
                    "Processing from the directory browser is only available for "
                    f"'{VALID_INVALID_GROUP_LABEL}' and {EXEMPT_QUERY_GROUP_LABEL} sub-processes."
                )
            elif (
                parent_folder == EXEMPT_QUERY_PARENT_FOLDER
                and process_type not in exempt_subprocesses
            ):
                process_disabled = True
                process_disabled_message = (
                    "Processing from the directory browser is only available for "
                    f"Exempt Query → {MERGE_NORMALIZE_GROUP_LABEL}, {MERGE_REMOVE_DUP_GROUP_LABEL}, and {FINAL_EXEMPT_GROUP_LABEL}."
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
            elif (
                merge_normalize_running
                and process_type == MERGE_NORMALIZE_SUBPROCESS_FOLDER
                and process_name == merge_normalize_process_name
                and file_kind == merge_normalize_file_kind
            ):
                process_disabled = True
                process_disabled_message = (
                    f"Process is disabled while {MERGE_NORMALIZE_GROUP_LABEL} "
                    f"({merge_normalize_file_kind_label(file_kind)}) is already running for '{process_name}'."
                )
            elif (
                merge_remove_dup_running
                and process_type == MERGE_REMOVE_DUP_SUBPROCESS_FOLDER
                and process_name == merge_remove_dup_process_name
            ):
                process_disabled = True
                process_disabled_message = (
                    f"Process is disabled while {MERGE_REMOVE_DUP_GROUP_LABEL} "
                    f"is already running for '{process_name}'."
                )
            elif (
                final_exempt_running
                and process_type == FINAL_EXEMPT_SUBPROCESS_FOLDER
                and process_name == final_exempt_process_name
            ):
                process_disabled = True
                process_disabled_message = (
                    f"Process is disabled while {FINAL_EXEMPT_GROUP_LABEL} "
                    f"is already running for '{process_name}'."
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
    process_dir, input_dir = get_valid_invalid_process_paths(
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

    header_keywords, kw_error = fetch_lc_etc_header_keywords_or_error()
    if kw_error:
        return False, kw_error

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


def process_merge_normalize_files(file_kind, process_name):
    if is_reserved_process_name(process_name):
        return False, (
            f"Process name '{process_name}' is reserved. "
            "Choose a different name (not Merge_&_Normalize, LC/ETC, VRN, input, or output)."
        )

    process_dir, input_dir, output_dir = ensure_merge_normalize_process_dirs(file_kind, process_name)
    if not process_dir or not input_dir or not output_dir:
        return False, "Please provide a valid process name and file type (LC/ETC or VRN)."

    if not input_dir.exists():
        return False, "Input folder does not exist. Upload files first."

    files_to_process = [p for p in input_dir.iterdir() if p.is_file()]
    if not files_to_process:
        return False, "Please upload at least one file before final submit."

    kind = normalize_merge_normalize_file_kind(file_kind)
    script_dir = Path(__file__).resolve().parent / "Scripts" / "Excempy_Query"
    header_keywords, kw_error = fetch_lc_etc_header_keywords_or_error()
    if kw_error:
        return False, kw_error

    if kind == MERGE_NORMALIZE_VRN_FOLDER:
        # script_path = script_dir / "normalize_vrn.py"
        script_path = script_dir / "normalize_vrn_faster.py"
    elif kind == MERGE_NORMALIZE_LC_ETC_FOLDER:
        # script_path = script_dir / "normalize_lcy.py"
        script_path = script_dir / "normalize_lcy_faster.py"
    else:
        return False, "File type must be LC/ETC or VRN."

    if not script_path.exists():
        return False, f"Normalize script not found: {script_path.name}"

    output_file = output_dir / MERGE_NORMALIZE_OUTPUT_FILENAME
    env = os.environ.copy()
    env["MERGE_NORMALIZE_INPUT_FOLDER"] = str(input_dir.resolve())
    env["MERGE_NORMALIZE_OUTPUT_FILE"] = str(output_file.resolve())
    env["MERGE_NORMALIZE_HEADER_KEYWORDS"] = json.dumps(header_keywords)

    try:
        returncode, combined = run_subprocess_streaming(
            ["python", "-u", str(script_path)],
            cwd=str(script_dir),
            env=env,
        )
    except Exception as exc:
        return False, f"{MERGE_NORMALIZE_GROUP_LABEL} failed to start: {exc}"

    if returncode != 0:
        fail_line = ""
        for line in reversed(combined.splitlines()):
            if "[CONVERT-FAIL]" in line or "Header detection failed" in line:
                fail_line = line.strip()
                break
        detail = fail_line or combined.strip()[-800:] or f"exit code {returncode}"
        return (
            False,
            f"{MERGE_NORMALIZE_GROUP_LABEL} ({kind}) failed: {detail}",
        )

    if not output_file.exists():
        return (
            False,
            "Normalize/merge did not create an output file. Check input formats and header keywords.",
        )

    return (
        True,
        f"{MERGE_NORMALIZE_GROUP_LABEL} finished for '{process_dir.name}' "
        f"({merge_normalize_file_kind_label(kind)}). "
        f"Output: output/{output_file.name}",
    )


def run_merge_normalize_in_background(file_kind, process_name):
    with merge_normalize_state_lock:
        merge_normalize_state["running"] = True
        merge_normalize_state["process_type"] = MERGE_NORMALIZE_SUBPROCESS_FOLDER
        merge_normalize_state["file_kind"] = normalize_merge_normalize_file_kind(file_kind)
        merge_normalize_state["process_name"] = process_name
        merge_normalize_state["started_at"] = time.time()
        merge_normalize_state["finished_at"] = None
        merge_normalize_state["last_status"] = f"{MERGE_NORMALIZE_GROUP_LABEL} started."

    ok, process_message = process_merge_normalize_files(file_kind, process_name)

    with merge_normalize_state_lock:
        merge_normalize_state["running"] = False
        merge_normalize_state["finished_at"] = time.time()
        merge_normalize_state["last_status"] = process_message if ok else f"FAILED: {process_message}"


def process_merge_remove_dup_files(process_name):
    if is_reserved_process_name(process_name):
        return False, (
            f"Process name '{process_name}' is reserved. "
            "Choose a different name (not Merge_&_Remove_Duplicate, vrn, lc_etc, input, or output)."
        )

    process_dir, vrn_dir, lc_dir, output_dir = ensure_merge_remove_dup_dirs(process_name)
    if not process_dir or not vrn_dir or not lc_dir or not output_dir:
        return False, "Please provide a valid process name."

    vrn_file = get_merge_remove_dup_slot_file(process_name, MERGE_REMOVE_DUP_VRN_SLOT)
    lc_file = get_merge_remove_dup_slot_file(process_name, MERGE_REMOVE_DUP_LC_SLOT)
    if not vrn_file:
        return False, "VRN file is required. Upload or import from Merge + Normalize (VRN) output."
    if not lc_file:
        return False, "LC/ETC file is required. Upload or import from Merge + Normalize (LC/ETC) output."

    script_dir = Path(__file__).resolve().parent / "Scripts" / "Excempy_Query"
    script_path = script_dir / "vrn_lc_final_prcess.py"
    if not script_path.exists():
        return False, f"Process script not found: {script_path.name}"

    output_file = output_dir / SEMI_FINAL_OUTPUT_FILENAME
    env = os.environ.copy()
    env["VRN_LC_FINAL_VRN_FILE"] = str(vrn_file.resolve())
    env["VRN_LC_FINAL_LC_FILE"] = str(lc_file.resolve())
    env["VRN_LC_FINAL_OUTPUT_FILE"] = str(output_file.resolve())
    env["VRN_LC_FINAL_DEBUG_EXPORT"] = "0"

    try:
        result = subprocess.run(
            ["python", str(script_path)],
            cwd=str(script_dir),
            env=env,
            check=False,
        )
    except Exception as exc:
        return False, f"{MERGE_REMOVE_DUP_GROUP_LABEL} failed to start: {exc}"

    if result.returncode != 0:
        return False, f"{MERGE_REMOVE_DUP_GROUP_LABEL} failed with exit code {result.returncode}"

    if not output_file.exists():
        return False, f"{MERGE_REMOVE_DUP_GROUP_LABEL} did not create output/{SEMI_FINAL_OUTPUT_FILENAME}."

    return (
        True,
        f"{MERGE_REMOVE_DUP_GROUP_LABEL} finished for '{process_dir.name}'. "
        f"Output: output/{SEMI_FINAL_OUTPUT_FILENAME}",
    )


def run_merge_remove_dup_in_background(process_name):
    with merge_remove_dup_state_lock:
        merge_remove_dup_state["running"] = True
        merge_remove_dup_state["process_type"] = MERGE_REMOVE_DUP_SUBPROCESS_FOLDER
        merge_remove_dup_state["process_name"] = process_name
        merge_remove_dup_state["started_at"] = time.time()
        merge_remove_dup_state["finished_at"] = None
        merge_remove_dup_state["last_status"] = f"{MERGE_REMOVE_DUP_GROUP_LABEL} started."

    ok, process_message = process_merge_remove_dup_files(process_name)

    with merge_remove_dup_state_lock:
        merge_remove_dup_state["running"] = False
        merge_remove_dup_state["finished_at"] = time.time()
        merge_remove_dup_state["last_status"] = process_message if ok else f"FAILED: {process_message}"


def _set_final_exempt_progress(message):
    with final_exempt_state_lock:
        if final_exempt_state["running"]:
            final_exempt_state["last_status"] = message


def _format_final_exempt_subprocess_error(stage_name, returncode, combined_output=""):
    detail = (combined_output or "").strip()
    if detail:
        detail = detail[-2000:]
        return f"{stage_name} failed with exit code {returncode}. Details: {detail}"
    return f"{stage_name} failed with exit code {returncode}."


def _validate_final_exempt_semi_file(semi_file):
    try:
        headers = pd.read_csv(semi_file, nrows=0).columns
    except Exception as exc:
        return False, f"Could not read the semi-final CSV: {exc}"
    normalized_headers = {
        str(header).replace("\ufeff", "").strip()
        for header in headers
    }
    missing = sorted(FINAL_EXEMPT_REQUIRED_SEMI_COLUMNS - normalized_headers)
    if missing:
        return (
            False,
            "Semi-final CSV is missing required column(s): " + ", ".join(missing),
        )
    return True, ""


def _run_final_exempt_script(script_dir, script_name, stage_name, env, expected_files=()):
    _set_final_exempt_progress(stage_name)
    try:
        returncode, combined = run_subprocess_streaming(
            [sys.executable, "-u", script_name],
            cwd=str(script_dir),
            env=env,
        )
    except Exception as exc:
        return False, f"{stage_name} could not start: {exc}"
    if returncode != 0:
        return False, _format_final_exempt_subprocess_error(stage_name, returncode, combined)
    missing = [Path(path).name for path in expected_files if not Path(path).is_file()]
    if missing:
        return False, f"{stage_name} completed but did not create: {', '.join(missing)}."
    return True, ""


def process_final_exempt_files(process_name):
    if is_reserved_process_name(process_name):
        return False, "Please provide a valid process name."

    process_dir, semi_dir, pass_dir, work_dir, output_dir = ensure_final_exempt_dirs(process_name)
    if not process_dir:
        return False, "Please provide a valid process name."
    semi_file = get_final_exempt_semi_file(process_name)
    pass_files = [
        path for path in pass_dir.iterdir()
        if path.is_file() and allowed_final_exempt_pass_file(path.name)
    ]
    if not semi_file:
        return False, "A semi-final CSV is required. Upload one or import a Merge + Remove Duplicate output."
    if not pass_files:
        return False, "Upload at least one .xls or .xlsx pass file before starting."

    valid_semi, semi_error = _validate_final_exempt_semi_file(semi_file)
    if not valid_semi:
        return False, semi_error

    process_config = load_final_exempt_config(process_name)
    plaza_name = str(process_config.get("plaza_name") or "").strip()
    output_file_name = final_exempt_output_file_name(
        process_name, process_config.get("output_file_name")
    )
    if not plaza_name:
        return False, "Select a plaza before starting the Final Exempt Process."
    if plaza_name not in get_final_exempt_plazas():
        return False, "Selected plaza is not available in codes_dump.json. Save the process again with a valid plaza."

    script_dir = (
        Path(__file__).resolve().parent
        / "Scripts"
        / "Excempy_Query"
        / "Final_7_scripts"
    )
    required_scripts = (
        "Pass_file_seperator.py",
        "main_script.py",
        "RF3_condition.py",
        "combined_with_rf3.py",
        "Date_validity_check.py",
        "Date_validity_check_LT.py",
        "Return_Journey_Logic.py",
    )
    missing_scripts = [name for name in required_scripts if not (script_dir / name).is_file()]
    if missing_scripts:
        return False, "Required final-stage script(s) not found: " + ", ".join(missing_scripts)
    codes_source = script_dir / "codes_dump.json"
    if not codes_source.is_file():
        return False, "codes_dump.json was not found with the final-stage scripts."

    # A re-run must not accidentally consume stale LT or intermediary files from
    # an earlier attempt. Original user inputs remain untouched under input/.
    try:
        if work_dir.exists():
            shutil.rmtree(work_dir)
        work_input = work_dir / "input"
        work_pass_raw = work_dir / "pass_raw"
        work_output = work_dir / "output"
        for directory in (work_input, work_pass_raw, work_output):
            directory.mkdir(parents=True, exist_ok=True)
        runtime_semi = work_input / "semi_final_output.csv"
        shutil.copy2(semi_file, runtime_semi)
        shutil.copy2(codes_source, work_input / "codes_dump.json")
        for pass_file in pass_files:
            shutil.copy2(pass_file, work_pass_raw / pass_file.name)
    except OSError as exc:
        return False, f"Could not prepare the per-run {FINAL_EXEMPT_INTERMEDIATE_FOLDER} folder: {exc}"

    env = os.environ.copy()
    env.update(
        {
            "PYTHONUNBUFFERED": "1",
            "EXEMPT_FINAL_WORK_DIR": str(work_dir.resolve()),
            "EXEMPT_FINAL_INPUT_DIR": str(work_input.resolve()),
            "EXEMPT_FINAL_OUTPUT_DIR": str(work_output.resolve()),
            "EXEMPT_FINAL_SEMI_CSV_NAME": runtime_semi.name,
            "EXEMPT_FINAL_OUTPUT_FILE_NAME": output_file_name,
            "EXEMPT_FINAL_PLAZA": plaza_name,
            "EXEMPT_FINAL_CODES_FILE": str((work_input / "codes_dump.json").resolve()),
            "EXEMPT_FINAL_PASS_INPUT": str(work_pass_raw.resolve()),
            "EXEMPT_FINAL_PASS_OUTPUT": str(work_input.resolve()),
        }
    )

    ok, error = _run_final_exempt_script(
        script_dir,
        "Pass_file_seperator.py",
        "Step 1: separating MP and LT pass files.",
        env,
        expected_files=(work_input / "MP_Pass.xlsx", work_input / "LT_Pass.xlsx"),
    )
    if not ok:
        return False, error

    summary_path = work_input / "pass_separator_summary.json"
    try:
        pass_summary = json.loads(summary_path.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return False, "Pass separation did not produce its validation summary."
    if not pass_summary.get("pass_type_detected"):
        return False, "No Pass Type column was found in the uploaded pass files. Use a supported pass report and try again."
    if not pass_summary.get("mp_rows"):
        return False, "Pass separation found 0 MP rows. MP passes are required for this process."
    has_lt = bool(pass_summary.get("lt_rows"))
    env["EXEMPT_FINAL_RETURN_INPUT"] = "LT_output.xlsx" if has_lt else "MP Output.xlsx"

    stage_plan = [
        (
            "main_script.py",
            "Step 2: detecting exempt exceptions.",
            (work_output / "updated_new_Combined_Side_Results.xlsx",),
        ),
        (
            "RF3_condition.py",
            "Step 3: calculating RF3 conditions.",
            (
                work_output / "RF3 Pivot.xlsx",
                work_output / "rf3_1.xlsx",
                work_output / "rf3_2.xlsx",
            ),
        ),
        (
            "combined_with_rf3.py",
            "Step 4: combining exceptions and RF3 results.",
            (work_output / "combined_with_rf3.xlsx",),
        ),
        (
            "Date_validity_check.py",
            "Step 5: checking MP pass validity.",
            (work_output / "MP Output.xlsx",),
        ),
    ]
    if has_lt:
        stage_plan.append(
            (
                "Date_validity_check_LT.py",
                "Step 6: checking LT pass validity.",
                (work_output / "LT_output.xlsx",),
            )
        )
    stage_plan.append(
        (
            "Return_Journey_Logic.py",
            "Final step: applying return-journey logic.",
            (work_output / output_file_name,),
        )
    )

    for script_name, stage_name, expected_files in stage_plan:
        ok, error = _run_final_exempt_script(
            script_dir, script_name, stage_name, env, expected_files=expected_files
        )
        if not ok:
            return False, error

    final_workbook = work_output / output_file_name
    final_destination = output_dir / output_file_name
    try:
        shutil.copy2(final_workbook, final_destination)
    except OSError as exc:
        return False, f"Final workbook was created but could not be copied to output: {exc}"

    other_count = int(pass_summary.get("other_rows") or 0)
    other_note = (
        f" {other_count} Other pass row(s) kept in "
        f"{FINAL_EXEMPT_INTERMEDIATE_FOLDER}/input for debugging."
        if other_count
        else ""
    )
    return (
        True,
        f"{FINAL_EXEMPT_GROUP_LABEL} finished for '{process_dir.name}'. "
        f"Output: output/{output_file_name}." + other_note,
    )


def run_final_exempt_in_background(process_name):
    with final_exempt_state_lock:
        final_exempt_state["running"] = True
        final_exempt_state["process_type"] = FINAL_EXEMPT_SUBPROCESS_FOLDER
        final_exempt_state["process_name"] = process_name
        final_exempt_state["started_at"] = time.time()
        final_exempt_state["finished_at"] = None
        final_exempt_state["last_status"] = "Preparing Final Exempt Process."

    ok, process_message = process_final_exempt_files(process_name)

    with final_exempt_state_lock:
        final_exempt_state["running"] = False
        final_exempt_state["finished_at"] = time.time()
        final_exempt_state["last_status"] = process_message if ok else f"FAILED: {process_message}"


@app.route("/life-cycle-merge-status")
def life_cycle_merge_status():
    with life_cycle_state_lock:
        state = dict(life_cycle_state)
    if state["started_at"]:
        state["elapsed_seconds"] = int(time.time() - state["started_at"])
    else:
        state["elapsed_seconds"] = 0
    return jsonify(state)


@app.route("/merge-normalize-status")
def merge_normalize_status():
    with merge_normalize_state_lock:
        state = dict(merge_normalize_state)
    if state["started_at"]:
        state["elapsed_seconds"] = int(time.time() - state["started_at"])
    else:
        state["elapsed_seconds"] = 0
    state["file_kind_label"] = merge_normalize_file_kind_label(state.get("file_kind"))
    return jsonify(state)


@app.route("/merge-remove-dup-status")
def merge_remove_dup_status():
    with merge_remove_dup_state_lock:
        state = dict(merge_remove_dup_state)
    if state["started_at"]:
        state["elapsed_seconds"] = int(time.time() - state["started_at"])
    else:
        state["elapsed_seconds"] = 0
    return jsonify(state)


@app.route("/final-exempt-status")
def final_exempt_status():
    with final_exempt_state_lock:
        state = dict(final_exempt_state)
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
        "valid_invalid_group_label": VALID_INVALID_GROUP_LABEL,
        "exempt_query_group_label": EXEMPT_QUERY_GROUP_LABEL,
        "merge_valid_lookup_group_label": VALID_INVALID_GROUP_LABEL,
        "valid_invalid_parent_folder": VALID_INVALID_PARENT_FOLDER,
        "merge_valid_lookup_parent_folder": VALID_INVALID_PARENT_FOLDER,
        "merge_normalize_group_label": MERGE_NORMALIZE_GROUP_LABEL,
        "merge_remove_dup_group_label": MERGE_REMOVE_DUP_GROUP_LABEL,
        "final_exempt_group_label": FINAL_EXEMPT_GROUP_LABEL,
        "exempt_query_parent_folder": EXEMPT_QUERY_PARENT_FOLDER,
        "merge_normalize_file_kind_labels": MERGE_NORMALIZE_FILE_KIND_LABELS,
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
        merge_normalize_folder=MERGE_NORMALIZE_SUBPROCESS_FOLDER,
        merge_remove_dup_folder=MERGE_REMOVE_DUP_SUBPROCESS_FOLDER,
        final_exempt_folder=FINAL_EXEMPT_SUBPROCESS_FOLDER,
    )


@app.route("/current-process")
def current_process():
    with life_cycle_state_lock:
        current_life_cycle_state = dict(life_cycle_state)
    with valid_invalid_state_lock:
        current_valid_invalid_state = dict(valid_invalid_state)
    with merge_normalize_state_lock:
        current_merge_normalize_state = dict(merge_normalize_state)
    with merge_remove_dup_state_lock:
        current_merge_remove_dup_state = dict(merge_remove_dup_state)
    with final_exempt_state_lock:
        current_final_exempt_state = dict(final_exempt_state)
    return render_template(
        "current_process.html",
        available_processes=list_available_processes(),
        life_cycle_state=current_life_cycle_state,
        valid_invalid_state=current_valid_invalid_state,
        merge_normalize_state=current_merge_normalize_state,
        merge_remove_dup_state=current_merge_remove_dup_state,
        final_exempt_state=current_final_exempt_state,
    )


def process_valid_invalid_files(process_name, header_mapping=None):
    if is_reserved_process_name(process_name):
        return False, (
            f"Process name '{process_name}' is reserved. "
            "Choose a different name (not Life_Cycle_Merge, Valid_Invalid_Lookup, input, output, or rate)."
        )
    process_dir, input_dir, rate_dir = ensure_valid_invalid_process_dirs(process_name)
    if not process_dir or not input_dir or not rate_dir:
        return False, "Please provide a valid process name."
    # Ensure folder structure exists, then validate required files explicitly.
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
    with merge_normalize_state_lock:
        merge_normalize_running = merge_normalize_state["running"]
        merge_normalize_process_name = merge_normalize_state["process_name"]
        merge_normalize_file_kind = merge_normalize_state["file_kind"]
    with merge_remove_dup_state_lock:
        merge_remove_dup_running = merge_remove_dup_state["running"]
        merge_remove_dup_process_name = merge_remove_dup_state["process_name"]
    with final_exempt_state_lock:
        final_exempt_running = final_exempt_state["running"]
        final_exempt_process_name = final_exempt_state["process_name"]

    delete_disabled = (
        life_cycle_running
        or valid_invalid_running
        or merge_normalize_running
        or merge_remove_dup_running
        or final_exempt_running
    )
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
    elif merge_normalize_running:
        delete_disabled_message = (
            f"Delete is disabled while {MERGE_NORMALIZE_GROUP_LABEL} "
            f"({merge_normalize_file_kind_label(merge_normalize_file_kind)}) is running for '{merge_normalize_process_name}'."
        )
    elif merge_remove_dup_running:
        delete_disabled_message = (
            f"Delete is disabled while {MERGE_REMOVE_DUP_GROUP_LABEL} "
            f"is running for '{merge_remove_dup_process_name}'."
        )
    elif final_exempt_running:
        delete_disabled_message = (
            f"Delete is disabled while {FINAL_EXEMPT_GROUP_LABEL} "
            f"is running for '{final_exempt_process_name}'."
        )

    def _annotate(entries):
        return annotate_process_button_state(
            entries,
            life_cycle_running,
            life_cycle_process_name,
            valid_invalid_running,
            valid_invalid_process_name,
            merge_normalize_running,
            merge_normalize_process_name,
            merge_normalize_file_kind,
            merge_remove_dup_running,
            merge_remove_dup_process_name,
            final_exempt_running,
            final_exempt_process_name,
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
                    entries=_annotate(entries),
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
                with merge_normalize_state_lock:
                    mn_running = merge_normalize_state["running"]
                    mn_file_kind = merge_normalize_state["file_kind"]
                    mn_process_name = merge_normalize_state["process_name"]
                with merge_remove_dup_state_lock:
                    mrd_running = merge_remove_dup_state["running"]
                    mrd_process_name = merge_remove_dup_state["process_name"]
                with final_exempt_state_lock:
                    fe_running = final_exempt_state["running"]
                    fe_process_name = final_exempt_state["process_name"]

                running_process_dir = None
                if running and running_process_name:
                    running_process_dir, _ = get_valid_invalid_process_paths(
                        running_process_type,
                        running_process_name,
                    )
                    if running_process_dir:
                        running_process_dir = running_process_dir.resolve()
                mn_process_dir = (
                    get_merge_normalize_paths(mn_file_kind, mn_process_name)[0]
                    if mn_running and mn_process_name and mn_file_kind
                    else None
                )
                if mn_process_dir:
                    mn_process_dir = mn_process_dir.resolve()
                mrd_process_dir = (
                    get_merge_remove_dup_paths(mrd_process_name)[0]
                    if mrd_running and mrd_process_name
                    else None
                )
                if mrd_process_dir:
                    mrd_process_dir = mrd_process_dir.resolve()
                fe_process_dir = (
                    get_final_exempt_paths(fe_process_name)[0]
                    if fe_running and fe_process_name
                    else None
                )
                if fe_process_dir:
                    fe_process_dir = fe_process_dir.resolve()

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
                elif (
                    mn_process_dir
                    and (target_path == mn_process_dir or mn_process_dir in target_path.parents)
                ):
                    messages.append(
                        (
                            "error",
                            f"Cannot delete '{mn_process_name}' while {MERGE_NORMALIZE_GROUP_LABEL} is running.",
                        )
                    )
                elif (
                    mrd_process_dir
                    and (target_path == mrd_process_dir or mrd_process_dir in target_path.parents)
                ):
                    messages.append(
                        (
                            "error",
                            f"Cannot delete '{mrd_process_name}' while {MERGE_REMOVE_DUP_GROUP_LABEL} is running.",
                        )
                    )
                elif (
                    fe_process_dir
                    and (target_path == fe_process_dir or fe_process_dir in target_path.parents)
                ):
                    messages.append(
                        (
                            "error",
                            f"Cannot delete '{fe_process_name}' while {FINAL_EXEMPT_GROUP_LABEL} is running.",
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
                context = parse_input_folder_context(target_path)
                process_name = (context or {}).get("process_name") or ""
                parent_folder = (context or {}).get("parent_folder") or ""
                process_type = (context or {}).get("process_type") or ""
                file_kind = (context or {}).get("file_kind") or ""
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
                elif (
                    parent_folder == EXEMPT_QUERY_PARENT_FOLDER
                    and process_type == MERGE_NORMALIZE_SUBPROCESS_FOLDER
                ):
                    keywords, kw_error = fetch_lc_etc_header_keywords_or_error()
                    if kw_error:
                        messages.append(("error", kw_error))
                    else:
                        with merge_normalize_state_lock:
                            running = merge_normalize_state["running"]
                        if running:
                            messages.append(
                                (
                                    "error",
                                    f"A {MERGE_NORMALIZE_GROUP_LABEL} process is already running. Please wait for it to finish.",
                                )
                            )
                        elif not file_kind:
                            messages.append(("error", "Could not determine LC/ETC or VRN file type for this folder."))
                        else:
                            worker = threading.Thread(
                                target=run_merge_normalize_in_background,
                                args=(file_kind, process_name),
                                daemon=True,
                            )
                            worker.start()
                            messages.append(
                                (
                                    "success",
                                    process_started_success_message(
                                        MERGE_NORMALIZE_GROUP_LABEL,
                                        process_name,
                                        merge_normalize_file_kind_label(file_kind),
                                    ),
                                )
                            )
                elif (
                    parent_folder == EXEMPT_QUERY_PARENT_FOLDER
                    and process_type == MERGE_REMOVE_DUP_SUBPROCESS_FOLDER
                ):
                    with merge_remove_dup_state_lock:
                        running = merge_remove_dup_state["running"]
                    if running:
                        messages.append(
                            (
                                "error",
                                f"A {MERGE_REMOVE_DUP_GROUP_LABEL} process is already running. Please wait for it to finish.",
                            )
                        )
                    else:
                        worker = threading.Thread(
                            target=run_merge_remove_dup_in_background,
                            args=(process_name,),
                            daemon=True,
                        )
                        worker.start()
                        messages.append(
                            (
                                "success",
                                process_started_success_message(
                                    MERGE_REMOVE_DUP_GROUP_LABEL, process_name
                                ),
                            )
                        )
                elif (
                    parent_folder == EXEMPT_QUERY_PARENT_FOLDER
                    and process_type == FINAL_EXEMPT_SUBPROCESS_FOLDER
                ):
                    with final_exempt_state_lock:
                        running = final_exempt_state["running"]
                    if running:
                        messages.append(
                            (
                                "error",
                                f"A {FINAL_EXEMPT_GROUP_LABEL} is already running. Please wait for it to finish.",
                            )
                        )
                    else:
                        worker = threading.Thread(
                            target=run_final_exempt_in_background,
                            args=(process_name,),
                            daemon=True,
                        )
                        worker.start()
                        messages.append(
                            (
                                "success",
                                process_started_success_message(
                                    FINAL_EXEMPT_GROUP_LABEL, process_name
                                ),
                            )
                        )
                elif parent_folder == EXEMPT_QUERY_PARENT_FOLDER:
                    messages.append(
                        (
                            "error",
                            f"Exempt Query processing from Directories is only available for "
                            f"{MERGE_NORMALIZE_GROUP_LABEL}, {MERGE_REMOVE_DUP_GROUP_LABEL}, and {FINAL_EXEMPT_GROUP_LABEL}.",
                        )
                    )
                elif not is_valid_invalid_parent_folder(parent_folder):
                    messages.append(("error", "Unsupported process group for this input folder."))
                elif process_type == LIFE_CYCLE_PROCESS_FOLDER:
                    keywords, kw_error = fetch_lc_etc_header_keywords_or_error()
                    if kw_error:
                        messages.append(("error", kw_error))
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
                                args=(process_name,),
                                daemon=True,
                            )
                            worker.start()
                            messages.append(
                                (
                                    "success",
                                    process_started_success_message(
                                        "Life cycle merge", process_name
                                    ),
                                )
                            )
                elif process_type == VALID_INVALID_PROCESS_FOLDER:
                    header_mapping = load_vil_header_mapping_from_disk(process_name)
                    if not is_vil_header_mapping_complete(header_mapping):
                        messages.append(
                            (
                                "error",
                                "Header mapping is incomplete. Open Valid/Invalid Lookup, "
                                "map headers, and confirm before processing from Directories.",
                            )
                        )
                    else:
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
                                args=(process_name, header_mapping),
                                daemon=True,
                            )
                            worker.start()
                            messages.append(
                                (
                                    "success",
                                    process_started_success_message(
                                        "Valid/Invalid Lookup", process_name
                                    ),
                                )
                            )
                else:
                    messages.append(("error", "Unsupported process type for this input folder."))

    current_path, entries = list_directory_entries(relative_path)
    if current_path is None:
        current_path, entries = list_directory_entries("")
        messages.append(("error", "Invalid directory path selected."))

    entries = _annotate(entries)

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


@app.route("/merge-normalize", methods=["GET", "POST"])
def merge_normalize():
    messages = []
    reset_form_after_submit = False
    current_process_name = request.form.get("process_name", "").strip() if request.method == "POST" else ""
    current_file_kind = (
        normalize_merge_normalize_file_kind(request.form.get("file_kind", ""))
        if request.method == "POST"
        else ""
    )

    if request.method == "POST":
        action = request.form.get("action")
        if current_process_name and is_reserved_process_name(current_process_name):
            messages.append(
                (
                    "error",
                    f"Process name '{current_process_name}' is reserved. "
                    "Use a different name (not Merge_&_Normalize, LC/ETC, VRN, input, or output).",
                )
            )
            process_dir, input_dir, output_dir = (None, None, None)
        else:
            process_dir, input_dir, output_dir = get_merge_normalize_paths(
                current_file_kind, current_process_name
            )

        if action == "upload":
            if not current_file_kind:
                messages.append(("error", "Select whether files are VRN or LC/ETC before upload."))
            elif not process_dir or not input_dir:
                messages.append(("error", "Process name is required before upload."))
            else:
                input_dir.mkdir(parents=True, exist_ok=True)
                (output_dir or process_dir / "output").mkdir(parents=True, exist_ok=True)
                uploaded_files = request.files.getlist("files")
                if not uploaded_files or all(not file.filename for file in uploaded_files):
                    messages.append(("error", "Please select at least one file to upload."))
                else:
                    existing_files = {
                        name.lower()
                        for name in list_merge_normalize_uploaded_files(
                            current_file_kind, current_process_name
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

            if not current_file_kind:
                messages.append(("error", "Select whether files are VRN or LC/ETC before delete."))
            elif not process_dir or not input_dir:
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
            elif not current_file_kind:
                messages.append(("error", "Select whether files are VRN or LC/ETC before submit."))
            elif is_reserved_process_name(current_process_name):
                messages.append(
                    (
                        "error",
                        f"Process name '{current_process_name}' is reserved. "
                        "Use a different name (not Merge_&_Normalize, LC/ETC, VRN, input, or output).",
                    )
                )
            else:
                process_dir, input_dir, output_dir = ensure_merge_normalize_process_dirs(
                    current_file_kind, current_process_name
                )
                uploaded_count = 0
                if input_dir and input_dir.exists():
                    uploaded_count = sum(1 for p in input_dir.iterdir() if p.is_file())
                if uploaded_count < MIN_MERGE_NORMALIZE_FILES:
                    messages.append(
                        (
                            "error",
                            f"Upload at least {MIN_MERGE_NORMALIZE_FILES} file(s) before starting.",
                        )
                    )
                else:
                    with merge_normalize_state_lock:
                        running = merge_normalize_state["running"]
                    if running:
                        messages.append(
                            (
                                "error",
                                f"A {MERGE_NORMALIZE_GROUP_LABEL} process is already running. Please wait for it to finish.",
                            )
                        )
                    else:
                        worker = threading.Thread(
                            target=run_merge_normalize_in_background,
                            args=(current_file_kind, current_process_name),
                            daemon=True,
                        )
                        worker.start()
                        messages.append(
                            (
                                "success",
                                process_started_success_message(
                                    MERGE_NORMALIZE_GROUP_LABEL,
                                    current_process_name,
                                    merge_normalize_file_kind_label(current_file_kind),
                                ),
                            )
                        )
                        reset_form_after_submit = True
                        current_process_name = ""
                        current_file_kind = ""

    uploaded_files = (
        list_merge_normalize_uploaded_files(current_file_kind, current_process_name)
        if current_process_name and current_file_kind
        else []
    )

    return render_template(
        "merge_normalize.html",
        current_process_name=current_process_name,
        current_file_kind=current_file_kind,
        uploaded_files=uploaded_files,
        messages=messages,
        allowed_extensions=sorted(ALLOWED_EXTENSIONS),
        reset_form_after_submit=reset_form_after_submit,
        min_merge_normalize_files=MIN_MERGE_NORMALIZE_FILES,
        min_header_keywords=MIN_LC_ETC_HEADER_KEYWORDS,
        file_kinds=MERGE_NORMALIZE_FILE_KINDS,
        file_kind_labels=MERGE_NORMALIZE_FILE_KIND_LABELS,
        header_keywords_modal_title=f"Header Keywords ({HEADER_KEYWORDS_LC_ETC_VRN_LABEL})",
        header_keywords_modal_description=(
            "Shared keyword list for LC/ETC and VRN header detection. "
            f"At least {MIN_LC_ETC_HEADER_KEYWORDS} keywords are required."
        ),
        merge_normalize_folder=MERGE_NORMALIZE_SUBPROCESS_FOLDER,
    )


@app.route("/merge-remove-duplicate", methods=["GET", "POST"])
def merge_remove_duplicate():
    messages = []
    reset_form_after_submit = False
    current_process_name = request.form.get("process_name", "").strip() if request.method == "POST" else ""

    vrn_outputs = list_merge_normalize_output_files(MERGE_NORMALIZE_VRN_FOLDER)
    lc_outputs = list_merge_normalize_output_files(MERGE_NORMALIZE_LC_ETC_FOLDER)

    if request.method == "POST":
        action = request.form.get("action")

        if current_process_name and is_reserved_process_name(current_process_name):
            messages.append(
                (
                    "error",
                    f"Process name '{current_process_name}' is reserved. "
                    "Use a different name (not Merge_&_Remove_Duplicate, vrn, lc_etc, input, or output).",
                )
            )
        elif action == "upload":
            if not current_process_name:
                messages.append(("error", "Process name is required before upload."))
            else:
                ensure_merge_remove_dup_dirs(current_process_name)
                vrn_source = request.form.get("vrn_source", "upload").strip()
                lc_source = request.form.get("lc_source", "upload").strip()

                if vrn_source == "import":
                    vrn_rel = request.form.get("vrn_import", "").strip()
                    vrn_path = resolve_merge_normalize_output_import(
                        vrn_rel, MERGE_NORMALIZE_VRN_FOLDER
                    )
                    if not vrn_path:
                        messages.append(("error", "Select a valid VRN file from Merge + Normalize output."))
                    else:
                        copy_to_merge_remove_dup_slot(
                            current_process_name, MERGE_REMOVE_DUP_VRN_SLOT, vrn_path
                        )
                        messages.append(("success", f"Imported VRN file from {vrn_path.name}."))
                else:
                    vrn_upload = request.files.get("vrn_file")
                    if vrn_upload and vrn_upload.filename:
                        original_name = secure_filename(vrn_upload.filename)
                        if not allowed_merge_remove_dup_file(original_name):
                            messages.append(("error", "VRN file must be a .csv."))
                        else:
                            _, vrn_dir, _, _ = ensure_merge_remove_dup_dirs(current_process_name)
                            clear_merge_remove_dup_slot(current_process_name, MERGE_REMOVE_DUP_VRN_SLOT)
                            vrn_upload.save(vrn_dir / original_name)
                            messages.append(("success", f"Uploaded VRN file: {original_name}"))

                if lc_source == "import":
                    lc_rel = request.form.get("lc_import", "").strip()
                    lc_path = resolve_merge_normalize_output_import(
                        lc_rel, MERGE_NORMALIZE_LC_ETC_FOLDER
                    )
                    if not lc_path:
                        messages.append(("error", "Select a valid LC/ETC file from Merge + Normalize output."))
                    else:
                        copy_to_merge_remove_dup_slot(
                            current_process_name, MERGE_REMOVE_DUP_LC_SLOT, lc_path
                        )
                        messages.append(("success", f"Imported LC/ETC file from {lc_path.name}."))
                else:
                    lc_upload = request.files.get("lc_file")
                    if lc_upload and lc_upload.filename:
                        original_name = secure_filename(lc_upload.filename)
                        if not allowed_merge_remove_dup_file(original_name):
                            messages.append(("error", "LC/ETC file must be a .csv."))
                        else:
                            _, _, lc_dir, _ = ensure_merge_remove_dup_dirs(current_process_name)
                            clear_merge_remove_dup_slot(current_process_name, MERGE_REMOVE_DUP_LC_SLOT)
                            lc_upload.save(lc_dir / original_name)
                            messages.append(("success", f"Uploaded LC/ETC file: {original_name}"))

        elif action == "delete":
            slot = request.form.get("slot", "").strip()
            filename = secure_filename(request.form.get("filename", ""))
            if not current_process_name:
                messages.append(("error", "Process name is required before delete."))
            elif slot not in (MERGE_REMOVE_DUP_VRN_SLOT, MERGE_REMOVE_DUP_LC_SLOT):
                messages.append(("error", "Invalid file slot."))
            elif not filename:
                messages.append(("error", "Invalid filename."))
            else:
                _, vrn_dir, lc_dir, _ = get_merge_remove_dup_paths(current_process_name)
                slot_dir = vrn_dir if slot == MERGE_REMOVE_DUP_VRN_SLOT else lc_dir
                file_path = slot_dir / filename if slot_dir else None
                if file_path and file_path.exists() and file_path.is_file():
                    try:
                        file_path.unlink()
                        messages.append(("success", f"Deleted file: {filename}"))
                    except OSError as exc:
                        messages.append(("error", f"Delete failed: {exc}"))
                else:
                    messages.append(("error", f"File not found: {filename}"))

        elif action == "process":
            if not current_process_name:
                messages.append(("error", "Process name is required before final submit."))
            elif is_reserved_process_name(current_process_name):
                messages.append(
                    (
                        "error",
                        f"Process name '{current_process_name}' is reserved. "
                        "Use a different name (not Merge_&_Remove_Duplicate, vrn, lc_etc, input, or output).",
                    )
                )
            elif not get_merge_remove_dup_slot_file(current_process_name, MERGE_REMOVE_DUP_VRN_SLOT):
                messages.append(
                    ("error", "VRN file is required. Upload or import from Merge + Normalize (VRN) output.")
                )
            elif not get_merge_remove_dup_slot_file(current_process_name, MERGE_REMOVE_DUP_LC_SLOT):
                messages.append(
                    ("error", "LC/ETC file is required. Upload or import from Merge + Normalize (LC/ETC) output.")
                )
            else:
                with merge_remove_dup_state_lock:
                    running = merge_remove_dup_state["running"]
                if running:
                    messages.append(
                        (
                            "error",
                            f"A {MERGE_REMOVE_DUP_GROUP_LABEL} process is already running. Please wait for it to finish.",
                        )
                    )
                else:
                    worker = threading.Thread(
                        target=run_merge_remove_dup_in_background,
                        args=(current_process_name,),
                        daemon=True,
                    )
                    worker.start()
                    messages.append(
                        (
                            "success",
                            process_started_success_message(
                                MERGE_REMOVE_DUP_GROUP_LABEL, current_process_name
                            ),
                        )
                    )
                    reset_form_after_submit = True
                    current_process_name = ""

    vrn_files = list_merge_remove_dup_slot_files(current_process_name) if current_process_name else []
    lc_files = (
        list_merge_remove_dup_slot_files(current_process_name, MERGE_REMOVE_DUP_LC_SLOT)
        if current_process_name
        else []
    )
    submit_ready = bool(current_process_name and vrn_files and lc_files)

    return render_template(
        "merge_remove_duplicate.html",
        current_process_name=current_process_name,
        vrn_files=vrn_files,
        lc_files=lc_files,
        vrn_outputs=vrn_outputs,
        lc_outputs=lc_outputs,
        messages=messages,
        reset_form_after_submit=reset_form_after_submit,
        submit_ready=submit_ready,
        semi_final_output_filename=SEMI_FINAL_OUTPUT_FILENAME,
        merge_normalize_output_filename=MERGE_NORMALIZE_OUTPUT_FILENAME,
    )


@app.route("/final-exempt-process", methods=["GET", "POST"])
def final_exempt_process():
    messages = []
    reset_form_after_submit = False
    current_process_name = (
        request.form.get("process_name", "").strip() if request.method == "POST" else ""
    )
    semi_outputs = list_merge_remove_dup_output_files()

    if request.method == "POST":
        action = request.form.get("action", "").strip()
        if current_process_name and is_reserved_process_name(current_process_name):
            messages.append(("error", "Choose a valid process name (not a reserved folder name)."))
        elif action == "apply":
            plaza_name = request.form.get("plaza_name", "").strip()
            output_file_name = final_exempt_output_file_name(
                current_process_name, request.form.get("output_file_name", "")
            )
            if not current_process_name:
                messages.append(("error", "Process name is required before applying inputs."))
            elif plaza_name not in get_final_exempt_plazas():
                messages.append(("error", "Select a plaza from the available list before applying inputs."))
            else:
                _, semi_dir, pass_dir, _, _ = ensure_final_exempt_dirs(current_process_name)
                save_final_exempt_config(current_process_name, plaza_name, output_file_name)
                messages.append(("success", "Process settings saved."))

                semi_source = request.form.get("semi_source", "upload").strip()
                if semi_source == "import":
                    imported_semi = resolve_merge_remove_dup_output_import(
                        request.form.get("semi_import", "").strip()
                    )
                    if not imported_semi:
                        messages.append(
                            (
                                "error",
                                "Select a valid semi-final output from Merge + Remove Duplicate.",
                            )
                        )
                    else:
                        for existing in semi_dir.iterdir():
                            if existing.is_file():
                                existing.unlink()
                        shutil.copy2(imported_semi, semi_dir / SEMI_FINAL_OUTPUT_FILENAME)
                        messages.append(("success", f"Imported semi-final file from {imported_semi.parent.parent.name}."))
                else:
                    semi_upload = request.files.get("semi_file")
                    if semi_upload and semi_upload.filename:
                        filename = secure_filename(semi_upload.filename)
                        if Path(filename).suffix.lower() != ".csv":
                            messages.append(("error", "Semi-final file must be a .csv."))
                        else:
                            for existing in semi_dir.iterdir():
                                if existing.is_file():
                                    existing.unlink()
                            semi_upload.save(semi_dir / filename)
                            messages.append(("success", f"Uploaded semi-final file: {filename}"))

                uploaded_pass_count = 0
                for pass_upload in request.files.getlist("pass_files"):
                    if not pass_upload or not pass_upload.filename:
                        continue
                    filename = secure_filename(pass_upload.filename)
                    if not allowed_final_exempt_pass_file(filename):
                        messages.append(("error", f"Pass file must be .xls or .xlsx: {filename or 'unnamed file'}"))
                        continue
                    destination = pass_dir / filename
                    suffix = 2
                    while destination.exists():
                        destination = pass_dir / f"{Path(filename).stem}_{suffix}{Path(filename).suffix}"
                        suffix += 1
                    pass_upload.save(destination)
                    uploaded_pass_count += 1
                if uploaded_pass_count:
                    messages.append(("success", f"Uploaded {uploaded_pass_count} pass file(s)."))

        elif action == "delete":
            slot = request.form.get("slot", "").strip()
            filename = secure_filename(request.form.get("filename", ""))
            if not current_process_name or not filename:
                messages.append(("error", "A process name and filename are required before delete."))
            else:
                _, semi_dir, pass_dir, _, _ = get_final_exempt_paths(current_process_name)
                folder = semi_dir if slot == FINAL_EXEMPT_SEMI_SLOT else pass_dir if slot == FINAL_EXEMPT_PASS_SLOT else None
                target = folder / filename if folder else None
                if target and target.is_file():
                    try:
                        target.unlink()
                        messages.append(("success", f"Deleted file: {filename}"))
                    except OSError as exc:
                        messages.append(("error", f"Delete failed: {exc}"))
                else:
                    messages.append(("error", "Selected input file was not found."))

        elif action == "process":
            process_config = load_final_exempt_config(current_process_name)
            semi_file = get_final_exempt_semi_file(current_process_name)
            pass_files = list_final_exempt_pass_files(current_process_name)
            if not current_process_name:
                messages.append(("error", "Process name is required before starting."))
            elif not process_config.get("plaza_name"):
                messages.append(("error", "Apply the process settings, including plaza, before starting."))
            elif not semi_file:
                messages.append(("error", "Attach a semi-final CSV before starting."))
            elif not pass_files:
                messages.append(("error", "Attach at least one pass file before starting."))
            else:
                with final_exempt_state_lock:
                    running = final_exempt_state["running"]
                if running:
                    messages.append(
                        (
                            "error",
                            f"A {FINAL_EXEMPT_GROUP_LABEL} is already running. Please wait for it to finish.",
                        )
                    )
                else:
                    worker = threading.Thread(
                        target=run_final_exempt_in_background,
                        args=(current_process_name,),
                        daemon=True,
                    )
                    worker.start()
                    messages.append(
                        (
                            "success",
                            process_started_success_message(
                                FINAL_EXEMPT_GROUP_LABEL, current_process_name
                            ),
                        )
                    )
                    reset_form_after_submit = True
                    current_process_name = ""

    process_config = load_final_exempt_config(current_process_name) if current_process_name else {}
    semi_file = get_final_exempt_semi_file(current_process_name) if current_process_name else None
    pass_files = list_final_exempt_pass_files(current_process_name) if current_process_name else []
    submit_ready = bool(
        current_process_name
        and semi_file
        and pass_files
        and process_config.get("plaza_name")
    )
    return render_template(
        "final_exempt_process.html",
        current_process_name=current_process_name,
        process_config=process_config,
        semi_file=semi_file.name if semi_file else "",
        pass_files=pass_files,
        semi_outputs=semi_outputs,
        plazas=get_final_exempt_plazas(),
        messages=messages,
        submit_ready=submit_ready,
        reset_form_after_submit=reset_form_after_submit,
    )


@app.route("/life-cycle-merge", methods=["GET", "POST"])
def life_cycle_merge():
    messages = []
    reset_form_after_submit = False
    current_process_name = request.form.get("process_name", "").strip() if request.method == "POST" else ""

    if request.method == "POST":
        action = request.form.get("action")
        process_dir, input_dir = get_valid_invalid_process_paths(
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
                process_dir, input_dir = get_valid_invalid_process_paths(
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
                                process_started_success_message(
                                    "Life cycle merge", current_process_name
                                ),
                            )
                        )
                        reset_form_after_submit = True
                        current_process_name = ""

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
        reset_form_after_submit=reset_form_after_submit,
        min_life_cycle_merge_files=MIN_LIFE_CYCLE_MERGE_FILES,
        min_header_keywords=MIN_LC_ETC_HEADER_KEYWORDS,
        header_keywords_modal_title=f"Header Keywords ({LC_ETC_FILE_TYPE})",
        header_keywords_modal_description=(
            "Each keyword is stored as one row in nhit_file_process. "
            f"At least {MIN_LC_ETC_HEADER_KEYWORDS} keywords are required for merge header detection."
        ),
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
        if current_process_name and is_reserved_process_name(current_process_name):
            messages.append(
                (
                    "error",
                    f"Process name '{current_process_name}' is reserved. "
                    "Use a different name (not Life_Cycle_Merge, Valid_Invalid_Lookup, input, output, or rate).",
                )
            )
            process_dir, input_dir, rate_dir = (None, None, None)
        else:
            process_dir, input_dir, rate_dir = get_valid_invalid_paths(current_process_name)

        if action == "upload":
            if not process_dir or not input_dir or not rate_dir:
                if not any(message_type == "error" for message_type, _ in messages):
                    messages.append(("error", "Process name is required before upload."))
            else:
                process_dir, input_dir, rate_dir = ensure_valid_invalid_process_dirs(current_process_name)

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
                            ensure_valid_invalid_process_dirs(current_process_name)
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
            elif is_reserved_process_name(current_process_name):
                messages.append(
                    (
                        "error",
                        f"Process name '{current_process_name}' is reserved. "
                        "Use a different name (not Life_Cycle_Merge, Valid_Invalid_Lookup, input, output, or rate).",
                    )
                )
            else:
                process_dir, input_dir, rate_dir = ensure_valid_invalid_process_dirs(current_process_name)
                if not process_dir or not input_dir or not rate_dir:
                    messages.append(("error", "Invalid process name."))
                else:
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
                            process_started_success_message(
                                "Valid/Invalid Lookup", current_process_name
                            ),
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
    # use_reloader=False: the debug reloader (watchdog) was restarting mid-job when
    # unrelated packages under site-packages changed, killing background Process runs.
    # Restart the portal manually after editing code.
    app.run(host="0.0.0.0", port=5000, debug=True, use_reloader=False)
