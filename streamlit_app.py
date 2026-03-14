import streamlit as st
from pprint import pformat

st.set_page_config(page_title="Aspen Plus Script Generator", layout="wide")

TEMPLATE = r"""# -*- coding: utf-8 -*-
# AUTOMATED ASPEN PLUS - PtMeOH SIMULATIONS
# Generated from a Streamlit app.
#
# DEPENDENCIES:
# pip install numpy pandas openpyxl pywin32 tqdm reportlab

import os
import sys
import math
import time
import zipfile
import shutil
import tempfile
import traceback
from copy import deepcopy
from pathlib import Path
from datetime import datetime
from xml.sax.saxutils import escape

import numpy as np
import pandas as pd

try:
    import win32com.client as win32
except ImportError as e:
    raise ImportError("Missing pywin32. Install it with: pip install pywin32") from e

try:
    from tqdm import tqdm
    TQDM_AVAILABLE = True
except ImportError:
    TQDM_AVAILABLE = False

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

CONFIG = __CONFIG_PYTHON__


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def safe_str(obj):
    try:
        return str(obj)
    except Exception:
        return repr(obj)


def ensure_dir(path_like):
    path = Path(path_like)
    path.mkdir(parents=True, exist_ok=True)
    return path


def is_blank(value):
    return value is None or str(value).strip() == ""


def normalize_user_path(text):
    return str(text).strip().strip('"').strip("'")


def log_message(msg, log_lines, print_also=True):
    line = f"[{now_str()}] {msg}"
    log_lines.append(line)
    if print_also:
        print(line)


def parse_output_paths_from_user(raw_text):
    result = {}
    if is_blank(raw_text):
        return result
    chunks = [c.strip() for c in raw_text.split(";") if c.strip()]
    for chunk in chunks:
        if "=" not in chunk:
            continue
        alias, path = chunk.split("=", 1)
        result[alias.strip()] = path.strip()
    return result


def prompt_if_missing(cfg):
    if is_blank(cfg["ASPEN_FILE"]):
        cfg["ASPEN_FILE"] = normalize_user_path(input("Aspen file path (.bkp, .apw or .apwz): "))
    if is_blank(cfg["INPUT_EXCEL"]):
        cfg["INPUT_EXCEL"] = normalize_user_path(input("Input Excel path: "))
    if is_blank(cfg["OUTPUT_DIR"]):
        cfg["OUTPUT_DIR"] = normalize_user_path(input("Output folder: "))
    if is_blank(cfg["INPUT_VARIABLES"]["H2_mass_flow"]):
        cfg["INPUT_VARIABLES"]["H2_mass_flow"] = normalize_user_path(input("Aspen internal path for H2 mass flow: "))

    print("
If you want to add extra output paths from the keyboard, use:")
    print("alias=\Data\... ; alias2=\Data\...")
    raw_outputs = input("Extra output paths (ENTER to keep CONFIG values): ").strip()
    extra_outputs = parse_output_paths_from_user(raw_outputs)
    if extra_outputs:
        cfg["OUTPUT_VARIABLES"].update(extra_outputs)
    return cfg


def resolve_aspen_model_path(user_file):
    user_path = Path(normalize_user_path(user_file)).expanduser().resolve()
    if not user_path.exists():
        raise FileNotFoundError(f"Aspen file not found: {user_path}")
    ext = user_path.suffix.lower()
    if ext in {".bkp", ".apw"}:
        return {"model_path": str(user_path), "temp_dir": None, "source_ext": ext, "message": f"Aspen file detected: {user_path.name}"}
    if ext == ".apwz":
        if not zipfile.is_zipfile(user_path):
            raise RuntimeError(f"The .apwz file could not be read as a valid container: {user_path}")
        temp_dir = tempfile.mkdtemp(prefix="aspen_apwz_")
        with zipfile.ZipFile(user_path, "r") as zf:
            zf.extractall(temp_dir)
        extracted = Path(temp_dir)
        bkp_files = list(extracted.rglob("*.bkp"))
        apw_files = list(extracted.rglob("*.apw"))
        if bkp_files:
            chosen = bkp_files[0]
            return {"model_path": str(chosen.resolve()), "temp_dir": temp_dir, "source_ext": ext, "message": f"APWZ extracted. Internal BKP will be used: {chosen.name}"}
        if apw_files:
            chosen = apw_files[0]
            return {"model_path": str(chosen.resolve()), "temp_dir": temp_dir, "source_ext": ext, "message": f"APWZ extracted. Internal APW will be used: {chosen.name}"}
        raise RuntimeError("The .apwz file does not contain a usable .bkp or .apw file.")
    raise ValueError(f"Unsupported extension: {ext}. Use .bkp, .apw or .apwz.")


def open_aspen_case(model_path, visible=True):
    aspen = win32.Dispatch("Apwn.Document")
    try:
        aspen.Visible = bool(visible)
    except Exception:
        pass
    try:
        aspen.SuppressDialogs = 1
    except Exception:
        pass
    aspen.InitFromArchive2(str(model_path))
    return aspen


def close_aspen_case(aspen):
    if aspen is None:
        return
    try:
        aspen.Close()
    except Exception:
        pass


def find_node(aspen, path):
    if is_blank(path):
        return None
    try:
        return aspen.Tree.FindNode(path)
    except Exception:
        return None


def set_node_value(aspen, path, value):
    node = find_node(aspen, path)
    if node is None:
        raise KeyError(f"Node not found: {path}")
    node.Value = value


def validate_nodes(aspen, cfg, log_lines):
    log_message("Validating Aspen internal paths...", log_lines)
    h2_path = cfg["INPUT_VARIABLES"].get("H2_mass_flow", "")
    if find_node(aspen, h2_path) is None:
        log_message(f'WARNING: input variable not found -> H2_mass_flow = "{h2_path}"', log_lines)
        raise KeyError(f'Input variable H2_mass_flow not found: "{h2_path}"')
    else:
        log_message(f'Input variable found -> H2_mass_flow = "{h2_path}"', log_lines)

    for alias, path in cfg["OUTPUT_VARIABLES"].items():
        if is_blank(path):
            log_message(f'WARNING: output "{alias}" has no configured path; it will be left as NaN.', log_lines)
            continue
        if find_node(aspen, path) is None:
            log_message(f'WARNING: output not found -> {alias} = "{path}"', log_lines)
        else:
            log_message(f'Output found -> {alias} = "{path}"', log_lines)

    for profile_name in ["CONVERGENCE_BASE", "CONVERGENCE_LEVEL2", "CONVERGENCE_LEVEL3"]:
        profile = cfg[profile_name]
        for key in ["max_iter_path", "tolerance_path", "method_path"]:
            path = profile.get(key, "")
            if is_blank(path):
                log_message(f'WARNING: {profile_name}.{key} was not configured.', log_lines)
                continue
            exists = find_node(aspen, path) is not None
            msg = "found" if exists else "not found"
            log_message(f'Aspen convergence path {profile_name}.{key}: {msg} -> "{path}"', log_lines)


def read_run_status(aspen, cfg):
    for path in cfg["RUN_STATUS_CANDIDATES"]:
        node = find_node(aspen, path)
        if node is not None:
            try:
                return node.Value, path
            except Exception:
                return None, path
    return None, None


def status_is_success(status_value, cfg):
    if status_value is None:
        return True
    status_text = safe_str(status_value).strip().lower()
    normalized_success = {safe_str(v).strip().lower() for v in cfg["SUCCESS_STATUS_VALUES"]}
    return status_text in normalized_success


def run_aspen_once(aspen, cfg):
    error_text = None
    tb_text = None
    try:
        aspen.Engine.Run2()
    except Exception as e:
        error_text = safe_str(e)
        tb_text = traceback.format_exc()
    status_value, status_path = read_run_status(aspen, cfg)
    converged = (error_text is None) and status_is_success(status_value, cfg)
    return {"converged": converged, "status_value": status_value, "status_path": status_path, "error_text": error_text, "traceback": tb_text}


def apply_convergence_profile(aspen, profile):
    actions = []
    for label_path, label_value, tag in [
        (profile.get("max_iter_path", ""), profile.get("max_iter_value", None), "max_iter"),
        (profile.get("tolerance_path", ""), profile.get("tolerance_value", None), "tolerance"),
        (profile.get("method_path", ""), profile.get("method_value", None), "method"),
    ]:
        if is_blank(label_path):
            continue
        node = find_node(aspen, label_path)
        if node is not None:
            try:
                node.Value = label_value
                actions.append(f"{tag}={label_value}")
            except Exception as e:
                actions.append(f"could not apply {tag} ({safe_str(e)})")
        else:
            actions.append(f"{tag} path not found")
    if not actions:
        actions.append("no convergence changes configured")
    return actions


def reset_convergence_to_base(aspen, cfg, log_lines=None):
    actions = apply_convergence_profile(aspen, cfg["CONVERGENCE_BASE"])
    if log_lines is not None:
        log_message(f"Convergence parameters reset to BASE -> {', '.join(actions)}", log_lines)


def try_warm_start_with_last_success(aspen, cfg, last_success_h2):
    if last_success_h2 is None:
        return {"attempted": False, "message": "No previous converged case was available for warm start."}
    h2_path = cfg["INPUT_VARIABLES"]["H2_mass_flow"]
    try:
        set_node_value(aspen, h2_path, last_success_h2)
        warm_result = run_aspen_once(aspen, cfg)
        if warm_result["converged"]:
            return {"attempted": True, "message": f"Warm start executed with last converged H2 = {last_success_h2}"}
        return {"attempted": True, "message": "Warm start was attempted but did not converge."}
    except Exception as e:
        return {"attempted": True, "message": f"Warm start could not be executed: {safe_str(e)}"}


def read_outputs(aspen, output_variables):
    data, comments = {}, []
    for alias, path in output_variables.items():
        if is_blank(path):
            data[alias] = np.nan
            comments.append(f"{alias}: empty path")
            continue
        node = find_node(aspen, path)
        if node is None:
            data[alias] = np.nan
            comments.append(f"{alias}: path not found")
            continue
        try:
            data[alias] = node.Value
        except Exception as e:
            data[alias] = np.nan
            comments.append(f"{alias}: read error ({safe_str(e)})")
    return data, comments


def save_excel_report(results_df, log_df, config_df, output_excel):
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        results_df.to_excel(writer, sheet_name="Results", index=False)
        log_df.to_excel(writer, sheet_name="Execution_Log", index=False)
        config_df.to_excel(writer, sheet_name="Config", index=False)


def save_txt_log(log_lines, txt_path):
    with open(txt_path, "w", encoding="utf-8") as f:
        for line in log_lines:
            f.write(line + "
")


def save_pdf_report(pdf_path, log_lines, results_df, cfg):
    if not REPORTLAB_AVAILABLE:
        raise ImportError("Missing reportlab. Install it with: pip install reportlab")
    doc = SimpleDocTemplate(str(pdf_path), pagesize=A4)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="SmallMono", fontName="Helvetica", fontSize=8, leading=10))
    story = [Paragraph("Aspen Plus execution report - PtMeOH", styles["Title"]), Spacer(1, 10)]
    story.append(Paragraph(f"Generated on: {escape(now_str())}", styles["Normal"]))
    story.append(Paragraph(f"Aspen file: {escape(str(cfg['ASPEN_FILE']))}", styles["Normal"]))
    story.append(Paragraph(f"Input Excel: {escape(str(cfg['INPUT_EXCEL']))}", styles["Normal"]))
    story.append(Spacer(1, 12))
    total_runs = len(results_df)
    feasible_runs = int(results_df["feasible"].fillna(0).sum()) if "feasible" in results_df.columns else 0
    failed_runs = total_runs - feasible_runs
    summary_data = [["Metric", "Value"], ["Total simulations", str(total_runs)], ["Successful convergence", str(feasible_runs)], ["Did not converge", str(failed_runs)], ["H2 column used", str(cfg["H2_COLUMN"])], ["H2 variable path", str(cfg["INPUT_VARIABLES"]["H2_mass_flow"])] ]
    table = Table(summary_data)
    table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9EAF7")), ("GRID", (0, 0), (-1, -1), 0.4, colors.grey)]))
    story.append(table)
    story.append(PageBreak())
    story.append(Paragraph("Log message sequence", styles["Heading2"]))
    for line in log_lines:
        story.append(Paragraph(escape(line), styles["SmallMono"]))
    doc.build(story)


def progress_wrapper(iterable, total, desc="Overall progress"):
    if TQDM_AVAILABLE:
        return tqdm(iterable, total=total, desc=desc, unit="sim")
    return iterable


def reopen_base_case_if_requested(cfg, current_aspen, model_info, log_lines):
    if not cfg["REOPEN_BASE_EACH_DAY"]:
        return current_aspen
    close_aspen_case(current_aspen)
    aspen = open_aspen_case(model_info["model_path"], visible=cfg["SHOW_ASPEN"])
    log_message("Base case reopened for new simulation.", log_lines)
    return aspen


def simulate_one_day(aspen, cfg, row, sim_number, log_lines, last_success_h2=None):
    day_value = row[cfg["DAY_COLUMN"]]
    h2_value = row[cfg["H2_COLUMN"]]
    h2_path = cfg["INPUT_VARIABLES"]["H2_mass_flow"]
    comments, countermeasures, error_messages = [], [], []
    convergence_level_used = 0
    corrective_routine_executed = 0
    log_message(f'Simulation "{sim_number}": running', log_lines)
    try:
        reset_convergence_to_base(aspen, cfg)
    except Exception as e:
        comments.append(f"BASE reset could not be applied: {safe_str(e)}")
    try:
        set_node_value(aspen, h2_path, h2_value)
    except Exception as e:
        msg = f'Could not write H2 to "{h2_path}": {safe_str(e)}'
        error_messages.append(msg)
        comments.append(msg)
        log_message(f'Simulation "{sim_number}": running -> "convergence recovery routine executed" -> "results could not be estimated"', log_lines)
        row_result = {"simulation_number": sim_number, "day": day_value, "H2_input_value": h2_value, "feasible": 0, "convergence_level_used": convergence_level_used, "corrective_routine_executed": 1, "run_status": "", "run_status_path": "", "error_message": " | ".join(error_messages), "countermeasures_taken": "input write failed", "comments": " | ".join(comments) if comments else "results could not be estimated"}
        for alias in cfg["OUTPUT_VARIABLES"]:
            row_result[alias] = np.nan
        return row_result, last_success_h2
    attempt1 = run_aspen_once(aspen, cfg)
    if attempt1["converged"]:
        convergence_level_used = 1
        outputs, output_comments = read_outputs(aspen, cfg["OUTPUT_VARIABLES"])
        comments.extend(output_comments)
        log_message(f'Simulation "{sim_number}": running -> "successful convergence" -> data recorded', log_lines)
        row_result = {"simulation_number": sim_number, "day": day_value, "H2_input_value": h2_value, "feasible": 1, "convergence_level_used": convergence_level_used, "corrective_routine_executed": corrective_routine_executed, "run_status": safe_str(attempt1["status_value"]), "run_status_path": safe_str(attempt1["status_path"]), "error_message": "", "countermeasures_taken": "", "comments": " | ".join(comments) if comments else ""}
        row_result.update(outputs)
        return row_result, h2_value
    corrective_routine_executed = 1
    error_messages.append(f'Attempt 1 failed: {safe_str(attempt1["error_text"])}')
    countermeasures.extend(apply_convergence_profile(aspen, cfg["CONVERGENCE_LEVEL2"]))
    comments.append("Level 2 applied: more iterations and/or looser tolerance")
    try:
        set_node_value(aspen, h2_path, h2_value)
    except Exception as e:
        error_messages.append(f"H2 rewrite before attempt 2 failed: {safe_str(e)}")
    attempt2 = run_aspen_once(aspen, cfg)
    if attempt2["converged"]:
        convergence_level_used = 2
        outputs, output_comments = read_outputs(aspen, cfg["OUTPUT_VARIABLES"])
        comments.extend(output_comments)
        log_message(f'Simulation "{sim_number}": running -> "convergence recovery routine executed" -> data recorded', log_lines)
        row_result = {"simulation_number": sim_number, "day": day_value, "H2_input_value": h2_value, "feasible": 1, "convergence_level_used": convergence_level_used, "corrective_routine_executed": corrective_routine_executed, "run_status": safe_str(attempt2["status_value"]), "run_status_path": safe_str(attempt2["status_path"]), "error_message": " | ".join([m for m in error_messages if m]), "countermeasures_taken": " | ".join(countermeasures), "comments": " | ".join(comments) if comments else ""}
        row_result.update(outputs)
        return row_result, h2_value
    error_messages.append(f'Attempt 2 failed: {safe_str(attempt2["error_text"])}')
    countermeasures.extend(apply_convergence_profile(aspen, cfg["CONVERGENCE_LEVEL3"]))
    comments.append("Level 3 applied: method change and/or warm start")
    warm_info = try_warm_start_with_last_success(aspen, cfg, last_success_h2)
    countermeasures.append(warm_info["message"])
    try:
        set_node_value(aspen, h2_path, h2_value)
    except Exception as e:
        error_messages.append(f"H2 rewrite before attempt 3 failed: {safe_str(e)}")
    attempt3 = run_aspen_once(aspen, cfg)
    if attempt3["converged"]:
        convergence_level_used = 3
        outputs, output_comments = read_outputs(aspen, cfg["OUTPUT_VARIABLES"])
        comments.extend(output_comments)
        log_message(f'Simulation "{sim_number}": running -> "convergence recovery routine executed" -> data recorded', log_lines)
        row_result = {"simulation_number": sim_number, "day": day_value, "H2_input_value": h2_value, "feasible": 1, "convergence_level_used": convergence_level_used, "corrective_routine_executed": corrective_routine_executed, "run_status": safe_str(attempt3["status_value"]), "run_status_path": safe_str(attempt3["status_path"]), "error_message": " | ".join([m for m in error_messages if m]), "countermeasures_taken": " | ".join(countermeasures), "comments": " | ".join(comments) if comments else ""}
        row_result.update(outputs)
        return row_result, h2_value
    comments.append("results could not be estimated")
    log_message(f'Simulation "{sim_number}": running -> "convergence recovery routine executed" -> "results could not be estimated"', log_lines)
    row_result = {"simulation_number": sim_number, "day": day_value, "H2_input_value": h2_value, "feasible": 0, "convergence_level_used": 0, "corrective_routine_executed": corrective_routine_executed, "run_status": safe_str(attempt3["status_value"]), "run_status_path": safe_str(attempt3["status_path"]), "error_message": " | ".join([m for m in error_messages if m]), "countermeasures_taken": " | ".join(countermeasures), "comments": " | ".join(comments)}
    for alias in cfg["OUTPUT_VARIABLES"]:
        row_result[alias] = np.nan
    return row_result, last_success_h2


def main():
    cfg = deepcopy(CONFIG)
    cfg = prompt_if_missing(cfg)
    output_dir = ensure_dir(cfg["OUTPUT_DIR"])
    output_excel = output_dir / cfg["OUTPUT_EXCEL_NAME"]
    output_pdf = output_dir / cfg["OUTPUT_PDF_NAME"]
    output_log_txt = output_dir / cfg["OUTPUT_LOG_TXT_NAME"]
    log_lines = []
    model_info = None
    aspen = None
    temp_dir_to_cleanup = None
    try:
        log_message("Automation process started.", log_lines)
        model_info = resolve_aspen_model_path(cfg["ASPEN_FILE"])
        temp_dir_to_cleanup = model_info["temp_dir"]
        log_message(model_info["message"], log_lines)
        input_excel = Path(cfg["INPUT_EXCEL"]).expanduser().resolve()
        if not input_excel.exists():
            raise FileNotFoundError(f"Input Excel file not found: {input_excel}")
        df = pd.read_excel(input_excel, sheet_name=cfg["INPUT_SHEET"])
        log_message(f"Excel loaded successfully: {input_excel}", log_lines)
        required_cols = [cfg["DAY_COLUMN"], cfg["H2_COLUMN"]]
        missing_cols = [c for c in required_cols if c not in df.columns]
        if missing_cols:
            raise KeyError(f"Missing columns in input Excel: {missing_cols}. Available columns: {list(df.columns)}")
        aspen = open_aspen_case(model_info["model_path"], visible=cfg["SHOW_ASPEN"])
        log_message("Aspen Plus opened successfully.", log_lines)
        validate_nodes(aspen, cfg, log_lines)
        all_results = []
        last_success_h2 = None
        iterable = list(df.itertuples(index=False))
        total = len(iterable)
        progress_iter = progress_wrapper(iterable, total=total, desc="Overall progress")
        for sim_number, row_tuple in enumerate(progress_iter, start=1):
            row = row_tuple._asdict()
            try:
                aspen = reopen_base_case_if_requested(cfg, aspen, model_info, log_lines)
                result_row, last_success_h2 = simulate_one_day(aspen, cfg, row, sim_number, log_lines, last_success_h2)
            except Exception as e:
                log_message(f'Simulation "{sim_number}": running -> "convergence recovery routine executed" -> "results could not be estimated"', log_lines)
                result_row = {"simulation_number": sim_number, "day": row.get(cfg["DAY_COLUMN"], ""), "H2_input_value": row.get(cfg["H2_COLUMN"], np.nan), "feasible": 0, "convergence_level_used": 0, "corrective_routine_executed": 1, "run_status": "", "run_status_path": "", "error_message": safe_str(e), "countermeasures_taken": "general error captured by main loop", "comments": "results could not be estimated"}
                for alias in cfg["OUTPUT_VARIABLES"]:
                    result_row[alias] = np.nan
            try:
                reset_convergence_to_base(aspen, cfg, log_lines)
            except Exception as e:
                log_message(f"Could not reset BASE parameters after simulation {sim_number}: {safe_str(e)}", log_lines)
            all_results.append(result_row)
            if cfg["PAUSE_BETWEEN_RUNS_SEC"] > 0:
                time.sleep(cfg["PAUSE_BETWEEN_RUNS_SEC"])
        results_df = pd.DataFrame(all_results)
        log_df = pd.DataFrame({"log_message": log_lines})
        config_flat = []
        for key, value in cfg.items():
            if isinstance(value, dict):
                for k2, v2 in value.items():
                    config_flat.append({"parameter": f"{key}.{k2}", "value": safe_str(v2)})
            elif isinstance(value, list):
                config_flat.append({"parameter": key, "value": " | ".join(map(safe_str, value))})
            else:
                config_flat.append({"parameter": key, "value": safe_str(value)})
        config_df = pd.DataFrame(config_flat)
        save_excel_report(results_df, log_df, config_df, output_excel)
        save_txt_log(log_lines, output_log_txt)
        save_pdf_report(output_pdf, log_lines, results_df, cfg)
        feasible_count = int(results_df["feasible"].fillna(0).sum()) if not results_df.empty else 0
        failed_count = len(results_df) - feasible_count
        print("
" + "=" * 70)
        print("EXECUTION FINISHED")
        print(f"Total simulations: {len(results_df)}")
        print(f"Successful convergence: {feasible_count}")
        print(f"Did not converge: {failed_count}")
        print(f"Excel: {output_excel}")
        print(f"PDF: {output_pdf}")
        print("=" * 70)
    finally:
        close_aspen_case(aspen)
        if temp_dir_to_cleanup and Path(temp_dir_to_cleanup).exists():
            shutil.rmtree(temp_dir_to_cleanup, ignore_errors=True)


if __name__ == "__main__":
    main()
"""

DEFAULT_OUTPUTS = """Methanol_prod_mol_h=
Methanol_prod_kg_h=
Methanol_purity=
CO2_conversion=
H2_utilization=
Vent_H2_loss=
Recycle_total_flow=
Compressor_power=
Reactor_heat_duty="""

DEFAULT_RUN_STATUS = """\Data\Results Summary\Run-Status\Output\UOSSTAT
\Data\Results Summary\Run-Status\Output\RUNSTAT
\Data\Results Summary\Output\RUN-STATUS
\Data\Convergence\Output\STATUS"""

DEFAULT_SUCCESS_VALUES = """8
OK
Converged
Results Available
Completed
Success"""


def parse_alias_lines(text: str):
    result = {}
    for raw in text.splitlines():
        line = raw.strip()
        if not line or "=" not in line:
            continue
        alias, path = line.split("=", 1)
        result[alias.strip()] = path.strip()
    return result


def parse_lines(text: str):
    return [line.strip() for line in text.splitlines() if line.strip()]


def parse_success_values(text: str):
    values = []
    for line in parse_lines(text):
        if line.lstrip("-").isdigit():
            values.append(int(line))
        else:
            values.append(line)
    return values


def python_literal(value):
    return pformat(value, width=110, sort_dicts=False)


def build_config():
    return {
        "ASPEN_FILE": st.session_state.aspen_file.strip(),
        "INPUT_EXCEL": st.session_state.input_excel.strip(),
        "INPUT_SHEET": int(st.session_state.input_sheet) if str(st.session_state.input_sheet).strip().isdigit() else st.session_state.input_sheet.strip(),
        "OUTPUT_DIR": st.session_state.output_dir.strip(),
        "OUTPUT_EXCEL_NAME": st.session_state.output_excel_name.strip() or "PtMeOH_365_results.xlsx",
        "OUTPUT_PDF_NAME": st.session_state.output_pdf_name.strip() or "PtMeOH_execution_report.pdf",
        "OUTPUT_LOG_TXT_NAME": st.session_state.output_log_name.strip() or "PtMeOH_execution_log.txt",
        "DAY_COLUMN": st.session_state.day_column.strip() or "day",
        "H2_COLUMN": st.session_state.h2_column.strip() or "H2_kg_h",
        "SHOW_ASPEN": bool(st.session_state.show_aspen),
        "REOPEN_BASE_EACH_DAY": bool(st.session_state.reopen_base_each_day),
        "PAUSE_BETWEEN_RUNS_SEC": float(st.session_state.pause_between_runs_sec),
        "INPUT_VARIABLES": {"H2_mass_flow": st.session_state.h2_path.strip()},
        "OUTPUT_VARIABLES": parse_alias_lines(st.session_state.output_variables),
        "RUN_STATUS_CANDIDATES": parse_lines(st.session_state.run_status_candidates),
        "SUCCESS_STATUS_VALUES": parse_success_values(st.session_state.success_status_values),
        "CONVERGENCE_BASE": {
            "label": "BASE",
            "max_iter_path": st.session_state.base_max_iter_path.strip(),
            "max_iter_value": int(st.session_state.base_max_iter_value),
            "tolerance_path": st.session_state.base_tolerance_path.strip(),
            "tolerance_value": float(st.session_state.base_tolerance_value),
            "method_path": st.session_state.base_method_path.strip(),
            "method_value": st.session_state.base_method_value.strip(),
        },
        "CONVERGENCE_LEVEL2": {
            "label": "LEVEL2",
            "max_iter_path": st.session_state.l2_max_iter_path.strip(),
            "max_iter_value": int(st.session_state.l2_max_iter_value),
            "tolerance_path": st.session_state.l2_tolerance_path.strip(),
            "tolerance_value": float(st.session_state.l2_tolerance_value),
            "method_path": st.session_state.l2_method_path.strip(),
            "method_value": st.session_state.l2_method_value.strip(),
        },
        "CONVERGENCE_LEVEL3": {
            "label": "LEVEL3",
            "max_iter_path": st.session_state.l3_max_iter_path.strip(),
            "max_iter_value": int(st.session_state.l3_max_iter_value),
            "tolerance_path": st.session_state.l3_tolerance_path.strip(),
            "tolerance_value": float(st.session_state.l3_tolerance_value),
            "method_path": st.session_state.l3_method_path.strip(),
            "method_value": st.session_state.l3_method_value.strip(),
        },
    }


def generate_script(config: dict):
    return TEMPLATE.replace("__CONFIG_PYTHON__", python_literal(config))


def load_example():
    st.session_state.aspen_file = "C:\project\PtMeOH_model.apwz"
    st.session_state.input_excel = "C:\project\input_365days.xlsx"
    st.session_state.output_dir = "C:\project results"
    st.session_state.h2_path = r"\Data\Streams\SH2\Input\FLOW"
    st.session_state.output_variables = """Methanol_prod_mol_h=\Data\Streams\CRUDEMEO\Output\MOLEFLOW\MIXED\METHANOL
Methanol_prod_kg_h=\Data\Streams\CRUDEMEO\Output\MASSFLOW\MIXED\METHANOL
Vent_H2_mol_h=\Data\Streams\VENT\Output\MOLEFLOW\MIXED\H2
Recycle_total_mol_h=\Data\Streams\RECLP\Output\TOTALFLOW\MIXED
Power_C1_W=\Data\Blocks\C1\Output\WNET
Reactor_heat_duty_W=\Data\Blocks\REACTOR\Output\QNET"""
    st.session_state.base_max_iter_path = r"\Data\Convergence\EO-Conv-Options\Input\MAXIT"
    st.session_state.base_tolerance_path = r"\Data\Convergence\EO-Conv-Options\Input\TOLF"
    st.session_state.base_method_path = ""
    st.session_state.l2_max_iter_path = r"\Data\Convergence\EO-Conv-Options\Input\MAXIT"
    st.session_state.l2_tolerance_path = r"\Data\Convergence\EO-Conv-Options\Input\TOLF"
    st.session_state.l2_method_path = ""
    st.session_state.l3_max_iter_path = r"\Data\Convergence\EO-Conv-Options\Input\MAXIT"
    st.session_state.l3_tolerance_path = r"\Data\Convergence\EO-Conv-Options\Input\TOLF"
    st.session_state.l3_method_path = ""


def init_state():
    defaults = {
        "script_name": "aspen_automation_generated.py",
        "aspen_file": "",
        "input_excel": "",
        "output_dir": "",
        "input_sheet": "0",
        "pause_between_runs_sec": 0.0,
        "show_aspen": True,
        "reopen_base_each_day": False,
        "output_excel_name": "PtMeOH_365_results.xlsx",
        "output_pdf_name": "PtMeOH_execution_report.pdf",
        "output_log_name": "PtMeOH_execution_log.txt",
        "day_column": "day",
        "h2_column": "H2_kg_h",
        "h2_path": "",
        "output_variables": DEFAULT_OUTPUTS,
        "run_status_candidates": DEFAULT_RUN_STATUS,
        "success_status_values": DEFAULT_SUCCESS_VALUES,
        "base_max_iter_path": "",
        "base_max_iter_value": 50,
        "base_tolerance_path": "",
        "base_tolerance_value": 1e-6,
        "base_method_path": "",
        "base_method_value": "Broyden",
        "l2_max_iter_path": "",
        "l2_max_iter_value": 100,
        "l2_tolerance_path": "",
        "l2_tolerance_value": 1e-5,
        "l2_method_path": "",
        "l2_method_value": "Broyden",
        "l3_max_iter_path": "",
        "l3_max_iter_value": 200,
        "l3_tolerance_path": "",
        "l3_tolerance_value": 1e-4,
        "l3_method_path": "",
        "l3_method_value": "Newton",
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


init_state()

st.title("Aspen Plus Automation Script Generator")
st.caption("Generate a Python automation script for batch Aspen Plus simulations, including Excel input, output extraction, and three-level convergence recovery.")

col_top_1, col_top_2 = st.columns([1, 1])
with col_top_1:
    if st.button("Load sample configuration"):
        load_example()
        st.rerun()
with col_top_2:
    st.info("Paste only Aspen internal paths such as \Data\Streams\... Do not paste Application.Tree.FindNode(...).")

left, right = st.columns([1, 1.2])
with left:
    st.subheader("Files and execution")
    st.text_input("Output Python filename", key="script_name")
    st.text_input("Aspen file path (.bkp, .apw, .apwz)", key="aspen_file")
    st.text_input("Input Excel path", key="input_excel")
    st.text_input("Output folder", key="output_dir")
    c1, c2 = st.columns(2)
    with c1:
        st.text_input("Excel sheet", key="input_sheet")
    with c2:
        st.number_input("Pause between runs (s)", min_value=0.0, step=0.1, key="pause_between_runs_sec")
    c3, c4 = st.columns(2)
    with c3:
        st.checkbox("Show Aspen during execution", key="show_aspen")
    with c4:
        st.checkbox("Reopen base case every day", key="reopen_base_each_day")
    st.text_input("Results Excel filename", key="output_excel_name")
    st.text_input("Results PDF filename", key="output_pdf_name")
    st.text_input("Log TXT filename", key="output_log_name")
    st.subheader("Excel input")
    c5, c6 = st.columns(2)
    with c5:
        st.text_input("Day column", key="day_column")
    with c6:
        st.text_input("H2 column", key="h2_column")
    st.subheader("Aspen input variable")
    st.text_input("Aspen internal path for H2_mass_flow", key="h2_path")
    st.subheader("Output variables")
    st.caption("One per line using alias=path")
    st.text_area("Output variables list", key="output_variables", height=220)
with right:
    st.subheader("Run status")
    st.caption("One Aspen internal path per line")
    st.text_area("Run status candidate paths", key="run_status_candidates", height=140)
    st.text_area("Values considered successful", key="success_status_values", height=120)
    st.subheader("Convergence profiles")
    st.markdown("**BASE**")
    b1, b2 = st.columns(2)
    with b1:
        st.text_input("BASE max_iter_path", key="base_max_iter_path")
        st.text_input("BASE tolerance_path", key="base_tolerance_path")
        st.text_input("BASE method_path", key="base_method_path")
    with b2:
        st.number_input("BASE max_iter_value", min_value=1, step=1, key="base_max_iter_value")
        st.number_input("BASE tolerance_value", min_value=0.0, format="%.10g", key="base_tolerance_value")
        st.text_input("BASE method_value", key="base_method_value")
    st.markdown("**LEVEL2**")
    l21, l22 = st.columns(2)
    with l21:
        st.text_input("LEVEL2 max_iter_path", key="l2_max_iter_path")
        st.text_input("LEVEL2 tolerance_path", key="l2_tolerance_path")
        st.text_input("LEVEL2 method_path", key="l2_method_path")
    with l22:
        st.number_input("LEVEL2 max_iter_value", min_value=1, step=1, key="l2_max_iter_value")
        st.number_input("LEVEL2 tolerance_value", min_value=0.0, format="%.10g", key="l2_tolerance_value")
        st.text_input("LEVEL2 method_value", key="l2_method_value")
    st.markdown("**LEVEL3**")
    l31, l32 = st.columns(2)
    with l31:
        st.text_input("LEVEL3 max_iter_path", key="l3_max_iter_path")
        st.text_input("LEVEL3 tolerance_path", key="l3_tolerance_path")
        st.text_input("LEVEL3 method_path", key="l3_method_path")
    with l32:
        st.number_input("LEVEL3 max_iter_value", min_value=1, step=1, key="l3_max_iter_value")
        st.number_input("LEVEL3 tolerance_value", min_value=0.0, format="%.10g", key="l3_tolerance_value")
        st.text_input("LEVEL3 method_value", key="l3_method_value")

config = build_config()
missing = []
if not config["ASPEN_FILE"]:
    missing.append("Aspen file path")
if not config["INPUT_EXCEL"]:
    missing.append("Input Excel path")
if not config["OUTPUT_DIR"]:
    missing.append("Output folder")
if not config["INPUT_VARIABLES"]["H2_mass_flow"]:
    missing.append("H2_mass_flow internal path")
if missing:
    st.warning("Missing fields: " + ", ".join(missing) + ". If you leave them empty, the generated script will ask for them in the console when it runs.")
else:
    st.success("Configuration looks complete. The generated script will embed Aspen file, Excel path, output folder, and H2 internal path directly inside CONFIG.")

script_text = generate_script(config)
st.subheader("Generated Python script")
st.code(script_text, language="python")
st.download_button("Download generated .py file", data=script_text, file_name=st.session_state.script_name.strip() or "aspen_automation_generated.py", mime="text/x-python")

st.subheader("How to deploy this app")
st.markdown("1. Upload `streamlit_app.py`, `requirements.txt`, and `README.md` to a GitHub repository. 2. In Streamlit Community Cloud, create a new app and point it to `streamlit_app.py`.3. Once deployed, open the app, fill the form, and download your generated Aspen automation script.")
