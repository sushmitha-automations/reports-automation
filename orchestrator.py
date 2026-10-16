"""
Master Orchestrator for Daily Report Automation.
Coordinates Polars loaders, Adherence & Onboarding evaluation engines,
and OpenPyXL report builders.
"""
import os
import sys
import glob
from datetime import datetime, date
from typing import Dict, List, Optional, Tuple, Callable

from config import (
    LM_PODS, HL_REGIONS, FM_REGIONS, HORO_SITES
)
from data_loaders import (
    load_facility_master, load_user_master, load_valinor_attendance,
    load_chroma_masters, load_onboarding_funnel
)
from adherence_engine import evaluate_attendance_adherence
from punch_engine import build_punch_adherence_dataset
from excel_builder import (
    build_attendance_adherence_report, build_onboarding_report,
    build_punch_adherence_report
)

def get_month_dates(year: int, month: int, cutoff_day: int) -> List[str]:
    """Generates a list of date strings 'YYYY-MM-DD' from day 1 up to cutoff_day."""
    dates = []
    for d in range(1, cutoff_day + 1):
        dates.append(f"{year:04d}-{month:02d}-{d:02d}")
    return dates

def find_file_by_pattern(base_dir: str, patterns: List[str]) -> Optional[str]:
    """Finds first matching file in base_dir given multiple case-insensitive patterns."""
    if not os.path.exists(base_dir):
        return None
    dir_files = [f for f in os.listdir(base_dir) if not f.startswith('~$')]
    for pattern in patterns:
        matches = glob.glob(os.path.join(base_dir, pattern))
        if matches:
            valid_m = [m for m in matches if not os.path.basename(m).startswith('~$')]
            if valid_m:
                return valid_m[0]
        # Case-insensitive substring check
        clean_p = pattern.replace('*', '').lower()
        for f in dir_files:
            if clean_p in f.lower():
                return os.path.join(base_dir, f)
    return None

def run_attendance_adherence(
    base_dir: str,
    output_dir: str,
    reporting_date_str: str = "30th Sep",
    year: int = 2026,
    month: int = 9,
    cutoff_day: int = 30,
    progress_callback: Optional[Callable[[float, str], None]] = None
) -> str:
    """
    Executes the complete Attendance Adherence Report pipeline.
    """
    def notify(pct: float, msg: str):
        if progress_callback:
            progress_callback(pct, msg)
        print(f"[{int(pct * 100)}%] {msg}")

    os.makedirs(output_dir, exist_ok=True)
    date_list = get_month_dates(year, month, cutoff_day)

    notify(0.02, f"Loading Facility Master CSV...")
    fac_csv = find_file_by_pattern(base_dir, ['*Facility*.csv', '*facility*.csv', 'Valinor Dump-Facility.csv'])
    if not fac_csv:
        raise FileNotFoundError(f"Valinor Dump-Facility.csv not found in {base_dir}")

    facilities, fac_pod_map = load_facility_master(fac_csv)
    print(f"      Loaded {len(facilities)} facilities ({len(fac_pod_map)} mapped PODs).")

    notify(0.08, "Loading & indexing Master User Database with Polars...")
    user_csvs = glob.glob(os.path.join(base_dir, '*all_user_data*.csv'))
    if not user_csvs:
        user_csvs = glob.glob(os.path.join(base_dir, '*user_data*.csv'))
    if not user_csvs:
        raise FileNotFoundError(f"No all_user_data CSV files found in {base_dir}")

    least_priority_file = find_file_by_pattern(base_dir, ['*Least Priority*.xlsx', '*least priority*.xlsx'])

    user_rows, val_map_by_aadh, val_map_by_mob, val_map_by_id = load_user_master(
        user_csvs, least_priority_file, fac_pod_map
    )
    val_dump_role_map = {u['numericId']: u.get('rolename', 'DE') for u in user_rows}
    print(f"      Loaded {len(user_rows)} users into Valinor Dump.")

    notify(0.20, "Loading Biometric Attendance punches (Polars)...")
    att_path = find_file_by_pattern(
        base_dir,
        [
            '*Attendance*.xlsx', '*attendance*.xlsx',
            '*Attendance*.csv', '*attendance*.csv',
            '*Attendance*.xls', '*attendance*.xls',
            'Valinor Attendance.xlsx', 'Valinor Attendance.csv',
            'Attendance'
        ]
    )
    if not att_path:
        raise FileNotFoundError(f"Valinor Attendance file not found in {base_dir}")

    val_att_rows, val_att_map = load_valinor_attendance(att_path, date_list)
    print(f"      Loaded attendance for {len(val_att_rows)} active records.")

    notify(0.30, "Loading Chroma Roster & Attendance masters (LM, FM, HL)...")
    lm_file = find_file_by_pattern(base_dir, ['*(LM)*.xlsx', '*LM*.xlsx'])
    fm_file = find_file_by_pattern(base_dir, ['*(FM)*.xlsx', '*FM*.xlsx'])
    hl_file = find_file_by_pattern(base_dir, ['*(HL*.xlsx', '*HL*.xlsx', '*Corp*.xlsx'])

    chroma_configs = []
    if lm_file:
        chroma_configs.append({'file_path': lm_file, 'type': 'LM'})
    if fm_file:
        chroma_configs.append({'file_path': fm_file, 'type': 'FM'})
    if hl_file:
        chroma_configs.append({'file_path': hl_file, 'type': 'HL'})

    if not chroma_configs:
        raise FileNotFoundError(f"No Chroma master files (LM, FM, HL) found in {base_dir}")

    chroma_rows = load_chroma_masters(chroma_configs, date_list, fac_pod_map)
    print(f"      Loaded {len(chroma_rows)} unique Chroma employee records.")

    notify(0.40, "Evaluating 3-Way Identity Resolution & Adherence Matrix...")
    evaluated_val_base = evaluate_attendance_adherence(
        chroma_rows, val_map_by_aadh, val_map_by_mob, val_map_by_id,
        val_att_map, fac_pod_map, date_list
    )

    out_file = os.path.join(output_dir, f"Attendance Adherence Report - {reporting_date_str}.xlsx")
    notify(0.50, "Generating Excel sheets with formulas...")
    excel_cb = (lambda p, m: notify(0.50 + p * 0.49, m)) if progress_callback else None
    build_attendance_adherence_report(
        out_file, date_list, evaluated_val_base, user_rows,
        facilities, val_att_rows, chroma_rows, val_dump_role_map,
        progress_callback=excel_cb
    )
    notify(1.00, f"Attendance Adherence Report created successfully: {os.path.basename(out_file)}")
    return out_file

def run_onboarding_report(
    base_dir: str,
    output_dir: str,
    reporting_date_str: str = "30th Sep",
    year: int = 2026,
    month: int = 9,
    cutoff_day: int = 30,
    progress_callback: Optional[Callable[[float, str], None]] = None
) -> str:
    """
    Executes the complete Onboarding Report pipeline.
    """
    def notify(pct: float, msg: str):
        if progress_callback:
            progress_callback(pct, msg)
        print(f"[{int(pct * 100)}%] {msg}")

    os.makedirs(output_dir, exist_ok=True)
    start_date = f"{year:04d}-{month:02d}-01"
    cutoff_date = f"{year:04d}-{month:02d}-{cutoff_day:02d}"

    notify(0.05, "Loading Onboarding Funnel CSV with Polars...")
    funnel_csv = find_file_by_pattern(base_dir, ['*onboarding_funnel*.csv', '*funnel*.csv'])
    if not funnel_csv:
        raise FileNotFoundError(f"No onboarding funnel CSV found in {base_dir}")

    mig_excel = find_file_by_pattern(base_dir, ['*migration*.xlsx', '*Migration*.xlsx'])

    notify(0.25, "Processing KYC compliance, Migrations & Ageing metrics...")
    sorted_dates, base_rows, base_headers = load_onboarding_funnel(
        funnel_csv, mig_excel, start_date, cutoff_date
    )
    print(f"      Net base records: {len(base_rows)} rows across {len(sorted_dates)} distinct dates.")

    out_file = os.path.join(output_dir, f"Onboarding Report - {reporting_date_str}.xlsx")
    notify(0.40, "Generating Excel sheets with live formulas...")
    excel_cb = (lambda p, m: notify(0.40 + p * 0.59, m)) if progress_callback else None
    build_onboarding_report(
        out_file, sorted_dates, base_rows, base_headers,
        progress_callback=excel_cb
    )

    notify(1.00, f"Onboarding Report created successfully: {os.path.basename(out_file)}")
    return out_file

def run_punch_adherence_report(
    base_dir: str,
    output_dir: str,
    reporting_date_str: str = "10th Oct",
    year: int = 2026,
    month: int = 10,
    cutoff_day: int = 10,
    progress_callback: Optional[Callable[[float, str], None]] = None
) -> str:
    """
    Executes the complete Punch Adherence Report pipeline.
    """
    def notify(pct: float, msg: str):
        if progress_callback:
            progress_callback(pct, msg)
        print(f"[{int(pct * 100)}%] {msg}")

    os.makedirs(output_dir, exist_ok=True)
    date_list = get_month_dates(year, month, cutoff_day)

    notify(0.04, "Locating Onroll & Offroll Headcount and Facility Master...")
    onroll_file = find_file_by_pattern(base_dir, ['*Onroll*.xlsx', '*onroll*.xlsx', 'Onroll HC.xlsx'])
    if not onroll_file:
        raise FileNotFoundError(f"Onroll HC.xlsx not found in {base_dir}")

    notify(0.18, "Locating Chroma masters / Valinor base files for Offroll workforce...")
    chroma_files = []
    for pat in ['*(LM)*.xlsx', '*LM*.xlsx', '*(FM)*.xlsx', '*FM*.xlsx', '*(HL*.xlsx', '*HL*.xlsx', '*Corp*.xlsx', '*Valinor*.xlsx']:
        matches = glob.glob(os.path.join(base_dir, pat))
        for m in matches:
            if m not in chroma_files and not os.path.basename(m).startswith('~$'):
                chroma_files.append(m)

    offroll_file = find_file_by_pattern(base_dir, ['*Offroll*.xlsx', '*offroll*.xlsx', 'Offroll HC.xlsx'])
    if not offroll_file and not chroma_files:
        raise FileNotFoundError(f"Neither Offroll HC.xlsx nor Valinor base files found in {base_dir}")

    fac_csv = find_file_by_pattern(base_dir, ['*Facility*.csv', '*facility*.csv', 'Valinor Dump-Facility.csv'])
    if not fac_csv:
        raise FileNotFoundError(f"Valinor Dump-Facility.csv not found in {base_dir}")

    notify(0.20, "Locating Biometric Attendance punch records...")
    att_csv = find_file_by_pattern(
        base_dir,
        [
            '*Attendance*.xlsx', '*attendance*.xlsx',
            '*Attendance*.csv', '*attendance*.csv',
            '*Attendance*.xls', '*attendance*.xls',
            'Valinor Attendance.xlsx', 'Valinor Attendance.csv',
            'Attendance'
        ]
    )
    if not att_csv:
        raise FileNotFoundError(f"Valinor Attendance file not found in {base_dir}")

    notify(0.25, f"Assembling workforce & joining daily punches with Polars ({len(date_list)} days)...")
    workforce_records, base_hub_combos = build_punch_adherence_dataset(
        onroll_path=onroll_file,
        offroll_path=offroll_file,
        attendance_csv_path=att_csv,
        facility_master_csv_path=fac_csv,
        chroma_file_paths=chroma_files,
        date_list=date_list
    )
    print(f"      Net workforce assembled: {len(workforce_records)} records, {len(base_hub_combos)} facility combos.")

    out_file = os.path.join(output_dir, f"Punch Adherence Report - {reporting_date_str}.xlsx")
    notify(0.40, "Generating Excel sheets with formulas and Aptos Narrow styling...")
    excel_cb = (lambda p, m: notify(0.40 + p * 0.59, m)) if progress_callback else None
    build_punch_adherence_report(
        output_path=out_file,
        date_list=date_list,
        workforce_records=workforce_records,
        base_hub_combos=base_hub_combos,
        progress_callback=excel_cb
    )

    notify(1.00, f"Punch Adherence Report created successfully: {os.path.basename(out_file)}")
    return out_file

if __name__ == '__main__':
    default_base = os.path.join(os.path.dirname(__file__), 'Base Files', '30th Sep')
    default_out = os.path.join(os.path.dirname(__file__), 'Reports')

    base_arg = sys.argv[1] if len(sys.argv) > 1 else default_base
    out_arg = sys.argv[2] if len(sys.argv) > 2 else default_out
    date_arg = sys.argv[3] if len(sys.argv) > 3 else "30th Sep"

    print("===============================================================================")
    print("           STARTING PYTHON REPORT AUTOMATION (POLARS + OPENPYXL)              ")
    print("===============================================================================")
    print(f"Base Directory:   {base_arg}")
    print(f"Output Directory: {out_arg}")
    print(f"Reporting Date:   {date_arg}")

    run_onboarding_report(base_arg, out_arg, date_arg)
    run_attendance_adherence(base_arg, out_arg, date_arg)
    print("\nALL REPORTS GENERATED SUCCESSFULLY!")
