"""
Punch Adherence Engine.
Assembles total workforce headcount from Onroll HC & Offroll HC files,
enriches supervisor roles from Chroma masters, joins daily biometric attendance punches,
and constructs Attendance Base, Punched Base, and Base data structures using Polars.
"""
import os
from typing import Dict, List, Tuple, Any, Optional
import polars as pl
import openpyxl

from config import (
    PUNCH_LM_PODS, PUNCH_HL_REGIONS, PUNCH_FM_REGIONS, PUNCH_HORO_SITES
)


def determine_vehicle_vertical(pod_name: str) -> str:
    """Classifies a POD/Site into LM, FM, HL, or HO-RO vertical."""
    p_upper = str(pod_name or '').strip().upper()
    if p_upper.startswith("FM_"):
        return "FM"
    if p_upper.startswith("HL_"):
        return "HL"
    if " - HO" in p_upper or " - RO" in p_upper or p_upper.endswith(" RO") or p_upper.endswith(" HO"):
        if "POD_" in p_upper:
            return "LM"
        return "HO-RO"
    return "LM"


def map_chroma_designation_to_role(desig: str) -> str:
    """Normalizes Chroma designations to punch report role abbreviations."""
    d = str(desig or '').strip().upper()
    if "TEAM LEADER" in d or d == "TL" or "FIELD TL" in d:
        return "TL"
    if "STORE MANAGER" in d:
        return "STORE MANAGER"
    if "PICKER" in d:
        return "PICKER"
    if "HELPER" in d:
        return "HELPER"
    if "DELIVERY" in d or "AMBASSADOR" in d or "DE" in d:
        return "DELIVERY AMBASSADORS"
    if "GC" in d:
        return "GC"
    if "FLEX" in d:
        return "Flex"
    if "IT EXECUTIVE" in d:
        return "IT EXECUTIVE"
    if "FINANCE" in d:
        return "FINANCE EXECUTIVE"
    if "MIS" in d:
        return "Mis Executive"
    if d:
        return str(desig).strip()
    return "Associate"


def build_punch_adherence_dataset(
    onroll_path: str,
    offroll_path: str,
    attendance_csv_path: str,
    facility_master_csv_path: str,
    chroma_file_paths: List[str],
    date_list: List[str]
) -> Tuple[List[Dict[str, Any]], List[Tuple[str, str, str]]]:
    """
    Constructs the complete 14.8k employee punch adherence dataset.
    Returns:
        workforce_records: list of dicts with employee attributes and daily punches
        base_hub_combos: list of unique (Hub Name, POD Name, Veticle) tuples for Base sheet
    """
    # 1. Load Facility Master POD Map
    df_fac = pl.read_csv(facility_master_csv_path, ignore_errors=True, infer_schema_length=0)
    fac_cols = {c.strip().lower(): c for c in df_fac.columns}
    fac_name_col = fac_cols.get('facility name') or df_fac.columns[1]
    pod_name_col = fac_cols.get('pod name') or df_fac.columns[5]

    fac_pod_map: Dict[str, str] = {}
    for r in df_fac.select([fac_name_col, pod_name_col]).iter_rows():
        f_name = str(r[0] or '').strip().lower()
        p_name = str(r[1] or '').strip()
        if f_name and p_name and p_name != 'None':
            fac_pod_map[f_name] = p_name

    # 2. Pre-index Chroma Master files for Off-roll enrichment
    chroma_lookup: Dict[str, Dict[str, str]] = {}
    for c_path in chroma_file_paths:
        if not os.path.exists(c_path):
            continue
        try:
            df_c = pl.read_excel(c_path)
            c_cols = {c.strip().lower(): c for c in df_c.columns}
            id_col = c_cols.get('empoyee id') or c_cols.get('employee id') or df_c.columns[0]
            desig_col = c_cols.get('designation')
            sub_cc_col = c_cols.get('sub cost center') or c_cols.get('cost centre')
            pod_col = c_cols.get('pod name') or c_cols.get('pod')
            branch_col = c_cols.get('branch code')

            for row in df_c.iter_rows(named=True):
                emp_id = str(row.get(id_col) or '').strip()
                if not emp_id or emp_id in chroma_lookup:
                    continue
                chroma_lookup[emp_id] = {
                    'designation': str(row.get(desig_col) or '').strip() if desig_col else '',
                    'department': str(row.get(sub_cc_col) or '').strip() if sub_cc_col else '',
                    'pod': str(row.get(pod_col) or '').strip() if pod_col else '',
                    'branch': str(row.get(branch_col) or '').strip() if branch_col else ''
                }
        except Exception as e:
            print(f"[Warning] Error parsing Chroma file {c_path}: {e}")

    # 3. Load On-Roll Employees (Onroll HC.xlsx -> Main data)
    df_on = pl.read_excel(onroll_path, sheet_name='Main data')
    on_cols = {c.strip().lower(): c for c in df_on.columns}
    on_id_col = on_cols.get('employeecode') or df_on.columns[0]
    on_name_col = on_cols.get('employee name') or df_on.columns[1]
    on_sit_col = on_cols.get('sitting location (office location)') or df_on.columns[4]
    on_dept_col = on_cols.get('sub cost centre') or df_on.columns[6]
    on_pod_col = on_cols.get('pod') or df_on.columns[7]

    workforce: List[Dict[str, Any]] = []

    for r in df_on.iter_rows(named=True):
        emp_id = str(r.get(on_id_col) or '').strip()
        if not emp_id:
            continue
        emp_name = str(r.get(on_name_col) or '').strip()
        sit_loc = str(r.get(on_sit_col) or '').strip()
        dept = str(r.get(on_dept_col) or '').strip()
        raw_pod = str(r.get(on_pod_col) or '').strip()

        # Derive POD with Facility Master priority
        fac_p = fac_pod_map.get(sit_loc.lower(), '')
        if fac_p:
            derived_pod = fac_p
        elif 'HO' in sit_loc or 'RO' in sit_loc:
            derived_pod = sit_loc
        elif raw_pod and raw_pod not in ('Not Applicable', '-', 'None'):
            derived_pod = raw_pod
        else:
            derived_pod = 'Not Applicable'

        workforce.append({
            'id': emp_id,
            'name': emp_name,
            'type': 'On-Roll',
            'pod': derived_pod,
            'facility': sit_loc,
            'department': dept if dept else 'Central',
            'role': 'Employee',
            'punches': {}
        })

    # 4. Load Off-Roll Employees (Offroll HC.xlsx -> Main Data)
    df_off = pl.read_excel(offroll_path, sheet_name='Main Data')
    off_cols = {c.strip().lower(): c for c in df_off.columns}
    off_id_col = off_cols.get('empoyee id') or off_cols.get('employee id') or df_off.columns[0]
    off_name_col = off_cols.get('name') or df_off.columns[2]
    off_branch_col = off_cols.get('branch code') or df_off.columns[3]

    for r in df_off.iter_rows(named=True):
        emp_id = str(r.get(off_id_col) or '').strip()
        if not emp_id:
            continue
        emp_name = str(r.get(off_name_col) or '').strip()
        branch = str(r.get(off_branch_col) or '').strip()

        c_info = chroma_lookup.get(emp_id)
        if c_info:
            role = map_chroma_designation_to_role(c_info.get('designation', ''))
            dept = c_info.get('department') or 'Operations'
            pod = c_info.get('pod') or fac_pod_map.get(branch.lower(), '') or 'Not Applicable'
        else:
            role = 'Associate'
            dept = 'Operations'
            pod = fac_pod_map.get(branch.lower(), '') or 'Not Applicable'

        workforce.append({
            'id': emp_id,
            'name': emp_name,
            'type': 'Off-Roll',
            'pod': pod,
            'facility': branch,
            'department': dept,
            'role': role,
            'punches': {}
        })

    # 5. Load Biometric Attendance Punches (Valinor Attendance.csv / .xlsx)
    if attendance_csv_path.lower().endswith(('.xlsx', '.xls')):
        try:
            df_att = pl.read_excel(attendance_csv_path)
            df_att = df_att.select([pl.col(c).cast(pl.String) for c in df_att.columns])
        except Exception:
            wb = openpyxl.load_workbook(attendance_csv_path, read_only=True, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            wb.close()
            if rows:
                headers = [str(c or '').strip() for c in rows[0]]
                data_dict = {h: [str(r[i] or '').strip() if i < len(r) else '' for r in rows[1:]] for i, h in enumerate(headers)}
                df_att = pl.DataFrame(data_dict)
            else:
                df_att = pl.DataFrame()
    else:
        df_att = pl.read_csv(attendance_csv_path, ignore_errors=True, infer_schema_length=0)

    att_cols = {str(c).strip().lower(): c for c in df_att.columns}
    att_id_col = att_cols.get('employeeid') or df_att.columns[1]

    # Map dates in attendance file: '01-10-2026' or '2026-10-01' -> date_list keys
    att_date_cols: Dict[str, str] = {}
    for col in df_att.columns:
        clean_c = str(col).strip()
        parts = clean_c.split('-')
        # DD-MM-YYYY
        if len(parts) == 3 and len(parts[0]) == 2 and parts[2][:4].isdigit():
            try:
                d_day = parts[0]
                m_num = int(parts[1])
                y_val = parts[2][:4]
                y_short = y_val[-2:]
                m_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
                m_str = m_names[m_num - 1]
                header_fmt = f"{d_day}-{m_str}-{y_short}"
                iso_fmt = f"{int(y_val):04d}-{m_num:02d}-{int(d_day):02d}"
                att_date_cols[header_fmt] = clean_c
                att_date_cols[iso_fmt] = clean_c
            except Exception:
                pass
        # YYYY-MM-DD
        elif len(parts) == 3 and len(parts[0]) == 4:
            try:
                y_val = parts[0]
                m_num = int(parts[1])
                d_day = parts[2][:2]
                y_short = y_val[-2:]
                m_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
                m_str = m_names[m_num - 1]
                header_fmt = f"{d_day}-{m_str}-{y_short}"
                iso_fmt = f"{int(y_val):04d}-{m_num:02d}-{int(d_day):02d}"
                att_date_cols[header_fmt] = clean_c
                att_date_cols[iso_fmt] = clean_c
            except Exception:
                pass

    att_punch_map: Dict[str, Dict[str, str]] = {}
    for r in df_att.iter_rows(named=True):
        raw_emp_id = str(r.get(att_id_col) or '').strip()
        if not raw_emp_id:
            continue
        punches = {}
        for d_hdr in date_list:
            orig_col = att_date_cols.get(d_hdr) or (d_hdr if d_hdr in r else None)
            if orig_col and orig_col in r:
                val = str(r.get(orig_col) or '').strip().upper()
                punches[d_hdr] = 'P' if val == 'P' else 'A'
            else:
                punches[d_hdr] = 'A'
        att_punch_map[raw_emp_id] = punches

    # Assign punches to workforce records
    for emp in workforce:
        emp_id = emp['id']
        clean_numeric_id = emp_id.replace('SF', '').replace('PROD', '').strip()
        p_data = att_punch_map.get(emp_id) or att_punch_map.get(clean_numeric_id) or {}
        for d_hdr in date_list:
            emp['punches'][d_hdr] = p_data.get(d_hdr, 'A')

    # 6. Extract unique Hub & POD combinations for Base sheet
    combo_set = set()
    base_hub_combos: List[Tuple[str, str, str]] = []
    for emp in workforce:
        hub = emp['facility']
        pod = emp['pod']
        key = (hub, pod)
        if key not in combo_set:
            combo_set.add(key)
            vehicle = determine_vehicle_vertical(pod)
            base_hub_combos.append((hub, pod, vehicle))

    return workforce, base_hub_combos
