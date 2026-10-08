"""
Attendance Adherence Evaluation Engine.
Implements the 3-Way Identity Resolution (Aadhaar -> Mobile -> ID)
and evaluates daily attendance status adherence between Chroma and Valinor.
"""
from typing import Dict, List, Any, Optional
from loaders import normalize_pod, map_hl_pod, map_fm_pod

def evaluate_attendance_adherence(
    chroma_rows: List[Dict[str, Any]],
    val_map_by_aadh: Dict[str, Any],
    val_map_by_mob: Dict[str, Any],
    val_map_by_id: Dict[int, Any],
    val_att_map: Dict[int, Dict[str, str]],
    facility_pod_map: Dict[str, str],
    date_list: List[str]
) -> List[Dict[str, Any]]:
    """
    Evaluates Chroma vs Valinor attendance adherence for all unique Chroma employees
    across the entire month (Day 1 to Day N).
    """
    evaluated_val_base: List[Dict[str, Any]] = []

    for item in chroma_rows:
        emp_id = item['emp_id']
        val_id_raw = item['valinor_id_raw']
        mob_str = item['mobile']
        aadh_str = item['aadhaar']
        branch_code = item['branch_code']
        state = item['state']
        base_pod_name = item['base_pod_name']
        f_type = item['file_type']

        # 3-Way Identity Resolution
        matched_user = None
        if aadh_str and aadh_str in val_map_by_aadh:
            matched_user = val_map_by_aadh[aadh_str]
        elif mob_str and mob_str in val_map_by_mob:
            matched_user = val_map_by_mob[mob_str]
        elif val_id_raw and val_id_raw in val_map_by_id:
            matched_user = val_map_by_id[val_id_raw]

        # Valinor POD Determination
        valinor_pod = matched_user.get('pod', '') if matched_user else ''
        if not valinor_pod:
            pod_from_fac = facility_pod_map.get(branch_code.lower(), '')
            if pod_from_fac:
                valinor_pod = pod_from_fac
            elif f_type == 'HL':
                valinor_pod = map_hl_pod(base_pod_name, state)
            elif f_type == 'FM':
                valinor_pod = map_fm_pod(branch_code, state)
            else:
                valinor_pod = normalize_pod(base_pod_name, branch_code) or base_pod_name or 'UNMAPPED'

        # Chroma POD Determination
        chroma_pod = normalize_pod(base_pod_name, branch_code)
        if not chroma_pod:
            pod_from_fac = facility_pod_map.get(branch_code.lower(), '')
            if pod_from_fac:
                chroma_pod = pod_from_fac
            elif f_type == 'HL':
                chroma_pod = map_hl_pod(base_pod_name, state)
            elif f_type == 'FM':
                chroma_pod = map_fm_pod(branch_code, state)
            else:
                chroma_pod = base_pod_name or 'UNMAPPED'

        matched_val_id = matched_user['numericId'] if matched_user else (val_id_raw or '')

        date_punches = item['date_punches']
        date_val_att = []
        date_validations = []

        # Day-by-Day Truth Table Comparison
        for d_idx, d_str in enumerate(date_list):
            w_val = date_punches[d_idx] if d_idx < len(date_punches) else 'NA'
            
            x_val = 'A'
            if matched_val_id and matched_val_id in val_att_map:
                x_val = val_att_map[matched_val_id].get(d_str, 'A')
            elif val_id_raw and val_id_raw in val_att_map:
                x_val = val_att_map[val_id_raw].get(d_str, 'A')
            
            date_val_att.append(x_val)

            # Adherence Truth Matrix
            val_result = 'MATCH'
            is_chroma_pres = (w_val in ('P', 'NPP', 'HD'))

            if x_val == 'Tech Issue':
                val_result = 'Tech Issue'
            elif is_chroma_pres and x_val == 'P':
                val_result = 'MATCH'
            elif is_chroma_pres and x_val == 'A':
                val_result = 'P in Chroma, A in Valinor'
            elif w_val in ('A', 'LWP') and x_val == 'A':
                val_result = 'MATCH'
            elif w_val in ('A', 'LWP') and x_val == 'P':
                val_result = 'A/LWP in Chroma, P in Valinor'
            elif 'NA' in w_val and x_val == 'P':
                val_result = 'NA in Chroma, P in Valinor'
            elif 'NA' in w_val and x_val == 'A':
                val_result = 'MATCH'
            elif w_val in ('WO', 'PL') and x_val == 'P':
                val_result = 'WO/PL/HD in Chroma, P in Valinor'
            elif w_val in ('WO', 'PL') and x_val == 'A':
                val_result = 'WO/PL/HD in Chroma, A in Valinor'

            date_validations.append(val_result)

        evaluated_val_base.append({
            'emp_id': emp_id,
            'valinor_id': matched_val_id,
            'valinor_user': matched_user,
            'chroma_item': item,
            'chroma_pod': chroma_pod,
            'valinor_pod': valinor_pod,
            'date_punches': date_punches,
            'date_val_att': date_val_att,
            'date_validations': date_validations,
            'match_count': date_validations.count('MATCH'),
            'p_chroma_a_val': date_validations.count('P in Chroma, A in Valinor'),
            'wo_chroma_a_val': date_validations.count('WO/PL/HD in Chroma, A in Valinor'),
            'tech_issue': date_validations.count('Tech Issue'),
            'na_chroma_p_val': date_validations.count('NA in Chroma, P in Valinor'),
            'a_chroma_p_val': date_validations.count('A/LWP in Chroma, P in Valinor'),
            'wo_chroma_p_val': date_validations.count('WO/PL/HD in Chroma, P in Valinor'),
        })

    return evaluated_val_base
