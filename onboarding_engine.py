"""
Onboarding Report Audit Engine.
Aggregates digital Aadhaar adoption, Hyperverge Red cases, and Manual KYC overrides
across hiring hubs and HR recruiters.
"""
from typing import Dict, List, Any, Set
from collections import defaultdict

def evaluate_onboarding_metrics(
    base_rows: List[List[Any]],
    base_headers: List[str],
    pod_list: List[str],
    date_list: List[str]
) -> Tuple[Dict[str, Dict[str, Dict[str, int]]], List[Tuple[str, int, int, int]]]:
    """
    Computes daily onboarding counts per POD/Region and extracts HR-wise manual approvals.
    """
    col_idx = {h.strip().lower(): i for i, h in enumerate(base_headers)}
    
    pod_col = col_idx.get('pod name', 10)
    hub_col = col_idx.get('hub name', 9)
    status_col = col_idx.get('status', 3)
    date_col = col_idx.get('onboarded date', 12)
    hr_col = col_idx.get('approved by', 18)
    aadh_col = col_idx.get('onboarded by aadhaar (y/n)', 19)
    auto_ver_col = col_idx.get('auto verification status', 23)
    hyp_status_col = col_idx.get('hyperverge status', 24)
    perm_addr_col = col_idx.get('permanent address', 30)

    pod_set = set(pod_list)
    daily_metrics: Dict[str, Dict[str, Dict[str, int]]] = {
        pod: {d: {'total': 0, 'aadhar': 0, 'manual': 0, 'red': 0} for d in date_list}
        for pod in pod_list
    }

    hr_counts: Dict[str, Dict[str, int]] = defaultdict(lambda: {'manual': 0, 'red': 0})

    for row in base_rows:
        if len(row) <= max(pod_col, hub_col, status_col, date_col, hr_col, aadh_col):
            continue

        p_name = str(row[pod_col] or '').strip()
        h_name = str(row[hub_col] or '').strip()
        status = str(row[status_col] or '').strip().lower()
        d_val = str(row[date_col] or '').strip()
        hr_name = str(row[hr_col] or '').strip() or 'Unknown / System'
        is_aadh_y = str(row[aadh_col] or '').strip().upper() == 'Y'
        auto_ver = str(row[auto_ver_col] or '').strip().lower() if auto_ver_col < len(row) else ''
        hyp_status = str(row[hyp_status_col] or '').strip().lower() if hyp_status_col < len(row) else ''
        perm_addr = str(row[perm_addr_col] or '').strip().lower() if perm_addr_col < len(row) else ''

        if not status.startswith('onboarded') or 'hub migration' in perm_addr:
            continue

        # Match POD or Hub Name
        target_pod = None
        if p_name in pod_set:
            target_pod = p_name
        elif h_name in pod_set:
            target_pod = h_name

        if target_pod and d_val in daily_metrics[target_pod]:
            m = daily_metrics[target_pod][d_val]
            m['total'] += 1
            if is_aadh_y:
                m['aadhar'] += 1
            else:
                m['manual'] += 1

            is_red = (auto_ver == 'red' or hyp_status == 'red' or auto_ver == 'failed')
            if is_red:
                m['red'] += 1

        # Track HR approvals for this operational vertical
        if target_pod:
            is_manual = not is_aadh_y
            is_red_flag = (auto_ver == 'red' or hyp_status == 'red' or auto_ver == 'failed')
            if is_manual or is_red_flag:
                if is_manual:
                    hr_counts[hr_name]['manual'] += 1
                if is_red_flag:
                    hr_counts[hr_name]['red'] += 1

    sorted_hr = []
    for hr, c in hr_counts.items():
        tot = c['manual'] + c['red']
        sorted_hr.append((hr, c['manual'], c['red'], tot))
    sorted_hr.sort(key=lambda x: x[3], reverse=True)

    return daily_metrics, sorted_hr
