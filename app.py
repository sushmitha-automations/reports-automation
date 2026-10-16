"""
Workforce & Turnstile Reports Studio.
A sleek, streamlined Streamlit application featuring a single unified multi-file uploader
with zero size restrictions, live base file validation, and persistent multi-report downloads
for Attendance Adherence, Punch Adherence, and Onboarding Reports.
"""
import os
import glob
import time
import io
import zipfile
from datetime import datetime, date
import streamlit as st
import polars as pl
import openpyxl

from orchestrator import (
    run_attendance_adherence,
    run_onboarding_report,
    run_punch_adherence_report
)

# Streamlit Page Config
st.set_page_config(
    page_title="Workforce & Turnstile Reports Studio",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Workspace Paths
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_STAGING_DIR = os.path.join(ROOT_DIR, "Uploaded_Base_Files")
DEFAULT_OUT_DIR = os.path.join(ROOT_DIR, "Reports")

os.makedirs(UPLOAD_STAGING_DIR, exist_ok=True)
os.makedirs(DEFAULT_OUT_DIR, exist_ok=True)

# Session State Initialization
if 'generated_reports' not in st.session_state:
    st.session_state['generated_reports'] = {}

# Custom High-Finish CSS
st.markdown("""
<style>
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1240px;
    }
    .header-title {
        font-size: 2.1rem;
        font-weight: 800;
        color: #1F497D;
        letter-spacing: -0.5px;
        margin-bottom: 0.1rem;
    }
    .header-sub {
        font-size: 0.95rem;
        color: #64748B;
        margin-bottom: 1.5rem;
    }
    .status-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1.1rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        margin-bottom: 0.8rem;
    }
    .badge-ready {
        background-color: #DEF7EC;
        color: #03543F;
        padding: 0.25rem 0.65rem;
        border-radius: 9999px;
        font-weight: 700;
        font-size: 0.8rem;
        display: inline-block;
    }
    .badge-missing {
        background-color: #FDE8E8;
        color: #9B1C1C;
        padding: 0.25rem 0.65rem;
        border-radius: 9999px;
        font-weight: 700;
        font-size: 0.8rem;
        display: inline-block;
    }
    .badge-optional {
        background-color: #EDF2F7;
        color: #4A5568;
        padding: 0.25rem 0.65rem;
        border-radius: 9999px;
        font-weight: 600;
        font-size: 0.8rem;
        display: inline-block;
    }
    .file-name-tag {
        font-family: monospace;
        font-size: 0.85rem;
        color: #1E293B;
        font-weight: 600;
    }
    .file-size-tag {
        font-size: 0.78rem;
        color: #94A3B8;
    }
    .empty-state-box {
        background-color: #F8FAFC;
        border: 2px dashed #CBD5E1;
        border-radius: 8px;
        padding: 1.5rem;
        text-align: center;
        color: #64748B;
        margin-bottom: 1rem;
    }
    .download-card {
        background-color: #F0FDF4;
        border: 1px solid #BBF7D0;
        border-radius: 8px;
        padding: 1rem;
        margin-bottom: 0.8rem;
    }
</style>
""", unsafe_allow_html=True)

# App Header
st.markdown('<div class="header-title">⚡ Workforce & Turnstile Reports Studio</div>', unsafe_allow_html=True)
st.markdown('<div class="header-sub">High-efficiency reconciliation engine powered by Polars. Upload base files, verify status with the validator, and generate executive reports.</div>', unsafe_allow_html=True)

# ==============================================================================
# SECTION 1: UNIFIED SINGLE MULTI-FILE UPLOADER
# ==============================================================================

st.markdown("### 1. Upload Base Files")

source_mode = st.radio(
    "Choose Base Files Source:",
    ["📤 Upload Files (Drag & Drop / Browse)", "📁 Use Local Folder"],
    horizontal=True
)

active_source_dir = None
active_files_list = []

if source_mode == "📤 Upload Files (Drag & Drop / Browse)":
    uploaded_files = st.file_uploader(
        "Upload all base files together (.csv, .xlsx) — No size restriction",
        type=["csv", "xlsx"],
        accept_multiple_files=True,
        help="Select or drag & drop base files: Master User CSVs, Headcounts, Biometric Punches, Chroma Masters, Facility Master, Onboarding Funnel."
    )

    if uploaded_files:
        for f in glob.glob(os.path.join(UPLOAD_STAGING_DIR, "*")):
            try:
                os.remove(f)
            except Exception:
                pass

        for uf in uploaded_files:
            save_path = os.path.join(UPLOAD_STAGING_DIR, uf.name)
            with open(save_path, "wb") as f:
                f.write(uf.getbuffer())

        active_source_dir = UPLOAD_STAGING_DIR
        active_files_list = [
            {'name': uf.name, 'size': f"{uf.size / (1024 * 1024):.2f} MB", 'path': os.path.join(UPLOAD_STAGING_DIR, uf.name)}
            for uf in uploaded_files
        ]
    else:
        for f in glob.glob(os.path.join(UPLOAD_STAGING_DIR, "*")):
            try:
                os.remove(f)
            except Exception:
                pass
        active_source_dir = None
        active_files_list = []

else:
    # Local Folder explicitly chosen
    candidate_folders = [
        "13-Oct-2026/Base Files",
        "Base Files - 30th sep/30th Sep"
    ]
    existing_folders = [
        c for c in candidate_folders
        if os.path.exists(os.path.join(ROOT_DIR, c.replace('/', os.sep)))
    ]

    if existing_folders:
        sel_folder = st.selectbox("Select Local Base Files Folder:", existing_folders)
        target_dir = os.path.join(ROOT_DIR, sel_folder.replace('/', os.sep))
        active_source_dir = target_dir
        for f in os.listdir(target_dir):
            fp = os.path.join(target_dir, f)
            if os.path.isfile(fp):
                size_mb = os.path.getsize(fp) / (1024 * 1024)
                active_files_list.append({'name': f, 'size': f"{size_mb:.2f} MB", 'path': fp})
        st.info(f"📁 Reading files directly from local disk: `{target_dir}`")
    else:
        st.error("No local base files directory found.")
        active_source_dir = None
        active_files_list = []

# ==============================================================================
# SECTION 2: LIVE BASE FILE VALIDATOR
# ==============================================================================

def match_files(files_list):
    """Categorizes files strictly from active_files_list."""
    if not files_list:
        return {}

    def find_file(patterns):
        for f_item in files_list:
            fl = f_item['name'].lower()
            if any(p.lower() in fl for p in patterns):
                return f_item
        return None

    def find_all_files(patterns):
        matched = []
        for f_item in files_list:
            fl = f_item['name'].lower()
            if any(p.lower() in fl for p in patterns):
                matched.append(f_item)
        return matched

    return {
        'facility': find_file(['facility']),
        'user_data': find_all_files(['all_user_data', 'user_data']),
        'attendance': find_file(['valinor attendance', 'attendance']),
        'chroma_lm': find_file(['(lm)', 'valinor (lm)']),
        'chroma_fm': find_file(['(fm)', 'valinor (fm)']),
        'chroma_hl': find_file(['(hl', 'hl & corp', 'corp)']),
        'onroll_hc': find_file(['onroll']),
        'offroll_hc': find_file(['offroll']),
        'onboarding_funnel': find_file(['onboarding_funnel', 'funnel']),
        'least_priority': find_file(['least priority']),
        'migration': find_file(['migration'])
    }

validation_results = match_files(active_files_list)
has_files = bool(active_files_list)

st.markdown("### 2. Base Files Validator")

if not has_files:
    st.markdown("""
    <div class="empty-state-box">
        <h4>📂 No base files uploaded yet</h4>
        <p>Please drop or select your CSV & Excel base files in the upload box above.</p>
    </div>
    """, unsafe_allow_html=True)
else:
    st.caption(f"Currently inspecting **{len(active_files_list)} file(s)**:")

fac = validation_results.get('facility')
users = validation_results.get('user_data', [])
att = validation_results.get('attendance')
c_lm = validation_results.get('chroma_lm')
c_fm = validation_results.get('chroma_fm')
c_hl = validation_results.get('chroma_hl')
onroll = validation_results.get('onroll_hc')
offroll = validation_results.get('offroll_hc')
funnel = validation_results.get('onboarding_funnel')
lp = validation_results.get('least_priority')
mig = validation_results.get('migration')

v_col1, v_col2, v_col3 = st.columns(3)

with v_col1:
    st.markdown("#### 🏢 Attendance Adherence")
    if fac:
        st.markdown(f"✅ **Facility Master:** <span class='file-name-tag'>{fac['name']}</span> <span class='file-size-tag'>({fac['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Facility Master:** <span class='badge-missing'>Not Found (`Valinor Dump-Facility.csv`)</span>", unsafe_allow_html=True)

    if users:
        u_names = ", ".join([f"<span class='file-name-tag'>{u['name']}</span> ({u['size']})" for u in users])
        st.markdown(f"✅ **Master User DB:** {u_names}", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Master User DB:** <span class='badge-missing'>Not Found (`all_user_data*.csv`)</span>", unsafe_allow_html=True)

    if att:
        st.markdown(f"✅ **Biometric Punch:** <span class='file-name-tag'>{att['name']}</span> <span class='file-size-tag'>({att['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Biometric Punch:** <span class='badge-missing'>Not Found (`Valinor Attendance.xlsx / .csv`)</span>", unsafe_allow_html=True)

    if c_lm:
        st.markdown(f"✅ **Chroma (LM):** <span class='file-name-tag'>{c_lm['name']}</span> <span class='file-size-tag'>({c_lm['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Chroma (LM):** <span class='badge-missing'>Not Found (`Valinor (LM)*.xlsx`)</span>", unsafe_allow_html=True)

    if c_fm:
        st.markdown(f"✅ **Chroma (FM):** <span class='file-name-tag'>{c_fm['name']}</span> <span class='file-size-tag'>({c_fm['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Chroma (FM):** <span class='badge-missing'>Not Found (`Valinor (FM)*.xlsx`)</span>", unsafe_allow_html=True)

    if c_hl:
        st.markdown(f"✅ **Chroma (HL):** <span class='file-name-tag'>{c_hl['name']}</span> <span class='file-size-tag'>({c_hl['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Chroma (HL):** <span class='badge-missing'>Not Found (`Valinor (HL)*.xlsx`)</span>", unsafe_allow_html=True)

with v_col2:
    st.markdown("#### 📌 Punch Adherence")
    if onroll:
        st.markdown(f"✅ **Onroll Headcount:** <span class='file-name-tag'>{onroll['name']}</span> <span class='file-size-tag'>({onroll['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Onroll Headcount:** <span class='badge-missing'>Not Found (`Onroll HC.xlsx`)</span>", unsafe_allow_html=True)

    if offroll:
        st.markdown(f"✅ **Offroll Headcount:** <span class='file-name-tag'>{offroll['name']}</span> <span class='file-size-tag'>({offroll['size']})</span>", unsafe_allow_html=True)
    elif c_lm or c_fm or c_hl:
        valinor_loaded = [c['name'] for c in [c_lm, c_fm, c_hl] if c]
        st.markdown(f"✅ **Offroll Headcount:** <span class='badge-ready'>Active ({len(valinor_loaded)} Valinor base files)</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Offroll Headcount:** <span class='badge-missing'>Not Found (`Offroll HC.xlsx` or `Valinor base files`)</span>", unsafe_allow_html=True)

    if att:
        st.markdown(f"✅ **Biometric Punch:** <span class='file-name-tag'>{att['name']}</span> <span class='file-size-tag'>({att['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Biometric Punch:** <span class='badge-missing'>Not Found (`Valinor Attendance.xlsx / .csv`)</span>", unsafe_allow_html=True)

    if fac:
        st.markdown(f"✅ **Facility Master:** <span class='file-name-tag'>{fac['name']}</span> <span class='file-size-tag'>({fac['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Facility Master:** <span class='badge-missing'>Not Found (`Valinor Dump-Facility.csv`)</span>", unsafe_allow_html=True)

    if c_lm or c_fm or c_hl:
        ch_found = [c['name'] for c in [c_lm, c_fm, c_hl] if c]
        st.markdown(f"✅ **Chroma Enrichment:** <span class='file-name-tag'>{len(ch_found)} loaded</span>", unsafe_allow_html=True)
    else:
        st.markdown("ℹ️ **Chroma Enrichment:** <span class='badge-optional'>Optional (Roles fallback to Associate)</span>", unsafe_allow_html=True)

with v_col3:
    st.markdown("#### 📝 Onboarding Report")
    if funnel:
        st.markdown(f"✅ **Onboarding Funnel:** <span class='file-name-tag'>{funnel['name']}</span> <span class='file-size-tag'>({funnel['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Onboarding Funnel:** <span class='badge-missing'>Not Found (`onboarding_funnel*.csv`)</span>", unsafe_allow_html=True)

    if lp:
        st.markdown(f"✅ **Historical Base:** <span class='file-name-tag'>{lp['name']}</span> <span class='file-size-tag'>({lp['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("ℹ️ **Historical Base:** <span class='badge-optional'>Optional (`Least Priority Base*.xlsx`)</span>", unsafe_allow_html=True)

    if mig:
        st.markdown(f"✅ **Migration Exclusions:** <span class='file-name-tag'>{mig['name']}</span> <span class='file-size-tag'>({mig['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("ℹ️ **Migration Exclusions:** <span class='badge-optional'>None detected (Optional)</span>", unsafe_allow_html=True)

# Status Badges
att_ready = bool(fac and users and att and c_lm and c_fm and c_hl)
punch_ready = bool(onroll and (offroll or (c_lm and c_hl and c_fm)) and fac and att)
onb_ready = bool(funnel)

st.markdown("<br>", unsafe_allow_html=True)
col_stat1, col_stat2, col_stat3 = st.columns(3)

with col_stat1:
    if not has_files:
        st.markdown("<span class='badge-missing'>🔴 Attendance: NO FILES</span>", unsafe_allow_html=True)
    elif att_ready:
        st.markdown("<span class='badge-ready'>🟢 Attendance Adherence: READY</span>", unsafe_allow_html=True)
    else:
        st.markdown("<span class='badge-missing'>🔴 Attendance: MISSING FILES</span>", unsafe_allow_html=True)

with col_stat2:
    if not has_files:
        st.markdown("<span class='badge-missing'>🔴 Punch Adherence: NO FILES</span>", unsafe_allow_html=True)
    elif punch_ready:
        st.markdown("<span class='badge-ready'>🟢 Punch Adherence: READY</span>", unsafe_allow_html=True)
    else:
        st.markdown("<span class='badge-missing'>🔴 Punch Adherence: MISSING FILES</span>", unsafe_allow_html=True)

with col_stat3:
    if not has_files:
        st.markdown("<span class='badge-missing'>🔴 Onboarding: NO FILES</span>", unsafe_allow_html=True)
    elif onb_ready:
        st.markdown("<span class='badge-ready'>🟢 Onboarding Report: READY</span>", unsafe_allow_html=True)
    else:
        st.markdown("<span class='badge-missing'>🔴 Onboarding: MISSING FUNNEL CSV</span>", unsafe_allow_html=True)

# ==============================================================================
# SECTION 3: REPORTING PERIOD CONFIGURATION (CALENDAR FORMAT)
# ==============================================================================

st.markdown("---")
st.markdown("### 3. Reporting Period")

def format_calendar_ordinal_label(dt: date) -> str:
    """Formats date into standard report label e.g., '10th Oct', '30th Sep', '14th Oct'."""
    day = dt.day
    if 11 <= (day % 100) <= 13:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    return f"{day}{suffix} {dt.strftime('%b')}"

# Determine smart default date based on detected files
if onroll:
    initial_cal_date = date(2026, 10, 10)
else:
    initial_cal_date = date(2026, 9, 30)

cal_col1, cal_col2 = st.columns([1, 1])

with cal_col1:
    picked_date = st.date_input(
        "📅 Select Cutoff Date (Calendar Picker)",
        value=initial_cal_date,
        format="YYYY-MM-DD",
        help="Select the exact reporting date directly from the calendar."
    )

computed_label = format_calendar_ordinal_label(picked_date)

with cal_col2:
    report_label = st.text_input(
        "🏷️ Report Date Label (Calendar Formatted)",
        value=computed_label,
        help="Automatically synced with your calendar selection. Used in generated report filenames (e.g., 'Punch Adherence Report - 10th Oct.xlsx')."
    )

sel_day = picked_date.day
sel_month = picked_date.month
sel_year = picked_date.year

st.caption(f"ℹ️ Configured Reporting Cutoff: **Day {sel_day}**, **Month {sel_month} ({picked_date.strftime('%B')})**, **Year {sel_year}** • Report Filename Label: **{report_label}**")


# ==============================================================================
# SECTION 4: EXECUTION ACTIONS
# ==============================================================================

st.markdown("---")
st.markdown("### 4. Generate Reports")

btn_run_all = st.button("⚡ Generate All 3 Reports", use_container_width=True, type="primary")

btn_c1, btn_c2, btn_c3 = st.columns(3)
with btn_c1:
    btn_run_att = st.button("📋 Attendance Report Only", use_container_width=True)
with btn_c2:
    btn_run_punch = st.button("📌 Punch Adherence Only", use_container_width=True)
with btn_c3:
    btn_run_onb = st.button("📝 Onboarding Report Only", use_container_width=True)

progress_placeholder = st.empty()
status_placeholder = st.empty()

# Execution Handling
if btn_run_all:
    if not active_source_dir or not (att_ready or punch_ready or onb_ready):
        st.error("Cannot generate reports: None of the reports have their required base files uploaded. Please upload them above.")
    else:
        p_bar = progress_placeholder.progress(0.0, text="Initializing report generation pipeline...")
        status_placeholder.info("⏳ Processing reports with Polars reconciliation engine...")
        t_start_all = time.time()
        successes = []

        try:
            # 1. Attendance Adherence (0.00 -> 0.35)
            if att_ready:
                def cb_att(pct: float, msg: str):
                    p_bar.progress(min(0.35, max(0.0, pct * 0.35)), text=f"📊 [1/3 Attendance] {msg}")

                t0 = time.time()
                att_out = run_attendance_adherence(
                    active_source_dir, DEFAULT_OUT_DIR, report_label,
                    int(sel_year), int(sel_month), int(sel_day),
                    progress_callback=cb_att
                )
                t_dur = time.time() - t0
                with open(att_out, "rb") as f:
                    st.session_state['generated_reports']['attendance'] = {
                        'path': att_out, 'filename': os.path.basename(att_out),
                        'bytes': f.read(), 'time': t_dur
                    }
                successes.append(f"Attendance Adherence ({t_dur:.1f}s)")

            # 2. Punch Adherence (0.35 -> 0.70)
            if punch_ready:
                def cb_punch(pct: float, msg: str):
                    p_bar.progress(min(0.70, max(0.35, 0.35 + pct * 0.35)), text=f"📌 [2/3 Punch Adherence] {msg}")

                t0 = time.time()
                punch_out = run_punch_adherence_report(
                    active_source_dir, DEFAULT_OUT_DIR, report_label,
                    int(sel_year), int(sel_month), int(sel_day),
                    progress_callback=cb_punch
                )
                t_dur = time.time() - t0
                with open(punch_out, "rb") as f:
                    st.session_state['generated_reports']['punch'] = {
                        'path': punch_out, 'filename': os.path.basename(punch_out),
                        'bytes': f.read(), 'time': t_dur
                    }
                successes.append(f"Punch Adherence ({t_dur:.1f}s)")

            # 3. Onboarding Report (0.70 -> 1.00)
            if onb_ready:
                def cb_onb(pct: float, msg: str):
                    p_bar.progress(min(1.00, max(0.70, 0.70 + pct * 0.30)), text=f"📝 [3/3 Onboarding] {msg}")

                t0 = time.time()
                onb_out = run_onboarding_report(
                    active_source_dir, DEFAULT_OUT_DIR, report_label,
                    int(sel_year), int(sel_month), int(sel_day),
                    progress_callback=cb_onb
                )
                t_dur = time.time() - t0
                with open(onb_out, "rb") as f:
                    st.session_state['generated_reports']['onboarding'] = {
                        'path': onb_out, 'filename': os.path.basename(onb_out),
                        'bytes': f.read(), 'time': t_dur
                    }
                successes.append(f"Onboarding ({t_dur:.1f}s)")

            p_bar.progress(1.0, text="🎉 Report generation completed!")
            t_total_all = time.time() - t_start_all
            status_placeholder.success(
                f"✅ Completed in {t_total_all:.1f}s: {', '.join(successes)}! Download below."
            )
        except Exception as err:
            status_placeholder.error(f"Execution Error: {err}")
            st.exception(err)

elif btn_run_att:
    if not active_source_dir or not att_ready:
        st.error("Cannot run Attendance Adherence: Base files missing above. Please upload them first.")
    else:
        p_bar = progress_placeholder.progress(0.0, text="Initializing Attendance pipeline...")
        status_placeholder.info("⏳ Processing Attendance Adherence Report with Polars...")
        t_start = time.time()
        try:
            def cb_att(pct: float, msg: str):
                p_bar.progress(min(1.0, max(0.0, pct)), text=f"📊 [Attendance] {msg}")

            att_out = run_attendance_adherence(
                active_source_dir, DEFAULT_OUT_DIR, report_label,
                int(sel_year), int(sel_month), int(sel_day),
                progress_callback=cb_att
            )
            p_bar.progress(1.0, text="🎉 Attendance Adherence Report generated successfully!")
            t_total = time.time() - t_start
            with open(att_out, "rb") as f:
                st.session_state['generated_reports']['attendance'] = {
                    'path': att_out, 'filename': os.path.basename(att_out),
                    'bytes': f.read(), 'time': t_total
                }
            status_placeholder.success(f"✅ Attendance Adherence Report generated in {t_total:.1f}s! Download below.")
        except Exception as err:
            status_placeholder.error(f"Execution Error: {err}")
            st.exception(err)

elif btn_run_punch:
    if not active_source_dir or not punch_ready:
        st.error("Cannot run Punch Adherence: Onroll/Offroll Headcounts, Punch or Facility Master missing above.")
    else:
        p_bar = progress_placeholder.progress(0.0, text="Initializing Punch Adherence pipeline...")
        status_placeholder.info("⏳ Processing Punch Adherence Report with Polars...")
        t_start = time.time()
        try:
            def cb_punch(pct: float, msg: str):
                p_bar.progress(min(1.0, max(0.0, pct)), text=f"📌 [Punch Adherence] {msg}")

            punch_out = run_punch_adherence_report(
                active_source_dir, DEFAULT_OUT_DIR, report_label,
                int(sel_year), int(sel_month), int(sel_day),
                progress_callback=cb_punch
            )
            p_bar.progress(1.0, text="🎉 Punch Adherence Report generated successfully!")
            t_total = time.time() - t_start
            with open(punch_out, "rb") as f:
                st.session_state['generated_reports']['punch'] = {
                    'path': punch_out, 'filename': os.path.basename(punch_out),
                    'bytes': f.read(), 'time': t_total
                }
            status_placeholder.success(f"✅ Punch Adherence Report generated in {t_total:.1f}s! Download below.")
        except Exception as err:
            status_placeholder.error(f"Execution Error: {err}")
            st.exception(err)

elif btn_run_onb:
    if not active_source_dir or not onb_ready:
        st.error("Cannot run Onboarding Report: Onboarding funnel CSV is missing above.")
    else:
        p_bar = progress_placeholder.progress(0.0, text="Initializing Onboarding pipeline...")
        status_placeholder.info("⏳ Processing Onboarding Report with Polars...")
        t_start = time.time()
        try:
            def cb_onb(pct: float, msg: str):
                p_bar.progress(min(1.0, max(0.0, pct)), text=f"📝 [Onboarding] {msg}")

            onb_out = run_onboarding_report(
                active_source_dir, DEFAULT_OUT_DIR, report_label,
                int(sel_year), int(sel_month), int(sel_day),
                progress_callback=cb_onb
            )
            p_bar.progress(1.0, text="🎉 Onboarding Report generated successfully!")
            t_total = time.time() - t_start
            with open(onb_out, "rb") as f:
                st.session_state['generated_reports']['onboarding'] = {
                    'path': onb_out, 'filename': os.path.basename(onb_out),
                    'bytes': f.read(), 'time': t_total
                }
            status_placeholder.success(f"✅ Onboarding Report generated in {t_total:.1f}s! Download below.")
        except Exception as err:
            status_placeholder.error(f"Execution Error: {err}")
            st.exception(err)

# ==============================================================================
# SECTION 5: PERSISTENT DOWNLOAD CENTER
# ==============================================================================

reports = st.session_state.get('generated_reports', {})

if reports:
    st.markdown("---")
    st.markdown("### 5. Download Generated Reports")
    st.caption("All generated reports remain available here persistently. Downloading one report does not hide any other.")

    # Show single ZIP option if 2 or more reports exist
    if len(reports) >= 2:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for rep_k, rep_data in reports.items():
                zip_file.writestr(rep_data['filename'], rep_data['bytes'])
        zip_buffer.seek(0)

        st.download_button(
            label=f"📦 Download All {len(reports)} Reports (ZIP Archive)",
            data=zip_buffer,
            file_name=f"Workforce_Reports_{report_label}.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary",
            key="dl_all_zip"
        )
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    col_dl1, col_dl2, col_dl3 = st.columns(3)

    with col_dl1:
        if 'attendance' in reports:
            rep = reports['attendance']
            size_mb = len(rep['bytes']) / (1024 * 1024)
            st.markdown(f"""
            <div class="download-card">
                <b>📊 Attendance Adherence Report</b><br>
                <span class="file-name-tag">{rep['filename']}</span> ({size_mb:.2f} MB)<br>
                <small style="color:#64748B;">Generated in {rep['time']:.1f}s • Saved to disk</small>
            </div>
            """, unsafe_allow_html=True)
            st.download_button(
                label=f"📥 Download {rep['filename']}",
                data=rep['bytes'],
                file_name=rep['filename'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_att_btn"
            )

    with col_dl2:
        if 'punch' in reports:
            rep = reports['punch']
            size_mb = len(rep['bytes']) / (1024 * 1024)
            st.markdown(f"""
            <div class="download-card">
                <b>📌 Punch Adherence Report</b><br>
                <span class="file-name-tag">{rep['filename']}</span> ({size_mb:.2f} MB)<br>
                <small style="color:#64748B;">Generated in {rep['time']:.1f}s • Saved to disk</small>
            </div>
            """, unsafe_allow_html=True)
            st.download_button(
                label=f"📥 Download {rep['filename']}",
                data=rep['bytes'],
                file_name=rep['filename'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_punch_btn"
            )

    with col_dl3:
        if 'onboarding' in reports:
            rep = reports['onboarding']
            size_mb = len(rep['bytes']) / (1024 * 1024)
            st.markdown(f"""
            <div class="download-card">
                <b>📋 Onboarding Report</b><br>
                <span class="file-name-tag">{rep['filename']}</span> ({size_mb:.2f} MB)<br>
                <small style="color:#64748B;">Generated in {rep['time']:.1f}s • Saved to disk</small>
            </div>
            """, unsafe_allow_html=True)
            st.download_button(
                label=f"📥 Download {rep['filename']}",
                data=rep['bytes'],
                file_name=rep['filename'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_onb_btn"
            )

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🗑️ Reset / Clear Downloaded Reports", help="Clears generated reports from the download center."):
        st.session_state['generated_reports'] = {}
        st.rerun()
