"""
Workforce & Turnstile Reports Studio.
A sleek, streamlined Streamlit application featuring a single unified multi-file uploader
with zero size restrictions, live base file validation, and persistent multi-report downloads.
"""
import os
import glob
import time
import io
import zipfile
from datetime import datetime
import streamlit as st
import polars as pl
import openpyxl

from orchestrator import run_attendance_adherence, run_onboarding_report

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
LOCAL_SAMPLE_DIR = os.path.join(ROOT_DIR, "Base Files", "30th Sep")
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
        max-width: 1200px;
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
st.markdown('<div class="header-sub">High-efficiency reconciliation engine powered by Polars. Upload base files, verify status with the validator, and generate audit reports.</div>', unsafe_allow_html=True)

# ==============================================================================
# SECTION 1: UNIFIED SINGLE MULTI-FILE UPLOADER
# ==============================================================================

st.markdown("### 1. Upload Base Files")

source_mode = st.radio(
    "Choose Base Files Source:",
    ["📤 Upload Files (Drag & Drop / Browse)", "📁 Use Local Folder (Base Files/30th Sep)"],
    horizontal=True
)

active_source_dir = None
active_files_list = []

if source_mode == "📤 Upload Files (Drag & Drop / Browse)":
    uploaded_files = st.file_uploader(
        "Upload all base files together (.csv, .xlsx) — No size restriction",
        type=["csv", "xlsx"],
        accept_multiple_files=True,
        help="Select or drag & drop all base files: Master User CSVs, Biometric Punches, Chroma Masters, Facility Master, Onboarding Funnel."
    )

    if uploaded_files:
        # User explicitly provided files
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
        # Uploader is empty -> clear staging and show empty state
        for f in glob.glob(os.path.join(UPLOAD_STAGING_DIR, "*")):
            try:
                os.remove(f)
            except Exception:
                pass
        active_source_dir = None
        active_files_list = []

else:
    # Local Folder explicitly chosen
    if os.path.exists(LOCAL_SAMPLE_DIR):
        active_source_dir = LOCAL_SAMPLE_DIR
        for f in os.listdir(LOCAL_SAMPLE_DIR):
            fp = os.path.join(LOCAL_SAMPLE_DIR, f)
            if os.path.isfile(fp):
                size_mb = os.path.getsize(fp) / (1024 * 1024)
                active_files_list.append({'name': f, 'size': f"{size_mb:.2f} MB", 'path': fp})
        st.info(f"📁 Reading files directly from local disk: `{LOCAL_SAMPLE_DIR}`")
    else:
        st.error(f"Local sample directory not found: {LOCAL_SAMPLE_DIR}")
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

v_col1, v_col2 = st.columns(2)

with v_col1:
    st.markdown("#### 🏢 Attendance Adherence Requirements")

    fac = validation_results.get('facility')
    if fac:
        st.markdown(f"✅ **Facility Master:** <span class='file-name-tag'>{fac['name']}</span> <span class='file-size-tag'>({fac['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Facility Master:** <span class='badge-missing'>Not Uploaded (`Valinor Dump-Facility.csv`)</span>", unsafe_allow_html=True)

    users = validation_results.get('user_data', [])
    if users:
        u_names = ", ".join([f"<span class='file-name-tag'>{u['name']}</span> ({u['size']})" for u in users])
        st.markdown(f"✅ **Master User DB:** {u_names}", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Master User DB:** <span class='badge-missing'>Not Uploaded (`all_user_data*.csv`)</span>", unsafe_allow_html=True)

    att = validation_results.get('attendance')
    if att:
        st.markdown(f"✅ **Biometric Attendance:** <span class='file-name-tag'>{att['name']}</span> <span class='file-size-tag'>({att['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Biometric Attendance:** <span class='badge-missing'>Not Uploaded (`Valinor Attendance.csv`)</span>", unsafe_allow_html=True)

    c_lm = validation_results.get('chroma_lm')
    c_fm = validation_results.get('chroma_fm')
    c_hl = validation_results.get('chroma_hl')

    if c_lm:
        st.markdown(f"✅ **Chroma (LM):** <span class='file-name-tag'>{c_lm['name']}</span> <span class='file-size-tag'>({c_lm['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Chroma (LM):** <span class='badge-missing'>Not Uploaded (`Valinor (LM)*.xlsx`)</span>", unsafe_allow_html=True)

    if c_fm:
        st.markdown(f"✅ **Chroma (FM):** <span class='file-name-tag'>{c_fm['name']}</span> <span class='file-size-tag'>({c_fm['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Chroma (FM):** <span class='badge-missing'>Not Uploaded (`Valinor (FM)*.xlsx`)</span>", unsafe_allow_html=True)

    if c_hl:
        st.markdown(f"✅ **Chroma (HL & Corp):** <span class='file-name-tag'>{c_hl['name']}</span> <span class='file-size-tag'>({c_hl['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Chroma (HL & Corp):** <span class='badge-missing'>Not Uploaded (`Valinor (HL & Corp)*.xlsx`)</span>", unsafe_allow_html=True)

with v_col2:
    st.markdown("#### 📝 Onboarding Report Requirements")

    funnel = validation_results.get('onboarding_funnel')
    if funnel:
        st.markdown(f"✅ **Onboarding Funnel:** <span class='file-name-tag'>{funnel['name']}</span> <span class='file-size-tag'>({funnel['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("❌ **Onboarding Funnel:** <span class='badge-missing'>Not Uploaded (`onboarding_funnel*.csv`)</span>", unsafe_allow_html=True)

    lp = validation_results.get('least_priority')
    if lp:
        st.markdown(f"✅ **Historical Base:** <span class='file-name-tag'>{lp['name']}</span> <span class='file-size-tag'>({lp['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("ℹ️ **Historical Base:** <span class='badge-optional'>Optional (`Least Priority Base*.xlsx`)</span>", unsafe_allow_html=True)

    mig = validation_results.get('migration')
    if mig:
        st.markdown(f"✅ **Migration Exclusions:** <span class='file-name-tag'>{mig['name']}</span> <span class='file-size-tag'>({mig['size']})</span>", unsafe_allow_html=True)
    else:
        st.markdown("ℹ️ **Migration Exclusions:** <span class='badge-optional'>None detected (Skipping migration exclusions)</span>", unsafe_allow_html=True)

# Status Badges
att_ready = bool(fac and users and att and c_lm and c_fm and c_hl)
onb_ready = bool(funnel)

st.markdown("<br>", unsafe_allow_html=True)
col_stat1, col_stat2 = st.columns(2)
with col_stat1:
    if not has_files:
        st.markdown("<span class='badge-missing'>🔴 Attendance Adherence: NO FILES UPLOADED</span>", unsafe_allow_html=True)
    elif att_ready:
        st.markdown("<span class='badge-ready'>🟢 Attendance Adherence: READY TO GENERATE</span>", unsafe_allow_html=True)
    else:
        st.markdown("<span class='badge-missing'>🔴 Attendance Adherence: MISSING REQUIRED FILES</span>", unsafe_allow_html=True)

with col_stat2:
    if not has_files:
        st.markdown("<span class='badge-missing'>🔴 Onboarding Report: NO FILES UPLOADED</span>", unsafe_allow_html=True)
    elif onb_ready:
        st.markdown("<span class='badge-ready'>🟢 Onboarding Report: READY TO GENERATE</span>", unsafe_allow_html=True)
    else:
        st.markdown("<span class='badge-missing'>🔴 Onboarding Report: MISSING FUNNEL CSV</span>", unsafe_allow_html=True)

# ==============================================================================
# SECTION 3: REPORTING PERIOD CONFIGURATION
# ==============================================================================

st.markdown("---")
st.markdown("### 3. Reporting Period")

rc1, rc2, rc3, rc4 = st.columns([2, 1, 1, 1])

with rc1:
    report_label = st.text_input("Report Date Label", value="30th Sep", help="Used in file naming: e.g., 'Attendance Adherence Report - 30th Sep.xlsx'")
with rc2:
    sel_day = st.number_input("Cutoff Day", min_value=1, max_value=31, value=30, help="D-0 or D-1 cutoff day")
with rc3:
    sel_month = st.number_input("Month", min_value=1, max_value=12, value=9)
with rc4:
    sel_year = st.number_input("Year", min_value=2024, max_value=2030, value=2026)

# ==============================================================================
# SECTION 4: EXECUTION ACTIONS
# ==============================================================================

st.markdown("---")
st.markdown("### 4. Generate Reports")

btn_col1, btn_col2, btn_col3 = st.columns(3)

with btn_col1:
    btn_run_both = st.button("⚡ Generate Both Reports", use_container_width=True, type="primary")

with btn_col2:
    btn_run_att = st.button("📋 Generate Attendance Report Only", use_container_width=True)

with btn_col3:
    btn_run_onb = st.button("📝 Generate Onboarding Report Only", use_container_width=True)

progress_placeholder = st.empty()
status_placeholder = st.empty()

# Execution Handling
if btn_run_both:
    if not active_source_dir or not att_ready or not onb_ready:
        missing = []
        if not att_ready:
            missing.append("Attendance Adherence base files")
        if not onb_ready:
            missing.append("Onboarding Funnel CSV")
        st.error(f"Cannot generate both reports: Missing {', '.join(missing)} above. Please upload them first.")
    else:
        p_bar = progress_placeholder.progress(0.0, text="Initializing report generation pipeline...")
        status_placeholder.info("⏳ Processing Attendance & Onboarding Reports with Polars...")
        t_start_all = time.time()
        try:
            # 1. Attendance Adherence (maps 0.0 -> 0.55)
            def cb_att(pct: float, msg: str):
                val = min(0.55, max(0.0, pct * 0.55))
                p_bar.progress(val, text=f"📊 [1/2 Attendance] {msg}")

            t_start_att = time.time()
            att_out = run_attendance_adherence(
                active_source_dir, DEFAULT_OUT_DIR, report_label,
                int(sel_year), int(sel_month), int(sel_day),
                progress_callback=cb_att
            )
            t_dur_att = time.time() - t_start_att

            with open(att_out, "rb") as f:
                att_bytes = f.read()

            st.session_state['generated_reports']['attendance'] = {
                'path': att_out,
                'filename': os.path.basename(att_out),
                'bytes': att_bytes,
                'time': t_dur_att
            }

            # 2. Onboarding Report (maps 0.55 -> 1.00)
            def cb_onb(pct: float, msg: str):
                val = min(1.0, max(0.55, 0.55 + pct * 0.45))
                p_bar.progress(val, text=f"📝 [2/2 Onboarding] {msg}")

            t_start_onb = time.time()
            onb_out = run_onboarding_report(
                active_source_dir, DEFAULT_OUT_DIR, report_label,
                int(sel_year), int(sel_month), int(sel_day),
                progress_callback=cb_onb
            )
            t_dur_onb = time.time() - t_start_onb

            with open(onb_out, "rb") as f:
                onb_bytes = f.read()

            st.session_state['generated_reports']['onboarding'] = {
                'path': onb_out,
                'filename': os.path.basename(onb_out),
                'bytes': onb_bytes,
                'time': t_dur_onb
            }

            p_bar.progress(1.0, text="🎉 Both reports generated successfully!")
            t_total_all = time.time() - t_start_all
            status_placeholder.success(
                f"✅ Both reports successfully generated in {t_total_all:.1f}s "
                f"(Attendance: {t_dur_att:.1f}s, Onboarding: {t_dur_onb:.1f}s)! Download them below."
            )
        except Exception as err:
            status_placeholder.error(f"Execution Error: {err}")
            st.exception(err)

elif btn_run_att:
    if not active_source_dir or not att_ready:
        st.error("Cannot run Attendance Adherence Report: Mandatory base files are missing above. Please upload them first.")
    else:
        p_bar = progress_placeholder.progress(0.0, text="Initializing Attendance pipeline...")
        status_placeholder.info("⏳ Processing Attendance Adherence Report with Polars...")
        t_start = time.time()
        try:
            def cb_att(pct: float, msg: str):
                val = min(1.0, max(0.0, pct))
                p_bar.progress(val, text=f"📊 [Attendance] {msg}")

            att_out = run_attendance_adherence(
                active_source_dir, DEFAULT_OUT_DIR, report_label,
                int(sel_year), int(sel_month), int(sel_day),
                progress_callback=cb_att
            )
            p_bar.progress(1.0, text="🎉 Attendance Adherence Report generated successfully!")
            t_total = time.time() - t_start

            with open(att_out, "rb") as f:
                att_bytes = f.read()

            st.session_state['generated_reports']['attendance'] = {
                'path': att_out,
                'filename': os.path.basename(att_out),
                'bytes': att_bytes,
                'time': t_total
            }
            status_placeholder.success(f"✅ Attendance Adherence Report generated in {t_total:.1f} seconds! Download below.")
        except Exception as err:
            status_placeholder.error(f"Execution Error: {err}")
            st.exception(err)

elif btn_run_onb:
    if not active_source_dir or not onb_ready:
        st.error("Cannot run Onboarding Report: Onboarding funnel CSV is missing above. Please upload it first.")
    else:
        p_bar = progress_placeholder.progress(0.0, text="Initializing Onboarding pipeline...")
        status_placeholder.info("⏳ Processing Onboarding Report with Polars...")
        t_start = time.time()
        try:
            def cb_onb(pct: float, msg: str):
                val = min(1.0, max(0.0, pct))
                p_bar.progress(val, text=f"📝 [Onboarding] {msg}")

            onb_out = run_onboarding_report(
                active_source_dir, DEFAULT_OUT_DIR, report_label,
                int(sel_year), int(sel_month), int(sel_day),
                progress_callback=cb_onb
            )
            p_bar.progress(1.0, text="🎉 Onboarding Report generated successfully!")
            t_total = time.time() - t_start

            with open(onb_out, "rb") as f:
                onb_bytes = f.read()

            st.session_state['generated_reports']['onboarding'] = {
                'path': onb_out,
                'filename': os.path.basename(onb_out),
                'bytes': onb_bytes,
                'time': t_total
            }
            status_placeholder.success(f"✅ Onboarding Report generated in {t_total:.1f} seconds! Download below.")
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
    st.caption("All generated reports remain available here. Downloading one report will not hide the other.")

    # Show single ZIP option if multiple reports exist
    if 'attendance' in reports and 'onboarding' in reports:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            zip_file.writestr(reports['attendance']['filename'], reports['attendance']['bytes'])
            zip_file.writestr(reports['onboarding']['filename'], reports['onboarding']['bytes'])
        zip_buffer.seek(0)

        st.download_button(
            label="📦 Download Both Reports (ZIP Archive)",
            data=zip_buffer,
            file_name=f"Workforce_Reports_{report_label}.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary",
            key="dl_both_zip"
        )
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    col_dl1, col_dl2 = st.columns(2)

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
    if st.button("🗑️ Reset / Clear Downloaded Reports", help="Clears the generated reports from the download center."):
        st.session_state['generated_reports'] = {}
        st.rerun()
