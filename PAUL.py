from datetime import datetime
import io
import multiprocessing
import os
import sys
import tempfile
import time
import zipfile

import fitz  # PyMuPDF
from pdf2docx import Converter
import streamlit as st

# ==========================================
# 1. PAGE CONFIG & SKLAUNCHER FLUID UI
# ==========================================
st.set_page_config(
    page_title="PAUL — PDF to Word Converter",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');

    /* Hide default Streamlit top bar & chrome */
    #MainMenu, footer, header, [data-testid="stHeader"] {
        visibility: hidden !important;
        height: 0px !important;
    }

    /* Ambient SKlauncher background */
    html, body, [data-testid="stAppViewContainer"], .stApp {
        background-color: #090b10 !important;
        background-image: radial-gradient(circle at 50% -10%, rgba(225, 29, 72, 0.09), transparent 55%) !important;
        color: #e2e8f0 !important;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }

    code, pre, .mono {
        font-family: 'JetBrains Mono', monospace !important;
    }

    .block-container {
        padding-top: 2.2rem !important;
        padding-bottom: 3.5rem !important;
        max-width: 1040px !important;
    }

    /* Smooth Keyframe Entrance for Tab Panels */
    @keyframes fadeSlideUp {
        from {
            opacity: 0;
            transform: translateY(6px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    [data-baseweb="tab-panel"] {
        animation: fadeSlideUp 0.28s cubic-bezier(0.22, 1, 0.36, 1) forwards !important;
        padding-top: 0.5rem !important;
    }

    /* ========================================================
       SKLAUNCHER HEADER CARD
       ======================================================== */
    .sk-header {
        background: linear-gradient(145deg, rgba(20, 24, 34, 0.88) 0%, rgba(13, 16, 23, 0.92) 100%);
        backdrop-filter: blur(16px);
        border: 1px solid rgba(255, 255, 255, 0.06);
        border-radius: 18px;
        padding: 1.4rem 1.85rem;
        margin-bottom: 1.6rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 16px 40px -12px rgba(0, 0, 0, 0.7);
        transition: border-color 0.35s cubic-bezier(0.22, 1, 0.36, 1),
                    box-shadow 0.35s cubic-bezier(0.22, 1, 0.36, 1);
    }
    .sk-header:hover {
        border-color: rgba(255, 255, 255, 0.12);
        box-shadow: 0 20px 45px -12px rgba(0, 0, 0, 0.85);
    }
    .sk-brand {
        display: flex;
        align-items: center;
        gap: 0.85rem;
    }
    .sk-title {
        font-size: 1.6rem;
        font-weight: 700;
        letter-spacing: -0.03em;
        color: #f8fafc;
        margin: 0;
        line-height: 1;
    }
    .sk-pill {
        background: rgba(225, 29, 72, 0.14);
        color: #fb7185;
        border: 1px solid rgba(225, 29, 72, 0.32);
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.68rem;
        font-weight: 600;
        padding: 0.22rem 0.65rem;
        border-radius: 9999px;
        letter-spacing: 0.06em;
        text-transform: uppercase;
    }
    .sk-subtitle {
        font-size: 0.86rem;
        color: #94a3b8;
        margin: 0.35rem 0 0 0;
    }
    .sk-sys-badge {
        background: rgba(11, 14, 20, 0.75);
        border: 1px solid rgba(255, 255, 255, 0.06);
        border-radius: 12px;
        padding: 0.55rem 0.95rem;
        font-size: 0.75rem;
        color: #94a3b8;
        text-align: right;
        line-height: 1.45;
    }
    .sk-sys-badge span {
        color: #f1f5f9;
        font-weight: 600;
    }

    /* ========================================================
       DETECTED PAGES PILL BAR
       ======================================================== */
    .detector-bar {
        background-color: #11151f;
        border: 1px solid rgba(244, 63, 94, 0.28);
        border-radius: 12px;
        padding: 0.65rem 1.1rem;
        margin-bottom: 1.15rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 0.5rem;
    }
    .detector-label {
        font-size: 0.8rem;
        font-weight: 600;
        color: #f8fafc;
    }
    .detector-chips {
        display: flex;
        flex-wrap: wrap;
        gap: 0.45rem;
    }
    .detector-chip {
        background: rgba(225, 29, 72, 0.14);
        border: 1px solid rgba(225, 29, 72, 0.32);
        color: #fda4af;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.74rem;
        font-weight: 500;
        padding: 0.2rem 0.65rem;
        border-radius: 9999px;
    }

    /* ========================================================
       TRUE ROUNDED PILL TABS
       ======================================================== */
    .stTabs [data-baseweb="tab-list"] {
        background-color: #10131c !important;
        border: 1px solid rgba(255, 255, 255, 0.06) !important;
        border-radius: 9999px !important;
        padding: 5px !important;
        gap: 6px !important;
        width: fit-content !important;
        margin-bottom: 1.15rem !important;
        box-shadow: inset 0 2px 6px rgba(0, 0, 0, 0.45) !important;
    }

    .stTabs [data-baseweb="tab-highlight"],
    .stTabs [data-baseweb="tab-border"] {
        display: none !important;
        height: 0px !important;
        background: transparent !important;
    }

    .stTabs button[data-baseweb="tab"] {
        border-radius: 9999px !important;
        padding: 0.5rem 1.35rem !important;
        height: auto !important;
        min-height: 36px !important;
        background-color: transparent !important;
        border: 1px solid transparent !important;
        color: #94a3b8 !important;
        transition: all 0.28s cubic-bezier(0.22, 1, 0.36, 1) !important;
        margin: 0 !important;
    }

    .stTabs button[data-baseweb="tab"] p,
    .stTabs button[data-baseweb="tab"] div {
        font-size: 0.84rem !important;
        font-weight: 500 !important;
        color: inherit !important;
        margin: 0 !important;
        padding: 0 !important;
        white-space: nowrap !important;
        transition: color 0.25s ease !important;
    }

    .stTabs button[data-baseweb="tab"]:hover {
        color: #f8fafc !important;
        background-color: rgba(255, 255, 255, 0.05) !important;
        transform: translateY(-1px) !important;
    }

    .stTabs button[data-baseweb="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, rgba(225, 29, 72, 0.22) 0%, rgba(159, 18, 57, 0.16) 100%) !important;
        color: #ffffff !important;
        border: 1px solid rgba(244, 63, 94, 0.4) !important;
        box-shadow: 0 4px 16px -2px rgba(225, 29, 72, 0.28),
                    inset 0 1px 0 rgba(255, 255, 255, 0.12) !important;
        transform: translateY(0px) !important;
    }

    .stTabs button[data-baseweb="tab"][aria-selected="true"] p {
        font-weight: 600 !important;
        color: #ffffff !important;
    }

    /* ========================================================
       FILE UPLOADER DROPZONE
       ======================================================== */
    [data-testid="stFileUploader"] {
        margin-bottom: 0.85rem !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        background-color: #11151f !important;
        border: 1.5px dashed rgba(255, 255, 255, 0.1) !important;
        border-radius: 16px !important;
        padding: 1.35rem 1.5rem !important;
        transition: all 0.32s cubic-bezier(0.22, 1, 0.36, 1) !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        background-color: #151a27 !important;
        border-color: rgba(244, 63, 94, 0.55) !important;
        box-shadow: 0 10px 30px -10px rgba(225, 29, 72, 0.18) !important;
        transform: translateY(-1px) !important;
    }
    [data-testid="stFileUploaderDropzone"] button {
        background-color: #1b2232 !important;
        color: #f1f5f9 !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        border-radius: 9999px !important;
        padding: 0.4rem 1.1rem !important;
        font-weight: 500 !important;
        transition: all 0.25s cubic-bezier(0.22, 1, 0.36, 1) !important;
    }
    [data-testid="stFileUploaderDropzone"] button:hover {
        background-color: #e11d48 !important;
        border-color: #fb7185 !important;
        color: #ffffff !important;
        transform: scale(1.02) !important;
    }

    /* ========================================================
       MODE CARDS & SUMMARY GRID
       ======================================================== */
    .mode-card {
        background: linear-gradient(180deg, rgba(19, 23, 34, 0.9) 0%, rgba(15, 18, 27, 0.9) 100%);
        border: 1px solid rgba(255, 255, 255, 0.06);
        border-radius: 16px;
        padding: 1.2rem 1.45rem;
        margin-bottom: 1.1rem;
        transition: all 0.3s cubic-bezier(0.22, 1, 0.36, 1);
    }
    .mode-card:hover {
        border-color: rgba(255, 255, 255, 0.12);
        box-shadow: 0 10px 28px -12px rgba(0, 0, 0, 0.6);
    }
    .mode-card-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 0.35rem;
    }
    .mode-card-title {
        font-size: 0.96rem;
        font-weight: 600;
        color: #f8fafc;
        letter-spacing: -0.01em;
    }
    .mode-card-tag {
        font-size: 0.72rem;
        font-weight: 500;
        color: #cbd5e1;
        background: rgba(255, 255, 255, 0.05);
        padding: 0.2rem 0.65rem;
        border-radius: 9999px;
        border: 1px solid rgba(255, 255, 255, 0.08);
    }
    .mode-card-desc {
        font-size: 0.84rem;
        color: #94a3b8;
        margin: 0;
        line-height: 1.45;
    }

    /* Summary Cards */
    .telemetry-grid {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 0.85rem;
        margin: 1.1rem 0 1.25rem 0;
    }
    .telemetry-cell {
        background-color: #11151f;
        border: 1px solid rgba(255, 255, 255, 0.06);
        border-radius: 14px;
        padding: 0.85rem 1.1rem;
        transition: all 0.28s cubic-bezier(0.22, 1, 0.36, 1);
    }
    .telemetry-cell:hover {
        transform: translateY(-2px);
        background-color: #151a27;
        border-color: rgba(255, 255, 255, 0.13);
        box-shadow: 0 8px 20px -8px rgba(0, 0, 0, 0.55);
    }
    .telemetry-label {
        font-size: 0.7rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #64748b;
        margin-bottom: 0.3rem;
    }
    .telemetry-val {
        font-size: 0.9rem;
        font-weight: 600;
        color: #f1f5f9;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }

    /* ========================================================
       INPUT FIELDS & LAUNCH BUTTONS
       ======================================================== */
    [data-baseweb="input"], [data-baseweb="select"] > div {
        background-color: #11151f !important;
        border-color: rgba(255, 255, 255, 0.08) !important;
        border-radius: 12px !important;
        transition: all 0.25s cubic-bezier(0.22, 1, 0.36, 1) !important;
    }
    [data-baseweb="input"]:hover, [data-baseweb="select"] > div:hover {
        border-color: rgba(255, 255, 255, 0.18) !important;
    }

    .stButton > button[kind="primary"],
    .stDownloadButton > button[kind="primary"] {
        background: linear-gradient(135deg, #e11d48 0%, #be123c 100%) !important;
        color: #ffffff !important;
        border: 1px solid rgba(255, 255, 255, 0.14) !important;
        border-radius: 12px !important;
        height: 46px !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
        letter-spacing: 0.01em !important;
        box-shadow: 0 8px 24px -6px rgba(225, 29, 72, 0.4) !important;
        transition: all 0.28s cubic-bezier(0.22, 1, 0.36, 1) !important;
    }
    .stButton > button[kind="primary"]:hover:not(:disabled),
    .stDownloadButton > button[kind="primary"]:hover:not(:disabled) {
        transform: translateY(-2px) !important;
        box-shadow: 0 12px 28px -4px rgba(225, 29, 72, 0.58) !important;
        filter: brightness(1.08);
    }
    .stButton > button[kind="primary"]:disabled {
        background: #161b26 !important;
        color: #475569 !important;
        border-color: rgba(255, 255, 255, 0.05) !important;
        box-shadow: none !important;
    }

    /* Deliverable Row */
    .deliverable-row {
        background-color: #11151f;
        border: 1px solid rgba(255, 255, 255, 0.06);
        border-radius: 12px;
        padding: 0.75rem 1.2rem;
        margin-bottom: 0.5rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
        transition: all 0.28s cubic-bezier(0.22, 1, 0.36, 1);
    }
    .deliverable-row:hover {
        transform: translateX(3px);
        border-color: rgba(244, 63, 94, 0.35);
        background-color: #151a27;
    }
    .deliverable-name {
        font-weight: 600;
        font-size: 0.88rem;
        color: #f8fafc;
    }
    .deliverable-meta {
        font-size: 0.78rem;
        color: #94a3b8;
    }
</style>
""",
    unsafe_allow_html=True,
)


# ==========================================
# 2. CONVERSION ENGINE & PAGE RESOLVERS
# ==========================================
HOST_CORES = multiprocessing.cpu_count()
AUTO_MP_PAGE_THRESHOLD = 12


def parse_page_expression(expr: str, max_pages: int) -> list[int]:
  pages_set = set()
  if not expr or not expr.strip() or max_pages <= 0:
    return []

  tokens = [t.strip() for t in expr.split(",") if t.strip()]
  for token in tokens:
    if "-" in token:
      parts = token.split("-")
      if (
          len(parts) == 2
          and parts[0].strip().isdigit()
          and parts[1].strip().isdigit()
      ):
        s_p, e_p = int(parts[0].strip()), int(parts[1].strip())
        if s_p > e_p:
          s_p, e_p = e_p, s_p
        for p in range(s_p, e_p + 1):
          if 1 <= p <= max_pages:
            pages_set.add(p - 1)
    elif token.isdigit():
      p = int(token)
      if 1 <= p <= max_pages:
        pages_set.add(p - 1)

  return sorted(pages_set)


def compute_auto_allocation(
    strategy: str, num_pages: int
) -> tuple[bool, int, str]:
  if strategy != "CONTINUOUS":
    return False, 1, "Standard"

  if num_pages >= AUTO_MP_PAGE_THRESHOLD and HOST_CORES > 1:
    optimal_cores = min(HOST_CORES, max(2, min(8, num_pages // 6)))
    return True, optimal_cores, "Fast Mode (Multi-Core)"

  return False, 1, "Standard"


def resolve_target_pages(
    mode: str,
    total_pages: int,
    seq_start: int = 1,
    seq_end: int = 1,
    exclusive_end: bool = False,
    non_seq_pills: list[int] = None,
    non_seq_expr: str = "",
) -> tuple[str, int, int, list[int]]:
  if total_pages <= 0:
    return "CONTINUOUS", 0, 0, []

  non_seq_pills = non_seq_pills or []

  if mode == "All Pages":
    return "CONTINUOUS", 0, total_pages, list(range(total_pages))

  elif mode == "Page Range":
    eff_end_1 = seq_end - 1 if exclusive_end else seq_end
    eff_end_1 = min(eff_end_1, total_pages)
    start_0 = max(0, seq_start - 1)
    end_0 = max(start_0, eff_end_1)
    return "CONTINUOUS", start_0, end_0, list(range(start_0, end_0))

  elif mode == "Specific Pages":
    combined_0 = set(p - 1 for p in non_seq_pills if 1 <= p <= total_pages)
    combined_0.update(parse_page_expression(non_seq_expr, total_pages))
    return "DISCRETE", 0, None, sorted(combined_0)

  else:  # Custom Mix
    eff_end_1 = seq_end - 1 if exclusive_end else seq_end
    eff_end_1 = min(eff_end_1, total_pages)
    start_0 = max(0, seq_start - 1)
    end_0 = max(start_0, eff_end_1)

    combined_0 = set(range(start_0, end_0))
    combined_0.update(p - 1 for p in non_seq_pills if 1 <= p <= total_pages)
    combined_0.update(parse_page_expression(non_seq_expr, total_pages))
    return "DISCRETE", 0, None, sorted(combined_0)


def execute_pdf2docx(
    pdf_path: str,
    docx_path: str,
    strategy: str,
    start_0: int,
    end_0: int,
    pages_0: list[int],
    use_mp: bool,
    cores: int,
):
  cv = Converter(pdf_path)
  try:
    if strategy == "CONTINUOUS":
      cv.convert(
          docx_path,
          start=start_0,
          end=end_0,
          multi_processing=use_mp,
          cpu_count=cores if use_mp else 1,
      )
    else:
      cv.convert(docx_path, pages=pages_0)
  finally:
    cv.close()


# ==========================================
# 3. SESSION STATE & SUMMARY HELPERS
# ==========================================
if "audit_logs" not in st.session_state:
  st.session_state.audit_logs = []
if "deliverables" not in st.session_state:
  st.session_state.deliverables = []


def log_event(status: str, filename: str, details: str):
  ts = datetime.now().strftime("%H:%M:%S")
  st.session_state.audit_logs.append(f"[{ts}] {status} — {filename}: {details}")


def render_summary_strip(
    queue_count: int,
    detected_max_pages: int,
    mode_label: str,
    pages_0: list[int],
    strategy: str,
):
  pages_1 = [p + 1 for p in pages_0]
  if not queue_count:
    files_str = "0 File(s)"
    scope_str = "Upload a PDF first"
    compute_str = "Automatic"
  else:
    files_str = f"{queue_count} File(s) • {detected_max_pages} Page(s)"
    if len(pages_1) > 7:
      scope_str = f"{len(pages_1)} of {detected_max_pages} pages"
    else:
      scope_str = (
          ", ".join(map(str, pages_1)) if pages_1 else "No pages selected"
      )
    _, _, compute_str = compute_auto_allocation(strategy, len(pages_0))

  st.markdown(
      f"""
        <div class="telemetry-grid">
            <div class="telemetry-cell">
                <div class="telemetry-label">Detected PDF</div>
                <div class="telemetry-val">{files_str}</div>
            </div>
            <div class="telemetry-cell">
                <div class="telemetry-label">Mode</div>
                <div class="telemetry-val">{mode_label}</div>
            </div>
            <div class="telemetry-cell">
                <div class="telemetry-label">Pages to Convert</div>
                <div class="telemetry-val">{scope_str}</div>
            </div>
            <div class="telemetry-cell">
                <div class="telemetry-label">Speed</div>
                <div class="telemetry-val">{compute_str}</div>
            </div>
        </div>
        """,
      unsafe_allow_html=True,
  )


def run_batch_pipeline(
    uploaded_files,
    file_page_counts: dict,
    mode: str,
    seq_start: int = 1,
    seq_end: int = 1,
    exclusive_end: bool = False,
    non_seq_pills: list[int] = None,
    non_seq_expr: str = "",
):
  st.session_state.deliverables = []
  progress_bar = st.progress(0, text="Starting conversion...")

  with st.status("Converting your PDF(s) to Word...", expanded=True) as status:
    total = len(uploaded_files)
    for idx, u_file in enumerate(uploaded_files, start=1):
      t0 = time.perf_counter()
      u_file.seek(0)
      raw_bytes = u_file.read()
      page_count = file_page_counts[u_file.name]

      strategy, s_0, e_0, pages_0 = resolve_target_pages(
          mode,
          page_count,
          seq_start,
          seq_end,
          exclusive_end,
          non_seq_pills,
          non_seq_expr,
      )

      if not pages_0:
        log_event(
            "Skipped",
            u_file.name,
            f"Selected pages are outside the document's {page_count} pages.",
        )
        st.warning(
            f"Skipped {u_file.name} — none of the selected pages exist in this"
            " file."
        )
        continue

      use_mp, auto_cores, _ = compute_auto_allocation(strategy, len(pages_0))

      st.write(
          f"Converting **{u_file.name}** ({len(pages_0)} of {page_count}"
          " pages)..."
      )

      with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_in:
        tmp_in.write(raw_bytes)
        tmp_in_path = tmp_in.name

      out_name = os.path.splitext(u_file.name)[0] + ".docx"
      tmp_out_path = tmp_in_path.replace(".pdf", ".docx")

      try:
        execute_pdf2docx(
            pdf_path=tmp_in_path,
            docx_path=tmp_out_path,
            strategy=strategy,
            start_0=s_0,
            end_0=e_0,
            pages_0=pages_0,
            use_mp=use_mp,
            cores=auto_cores,
        )

        with open(tmp_out_path, "rb") as f_docx:
          docx_bytes = f_docx.read()

        dt = round(time.perf_counter() - t0, 2)
        size_kb = round(len(docx_bytes) / 1024, 1)

        st.session_state.deliverables.append({
            "name": out_name,
            "bytes": docx_bytes,
            "elapsed": dt,
            "size_kb": size_kb,
            "pages": f"{len(pages_0)} of {page_count} pages",
        })
        log_event(
            "Done",
            u_file.name,
            f"Saved as {out_name} ({len(pages_0)} pages, {size_kb} KB in"
            f" {dt}s)",
        )
      except Exception as exc:
        log_event("Error", u_file.name, str(exc))
        st.error(f"Could not convert {u_file.name}: {exc}")
      finally:
        if os.path.exists(tmp_in_path):
          os.remove(tmp_in_path)
        if os.path.exists(tmp_out_path):
          os.remove(tmp_out_path)

      progress_bar.progress(
          idx / total, text=f"Converted {idx} of {total} file(s)"
      )

    status.update(
        label=(
            f"Finished! Converted {len(st.session_state.deliverables)} of"
            f" {total} file(s)."
        ),
        state="complete",
        expanded=False,
    )


# ==========================================
# 4. HEADER & SIMPLE NAVIGATION
# ==========================================
st.markdown(
    """
    <div class="sk-header">
        <div>
            <div class="sk-brand">
                <h1 class="sk-title">PAUL</h1>
                <span class="sk-pill">PDF to Word</span>
            </div>
            <p class="sk-subtitle">Fast, clean PDF to editable Word (.docx) converter</p>
        </div>
        <div class="sk-sys-badge">
            Page Detector: <span>Active</span><br>
            Format: <span>Word (.docx)</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

nav_workspace, nav_inspector, nav_telemetry = st.tabs([
    "Convert",
    "Page Preview",
    "Activity Log",
])

# ------------------------------------------
# TOP TAB 1: CONVERT
# ------------------------------------------
with nav_workspace:
  uploaded_files = st.file_uploader(
      "Upload PDF files",
      type=["pdf"],
      accept_multiple_files=True,
      label_visibility="collapsed",
  )

  # --- AUTOMATIC PAGE DETECTOR ---
  queue_count = len(uploaded_files) if uploaded_files else 0
  file_page_counts = {}
  detected_max_pages = 0

  if uploaded_files:
    chips_html = ""
    for uf in uploaded_files:
      uf.seek(0)
      with fitz.open(stream=uf.read(), filetype="pdf") as doc:
        p_len = len(doc)
        file_page_counts[uf.name] = p_len
      chips_html += (
          f'<span class="detector-chip">{uf.name} • {p_len}'
          f' {"Page" if p_len == 1 else "Pages"}</span>'
      )

    detected_max_pages = max(file_page_counts.values())

    st.markdown(
        f"""
            <div class="detector-bar">
                <span class="detector-label">Page Detector — Only showing pages available in your file:</span>
                <div class="detector-chips">{chips_html}</div>
            </div>
            """,
        unsafe_allow_html=True,
    )

  # --- Strictly bound controls to the detected page count ---
  bound_max = detected_max_pages if queue_count > 0 else 1
  exact_page_options = (
      list(range(1, detected_max_pages + 1)) if queue_count > 0 else []
  )

  # Clean up any previously selected pages in session_state that exceed the newly uploaded PDF's page count
  for state_key in ["ns_only_pills", "hyb_pills"]:
    if state_key in st.session_state:
      st.session_state[state_key] = [
          p for p in st.session_state[state_key] if 1 <= p <= detected_max_pages
      ]

  # --- USER-FRIENDLY MODE TABS ---
  mode_tab_full, mode_tab_seq, mode_tab_nonseq, mode_tab_hybrid = st.tabs([
      "All Pages",
      "Page Range",
      "Specific Pages",
      "Custom Mix",
  ])

  # === MODE TAB 1: ALL PAGES ===
  with mode_tab_full:
    st.markdown(
        """
            <div class="mode-card">
                <div class="mode-card-header">
                    <span class="mode-card-title">Convert Entire PDF</span>
                    <span class="mode-card-tag">All Pages</span>
                </div>
                <p class="mode-card-desc">
                    Converts every page in your PDF into an editable Word document.
                </p>
            </div>
            """,
        unsafe_allow_html=True,
    )

    strat, _, _, preview_p0 = resolve_target_pages(
        "All Pages", detected_max_pages
    )
    render_summary_strip(
        queue_count, detected_max_pages, "All Pages", preview_p0, strat
    )

    if st.button(
        "Convert to Word (.docx)",
        key="btn_launch_full",
        type="primary",
        use_container_width=True,
        disabled=not queue_count,
    ):
      run_batch_pipeline(uploaded_files, file_page_counts, mode="All Pages")

  # === MODE TAB 2: PAGE RANGE ===
  with mode_tab_seq:
    st.markdown(
        """
            <div class="mode-card">
                <div class="mode-card-header">
                    <span class="mode-card-title">Convert a Page Range</span>
                    <span class="mode-card-tag">From – To</span>
                </div>
                <p class="mode-card-desc">
                    Extracts a continuous set of pages from your starting page to your ending page.
                </p>
            </div>
            """,
        unsafe_allow_html=True,
    )

    sc1, sc2, sc3 = st.columns([1, 1, 1.2])
    with sc1:
      seq_start = st.number_input(
          f"From Page (Max: {detected_max_pages})"
          if queue_count
          else "From Page",
          min_value=1,
          max_value=bound_max,
          value=1,
          step=1,
          disabled=not queue_count,
      )
    with sc2:
      seq_end = st.number_input(
          f"To Page (Max: {detected_max_pages})" if queue_count else "To Page",
          min_value=1,
          max_value=bound_max,
          value=bound_max,
          step=1,
          disabled=not queue_count,
      )
    with sc3:
      st.write("")
      seq_excl = st.checkbox(
          "Exclude last page",
          value=False,
          disabled=not queue_count,
          help="If checked, conversion stops right before the 'To Page'.",
      )

    strat, _, _, preview_p0 = resolve_target_pages(
        "Page Range",
        detected_max_pages,
        seq_start=seq_start,
        seq_end=seq_end,
        exclusive_end=seq_excl,
    )
    render_summary_strip(
        queue_count, detected_max_pages, "Page Range", preview_p0, strat
    )

    if st.button(
        "Convert to Word (.docx)",
        key="btn_launch_seq",
        type="primary",
        use_container_width=True,
        disabled=not queue_count or len(preview_p0) == 0,
    ):
      run_batch_pipeline(
          uploaded_files,
          file_page_counts,
          mode="Page Range",
          seq_start=seq_start,
          seq_end=seq_end,
          exclusive_end=seq_excl,
      )

  # === MODE TAB 3: SPECIFIC PAGES ===
  with mode_tab_nonseq:
    st.markdown(
        """
            <div class="mode-card">
                <div class="mode-card-header">
                    <span class="mode-card-title">Pick Specific Pages</span>
                    <span class="mode-card-tag">Custom Pages</span>
                </div>
                <p class="mode-card-desc">
                    Choose individual pages from your uploaded PDF (only pages that exist in your file are shown).
                </p>
            </div>
            """,
        unsafe_allow_html=True,
    )

    nc1, nc2 = st.columns([1.4, 1])
    with nc1:
      ns_pills = st.multiselect(
          f"Select pages (1 to {detected_max_pages} available)"
          if queue_count
          else "Select pages (Upload a PDF first)",
          options=exact_page_options,
          placeholder="Click to pick pages..."
          if queue_count
          else "Upload a PDF above to detect pages",
          disabled=not queue_count,
          key="ns_only_pills",
      )
    with nc2:
      ns_expr = st.text_input(
          f"Or type pages (1–{detected_max_pages})"
          if queue_count
          else "Or type pages",
          value="",
          placeholder=f"e.g. 1, {min(3, bound_max)}"
          if queue_count
          else "Upload a PDF first",
          disabled=not queue_count,
          key="ns_only_expr",
      )

    strat, _, _, preview_p0 = resolve_target_pages(
        "Specific Pages",
        detected_max_pages,
        non_seq_pills=ns_pills,
        non_seq_expr=ns_expr,
    )
    render_summary_strip(
        queue_count, detected_max_pages, "Specific Pages", preview_p0, strat
    )

    if st.button(
        "Convert to Word (.docx)",
        key="btn_launch_nonseq",
        type="primary",
        use_container_width=True,
        disabled=not queue_count or len(preview_p0) == 0,
    ):
      run_batch_pipeline(
          uploaded_files,
          file_page_counts,
          mode="Specific Pages",
          non_seq_pills=ns_pills,
          non_seq_expr=ns_expr,
      )

  # === MODE TAB 4: CUSTOM MIX ===
  with mode_tab_hybrid:
    st.markdown(
        """
            <div class="mode-card">
                <div class="mode-card-header">
                    <span class="mode-card-title">Combine Range & Extra Pages</span>
                    <span class="mode-card-tag">Range + Extra Pages</span>
                </div>
                <p class="mode-card-desc">
                    Convert a main page range and include extra individual pages in a single Word file.
                </p>
            </div>
            """,
        unsafe_allow_html=True,
    )

    hc1, hc2, hc3 = st.columns([1, 1, 1.2])
    with hc1:
      hyb_start = st.number_input(
          f"From Page (Max: {detected_max_pages})"
          if queue_count
          else "From Page",
          min_value=1,
          max_value=bound_max,
          value=1,
          step=1,
          disabled=not queue_count,
          key="hyb_start",
      )
    with hc2:
      hyb_end = st.number_input(
          f"To Page (Max: {detected_max_pages})" if queue_count else "To Page",
          min_value=1,
          max_value=bound_max,
          value=min(3, bound_max),
          step=1,
          disabled=not queue_count,
          key="hyb_end",
      )
    with hc3:
      st.write("")
      hyb_excl = st.checkbox(
          "Exclude last page",
          value=False,
          disabled=not queue_count,
          key="hyb_excl",
      )

    hnc1, hnc2 = st.columns([1.4, 1])
    with hnc1:
      hyb_pills = st.multiselect(
          f"Extra pages to include (1 to {detected_max_pages})"
          if queue_count
          else "Extra pages to include",
          options=exact_page_options,
          placeholder="Click to pick extra pages..."
          if queue_count
          else "Upload a PDF first",
          disabled=not queue_count,
          key="hyb_pills",
      )
    with hnc2:
      hyb_expr = st.text_input(
          f"Or type extra pages (1–{detected_max_pages})"
          if queue_count
          else "Or type extra pages",
          value="",
          placeholder="e.g. 5, 7" if queue_count else "Upload a PDF first",
          disabled=not queue_count,
          key="hyb_expr",
      )

    strat, _, _, preview_p0 = resolve_target_pages(
        "Custom Mix",
        detected_max_pages,
        seq_start=hyb_start,
        seq_end=hyb_end,
        exclusive_end=hyb_excl,
        non_seq_pills=hyb_pills,
        non_seq_expr=hyb_expr,
    )
    render_summary_strip(
        queue_count, detected_max_pages, "Custom Mix", preview_p0, strat
    )

    if st.button(
        "Convert to Word (.docx)",
        key="btn_launch_hybrid",
        type="primary",
        use_container_width=True,
        disabled=not queue_count or len(preview_p0) == 0,
    ):
      run_batch_pipeline(
          uploaded_files,
          file_page_counts,
          mode="Custom Mix",
          seq_start=hyb_start,
          seq_end=hyb_end,
          exclusive_end=hyb_excl,
          non_seq_pills=hyb_pills,
          non_seq_expr=hyb_expr,
      )

  # --- DOWNLOAD READY FILES ---
  if st.session_state.deliverables:
    st.divider()
    st.markdown("##### Ready to Download")

    if len(st.session_state.deliverables) > 1:
      zip_buf = io.BytesIO()
      with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for d in st.session_state.deliverables:
          zf.writestr(d["name"], d["bytes"])

      st.download_button(
          label=(
              f"Download All as ZIP ({len(st.session_state.deliverables)}"
              " files)"
          ),
          data=zip_buf.getvalue(),
          file_name="Converted_Word_Files.zip",
          mime="application/zip",
          type="primary",
          use_container_width=True,
      )

    for item in st.session_state.deliverables:
      col_meta, col_btn = st.columns([4.2, 1])
      with col_meta:
        st.markdown(
            f"""
                    <div class="deliverable-row">
                        <span class="deliverable-name">{item['name']}</span>
                        <span class="deliverable-meta">{item['pages']} &nbsp;•&nbsp; {item['size_kb']} KB &nbsp;•&nbsp; {item['elapsed']}s</span>
                    </div>
                    """,
            unsafe_allow_html=True,
        )
      with col_btn:
        st.download_button(
            label="Download",
            data=item["bytes"],
            file_name=item["name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"dl_{item['name']}",
            use_container_width=True,
        )

# ------------------------------------------
# TOP TAB 2: PAGE PREVIEW
# ------------------------------------------
with nav_inspector:
  if not uploaded_files:
    st.caption("Upload a PDF file first to see its detected page count.")
  else:
    inspector_rows = []
    for uf in uploaded_files:
      total_p = file_page_counts[uf.name]
      _, _, desc_auto = compute_auto_allocation("CONTINUOUS", total_p)
      inspector_rows.append({
          "File Name": uf.name,
          "Detected Pages": f"1 to {total_p} ({total_p} total)",
          "Conversion Speed": desc_auto,
      })
    st.dataframe(inspector_rows, use_container_width=True, hide_index=True)

# ------------------------------------------
# TOP TAB 3: ACTIVITY LOG
# ------------------------------------------
with nav_telemetry:
  if st.session_state.audit_logs:
    st.code("\n".join(st.session_state.audit_logs), language="text")
    if st.button("Clear Log"):
      st.session_state.audit_logs = []
      st.rerun()
  else:
    st.caption("No conversions run yet.")


# ==========================================
# 5. VS CODE PLAY BUTTON AUTO-LAUNCHER
# ==========================================
if __name__ == "__main__":
  from streamlit import runtime
  from streamlit.web import cli as stcli

  if not runtime.exists():
    sys.argv = ["streamlit", "run", sys.argv[0]]
    sys.exit(stcli.main())