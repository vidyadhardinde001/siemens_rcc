try:
    import cv2
except ImportError:
    import sys
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "opencv-python-headless"])
    import cv2
import pytesseract
import pandas as pd
from difflib import SequenceMatcher
import re
from pdf2image import convert_from_path
import os
import numpy as np
from PIL import Image
import streamlit as st
import threading
import platform
import time
from datetime import datetime

# Configure paths for Tesseract and Poppler
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_PATH = r"C:\Users\Vidyadhar\Downloads\Release-24.08.0-0\poppler-24.08.0\Library\bin"

# --- UI Layout ---
st.set_page_config(page_title="SIEMENS RCC Comparator", layout="wide")
st.markdown(
    f"""
    <style>
    .reportview-container {{
        background: #f5f5f5;
    }}
    .sidebar .sidebar-content {{
        background: #e6f2ff;
    }}
    .stButton>button {{
        color: white;
        background: #0066b3;
        border-radius: 6px;
        font-weight: bold;
        padding: 0.5em 1.5em;
    }}
    .stButton>button:hover {{
        background: #00b0f0;
        color: #333333;
    }}
    .stProgress>div>div>div>div {{
        background-color: #0066b3;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

st.title("SIEMENS RCC Comparator")
st.caption("Document Comparison Tool")

with st.expander("Document Selection", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        orig_file = st.file_uploader("Original Document", type=["pdf"], key="orig")
    with col2:
        mod_file = st.file_uploader("Modified Document", type=["pdf"], key="mod")

with st.expander("Comparison Settings", expanded=True):
    output_dir = st.text_input("Output Folder", value="comparison_results")
    sensitivity = st.slider("Comparison Sensitivity", min_value=0.5, max_value=1.0, value=0.75, step=0.01, format="%.2f")
    st.caption(f"Sensitivity: {int(sensitivity*100)}%")

progress_placeholder = st.empty()
status_placeholder = st.empty()
time_placeholder = st.empty()

results_expander = st.expander("Comparison Results", expanded=True)
results_log = st.session_state.get("results_log", [])

def extract_text_data(pil_image):
    img_cv = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
    config = r'--oem 3 --psm 6'
    data = pytesseract.image_to_data(
        img_cv, 
        output_type=pytesseract.Output.DATAFRAME, 
        config=config
    )
    return data[data.text.notnull()].reset_index(drop=True), img_cv

def is_text_similar(a, b, threshold=0.75):
    def normalize(text):
        text = text.strip().lower()
        text = re.sub(r'[^a-z0-9]', '', text)
        text = text.replace('0', 'o').replace('1', 'i').replace('5', 's').replace('8', 'b')
        return text
    norm_a = normalize(a)
    norm_b = normalize(b)
    if norm_a == norm_b:
        return True
    return SequenceMatcher(None, norm_a, norm_b).ratio() > threshold

def highlight_changes(image, text_data, change_indices):
    overlay = image.copy()
    padding = 4
    for idx in change_indices:
        if idx < len(text_data):
            row = text_data.loc[idx]
            x, y, w, h = int(row['left']), int(row['top']), int(row['width']), int(row['height'])
            x1 = max(x - padding, 0)
            y1 = max(y - padding, 0)
            x2 = x + w + padding
            y2 = y + h + padding
            cv2.rectangle(overlay, (x1, y1), (x2, y2), (0, 0, 255), -1)
    return cv2.addWeighted(overlay, 0.3, image, 0.7, 0)

def create_pdf_report(image_paths, output_path):
    images = [Image.open(img).convert("RGB") for img in image_paths]
    images[0].save(output_path, save_all=True, append_images=images[1:])

def log_result(message, tag=None):
    timestamp = datetime.now().strftime("%H:%M:%S")
    results_log.append((timestamp, message, tag))
    st.session_state["results_log"] = results_log

def compare_documents(original_path, modified_path, output_dir, sensitivity, progress_cb, status_cb, time_cb, cancel_flag):
    try:
        progress_cb(5)
        status_cb("Converting documents to images...")
        original_pages = convert_from_path(
            original_path, 
            dpi=300, 
            poppler_path=POPPLER_PATH
        )
        modified_pages = convert_from_path(
            modified_path, 
            dpi=300, 
            poppler_path=POPPLER_PATH
        )
        total_pages = min(len(original_pages), len(modified_pages))
        if total_pages == 0:
            raise ValueError("No pages found in one or both documents")
        log_result(f"Processing {total_pages} page comparisons...", "info")
        result_images = []
        start_time = time.time()
        for i, (orig_page, mod_page) in enumerate(zip(original_pages, modified_pages)):
            if cancel_flag["cancel"]:
                log_result("Comparison cancelled by user", "warning")
                break
            progress = 10 + (i * 80 / total_pages)
            progress_cb(progress)
            status_cb(f"Comparing page {i+1}/{total_pages}...")
            orig_data, _ = extract_text_data(orig_page)
            mod_data, mod_img = extract_text_data(mod_page)
            orig_words = orig_data['text'].tolist()
            mod_words = mod_data['text'].tolist()
            changed_indices = [
                idx for idx, mod_word in enumerate(mod_words)
                if not any(is_text_similar(mod_word, orig_word, sensitivity) 
                         for orig_word in orig_words)
            ]
            if changed_indices:
                highlighted_img = highlight_changes(mod_img, mod_data, changed_indices)
                output_path = os.path.join(output_dir, f"page_{i+1}_diff.png")
                cv2.imwrite(output_path, highlighted_img)
                result_images.append(output_path)
                log_result(
                    f"Page {i+1}: Found {len(changed_indices)} differences → {output_path}",
                    "success"
                )
            else:
                log_result(f"Page {i+1}: No differences found", "info")
            elapsed = time.time() - start_time
            if progress > 5:
                estimated_total = elapsed / (progress / 100)
                remaining = max(0, estimated_total - elapsed)
                if remaining > 60:
                    mins = int(remaining // 60)
                    secs = int(remaining % 60)
                    time_str = f"Estimated time remaining: {mins}m {secs}s"
                else:
                    time_str = f"Estimated time remaining: {int(remaining)}s"
                time_cb(time_str)
        if result_images and not cancel_flag["cancel"]:
            progress_cb(95)
            status_cb("Generating final report...")
            pdf_path = os.path.join(output_dir, "comparison_report.pdf")
            create_pdf_report(result_images, pdf_path)
            log_result(f"Final report generated: {pdf_path}", "success")
            st.session_state["result_pdf_path"] = pdf_path
        if not cancel_flag["cancel"]:
            progress_cb(100)
            status_cb("Comparison complete!")
            log_result("Document comparison finished successfully", "success")
        else:
            progress_cb(0)
            status_cb("Comparison cancelled")
    except Exception as e:
        progress_cb(0)
        status_cb(f"Error: {str(e)}")
        log_result(f"Error during comparison: {str(e)}", "error")

def save_uploaded_file(uploadedfile, save_path):
    with open(save_path, "wb") as f:
        f.write(uploadedfile.getbuffer())

# --- Main UI Actions ---
if "cancel_flag" not in st.session_state:
    st.session_state["cancel_flag"] = {"cancel": False}
if "result_pdf_path" not in st.session_state:
    st.session_state["result_pdf_path"] = None

compare_btn = st.button("Compare Documents", key="compare")
cancel_btn = st.button("Cancel", key="cancel")
open_btn = st.button("Open Results", key="open", disabled=not st.session_state.get("result_pdf_path"))

if compare_btn:
    if not orig_file or not mod_file:
        st.error("Please select both documents to compare")
    else:
        # Save uploaded files to disk for processing
        os.makedirs(output_dir, exist_ok=True)
        orig_path = os.path.join(output_dir, "original.pdf")
        mod_path = os.path.join(output_dir, "modified.pdf")
        save_uploaded_file(orig_file, orig_path)
        save_uploaded_file(mod_file, mod_path)
        st.session_state["cancel_flag"]["cancel"] = False
        progress_placeholder.progress(0)
        status_placeholder.info("Preparing to compare documents...")
        time_placeholder.text("")
        st.session_state["results_log"] = []
        def progress_cb(val):
            progress_placeholder.progress(int(val))
        def status_cb(msg):
            status_placeholder.info(msg)
        def time_cb(msg):
            time_placeholder.text(msg)
        cancel_flag = st.session_state["cancel_flag"]
        compare_documents(
            orig_path, mod_path, output_dir, sensitivity,
            progress_cb, status_cb, time_cb, cancel_flag
        )

if cancel_btn:
    st.session_state["cancel_flag"]["cancel"] = True
    status_placeholder.warning("Cancelling... Please wait")

# ...existing code...

if open_btn and st.session_state.get("result_pdf_path"):
    result_pdf_path = st.session_state["result_pdf_path"]
    if os.path.exists(result_pdf_path):
        st.success(f"Opening results folder: {os.path.dirname(result_pdf_path)}")
        if platform.system() == "Windows":
            os.startfile(os.path.dirname(result_pdf_path))
        elif platform.system() == "Darwin":
            os.system(f'open "{os.path.dirname(result_pdf_path)}"')
        else:
            os.system(f'xdg-open "{os.path.dirname(result_pdf_path)}"')
        
        # --- Show page-by-page diff images ---
        diff_images = sorted([
            f for f in os.listdir(os.path.dirname(result_pdf_path))
            if f.startswith("page_") and f.endswith("_diff.png")
        ])
        if diff_images:
            st.markdown("### Page-by-Page Differences")
            for img_name in diff_images:
                img_path = os.path.join(os.path.dirname(result_pdf_path), img_name)
                st.image(img_path, caption=img_name, use_column_width=True)
        else:
            st.info("No page differences found to display.")
    else:
        st.warning("No comparison results available yet")

# ...existing code...

# --- Results Log ---
with results_expander:
    for ts, msg, tag in results_log:
        if tag == "success":
            st.success(f"[{ts}] {msg}")
        elif tag == "error":
            st.error(f"[{ts}] {msg}")
        elif tag == "warning":
            st.warning(f"[{ts}] {msg}")
        else:
            st.info(f"[{ts}] {msg}")

# --- Help Section ---
with st.expander("Help Guide"):
    st.markdown("""
**SIEMENS RCC Comparator - Help Guide**

1. **SELECTING DOCUMENTS**
   - Click "Browse" to select original and modified PDFs
   - Both documents must be PDF files with readable text

2. **CONFIGURING COMPARISON**
   - Output folder: Where results will be saved
   - Sensitivity: Adjust how strict the comparison should be
     - Higher values = only major changes
     - Lower values = more subtle differences

3. **RUNNING COMPARISON**
   - Click "Compare Documents" to start
   - Progress bar shows current status
   - Cancel button stops the operation

4. **VIEWING RESULTS**
   - Results show differences found
   - "Open Results" opens the output folder
   - Each page with changes is saved as an image
   - Complete report is generated as PDF

**TIPS:**
- Use high-quality PDFs for best results
- Standard fonts improve OCR accuracy
- Larger documents take more time to process

For support contact: automation.support@siemens.com
""")