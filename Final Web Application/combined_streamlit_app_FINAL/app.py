import os
import subprocess
import pandas as pd
import streamlit as st
from pathlib import Path

# --- CONFIG ---
FOLDER_PATH = r"D:\internship\FINAL INTERNSHIP PROJECT\Indivisual Dashboards"  # 👈 Update if needed
EXE_COMBINE = os.path.join(FOLDER_PATH, "Combining_Excel_Files.exe")
EXE_COPY = os.path.join(FOLDER_PATH, "Copying_FROM_Combined_to_Connected.exe")
SUPPORTED_EXTENSIONS = ['.xlsx', '.xlsm', '.pbix']

# --- PAGE SETUP ---
st.set_page_config(page_title="FileSync Pro - Bajaj Mukand", layout="wide", page_icon="📁")

# --- CSS STYLING ---
st.markdown("""
<style>
.logo {
    font-weight: 900;
    font-size: 2.8rem;
    background: linear-gradient(90deg, #004aad, #0085ff);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin-bottom: -10px;
}
.logo-subtitle {
    font-weight: 600;
    font-size: 1.1rem;
    color: #555f77;
    letter-spacing: 1.2px;
    margin-bottom: 20px;
}
h1, h2, h3 {
    color: #003b70;
}
.file-row {
    padding: 12px 18px;
    border-radius: 10px;
    margin-bottom: 6px;
    background: white;
    display: flex;
    justify-content: space-between;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
}
.file-row:hover {
    background: #e6f0ff;
    box-shadow: 0 4px 14px rgba(0, 75, 210, 0.25);
}
.file-name {
    font-weight: 700;
    font-size: 1.05rem;
    color: #003b70;
}
.file-type {
    color: #0077ff;
    font-weight: 600;
    margin-left: 8px;
    font-size: 0.9rem;
}
.open-btn button {
    background: #0077ff !important;
    color: white !important;
    font-weight: 700 !important;
    border-radius: 8px !important;
    padding: 6px 18px !important;
    box-shadow: 0 2px 8px rgba(0, 119, 255, 0.4);
}
.open-btn button:hover {
    background: #005bb5 !important;
    box-shadow: 0 4px 20px rgba(0, 91, 181, 0.6);
}
footer {
    margin-top: 40px;
    text-align: center;
    color: #a1a9bf;
    font-size: 12px;
    letter-spacing: 1.1px;
}
</style>
""", unsafe_allow_html=True)

# --- HEADER ---
st.markdown('<div class="logo">Bajaj Mukand</div>', unsafe_allow_html=True)
st.markdown('<div class="logo-subtitle">FileSync Pro — Your Live Folder Dashboard</div>', unsafe_allow_html=True)
st.markdown("---")

# --- EXE ACTION BUTTONS ---
st.markdown("## ⚙️ Run Data Processing Tools")

col1, col2 = st.columns(2)
with col1:
    if st.button("▶️ Run: Combine Excel Files"):
        if os.path.exists(EXE_COMBINE):
            subprocess.Popen([EXE_COMBINE], shell=True)
            st.success("✅ Combining Excel Files... launched.")
        else:
            st.error("❌ Combining_Excel_Files.exe not found.")

with col2:
    if st.button("▶️ Run: Copy to Connected"):
        if os.path.exists(EXE_COPY):
            subprocess.Popen([EXE_COPY], shell=True)
            st.success("✅ Copying to Connected... launched.")
        else:
            st.error("❌ Copying_FROM_Combined_to_Connected.exe not found.")

st.markdown("---")

# --- FUNCTION TO LIST FILES ---
def get_files(folder):
    file_list = []
    for root, _, files in os.walk(folder):
        for file in files:
            ext = os.path.splitext(file)[1].lower()
            if ext in SUPPORTED_EXTENSIONS and not file.startswith("combined_output_"):
                full_path = os.path.join(root, file)
                file_list.append({
                    'File Name': file,
                    'Type': ext[1:].upper(),
                    'Path': full_path
                })
    return pd.DataFrame(file_list)

# --- DISPLAY REGULAR FILES ---
df_files = get_files(FOLDER_PATH)

search = st.text_input("🔍 Search files by name:")
if search:
    df_files = df_files[df_files['File Name'].str.contains(search, case=False)]

st.markdown(f"### 📄 Found **{len(df_files)}** regular files")

file_type_emoji = {
    'XLSX': "📊",
    'XLSM': "📈",
    'PBIX': "📉"
}

for idx, row in df_files.iterrows():
    file_icon = file_type_emoji.get(row['Type'], "📁")
    file_display = f"<span class='file-name'>{file_icon} {row['File Name']}</span><span class='file-type'>({row['Type']})</span>"
    cols = st.columns([8, 1])
    with cols[0]:
        st.markdown(f"<div class='file-row'>{file_display}</div>", unsafe_allow_html=True)
    with cols[1]:
        if st.button("Open", key=f"open_{idx}", help=f"Open {row['File Name']}"):
            subprocess.Popen(['start', '', row['Path']], shell=True)

# --- SEPARATOR ---
st.markdown("---")
st.markdown("## 🧩 Combined Output Files")

# --- DISPLAY COMBINED FILES ---
combined_candidates = [
    os.path.join(FOLDER_PATH, f)
    for f in os.listdir(FOLDER_PATH)
    if f.startswith("combined_output_") and f.lower().endswith(('.xlsx', '.pbix'))
]

if combined_candidates:
    st.success(f"✅ Found {len(combined_candidates)} combined output files")

    for i, filepath in enumerate(sorted(combined_candidates, key=os.path.getctime, reverse=True)):
        filename = os.path.basename(filepath)
        ext = filename.lower().split('.')[-1]
        st.markdown(f"**📦 {filename}**")

        # Show preview for the latest xlsx
        if i == 0 and ext == "xlsx":
            try:
                df_combined = pd.read_excel(filepath, sheet_name="CombinedData")
                st.dataframe(df_combined, use_container_width=True)
            except Exception as e:
                st.error(f"❌ Could not preview Excel file: {e}")

        if st.button(f"📂 Open {filename}", key=f"open_combined_{i}"):
            subprocess.Popen(['start', '', str(filepath)], shell=True)
else:
    st.warning("⚠️ No combined output files (.xlsx or .pbix) found in the folder.")

# --- FOOTER ---
st.markdown("---")
st.markdown("<footer>© 2025 FileSync Pro — Developed by Omkar Patil for Bajaj Mukand</footer>", unsafe_allow_html=True)
