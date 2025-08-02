import os
import subprocess
import pandas as pd
import streamlit as st

# --- PAGE SETUP ---
st.set_page_config(page_title="FileSync Pro - Bajaj Mukand", layout="wide", page_icon="📁")

# --- CSS STYLING ---
st.markdown("""<style>
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
</style>""", unsafe_allow_html=True)

# --- HEADER ---
st.markdown('<div class="logo">Bajaj Mukand</div>', unsafe_allow_html=True)
st.markdown('<div class="logo-subtitle">FileSync Pro — Your Live Folder Dashboard</div>', unsafe_allow_html=True)
st.markdown("---")

# --- SESSION STATE ---
if "recent_folders" not in st.session_state:
    st.session_state.recent_folders = []

# --- FOLDER SELECTION ---
st.markdown("### 📁 Select or Enter Folder Path")
recent_folders = st.session_state.recent_folders
selected_folder = st.selectbox(
    "⬇️ Recently Used Folders", options=["-- Select --"] + recent_folders, index=0
)
manual_path = st.text_input("📂 Or manually enter a folder path:")

folder_path = ""
if selected_folder != "-- Select --":
    folder_path = selected_folder
elif manual_path:
    folder_path = manual_path.strip()

if not folder_path:
    st.info("📂 Please enter or select a folder path to start.")
elif not os.path.isdir(folder_path):
    st.error("❌ Invalid path. Please check the folder path again.")
else:
    # Save to recent folders list
    if folder_path not in st.session_state.recent_folders:
        st.session_state.recent_folders.insert(0, folder_path)
        if len(st.session_state.recent_folders) > 5:
            st.session_state.recent_folders.pop()

    # --- CONFIG ---
    SUPPORTED_EXTENSIONS = ['.xlsx', '.xlsm', '.pbix']
    EXE_COMBINE = os.path.join(folder_path, "Combining_Excel_Files.exe")

    # --- EXE ACTIONS + SUMMARY TOOL ---
    st.markdown("## ⚙️ Tools & Actions")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("▶️ Run: Combine Excel Files"):
            if os.path.exists(EXE_COMBINE):
                subprocess.Popen([EXE_COMBINE], shell=True)
                st.success("✅ Combining Excel Files... launched.")
            else:
                st.error("❌ Combining_Excel_Files.exe not found in selected folder.")

    with col2:
        if st.button("📊 Show Folder Summary"):
            try:
                df_summary = pd.DataFrame()
                df_summary = []
                total_size = 0
                for root, _, files in os.walk(folder_path):
                    for file in files:
                        ext = os.path.splitext(file)[1].lower()
                        if ext in SUPPORTED_EXTENSIONS and not file.startswith("combined_output_"):
                            full_path = os.path.join(root, file)
                            df_summary.append({
                                'File Name': file,
                                'Type': ext[1:].upper(),
                                'Path': full_path
                            })
                            total_size += os.path.getsize(full_path)

                df_summary = pd.DataFrame(df_summary)
                total_files = len(df_summary)
                total_size_mb = round(total_size / (1024 * 1024), 2)

                st.success(f"📁 Folder contains {total_files} supported files")
                st.info(f"📦 Total Size: {total_size_mb} MB")

                type_counts = df_summary['Type'].value_counts().reset_index()
                type_counts.columns = ['File Type', 'Count']
                st.markdown("### 📌 File Type Breakdown")
                st.table(type_counts)

            except Exception as e:
                st.error(f"❌ Failed to generate summary: {e}")

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
    df_files = get_files(folder_path)
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

    # --- COMBINED OUTPUT FILES ---
    st.markdown("---")
    st.markdown("## 🧩 Combined Output Files")

    combined_candidates = [
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.startswith("combined_output_") and f.lower().endswith(('.xlsx', '.pbix'))
    ]

    if combined_candidates:
        st.success(f"✅ Found {len(combined_candidates)} combined output files")

        for i, filepath in enumerate(sorted(combined_candidates, key=os.path.getctime, reverse=True)):
            filename = os.path.basename(filepath)
            ext = filename.lower().split('.')[-1]
            st.markdown(f"**📦 {filename}**")

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
