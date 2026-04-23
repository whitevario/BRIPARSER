import streamlit as st
import pandas as pd
import re
import io

# ==== KONFIGURASI ====
st.set_page_config(page_title="BRIVA Filter Tool", layout="wide")
st.title("🔍 BRIVA Transaction Filter")
st.markdown("Upload rekening koran (Excel) dan dapatkan transaksi BRIVA yang valid")

# ==== LOAD PREFIX ====
@st.cache_data
def load_prefixes(uploaded_file):
    if uploaded_file is not None:
        df_prefix = pd.read_excel(uploaded_file)
        df_prefix.columns = df_prefix.columns.str.strip().str.lower()
        prefixes = df_prefix["corporate_code"].astype(str).tolist()
        return prefixes
    return []

# ==== AMBIL BRIVA DARI REMARK ====
def ambil_briva(remark, prefixes):
    text = str(remark)
    text = re.sub(r"[^0-9]", "", text)
    for prefix in prefixes:
        match = re.search(prefix + r"\d{10}", text)
        if match:
            return match.group(0)
    return None

# ==== BERSIHKAN NOMINAL ====
def bersihkan_nominal(x):
    if pd.isna(x):
        return 0
    if isinstance(x, (int, float)):
        return int(x) if isinstance(x, float) else x
    s = str(x).strip()
    s = s.replace(",", "")
    s = re.sub(r"\.00$", "", s)
    if '.' in s and s.replace('.', '').isdigit():
        if ',' in s:
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace('.', '')
    s = re.sub(r'[^\d\-]', '', s)
    try:
        return int(float(s))
    except:
        return 0

# ==== BACA REKENING KORAN ====
def baca_rekening_koran(uploaded_file):
    df_raw = pd.read_excel(uploaded_file, header=None)
    header_row = None
    for idx, row in df_raw.iterrows():
        for cell in row:
            if isinstance(cell, str) and "Tanggal Transaksi" in cell:
                header_row = idx
                break
        if header_row is not None:
            break
    if header_row is None:
        st.error("Tidak ditemukan header 'Tanggal Transaksi'")
        return None

    df = pd.read_excel(uploaded_file, header=header_row)
    df = df.dropna(axis=1, how='all')

    kolom_mapping = {}
    for col in df.columns:
        col_str = str(col).strip()
        if 'Tanggal Transaksi' in col_str:
            kolom_mapping[col] = 'TANGGAL'
        elif 'Uraian Transaksi' in col_str:
            kolom_mapping[col] = 'REMARK'
        elif 'Debet' in col_str:
            kolom_mapping[col] = 'DEBET'
        elif 'Kredit' in col_str:
            kolom_mapping[col] = 'KREDIT'

    df = df.rename(columns=kolom_mapping)
    return df

# ==== SIDE BAR: UPLOAD PREFIX ====
st.sidebar.header("📂 1. Upload Corporate Code")
prefix_file = st.sidebar.file_uploader(
    "Upload corporate_code.xlsx / .xls",
    type=["xlsx", "xls"]
)

if prefix_file:
    prefixes = load_prefixes(prefix_file)
    st.sidebar.success(f"✅ {len(prefixes)} prefix loaded")
    st.sidebar.write("Prefix:", prefixes[:5], "..." if len(prefixes) > 5 else "")
else:
    st.sidebar.warning("⚠️ Upload file corporate_code terlebih dahulu")
    st.stop()

# ==== MAIN AREA: UPLOAD REK KORAN ====
st.header("📄 2. Upload Rekening Koran")
uploaded_files = st.file_uploader(
    "Upload satu atau lebih file Excel rekening koran",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("Silakan upload file rekening koran")
    st.stop()

# ==== PROSES SEMUA FILE ====
all_match = []
all_lain = []

for file in uploaded_files:
    st.markdown(f"### 📁 {file.name}")
    
    # Baca file
    try:
        df = baca_rekening_koran(file)
        if df is None:
            st.error(f"Gagal membaca {file.name}")
            continue
    except Exception as e:
        st.error(f"Error membaca {file.name}: {e}")
        continue

    # Cek kolom yang diperlukan
    if not all(col in df.columns for col in ['TANGGAL', 'REMARK', 'DEBET', 'KREDIT']):
        st.error(f"Kolom tidak lengkap di {file.name}")
        continue

    # Bersihkan nominal
    df['DEBET'] = df['DEBET'].apply(bersihkan_nominal)
    df['KREDIT'] = df['KREDIT'].apply(bersihkan_nominal)

    # Ambil BRIVA
    df['BRIVA'] = df['REMARK'].apply(lambda x: ambil_briva(x, prefixes))

    # Tentukan tipe
    df['TIPE'] = df.apply(
        lambda row: "MASUK" if row['KREDIT'] > 0 else ("KELUAR" if row['DEBET'] > 0 else ""),
        axis=1
    )

    df['ASAL_FILE'] = file.name

    # Pisahkan match vs lain
    df_match = df[df['BRIVA'].notna() & df['BRIVA'].str[:5].isin(prefixes)].copy()
    df_lain = df.drop(df_match.index).copy()

    # Ganti remark dengan BRIVA untuk yg match
    if not df_match.empty:
        df_match['REMARK'] = df_match['BRIVA']

    # Simpan ke rekap
    all_match.append(df_match)
    all_lain.append(df_lain)

    # Tampilkan preview
    col1, col2 = st.columns(2)
    with col1:
        st.metric("✅ BRIVA MATCH", len(df_match))
    with col2:
        st.metric("📄 LAIN-LAIN", len(df_lain))

    if not df_match.empty:
        st.dataframe(df_match[['TANGGAL', 'REMARK', 'DEBET', 'KREDIT', 'TIPE']].head(10))

# ==== GABUNG & DOWNLOAD ====
st.header("📥 3. Download Hasil")

if all_match or all_lain:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if all_match:
            df_all_match = pd.concat(all_match, ignore_index=True)
            df_all_match.to_excel(writer, index=False, sheet_name='BRIVA_MATCH')
            st.success(f"📊 Total BRIVA_MATCH: {len(df_all_match)} baris")
        if all_lain:
            df_all_lain = pd.concat(all_lain, ignore_index=True)
            df_all_lain.to_excel(writer, index=False, sheet_name='LAIN-LAIN')
            st.success(f"📊 Total LAIN-LAIN: {len(df_all_lain)} baris")
        # Tambah sheet prefix
        if prefix_file:
            df_prefix_display = pd.DataFrame({"corporate_code": prefixes})
            df_prefix_display.to_excel(writer, index=False, sheet_name='PREFIX_LIST')

    st.download_button(
        label="💾 Download Rekap BRIVA.xlsx",
        data=output.getvalue(),
        file_name="rekap_briva.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.warning("Tidak ada data yang diproses")