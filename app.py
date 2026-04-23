import streamlit as st
import pandas as pd
import re
import io

# ==== IMPORT PDF LIBRARIES ====
import pdfplumber
import PyPDF2

# ==== KONFIGURASI ====
st.set_page_config(page_title="BRIVA Filter Tool - Support Excel & PDF", layout="wide")
st.title("🔍 BRIVA Transaction Filter")
st.markdown("Upload **Rekening Koran (Excel/PDF)** dan dapatkan transaksi BRIVA yang valid")

# ==== LOAD PREFIX ====
@st.cache_data
def load_prefixes(uploaded_file):
    if uploaded_file is not None:
        df_prefix = pd.read_excel(uploaded_file)
        df_prefix.columns = df_prefix.columns.str.strip().str.lower()
        prefixes = df_prefix["corporate_code"].astype(str).tolist()
        return prefixes
    return []

# ==== EKSTRAK TEKS DARI PDF ====
def ekstrak_teks_pdf_pypdf2(uploaded_file):
    """Ekstrak teks dari PDF menggunakan PyPDF2"""
    reader = PyPDF2.PdfReader(uploaded_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text()
    return text

def ekstrak_teks_pdf_pdfplumber(uploaded_file):
    """Ekstrak teks dari PDF menggunakan pdfplumber (lebih akurat)"""
    text = ""
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text += page.extract_text()
    return text

def ekstrak_teks_pdf(uploaded_file):
    """Coba ekstrak dengan pdfplumber dulu, fallback ke PyPDF2"""
    try:
        # Reset file pointer
        uploaded_file.seek(0)
        text = ekstrak_teks_pdf_pdfplumber(uploaded_file)
        if text.strip():
            return text
    except:
        pass
    
    try:
        uploaded_file.seek(0)
        text = ekstrak_teks_pdf_pypdf2(uploaded_file)
        return text
    except Exception as e:
        st.error(f"Gagal ekstrak PDF: {e}")
        return ""

# ==== EKSTRAK TRANSAKSI DARI TEKS PDF ====
def ekstrak_transaksi_dari_teks(text, prefixes):
    """
    Ambil transaksi dari teks PDF
    Format umum di rekening koran PDF:
    Tanggal Waktu Uraian Debet Kredit Saldo
    """
    lines = text.split('\n')
    transaksi = []
    
    # Pattern untuk date (format: 21/04/26 atau 21-04-26)
    date_pattern = r'(\d{2}[/\-]\d{2}[/\-]\d{2})'
    # Pattern untuk waktu (HH:MM:SS)
    time_pattern = r'(\d{2}:\d{2}:\d{2})'
    # Pattern untuk nominal (contoh: 40,000,000.00 atau 40.000.000)
    nominal_pattern = r'[\d\.,]+'
    
    for line in lines:
        # Cari tanggal di awal line
        date_match = re.search(date_pattern, line)
        if date_match:
            tanggal = date_match.group(1)
            
            # Cari waktu
            time_match = re.search(time_pattern, line)
            waktu = time_match.group(1) if time_match else "00:00:00"
            
            # Cari BRIVA di dalam line
            briva = ambil_briva_from_text(line, prefixes)
            
            # Cari nominal (debet/kredit)
            nominal_matches = re.findall(nominal_pattern, line)
            debet = 0
            kredit = 0
            
            # Parse nominal (ambil 2 nominal terakhir biasanya)
            for nom in nominal_matches:
                nom_clean = bersihkan_nominal_pdf(nom)
                if nom_clean > 0:
                    if "debet" in line.lower() or "debit" in line.lower():
                        debet = nom_clean
                    elif "kredit" in line.lower() or "credit" in line.lower():
                        kredit = nom_clean
                    else:
                        # Jika tidak ada label, asumsikan kredit
                        kredit = nom_clean
            
            if briva or debet > 0 or kredit > 0:
                transaksi.append({
                    'TANGGAL': tanggal,
                    'WAKTU': waktu,
                    'REMARK': line,
                    'BRIVA': briva,
                    'DEBET': debet,
                    'KREDIT': kredit
                })
    
    return pd.DataFrame(transaksi)

def ambil_briva_from_text(text, prefixes):
    """Ambil nomor BRIVA dari text (tanpa harus hapus non-digit dulu)"""
    # Pattern untuk BRIVA (5 digit prefix + 10 digit)
    for prefix in prefixes:
        pattern = prefix + r'\d{10}'
        match = re.search(pattern, text)
        if match:
            return match.group(0)
    return None

def bersihkan_nominal_pdf(nominal_str):
    """Bersihkan nominal dari PDF (format: 40,000,000.00 atau 40.000.000)"""
    if not nominal_str:
        return 0
    # Hapus koma (pemisah ribuan)
    s = nominal_str.replace(',', '')
    # Hapus .00 di akhir
    s = re.sub(r'\.00$', '', s)
    # Hapus titik pemisah ribuan
    s = s.replace('.', '')
    # Hanya ambil digit
    s = re.sub(r'[^\d]', '', s)
    try:
        return int(s)
    except:
        return 0

# ==== AMBIL BRIVA DARI REMARK (Excel) ====
def ambil_briva(remark, prefixes):
    text = str(remark)
    text = re.sub(r"[^0-9]", "", text)
    for prefix in prefixes:
        match = re.search(prefix + r"\d{10}", text)
        if match:
            return match.group(0)
    return None

# ==== BERSIHKAN NOMINAL (Excel) ====
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

# ==== BACA REKENING KORAN EXCEL ====
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

# ==== PROSES EXCEL ====
def proses_excel(file, prefixes):
    df = baca_rekening_koran(file)
    if df is None:
        return None, None
    
    if not all(col in df.columns for col in ['TANGGAL', 'REMARK', 'DEBET', 'KREDIT']):
        st.error(f"Kolom tidak lengkap di {file.name}")
        return None, None
    
    df['DEBET'] = df['DEBET'].apply(bersihkan_nominal)
    df['KREDIT'] = df['KREDIT'].apply(bersihkan_nominal)
    df['BRIVA'] = df['REMARK'].apply(lambda x: ambil_briva(x, prefixes))
    df['TIPE'] = df.apply(
        lambda row: "MASUK" if row['KREDIT'] > 0 else ("KELUAR" if row['DEBET'] > 0 else ""),
        axis=1
    )
    df['ASAL_FILE'] = file.name
    
    df_match = df[df['BRIVA'].notna() & df['BRIVA'].str[:5].isin(prefixes)].copy()
    df_lain = df.drop(df_match.index).copy()
    
    if not df_match.empty:
        df_match['REMARK'] = df_match['BRIVA']
    
    return df_match, df_lain

# ==== PROSES PDF ====
def proses_pdf(file, prefixes):
    # Ekstrak teks dari PDF
    text = ekstrak_teks_pdf(file)
    if not text:
        st.error(f"Tidak bisa ekstrak teks dari {file.name}")
        return None, None
    
    # Ekstrak transaksi dari teks
    df = ekstrak_transaksi_dari_teks(text, prefixes)
    if df.empty:
        st.warning(f"Tidak ada transaksi yang ditemukan di {file.name}")
        return None, None
    
    df['TIPE'] = df.apply(
        lambda row: "MASUK" if row['KREDIT'] > 0 else ("KELUAR" if row['DEBET'] > 0 else ""),
        axis=1
    )
    df['ASAL_FILE'] = file.name
    
    df_match = df[df['BRIVA'].notna() & df['BRIVA'].str[:5].isin(prefixes)].copy()
    df_lain = df.drop(df_match.index).copy()
    
    if not df_match.empty:
        df_match['REMARK'] = df_match['BRIVA']
    
    # Kolom output standar
    kolom_standar = ['TANGGAL', 'WAKTU', 'REMARK', 'DEBET', 'KREDIT', 'TIPE', 'ASAL_FILE']
    for col in kolom_standar:
        if col not in df_match.columns and col != 'WAKTU':
            df_match[col] = ''
        if col not in df_lain.columns and col != 'WAKTU':
            df_lain[col] = ''
    
    return df_match, df_lain

# ==== SIDEBAR: UPLOAD PREFIX ====
st.sidebar.header("📂 1. Upload Corporate Code")
prefix_file = st.sidebar.file_uploader(
    "Upload corporate_code.xlsx / .xls",
    type=["xlsx", "xls"]
)

if prefix_file:
    prefixes = load_prefixes(prefix_file)
    st.sidebar.success(f"✅ {len(prefixes)} prefix loaded")
    with st.sidebar.expander("Lihat daftar prefix"):
        st.write(prefixes)
else:
    st.sidebar.warning("⚠️ Upload file corporate_code terlebih dahulu")
    st.stop()

# ==== MAIN AREA: UPLOAD FILE ====
st.header("📄 2. Upload Rekening Koran (Excel atau PDF)")
uploaded_files = st.file_uploader(
    "Upload satu atau lebih file (Excel .xlsx/.xls atau PDF)",
    type=["xlsx", "xls", "pdf"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("Silakan upload file rekening koran (Excel atau PDF)")
    st.stop()

# ==== PROSES SEMUA FILE ====
all_match = []
all_lain = []
progress_bar = st.progress(0)

for i, file in enumerate(uploaded_files):
    st.markdown(f"### 📁 {file.name}")
    st.write(f"**Tipe:** {file.type}")
    
    if file.type == "application/pdf":
        df_match, df_lain = proses_pdf(file, prefixes)
    else:
        df_match, df_lain = proses_excel(file, prefixes)
    
    if df_match is not None and not df_match.empty:
        all_match.append(df_match)
        st.success(f"✅ BRIVA MATCH: {len(df_match)} transaksi")
        st.dataframe(df_match[['TANGGAL', 'REMARK', 'DEBET', 'KREDIT', 'TIPE']].head(5))
    else:
        st.info("📭 Tidak ada transaksi BRIVA yang ditemukan")
    
    if df_lain is not None and not df_lain.empty:
        all_lain.append(df_lain)
        st.info(f"📄 LAIN-LAIN: {len(df_lain)} transaksi")
    
    progress_bar.progress((i + 1) / len(uploaded_files))

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
