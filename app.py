import streamlit as st
import pandas as pd
import os
import xlrd
import tempfile
import zipfile
import warnings
from io import BytesIO

warnings.filterwarnings('ignore', category=UserWarning)

# --- UI STREAMLIT CONFIGURATION ---
st.set_page_config(
    page_title="Koreksi Data TK BPJS Ketenagakerjaan",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling (Eye-Catching, Modern & Clean UI)
st.markdown("""
    <style>
    .main {
        background-color: #f8fafc;
    }
    .stButton>button {
        width: 100%;
        background-color: #2563eb;
        color: white;
        font-weight: 600;
        border-radius: 8px;
        padding: 0.6rem 1rem;
        border: none;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #1d4ed8;
        box-shadow: 0 6px 8px -1px rgba(0, 0, 0, 0.15);
    }
    .card-box {
        background: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        margin-bottom: 20px;
        border-left: 5px solid #2563eb;
    }
    </style>
""", unsafe_allow_html=True)

# Header Section
st.markdown("<h1 style='color: #1e3a8a; text-align: center;'>🛡️ Smart BPJS Data Corrector Pro</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #64748b; font-size: 1.1rem;'>Aplikasi Otomasi VLOOKUP & Koreksi Elemen Data Tenaga Kerja Aktif BPJS Ketenagakerjaan</p>", unsafe_allow_html=True)
st.markdown("---")

# Sidebar for Configuration & Uploads
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/microsoft-excel-2019.png", width=60)
    st.header("📂 Panel Unggah Berkas")
    st.markdown("Unggah file master referensi dan file batch input koreksi Anda di bawah ini.")
    
    ref_file = st.file_uploader("1. File Referensi VLOOKUP (.xlsx)", type=['xlsx'])
    input_files = st.file_uploader("2. Berkas Input (.xls / .xlsx)", type=['xls', 'xlsx'], accept_multiple_files=True)
    
    st.markdown("---")
    st.markdown("### 📌 Aturan Validasi Ketat:")
    st.markdown("• Sheet referensi bernama **Sheet2**.")
    st.markdown("• **Nama Ibu Kandung** wajib ada/valid. Jika kosong, baris diabaikan.")
    st.markdown("• Jika HP/Email kosong tetapi Nama Ibu ada, maka Nama Ibu tetap diisi.")

# Main Interface Area
col1, col2 = st.columns([1, 1], gap="medium")

with col1:
    st.markdown("### 📄 Status Berkas Masuk")
    if ref_file:
        st.success(f"✅ Referensi dimuat: **{ref_file.name}**")
    else:
        st.info("ℹ️ Menunggu unggahan file referensi...")

    if input_files:
        st.success(f"✅ Total {len(input_files)} berkas input siap diproses.")
        with st.expander("Lihat Daftar Berkas Input"):
            for f in input_files:
                st.text(f"• {f.name}")
    else:
        st.info("ℹ️ Belum ada berkas input yang diunggah.")

with col2:
    st.markdown("### ⚙️ Panel Kontrol Eksekusi")
    st.markdown("Klik tombol di bawah untuk menjalankan mesin transformasi data dan validasi ketat.")
    
    process_btn = st.button("🚀 Proses Koreksi Data Sekarang")

# Execution Logic & Display Preview Before/After
if process_btn:
    if not ref_file:
        st.error("❌ Silakan unggah file referensi terlebih dahulu di sidebar.")
    elif not input_files:
        st.error("❌ Silakan unggah minimal 1 file input untuk dikoreksi.")
    else:
        with st.spinner("🔄 Sedang memproses logika VLOOKUP, validasi ketat, dan normalisasi data..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                input_dir = os.path.join(temp_dir, 'input')
                output_dir = os.path.join(temp_dir, 'output')
                os.makedirs(input_dir)
                os.makedirs(output_dir)

                ref_path = os.path.join(temp_dir, ref_file.name)
                with open(ref_path, "wb") as f:
                    f.write(ref_file.getbuffer())

                for uploaded_file in input_files:
                    in_path = os.path.join(input_dir, uploaded_file.name)
                    with open(in_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                try:
                    df_ref = pd.read_excel(ref_path, sheet_name='Sheet2', engine='openpyxl')
                    df_ref.columns = df_ref.columns.str.strip()
                    
                    kpj_idx = df_ref.columns.get_loc('KPJ')
                    col_ibu = df_ref.columns[kpj_idx + 8]
                    col_hp = df_ref.columns[kpj_idx + 9]
                    col_email = df_ref.columns[kpj_idx + 10]

                    df_ref['KPJ'] = df_ref['KPJ'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                    df_ref = df_ref.drop_duplicates(subset=['KPJ'], keep='first')
                    
                    ref_dict = df_ref.set_index('KPJ')[[col_ibu, col_hp, col_email]].to_dict('index')
                    
                    processed_count = 0
                    preview_data_store = []
                    
                    for file_name in os.listdir(input_dir):
                        file_path = os.path.join(input_dir, file_name)
                        output_file_name = os.path.splitext(file_name)[0] + '.xlsx'
                        output_path = os.path.join(output_dir, output_file_name)
                        
                        df_raw = None
                        if file_name.endswith('.xls'):
                            try:
                                wb = xlrd.open_workbook(file_path, ignore_workbook_corruption=True)
                                df_raw = pd.read_excel(wb, sheet_name='Koreksi Elemen TK', header=None, engine='xlrd')
                            except Exception:
                                try:
                                    dfs = pd.read_html(file_path, header=None)
                                    if dfs: df_raw = dfs[0]
                                except Exception: pass
                        
                        if df_raw is None:
                            for engine_choice in ['openpyxl', 'xlrd', None]:
                                try:
                                    df_raw = pd.read_excel(file_path, sheet_name='Koreksi Elemen TK', header=None, engine=engine_choice)
                                    break
                                except Exception: continue
                        
                        if df_raw is None:
                            continue

                        raw_data_clone = df_raw.iloc[2:].copy()
                        raw_headers = df_raw.iloc[1:2].copy()
                        raw_data_clone.columns = raw_headers.iloc[0].values
                        
                        metadata = df_raw.iloc[0:1].copy()   
                        headers = df_raw.iloc[1:2].copy()    
                        data = df_raw.iloc[2:].copy()        
                        
                        data.columns = headers.iloc[0].values
                        data['KPJ*'] = data['KPJ*'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                        
                        valid_lokasi = data['Lokasi Pekerjaan'].dropna().astype(str).str.strip()
                        valid_lokasi = valid_lokasi[~valid_lokasi.isin(['', 'nan', 'None'])]
                        default_lokasi = valid_lokasi.mode().iloc[0] if not valid_lokasi.empty else '3515'
                        
                        for idx, row in data.iterrows():
                            kpj = row['KPJ*']
                            
                            lokasi = row.get('Lokasi Pekerjaan', None)
                            if pd.isna(lokasi) or str(lokasi).strip() in ['', 'nan', 'None']:
                                data.at[idx, 'Lokasi Pekerjaan'] = default_lokasi
                            
                            pkwt = str(row.get('PKWT', '')).strip().upper()
                            if pkwt == 'Y': 
                                data.at[idx, 'Tanggal Akhir Kontrak'] = '31-12-2026'
                            elif pkwt == 'T': 
                                data.at[idx, 'Tanggal Akhir Kontrak'] = ''

                            # Logika VLOOKUP Ketat Berdasarkan Nama Ibu Kandung
                            if kpj in ref_dict:
                                ibu = ref_dict[kpj][col_ibu]
                                hp = ref_dict[kpj][col_hp]
                                email = ref_dict[kpj][col_email]
                                
                                is_ibu_valid = pd.notna(ibu) and str(ibu).strip() != '' and str(ibu).strip().lower() != 'nan'
                                
                                if is_ibu_valid:
                                    # Update Nama Ibu Kandung (Wajib)
                                    data.at[idx, 'Nama Ibu Kandung'] = str(ibu).strip()
                                    
                                    # Update HP jika valid
                                    if pd.notna(hp):
                                        hp_str = str(hp).replace('.0', '').strip()
                                        if hp_str and hp_str.lower() != 'nan':
                                            data.at[idx, 'Handphone'] = hp_str
                                            
                                    # Update Email jika valid
                                    if pd.notna(email):
                                        email_str = str(email).strip()
                                        if email_str and email_str.lower() != 'nan':
                                            data.at[idx, 'Email'] = email_str
                                else:
                                    # Jika Nama Ibu Kandung kosong/invalid, data vlookup tidak diterapkan
                                    pass

                        preview_data_store.append({
                            "filename": file_name,
                            "before": raw_data_clone,
                            "after": data.copy()
                        })

                        metadata.columns = range(len(metadata.columns))
                        headers.columns = range(len(headers.columns))
                        data.columns = range(len(data.columns))
                        
                        final_df = pd.concat([metadata, headers, data], ignore_index=True)
                        final_df.to_excel(output_path, index=False, header=False, sheet_name='Koreksi Elemen TK', engine='openpyxl')
                        processed_count += 1
                        
                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan sistem saat pemrosesan: {e}")
                    processed_count = 0

                if processed_count > 0:
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for root, _, files in os.walk(output_dir):
                            for file in files:
                                file_path = os.path.join(root, file)
                                zip_file.write(file_path, arcname=file)
                    
                    st.markdown("---")
                    st.success(f"🎉 Berhasil memproses **{processed_count}** berkas input dengan validasi ketat!")
                    
                    st.markdown("## 📊 Pratinjau Koreksi Data (Before vs After)")
                    st.markdown("Perbandingan langsung elemen data berdasarkan aturan validasi Nama Ibu Kandung:")
                    
                    for item in preview_data_store:
                        with st.expander(f"📁 Detail Berkas: {item['filename']}", expanded=True):
                            col_b, col_a = st.columns(2)
                            cols_to_show = ['KPJ*', 'Nama Lengkap', 'Lokasi Pekerjaan', 'PKWT', 'Tanggal Akhir Kontrak', 'Nama Ibu Kandung', 'Handphone', 'Email']
                            
                            available_cols_b = [c for c in cols_to_show if c in item['before'].columns]
                            available_cols_a = [c for c in cols_to_show if c in item['after'].columns]
                            
                            with col_b:
                                st.markdown("🔴 **SEBELUM (Data Mentah Input)**")
                                st.dataframe(item['before'][available_cols_b].head(10), use_container_width=True)
                                
                            with col_a:
                                st.markdown("🟢 **SESUDAH (Hasil Koreksi Validasi)**")
                                st.dataframe(item['after'][available_cols_a].head(10), use_container_width=True)

                    st.markdown("---")
                    st.markdown("### 📥 Unduh Hasil Akhir")
                    st.download_button(
                        label="📦 Unduh Seluruh Berkas Koreksi (.ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name="Hasil_Koreksi_BPJS_Final.zip",
                        mime="application/zip"
                    )
