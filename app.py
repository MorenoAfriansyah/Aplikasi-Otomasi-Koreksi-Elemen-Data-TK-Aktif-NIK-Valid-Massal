import os
import pandas as pd
import streamlit as st
import zipfile
import io
import warnings

warnings.filterwarnings('ignore', category=UserWarning)

st.set_page_config(
    page_title="Koreksi Data TK BPJS Ketenagakerjaan",
    page_icon="🛡️",
    layout="wide"
)

st.title("🛡️ Aplikasi Koreksi Data TK Aktif BPJS Ketenagakerjaan")
st.markdown("Otomatisasi *VLOOKUP* elemen data peserta, penyesuaian lokasi pekerjaan, status PKWT, lengkap dengan fitur *Preview Before-After* dan unduh ZIP.")

# Setup Direktori Sesi Web
INPUT_DIR = "web_input"
OUTPUT_DIR = "web_output"
for d in [INPUT_DIR, OUTPUT_DIR]:
    if not os.path.exists(d):
        os.makedirs(d)

# Sidebar Panduan
with st.sidebar:
    st.header("Panduan Penggunaan")
    st.markdown("1. **Unggah Referensi**: Masukkan file pembanding (`Sheet2`).")
    st.markdown("2. **Unggah Target**: Pilih banyak file template sekaligus (`.xls/.xlsx`).")
    st.markdown("3. **Preview**: Cek tabel perbandingan perubahan data (*Before-After*).")
    st.markdown("4. **Unduh**: Dapatkan seluruh hasil dalam satu klik format ZIP.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. File Referensi VLOOKUP")
    ref_file_uploaded = st.file_uploader("Unggah file referensi (.xlsx/.xls)", type=['xlsx', 'xls'], key="ref")

ref_file_path = "TESTING TEMPLATE vlookup.xlsx" # Default lokal jika ada
if ref_file_uploaded is not None:
    ref_file_path = os.path.join(INPUT_DIR, ref_file_uploaded.name)
    with open(ref_file_path, "wb") as f:
        f.write(ref_file_uploaded.getbuffer())
    st.sidebar.success(f"Referensi aktif: {ref_file_uploaded.name}")

with col2:
    st.subheader("2. File Template Target Koreksi")
    uploaded_files = st.file_uploader(
        "Pilih banyak file template sekaligus",
        type=['xls', 'xlsx'],
        accept_multiple_files=True,
        key="targets"
    )

if st.button("🚀 Jalankan Proses Koreksi & Preview", type="primary"):
    if not os.path.exists(ref_file_path):
        st.error("Harap unggah atau sediakan file referensi terlebih dahulu!")
    elif not uploaded_files:
        st.error("Harap unggah minimal satu file template koreksi!")
    else:
        with st.spinner("Memproses data koreksi dan menyiapkan preview..."):
            try:
                # Baca referensi
                df_ref = pd.read_excel(ref_file_path, sheet_name='Sheet2')
                df_ref['KPJ'] = df_ref['KPJ'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df_ref = df_ref.drop_duplicates(subset=['KPJ'], keep='first')
                ref_dict = df_ref.set_index('KPJ')[['NAMA_IBU_KANDUNG', 'HANDPHONE', 'EMAIL']].to_dict('index')
                
                success_count = 0
                log_results = []
                preview_data_store = []

                for uploaded_file in uploaded_files:
                    file_name = uploaded_file.name
                    input_path = os.path.join(INPUT_DIR, file_name)
                    
                    with open(input_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                        
                    base_name = os.path.splitext(file_name)[0]
                    output_file_name = base_name + '.xlsx'
                    output_path = os.path.join(OUTPUT_DIR, output_file_name)
                    
                    # Baca file input
                    df_raw = None
                    for engine_choice in [None, 'openpyxl', 'xlrd']:
                        try:
                            df_raw = pd.read_excel(input_path, sheet_name='Koreksi Elemen TK', header=None, engine=engine_choice)
                            break
                        except:
                            pass
                            
                    if df_raw is not None:
                        metadata = df_raw.iloc[0:1].copy()
                        headers = df_raw.iloc[1:2].copy()
                        data_original = df_raw.iloc[2:].copy()
                        
                        data_original.columns = headers.iloc[0].values
                        
                        # Buat salinan untuk data setelah dikoreksi (untuk perbandingan preview)
                        data_modified = data_original.copy()
                        data_modified['KPJ*'] = data_modified['KPJ*'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                        
                        updated_rows = 0
                        for idx, row in data_modified.iterrows():
                            kpj = row['KPJ*']
                            
                            # Rule 1: Lokasi Pekerjaan
                            lokasi = row['Lokasi Pekerjaan']
                            if pd.isna(lokasi) or str(lokasi).strip() in ['', 'nan', 'None']:
                                data_modified.at[idx, 'Lokasi Pekerjaan'] = '3515'
                                
                            # Rule 2: PKWT Kontrak
                            pkwt = str(row['PKWT']).strip().upper()
                            if pkwt == 'Y':
                                data_modified.at[idx, 'Tanggal Akhir Kontrak'] = '31-12-2026'
                            elif pkwt == 'T':
                                data_modified.at[idx, 'Tanggal Akhir Kontrak'] = ''
                                
                            # Rule 3: VLOOKUP
                            if kpj in ref_dict:
                                updated_rows += 1
                                ibu = ref_dict[kpj]['NAMA_IBU_KANDUNG']
                                if pd.notna(ibu):
                                    data_modified.at[idx, 'Nama Ibu Kandung'] = str(ibu).strip()
                                hp = ref_dict[kpj]['HANDPHONE']
                                if pd.notna(hp):
                                    hp_str = str(hp).replace('.0', '').strip()
                                    if hp_str:
                                        data_modified.at[idx, 'Handphone'] = hp_str + '\t'
                                email = ref_dict[kpj]['EMAIL']
                                if pd.notna(email):
                                    data_modified.at[idx, 'Email'] = str(email).strip()
                                    
                        # Simpan hasil akhir ke file fisik
                        metadata.columns = range(len(metadata.columns))
                        headers.columns = range(len(headers.columns))
                        data_modified.columns = range(len(data_modified.columns))
                        
                        final_df = pd.concat([metadata, headers, data_modified], ignore_index=True)
                        final_df.to_excel(output_path, index=False, header=False, sheet_name='Koreksi Elemen TK', engine='openpyxl')
                        
                        success_count += 1
                        log_results.append({"File Asal": file_name, "Status": "Berhasil", "Baris Terkoreksi": updated_rows, "File Output": output_file_name})
                        
                        # Simpan ringkasan untuk Preview Before-After (Kolom penting saja)
                        cols_to_show = ['KPJ*', 'Nama Lengkap', 'Nama Ibu Kandung', 'Handphone', 'Email', 'Lokasi Pekerjaan', 'Tanggal Akhir Kontrak']
                        preview_data_store.append({
                            "file": file_name,
                            "before": data_original[[c for c in cols_to_show if c in data_original.columns]],
                            "after": data_modified[[c for c in cols_to_show if c in data_modified.columns]]
                        })
                    else:
                        log_results.append({"File Asal": file_name, "Status": "Gagal Dibaca", "Baris Terkoreksi": 0, "File Output": "-"})

                st.success(f"Berhasil memproses {success_count} dari {len(uploaded_files)} file!")
                
                # --- TAMPILAN PREVIEW BEFORE - AFTER ---
                st.markdown("---")
                st.subheader("🔍 Pengecekan Preview (Before vs After Koreksi)")
                
                for store in preview_data_store:
                    with st.expander(f"📁 Detail File: {store['file']}"):
                        tab_before, tab_after = st.tabs(["❌ Data Sebelum (Before)", "✅ Data Sesudah (After VLOOKUP)"])
                        with tab_before:
                            st.dataframe(store['before'], use_container_width=True)
                        with tab_after:
                            st.dataframe(store['after'], use_container_width=True)

                # --- BUAT FILE ZIP OTOMATIS ---
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for res in log_results:
                        if res["Status"] == "Berhasil":
                            out_p = os.path.join(OUTPUT_DIR, res["File Output"])
                            zip_file.write(out_p, arcname=res["File Output"])
                zip_buffer.seek(0)

                # --- TOMBOL DOWNLOAD ZIP ---
                st.markdown("---")
                st.subheader("📦 Unduh Seluruh Hasil (Format Folder ZIP)")
                st.download_button(
                    label="📥 Unduh Semua File Hasil (ZIP)",
                    data=zip_buffer,
                    file_name="Hasil_Koreksi_TK_BPJS.zip",
                    mime="application/zip",
                    type="primary"
                )

            except Exception as e:
                st.error(f"Terjadi kesalahan sistem saat memproses: {e}")