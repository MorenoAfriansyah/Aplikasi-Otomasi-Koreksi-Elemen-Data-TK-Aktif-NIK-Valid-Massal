import streamlit as st
import pandas as pd
import os
import xlrd
import tempfile
import zipfile
import warnings
from io import BytesIO

warnings.filterwarnings('ignore', category=UserWarning)

# --- UI STREAMLIT ---
st.set_page_config(page_title="Koreksi Data TK BPJS", layout="wide")
st.title("APLIKASI OTOMASI KOREKSI ELEMEN DATA TK AKTIF NIK VALID MASSAL")

ref_file = st.file_uploader("1. Upload File Referensi (.xlsx)", type=['xlsx'])
input_files = st.file_uploader("2. Upload File Input (.xls / .xlsx)", type=['xls', 'xlsx'], accept_multiple_files=True)

if st.button("Proses Data"):
    if not ref_file:
        st.error("Silakan upload file referensi.")
    elif not input_files:
        st.error("Silakan upload minimal 1 file input.")
    else:
        with st.spinner("Sedang memproses data..."):
            # Membuat folder virtual/sementara di server Streamlit
            with tempfile.TemporaryDirectory() as temp_dir:
                input_dir = os.path.join(temp_dir, 'input')
                output_dir = os.path.join(temp_dir, 'output')
                os.makedirs(input_dir)
                os.makedirs(output_dir)

                # Simpan objek memori Streamlit menjadi file fisik di folder virtual
                ref_path = os.path.join(temp_dir, ref_file.name)
                with open(ref_path, "wb") as f:
                    f.write(ref_file.getbuffer())

                for uploaded_file in input_files:
                    in_path = os.path.join(input_dir, uploaded_file.name)
                    with open(in_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                # --- LOGIKA PEMROSESAN (Sama persis dengan script lokal) ---
                try:
                    df_ref = pd.read_excel(ref_path, sheet_name='Sheet2', engine='openpyxl')
                    df_ref['KPJ'] = df_ref['KPJ'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                    df_ref = df_ref.drop_duplicates(subset=['KPJ'], keep='first')
                    ref_dict = df_ref.set_index('KPJ')[['NAMA_IBU_KANDUNG', 'HANDPHONE', 'EMAIL']].to_dict('index')
                    
                    processed_count = 0
                    
                    for file_name in os.listdir(input_dir):
                        file_path = os.path.join(input_dir, file_name)
                        output_file_name = os.path.splitext(file_name)[0] + '.xlsx'
                        output_path = os.path.join(output_dir, output_file_name)
                        
                        df_raw = None
                        
                        # Bypass xlrd korupsi / fallback HTML
                        if file_name.endswith('.xls'):
                            try:
                                wb = xlrd.open_workbook(file_path, ignore_workbook_corruption=True)
                                df_raw = pd.read_excel(wb, sheet_name='Koreksi Elemen TK', header=None, engine='xlrd')
                            except Exception:
                                try:
                                    dfs = pd.read_html(file_path, header=None)
                                    if dfs: df_raw = dfs[0]
                                except Exception: pass
                        
                        # Fallback eksekusi standard (untuk .xlsx / jika lapis pertama gagal)
                        if df_raw is None:
                            for engine_choice in ['openpyxl', 'xlrd', None]:
                                try:
                                    df_raw = pd.read_excel(file_path, sheet_name='Koreksi Elemen TK', header=None, engine=engine_choice)
                                    break
                                except Exception: continue
                        
                        if df_raw is None:
                            st.warning(f"Gagal membaca {file_name}")
                            continue

                        # Pemecahan template dan Eksekusi Rules
                        metadata = df_raw.iloc[0:1].copy()   
                        headers = df_raw.iloc[1:2].copy()    
                        data = df_raw.iloc[2:].copy()        
                        
                        data.columns = headers.iloc[0].values
                        data['KPJ*'] = data['KPJ*'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                        
                        # RULE: Kode Lokasi Pekerjaan
                        valid_lokasi = data['Lokasi Pekerjaan'].dropna().astype(str).str.strip()
                        valid_lokasi = valid_lokasi[~valid_lokasi.isin(['', 'nan', 'None'])]
                        
                        if not valid_lokasi.empty:
                            default_lokasi = valid_lokasi.mode().iloc[0]
                        else:
                            default_lokasi = '3515'
                            
                        for idx, row in data.iterrows():
                            kpj = row['KPJ*']
                            
                            lokasi = row.get('Lokasi Pekerjaan', None)
                            if pd.isna(lokasi) or str(lokasi).strip() in ['', 'nan', 'None']:
                                data.at[idx, 'Lokasi Pekerjaan'] = default_lokasi
                            
                            pkwt = str(row.get('PKWT', '')).strip().upper()
                            if pkwt == 'Y': data.at[idx, 'Tanggal Akhir Kontrak'] = '31-12-2026'
                            elif pkwt == 'T': data.at[idx, 'Tanggal Akhir Kontrak'] = ''

                            if kpj in ref_dict:
                                ibu = ref_dict[kpj]['NAMA_IBU_KANDUNG']
                                if pd.notna(ibu): data.at[idx, 'Nama Ibu Kandung'] = str(ibu).strip()
                                    
                                hp = ref_dict[kpj]['HANDPHONE']
                                if pd.notna(hp):
                                    hp_str = str(hp).replace('.0', '').strip()
                                    if hp_str: data.at[idx, 'Handphone'] = hp_str + '\t'
                                    
                                email = ref_dict[kpj]['EMAIL']
                                if pd.notna(email): data.at[idx, 'Email'] = str(email).strip()

                        # Reassembly & Simpan
                        metadata.columns = range(len(metadata.columns))
                        headers.columns = range(len(headers.columns))
                        data.columns = range(len(data.columns))
                        
                        final_df = pd.concat([metadata, headers, data], ignore_index=True)
                        final_df.to_excel(output_path, index=False, header=False, sheet_name='Koreksi Elemen TK', engine='openpyxl')
                        processed_count += 1
                        
                except Exception as e:
                    st.error(f"Terjadi kesalahan saat mengeksekusi data: {e}")

                # --- PROSES PEMBUATAN FILE ZIP OUTPUT ---
                if processed_count > 0:
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for root, _, files in os.walk(output_dir):
                            for file in files:
                                file_path = os.path.join(root, file)
                                zip_file.write(file_path, arcname=file) # Menyimpan file ke root ZIP tanpa folder luar
                    
                    st.success(f"✅ Selesai! Berhasil memproses {processed_count} file input.")
                    
                    st.download_button(
                        label="⬇️ Download Hasil Koreksi (.zip)",
                        data=zip_buffer.getvalue(),
                        file_name="Output_Koreksi_BPJS.zip",
                        mime="application/zip"
                    )
