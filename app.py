import streamlit as st
import pandas as pd
import xlrd
import io
import warnings

# Abaikan peringatan formatting
warnings.filterwarnings('ignore', category=UserWarning)

st.title("Koreksi Data TK Aktif BPJS Ketenagakerjaan")

# 1. Upload File Referensi & Input File
ref_file = st.file_uploader("Upload File Referensi (Format .xlsx)", type=['xlsx'])
input_files = st.file_uploader("Upload File SMILE (Format .xls / .xlsx)", type=['xls', 'xlsx'], accept_multiple_files=True)

if st.button("Mulai Pemrosesan Data"):
    if not ref_file or not input_files:
        st.warning("Pastikan file referensi dan minimal 1 file input telah diunggah.")
        st.stop()

    # 2. Proses File Referensi
    with st.spinner("Membaca data referensi..."):
        try:
            df_ref = pd.read_excel(ref_file, sheet_name='Sheet2', engine='openpyxl')
            df_ref['KPJ'] = df_ref['KPJ'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            df_ref = df_ref.drop_duplicates(subset=['KPJ'], keep='first')
            ref_dict = df_ref.set_index('KPJ')[['NAMA_IBU_KANDUNG', 'HANDPHONE', 'EMAIL']].to_dict('index')
            st.success(f"Berhasil memuat {len(ref_dict)} data KPJ referensi unik.")
        except Exception as e:
            st.error(f"Gagal membaca file referensi: {e}")
            st.stop()

    # 3. Proses File Input
    for uploaded_file in input_files:
        file_name = uploaded_file.name
        file_bytes = uploaded_file.getvalue()
        df_raw = None
        
        st.write(f"---")
        st.write(f"**Memproses: {file_name}**")
        
        # Penanganan khusus ekstensi .xls via object Stream (BytesIO)
        if file_name.endswith('.xls'):
            try:
                # Bypass xlrd corruption via memory stream
                wb = xlrd.open_workbook(file_contents=file_bytes, ignore_workbook_corruption=True)
                df_raw = pd.read_excel(wb, sheet_name='Koreksi Elemen TK', header=None, engine='xlrd')
            except Exception:
                # Fallback HTML parser via memory stream
                try:
                    dfs = pd.read_html(io.BytesIO(file_bytes), header=None)
                    if dfs:
                        df_raw = dfs[0]
                except Exception:
                    pass
        
        # Eksekusi standar file normal/xlsx
        if df_raw is None:
            for engine_choice in ['openpyxl', 'xlrd', None]:
                try:
                    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name='Koreksi Elemen TK', header=None, engine=engine_choice)
                    break
                except Exception:
                    continue
        
        if df_raw is None:
            st.error(f"Gagal membaca data dari {file_name}. File mungkin rusak.")
            continue

        try:
            # Pecah metadata dan header
            metadata = df_raw.iloc[0:1].copy()   
            headers = df_raw.iloc[1:2].copy()    
            data = df_raw.iloc[2:].copy()        
            
            data.columns = headers.iloc[0].values
            data['KPJ*'] = data['KPJ*'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
            # Tentukan default lokasi terbanyak
            valid_lokasi = data['Lokasi Pekerjaan'].dropna().astype(str).str.strip()
            valid_lokasi = valid_lokasi[~valid_lokasi.isin(['', 'nan', 'None'])]
            default_lokasi = valid_lokasi.mode().iloc[0] if not valid_lokasi.empty else '3515'
            
            counter_updated = 0
            
            # Iterasi Data
            for idx, row in data.iterrows():
                kpj = row['KPJ*']
                
                # Rule 1: Lokasi
                lokasi = row.get('Lokasi Pekerjaan', None)
                if pd.isna(lokasi) or str(lokasi).strip() in ['', 'nan', 'None']:
                    data.at[idx, 'Lokasi Pekerjaan'] = default_lokasi
                
                # Rule 2: PKWT
                pkwt = str(row.get('PKWT', '')).strip().upper()
                if pkwt == 'Y':
                    data.at[idx, 'Tanggal Akhir Kontrak'] = '31-12-2026'
                elif pkwt == 'T':
                    data.at[idx, 'Tanggal Akhir Kontrak'] = ''

                # Rule 3: VLOOKUP
                if kpj in ref_dict:
                    counter_updated += 1
                    
                    ibu = ref_dict[kpj]['NAMA_IBU_KANDUNG']
                    if pd.notna(ibu):
                        data.at[idx, 'Nama Ibu Kandung'] = str(ibu).strip()
                        
                    hp = ref_dict[kpj]['HANDPHONE']
                    if pd.notna(hp):
                        hp_str = str(hp).replace('.0', '').strip()
                        if hp_str:
                            data.at[idx, 'Handphone'] = hp_str + '\t'
                        
                    email = ref_dict[kpj]['EMAIL']
                    if pd.notna(email):
                        data.at[idx, 'Email'] = str(email).strip()

            # Gabungkan dan jadikan format BytesIO untuk diunduh
            metadata.columns = range(len(metadata.columns))
            headers.columns = range(len(headers.columns))
            data.columns = range(len(data.columns))
            
            final_df = pd.concat([metadata, headers, data], ignore_index=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, header=False, sheet_name='Koreksi Elemen TK')
            
            output.seek(0)
            
            st.info(f"Koreksi berhasil: {counter_updated} baris diupdate (Default Lokasi: {default_lokasi}).")
            
            # Tombol Download per file
            output_filename = file_name.rsplit('.', 1)[0] + ".xlsx"
            st.download_button(
                label=f"📥 Download {output_filename}",
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses isi {file_name}: {e}")
