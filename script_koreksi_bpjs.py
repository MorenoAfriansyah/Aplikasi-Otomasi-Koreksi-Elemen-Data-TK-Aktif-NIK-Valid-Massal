import os
import pandas as pd
import warnings
import xlrd

# Mengabaikan warning styling Excel
warnings.filterwarnings('ignore', category=UserWarning)

def process_koreksi_data():
    input_dir = 'input'
    output_dir = 'output'
    ref_file = 'TESTING TEMPLATE vlookup.xlsx'
    
    print("="*50)
    print(" PROGRAM KOREKSI DATA TK AKTIF BPJS KETENAGAKERJAAN ")
    print("="*50)

    # 1. Setup Direktori
    for directory in [input_dir, output_dir]:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print(f"[INFO] Folder '{directory}' dibuat.")
            
    if not os.path.exists(ref_file):
        print(f"[ERROR] File referensi '{ref_file}' tidak ditemukan.")
        return

    # 2. Baca File Referensi (Sheet 2)
    print(f"\n[PROSES] Membaca file referensi: {ref_file}...")
    try:
        df_ref = pd.read_excel(ref_file, sheet_name='Sheet2', engine='openpyxl')
    except Exception as e:
        print(f"[ERROR] Gagal membaca referensi: {e}")
        return

    # --- LOGIKA VLOOKUP INDEX 9, 10, 11 ---
    df_ref.columns = df_ref.columns.str.strip()

    try:
        kpj_idx = df_ref.columns.get_loc('KPJ')
    except KeyError:
        print("[ERROR] Kolom 'KPJ' tidak ditemukan di Sheet2. Pastikan penulisan header benar.")
        return

    try:
        col_ibu = df_ref.columns[kpj_idx + 8]
        col_hp = df_ref.columns[kpj_idx + 9]
        col_email = df_ref.columns[kpj_idx + 10]
    except IndexError:
        print("[ERROR] Jumlah kolom di Sheet2 kurang. Tidak dapat mencapai index 9, 10, atau 11.")
        return

    print(f"[INFO] Kolom target VLOOKUP ditemukan: ")
    print(f"       - Index 9  : {col_ibu}")
    print(f"       - Index 10 : {col_hp}")
    print(f"       - Index 11 : {col_email}")

    # Membersihkan isi data KPJ di file referensi
    df_ref['KPJ'] = df_ref['KPJ'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df_ref = df_ref.drop_duplicates(subset=['KPJ'], keep='first')
    
    ref_dict = df_ref.set_index('KPJ')[[col_ibu, col_hp, col_email]].to_dict('index')
    print(f"[INFO] {len(ref_dict)} data referensi KPJ unik berhasil dimuat dan siap dicocokkan.")
    
    # 3. Looping File Input
    input_files = [f for f in os.listdir(input_dir) if f.startswith('TEMPLATE_KOREKSI_TK_VLD_') and f.endswith(('.xls', '.xlsx'))]
    
    if not input_files:
        print(f"\n[WARNING] Tidak ada file input ditemukan di folder '{input_dir}/'.")
        return
        
    for file_name in input_files:
        file_path = os.path.join(input_dir, file_name)
        output_file_name = os.path.splitext(file_name)[0] + '.xlsx'
        output_path = os.path.join(output_dir, output_file_name)
        
        print(f"\n[PROSES] Menganalisis file input: {file_name}...")
        df_raw = None
        
        if file_name.endswith('.xls'):
            try:
                wb = xlrd.open_workbook(file_path, ignore_workbook_corruption=True)
                df_raw = pd.read_excel(wb, sheet_name='Koreksi Elemen TK', header=None, engine='xlrd')
            except Exception:
                try:
                    dfs = pd.read_html(file_path, header=None)
                    if dfs:
                        df_raw = dfs[0]
                except Exception:
                    pass
        
        if df_raw is None:
            for engine_choice in ['openpyxl', 'xlrd', None]:
                try:
                    df_raw = pd.read_excel(file_path, sheet_name='Koreksi Elemen TK', header=None, engine=engine_choice)
                    break
                except Exception:
                    continue
        
        if df_raw is None:
            print(f"[ERROR] Gagal membaca {file_name} secara keseluruhan. Lewati file ini.")
            continue

        try:
            metadata = df_raw.iloc[0:1].copy()   
            headers = df_raw.iloc[1:2].copy()    
            data = df_raw.iloc[2:].copy()        
            
            data.columns = headers.iloc[0].values
            data['KPJ*'] = data['KPJ*'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
            # --- PENENTUAN DEFAULT LOKASI PEKERJAAN ---
            valid_lokasi = data['Lokasi Pekerjaan'].dropna().astype(str).str.strip()
            valid_lokasi = valid_lokasi[~valid_lokasi.isin(['', 'nan', 'None'])]
            
            if not valid_lokasi.empty:
                default_lokasi = valid_lokasi.mode().iloc[0]
            else:
                default_lokasi = '3515'
            
            print(f"[INFO] Kode Lokasi default untuk file ini diatur ke: {default_lokasi}")

            # --- PROSES EKSEKUSI & PENYARINGAN BARIS ---
            valid_rows = []
            counter_updated = 0
            
            for idx, row in data.iterrows():
                kpj = row['KPJ*']
                
                # Jika KPJ tidak ditemukan di referensi, baris dibuang
                if kpj not in ref_dict:
                    continue
                
                ibu = ref_dict[kpj][col_ibu]
                hp = ref_dict[kpj][col_hp]
                email = ref_dict[kpj][col_email]
                
                # Validasi ketat: Nama Ibu Kandung WAJIB ada dan valid
                is_ibu_valid = pd.notna(ibu) and str(ibu).strip() != '' and str(ibu).strip().lower() != 'nan'
                
                if not is_ibu_valid:
                    # Jika Nama Ibu Kandung tidak ada, baris ini dibuang total (tidak ditampilkan)
                    continue
                
                counter_updated += 1
                
                # 1. Update Lokasi Pekerjaan
                lokasi = row.get('Lokasi Pekerjaan', None)
                if pd.isna(lokasi) or str(lokasi).strip() in ['', 'nan', 'None']:
                    row['Lokasi Pekerjaan'] = default_lokasi
                
                # 2. Update Tanggal Akhir Kontrak (PKWT)
                pkwt = str(row.get('PKWT', '')).strip().upper()
                if pkwt == 'Y':
                    row['Tanggal Akhir Kontrak'] = '31-12-2026'
                elif pkwt == 'T':
                    row['Tanggal Akhir Kontrak'] = ''

                # 3. Update Nama Ibu Kandung
                row['Nama Ibu Kandung'] = str(ibu).strip()
                
                # 4. Update Handphone (jika ada, jika tidak dikosongkan)
                if pd.notna(hp):
                    hp_str = str(hp).replace('.0', '').strip()
                    if hp_str and hp_str.lower() != 'nan':
                        row['Handphone'] = hp_str
                    else:
                        row['Handphone'] = ''
                else:
                    row['Handphone'] = ''
                    
                # 5. Update Email (jika ada, jika tidak dikosongkan)
                if pd.notna(email):
                    email_str = str(email).strip()
                    if email_str and email_str.lower() != 'nan':
                        row['Email'] = email_str
                    else:
                        row['Email'] = ''
                else:
                    row['Email'] = ''
                
                valid_rows.append(row)

            # Buat DataFrame dari baris yang lolos validasi
            if valid_rows:
                data_filtered = pd.DataFrame(valid_rows)
            else:
                data_filtered = pd.DataFrame(columns=data.columns)

            metadata.columns = range(len(metadata.columns))
            headers.columns = range(len(headers.columns))
            data_filtered.columns = range(len(data_filtered.columns))
            
            final_df = pd.concat([metadata, headers, data_filtered], ignore_index=True)
            final_df.to_excel(output_path, index=False, header=False, sheet_name='Koreksi Elemen TK', engine='openpyxl')
            
            print(f"[SUKSES] Berhasil memproses dan menyaring {counter_updated} baris valid.")
            print(f"[SUKSES] Tersimpan -> {output_path}")
                
        except Exception as e:
            print(f"[ERROR] Kesalahan pada pemrosesan logika {file_name}: {e}")

if __name__ == "__main__":
    process_koreksi_data()
    print("\n"+"="*50)
    print(" SELESAI ")
