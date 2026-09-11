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
    # Bersihkan spasi tersembunyi pada nama kolom agar pencarian nama header aman
    df_ref.columns = df_ref.columns.str.strip()

    # Cari posisi kolom 'KPJ'
    try:
        kpj_idx = df_ref.columns.get_loc('KPJ')
    except KeyError:
        print("[ERROR] Kolom 'KPJ' tidak ditemukan di Sheet2. Pastikan penulisan header benar.")
        return

    # Tentukan target kolom berdasarkan perhitungan index
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
    
    # Buat dictionary sebagai Mesin Pencari VLOOKUP
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
        
        # Penanganan format .xls bawaan sistem
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
        
        # Fallback jika gagal dengan xlrd / html
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
            # Memisahkan metadata, header, dan data asli
            metadata = df_raw.iloc[0:1].copy()   
            headers = df_raw.iloc[1:2].copy()    
            data = df_raw.iloc[2:].copy()        
            
            # Set kolom dengan header asli
            data.columns = headers.iloc[0].values
            
            # Membersihkan KPJ di file input
            data['KPJ*'] = data['KPJ*'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
            # --- PENENTUAN DEFAULT LOKASI PEKERJAAN ---
            valid_lokasi = data['Lokasi Pekerjaan'].dropna().astype(str).str.strip()
            valid_lokasi = valid_lokasi[~valid_lokasi.isin(['', 'nan', 'None'])]
            
            if not valid_lokasi.empty:
                default_lokasi = valid_lokasi.mode().iloc[0]
            else:
                default_lokasi = '3515'
            
            print(f"[INFO] Kode Lokasi default untuk file ini diatur ke: {default_lokasi}")

            # --- PROSES EKSEKUSI PENGISIAN DATA ---
            counter_updated = 0
            for idx, row in data.iterrows():
                kpj = row['KPJ*']
                
                # 1. Update Lokasi Pekerjaan
                lokasi = row.get('Lokasi Pekerjaan', None)
                if pd.isna(lokasi) or str(lokasi).strip() in ['', 'nan', 'None']:
                    data.at[idx, 'Lokasi Pekerjaan'] = default_lokasi
                
                # 2. Update Tanggal Akhir Kontrak (PKWT)
                pkwt = str(row.get('PKWT', '')).strip().upper()
                if pkwt == 'Y':
                    data.at[idx, 'Tanggal Akhir Kontrak'] = '31-12-2026'
                elif pkwt == 'T':
                    data.at[idx, 'Tanggal Akhir Kontrak'] = ''

                # 3. Eksekusi VLOOKUP berdasarkan Ref Dict
                if kpj in ref_dict:
                    counter_updated += 1
                    
                    # Isi Nama Ibu (dengan validasi anti 'nan')
                    ibu = ref_dict[kpj][col_ibu]
                    if pd.notna(ibu) and str(ibu).strip() != '':
                        data.at[idx, 'Nama Ibu Kandung'] = str(ibu).strip()
                        
                    # Isi No Handphone (Murni teks angka tanpa \t)
                    hp = ref_dict[kpj][col_hp]
                    if pd.notna(hp):
                        hp_str = str(hp).replace('.0', '').strip()
                        if hp_str and hp_str.lower() != 'nan':
                            data.at[idx, 'Handphone'] = hp_str
                        
                    # Isi Email
                    email = ref_dict[kpj][col_email]
                    if pd.notna(email) and str(email).strip() != '':
                        data.at[idx, 'Email'] = str(email).strip()

            # Kembalikan struktur kolom menjadi angka agar bisa disatukan kembali
            metadata.columns = range(len(metadata.columns))
            headers.columns = range(len(headers.columns))
            data.columns = range(len(data.columns))
            
            # Gabungkan metadata, header, dan data yang sudah dikoreksi
            final_df = pd.concat([metadata, headers, data], ignore_index=True)
            
            # Simpan file ke folder output
            final_df.to_excel(output_path, index=False, header=False, sheet_name='Koreksi Elemen TK', engine='openpyxl')
            
            print(f"[SUKSES] VLOOKUP berhasil menemukan dan mengubah {counter_updated} baris.")
            print(f"[SUKSES] Tersimpan -> {output_path}")
                
        except Exception as e:
            print(f"[ERROR] Kesalahan pada pemrosesan logika {file_name}: {e}")

if __name__ == "__main__":
    process_koreksi_data()
    print("\n"+"="*50)
    print(" SELESAI ")
