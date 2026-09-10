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

    # 2. Baca Referensi
    print(f"\n[PROSES] Membaca file referensi: {ref_file}...")
    try:
        df_ref = pd.read_excel(ref_file, sheet_name='Sheet2', engine='openpyxl')
    except Exception as e:
        print(f"[ERROR] Gagal membaca referensi: {e}")
        return

    df_ref['KPJ'] = df_ref['KPJ'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df_ref = df_ref.drop_duplicates(subset=['KPJ'], keep='first')
    ref_dict = df_ref.set_index('KPJ')[['NAMA_IBU_KANDUNG', 'HANDPHONE', 'EMAIL']].to_dict('index')
    print(f"[INFO] {len(ref_dict)} data KPJ referensi unik berhasil dimuat.")
    
    # 3. Looping File Input
    input_files = [f for f in os.listdir(input_dir) if f.startswith('TEMPLATE_KOREKSI_TK_VLD_') and f.endswith(('.xls', '.xlsx'))]
    
    if not input_files:
        print(f"\n[WARNING] Tidak ada file input ditemukan di folder '{input_dir}/'.")
        return
        
    for file_name in input_files:
        file_path = os.path.join(input_dir, file_name)
        output_file_name = os.path.splitext(file_name)[0] + '.xlsx'
        output_path = os.path.join(output_dir, output_file_name)
        
        print(f"\n[PROSES] Menganalisis file: {file_name}...")
        df_raw = None
        
        if file_name.endswith('.xls'):
            try:
                wb = xlrd.open_workbook(file_path, ignore_workbook_corruption=True)
                df_raw = pd.read_excel(wb, sheet_name='Koreksi Elemen TK', header=None, engine='xlrd')
                print(f"[INFO] Berhasil membaca biner .xls asli.")
            except Exception:
                try:
                    dfs = pd.read_html(file_path, header=None)
                    if dfs:
                        df_raw = dfs[0]
                        print(f"[INFO] Berhasil membaca menggunakan parser HTML.")
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
            print(f"[ERROR] Gagal membaca {file_name} secara keseluruhan.")
            continue

        try:
            metadata = df_raw.iloc[0:1].copy()   
            headers = df_raw.iloc[1:2].copy()    
            data = df_raw.iloc[2:].copy()        
            
            data.columns = headers.iloc[0].values
            data['KPJ*'] = data['KPJ*'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
            # --- RULE BARU: Penentuan Default Lokasi Pekerjaan ---
            valid_lokasi = data['Lokasi Pekerjaan'].dropna().astype(str).str.strip()
            valid_lokasi = valid_lokasi[~valid_lokasi.isin(['', 'nan', 'None'])]
            
            if not valid_lokasi.empty:
                default_lokasi = valid_lokasi.mode().iloc[0]
            else:
                default_lokasi = '3515'
            
            print(f"[INFO] Kode Lokasi default untuk file ini diatur ke: {default_lokasi}")

            counter_updated = 0
            for idx, row in data.iterrows():
                kpj = row['KPJ*']
                
                # --- Eksekusi Pengisian Lokasi Pekerjaan ---
                lokasi = row.get('Lokasi Pekerjaan', None)
                if pd.isna(lokasi) or str(lokasi).strip() in ['', 'nan', 'None']:
                    data.at[idx, 'Lokasi Pekerjaan'] = default_lokasi
                
                pkwt = str(row.get('PKWT', '')).strip().upper()
                if pkwt == 'Y':
                    data.at[idx, 'Tanggal Akhir Kontrak'] = '31-12-2026'
                elif pkwt == 'T':
                    data.at[idx, 'Tanggal Akhir Kontrak'] = ''

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

            metadata.columns = range(len(metadata.columns))
            headers.columns = range(len(headers.columns))
            data.columns = range(len(data.columns))
            
            final_df = pd.concat([metadata, headers, data], ignore_index=True)
            final_df.to_excel(output_path, index=False, header=False, sheet_name='Koreksi Elemen TK', engine='openpyxl')
            
            print(f"[SUKSES] Berhasil koreksi {counter_updated} baris.")
            print(f"[SUKSES] Tersimpan -> {output_path}")
                
        except Exception as e:
            print(f"[ERROR] Kesalahan pada pemrosesan logika {file_name}: {e}")

if __name__ == "__main__":
    process_koreksi_data()
    print("\n"+"="*50)
    print(" SELESAI ")
    print("="*50)
