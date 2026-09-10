import os
import pandas as pd
import warnings

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
        print(f"[ERROR] File referensi '{ref_file}' tidak ditemukan di direktori yang sama.")
        return

    # 2. Baca Referensi (Sheet2)
    print(f"\n[PROSES] Membaca file referensi: {ref_file}...")
    try:
        df_ref = pd.read_excel(ref_file, sheet_name='Sheet2')
    except Exception as e:
        print(f"[ERROR] Gagal membaca referensi: {e}")
        return

    # Normalisasi KPJ di referensi
    df_ref['KPJ'] = df_ref['KPJ'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    
    # Bersihkan duplikasi KPJ referensi
    before_count = len(df_ref)
    df_ref = df_ref.drop_duplicates(subset=['KPJ'], keep='first')
    after_count = len(df_ref)
    if before_count != after_count:
        print(f"[INFO] Ditemukan {before_count - after_count} baris KPJ duplikat di referensi. Telah dibersihkan.")

    ref_dict = df_ref.set_index('KPJ')[['NAMA_IBU_KANDUNG', 'HANDPHONE', 'EMAIL']].to_dict('index')
    print(f"[INFO] {len(ref_dict)} data KPJ referensi unik berhasil dimuat.")
    
    # 3. Looping File Input (Mendukung .xls dan .xlsx)
    input_files = [f for f in os.listdir(input_dir) if f.startswith('TEMPLATE_KOREKSI_TK_VLD_') and f.endswith(('.xls', '.xlsx'))]
    
    if not input_files:
        print(f"\n[WARNING] Tidak ada file input ditemukan di folder '{input_dir}/'.")
        return
        
    for file_name in input_files:
        file_path = os.path.join(input_dir, file_name)
        
        # SOLUSI: Ubah ekstensi output ke .xlsx agar menggunakan engine openpyxl (Bebas dari error xlwt & korupsi stream .xls)
        base_name_no_ext = os.path.splitext(file_name)[0]
        output_file_name = base_name_no_ext + '.xlsx'
        output_path = os.path.join(output_dir, output_file_name)
        
        print(f"\n[PROSES] Menganalisis file: {file_name}...")
        
        try:
            # Mencoba membaca file (mendukung engine otomatis untuk xlrd/openpyxl)
            df_raw = None
            read_errors = []
            
            # Coba baca dengan engine default / openpyxl / xlrd secara bergantian
            for engine_choice in [None, 'openpyxl', 'xlrd']:
                try:
                    df_raw = pd.read_excel(file_path, sheet_name='Koreksi Elemen TK', header=None, engine=engine_choice)
                    break
                except Exception as ex_eng:
                    read_errors.append(str(ex_eng))
            
            if df_raw is None:
                print(f"[ERROR] Gagal membaca {file_name}. File mungkin korup atau formatnya tidak valid.")
                continue

            # Pecah struktur file template
            metadata = df_raw.iloc[0:1].copy()   
            headers = df_raw.iloc[1:2].copy()    
            data = df_raw.iloc[2:].copy()        
            
            # Mapping nama kolom
            data.columns = headers.iloc[0].values
            data['KPJ*'] = data['KPJ*'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
            counter_updated = 0
            
            # 4. Iterasi Eksekusi VLOOKUP & Rules
            for idx, row in data.iterrows():
                kpj = row['KPJ*']
                
                # --- RULE 1: Lokasi Pekerjaan ---
                lokasi = row['Lokasi Pekerjaan']
                if pd.isna(lokasi) or str(lokasi).strip() in ['', 'nan', 'None']:
                    data.at[idx, 'Lokasi Pekerjaan'] = '3515'
                
                # --- RULE 2: Tanggal Akhir Kontrak berdasarkan PKWT ---
                pkwt = str(row['PKWT']).strip().upper()
                if pkwt == 'Y':
                    data.at[idx, 'Tanggal Akhir Kontrak'] = '31-12-2026'
                elif pkwt == 'T':
                    data.at[idx, 'Tanggal Akhir Kontrak'] = ''

                # --- RULE 3: VLOOKUP Data dari Referensi ---
                if kpj in ref_dict:
                    counter_updated += 1
                    
                    # Update Nama Ibu Kandung
                    ibu = ref_dict[kpj]['NAMA_IBU_KANDUNG']
                    if pd.notna(ibu):
                        data.at[idx, 'Nama Ibu Kandung'] = str(ibu).strip()
                        
                    # Update Handphone (Tab Delimited untuk Format Text)
                    hp = ref_dict[kpj]['HANDPHONE']
                    if pd.notna(hp):
                        hp_str = str(hp).replace('.0', '').strip()
                        if hp_str:
                            data.at[idx, 'Handphone'] = hp_str + '\t'
                        
                    # Update Email
                    email = ref_dict[kpj]['EMAIL']
                    if pd.notna(email):
                        data.at[idx, 'Email'] = str(email).strip()

            # 5. Gabungkan dan Ekspor File Baru (.xlsx)
            metadata.columns = range(len(metadata.columns))
            headers.columns = range(len(headers.columns))
            data.columns = range(len(data.columns))
            
            final_df = pd.concat([metadata, headers, data], ignore_index=True)
            
            # Simpan menggunakan engine openpyxl ke format .xlsx
            final_df.to_excel(output_path, index=False, header=False, sheet_name='Koreksi Elemen TK', engine='openpyxl')
            print(f"[SUKSES] Berhasil koreksi {counter_updated} baris.")
            print(f"[SUKSES] File tersimpan: {output_path}")
                
        except Exception as e:
            print(f"[ERROR] Gagal memproses file {file_name}: {e}")

if __name__ == "__main__":
    process_koreksi_data()
    print("\n"+"="*50)
    print(" SELESAI ")
    print("="*50)