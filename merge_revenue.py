import os
from datetime import datetime, timedelta
import pandas as pd

START_DATE = datetime.strptime("2026-09-16", "%Y-%m-%d").date()
END_DATE = datetime.strptime("2026-10-16", "%Y-%m-%d").date()

target_dates = []
curr = START_DATE
while curr <= END_DATE:
    target_dates.append(curr)
    curr += timedelta(days=1)

df_list = []
found_count = 0

print(f"Memeriksa file transaksi dari {START_DATE} hingga {END_DATE}...")

for d in target_dates:
    date_str = d.strftime("%Y-%m-%d")
    
    candidates = [
        f"laporan_padel_dki_jakarta_{date_str}.xlsx",
        f"laporan_padel_{date_str}_final.csv",
        f"laporan_padel_{date_str}.xlsx",
        f"laporan_padel_{date_str}.csv"
    ]
    
    selected_file = None
    for fname in candidates:
        if os.path.exists(fname):
            selected_file = fname
            break
            
    if selected_file:
        print(f"-> Membaca data: {selected_file}")
        try:
            if selected_file.endswith(".xlsx"):
                df = pd.read_excel(selected_file)
            else:
                df = pd.read_csv(selected_file)
                
            if "tanggal" not in df.columns:
                df.insert(0, "tanggal", date_str)
                
            df_list.append(df)
            found_count += 1
        except Exception as e:
            print(f"Gagal memproses file {selected_file}: {e}")
    else:
        print(f"[LEWAT] File tanggal {date_str} tidak ditemukan.")

if not df_list:
    print("\nTidak ada file dalam rentang 16 Sep - 16 Okt yang ditemukan!")
    exit(1)

merged_df = pd.concat(df_list, ignore_index=True)

revenue_col = None
candidate_cols = ["revenue", "total_harga", "omzet", "pendapatan", "nominal", "harga", "total_pendapatan"]
for col in merged_df.columns:
    if str(col).strip().lower() in candidate_cols:
        revenue_col = col
        break

if revenue_col is None:
    numeric_cols = merged_df.select_dtypes(include=["number"]).columns
    revenue_col = numeric_cols[-1] if len(numeric_cols) > 0 else merged_df.columns[-1]

merged_df[revenue_col] = pd.to_numeric(merged_df[revenue_col], errors="coerce").fillna(0)
total_pendapatan = merged_df[revenue_col].sum()

summary_row = {col: "" for col in merged_df.columns}
first_col = merged_df.columns[0]
summary_row[first_col] = "TOTAL REVENUE (16 SEP - 16 OKT)"
summary_row[revenue_col] = total_pendapatan

final_output_df = pd.concat([merged_df, pd.DataFrame([summary_row])], ignore_index=True)

output_filename = "laporan_padel_dki_jakarta_16Sep_16Okt_lengkap.xlsx"
final_output_df.to_excel(output_filename, index=False)

print("\n" + "=" * 55)
print(f"Total file diproses   : {found_count} hari")
print(f"Total baris transaksi : {len(merged_df)}")
print(f"Total Pendapatan      : Rp {total_pendapatan:,.2f}")
print(f"File berhasil dibuat  : {output_filename}")
print("=" * 55)
