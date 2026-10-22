from datetime import datetime
import glob
import os
import re
import openpyxl
import pandas as pd

DATA_DIR = "."

START_DATE = datetime.strptime("2026-09-16", "%Y-%m-%d").date()
END_DATE = datetime.strptime("2026-10-16", "%Y-%m-%d").date()

pattern = re.compile(r".*laporan_padel_dki_jakarta_(\d{4}-\d{2}-\d{2})\.xlsx$")

all_files = glob.glob(
    os.path.join(DATA_DIR, "laporan_padel_dki_jakarta_*.xlsx")
)
df_list = []
matched_files = []

print("Mencari dan memproses file Excel...")

for file_path in all_files:
  filename = os.path.basename(file_path)
  match = pattern.match(filename)

  if match:
    file_date_str = match.group(1)
    file_date = datetime.strptime(file_date_str, "%Y-%m-%d").date()

    if START_DATE <= file_date <= END_DATE:
      print(f"-> Membaca: {filename}")
      df = pd.read_excel(file_path)

      if "tanggal" not in df.columns:
        df["tanggal"] = file_date

      df_list.append(df)
      matched_files.append((file_date, filename))

if not df_list:
  print("Tidak ditemukan file yang sesuai rentang tanggal!")
else:
  df_list_sorted = [
      x
      for _, x in sorted(
          zip([d[0] for d in matched_files], df_list), key=lambda pair: pair[0]
      )
  ]

  merged_df = pd.concat(df_list_sorted, ignore_index=True)

  col_revenue = "revenue"
  if col_revenue in merged_df.columns:
    total_rev = merged_df[col_revenue].sum()
    print("\n" + "=" * 50)
    print(f"Berhasil menggabungkan : {len(matched_files)} file")
    print(f"Total baris transaksi   : {len(merged_df)}")
    print(f"Total Revenue 1 Bulan   : Rp {total_rev:,.2f}")
    print("=" * 50)

  output_filename = "laporan_padel_dki_jakarta_revenue_1bulan_total.xlsx"
  merged_df.to_excel(output_filename, index=False)
  print(f"\nFile rangkuman 1 bulan berhasil disimpan ke: {output_filename}")
