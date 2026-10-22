from datetime import datetime
import glob
import os
import re
import pandas as pd

START_DATE = datetime.strptime("2026-09-16", "%Y-%m-%d").date()
END_DATE = datetime.strptime("2026-10-16", "%Y-%m-%d").date()

pattern = re.compile(r".*laporan_padel_(\d{4}-\d{2}-\d{2})_final\.csv$")

all_files = glob.glob("laporan_padel_*_final.csv")
df_list = []
matched_files = []

print(f"Total file csv terdeteksi di repositori: {len(all_files)}")

for file_path in all_files:
  filename = os.path.basename(file_path)
  match = pattern.match(filename)

  if match:
    file_date_str = match.group(1)
    file_date = datetime.strptime(file_date_str, "%Y-%m-%d").date()

    if START_DATE <= file_date <= END_DATE:
      print(f"-> Membaca: {filename}")
      df = pd.read_csv(file_path)

      if "tanggal" not in df.columns:
        df["tanggal"] = file_date

      df_list.append(df)
      matched_files.append((file_date, filename))

if not df_list:
  print("Tidak ditemukan file yang masuk rentang 16 Sep - 16 Okt 2026!")
else:
  df_list_sorted = [
      x
      for _, x in sorted(
          zip([d[0] for d in matched_files], df_list), key=lambda pair: pair[0]
      )
  ]
  merged_df = pd.concat(df_list_sorted, ignore_index=True)

  output_filename = "laporan_padel_dki_jakarta_revenue_1bulan_total.csv"
  merged_df.to_csv(output_filename, index=False)

  print("\n" + "=" * 50)
  print(f"Berhasil menggabungkan : {len(matched_files)} file")
  print(f"Total baris data        : {len(merged_df)}")
  print(f"Hasil disimpan ke       : {output_filename}")
  print("=" * 50)
