import os
import re
from datetime import datetime, timedelta
import openpyxl
import pandas as pd

START_DATE = datetime.strptime("2026-09-16", "%Y-%m-%d").date()
END_DATE = datetime.strptime("2026-10-16", "%Y-%m-%d").date()

dates_list = []
curr = START_DATE
while curr <= END_DATE:
  dates_list.append(curr.strftime("%Y-%m-%d"))
  curr += timedelta(days=1)


def clean_currency_to_number(val):
  """Membersihkan format 'Rp1.940.000' atau angka biasa menjadi integer murni."""
  if pd.isna(val):
    return 0
  numeric_str = re.sub(r"[^\d]", "", str(val))
  return int(numeric_str) if numeric_str else 0


records = []
print(
    f"Membaca file harian dari {dates_list[0]} sampai {dates_list[-1]}..."
)

for date_str in dates_list:
  candidates = [
      f"laporan_padel_dki_jakarta_{date_str}.xlsx",
      f"laporan_padel_{date_str}.xlsx",
      f"laporan_padel_{date_str}_final.csv",
      f"laporan_padel_{date_str}.csv",
  ]

  found_file = next((f for f in candidates if os.path.exists(f)), None)
  if not found_file:
    print(f"[LEWAT] File {date_str} tidak ditemukan.")
    continue

  print(f"-> Membaca: {found_file}")

  if found_file.endswith(".xlsx"):
    xls = pd.ExcelFile(found_file)
    for sheet_name in xls.sheet_names:
      if "TOTAL" in sheet_name.upper():
        continue

      df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
      if df_sheet.empty or "Venue" not in df_sheet.columns:
        continue

      for _, row in df_sheet.iterrows():
        venue = row.get("Venue")
        if pd.isna(venue) or "TOTAL" in str(venue).upper():
          continue

        raw_revenue = row.get("Revenue", 0)
        clean_revenue = clean_currency_to_number(raw_revenue)

        records.append({
            "Wilayah": str(sheet_name).strip(),
            "Venue": str(venue).strip(),
            "Tanggal": date_str,
            "Revenue": clean_revenue,
        })
  else:
    df_csv = pd.read_csv(found_file)
    if "Venue" in df_csv.columns:
      for _, row in df_csv.iterrows():
        venue = row.get("Venue")
        if pd.isna(venue) or "TOTAL" in str(venue).upper():
          continue

        wilayah = (
            row.get("Wilayah") if "Wilayah" in row else "Jakarta Barat"
        )
        raw_revenue = row.get("Revenue", 0)
        clean_revenue = clean_currency_to_number(raw_revenue)

        records.append({
            "Wilayah": str(wilayah).strip(),
            "Venue": str(venue).strip(),
            "Tanggal": date_str,
            "Revenue": clean_revenue,
        })

if not records:
  print("Tidak ada data transaksi yang berhasil diekstrak!")
  exit(1)

full_df = pd.DataFrame(records)

output_filename = "laporan_padel_dki_jakarta_16Sep_16Okt_lengkap.xlsx"
writer = pd.ExcelWriter(output_filename, engine="openpyxl")

wilayah_order = sorted(full_df["Wilayah"].unique())

for wil in wilayah_order:
  sub_df = full_df[full_df["Wilayah"] == wil]

  pivot_df = sub_df.pivot_table(
      index="Venue",
      columns="Tanggal",
      values="Revenue",
      aggfunc="sum",
      fill_value=0,
  )

  for d in dates_list:
    if d not in pivot_df.columns:
      pivot_df[d] = 0
  pivot_df = pivot_df[dates_list]

  pivot_df["Revenue Total"] = pivot_df.sum(axis=1)
  pivot_df = pivot_df.reset_index()

  pivot_df.insert(0, "No", range(1, len(pivot_df) + 1))
  pivot_df.rename(columns={"Venue": "Nama Venue"}, inplace=True)

  total_row = {"No": "", "Nama Venue": "TOTAL REVENUE"}
  for col in dates_list:
    total_row[col] = pivot_df[col].sum()
  total_row["Revenue Total"] = pivot_df["Revenue Total"].sum()

  pivot_df = pd.concat([pivot_df, pd.DataFrame([total_row])], ignore_index=True)

  sheet_name = str(wil)[:31]
  pivot_df.to_excel(writer, sheet_name=sheet_name, index=False)
  print(f"Sheet '{sheet_name}' berhasil diproses: {len(pivot_df)-1} venue.")

writer.close()
print("\n" + "=" * 55)
print(f"Selesai! Hasil akhir tersimpan di: {output_filename}")
print("=" * 55)
