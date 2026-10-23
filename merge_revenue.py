from datetime import datetime, timedelta
import os
import re
import openpyxl
import pandas as pd

START_DATE = datetime.strptime("2026-09-16", "%Y-%m-%d").date()
END_DATE = datetime.strptime("2026-10-16", "%Y-%m-%d").date()

dates_list = []
curr = START_DATE
while curr <= END_DATE:
  dates_list.append(curr.strftime("%Y-%m-%d"))
  curr += timedelta(days=1)

records = []

def tentukan_wilayah(venue_name):
  v = str(venue_name).lower()
  if any(
      k in v
      for k in [
          "selatan",
          "cilandak",
          "kemang",
          "senopati",
          "pondok indah",
          "kebayoran",
          "scbd",
          "fatmawati",
      ]
  ):
    return "Jakarta Selatan"
  if any(k in v for k in ["utara", "pik", "pluit", "kelapa gading", "sunter"]):
    return "Jakarta Utara"
  if any(
      k in v
      for k in [
          "pusat",
          "menteng",
          "sudirman",
          "thamrin",
          "tanah abang",
          "monas",
      ]
  ):
    return "Jakarta Pusat"
  if any(
      k in v
      for k in ["timur", "rawamangun", "cakung", "duren sawit", "pulomas"]
  ):
    return "Jakarta Timur"
  return "Jakarta Barat"


print("Membaca dan memproses laporan harian...")

for date_str in dates_list:
  candidates = [
      f"laporan_padel_dki_jakarta_{date_str}.xlsx",
      f"laporan_padel_{date_str}.xlsx",
      f"laporan_padel_{date_str}_final.csv",
      f"laporan_padel_{date_str}.csv",
  ]

  found_file = None
  for f in candidates:
    if os.path.exists(f):
      found_file = f
      break

  if not found_file:
    print(f"[LEWAT] File tanggal {date_str} tidak ditemukan.")
    continue

  print(f"-> Memproses: {found_file}")

  if found_file.endswith(".xlsx"):
    xls = pd.ExcelFile(found_file)
    for sheet_name in xls.sheet_names:
      df_temp = pd.read_excel(xls, sheet_name=sheet_name)
      if df_temp.empty or "Venue" not in df_temp.columns:
        continue

      wilayah = sheet_name if len(xls.sheet_names) > 1 else None

      for _, row in df_temp.iterrows():
        v_name = row.get("Venue")
        if pd.isna(v_name) or "TOTAL" in str(v_name).upper():
          continue

        w = (
            wilayah
            if wilayah
            else (
                row.get("Wilayah")
                if "Wilayah" in row
                else tentukan_wilayah(v_name)
            )
        )
        rev = row.get("Revenue", 0)

        records.append({
            "Wilayah": w,
            "Venue": str(v_name).strip(),
            "Tanggal": date_str,
            "Revenue": (
                pd.to_numeric(rev, errors="coerce")
                if pd.notnull(pd.to_numeric(rev, errors="coerce"))
                else 0
            ),
        })

  else:
    df_temp = pd.read_csv(found_file)
    if "Venue" in df_temp.columns:
      for _, row in df_temp.iterrows():
        v_name = row.get("Venue")
        if pd.isna(v_name) or "TOTAL" in str(v_name).upper():
          continue

        w = (
            row.get("Wilayah")
            if "Wilayah" in row
            else tentukan_wilayah(str(v_name))
        )
        rev = row.get("Revenue", 0)

        records.append({
            "Wilayah": w,
            "Venue": str(v_name).strip(),
            "Tanggal": date_str,
            "Revenue": (
                pd.to_numeric(rev, errors="coerce")
                if pd.notnull(pd.to_numeric(rev, errors="coerce"))
                else 0
            ),
        })

if not records:
  print("Tidak ada data yang berhasil diekstrak!")
  exit(1)

full_data = pd.DataFrame(records)

output_filename = "laporan_padel_dki_jakarta_16Sep_16Okt_lengkap.xlsx"
writer = pd.ExcelWriter(output_filename, engine="openpyxl")

wilayah_list = sorted(full_data["Wilayah"].unique())

for wil in wilayah_list:
  sub_df = full_data[full_data["Wilayah"] == wil]

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

  sheet_name_clean = str(wil)[:31]
  pivot_df.to_excel(writer, sheet_name=sheet_name_clean, index=False)
  print(f"-> Sheet '{sheet_name_clean}' selesai dibuat ({len(pivot_df)-1} venue)")

writer.close()
print("\n" + "=" * 55)
print(f"File multi-sheet berhasil dibuat: {output_filename}")
print("=" * 55)
