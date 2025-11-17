import pandas as pd
import holidays
from datetime import date, timedelta

# ------------------------------------------------------------------
# BAGIAN 1: KONFIGURASI (KHUSUS TAHUN 2025)
# ------------------------------------------------------------------

# 1. Tentukan rentang tahun (Hanya 2025)
YEARS_TO_GET = [2025]

# 2. Tentukan MUSIM F&B (Rentang Waktu) untuk 2025.
# Catatan: Anda TIDAK PERLU menulis "Tahun Baru", "Imlek", "Nyepi" di sini.
# Program sudah otomatis mengambilnya dari 'YEARS_TO_GET'.
# Cukup masukkan musim libur sekolah atau puasa yang panjang.

CUSTOM_SEASONS = [
    # Format: (Mulai, Selesai, Nama Event, Tipe Event)
    # Musim Puasa (Perkiraan)
    ("2025-03-01", "2025-03-30", "Musim Ramadan 2025", "Musim Keagamaan"),
    # Musim Libur Sekolah (Perkiraan - Sesuaikan dengan provinsi Anda)
    ("2025-06-21", "2025-07-13", "Libur Sekolah (Juni) 2025", "Musim Liburan"),
    # Musim Akhir Tahun
    ("2025-12-20", "2025-12-31", "Libur Nataru 2025", "Musim Liburan"),
    ("2025-01-27", "2025-01-27", "Isra Mikraj 2025", "Libur Nasional"),
    ("2025-01-29", "2025-01-29", "Tahun Baru Imlek 2025", "Libur Nasional"),
    # (Opsional: Cuti Bersama Imlek jika mau dimasukkan juga)
    ("2025-01-28", "2025-01-28", "Cuti Bersama Imlek", "Libur Nasional"),
]

# 3. Nama file output
OUTPUT_FILENAME = "kalender_event1.csv"

# ------------------------------------------------------------------
# BAGIAN 2: PROSES OTOMATIS (Tidak perlu diubah)
# ------------------------------------------------------------------


def expand_date_range(start_date_str, end_date_str, event_name, event_type):
    """Membantu mengubah rentang tanggal (Mulai-Selesai) menjadi daftar harian."""
    dates = []
    try:
        current_date = date.fromisoformat(start_date_str)
        end_date = date.fromisoformat(end_date_str)
    except ValueError:
        print(
            f"ERROR: Format tanggal salah untuk event '{event_name}'. Gunakan YYYY-MM-DD."
        )
        return []

    while current_date <= end_date:
        dates.append(
            {
                "Tanggal": current_date,
                "Nama_Event": event_name,
                "Tipe_Event": event_type,
            }
        )
        current_date += timedelta(days=1)
    return dates


print("Memulai program pembuat kalender event (Khusus 2025)...")

# --- LANGKAH 1: Ambil Libur Nasional (Otomatis) ---
print(f"Mengambil data 'Libur Nasional' Indonesia untuk tahun: {YEARS_TO_GET}...")
national_holidays_data = []
try:
    # Ambil libur otomatis
    indonesia_holidays = holidays.ID(years=YEARS_TO_GET, include_public=True)

    for tanggal, nama in indonesia_holidays.items():
        national_holidays_data.append(
            {"Tanggal": tanggal, "Nama_Event": nama, "Tipe_Event": "Libur Nasional"}
        )
    print(
        f"Ditemukan {len(national_holidays_data)} hari Libur Nasional & Cuti Bersama otomatis."
    )
except Exception as e:
    print(f"Gagal mengambil data liburan: {e}")
    print("Pastikan Anda terhubung ke internet.")

df_national_holidays = pd.DataFrame(national_holidays_data)
if not df_national_holidays.empty:
    df_national_holidays["Tanggal"] = pd.to_datetime(
        df_national_holidays["Tanggal"]
    ).dt.date

# --- LANGKAH 2: Tambahkan Musim Kustom (Manual) ---
print("Menambahkan data 'Musim Kustom'...")
custom_seasons_data = []
for start, end, name, type in CUSTOM_SEASONS:
    custom_seasons_data.extend(expand_date_range(start, end, name, type))

df_custom_seasons = pd.DataFrame(custom_seasons_data)
if not df_custom_seasons.empty:
    df_custom_seasons["Tanggal"] = pd.to_datetime(df_custom_seasons["Tanggal"]).dt.date
    print(f"Ditambahkan {len(df_custom_seasons)} baris data dari musim kustom.")
else:
    print("Tidak ada musim kustom yang ditambahkan.")

# --- LANGKAH 3: Gabungkan & Prioritaskan ---
print("Menggabungkan data...")
df_final = pd.concat([df_national_holidays, df_custom_seasons], ignore_index=True)

if df_final.empty:
    print("Tidak ada data untuk diproses.")
else:
    # Prioritaskan: 'Libur Nasional' menang di atas 'Musim Liburan'
    df_final["Prioritas"] = df_final["Tipe_Event"].apply(
        lambda x: 1 if x == "Libur Nasional" else (2 if x == "Event Lokal" else 3)
    )
    df_final = df_final.sort_values(by=["Tanggal", "Prioritas"], ascending=[True, True])

    # Hapus duplikat, ambil prioritas tertinggi
    df_final = df_final.drop_duplicates(subset=["Tanggal"], keep="first")
    df_final = df_final.drop(columns=["Prioritas"])
    df_final = df_final.sort_values(by="Tanggal")

    # --- LANGKAH 4: Simpan ke CSV ---
    try:
        df_final.to_csv(OUTPUT_FILENAME, index=False, encoding="utf-8-sig")
        print("\n" + "=" * 50)
        print(f"✅ SUKSES! Data Kalender 2025 Siap.")
        print(f"File '{OUTPUT_FILENAME}' telah dibuat.")
        print("=" * 50)
        print(df_final.head())  # Tampilkan awal data
    except Exception as e:
        print(f"\n❌ GAGAL menyimpan file: {e}")
