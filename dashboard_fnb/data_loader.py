### data_loader.py ###
import pandas as pd
import sqlite3
import openpyxl  # Diperlukan oleh pandas
import glob
import time
import re

# --- FUNGSI PEMUATAN FILE (Sama seperti kode Anda, tanpa cache) ---
# Catatan: Kita tidak butuh @st.cache_data di sini


def load_data_gmv(uploaded_file):
    df = None
    try:
        if uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, header=9, engine="openpyxl")
        elif uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, header=9, encoding="latin1")
        else:
            print("Format file 1 tidak didukung.")
            return None
    except Exception as e:
        print(f"Error membaca file GMV (File 1): {e}")
        return None

    numeric_cols = [
        "Qty",
        "Price (Net)",
        "Service Charge",
        "Tax",
        "Total Nett Sales",
        "Bill Discount",
        "Total Gross Sales",
        "Total After Bill Discount",
        "Difference Price",
        "Discount",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    if "Sales Date In" in df.columns:
        df["Sales Date In"] = pd.to_datetime(df["Sales Date In"], errors="coerce")
    if "Sales Date Out" in df.columns:
        df["Sales Date Out"] = pd.to_datetime(df["Sales Date Out"], errors="coerce")

    # Hapus kolom yang tidak bisa disimpan di SQL (jika ada)
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
    df.dropna(subset=["Bill Number", "Sales Date In"], inplace=True)
    return df


def load_cogs_data(uploaded_file):
    df = None
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, header=12, encoding="latin1")
        elif uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, header=12, engine="openpyxl")
        else:
            print("Format file COGS (File 2) tidak didukung.")
            return None
    except Exception as e:
        print(f"Error membaca file COGS (File 2): {e}")
        return None

    column_mapping = {"Price": "Harga Jual", "COGS Total": "COGS"}
    df.rename(columns=column_mapping, inplace=True)

    required_cols = ["Menu", "Harga Jual", "COGS", "Qty", "Total", "Sales Date"]
    for col in required_cols:
        if col not in df.columns:
            print(f"File COGS (File 2) kekurangan kolom: {col}")
            return None

    df["Menu"] = df["Menu"].astype(str)
    df["Harga Jual"] = pd.to_numeric(df["Harga Jual"], errors="coerce").fillna(0)
    df["COGS"] = pd.to_numeric(df["COGS"], errors="coerce").fillna(0)
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
    df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0)
    df["Sales Date"] = pd.to_datetime(df["Sales Date"], errors="coerce")
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
    return df


def load_data_waiter(uploaded_file):
    df = None
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, header=11, encoding="latin1")
        elif uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, header=11, engine="openpyxl")
        else:
            print("Format file Waiter (File 3) tidak didukung.")
            return None
    except Exception as e:
        print(f"Error membaca file Waiter (File 3): {e}")
        return None

    required_cols = ["Bill Number", "Waiter", "Order Time", "Total After Bill Discount"]
    if not all(col in df.columns for col in required_cols):
        print(f"File Waiter (File 3) harus memiliki kolom: {', '.join(required_cols)}")
        return None

    df["Order Time"] = pd.to_datetime(df["Order Time"], errors="coerce")
    df["Total After Bill Discount"] = pd.to_numeric(
        df["Total After Bill Discount"], errors="coerce"
    ).fillna(0)
    df.dropna(subset=["Bill Number", "Order Time"], inplace=True)
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
    return df


def load_data_ulasan(uploaded_file):
    df = None
    try:
        if uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        elif uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, encoding="latin1")
        else:
            print("Format file Ulasan (File 4) tidak didukung.")
            return None
    except Exception as e:
        print(f"Error membaca file Ulasan (File 4): {e}")
        return None

    if "Rating" not in df.columns or "Ulasan" not in df.columns:
        print("File Ulasan (File 4) harus memiliki kolom 'Rating' dan 'Ulasan'.")
        return None

    df["Rating_Clean"] = (
        df["Rating"].astype(str).str.extract(r"(\d+)").fillna(0).astype(int)
    )
    df.dropna(subset=["Ulasan"], inplace=True)
    df = df[df["Rating_Clean"] > 0]
    df["Ulasan"] = df["Ulasan"].astype(str)
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
    return df


# --- FUNGSI UTAMA UNTUK MEMBUAT DATABASE ---
def create_database():
    DB_NAME = "fnb_data.db"

    # 1. Temukan file (Asumsi nama file mengandung kata kunci unik)
    #    Sesuaikan pola 'glob' ini jika nama file Anda berbeda
    try:
        gmv_file_path = glob.glob("*[Gg][Mm][Vv]*.xls*")[0]
        cogs_file_path = glob.glob("*[Cc][Oo][Gg][Ss]*.xls*")[0]
        waiter_file_path = glob.glob("*[Ww][Aa][Ii][Tt][Ee][Rr]*.xls*")[
            0
        ]  # Asumsi mengandung kata "Waiter" atau "Rekapitulasi"
        ulasan_file_path = glob.glob("*[Uu][Ll][Aa][Ss][Aa][Nn]*.xls*")[
            0
        ]  # Asumsi mengandung kata "Ulasan"
    except IndexError as e:
        print(f"Error: Tidak dapat menemukan file. Pastikan file ada di folder ini.")
        print(f"Detail: {e}")
        print(
            "\nPastikan nama file Anda mengandung 'gmv', 'cogs', 'waiter', dan 'ulasan'."
        )
        return
    except Exception as e:
        print(f"Error: {e}")
        return

    print("File ditemukan, mulai memuat...")

    # 2. Muat data ke DataFrames
    #    Kita gunakan 'open(file, 'rb')' untuk simulasi 'uploaded_file' object
    with open(gmv_file_path, "rb") as f:
        f.name = gmv_file_path
        df_gmv = load_data_gmv(f)

    with open(cogs_file_path, "rb") as f:
        f.name = cogs_file_path
        df_cogs = load_cogs_data(f)

    with open(waiter_file_path, "rb") as f:
        f.name = waiter_file_path
        df_waiter = load_data_waiter(f)

    with open(ulasan_file_path, "rb") as f:
        f.name = ulasan_file_path
        df_ulasan = load_data_ulasan(f)

    print("Data berhasil dimuat ke memori.")

    # 3. Buat koneksi database & simpan
    conn = sqlite3.connect(DB_NAME)
    print(f"Koneksi ke {DB_NAME} berhasil.")

    try:
        # Gunakan 'if_exists='replace'' untuk menimpa data lama setiap kali skrip dijalankan
        df_gmv.to_sql("gmv", conn, if_exists="replace", index=False)
        print("Tabel 'gmv' berhasil disimpan.")

        df_cogs.to_sql("cogs", conn, if_exists="replace", index=False)
        print("Tabel 'cogs' berhasil disimpan.")

        df_waiter.to_sql("waiter", conn, if_exists="replace", index=False)
        print("Tabel 'waiter' berhasil disimpan.")

        df_ulasan.to_sql("ulasan", conn, if_exists="replace", index=False)
        print("Tabel 'ulasan' berhasil disimpan.")

    except Exception as e:
        print(f"Gagal menyimpan ke database: {e}")
    finally:
        conn.close()
        print(f"Koneksi {DB_NAME} ditutup.")
        print("\n--- DATABASE BERHASIL DIBUAT ---")


if __name__ == "__main__":
    create_database()
