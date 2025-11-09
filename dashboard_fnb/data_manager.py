import pandas as pd
import streamlit as st
import sqlite3
import os

# Nama file database
DB_FILE = "fnb_dashboard.db"

# --- 1. Fungsi Koneksi Database ---


def get_db_connection():
    """Membuat koneksi ke database SQLite."""
    try:
        conn = sqlite3.connect(DB_FILE)
        return conn
    except Exception as e:
        st.error(f"Gagal terhubung ke database {DB_FILE}: {e}")
        return None


def init_db():
    """Membuat tabel jika belum ada (df.to_sql akan menangani ini)."""
    conn = get_db_connection()
    if conn:
        try:
            # Tidak perlu membuat skema secara eksplisit
            pass
        except Exception as e:
            st.warning(f"Gagal menginisialisasi tabel: {e}")
        finally:
            conn.close()


def save_dataframe_to_db(df, table_name):
    """Menyimpan DataFrame ke tabel di database, menimpa data lama."""
    conn = get_db_connection()
    if conn and df is not None:
        try:
            # Salin df untuk menghindari modifikasi data di session state
            df_to_save = df.copy()

            # Ubah kolom datetime menjadi string ISO 8601 untuk SQLite
            for col in df_to_save.select_dtypes(include=["datetime64[ns]"]).columns:
                df_to_save[col] = df_to_save[col].astype(str)

            df_to_save.to_sql(table_name, conn, if_exists="replace", index=False)
            st.sidebar.success(f"Data '{table_name}' berhasil disimpan ke DB.")
        except Exception as e:
            st.sidebar.error(f"Gagal menyimpan {table_name} ke DB: {e}")
        finally:
            conn.close()


def load_dataframe_from_db(table_name):
    """Memuat DataFrame dari tabel di database."""
    conn = get_db_connection()
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';"
            )
            if cursor.fetchone():
                df = pd.read_sql(f"SELECT * FROM {table_name}", conn)

                # Konversi kolom 'Tanggal' (dan kolom tanggal lainnya) kembali ke datetime
                for col in df.columns:
                    if "tanggal" in col.lower():
                        try:
                            df[col] = pd.to_datetime(df[col])
                        except Exception:
                            pass  # Biarkan jika gagal konversi

                # Konversi khusus untuk COGS (disimpan sebagai master list)
                if table_name == "cogs_data" and "COGS" in df.columns:
                    df["COGS"] = pd.to_numeric(df["COGS"], errors="coerce")

                return df
            else:
                return None
        except Exception as e:
            st.sidebar.warning(f"Data '{table_name}' tidak ditemukan di DB: {e}")
            return None
        finally:
            conn.close()
    return None


# --- 2. Fungsi Pemuatan Data (Data Loaders) ---


def load_data_gmv(file, from_db=False):
    """Memuat data GMV dari file (header 10) atau DB."""
    if from_db:
        return load_dataframe_from_db("gmv_data")

    if file:
        try:
            df = pd.read_excel(file, header=9)  # Header di baris 10 (indeks 9)

            original_columns = list(df.columns)
            df.columns = df.columns.str.strip().str.lower()

            # Opsi nama kolom (dari daftar yang Anda berikan)
            key_date = "sales date in"
            key_bill = "bill number"
            key_payment = "payment method"
            key_visit = "visit purpose"
            key_menu = "menu"
            key_qty = "qty"
            key_discount = "discount"
            key_total = "total"  # 'Total' (setelah diskon per item, sebelum pajak)

            required_cols = [key_date, key_bill, key_menu, key_qty, key_total]
            missing_cols = [col for col in required_cols if col not in df.columns]

            if missing_cols:
                st.error(
                    f"Gagal memuat file GMV: Kolom berikut tidak ditemukan: {', '.join(missing_cols)}"
                )
                st.error(f"Kolom terdeteksi (dari header baris 10): {original_columns}")
                return None

            df = df[df[key_menu] != "Grand Total"].copy()
            df[key_date] = pd.to_datetime(df[key_date], errors="coerce")
            df.dropna(subset=[key_date, key_bill], inplace=True)

            # Standarisasi nama kolom untuk 'analysis.py'
            rename_map = {
                key_date: "Tanggal",
                key_bill: "Nomor Transaksi",
                key_payment: "Metode Pembayaran",
                key_visit: "Tujuan Kunjungan",
                key_menu: "Nama Menu",
                key_qty: "Qty",
                key_discount: "Diskon",
                key_total: "Total",
            }

            # Hanya rename kolom yang ada
            df.rename(
                columns={k: v for k, v in rename_map.items() if k in df.columns},
                inplace=True,
            )

            return df
        except Exception as e:
            st.error(f"Gagal memuat file GMV: {e}")
            return None
    return None


def load_cogs_data(file, from_db=False):
    """Memuat data COGS dari file (header 13) atau DB."""
    if from_db:
        return load_dataframe_from_db("cogs_data")

    if file:
        try:
            df = pd.read_excel(file, header=12)  # Header di baris 13 (indeks 12)

            original_columns = list(df.columns)
            df.columns = df.columns.str.strip().str.lower()

            # Opsi nama kolom (dari daftar yang Anda berikan)
            key_menu = "menu"
            key_qty = "qty"
            key_cogs_total = "cogs total"
            key_date = "sales date"  # Kita perlu tanggal untuk konversi DB

            required_cols = [key_menu, key_qty, key_cogs_total, key_date]
            missing_cols = [col for col in required_cols if col not in df.columns]

            if missing_cols:
                st.error(
                    f"Gagal memuat file COGS: Kolom berikut tidak ditemukan: {', '.join(missing_cols)}"
                )
                st.error(f"Kolom terdeteksi (dari header baris 13): {original_columns}")
                return None

            df = df[df[key_menu] != "Grand Total"].copy()

            # Konversi tipe data
            df[key_qty] = pd.to_numeric(df[key_qty], errors="coerce")
            df[key_cogs_total] = pd.to_numeric(df[key_cogs_total], errors="coerce")
            df[key_date] = pd.to_datetime(df[key_date], errors="coerce")

            # Hapus baris di mana Qty atau COGS tidak valid (misal 0 atau NaN)
            df.dropna(subset=[key_qty, key_cogs_total, key_date], inplace=True)
            df = df[df[key_qty] != 0]

            # Hitung COGS per unit
            df["cogs_per_unit"] = df[key_cogs_total] / df[key_qty]

            # Buat 'master list' COGS: Ambil rata-rata COGS per unit untuk setiap menu
            df_agg = (
                df.groupby(key_menu).agg(COGS=("cogs_per_unit", "mean")).reset_index()
            )

            # Standarisasi nama kolom
            df_agg.rename(columns={key_menu: "Nama Menu"}, inplace=True)

            # Hapus baris dengan COGS 0 atau NaN setelah agregasi
            df_agg.dropna(subset=["COGS"], inplace=True)
            df_agg = df_agg[df_agg["COGS"] > 0]

            return df_agg

        except Exception as e:
            st.error(f"Gagal memuat file COGS: {e}")
            return None
    return None


def load_data_waiter(file, from_db=False):
    """Memuat data Waiter (Sales Recapitulation Detail) dari file (header 12) atau DB."""
    if from_db:
        return load_dataframe_from_db("waiter_data")

    if file:
        try:
            df = pd.read_excel(file, header=11)  # Header di baris 12 (indeks 11)

            original_columns = list(df.columns)
            df.columns = df.columns.str.strip().str.lower()

            # Cek kolom berdasarkan list baru Anda
            key_waiter = "waiter"
            key_date = "sales date"
            key_total = "total"

            required_cols_lower = [key_waiter, key_date, key_total]
            missing_cols = [col for col in required_cols_lower if col not in df.columns]

            if missing_cols:
                st.error(
                    f"Gagal memuat file Waiter: Kolom berikut tidak ditemukan: {', '.join(missing_cols)}"
                )
                st.error(
                    f"Kolom yang terdeteksi (dari header baris 12): {original_columns}"
                )
                st.warning("Pastikan file Anda adalah 'Sales Recapitulation Detail'.")
                return None

            df = df[df[key_waiter] != "Grand Total"].copy()
            df[key_date] = pd.to_datetime(df[key_date], errors="coerce")
            df[key_total] = pd.to_numeric(df[key_total], errors="coerce")

            df.dropna(subset=[key_date, key_waiter, key_total], inplace=True)

            # Standarisasi
            rename_map = {
                key_waiter: "Nama Waiter",
                key_date: "Tanggal",
                key_total: "Total",
            }
            df.rename(
                columns={k: v for k, v in rename_map.items() if k in df.columns},
                inplace=True,
            )

            # OPTIMASI PENTING: Agregasi data agar lebih ringan
            final_cols = ["Tanggal", "Nama Waiter", "Total"]

            df_agg = (
                df[final_cols]
                .groupby(["Tanggal", "Nama Waiter"])
                .agg(Total=("Total", "sum"))
                .reset_index()
            )

            return df_agg

        except Exception as e:
            st.error(f"Gagal memuat file Waiter: {e}")
            return None
    return None


def load_data_ulasan(file, from_db=False):
    """Memuat data Ulasan Pelanggan dari file atau DB."""
    if from_db:
        return load_dataframe_from_db("ulasan_data")

    if file:
        try:
            if file.name.endswith(".csv"):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)

            original_columns = list(df.columns)
            df.columns = df.columns.str.strip().str.lower()

            required_cols_lower = ["ulasan", "rating"]
            missing_cols = [col for col in required_cols_lower if col not in df.columns]

            if missing_cols:
                st.error(
                    f"Gagal memuat file Ulasan: Kolom berikut tidak ditemukan: {', '.join(missing_cols)}"
                )
                st.error(f"Kolom yang terdeteksi: {original_columns}")
                return None

            # PERBAIKAN: Ekstrak angka dari teks 'rating' (misal: "5 bintang")
            df["rating"] = df["rating"].astype(str).str.extract(r"(\d+)", expand=False)
            df["rating"] = pd.to_numeric(df["rating"], errors="coerce")

            df.dropna(subset=["ulasan", "rating"], inplace=True)

            # Standarisasi
            rename_map = {"ulasan": "Ulasan", "rating": "Rating", "nama": "Nama"}
            df.rename(
                columns={k: v for k, v in rename_map.items() if k in df.columns},
                inplace=True,
            )

            return df
        except Exception as e:
            st.error(f"Gagal memuat file Ulasan: {e}")
            return None
    return None


def load_data_purchase(file, from_db=False):
    """Memuat data Pembelian (Purchase Recapitulation) dari file (header 12) atau DB."""
    if from_db:
        return load_dataframe_from_db("purchase_data")

    if file:
        try:
            df = pd.read_excel(file, header=11)  # Header di baris 12 (indeks 11)

            original_columns = list(df.columns)
            df.columns = df.columns.str.strip().str.lower()

            # Opsi nama kolom (dari file baru Anda)
            key_item = "product name"
            key_po = "purchase number"
            key_date = "purchase date"
            key_qty = "po qty"
            key_supplier = "supplier name"

            required_cols = [key_item, key_po, key_date, key_qty]
            missing_cols = [col for col in required_cols if col not in df.columns]

            if missing_cols:
                st.error(
                    f"Gagal memuat file Pembelian: Kolom berikut tidak ditemukan: {', '.join(missing_cols)}"
                )
                st.error(f"Kolom terdeteksi: {original_columns}")
                return None

            df = df[df[key_item] != "Grand Total"].copy()
            df[key_date] = pd.to_datetime(df[key_date], errors="coerce")
            df.dropna(subset=[key_date, key_po], inplace=True)

            # Peta Standarisasi
            rename_map = {
                key_item: "Nama",
                key_date: "Tanggal",
                key_po: "Nomor",
                key_qty: "Qty",
                key_supplier: "Nama Supplier",
                "total": "Total",
            }

            df.rename(
                columns={k: v for k, v in rename_map.items() if k in df.columns},
                inplace=True,
            )

            return df
        except Exception as e:
            st.error(f"Gagal memuat file Pembelian: {e}")
            if "original_columns" in locals():
                st.error(
                    f"Kolom yang terdeteksi (dari header baris 12): {original_columns}"
                )
            return None
    return None
