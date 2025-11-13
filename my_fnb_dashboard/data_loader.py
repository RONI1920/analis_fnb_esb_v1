# data_loader.py
import os
import sqlite3
import pandas as pd
import streamlit as st
import openpyxl  # Diperlukan agar pandas bisa membaca file .xlsx

DB_FILE = "fnb_analyst_data.db"


def get_db_connection():
    """Membuat koneksi ke database SQLite."""
    conn = sqlite3.connect(DB_FILE)
    return conn


def init_db():
    """
    Membuat skema tabel database yang BENAR jika belum ada.
    """
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        # 1. Skema Tabel GMV
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS gmv_data (
            "Sales Date In" DATETIME, "Sales Date Out" DATETIME, "Bill Number" TEXT,
            "Menu" TEXT, "Qty" REAL, "Price (Net)" REAL, "Service Charge" REAL,
            "Tax" REAL, "Total Nett Sales" REAL, "Bill Discount" REAL,
            "Total Gross Sales" REAL, "Total After Bill Discount" REAL,
            "Difference Price" REAL, "Discount" REAL, "Payment Method" TEXT,
            "Visit Purpose" TEXT, "Menu Category" TEXT, "Menu Category Detail" TEXT,
            "Waiter" TEXT, "Order Time" DATETIME, "Company" TEXT, "Period" TEXT, "Branch" TEXT
        );
        """
        )
        # 2. Skema Tabel COGS
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS cogs_data (
            "Sales Date" DATETIME, "Branch" TEXT, "Menu Category" TEXT, "Menu" TEXT,
            "Harga Jual" REAL, "COGS" REAL, "Qty" REAL, "Total" REAL
        );
        """
        )
        # 3. Skema Tabel Waiter
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS waiter_data (
            "Bill Number" TEXT, "Waiter" TEXT, "Order Time" DATETIME,
            "Total After Bill Discount" REAL, "Branch" TEXT
        );
        """
        )
        # 4. Skema Tabel Ulasan
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS ulasan_data (
            "Nama" TEXT, "Rating" TEXT, "Ulasan" TEXT, "Rating_Clean" INTEGER
        );
        """
        )
        # 5. Skema Tabel Purchase
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS purchase_data (
            "Purchase Date" DATETIME, "Required Date" DATETIME, "Purchase Number" TEXT,
            "Supplier Name" TEXT, "Category" TEXT, "Sub Category" TEXT, "Product Name" TEXT,
            "PO Qty" REAL, "Receipt Qty" REAL, "Pricelist Price" REAL, "Price" REAL,
            "Discount" REAL, "VAT" REAL, "Total" REAL, "Branch" TEXT
        );
        """
        )
        conn.commit()
    except Exception as e:
        st.error(f"Gagal total saat inisialisasi database: {e}")
    finally:
        if conn:
            conn.close()


def save_dataframe_to_db(df, table_name):
    """Menyimpan DataFrame ke tabel, mengganti yang lama."""
    if df is None or df.empty:
        st.error(f"Dataframe untuk {table_name} kosong, tidak ada yang disimpan.")
        return
    try:
        conn = get_db_connection()
        df.to_sql(table_name, conn, if_exists="replace", index=False)
        conn.commit()
        conn.close()
        print(f"Data {table_name} (mode replace) berhasil disimpan.")
    except Exception as e:
        st.error(f"Gagal menyimpan data {table_name} ke DB: {e}")


def save_dataframe_smart_append(df, table_name, date_col_name):
    """
    Menyimpan DataFrame ke DB dengan strategi "Smart Append".
    """
    if df is None or df.empty:
        st.error(f"Dataframe untuk {table_name} kosong, tidak ada yang disimpan.")
        return
    if date_col_name not in df.columns:
        st.error(f"Kolom tanggal '{date_col_name}' tidak ditemukan. Gagal menyimpan.")
        return
    try:
        df[date_col_name] = pd.to_datetime(df[date_col_name])
        min_date = df[date_col_name].min().strftime("%Y-%m-%d %H:%M:%S")
        max_date = df[date_col_name].max().strftime("%Y-%m-%d %H:%M:%S")
    except Exception as e:
        st.error(f"Gagal memproses kolom tanggal '{date_col_name}': {e}")
        return

    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        delete_query = (
            f'DELETE FROM {table_name} WHERE "{date_col_name}" BETWEEN ? AND ?'
        )
        cursor.execute(delete_query, (min_date, max_date))
        deleted_rows = cursor.rowcount
        if deleted_rows > 0:
            st.warning(
                f"Menghapus {deleted_rows} data lama dari rentang {min_date} s/d {max_date}..."
            )
        df.to_sql(table_name, conn, if_exists="append", index=False)
        conn.commit()
        st.success(
            f"Sukses! {len(df)} baris data baru untuk {table_name} berhasil disimpan."
        )
    except Exception as e:
        st.error(f"Gagal menyimpan data {table_name} ke DB: {e}")
        if conn:
            conn.rollback()
    finally:
        if conn:
            conn.close()


@st.cache_data(show_spinner=False)
def load_dataframe_from_db(table_name, date_cols=[], numeric_cols_config={}):
    """Memuat DataFrame dari tabel dan memperbaiki tipe data."""
    if not os.path.exists(DB_FILE):
        return None
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute(
            f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';"
        )
        if cursor.fetchone() is None:
            conn.close()
            return None
        df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        conn.close()
        if df.empty:
            return None
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")
        for col_name, col_type in numeric_cols_config.items():
            if col_name in df.columns:
                if col_type == "int":
                    df[col_name] = (
                        pd.to_numeric(df[col_name], errors="coerce")
                        .fillna(0)
                        .astype(int)
                    )
                else:
                    df[col_name] = pd.to_numeric(df[col_name], errors="coerce").fillna(
                        0
                    )
        return df
    except Exception as e:
        if conn:
            conn.close()
        return None


@st.cache_data
def load_data_gmv(uploaded_file, use_db=False):
    """Memuat dan membersihkan data GMV (File 1) DAN membaca headernya."""
    if use_db and uploaded_file is None:
        with st.spinner("Memuat data GMV dari database..."):
            numeric_config = {
                "Qty": "float",
                "Price (Net)": "float",
                "Service Charge": "float",
                "Tax": "float",
                "Total Nett Sales": "float",
                "Bill Discount": "float",
                "Total Gross Sales": "float",
                "Total After Bill Discount": "float",
                "Difference Price": "float",
                "Discount": "float",
            }
            df = load_dataframe_from_db(
                "gmv_data",
                date_cols=["Sales Date In", "Sales Date Out", "Order Time"],
                numeric_cols_config=numeric_config,
            )
            if df is None:
                st.info("Database GMV kosong. Silakan upload file baru.", icon="ℹ️")
                return None, None, None, None
            return df, "DB_MODE", "DB_MODE", "DB_MODE"
    if uploaded_file is None:
        return None, None, None, None

    df_data = None
    company_name = "N/A"
    period_str = "N/A"
    branch_name_header = "N/A"
    try:
        if uploaded_file.name.endswith(".xlsx"):
            df_header = pd.read_excel(
                uploaded_file, header=None, nrows=7, usecols=[0, 1], engine="openpyxl"
            )
            company_name = str(df_header.iloc[1, 0])
            period_str = str(df_header.iloc[4, 1])
            branch_name_header = str(df_header.iloc[5, 1])
            uploaded_file.seek(0)
            df_data = pd.read_excel(uploaded_file, header=9, engine="openpyxl")
        elif uploaded_file.name.endswith(".csv"):
            uploaded_file.seek(0)
            df_header = pd.read_csv(
                uploaded_file, header=None, nrows=7, encoding="latin1"
            )
            company_name = str(df_header.iloc[1, 0])
            period_str = str(df_header.iloc[4, 1])
            branch_name_header = str(df_header.iloc[5, 1])
            uploaded_file.seek(0)
            df_data = pd.read_csv(uploaded_file, header=9, encoding="latin1")
        else:
            st.error(f"Format file {uploaded_file.name} tidak didukung.")
            return None, None, None, None
    except Exception as e:
        st.error(f"Error membaca file GMV: {e}")
        st.error("Pastikan header file GMV ada di baris 10.")
        return None, None, None, None

    if df_data is not None:
        df_data.columns = [col.strip().title() for col in df_data.columns]
    else:
        st.error("Gagal memuat df_data.")
        return None, None, None, None

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
        if col in df_data.columns:
            df_data[col] = pd.to_numeric(df_data[col], errors="coerce").fillna(0)
        else:
            st.warning(f"Peringatan (File 1): Kolom '{col}' tidak ditemukan.")

    if "Sales Date In" in df_data.columns:
        df_data["Sales Date In"] = pd.to_datetime(
            df_data["Sales Date In"], errors="coerce"
        )
    if "Sales Date Out" in df_data.columns:
        df_data["Sales Date Out"] = pd.to_datetime(
            df_data["Sales Date Out"], errors="coerce"
        )
    if "Order Time" in df_data.columns:
        df_data["Order Time"] = pd.to_datetime(df_data["Order Time"], errors="coerce")

    df_data["Company"] = company_name
    df_data["Period"] = period_str
    df_data["Branch"] = branch_name_header

    if "Bill Number" not in df_data.columns or "Sales Date In" not in df_data.columns:
        st.warning(
            "Peringatan: Kolom 'Bill Number' atau 'Sales Date In' tidak ditemukan."
        )
    else:
        df_data.dropna(subset=["Bill Number", "Sales Date In"], inplace=True)

    return df_data, company_name, period_str, branch_name_header


@st.cache_data
def load_data_purchase(uploaded_file, use_db=False):
    """Memuat dan membersihkan data Laporan Pembelian (File 5)."""
    if use_db and uploaded_file is None:
        with st.spinner("Memuat data Pembelian dari database..."):
            numeric_config = {
                "PO Qty": "float",
                "Receipt Qty": "float",
                "Pricelist Price": "float",
                "Price": "float",
                "Discount": "float",
                "VAT": "float",
                "Total": "float",
            }
            df = load_dataframe_from_db(
                "purchase_data",
                date_cols=["Purchase Date", "Required Date"],
                numeric_cols_config=numeric_config,
            )
            if df is None:
                st.info(
                    "Database Pembelian kosong. Silakan upload file baru.", icon="ℹ️"
                )
                return None
            return df
    if uploaded_file is None:
        return None

    df = None
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, header=11, encoding="latin1")
        elif uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, header=11, engine="openpyxl")
        else:
            st.error("Format file Pembelian (File 5) tidak didukung.")
            return None
    except Exception as e:
        st.error(f"Error membaca file Pembelian (File 5): {e}")
        st.error("Pastikan header file Pembelian ada di baris 12.")
        return None

    numeric_cols = [
        "PO Qty",
        "Receipt Qty",
        "Pricelist Price",
        "Price",
        "Discount",
        "VAT",
        "Total",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            st.warning(f"Peringatan (File 5): Kolom wajib '{col}' tidak ditemukan.")
            if col == "Total":
                return None

    date_cols = ["Purchase Date", "Required Date"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    if (
        "Category" not in df.columns
        or "Product Name" not in df.columns
        or "Supplier Name" not in df.columns
    ):
        st.error(
            "File 5 kekurangan kolom 'Category', 'Product Name', atau 'Supplier Name'."
        )
        return None
    if "Branch" not in df.columns:
        st.warning("Peringatan: Kolom 'Branch' tidak ditemukan di File Pembelian.")

    df.dropna(subset=["Purchase Number", "Purchase Date"], inplace=True)
    if "Branch" in df.columns:
        df["Branch"] = df["Branch"].astype(str).str.strip().str.title()
    return df


@st.cache_data
def load_cogs_data(uploaded_file, use_db=False):
    """Memuat dan membersihkan data COGS (File 2)."""
    if use_db and uploaded_file is None:
        with st.spinner("Memuat data COGS dari database..."):
            numeric_config = {
                "Harga Jual": "float",
                "COGS": "float",
                "Qty": "float",
                "Total": "float",
            }
            df = load_dataframe_from_db(
                "cogs_data",
                date_cols=["Sales Date"],
                numeric_cols_config=numeric_config,
            )
            if df is None:
                st.info("Database COGS kosong. Silakan upload file baru.", icon="ℹ️")
                return None
            return df
    if uploaded_file is None:
        return None

    df = None
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, header=12, encoding="latin1")
        elif uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, header=12, engine="openpyxl")
        else:
            st.error("Format file COGS (File 2) tidak didukung.")
            return None
    except Exception as e:
        st.error(f"Error membaca file COGS (File 2): {e}")
        st.error("Pastikan header file COGS ada di baris 13.")
        return None

    column_mapping = {"Price": "Harga Jual", "COGS Total": "COGS"}
    df.rename(columns=column_mapping, inplace=True)

    required_cols = ["Menu", "Harga Jual", "COGS", "Qty", "Total", "Sales Date"]
    missing_cols = [col for col in required_cols if col not in df.columns]

    if "Branch" not in df.columns:
        st.warning("Peringatan: Kolom 'Branch' tidak ditemukan di File COGS.")

    if missing_cols:
        st.error(f"File COGS (File 2) kekurangan kolom: {missing_cols}")
        st.error(f"Kolom yang ditemukan: {list(df.columns)}")
        return None

    df["Menu"] = df["Menu"].astype(str)
    df["Harga Jual"] = pd.to_numeric(df["Harga Jual"], errors="coerce").fillna(0)
    df["COGS"] = pd.to_numeric(df["COGS"], errors="coerce").fillna(0)
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
    df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0)
    df["Sales Date"] = pd.to_datetime(df["Sales Date"], errors="coerce")

    if "Branch" in df.columns:
        df["Branch"] = df["Branch"].astype(str).str.strip().str.title()

    return df


@st.cache_data
def load_data_waiter(uploaded_file, use_db=False):
    """Memuat dan membersihkan data Waiter (File 3)."""
    if use_db and uploaded_file is None:
        with st.spinner("Memuat data Waiter dari database..."):
            numeric_config = {"Total After Bill Discount": "float"}
            df = load_dataframe_from_db(
                "waiter_data",
                date_cols=["Order Time"],
                numeric_cols_config=numeric_config,
            )
            if df is None:
                st.info("Database Waiter kosong. Silakan upload file baru.", icon="ℹ️")
                return None
            return df
    if uploaded_file is None:
        return None

    df = None
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, header=11, encoding="latin1")
        elif uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, header=11, engine="openpyxl")
        else:
            st.error("Format file Waiter (File 3) tidak didukung.")
            return None
    except Exception as e:
        st.error(f"Error membaca file Waiter (File 3): {e}")
        st.error("Pastikan header file Waiter ada di baris 12.")
        return None

    required_cols = ["Bill Number", "Waiter", "Order Time", "Total After Bill Discount"]
    if not all(col in df.columns for col in required_cols):
        st.error(
            f"File Waiter (File 3) harus memiliki kolom: {', '.join(required_cols)}"
        )
        st.error(f"Kolom ditemukan: {list(df.columns)}")
        return None

    if "Branch" not in df.columns:
        st.warning("Peringatan: Kolom 'Branch' tidak ditemukan di File Waiter.")

    df["Order Time"] = pd.to_datetime(df["Order Time"], errors="coerce")
    df["Total After Bill Discount"] = pd.to_numeric(
        df["Total After Bill Discount"], errors="coerce"
    ).fillna(0)
    df.dropna(subset=["Bill Number", "Order Time"], inplace=True)

    if "Branch" in df.columns:
        df["Branch"] = df["Branch"].astype(str).str.strip().str.title()

    return df


@st.cache_data
def load_data_ulasan(uploaded_file, use_db=False):
    """Memuat dan membersihkan data Ulasan (File 4)."""
    if use_db and uploaded_file is None:
        with st.spinner("Memuat data Ulasan dari database..."):
            numeric_config = {"Rating_Clean": "int"}
            df = load_dataframe_from_db(
                "ulasan_data", numeric_cols_config=numeric_config
            )
            if df is None:
                st.info("Database Ulasan kosong. Silakan upload file baru.", icon="ℹ️")
                return None
            return df
    if uploaded_file is None:
        return None

    df = None
    try:
        if uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        elif uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, encoding="latin1")
        else:
            st.error("Format file Ulasan (File 4) tidak didukung.")
            return None
    except Exception as e:
        st.error(f"Error membaca file Ulasan (File 4): {e}")
        st.error(
            "Pastikan file adalah .csv atau .xlsx dengan kolom: Nama, Rating, Ulasan"
        )
        return None

    if "Rating" not in df.columns or "Ulasan" not in df.columns:
        st.error("File Ulasan (File 4) harus memiliki kolom 'Rating' dan 'Ulasan'.")
        return None

    df["Rating_Clean"] = (
        df["Rating"].astype(str).str.extract(r"(\d+)").fillna(0).astype(int)
    )
    df.dropna(subset=["Ulasan"], inplace=True)
    df = df[df["Rating_Clean"] > 0]
    df["Ulasan"] = df["Ulasan"].astype(str)

    return df
