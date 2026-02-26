import time
import pandas as pd
import streamlit as st
import openpyxl  # Diperlukan agar pandas bisa membaca file .xlsx
import altair as alt  # Library untuk grafik yang lebih baik
import numpy as np  # Diperlukan untuk kalkulasi margin
from prophet import Prophet
from prophet.plot import plot_plotly, plot_components_plotly
import re  # <-- Diperlukan untuk Tab 7
import os
import sqlite3
import plotly.express as px  # <-- DIGANTI: Ditambahkan untuk grafik baru
import plotly.graph_objects as go  # <-- DIGANTI: Ditambahkan untuk grafik baru
import requests  # Pastikan ini ada di atas file
from streamlit_lottie import st_lottie  # Pastikan ini ada di atas file
import json  # <-- Dipindahkan dari bawah
import plotly.express as px

# Tambahkan ini di bagian import paling atas
from mlxtend.frequent_patterns import apriori, association_rules
from mlxtend.preprocessing import TransactionEncoder

# -----------------------------

# --- Konfigurasi Halaman Streamlit ---
st.set_page_config(
    page_title="Data Driven Analyst Specialyst FnB",
    layout="wide",
    initial_sidebar_state="expanded",
)


# --- FUNGSI UNTUK MEMUAT CSS EKSTERNAL ---
def load_css(file_name):
    """Membaca file CSS dan menerapkannya ke aplikasi."""
    try:
        # Pengecekan file di sini
        if os.path.exists(file_name):
            with open(file_name) as f:
                st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
        else:
            # Jika style.css tidak ada, kita tetap jalankan tanpa error
            pass
    except Exception as e:
        st.error(f"Gagal memuat CSS: {e}")


# #################################################################
# --- BAGIAN 1: FUNGSI HELPER (KALKULASI & FORMATTING) ---
# #################################################################


def format_rupiah(amount):
    """Format angka menjadi string Rupiah DENGAN titik pemisah ribuan."""
    # Pastikan input adalah angka, jika tidak setel ke 0
    if pd.isna(amount):
        amount = 0
    return f"Rp {amount:,.0f}".replace(",", ".")


def format_angka_bulat(number):
    """Format angka kuantitas menjadi bulat tanpa desimal."""
    if pd.isna(number):
        number = 0
    return f"{number:.0f}"


def format_persen(number):
    """Format angka menjadi string persentase."""
    if pd.isna(number):
        number = 0
    return f"{number:,.1f}%"


# #################################################################
# --- BAGIAN 1.5: FUNGSI HELPER DATABASE (GANTI FUNGSI INI) ---
# #################################################################

DB_FILE = "database_bisnis_saya.db"  # <-- DISESUAIKAN dengan kode terbaru Anda


def get_db_connection():
    """Membuat koneksi ke database SQLite."""
    conn = sqlite3.connect(DB_FILE)
    return conn


def init_db():
    """
    Membuat skema tabel database yang BENAR.
    DIPERBARUI: Menambahkan kolom 'Sales Type' agar data Void/Sales tersimpan.
    """
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        # 1. Skema Tabel GMV (DIPERBAIKI - Ditambah Sales Type)
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS gmv_data (
            "Sales Date In" DATETIME,
            "Sales Date Out" DATETIME,
            "Bill Number" TEXT,
            "Menu" TEXT,
            "Menu Code" TEXT,
            "Sales Number" TEXT,
            "Sales Type" TEXT,          -- <--- TAMBAHAN PENTING DI SINI!
            "Qty" REAL,
            "Price (Net)" REAL,
            "Service Charge" REAL,
            "Tax" REAL,
            "Total" REAL,
            "Price (Pricelist)" REAL,
            "Total Nett Sales" REAL,
            "Bill Discount" REAL,
            "Total Gross Sales" REAL,
            "Total After Bill Discount" REAL,
            "Difference Price" REAL,
            "Discount" REAL,
            "Payment Method" TEXT,
            "Visit Purpose" TEXT,
            "Menu Category" TEXT,
            "Menu Category Detail" TEXT,
            "Waiter" TEXT,
            "Order Time" DATETIME,  
            "Company" TEXT,
            "Period" TEXT,
            "Branch" TEXT
        );
        """
        )

        # ... (Sisa tabel cogs_data, waiter_data, dll biarkan tetap sama) ...
        # 2. Skema Tabel COGS
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS cogs_data (
            "Sales Date" DATETIME,
            "Branch" TEXT,
            "Menu Category" TEXT,
            "Menu" TEXT,
            "Harga Jual" REAL,
            "COGS" REAL,
            "Qty" REAL,
            "Total" REAL
        );
        """
        )

        # 3. Skema Tabel Waiter
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS waiter_data (
            "Bill Number" TEXT,
            "Waiter" TEXT,
            "Order Time" DATETIME,
            "Total After Bill Discount" REAL,
            "Branch" TEXT,
            "Sales Type" TEXT           -- <--- OPSI: Tambahkan di sini juga jika perlu untuk Tab 3
        );
        """
        )

        # 4. Skema Tabel Ulasan
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS ulasan_data (
            "Nama" TEXT,
            "Rating" TEXT,
            "Ulasan" TEXT,
            "Rating_Clean" INTEGER
        );
        """
        )

        # 5. Skema Tabel Purchase
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS purchase_data (
            "Purchase Date" DATETIME,
            "Required Date" DATETIME,
            "Purchase Number" TEXT,
            "Supplier Name" TEXT,
            "Category" TEXT,
            "Sub Category" TEXT,
            "Product Name" TEXT,
            "PO Qty" REAL,
            "Receipt Qty" REAL,
            "Pricelist Price" REAL,
            "Price" REAL,
            "Discount" REAL,
            "VAT" REAL,
            "Total" REAL,
            "Branch" TEXT
        );
        """
        )

        # 6. Skema Tabel Profit & Loss (BARU)
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS pl_data (
            "Account" TEXT,
            "Description" TEXT,
            "Date" DATETIME,
            "Month_Name" TEXT,
            "Year_Type" TEXT, -- Current Year / Last Year
            "Value" REAL,
            "Branch" TEXT,
            "Category" TEXT -- Revenue/Expense/COGS (Derived)
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
    """
    Menyimpan DataFrame ke tabel, mengganti yang lama.
    VERSI BARU: Otomatis memfilter kolom agar sesuai dengan skema DB.
    """
    if df is None or df.empty:
        st.error(f"Dataframe untuk {table_name} kosong, tidak ada yang disimpan.")
        return

    conn = None
    try:
        conn = get_db_connection()

        # 1. Dapatkan skema kolom dari tabel database
        db_schema_df = pd.read_sql_query(f"SELECT * FROM {table_name} LIMIT 0", conn)
        db_columns = db_schema_df.columns.tolist()

        # 2. Dapatkan kolom dari DataFrame Excel Anda
        df_columns = df.columns.tolist()

        # 3. Temukan kolom yang cocok
        columns_to_keep = [col for col in df_columns if col in db_columns]

        # 4. Temukan kolom yang akan diabaikan (untuk info)
        extra_columns = [col for col in df_columns if col not in db_columns]
        if extra_columns:
            st.warning(
                f"Kolom berikut diabaikan karena tidak ada di database: {', '.join(extra_columns)}"
            )

        # 5. Buat DataFrame bersih yang hanya berisi kolom yang cocok
        df_clean = df[columns_to_keep]

        # 6. Simpan DataFrame yang bersih
        df_clean.to_sql(table_name, conn, if_exists="replace", index=False)
        conn.commit()

        print(f"Data {table_name} (mode replace) berhasil disimpan.")

    except Exception as e:
        st.error(f"Gagal menyimpan data {table_name} ke DB: {e}")
    finally:
        if conn:
            conn.close()


def save_dataframe_smart_append(df, table_name, date_col_name):
    """
    Menyimpan DataFrame ke DB dengan strategi "Smart Append".
    VERSI BARU: Otomatis memfilter kolom agar sesuai dengan skema DB.
    """
    if df is None or df.empty:
        st.error(f"Dataframe untuk {table_name} kosong, tidak ada yang disimpan.")
        return

    # --- BLOK BARU UNTUK MEMFILTER KOLOM ---
    conn_schema = None
    try:
        conn_schema = get_db_connection()
        # 1. Dapatkan skema kolom dari tabel database
        db_schema_df = pd.read_sql_query(
            f"SELECT * FROM {table_name} LIMIT 0", conn_schema
        )
        db_columns = db_schema_df.columns.tolist()

        # 2. Dapatkan kolom dari DataFrame Excel Anda
        df_columns = df.columns.tolist()

        # 3. Temukan kolom yang cocok
        columns_to_keep = [col for col in df_columns if col in db_columns]

        # 4. Temukan kolom yang akan diabaikan (untuk info)
        extra_columns = [col for col in df_columns if col not in db_columns]
        if extra_columns:
            st.warning(
                f"Kolom berikut diabaikan karena tidak ada di database: {', '.join(extra_columns)}"
            )

        # 5. Buat DataFrame bersih yang hanya berisi kolom yang cocok
        df_clean = df[columns_to_keep]

    except Exception as e:
        st.error(f"Gagal mencocokkan skema database: {e}")
        return
    finally:
        if conn_schema:
            conn_schema.close()
    # --- AKHIR BLOK BARU ---

    # 1. Pastikan kolom tanggal ada di data yang SUDAH BERSIH
    if date_col_name not in df_clean.columns:
        st.error(
            f"Kolom tanggal '{date_col_name}' tidak ditemukan di data. Gagal menyimpan."
        )
        return

    # 2. Konversi ke datetime (jika belum) dan cari rentang tanggal
    try:
        df_clean[date_col_name] = pd.to_datetime(df_clean[date_col_name])
        min_date = df_clean[date_col_name].min().strftime("%Y-%m-%d %H:%M:%S")
        max_date = df_clean[date_col_name].max().strftime("%Y-%m-%d %H:%M:%S")
    except Exception as e:
        st.error(f"Gagal memproses kolom tanggal '{date_col_name}': {e}")
        return

    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        # 3. HAPUS data lama yang tumpang tindih
        delete_query = (
            f'DELETE FROM {table_name} WHERE "{date_col_name}" BETWEEN ? AND ?'
        )
        cursor.execute(delete_query, (min_date, max_date))

        deleted_rows = cursor.rowcount
        if deleted_rows > 0:
            st.warning(
                f"Menghapus {deleted_rows} data lama dari rentang {min_date} s/d {max_date}..."
            )

        # 4. TAMBAHKAN (Append) data baru yang sudah bersih
        df_clean.to_sql(table_name, conn, if_exists="append", index=False)

        conn.commit()
        st.success(
            f"Sukses! {len(df_clean)} baris data baru untuk {table_name} berhasil disimpan ke database."
        )

    except Exception as e:
        st.error(f"Gagal menyimpan data {table_name} ke DB: {e}")
        if conn:
            conn.rollback()
    finally:
        if conn:
            conn.close()


@st.cache_data
def load_dataframe_from_db(table_name, date_cols=[], numeric_cols_config={}):
    """Memuat DataFrame dari tabel dan memperbaiki tipe data."""
    if not os.path.exists(DB_FILE):
        return None  # DB file tidak ada

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute(
            f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';"
        )
        if cursor.fetchone() is None:
            conn.close()
            return None  # Tabel tidak ada

        df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        conn.close()

        if df.empty:
            return None  # Tabel ada tapi kosong

        # --- PERBAIKAN TIPE DATA (PENTING!) ---
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")

        # Perbaikan tipe data numerik
        for col_name, col_type in numeric_cols_config.items():
            if col_name in df.columns:
                if col_type == "int":
                    df[col_name] = (
                        pd.to_numeric(df[col_name], errors="coerce")
                        .fillna(0)
                        .astype(int)
                    )
                else:  # float
                    df[col_name] = pd.to_numeric(df[col_name], errors="coerce").fillna(
                        0
                    )

        return df

    except Exception as e:
        # st.error(f"Gagal memuat {table_name} dari DB: {e}")
        if conn:
            conn.close()
        return None


# Fungsi Analisis & Grafik
def clean_payment_method(method_str):
    """Mengelompokkan metode pembayaran yang berantakan."""
    method_str = str(method_str).upper()
    if "VOUCHER" in method_str or "," in method_str:
        return "Voucher / Split"
    if "QRIS" in method_str:
        return "QRIS"
    if "VISA" in method_str:
        return "VISA"
    if "DEBIT CARD" in method_str:
        return "DEBIT CARD"
    if "BCA CARD" in method_str:
        return "BCA CARD"
    if "MASTER" in method_str:
        return "MASTER"
    if "CC " in method_str or "CREDIT CARD" in method_str:
        return "CREDIT CARD (Lainnya)"
    if "CASH" in method_str:
        return "CASH"
    if "TRANSFER" in method_str:
        return "TRANSFER"
    if "BRI CARD" in method_str:
        return "BRI CARD"
    if "BNI CARD" in method_str:
        return "BNI CARD"
    return "Lainnya"


@st.cache_data
def get_prophet_projection(prophet_data, sisa_hari):
    """
    Melatih model Prophet yang di-cache dan mengembalikan nilai ramalan.
    Fungsi ini hanya akan berjalan jika input (data & sisa_hari) berubah.
    """
    try:
        # Kita tambahkan pengecekan data di sini
        if len(prophet_data) < 7:
            st.warning("Data bulan ini < 7 hari, ramalan Prophet tidak akurat.")
            return None  # Tidak cukup data

        # Tampilkan spinner HANYA saat fungsi ini benar-benar berjalan
        with st.spinner("Menjalankan model ramalan Prophet (pertama kali)..."):
            model_prophet = Prophet(
                weekly_seasonality=True,
                daily_seasonality=False,
                changepoint_prior_scale=0.1,
            )
            model_prophet.fit(prophet_data)

            future_df_prophet = model_prophet.make_future_dataframe(periods=sisa_hari)
            forecast_df_prophet = model_prophet.predict(future_df_prophet)

            # Kembalikan hanya total ramalan untuk sisa hari
            # Pastikan sisa_hari > 0
            if sisa_hari > 0:
                ramalan_sisa_hari = forecast_df_prophet.iloc[-sisa_hari:]["yhat"].sum()
            else:
                ramalan_sisa_hari = 0

            return ramalan_sisa_hari

    except Exception as e:
        st.error(f"Gagal menjalankan ramalan Prophet: {e}", icon="🤖")
        return None


@st.cache_data
def calculate_sales_kpi(df):
    """Menghitung KPI Penjualan Utama."""
    if df is None or df.empty:
        return {
            "Total Pendapatan Kotor": 0,
            "Total Penjualan Bersih (Nett)": 0,
            "Total Transaksi": 0,
            "Rata-rata Nilai Transaksi (ATV)": 0,
            "Total Item Terjual": 0,
            "Item per Transaksi (IPB)": 0,
            "Total Diskon": 0,
            "Total Service Charge": 0,
            "Total Pajak": 0,
        }

    total_revenue = df["Total After Bill Discount"].sum()
    total_nett_sales = df["Total Nett Sales"].sum()
    unique_bills = df["Bill Number"].nunique()
    atv = total_revenue / unique_bills if unique_bills > 0 else 0
    total_items_sold = df["Qty"].sum()
    ipb = total_items_sold / unique_bills if unique_bills > 0 else 0
    total_discounts = (
        df["Difference Price"].sum() + df["Discount"].sum() + df["Bill Discount"].sum()
    )
    total_service_charge = df["Service Charge"].sum()
    total_tax = df["Tax"].sum()

    return {
        "Total Pendapatan Kotor": total_revenue,
        "Total Penjualan Bersih (Nett)": total_nett_sales,
        "Total Transaksi": unique_bills,
        "Rata-rata Nilai Transaksi (ATV)": atv,
        "Total Item Terjual": total_items_sold,
        "Item per Transaksi (IPB)": ipb,
        "Total Diskon": total_discounts,
        "Total Service Charge": total_service_charge,
        "Total Pajak": total_tax,
    }


@st.cache_data
def get_payment_analysis(df):
    """Menganalisis penjualan berdasarkan metode pembayaran."""
    bill_data = (
        df.groupby("Bill Number")
        .agg(
            Bill_Revenue=("Total After Bill Discount", "sum"),
            Payment_Method=("Payment Method", "first"),
        )
        .reset_index()
    )
    bill_data["Cleaned_Payment"] = bill_data["Payment_Method"].apply(
        clean_payment_method
    )
    payment_analysis = (
        bill_data.groupby("Cleaned_Payment")["Bill_Revenue"]
        .agg(Total_Penjualan="sum", Jumlah_Transaksi="count")
        .sort_values(by="Total_Penjualan", ascending=False)
    )
    return payment_analysis.reset_index()


@st.cache_data
def get_visit_purpose_analysis(df):
    """Menganalisis penjualan berdasarkan tujuan kunjungan."""
    bill_data = (
        df.groupby("Bill Number")
        .agg(
            Bill_Revenue=("Total After Bill Discount", "sum"),
            Visit_Purpose=("Visit Purpose", "first"),
        )
        .reset_index()
    )
    sales_by_visit = (
        bill_data.groupby("Visit_Purpose")["Bill_Revenue"]
        .sum()
        .sort_values(ascending=False)
    )
    return sales_by_visit.reset_index().rename(
        columns={
            "Visit_Purpose": "Visit Purpose",
            "Bill_Revenue": "Total After Bill Discount",
        }
    )


@st.cache_data
def get_menu_performance(df, filter_regex_items_str):  # <-- 1. TAMBAHKAN ARGUMEN
    """
    Menganalisis performa menu dan kategori.
    VERSI HYBRID: Mengembalikan data untuk Top/Bottom 10 DAN data mentah
    untuk drill-down interaktif.
    """

    # --- Filter awal (tetap sama) ---
    menu_sales = df[~df["Menu"].str.contains("PACKAGE", na=False, case=False)]
    menu_sales = menu_sales[menu_sales["Price (Net)"] > 0]

    # #################################################################
    # --- PERBAIKAN: Filter dinamis berdasarkan input UI ---

    # 2. HAPUS SEMUA filter_regex yang di-hardcode

    # 3. Gunakan argumen 'filter_regex_items_str' untuk semua filter
    if filter_regex_items_str:  # Hanya jalankan filter jika string tidak kosong

        # Filter Kolom 1: 'Menu Category'
        if "Menu Category" in df.columns:
            menu_sales = menu_sales[
                ~menu_sales["Menu Category"].str.contains(
                    filter_regex_items_str, na=False, case=False, regex=True
                )
            ]

        # Filter Kolom 2: 'Menu Category Detail'
        NAMA_KOLOM_DETAIL = "Menu Category Detail"
        if NAMA_KOLOM_DETAIL in df.columns:
            menu_sales = menu_sales[
                ~menu_sales[NAMA_KOLOM_DETAIL].str.contains(
                    filter_regex_items_str, na=False, case=False, regex=True
                )
            ]

        # Filter Kolom 3: 'Menu'
        if "Menu" in menu_sales.columns:
            menu_sales = menu_sales[
                ~menu_sales["Menu"].str.contains(
                    filter_regex_items_str, na=False, case=False, regex=True
                )
            ]

    # --- BATAS PERBAIKAN ---
    # #################################################################

    # --- Data untuk Drill-Down & Kategori Statis ---
    # (Sisa kode di bawah ini tidak perlu diubah)

    top_selling_categories = pd.DataFrame(columns=["Menu Category", "Qty"])
    top_grossing_categories = pd.DataFrame(
        columns=["Menu Category", "Total Nett Sales"]
    )
    menu_sales_cat_df = pd.DataFrame()

    if "Menu Category" in df.columns:
        menu_sales_cat_df = menu_sales.copy()  # menu_sales sudah bersih

        top_selling_categories = (
            menu_sales_cat_df.groupby("Menu Category")["Qty"]
            .sum()
            .nlargest(10)
            .sort_values(ascending=False)
        )
        top_grossing_categories = (
            menu_sales_cat_df.groupby("Menu Category")["Total Nett Sales"]
            .sum()
            .nlargest(10)
            .sort_values(ascending=False)
        )

    # --- Data untuk Expander Top 10 ---
    # menu_sales di sini sudah bersih
    top_selling_items = menu_sales.groupby("Menu")["Qty"].sum().nlargest(10)
    top_grossing_items = (
        menu_sales.groupby("Menu")["Total Nett Sales"].sum().nlargest(10)
    )

    # --- Data untuk Expander Bottom 10 ---
    all_menu_sales = menu_sales.groupby("Menu")["Qty"].sum()
    bottom_selling_items = all_menu_sales.nsmallest(10).sort_values(ascending=True)
    all_menu_revenue = menu_sales.groupby("Menu")["Total Nett Sales"].sum()
    bottom_grossing_items = all_menu_revenue.nsmallest(10).sort_values(ascending=True)

    return (
        top_selling_items.reset_index(),
        top_grossing_items.reset_index(),
        top_selling_categories.reset_index(),
        top_grossing_categories.reset_index(),
        bottom_selling_items.reset_index(),
        bottom_grossing_items.reset_index(),
        menu_sales_cat_df,
    )


@st.cache_data
def get_operational_kpi(df):
    """Menghitung KPI Operasional (Jam Sibuk, Hari Sibuk, Durasi)."""
    avg_dining_time = 0.0
    if "Visit Purpose" in df.columns and "Sales Date Out" in df.columns:
        df_dine_in = df[
            df["Visit Purpose"].str.contains("DINE IN", na=False, case=False)
        ].copy()
        df_dine_in.dropna(subset=["Sales Date In", "Sales Date Out"], inplace=True)
        bill_times = df_dine_in.groupby("Bill Number").agg(
            Start=("Sales Date In", "min"), End=("Sales Date Out", "max")
        )
        bill_times["Duration_minutes"] = (
            bill_times["End"] - bill_times["Start"]
        ).dt.total_seconds() / 60
        bill_times_filtered = bill_times[
            (bill_times["Duration_minutes"] > 1)
            & (bill_times["Duration_minutes"] < 480)
        ]
        avg_dining_time = bill_times_filtered["Duration_minutes"].mean()
        if pd.isna(avg_dining_time):
            avg_dining_time = 0.0
    else:
        pass

    df_hourly = df.copy()
    df_hourly["Hour"] = df_hourly["Sales Date In"].dt.hour
    peak_hours = (
        df_hourly.groupby("Hour")["Bill Number"].nunique().sort_values(ascending=False)
    )

    day_map = {
        "Monday": "Senin",
        "Tuesday": "Selasa",
        "Wednesday": "Rabu",
        "Thursday": "Kamis",
        "Friday": "Jumat",
        "Saturday": "Sabtu",
        "Sunday": "Minggu",
    }
    df_daily = df.copy()
    df_daily.dropna(subset=["Sales Date In"], inplace=True)
    df_daily["Day Name"] = df_daily["Sales Date In"].dt.day_name().map(day_map)
    peak_days_of_week = (
        df_daily.groupby("Day Name")["Bill Number"].nunique().reset_index()
    )

    return avg_dining_time, peak_hours.reset_index(), peak_days_of_week


@st.cache_data
def analyze_profit(df_cogs):
    """Menganalisis profitabilitas HANYA dari file COGS."""
    if df_cogs is None or df_cogs.empty:
        final_cols = [
            "Menu",
            "Qty",
            "Harga Jual",
            "COGS",
            "Margin (Rp)",
            "Margin (%)",
            "Total Revenue (Rp)",
            "Total COGS (Rp)",
            "Total Profit (Rp)",
        ]
        return pd.DataFrame(columns=final_cols)

    profit_df = df_cogs.copy()

    # #############################################################
    # --- BLOK FILTER YANG DIPERBARUI (LEBIH KUAT) ---
    # #############################################################

    # Definisikan SEMUA kata kunci yang tidak diinginkan dalam satu regex
    # r"ADD[ -]?ON" -> akan menangkap "ADD ON", "ADD-ON", dan "ADDON"
    filter_regex_all = r"ADDITIONAL|ADD[ -]?ON|New Add-ons|Level"

    # Filter 1: Berdasarkan 'Menu Category' (jika kolomnya ada)
    if "Menu Category" in profit_df.columns:
        profit_df = profit_df[
            ~profit_df["Menu Category"].str.contains(
                filter_regex_all, na=False, case=False, regex=True
            )
        ]

    # Filter 2: Berdasarkan 'Menu' (kolom ini pasti ada)
    if "Menu" in profit_df.columns:
        profit_df = profit_df[
            ~profit_df["Menu"].str.contains(
                filter_regex_all, na=False, case=False, regex=True
            )
        ]

    # #############################################################
    # --- BATAS BLOK FILTER ---
    # #############################################################

    # Sisa fungsi berjalan seperti biasa, tapi di data yang sudah bersih
    profit_df["Margin (Rp)"] = profit_df["Harga Jual"] - profit_df["COGS"]
    # Total disini adalah Total Revenue/Penjualan (Harga Jual * Qty)
    profit_df["Total Revenue (Rp)"] = profit_df["Total"]
    profit_df["Total COGS (Rp)"] = profit_df["COGS"] * profit_df["Qty"]
    profit_df["Total Profit (Rp)"] = profit_df["Margin (Rp)"] * profit_df["Qty"]

    agg_df = (
        profit_df.groupby("Menu")
        .agg(
            Qty=("Qty", "sum"),
            Total_Revenue_Rp=("Total Revenue (Rp)", "sum"),
            Total_COGS_Rp=("Total COGS (Rp)", "sum"),
            Total_Profit_Rp=("Total Profit (Rp)", "sum"),
        )
        .reset_index()
    )

    agg_df["Margin (Rp)"] = np.where(
        agg_df["Qty"] > 0, agg_df["Total_Profit_Rp"] / agg_df["Qty"], 0
    )
    agg_df["Margin (%)"] = np.where(
        agg_df["Total_Revenue_Rp"] > 0,
        (agg_df["Total_Profit_Rp"] / agg_df["Total_Revenue_Rp"]) * 100,
        0,
    )

    unit_costs = (
        profit_df[profit_df["Harga Jual"] > 0]
        .groupby("Menu")
        .agg(Harga_Jual_Unit=("Harga Jual", "mean"), COGS_Unit=("COGS", "mean"))
        .reset_index()
    )

    final_df = pd.merge(agg_df, unit_costs, on="Menu", how="left")
    final_df.rename(
        columns={
            "Total_Revenue_Rp": "Total Revenue (Rp)",
            "Total_COGS_Rp": "Total COGS (Rp)",
            "Total_Profit_Rp": "Total Profit (Rp)",
            "Harga_Jual_Unit": "Harga Jual",
            "COGS_Unit": "COGS",
        },
        inplace=True,
    )

    final_cols = [
        "Menu",
        "Qty",
        "Harga Jual",
        "COGS",
        "Margin (Rp)",
        "Margin (%)",
        "Total Revenue (Rp)",
        "Total COGS (Rp)",
        "Total Profit (Rp)",
    ]
    for col in final_cols:
        if col not in final_df.columns:
            final_df[col] = 0
    final_df.fillna(0, inplace=True)

    return final_df[final_cols].sort_values(by="Total Profit (Rp)", ascending=False)


@st.cache_data
def get_peak_time_analysis(df):
    """Menganalisis transaksi berdasarkan waktu (Breakfast, Lunch, Dinner)."""
    if df is None or df.empty:
        return pd.DataFrame(
            columns=["Waktu Kunjungan", "Jumlah_Transaksi", "Total_Penjualan"]
        )

    bill_df = (
        df.groupby("Bill Number")
        .agg(
            Order_Time=("Order Time", "first"),
            Total_Sales=("Total After Bill Discount", "sum"),
        )
        .reset_index()
    )
    bill_df.dropna(subset=["Order_Time"], inplace=True)
    bill_df["Hour"] = bill_df["Order_Time"].dt.hour

    conditions = [
        (bill_df["Hour"] >= 10) & (bill_df["Hour"] < 12),
        (bill_df["Hour"] >= 12) & (bill_df["Hour"] < 17),
        (bill_df["Hour"] >= 17) & (bill_df["Hour"] < 22),
    ]
    choices = ["Breakfast/Brunch (10-12)", "Lunch (12-17)", "Dinner (17-22)"]
    bill_df["Waktu Kunjungan"] = np.select(conditions, choices, default="Luar Jam Buka")

    time_analysis = (
        bill_df.groupby("Waktu Kunjungan")
        .agg(
            Jumlah_Transaksi=("Bill Number", "nunique"),
            Total_Penjualan=("Total_Sales", "sum"),
        )
        .reset_index()
    )

    time_order = [
        "Breakfast/Brunch (10-12)",
        "Lunch (12-17)",
        "Dinner (17-22)",
        "Luar Jam Buka",
    ]
    try:
        time_analysis["Waktu Kunjungan"] = pd.Categorical(
            time_analysis["Waktu Kunjungan"], categories=time_order, ordered=True
        )
        time_analysis.sort_values("Waktu Kunjungan", inplace=True)
    except Exception as e:
        st.warning(f"Gagal mengurutkan waktu: {e}")

    return time_analysis


@st.cache_data
def get_waiter_performance(df):
    """Menganalisis performa waiter (Top 10)."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["Waiter", "Total_Penjualan", "Jumlah_Transaksi"])

    bill_df = (
        df.groupby("Bill Number")
        .agg(
            Waiter=("Waiter", "first"), Total_Sales=("Total After Bill Discount", "sum")
        )
        .reset_index()
    )
    bill_df["Waiter"] = bill_df["Waiter"].fillna("Tidak Diketahui")

    waiter_perf = (
        bill_df.groupby("Waiter")
        .agg(
            Total_Penjualan=("Total_Sales", "sum"),
            Jumlah_Transaksi=("Bill Number", "nunique"),
        )
        .reset_index()
    )
    return waiter_perf.nlargest(10, "Total_Penjualan")


# #################################################################
# --- BAGIAN GRAFIK YANG DIMODIFIKASI (MENGGUNAKAN PLOTLY) ---
# #################################################################


def create_horizontal_bar_chart(data, x_col, y_col, x_title, y_title, sort_order="-x"):
    """
    Membuat grafik batang horizontal Plotly Express yang profesional.
    """
    # Tentukan urutan sorting
    if sort_order == "-x":
        sort_ascending = False
    else:  # asumsikan "x"
        sort_ascending = True

    # Sort data untuk memastikan urutan bar di Plotly
    data_sorted = data.sort_values(by=x_col, ascending=sort_ascending)

    # Buat grafik
    fig = px.bar(
        data_sorted,
        x=x_col,
        y=y_col,
        orientation="h",
        labels={x_col: x_title, y_col: ""},  # Sembunyikan judul sumbu Y
        title=y_title,  # Gunakan y_title sebagai judul utama grafik
        color=x_col,
        color_continuous_scale=px.colors.sequential.Blues,  # Skala warna profesional
        template="plotly_white",
        text=x_col,  # Tambahkan label data
    )

    fig.update_layout(
        xaxis_title=x_title,
        yaxis_title="",  # Pastikan kosong
        xaxis_side="top",  # Pindahkan sumbu X ke atas
        coloraxis_showscale=False,  # Sembunyikan color bar
        title_x=0.01,  # Judul rata kiri
        title_font_size=18,
        margin=dict(l=0, r=20, t=60, b=20),  # Margin
        yaxis=(
            {"categoryorder": "total ascending"}
            if sort_ascending
            else {"categoryorder": "total descending"}
        ),
    )

    # Format label data dan tooltip
    fig.update_traces(
        texttemplate="%{x:.2s}",  # Format label (misal: 1.5M, 250k)
        textposition="outside",
        hovertemplate=f"<b>%{{y}}</b><br>{x_title}: %{{x:,.0f}}<extra></extra>",
    )

    # Hapus grid y-axis dan atur grid x-axis
    fig.update_yaxes(showgrid=False)
    fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor="#E5E5E5")

    return fig  # Kembalikan objek fig Plotly


def create_vertical_bar_chart(
    data, x_col, y_col, x_title, y_title, x_type="N", sort_order=None
):
    """
    Membuat grafik batang vertikal Plotly Express yang profesional.
    """

    # Siapkan urutan kategori jika ada
    category_orders = {}
    if sort_order:
        category_orders[x_col] = sort_order

    fig = px.bar(
        data,
        x=x_col,
        y=y_col,
        title=f"{y_title} vs {x_title}",
        labels={x_col: x_title, y_col: y_title},
        color=x_col,  # Warnai berdasarkan kategori x
        template="plotly_white",
        category_orders=category_orders,
        text=y_col,  # Tambahkan label data
    )

    fig.update_layout(
        xaxis_title=x_title,
        yaxis_title=y_title,
        showlegend=False,  # Legenda tidak perlu jika diwarnai by x
        title_x=0.01,  # Judul rata kiri
        title_font_size=18,
        margin=dict(l=0, r=0, t=60, b=0),
        yaxis_tickformat=".2s",  # Format sumbu Y (misal: 1.5M, 250k)
    )

    # Format label data dan tooltip
    fig.update_traces(
        texttemplate="%{y:.2s}",  # Format label (misal: 1.5M, 250k)
        textposition="outside",
        hovertemplate=f"<b>%{{x}}</b><br>{y_title}: %{{y:,.0f}}<extra></extra>",
    )

    # Hapus grid x-axis dan atur grid y-axis
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor="#E5E5E5")

    return fig  # Kembalikan objek fig Plotly


# #################################################################
# --- BATAS MODIFIKASI FUNGSI GRAFIK ---
# #################################################################


def calculate_delta(value_A, value_B, formatter_func, higher_is_better=True):
    """
    Menghitung delta antara A dan B, mengembalikan string dan warna.
    """
    delta_abs = value_A - value_B

    # Tentukan fungsi format
    if formatter_func == format_rupiah:
        delta_abs_formatted = formatter_func(delta_abs)
    elif formatter_func == format_angka_bulat:
        delta_abs_formatted = formatter_func(delta_abs)
    else:  # Asumsi format_persen atau float
        # Pastikan delta non-moneter diformat dengan 2 desimal
        delta_abs_formatted = f"{delta_abs:,.2f}"

    delta_pct_str = ""
    delta_color = "off"  # Default abu-abu

    if value_B != 0:
        delta_pct = (delta_abs / value_B) * 100
        arrow = "🔼" if delta_pct > 0 else "🔽"
        delta_pct_str = f"{arrow} {delta_pct:.1f}%"

        if delta_abs > 0:
            delta_color = (
                "normal" if higher_is_better else "inverse"
            )  # Hijau jika naik (dan naik itu bagus)
        elif delta_abs < 0:
            delta_color = (
                "inverse" if higher_is_better else "normal"
            )  # Merah jika turun (dan turun itu jelek)
    elif value_A != 0:
        # Kasus B=0 tapi A>0 (pertumbuhan tak terhingga)
        delta_pct_str = "🔼 100% +"
        delta_color = "normal" if higher_is_better else "inverse"
    else:
        # Kasus A=0 dan B=0
        delta_pct_str = "-"

    # Kembalikan 3 nilai: Nilai absolut, String Persen, dan Warna
    return delta_abs_formatted, delta_pct_str, delta_color


# #################################################################
# --- FUNGSI BARU UNTUK INSIGHT DI TAB 1 ---
# #################################################################
@st.cache_data
def generate_gmv_insights(
    kpi, top_selling, bottom_selling, peak_hours, peak_days_of_week
):
    """
    Menganalisis data GMV yang sudah diproses dan menghasilkan insight
    dalam bahasa alami.
    """
    insights = []

    # 1. Insight Menu Paling Laris (dari data top_selling)
    try:
        if not top_selling.empty:
            top_item = top_selling.iloc[0]["Menu"]
            top_qty = top_selling.iloc[0]["Qty"]
            insights.append(
                f"**🚀 Menu Paling Laris:** `{top_item}` adalah bintang utama Anda, "
                f"terjual sebanyak **{top_qty:,.0f} porsi** pada periode ini."
            )
    except Exception as e:
        print(f"Gagal generate insight top_selling: {e}")

    # 2. Insight Menu Jarang Laku (dari data bottom_selling)
    try:
        if not bottom_selling.empty:
            bottom_item = bottom_selling.iloc[0]["Menu"]
            bottom_qty = bottom_selling.iloc[0]["Qty"]
            insights.append(
                f"**📉 Menu Jarang Laku:** `{bottom_item}` perlu dievaluasi. "
                f"Menu ini hanya terjual **{bottom_qty:,.0f} porsi**."
            )
    except Exception as e:
        print(f"Gagal generate insight bottom_selling: {e}")

    # 3. Insight Hari Paling Ramai (dari data peak_days_of_week)
    try:
        if not peak_days_of_week.empty:
            peak_day_data = peak_days_of_week.sort_values(
                by="Bill Number", ascending=False
            ).iloc[0]
            peak_day = peak_day_data["Day Name"]
            peak_day_trx = peak_day_data["Bill Number"]
            insights.append(
                f"**🗓️ Hari Paling Ramai:** **{peak_day}** adalah hari tersibuk Anda "
                f"dengan total **{peak_day_trx:,.0f} transaksi**."
            )
    except Exception as e:
        print(f"Gagal generate insight peak_days: {e}")

    # 4. Insight Jam Paling Ramai (dari data peak_hours)
    try:
        if not peak_hours.empty:
            peak_hour_data = peak_hours.sort_values(
                by="Bill Number", ascending=False
            ).iloc[0]
            peak_hour = peak_hour_data["Hour"]
            peak_hour_trx = peak_hour_data["Bill Number"]
            insights.append(
                f"**🕒 Jam Paling Ramai:** Puncak kunjungan terjadi pada pukul **{peak_hour}:00**, "
                f"mencatat **{peak_hour_trx:,.0f} transaksi**."
            )
    except Exception as e:
        print(f"Gagal generate insight peak_hours: {e}")

    # 5. Insight Rata-rata Transaksi (dari data kpi)
    try:
        atv = kpi.get("Rata-rata Nilai Transaksi (ATV)", 0)
        ipb = kpi.get("Item per Transaksi (IPB)", 0)
        insights.append(
            f"**💸 Pola Belanja:** Rata-rata pelanggan Anda menghabiskan **{format_rupiah(atv)}** "
            f"dengan membeli **{ipb:.2f} item** per transaksi (IPB)."
        )
    except Exception as e:
        print(f"Gagal generate insight kpi: {e}")

    # 6. Catatan Penting tentang PROFIT
    insights.append(
        "**💡 Catatan Profit:** Untuk melihat *menu paling profit* (berdasarkan COGS), "
        "silakan cek di tab **'💰 COGS & Profit'**. Tab ini hanya menganalisis *penjualan* (GMV)."
    )

    return insights


# Fungsi Pemuatan Data (File Upload)

# #################################################################
# --- BAGIAN 2: FUNGSI PEMUATAN DATA (DENGAN CACHE) ---
# #################################################################


@st.cache_data
def load_data_gmv(uploaded_file, use_db=False):
    """Memuat dan membersihkan data GMV (File 1) DAN membaca headernya."""

    # --- PRIORITAS 1: Cekbox Database Dicentang ---
    if use_db:
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
                print("Database GMV kosong atau tidak ada.")
                return None, None, None, None

            return df, "DB_MODE", "DB_MODE", "DB_MODE"

    # --- PRIORITAS 2: Cekbox TIDAK Dicentang, TAPI Ada File di Uploader ---
    if uploaded_file is not None:
        df_data = None
        company_name = "N/A"
        period_str = "N/A"
        branch_name_header = "N/A"

        try:
            if uploaded_file.name.endswith(".xlsx"):
                df_header = pd.read_excel(
                    uploaded_file,
                    header=None,
                    nrows=7,
                    usecols=[0, 1],
                    engine="openpyxl",
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
                st.error(
                    f"Format file {uploaded_file.name} tidak didukung. Harap upload .xlsx atau .csv"
                )
                return None, None, None, None
        except Exception as e:
            st.error(f"Error membaca file GMV: {e}")
            st.error("Pastikan header file GMV ada di baris 10.")
            return None, None, None, None

        if df_data is not None:
            # Standarisasi nama kolom menjadi Title Case (Contoh: "sales date" -> "Sales Date")
            df_data.columns = [col.strip().title() for col in df_data.columns]
        else:
            st.error("Gagal memuat df_data.")
            return None, None, None, None

        # --- Konversi Tipe Data Numerik ---
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

        # --- Konversi Tanggal ---
        if "Sales Date In" in df_data.columns:
            df_data["Sales Date In"] = pd.to_datetime(
                df_data["Sales Date In"], errors="coerce"
            )
        if "Sales Date Out" in df_data.columns:
            df_data["Sales Date Out"] = pd.to_datetime(
                df_data["Sales Date Out"], errors="coerce"
            )
        if "Order Time" in df_data.columns:
            df_data["Order Time"] = pd.to_datetime(
                df_data["Order Time"], errors="coerce"
            )

        # --- PENGISIAN METADATA UTAMA ---
        df_data["Company"] = company_name
        df_data["Period"] = period_str

        # #################################################################
        # --- PERBAIKAN LOGIKA CABANG (BRANCH) ---
        # #################################################################

        # 1. Cari apakah ada kolom "Branch" atau "Outlet" di dalam data tabel
        found_branch_col = None
        for col in df_data.columns:
            if "BRANCH" in col.upper() or "OUTLET" in col.upper():
                found_branch_col = col
                break

        if found_branch_col:
            # Jika kolom ditemukan, rename menjadi 'Branch' standar jika namanya beda
            if found_branch_col != "Branch":
                df_data.rename(columns={found_branch_col: "Branch"}, inplace=True)

            # Isi nilai yang kosong (NaN) dengan header sebagai cadangan
            df_data["Branch"] = df_data["Branch"].fillna(branch_name_header)
        else:
            # Jika kolom branch sama sekali tidak ada di tabel, baru gunakan Header untuk semua baris
            df_data["Branch"] = branch_name_header

        # #################################################################
        # --- AKHIR PERBAIKAN ---
        # #################################################################

        # Validasi Akhir
        if (
            "Bill Number" not in df_data.columns
            or "Sales Date In" not in df_data.columns
        ):
            st.warning(
                "Peringatan: Kolom 'Bill Number' atau 'Sales Date In' tidak ditemukan setelah pembersihan."
            )
        else:
            df_data.dropna(subset=["Bill Number", "Sales Date In"], inplace=True)

        return df_data, company_name, period_str, branch_name_header

    # --- PRIORITAS 3: Cekbox TIDAK Dicentang, Uploader KOSONG ---
    return None, None, None, None


@st.cache_data
def load_pl_data(uploaded_file, use_db=False):
    """
    Memuat dan membersihkan data Profit & Loss (File 6).
    VERSI FIX: Menangani format akun dengan spasi dan deteksi tahun lebih kuat.
    """

    # --- Mode Database ---
    if use_db:
        with st.spinner("Memuat data P&L dari database..."):
            df = load_dataframe_from_db(
                "pl_data", date_cols=["Date"], numeric_cols_config={"Value": "float"}
            )
            return df

    # --- Mode Upload File ---
    if uploaded_file is not None:
        try:
            # 1. Baca Metadata untuk mendapatkan Tahun Periode
            period_year = (
                pd.Timestamp.now().year
            )  # Default ke tahun sekarang jika gagal

            uploaded_file.seek(0)
            # Baca 15 baris pertama untuk mencari metadata dan header
            if uploaded_file.name.endswith(".csv"):
                raw_content = pd.read_csv(uploaded_file, header=None, nrows=15)
            else:
                raw_content = pd.read_excel(uploaded_file, header=None, nrows=15)

            # Cari baris yang berisi kata "Period" untuk ambil tahun (misal: Period, 2025)
            try:
                for i, row in raw_content.iterrows():
                    row_str = row.astype(str).str.cat(sep=" ")
                    if "Period" in row_str:
                        # Cari angka 4 digit di baris ini (tahun)
                        found_years = re.findall(r"\b20\d{2}\b", row_str)
                        if found_years:
                            period_year = int(found_years[0])
                            break
            except:
                pass  # Gunakan default tahun ini

            # 2. Baca Data Utama (Header biasanya di baris ke-14, index 13)
            uploaded_file.seek(0)
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file, header=13)
            else:
                df = pd.read_excel(uploaded_file, header=13)

            # 3. Identifikasi Kolom
            # Kolom identitas adalah Account dan Description
            id_vars = [
                c for c in df.columns if "Account" in str(c) or "Description" in str(c)
            ]

            # Kolom nilai adalah yang mengandung nama bulan
            valid_months = [
                "January",
                "February",
                "March",
                "April",
                "May",
                "June",
                "July",
                "August",
                "September",
                "October",
                "November",
                "December",
            ]

            value_vars = []
            for c in df.columns:
                c_str = str(c)
                # Ambil kolom jika bukan ID dan mengandung nama bulan
                if c_str not in id_vars and any(m in c_str for m in valid_months):
                    value_vars.append(c)

            if not value_vars:
                st.error(
                    "Tidak ditemukan kolom bulan (January, February, dst) di file P&L."
                )
                return None

            # 4. Unpivot / Melt Data
            df_melted = df.melt(
                id_vars=id_vars,
                value_vars=value_vars,
                var_name="Raw_Column",
                value_name="Value",
            )

            # 5. Parsing Kolom Header (Tanggal & Branch)
            def parse_column(row):
                raw = str(row["Raw_Column"])
                is_last_year = "Last Year" in raw

                # Tentukan Bulan
                month_num = 1
                month_name = "Unknown"
                for i, m in enumerate(valid_months):
                    if m in raw:
                        month_num = i + 1
                        month_name = m
                        break

                # Tentukan Tahun
                # Jika header ada "Last Year", kurangi 1 dari period_year
                year = period_year - 1 if is_last_year else period_year

                # Buat Tanggal (Tanggal 1 setiap bulan)
                try:
                    date_val = pd.Timestamp(year=year, month=month_num, day=1)
                except:
                    date_val = pd.Timestamp.now()  # Fallback

                # Ekstrak Branch (String sebelum kurung buka pertama)
                if "(" in raw:
                    branch_name = raw.split("(")[0].strip()
                else:
                    branch_name = "All Branch"

                return pd.Series(
                    [
                        date_val,
                        month_name,
                        "Last Year" if is_last_year else "Current Year",
                        branch_name,
                    ]
                )

            df_melted[["Date", "Month_Name", "Year_Type", "Branch"]] = df_melted.apply(
                parse_column, axis=1
            )

            # 6. Bersihkan Nilai dan Kategori Akun
            # Hapus koma atau titik ribuan jika ada, lalu konversi ke float
            df_melted["Value"] = pd.to_numeric(
                df_melted["Value"], errors="coerce"
            ).fillna(0)

            # Kategorisasi Akun (PENTING: Bersihkan kode akun dari spasi/titik)
            def categorize_account(code):
                # Ubah ke string, hapus spasi, hapus titik
                code_clean = (
                    str(code).replace(" ", "").replace(".", "").replace("-", "").strip()
                )

                if code_clean.startswith("4"):
                    return "Revenue"
                elif code_clean.startswith("5"):
                    return "COGS"
                # Akun 6, 7, 8 biasanya Expense
                elif (
                    code_clean.startswith("6")
                    or code_clean.startswith("7")
                    or code_clean.startswith("8")
                ):
                    return "Expense"
                return "Other"

            df_melted["Category"] = df_melted["Account"].apply(categorize_account)

            # Hapus baris yang nilainya 0 untuk menghemat memori
            df_melted = df_melted[df_melted["Value"] != 0]

            return df_melted

        except Exception as e:
            st.error(f"Gagal memproses file P&L: {e}")
            return None
    return None


# #################################################################
# --- FUNGSI BARU UNTUK MEMUAT KALENDER (DARI DISKUSI KITA) ---
# --- (Letakkan ini di bawah `load_data_ulasan`) ---
# #################################################################


# KODE BARU (PERBAIKAN)
@st.cache_data
def load_kalender_data(file_path):  # <-- 1. Tambahkan 'file_path' di sini
    """
    Memuat data kalender event dari file CSV.
    (Versi ini menerima path sebagai argumen)
    """
    # 2. Gunakan 'file_path' yang Anda berikan
    kalender_file_path = file_path

    try:
        df = pd.read_csv(kalender_file_path)
        df["Tanggal"] = pd.to_datetime(df["Tanggal"]).dt.date
        return df
    except FileNotFoundError:
        st.error(f"FILE TIDAK DITEMUKAN: '{kalender_file_path}'")
        st.warning(
            "Harap jalankan skrip `buat_kalender_event.py` terlebih dahulu "
            "dan pastikan file .csv-nya ada di lokasi yang benar."
        )
        return None
    except Exception as e:
        st.error(f"Error saat membaca file kalender: {e}")
        return None


@st.cache_data
def load_cogs_data(uploaded_file, use_db=False):
    """Memuat dan membersihkan data COGS (File 2)."""

    # --- PRIORITAS 1: Cekbox Database Dicentang ---
    if use_db:
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
                print("Database COGS kosong atau tidak ada.")
                return None
            return df  # Kembalikan data dari DB

    # --- PRIORITAS 2: Cekbox TIDAK Dicentang, TAPI Ada File di Uploader ---
    if uploaded_file is not None:
        # --- Logika Asli (jika file di-upload) ---
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

        if "Branch" in df.columns:
            st.warning(
                "Peringatan: Kolom 'Branch' tidak ditemukan di File COGS. Filter cabang mungkin tidak berfungsi."
            )

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

    # --- PRIORITAS 3: Cekbox TIDAK Dicentang, Uploader KOSONG ---
    return None


@st.cache_data
def load_data_purchase(uploaded_file, use_db=False):
    """Memuat dan membersihkan data Laporan Pembelian (File 5)."""

    # --- PRIORITAS 1: Cekbox Database Dicentang ---
    if use_db:
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
                print("Database Pembelian kosong atau tidak ada.")
                return None
            return df

    # --- PRIORITAS 2: Cekbox TIDAK Dicentang, TAPI Ada File di Uploader ---
    if uploaded_file is not None:
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
            st.warning(
                "Peringatan: Kolom 'Branch' tidak ditemukan di File Pembelian. Filter cabang mungkin tidak berfungsi."
            )

        df.dropna(subset=["Purchase Number", "Purchase Date"], inplace=True)
        if "Branch" in df.columns:
            df["Branch"] = df["Branch"].astype(str).str.strip().str.title()
        return df

    # --- PRIORITAS 3: Cekbox TIDAK Dicentang, Uploader KOSONG ---
    return None


@st.cache_data
def load_data_waiter(uploaded_file, use_db=False):
    """Memuat dan membersihkan data Waiter (File 3)."""

    # --- PRIORITAS 1: Cekbox Database Dicentang ---
    if use_db:
        with st.spinner("Memuat data Waiter dari database..."):
            numeric_config = {"Total After Bill Discount": "float"}
            df = load_dataframe_from_db(
                "waiter_data",
                date_cols=["Order Time"],
                numeric_cols_config=numeric_config,
            )
            if df is None:
                print("Database Waiter kosong atau tidak ada.")
                return None
            return df  # Kembalikan data dari DB

    # --- PRIORITAS 2: Cekbox TIDAK Dicentang, TAPI Ada File di Uploader ---
    if uploaded_file is not None:
        # --- Logika Asli (jika file di-upload) ---
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

        required_cols = [
            "Bill Number",
            "Waiter",
            "Order Time",
            "Total After Bill Discount",
        ]
        if not all(col in df.columns for col in required_cols):
            st.error(
                f"File Waiter (File 3) harus memiliki kolom: {', '.join(required_cols)}"
            )
            st.error(f"Kolom ditemukan: {list(df.columns)}")
            return None

        if "Branch" not in df.columns:
            st.warning(
                "Peringatan: Kolom 'Branch' tidak ditemukan di File Waiter. Filter cabang mungkin tidak berfungsi."
            )

        df["Order Time"] = pd.to_datetime(df["Order Time"], errors="coerce")
        df["Total After Bill Discount"] = pd.to_numeric(
            df["Total After Bill Discount"], errors="coerce"
        ).fillna(0)
        df.dropna(subset=["Bill Number", "Order Time"], inplace=True)

        if "Branch" in df.columns:
            df["Branch"] = df["Branch"].astype(str).str.strip().str.title()

        return df

    # --- PRIORITAS 3: Cekbox TIDAK Dicentang, Uploader KOSONG ---
    return None


@st.cache_data
def load_data_ulasan(uploaded_file, use_db=False):
    """Memuat dan membersihkan data Ulasan (File 4)."""

    # --- PRIORITAS 1: Cekbox Database Dicentang ---
    if use_db:
        with st.spinner("Memuat data Ulasan dari database..."):
            numeric_config = {"Rating_Clean": "int"}
            df = load_dataframe_from_db(
                "ulasan_data", numeric_cols_config=numeric_config
            )
            if df is None:
                print("Database Ulasan kosong atau tidak ada.")
                return None
            return df  # Kembalikan data dari DB

    # --- PRIORITAS 2: Cekbox TIDAK Dicentang, TAPI Ada File di Uploader ---
    if uploaded_file is not None:
        # --- Logika Asli (jika file di-upload) ---
        df = None
        try:
            if uploaded_file.name.endswith(".xlsx"):
                df = pd.read_excel(uploaded_file, engine="openpyxl")
            elif uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file, encoding="latin1")
            else:
                st.error(
                    "Format file Ulasan (File 4) tidak didukung. Harap upload .xlsx atau .csv"
                )
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

    # --- PRIORITAS 3: Cekbox TIDAK Dicentang, Uploader KOSONG ---
    return None


# #################################################################
# --- BAGIAN 2.5: FUNGSI BARU UNTUK ANALISIS PEMBELIAN ---
# #################################################################


@st.cache_data
def analyze_purchase_data(df):
    """
    Menganalisis data pembelian yang sudah difilter.
    Sesuai permintaan: Menjumlahkan SEMUA kategori sebagai 'Total Biaya Pembelian'.
    """
    if df is None or df.empty:
        return 0, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # PENTING: Hanya analisis item yang memiliki biaya tercatat (Total > 0)
    # Ini akan mengabaikan item dari "Santhai HO LINK" atau "Pasar Cash" yang bernilai 0
    df_with_cost = df[df["Total"] > 0].copy()

    # 1. Total Biaya (Sesuai permintaan, SEMUA kategori dijumlahkan)
    total_cost = df_with_cost["Total"].sum()

    # 2. Biaya per Kategori
    cost_by_category = (
        df_with_cost.groupby("Category")["Total"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )

    # 3. Biaya per Supplier
    cost_by_supplier = (
        df_with_cost.groupby("Supplier Name")["Total"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )

    # 4. Top 20 Item Termahal
    top_items = (
        df_with_cost.groupby("Product Name")["Total"]
        .sum()
        .nlargest(20)
        .sort_values(ascending=False)
        .reset_index()
    )

    # 5. Data mentah yang sudah difilter (untuk ditampilkan)
    raw_data_filtered = df_with_cost[
        [
            "Purchase Date",
            "Supplier Name",
            "Category",
            "Sub Category",
            "Product Name",
            "Receipt Qty",
            "Price",
            "Total",
        ]
    ].sort_values(by="Total", ascending=False)

    return total_cost, cost_by_category, cost_by_supplier, top_items, raw_data_filtered


# Fungsi Pembangun UI (Interface)

# #################################################################
# --- BAGIAN 3: FUNGSI PEMBANGUN UI (INTERFACE) ---
# #################################################################


def build_sidebar():
    """Menggambar sidebar dan mengembalikan file yang di-upload."""

    # --- FUNGSI CALLBACK BARU ---
    def on_checkbox_change():
        """Memperbarui variabel KONTROL 'use_db' kita saat widget checkbox diklik."""
        st.session_state.use_db = st.session_state.use_db_widget_key

    # #################################################################
    # --- TAMBAHKAN FUNGSI BARU 'on_file_change' DI SINI ---
    # #################################################################
    def on_file_change():
        """
        Callback yang dipanggil SETIAP KALI file di uploader berubah
        (baik ditambah, diganti, atau dihapus).
        """

        # 1. Bersihkan semua cache pemuatan data
        load_data_gmv.clear()
        load_cogs_data.clear()
        load_data_waiter.clear()
        load_data_ulasan.clear()
        load_data_purchase.clear()

        # 2. (PERBAIKAN) Selalu nonaktifkan mode DB saat file diubah
        st.session_state.use_db = False

        # 3. Reset semua status "tersimpan"
        st.session_state.gmv_saved_status = False
        st.session_state.cogs_saved_status = False
        st.session_state.waiter_saved_status = False
        st.session_state.ulasan_saved_status = False
        st.session_state.purchase_saved_status = False

    # #################################################################
    # --- BATAS FUNGSI BARU ---
    # #################################################################

    with st.sidebar:
        st.markdown(
            """
            <div style='text-align: center; margin-bottom: 20px;'>
                <h2>DATA DRIVEN</h2>
                <h2>SPECIALYST FNB</h2>
                <p style='font-size: 0.9rem; color: #aaaaaa; margin-top: 5px;'> </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.header("Mode Pemuatan Data")

        # Inisialisasi 'use_db' jika belum ada (ini adalah variabel KONTROL kita)
        if "use_db" not in st.session_state:
            st.session_state.use_db = False

        # --- PERBAIKAN BESAR DI SINI ---
        # Kita memisahkan 'key' widget dari variabel 'use_db' kita
        use_db = st.checkbox(
            "Gunakan data terakhir dari database",
            value=st.session_state.use_db,  # Nilai checkbox diatur oleh variabel KONTROL kita
            key="use_db_widget_key",  # Widget ini punya 'key' UNIK sendiri
            on_change=on_checkbox_change,  # Saat diklik, panggil callback untuk update var KONTROL
            help="Centang untuk memuat data terakhir. Hapus centang untuk mengabaikan database. Meng-upload file baru akan otomatis mengabaikan database.",
        )
        # --- BATAS PERBAIKAN ---

        st.markdown("---")

        st.header("Upload Data")

        tipe_file_standar = [
            "xlsx",
            "csv",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "text/csv",
        ]

        # --- Uploader 1: GMV ---
        gmv_file = st.file_uploader(
            "1. Upload Laporan GMV (Operasional)",
            type=tipe_file_standar,
            key="uploader_gmv",
            on_change=on_file_change,  # <-- TAMBAHKAN INI
        )
        if gmv_file is not None:
            if st.session_state.gmv_saved_status:
                st.success("Data GMV berhasil disimpan!", icon="✅")
            else:
                if st.button("Simpan Data GMV ini ke Database", key="save_gmv"):
                    st.session_state.save_gmv_flag = True
        st.info("ℹ️ Laporan GMV asli (header di baris ke-10).")

        # --- Uploader 2: COGS ---
        cogs_file = st.file_uploader(
            "2. Upload Laporan COGS (Menu COGS Report)",
            type=tipe_file_standar,
            key="uploader_cogs",
            on_change=on_file_change,  # <-- TAMBAHKAN INI
        )
        if cogs_file is not None:
            if st.session_state.cogs_saved_status:
                # --- PERBAIKAN DI SINI ---
                st.success("Data COGS berhasil disimpan!", icon="✅")
            else:
                if st.button("Simpan Data COGS ini ke Database", key="save_cogs"):
                    st.session_state.save_cogs_flag = True
        st.info("ℹ️ File COGS (header di baris ke-13).")

        # --- Uploader 3: Waiter ---
        waiter_file = st.file_uploader(
            "3. Upload Sales Recapitulation Detail  (Rekapitulasi Detail)",
            type=tipe_file_standar,
            key="uploader_waiter",
            on_change=on_file_change,  # <-- TAMBAHKAN INI
        )
        if waiter_file is not None:
            if st.session_state.waiter_saved_status:
                # --- PERBAIKAN DI SINI ---
                st.success("Data Waiter berhasil disimpan!", icon="✅")
            else:
                if st.button("Simpan Data Waiter ini ke Database", key="save_waiter"):
                    st.session_state.save_waiter_flag = True
        st.info("ℹ️ Laporan Rekapitulasi Detail (header di baris ke-12).")

        # --- Uploader 4: Ulasan ---
        ulasan_file = st.file_uploader(
            "4. Upload Laporan Ulasan Pelanggan",
            type=tipe_file_standar,
            key="uploader_ulasan",
            on_change=on_file_change,  # <-- TAMBAHKAN INI
        )
        if ulasan_file is not None:
            if st.session_state.ulasan_saved_status:
                # --- PERBAIKAN DI SINI ---
                st.success("Data Ulasan berhasil disimpan!", icon="✅")
            else:
                if st.button("Simpan Data Ulasan ini ke Database", key="save_ulasan"):
                    st.session_state.save_ulasan_flag = True
        st.info("ℹ️ File .csv atau .xlsx berisi kolom: Nama, Rating, Ulasan.")

        # --- UPLOADER KE-5 ---
        purchase_file = st.file_uploader(
            "5. Upload Laporan Pembelian (Purchase Recapitulation)",
            type=tipe_file_standar,
            key="uploader_purchase",
            on_change=on_file_change,  # <-- TAMBAHKAN INI
        )
        if purchase_file is not None:
            if st.session_state.purchase_saved_status:
                # --- PERBAIKAN DI SINI ---
                st.success("Data Pembelian berhasil disimpan!", icon="✅")
            else:
                if st.button(
                    "Simpan Data Pembelian ini ke Database", key="save_purchase"
                ):
                    st.session_state.save_purchase_flag = True
        st.info("ℹ️ Laporan Pembelian (header di baris ke-12).")
        # --- BATAS TAMBAHAN ---

        # --- UPLOADER KE-6 (P&L) ---
        pl_file = st.file_uploader(
            "6. Upload Profit Loss Report (Laba Rugi)",
            type=tipe_file_standar,
            key="uploader_pl",
            on_change=on_file_change,
        )
        if pl_file is not None:
            if st.session_state.get(
                "pl_saved_status"
            ):  # Pastikan inisialisasi state ini di main()
                st.success("Data P&L berhasil disimpan!", icon="✅")
            else:
                if st.button("Simpan Data P&L ini ke Database", key="save_pl"):
                    st.session_state.save_pl_flag = True
        st.info("ℹ️ File Report CSV/Excel (Format Wide dengan kolom bulan).")

    # Update return statement sidebar
    return (
        gmv_file,
        cogs_file,
        waiter_file,
        ulasan_file,
        purchase_file,
        pl_file,  # <--- Tambahkan ini
        st.session_state.use_db,
    )

    # Kembalikan file DAN status checkbox
    return (
        gmv_file,
        cogs_file,
        waiter_file,
        ulasan_file,
        purchase_file,
        st.session_state.use_db,  # Kita tetap kembalikan var KONTROL kita
    )


# FILTER GLOBAL


def build_global_filters(data_gmv, data_cogs, data_waiter, data_purchase, data_pl):
    """
    Menggambar filter global.
    MODIFIKASI: Filter Bulanan sekarang menggunakan Dropdown (Bulan & Tahun), bukan Date Picker.
    """

    # Inisialisasi data yang difilter sebagai data asli
    filtered_gmv = data_gmv
    filtered_cogs = data_cogs
    filtered_waiter = data_waiter
    filtered_purchase = data_purchase
    filtered_pl = data_pl

    # --- 1. Tentukan Rentang Tanggal Master & Data Source Utama ---
    master_min_date = pd.Timestamp.now().date()
    master_max_date = pd.Timestamp.now().date()
    filter_source_df = None  # Dataframe utama untuk mengambil daftar bulan
    date_col_ref = None  # Nama kolom tanggal di dataframe utama

    try:
        if data_gmv is not None and not data_gmv.empty:
            master_min_date = data_gmv["Sales Date In"].min().date()
            master_max_date = data_gmv["Sales Date In"].max().date()
            filter_source_df = data_gmv
            date_col_ref = "Sales Date In"
        elif data_cogs is not None and not data_cogs.empty:
            master_min_date = data_cogs["Sales Date"].min().date()
            master_max_date = data_cogs["Sales Date"].max().date()
            filter_source_df = data_cogs
            date_col_ref = "Sales Date"
        # ... (Logika fallback lainnya biarkan saja jika ada) ...
    except Exception as e:
        st.error(f"Gagal membaca rentang tanggal: {e}")

    # --- 2. TAMPILKAN WIDGET FILTER ---
    if filter_source_df is not None or data_pl is not None:
        st.subheader("Filter Analisis Global")

        # --- FILTER CABANG ---
        if data_gmv is not None and "Branch" in data_gmv.columns:
            all_branches = sorted(data_gmv["Branch"].unique())
            selected_branches = st.multiselect(
                "Pilih Cabang (Branch):",
                options=all_branches,
                default=all_branches,
                key="branch_filter",
            )

            # Terapkan filter cabang
            filtered_gmv = data_gmv[data_gmv["Branch"].isin(selected_branches)]

            if selected_branches:
                if data_cogs is not None and "Branch" in data_cogs.columns:
                    filtered_cogs = data_cogs[
                        data_cogs["Branch"].isin(selected_branches)
                    ]
                if data_waiter is not None and "Branch" in data_waiter.columns:
                    filtered_waiter = data_waiter[
                        data_waiter["Branch"].isin(selected_branches)
                    ]
                if data_purchase is not None and "Branch" in data_purchase.columns:
                    filtered_purchase = data_purchase[
                        data_purchase["Branch"].isin(selected_branches)
                    ]
                if data_pl is not None and "Branch" in data_pl.columns:
                    filtered_pl = data_pl[data_pl["Branch"].isin(selected_branches)]

        # --- FILTER WAKTU ---
        filter_type = st.radio(
            "Pilih rentang waktu:",
            ["Semua Periode", "Harian", "Mingguan", "Bulanan", "Tahunan"],
            horizontal=True,
            key="filter_type_global",
        )

        if filter_type == "Harian":
            selected_date = st.date_input(
                "Pilih Tanggal",
                value=master_max_date,
                min_value=master_min_date,
                max_value=master_max_date,
                key="global_date",
            )
            if filtered_gmv is not None:
                filtered_gmv = filtered_gmv[
                    filtered_gmv["Sales Date In"].dt.date == selected_date
                ]
            if filtered_cogs is not None:
                filtered_cogs = filtered_cogs[
                    filtered_cogs["Sales Date"].dt.date == selected_date
                ]
            if filtered_waiter is not None:
                filtered_waiter = filtered_waiter[
                    filtered_waiter["Order Time"].dt.date == selected_date
                ]
            if filtered_purchase is not None:
                filtered_purchase = filtered_purchase[
                    filtered_purchase["Purchase Date"].dt.date == selected_date
                ]

        elif filter_type == "Mingguan":
            default_start = master_max_date - pd.to_timedelta(6, unit="d")
            if default_start < master_min_date:
                default_start = master_min_date

            start_date = st.date_input(
                "Pilih Tanggal Mulai (7 hari)",
                value=default_start,
                min_value=master_min_date,
                max_value=master_max_date,
                key="global_week_start",
            )
            end_date = start_date + pd.to_timedelta(6, unit="d")
            st.info(
                f"Menampilkan: {start_date.strftime('%d-%m-%Y')} s.d. {end_date.strftime('%d-%m-%Y')}"
            )

            if filtered_gmv is not None:
                filtered_gmv = filtered_gmv[
                    (filtered_gmv["Sales Date In"].dt.date >= start_date)
                    & (filtered_gmv["Sales Date In"].dt.date <= end_date)
                ]
            if filtered_cogs is not None:
                filtered_cogs = filtered_cogs[
                    (filtered_cogs["Sales Date"].dt.date >= start_date)
                    & (filtered_cogs["Sales Date"].dt.date <= end_date)
                ]
            if filtered_waiter is not None:
                filtered_waiter = filtered_waiter[
                    (filtered_waiter["Order Time"].dt.date >= start_date)
                    & (filtered_waiter["Order Time"].dt.date <= end_date)
                ]
            if filtered_purchase is not None:
                filtered_purchase = filtered_purchase[
                    (filtered_purchase["Purchase Date"].dt.date >= start_date)
                    & (filtered_purchase["Purchase Date"].dt.date <= end_date)
                ]

        # #############################################################
        # --- PERUBAHAN UTAMA: FILTER BULANAN (DROPDOWN) ---
        # #############################################################
        elif filter_type == "Bulanan":
            # 1. Siapkan Opsi Bulan dari Data yang Ada
            period_options = []
            period_map = {}  # Untuk memetakan "November 2025" -> (11, 2025)

            if filter_source_df is not None and date_col_ref is not None:
                # Buat kolom sementara Period (Y-M) untuk sorting
                temp_df = filter_source_df[[date_col_ref]].copy()
                temp_df["Period_Obj"] = temp_df[date_col_ref].dt.to_period("M")

                # Ambil periode unik dan urutkan descending (terbaru di atas)
                unique_periods = sorted(temp_df["Period_Obj"].unique(), reverse=True)

                for p in unique_periods:
                    label = p.strftime("%B %Y")  # Contoh: "November 2025"
                    period_options.append(label)
                    period_map[label] = (p.month, p.year)

            # Fallback jika data kosong
            if not period_options:
                current_month_label = pd.Timestamp.now().strftime("%B %Y")
                period_options = [current_month_label]
                period_map[current_month_label] = (
                    pd.Timestamp.now().month,
                    pd.Timestamp.now().year,
                )

            # 2. Tampilkan Selectbox
            selected_month_str = st.selectbox(
                "Pilih Bulan:", options=period_options, key="global_month_select"
            )

            # 3. Ambil Angka Bulan dan Tahun dari pilihan
            sel_month, sel_year = period_map[selected_month_str]

            # st.info(f"Menampilkan data bulan: {selected_month_str}")

            # 4. Terapkan Filter
            if filtered_gmv is not None:
                filtered_gmv = filtered_gmv[
                    (filtered_gmv["Sales Date In"].dt.month == sel_month)
                    & (filtered_gmv["Sales Date In"].dt.year == sel_year)
                ]
            if filtered_cogs is not None:
                filtered_cogs = filtered_cogs[
                    (filtered_cogs["Sales Date"].dt.month == sel_month)
                    & (filtered_cogs["Sales Date"].dt.year == sel_year)
                ]
            if filtered_waiter is not None:
                filtered_waiter = filtered_waiter[
                    (filtered_waiter["Order Time"].dt.month == sel_month)
                    & (filtered_waiter["Order Time"].dt.year == sel_year)
                ]
            if filtered_purchase is not None:
                filtered_purchase = filtered_purchase[
                    (filtered_purchase["Purchase Date"].dt.month == sel_month)
                    & (filtered_purchase["Purchase Date"].dt.year == sel_year)
                ]
        # #############################################################
        # --- AKHIR PERUBAHAN ---
        # #############################################################

        elif filter_type == "Tahunan":
            # Gunakan logika serupa (ambil tahun unik dari data) untuk UX lebih baik
            year_options = []
            if filter_source_df is not None and date_col_ref is not None:
                unique_years = sorted(
                    filter_source_df[date_col_ref].dt.year.unique(), reverse=True
                )
                year_options = unique_years

            if not year_options:
                year_options = [pd.Timestamp.now().year]

            sel_year = st.selectbox(
                "Pilih Tahun:", options=year_options, key="global_year_select"
            )
            # st.info(f"Menampilkan data tahun {sel_year}")

            if filtered_gmv is not None:
                filtered_gmv = filtered_gmv[
                    filtered_gmv["Sales Date In"].dt.year == sel_year
                ]
            if filtered_cogs is not None:
                filtered_cogs = filtered_cogs[
                    filtered_cogs["Sales Date"].dt.year == sel_year
                ]
            if filtered_waiter is not None:
                filtered_waiter = filtered_waiter[
                    filtered_waiter["Order Time"].dt.year == sel_year
                ]
            if filtered_purchase is not None:
                filtered_purchase = filtered_purchase[
                    filtered_purchase["Purchase Date"].dt.year == sel_year
                ]

    elif (
        data_gmv is None
        and data_cogs is None
        and data_waiter is None
        and data_purchase is None
    ):
        st.markdown("---")

    return filtered_gmv, filtered_cogs, filtered_waiter, filtered_purchase, filtered_pl


# --- INI FUNGSI TAB 1 PENGGANTI YANG SUDAH LENGKAP ---


def build_tab1_sales(filtered_gmv):
    """
    Menggambar semua elemen untuk Tab 1.
    (VERSI DENGAN FORMAT JAM BARU dan INSIGHT DI BAWAH)
    """
    if filtered_gmv is not None:
        if not filtered_gmv.empty:
            start_date = filtered_gmv["Sales Date In"].min().strftime("%d-%m-%Y")
            end_date = filtered_gmv["Sales Date In"].max().strftime("%d-%m-%Y")
            st.subheader(f"Periode Analisis: {start_date} s.d. {end_date}")

            # --- [BLOK BARU UNTUK FILTER KUSTOM] ---
            with st.expander("⚙️ Pengaturan Filter Menu"):
                st.info(
                    "Masukkan nama menu yang ingin DIKECUALIKAN dari analisis Top/Bottom. Pisahkan dengan tanda | (pipa)."
                )

                # Ini adalah gabungan dari SEMUA filter Anda sebelumnya
                default_filter_string = (
                    r"Ocha|Refill|Mineral Water|ADD[ -]?ON|ADDITIONAL|Level"
                )

                filter_regex_items_input = st.text_input(
                    "Menu yang Dikecualikan (Regex)",
                    default_filter_string,
                    help="Menggunakan Regex. | berarti ATAU. 'ADD[ -]?ON' akan memfilter 'ADD ON' dan 'ADD-ON'.",
                )
            # --- [AKHIR BLOK BARU] ---

            # === 1. HITUNG SEMUA DATA DULU (PENTING!) ===
            kpi = calculate_sales_kpi(filtered_gmv)
            (
                top_selling,
                top_grossing,
                top_sell_cat,
                top_gross_cat,
                bottom_selling,
                bottom_grossing,
                menu_sales_cat_df,
            ) = get_menu_performance(
                filtered_gmv, filter_regex_items_input
            )  # <-- KIRIM ARGUMEN

            # === 2. TAMPILKAN KPI UTAMA (Expander Asli Anda) ===
            with st.expander(
                "📈 KPI Kinerja Penjualan (Revenue, ATV, IPB)", expanded=True
            ):
                # ... (sisa kode KPI Anda tidak berubah) ...
                st.header("📊 KPI Kinerja Penjualan")
                col1, col2, col3 = st.columns(3)
                col1.metric(
                    "💰 Total Pendapatan Kotor",
                    format_rupiah(kpi["Total Pendapatan Kotor"]),
                )
                col2.metric(
                    "💵 Total Penjualan Bersih (Nett)",
                    format_rupiah(kpi["Total Penjualan Bersih (Nett)"]),
                )
                col3.metric("🧾 Total Transaksi", f"{kpi['Total Transaksi']} Transaksi")

                col4, col5, col6 = st.columns(3)
                col4.metric(
                    "💸 Rata-rata Nilai Transaksi (ATV)",
                    format_rupiah(kpi["Rata-rata Nilai Transaksi (ATV)"]),
                )
                col5.metric(
                    "📦 Total Item Terjual", f"{kpi['Total Item Terjual']:,.0f} Items"
                )
                col6.metric(
                    "🛍️ Item per Transaksi (IPB)",
                    f"{kpi['Item per Transaksi (IPB)']:.2f}",
                )

                with st.expander("Lihat Rincian Pendapatan (Diskon, Service, Pajak)"):
                    exp_col1, exp_col2, exp_col3 = st.columns(3)
                    exp_col1.metric(
                        "📉 Total Diskon", format_rupiah(kpi["Total Diskon"])
                    )
                    exp_col2.metric(
                        "🛎️ Total Service Charge",
                        format_rupiah(kpi["Total Service Charge"]),
                    )
                    exp_col3.metric("🧾 Total Pajak", format_rupiah(kpi["Total Pajak"]))

            st.markdown("---")

            # === 3. SISA TAB (Semua expander Anda yang lain) ===
            st.header("🍽️ Analisis Menu & Kategori")

            if "Menu Category" in filtered_gmv.columns and not menu_sales_cat_df.empty:
                # ... (sisa kode expander kategori interaktif tidak berubah) ...
                with st.expander(
                    "🍰 Analisis Kategori Menu Interaktif (Klik untuk Detail)",
                    expanded=True,
                ):
                    st.subheader("Grafik Kuantiti Terjual (Drill-Down)")
                    st.info(
                        "Klik pada salah satu batang di **Grafik Kategori** untuk melihat detail itemnya di **Grafik Menu Item** di bawah."
                    )

                    data_kategori = (
                        menu_sales_cat_df.groupby("Menu Category")["Qty"]
                        .sum()
                        .reset_index()
                    )
                    data_kategori_sorted = data_kategori.sort_values(
                        "Qty", ascending=False
                    )

                    N_TOP = 15
                    top_n_categories = data_kategori_sorted.head(N_TOP)
                    other_categories = data_kategori_sorted.iloc[N_TOP:]

                    if not other_categories.empty:
                        other_sum = pd.DataFrame(
                            {
                                "Menu Category": ["Lainnya (Others)"],
                                "Qty": [other_categories["Qty"].sum()],
                            }
                        )
                        data_kategori_display_default = pd.concat(
                            [top_n_categories, other_sum]
                        ).reset_index(drop=True)
                    else:
                        data_kategori_display_default = top_n_categories.reset_index(
                            drop=True
                        )

                    st.markdown("---")
                    show_all_categories = st.checkbox(
                        "Tampilkan Semua Kategori (Drill-Down)",
                        value=False,
                        key="checkbox_kategori_tab1",
                        help="Centang untuk menampilkan semua kategori (akan mengaktifkan scroll vertikal jika datanya sangat banyak)",
                    )
                    st.markdown("---")

                    if show_all_categories:
                        data_untuk_grafik_atas = data_kategori_sorted
                        title_grafik_atas = "Total Penjualan per Kategori (Semua)"
                    else:
                        data_untuk_grafik_atas = data_kategori_display_default
                        title_grafik_atas = (
                            f"Total Penjualan per Kategori (Top {N_TOP} & Lainnya)"
                        )

                    data_menu_item = (
                        menu_sales_cat_df.groupby(["Menu Category", "Menu"])["Qty"]
                        .sum()
                        .reset_index()
                    )

                    selection_kategori = alt.selection_point(fields=["Menu Category"])

                    bar_height_px = 25
                    max_height_before_scroll = 400
                    min_height_px = 150
                    num_bars_kategori = len(data_untuk_grafik_atas)
                    ideal_height = num_bars_kategori * bar_height_px

                    chart_height_kategori = min(
                        max(ideal_height, min_height_px), max_height_before_scroll
                    )

                    chart_kategori = (
                        alt.Chart(data_untuk_grafik_atas)
                        .mark_bar()
                        .encode(
                            x=alt.X(
                                "Qty:Q",
                                title="Total Kuantiti Terjual",
                                axis=alt.Axis(orient="top"),
                            ),
                            y=alt.Y(
                                "Menu Category:N",
                                title="Kategori Menu",
                                sort="-x",
                                axis=alt.Axis(labelLimit=300),
                            ),
                            tooltip=["Menu Category", "Qty"],
                            color=alt.condition(
                                selection_kategori,
                                alt.value("orange"),
                                alt.value("steelblue"),
                            ),
                        )
                        .add_params(selection_kategori)
                        .properties(
                            title=title_grafik_atas,
                            height=chart_height_kategori,
                        )
                    )

                    num_bars_detail = len(data_menu_item["Menu"].unique())
                    ideal_height_detail = num_bars_detail * bar_height_px
                    chart_height_detail = min(
                        max(ideal_height_detail, min_height_px),
                        max_height_before_scroll,
                    )

                    chart_detail = (
                        alt.Chart(data_menu_item)
                        .mark_bar()
                        .encode(
                            x=alt.X(
                                "Qty:Q",
                                title="Total Kuantiti Terjual",
                                axis=alt.Axis(orient="top"),
                            ),
                            y=alt.Y(
                                "Menu:N",
                                title="Menu Item",
                                sort="-x",
                                axis=alt.Axis(labelLimit=300),
                            ),
                            tooltip=["Menu Category", "Menu", "Qty"],
                        )
                        .transform_filter(selection_kategori)
                        .properties(
                            title="Detail Penjualan per Menu Item (Berdasarkan Kategori Dipilih)",
                            height=chart_height_detail,
                        )
                    )

                    combined_chart = alt.vconcat(
                        chart_kategori, chart_detail, spacing=40
                    ).resolve_scale(y="independent")

                    st.altair_chart(combined_chart, use_container_width=True)

            elif "Menu Category" in filtered_gmv.columns:
                st.warning("Data kategori menu tidak ditemukan untuk periode ini.")
            else:
                st.warning(
                    "Kolom 'Menu Category' tidak ditemukan di File 1 untuk analisis interaktif."
                )

            # === 3. URUTAN BARU 2: BOTTOM 10 MENU ===
            with st.expander("📉 KPI Kinerja Menu (Bottom 10 Selling/Grossing)"):
                st.subheader("📉 Performa Menu Kurang Laku (Bottom 10)")
                col13, col14 = st.columns(2)
                with col13:
                    st.markdown("##### Menu Paling Jarang Terjual (by Kuantitas)")
                    st.dataframe(
                        bottom_selling.set_index("Menu").style.format(
                            {"Qty": format_angka_bulat}
                        )
                    )
                with col14:
                    st.markdown("##### Menu Pendapatan Terendah (by Nett Sales)")
                    st.dataframe(
                        bottom_grossing.set_index("Menu").style.format(
                            {"Total Nett Sales": format_rupiah}
                        )
                    )

            # === 4. URUTAN BARU 3: TOP 10 MENU ===
            with st.expander("🏆 KPI Kinerja Menu (Top 10 Selling/Grossing)"):
                st.subheader("🏆 Performa Menu Teratas (Top 10)")
                col9, col10 = st.columns(2)
                with col9:
                    st.markdown("##### Menu Terlaris (by Kuantitas)")
                    st.dataframe(
                        top_selling.set_index("Menu").style.format(
                            {"Qty": format_angka_bulat}
                        )
                    )
                with col10:
                    st.markdown("##### Menu Pendapatan Tertinggi (by Nett Sales)")
                    st.dataframe(
                        top_grossing.set_index("Menu").style.format(
                            {"Total Nett Sales": format_rupiah}
                        )
                    )

                col11, col12 = st.columns(2)
                with col11:
                    chart = create_horizontal_bar_chart(
                        top_selling,
                        "Qty",
                        "Menu",
                        "Kuantitas Terjual",
                        "Menu Terlaris (by Kuantitas)",
                    )
                    st.plotly_chart(chart, use_container_width=True)
                with col12:
                    chart = create_horizontal_bar_chart(
                        top_grossing,
                        "Total Nett Sales",
                        "Menu",
                        "Total Nett Sales (Rp)",
                        "Menu Pendapatan Tertinggi (by Nett Sales)",
                    )
                    st.plotly_chart(chart, use_container_width=True)

            st.markdown("---")

            # === 5. URUTAN BARU 4: KPI OPERASIONAL ===
            st.header("⚙️ Analisis Operasional & Pelanggan")

            with st.expander(
                "⚙️ KPI Operasional (Jam Sibuk, Hari Sibuk, Durasi Makan)", expanded=True
            ):
                avg_time, peak_hours, peak_days_of_week = get_operational_kpi(
                    filtered_gmv
                )

                if avg_time > 0:
                    st.metric(
                        "⏱️ Rata-rata Durasi Makan (Dine In)", f"{avg_time:.1f} menit"
                    )

                col19, col20 = st.columns(2)

                with col19:
                    st.subheader("🕒 Jam Sibuk (Berdasarkan Transaksi)")

                    peak_hours_formatted = peak_hours.copy()
                    peak_hours_formatted["Jam_Label"] = peak_hours_formatted[
                        "Hour"
                    ].apply(lambda h: f"{h:02d}:00")
                    hour_sort_order = peak_hours_formatted.sort_values(by="Hour")[
                        "Jam_Label"
                    ].tolist()

                    chart = create_vertical_bar_chart(
                        peak_hours_formatted,  # <-- Gunakan data baru
                        "Jam_Label",  # <-- Ganti "Hour" dengan kolom label baru
                        "Bill Number",
                        "Jam",
                        "Jumlah Transaksi",
                        x_type="O",
                        sort_order=hour_sort_order,  # <-- Berikan urutan sort manual
                    )
                    st.plotly_chart(chart, use_container_width=True)

                with col20:
                    st.subheader("🗓️ Hari Sibuk (Berdasarkan Transaksi)")
                    day_sort_order = [
                        "Senin",
                        "Selasa",
                        "Rabu",
                        "Kamis",
                        "Jumat",
                        "Sabtu",
                        "Minggu",
                    ]
                    chart = create_vertical_bar_chart(
                        peak_days_of_week,
                        "Day Name",
                        "Bill Number",
                        "Hari",
                        "Jumlah Transaksi",
                        x_type="O",
                        sort_order=day_sort_order,
                    )
                    st.plotly_chart(chart, use_container_width=True)

            # === 6. URUTAN BARU 5: ANALISIS TRANSAKSI (TATA LETAK VERTIKAL) ===
            with st.expander(
                "💳 Analisis Transaksi (Pembayaran & Kunjungan)", expanded=True
            ):

                if "Payment Method" in filtered_gmv.columns:
                    st.subheader("💳 Penjualan per Metode Pembayaran")
                    payment_data = get_payment_analysis(filtered_gmv)
                    chart = create_horizontal_bar_chart(
                        payment_data,
                        "Total_Penjualan",
                        "Cleaned_Payment",
                        "Total Penjualan (Rp)",
                        "Penjualan per Metode Pembayaran",
                    )
                    st.plotly_chart(chart, use_container_width=True)
                else:
                    st.warning("Kolom 'Payment Method' tidak ditemukan di File 1.")

                st.markdown("---")

                if "Visit Purpose" in filtered_gmv.columns:
                    st.subheader("🏪 Penjualan per Tipe Kunjungan")
                    visit_data = get_visit_purpose_analysis(filtered_gmv)
                    chart = create_horizontal_bar_chart(
                        visit_data,
                        "Total After Bill Discount",
                        "Visit Purpose",
                        "Total Penjualan (Rp)",
                        "Penjualan per Tipe Kunjungan",
                    )
                    st.plotly_chart(chart, use_container_width=True)
                else:
                    st.warning("Kolom 'Visit Purpose' tidak ditemukan di File 1.")

            # #############################################################
            # --- BLOK INSIGHT (TETAP DI PALING BAWAH) ---
            # #############################################################

            st.markdown("---")  # Tambahkan pemisah visual
            st.header("💡 Insight Otomatis (Ringkasan)")

            # Panggil fungsi 'pencari insight' kita
            # Kita gunakan data yang sudah dihitung di atas

            insights = generate_gmv_insights(
                kpi, top_selling, bottom_selling, peak_hours, peak_days_of_week
            )

            # Tampilkan dalam expander baru
            with st.expander(
                "Klik untuk melihat Temuan Kunci dari Data Penjualan Anda",
                expanded=True,
            ):
                if insights:
                    for insight in insights:
                        st.markdown(f"&bull; {insight}")  # Tampilkan sebagai daftar
                else:
                    st.info(
                        "Tidak ada insight otomatis yang dapat dibuat dari data ini."
                    )

        elif filtered_gmv is not None and filtered_gmv.empty:
            st.warning(
                "Tidak ada data ditemukan di File GMV untuk rentang waktu yang dipilih."
            )
    else:
        st.info(
            "Silakan upload file Laporan GMV (File 1) di sidebar untuk melihat analisis penjualan."
        )


# #################################################################
# --- FUNGSI BARU UNTUK INSIGHT DI TAB 2 (COGS) ---
# #################################################################
@st.cache_data(show_spinner=False)
def generate_cogs_insights(profit_df):
    """
    Menganalisis data profit_df yang sudah diproses dan menghasilkan insight
    dalam bahasa alami untuk Tab 2.
    """
    insights = []

    # Pastikan data tidak kosong
    if profit_df is None or profit_df.empty:
        return ["Tidak ada data profit untuk dianalisis."]

    try:
        # 1. Insight Profitabilitas Keseluruhan
        total_revenue = profit_df["Total Revenue (Rp)"].sum()
        total_profit = profit_df["Total Profit (Rp)"].sum()
        avg_margin_percent = (
            (total_profit / total_revenue) * 100 if total_revenue > 0 else 0
        )

        insights.append(
            f"**💰 Profitabilitas Umum:** Dari total revenue **{format_rupiah(total_revenue)}** (berdasarkan file COGS), "
            f"Anda menghasilkan profit kotor sebesar **{format_rupiah(total_profit)}**, "
            f"dengan rata-rata margin profit **{avg_margin_percent:,.1f}%**."
        )

        # Filter item yang valid (pernah terjual)
        df_valid = profit_df[profit_df["Qty"] > 0].copy()
        if not df_valid.empty:

            # 2. Insight Menu Paling Untung (Rp)
            top_profit_item = df_valid.nlargest(1, "Total Profit (Rp)").iloc[0]
            insights.append(
                f"**🏆 Bintang Profit (Rp):** `{top_profit_item['Menu']}` adalah penyumbang profit terbesar Anda, "
                f"menghasilkan **{format_rupiah(top_profit_item['Total Profit (Rp)'])}** sendirian."
            )

            # 3. Insight Menu Margin Tertinggi (%)
            top_margin_item = df_valid.nlargest(1, "Margin (%)").iloc[0]
            insights.append(
                f"**📈 Efisiensi Terbaik (%):** `{top_margin_item['Menu']}` adalah menu paling efisien "
                f"dengan margin profit **{format_persen(top_margin_item['Margin (%)'])}**. "
                f"Meskipun mungkin bukan penyumbang profit terbesar, COGS-nya sangat sehat."
            )

            # 4. Insight Menu Paling Rugi / Profit Terendah (Rp)
            bottom_profit_item = df_valid.nsmallest(1, "Total Profit (Rp)").iloc[0]
            insights.append(
                f"**💸 Perlu Perhatian (Rp):** Waspadai `{bottom_profit_item['Menu']}`. "
                f"Menu ini hanya menghasilkan profit **{format_rupiah(bottom_profit_item['Total Profit (Rp)'])}** "
                f"dari **{bottom_profit_item['Qty']:,.0f} porsi** terjual. (Evaluasi resep atau harga jual)."
            )

    except Exception as e:
        print(f"Gagal generate insight COGS: {e}")
        insights.append(f"Gagal membuat insight: {e}")

    return insights


def build_tab2_cogs(filtered_cogs):
    """Menggambar semua elemen untuk Tab 2.
    (VERSI BARU DENGAN KOTAK INSIGHT DI BAWAH)
    """
    if filtered_cogs is not None:
        if not filtered_cogs.empty:
            st.header("💰 Analisis Profitabilitas Menu (COGS)")

            # 1. Panggil fungsi analisis (ini sudah di-cache)
            profit_df = analyze_profit(filtered_cogs)

            # 2. Expander Ringkasan Profitabilitas (tidak berubah)
            with st.expander(
                "💰 Ringkasan Profitabilitas (Total Revenue, COGS, Profit)",
                expanded=True,
            ):
                total_revenue = profit_df["Total Revenue (Rp)"].sum()
                total_cogs_cost = profit_df["Total COGS (Rp)"].sum()
                total_profit = profit_df["Total Profit (Rp)"].sum()
                avg_margin_percent = (
                    (total_profit / total_revenue) * 100 if total_revenue > 0 else 0
                )

                st.subheader("Ringkasan Profitabilitas")
                p_col1, p_col2, p_col3, p_col4 = st.columns(4)
                p_col1.metric(
                    "📈 Total Revenue (dari COGS)", format_rupiah(total_revenue)
                )
                p_col2.metric("📉 Total COGS", format_rupiah(total_cogs_cost))
                p_col3.metric("💸 Total Profit", format_rupiah(total_profit))
                p_col4.metric(
                    "📊 Rata-rata Margin Profit", format_persen(avg_margin_percent)
                )

            st.markdown("---")

            # 3. Expander Rincian Profitabilitas (tidak berubah)
            with st.expander("📝 Rincian Profitabilitas per Menu (Tabel)"):
                st.subheader("Rincian Profitabilitas per Menu")
                st.info(
                    "Data ini dijumlahkan (agregasi) HANYA dari file Laporan COGS (sesuai filter waktu dan **cabang** yang dipilih)."
                )

                formatted_df = profit_df.copy()
                format_cols_rupiah = [
                    "Harga Jual",
                    "COGS",
                    "Margin (Rp)",
                    "Total Revenue (Rp)",
                    "Total COGS (Rp)",
                    "Total Profit (Rp)",
                ]
                for col in format_cols_rupiah:
                    formatted_df[col] = formatted_df[col].apply(format_rupiah)
                formatted_df["Margin (%)"] = formatted_df["Margin (%)"].apply(
                    format_persen
                )

                st.dataframe(formatted_df.set_index("Menu"), use_container_width=True)

            st.markdown("---")

            # 4. Expander Analisis Performa (tidak berubah)
            with st.expander(
                "📊 Analisis Performa Profit Menu (Grafik Top & Bottom 10)",
                expanded=True,
            ):
                st.subheader("Analisis Performa Profit Menu")
                top_10_profit = profit_df.nlargest(10, "Total Profit (Rp)")
                bottom_10_profit = profit_df[profit_df["Qty"] > 0].nsmallest(
                    10, "Total Profit (Rp)"
                )
                top_10_margin_pct = profit_df[profit_df["Qty"] > 0].nlargest(
                    10, "Margin (%)"
                )
                bottom_10_margin_pct = profit_df[profit_df["Qty"] > 0].nsmallest(
                    10, "Margin (%)"
                )

                p_col5, p_col6 = st.columns(2)
                with p_col5:
                    st.markdown("##### 🏆 Menu Paling Untung (by Total Profit Rp)")
                    chart = create_horizontal_bar_chart(
                        top_10_profit,
                        "Total Profit (Rp)",
                        "Menu",
                        "Total Profit (Rp)",
                        "Menu Paling Untung (by Total Profit Rp)",
                    )
                    # --- DIGANTI ---
                    st.plotly_chart(chart, use_container_width=True)
                with p_col6:
                    st.markdown("##### 📈 Menu Margin Tertinggi (by %)")
                    chart = create_horizontal_bar_chart(
                        top_10_margin_pct,
                        "Margin (%)",
                        "Menu",
                        "Margin (%)",
                        "Menu Margin Tertinggi (by %)",
                    )
                    # --- DIGANTI ---
                    st.plotly_chart(chart, use_container_width=True)

                p_col7, p_col8 = st.columns(2)
                with p_col7:
                    st.markdown(
                        "##### 📉 Menu Paling Tidak Untung (by Total Profit Rp)"
                    )
                    chart = create_horizontal_bar_chart(
                        bottom_10_profit,
                        "Total Profit (Rp)",
                        "Menu",
                        "Total Profit (Rp)",
                        "Menu Paling Tidak Untung (by Total Profit Rp)",
                        sort_order="x",
                    )
                    # --- DIGANTI ---
                    st.plotly_chart(chart, use_container_width=True)
                with p_col8:
                    st.markdown("##### 📉 Menu Margin Terendah (by %)")
                    chart = create_horizontal_bar_chart(
                        bottom_10_margin_pct,
                        "Margin (%)",
                        "Menu",
                        "Margin (%)",
                        "Menu Margin Terendah (by %)",
                        sort_order="x",
                    )
                    # --- DIGANTI ---
                    st.plotly_chart(chart, use_container_width=True)

            # #############################################################
            # --- BLOK INSIGHT BARU DITEMPATKAN DI SINI (PALING BAWAH) ---
            # #############################################################

            st.markdown("---")  # Tambahkan pemisah visual
            st.header("💡 Insight Otomatis (COGS & Profit)")

            # Panggil fungsi 'pencari insight' kita
            # Kita gunakan 'profit_df' yang sudah dihitung di awal
            insights = generate_cogs_insights(profit_df)

            # Tampilkan dalam expander baru
            with st.expander(
                "Klik untuk melihat Temuan Kunci dari Data Profit Anda", expanded=True
            ):
                if insights:
                    for insight in insights:
                        st.markdown(f"&bull; {insight}")  # Tampilkan sebagai daftar
                else:
                    st.info(
                        "Tidak ada insight otomatis yang dapat dibuat dari data ini."
                    )

            # #############################################################
            # --- BATAS BLOK INSIGHT BARU ---
            # #############################################################

        elif filtered_cogs is not None and filtered_cogs.empty:
            st.warning(
                "Tidak ada data ditemukan di File COGS untuk rentang waktu yang dipilih."
            )
    else:
        st.info(
            "Silakan upload file Laporan COGS (File 2) di sidebar untuk melihat analisis profitabilitas."
        )


# #################################################################
# --- FUNGSI BARU UNTUK INSIGHT DI TAB 3 (SDM & WAKTU) ---
# #################################################################
@st.cache_data(show_spinner=False)
def generate_hr_insights(time_data, waiter_data):
    """
    Menganalisis data SDM & Waktu yang sudah diproses dan menghasilkan insight
    dalam bahasa alami untuk Tab 3.
    """
    insights = []

    # Pastikan data tidak kosong
    if (time_data is None or time_data.empty) and (
        waiter_data is None or waiter_data.empty
    ):
        return ["Tidak ada data SDM & Waktu untuk dianalisis."]

    try:
        # 1. Insight Waktu Paling Ramai (berdasarkan Penjualan)
        if time_data is not None and not time_data.empty:
            # Urutkan berdasarkan Total_Penjualan
            peak_time_sales = time_data.nlargest(1, "Total_Penjualan").iloc[0]
            insights.append(
                f"**🕒 Waktu Emas (Penjualan):** Sesi **{peak_time_sales['Waktu Kunjungan']}** "
                f"adalah penghasil revenue terbesar, menyumbang **{format_rupiah(peak_time_sales['Total_Penjualan'])}**."
            )

            # 2. Insight Waktu Paling Sibuk (berdasarkan Transaksi)
            peak_time_trx = time_data.nlargest(1, "Jumlah_Transaksi").iloc[0]
            insights.append(
                f"**🏃 Waktu Tersibuk (Transaksi):** Sesi **{peak_time_trx['Waktu Kunjungan']}** "
                f"memiliki lalu lintas transaksi tertinggi dengan **{format_angka_bulat(peak_time_trx['Jumlah_Transaksi'])}** transaksi."
            )

    except Exception as e:
        print(f"Gagal generate insight time_data: {e}")

    try:
        # 3. Insight Waiter Terbaik (berdasarkan Penjualan)
        if waiter_data is not None and not waiter_data.empty:
            top_waiter = waiter_data.nlargest(1, "Total_Penjualan").iloc[0]
            insights.append(
                f"**🏆 Waiter Performa Terbaik:** **{top_waiter['Waiter']}** adalah top sales Anda, "
                f"menghasilkan **{format_rupiah(top_waiter['Total_Penjualan'])}** dari "
                f"**{format_angka_bulat(top_waiter['Jumlah_Transaksi'])}** transaksi."
            )

            # 4. Insight Performa Tim
            avg_sales_per_waiter = waiter_data["Total_Penjualan"].mean()
            avg_trx_per_waiter = waiter_data["Jumlah_Transaksi"].mean()
            insights.append(
                f"**🧑‍🍳 Performa Tim (Top 10):** Rata-rata, 10 waiter teratas Anda "
                f"menghasilkan **{format_rupiah(avg_sales_per_waiter)}** "
                f"dari **{avg_trx_per_waiter:,.1f}** transaksi per orang."
            )

    except Exception as e:
        print(f"Gagal generate insight waiter_data: {e}")

    return insights


# #################################################################
# --- FUNGSI BARU UNTUK INSIGHT DI TAB 6 (TARGET) - VERSI 2.0 INTERAKTIF ---
# #################################################################
@st.cache_data(show_spinner=False)
def generate_target_insights(kpi_dict):
    """
    Menganalisis kamus KPI dari Tab 6 dan menghasilkan insight
    dalam bahasa alami (DENGAN TIPE INTERAKTIF).
    """
    insights = []  # Ini akan menjadi daftar kamus: [{"type": "...", "text": "..."}]

    try:
        sisa_hari = kpi_dict.get("sisa_hari", 0)

        # --- KASUS 1: BULAN SUDAH SELESAI ---
        if sisa_hari <= 0:
            penjualan_akhir = kpi_dict.get("penjualan_saat_ini", 0)
            target_bulanan = kpi_dict.get("target_bulanan", 1)
            pencapaian_persen = (
                (penjualan_akhir / target_bulanan) if target_bulanan > 0 else 0
            )

            if pencapaian_persen >= 1:  # 1.0 = 100%
                insights.append(
                    {
                        "type": "success",
                        "text": f"**TARGET TERCAPAI:** Selamat! Bulan ini ditutup dengan pencapaian **{pencapaian_persen*100:,.1f}%** dari target.",
                    }
                )
            else:
                insights.append(
                    {
                        "type": "warning",
                        "text": f"**LAPORAN FINAL:** Bulan ini ditutup dengan pencapaian **{pencapaian_persen*100:,.1f}%** dari target.",
                    }
                )
            return insights

        # --- KASUS 2: BULAN MASIH BERJALAN ---
        proyeksi_pct = kpi_dict.get("proyeksi_vs_target_persen", 0)
        rdr_weekday = kpi_dict.get("rdr_weekday", 0)
        avg_sales_weekday = kpi_dict.get("avg_sales_weekday", 0)
        rdr_weekend = kpi_dict.get("rdr_weekend", 0)
        avg_sales_weekend = kpi_dict.get("avg_sales_weekend", 0)

        # Insight 1: Status On/Off Track
        if proyeksi_pct > 1.05:  # Di atas 105%
            insights.append(
                {
                    "type": "success",
                    "text": f"**SANGAT ON TRACK:** Performa luar biasa! Proyeksi cerdas Anda mencapai **{proyeksi_pct*100:,.1f}%** dari target. Pertahankan!",
                }
            )
        elif proyeksi_pct >= 0.98:  # Antara 98% - 105%
            insights.append(
                {
                    "type": "info",
                    "text": f"**ON TRACK:** Kerja bagus! Proyeksi cerdas Anda saat ini **{proyeksi_pct*100:,.1f}%** dari target. Jaga momentum ini.",
                }
            )
        else:  # Di bawah 98%
            insights.append(
                {
                    "type": "error",
                    "text": f"**OFF TRACK:** Perlu perhatian! Proyeksi cerdas Anda hanya **{proyeksi_pct*100:,.1f}%** dari target. Rencana aksi di bawah ini Wajib dijalankan.",
                }
            )

        # Insight 2: Rencana Aksi Weekday
        delta_weekday = rdr_weekday - avg_sales_weekday
        if delta_weekday > 0:
            insights.append(
                {
                    "type": "warning",
                    "text": f"**FOKUS WEEKDAY:** Untuk mengejar target, penjualan **Weekday (Sen-Kam)** harus ditingkatkan dari rata-rata saat ini ({format_rupiah(avg_sales_weekday)}) menjadi **{format_rupiah(rdr_weekday)}** (perlu tambahan {format_rupiah(delta_weekday)}/hari).",
                }
            )
        else:
            insights.append(
                {
                    "type": "success",
                    "text": f"**PERFORMA WEEKDAY:** Penjualan Weekday (Sen-Kam) Anda (rata-rata {format_rupiah(avg_sales_weekday)}) **SUDAH BAIK** dan di atas target harian baru ({format_rupiah(rdr_weekday)}).",
                }
            )

        # Insight 3: Rencana Aksi Weekend
        delta_weekend = rdr_weekend - avg_sales_weekend
        if delta_weekend > 0 and rdr_weekend > 0:
            insights.append(
                {
                    "type": "warning",
                    "text": f"**FOKUS WEEKEND:** Penjualan **Weekend (Jum-Min)** juga harus ditingkatkan dari rata-rata saat ini ({format_rupiah(avg_sales_weekend)}) menjadi **{format_rupiah(rdr_weekend)}** (perlu tambahan {format_rupiah(delta_weekend)}/hari).",
                }
            )

        # Insight 4: Perbandingan Model
        proyeksi_cerdas = kpi_dict.get("proyeksi_akhir_bulan", 0)
        proyeksi_prophet = kpi_dict.get("proyeksi_prophet", 0)

        if proyeksi_prophet > (proyeksi_cerdas * 1.05):  # Prophet 5% lebih tinggi
            insights.append(
                {
                    "type": "info",
                    "text": f"**CATATAN AI:** Model Prophet (**{format_rupiah(proyeksi_prophet)}**) lebih optimis daripada Proyeksi Cerdas (**{format_rupiah(proyeksi_cerdas)}**). Ini mungkin karena tren atau musiman yang positif.",
                }
            )
        elif proyeksi_prophet < (proyeksi_cerdas * 0.95):  # Prophet 5% lebih rendah
            insights.append(
                {
                    "type": "warning",
                    "text": f"**CATATAN AI:** Model Prophet (**{format_rupiah(proyeksi_prophet)}**) lebih pesimis daripada Proyeksi Cerdas (**{format_rupiah(proyeksi_cerdas)}**). Ini mungkin karena tren yang melambat.",
                }
            )

    except Exception as e:
        print(f"Gagal generate insight target: {e}")
        insights.append({"type": "error", "text": f"Gagal membuat insight target: {e}"})

    return insights


# #################################################################
# --- FUNGSI BARU UNTUK INSIGHT DI TAB 5 (FORECAST) ---
# #################################################################
@st.cache_data(show_spinner=False)
def generate_forecast_insights(forecast_df, last_date):
    """
    Menganalisis data forecast_df yang sudah diproses dan menghasilkan insight
    dalam bahasa alami untuk Tab 5.
    """
    insights = []

    if forecast_df is None or forecast_df.empty:
        return ["Tidak ada data ramalan untuk dianalisis."]

    try:
        # 1. Insight Tren Jangka Panjang
        # Ambil tren pada hari terakhir data aktual
        trend_now = forecast_df[forecast_df["ds"] <= last_date]["trend"].iloc[-1]
        # Ambil tren pada hari terakhir ramalan
        trend_future_end = forecast_df["trend"].iloc[-1]

        trend_pct = (trend_future_end - trend_now) / trend_now if trend_now != 0 else 0

        if trend_pct > 0.01:
            insights.append(
                f"**📈 Tren Jangka Panjang:** Model mendeteksi **tren NAIK** positif "
                f"({trend_pct:+.1%}) untuk periode ke depan."
            )
        elif trend_pct < -0.01:
            insights.append(
                f"**📉 Tren Jangka Panjang:** Model mendeteksi **tren TURUN** "
                f"({trend_pct:+.1%}). Waspadai potensi perlambatan bisnis."
            )
        else:
            insights.append(
                f"**⚖️ Tren Jangka Panjang:** Model mendeteksi tren penjualan yang **STABIL** "
                f"(perubahan {trend_pct:+.1%})."
            )

        # 2. Insight Ramalan 7 Hari ke Depan
        future_df = forecast_df[forecast_df["ds"] > last_date]
        if len(future_df) >= 7:
            next_7_days_sales = future_df.iloc[:7]["yhat"].sum()
            insights.append(
                f"**🔮 Ramalan 7 Hari:** Berdasarkan data Anda, model memprediksi "
                f"penjualan sekitar **{format_rupiah(next_7_days_sales)}** untuk 7 hari ke depan."
            )

        # 3. Insight Pola Mingguan
        insights.append(
            f"**🗓️ Pola Mingguan:** Untuk melihat hari apa yang paling kuat dan paling lemah "
            f"dalam seminggu, periksa grafik **'Analisis Komponen Tren & Musiman'** di atas."
        )

    except Exception as e:
        print(f"Gagal generate insight forecast: {e}")
        insights.append(f"Gagal membuat insight ramalan: {e}")

    return insights


# #################################################################
# --- FUNGSI BARU UNTUK INSIGHT DI TAB 4 (A/B COMPARISON) ---
# #################################################################
@st.cache_data(show_spinner=False)
def generate_comparison_insights(
    kpi_A_gmv, kpi_B_gmv, kpi_A_cogs, kpi_B_cogs, kpi_A_waiter, kpi_B_waiter
):
    """
    Menganalisis perbandingan KPI A vs B dan menghasilkan insight
    dalam bahasa alami untuk Tab 4.
    """
    insights = []

    # Fungsi helper kecil untuk menghitung delta persen dengan aman
    def calc_pct_delta(a, b):
        if b is None or b == 0:
            return 1.0 if (a is not None and a != 0) else 0.0
        if a is None:
            a = 0
        return (a - b) / b

    try:
        # 1. Insight Performa GMV (Revenue)
        rev_A = kpi_A_gmv.get("Total Pendapatan Kotor", 0)
        rev_B = kpi_B_gmv.get("Total Pendapatan Kotor", 0)
        pct_rev = calc_pct_delta(rev_A, rev_B)

        if pct_rev > 0.01:  # Naik
            insights.append(
                f"**📈 Performa Penjualan:** Kinerja penjualan **NAIK** signifikan sebesar **{pct_rev:+.1%}** "
                f"dibanding periode B."
            )
        elif pct_rev < -0.01:  # Turun
            insights.append(
                f"**📉 Performa Penjualan:** Kinerja penjualan **TURUN** sebesar **{pct_rev:+.1%}** "
                f"dibanding periode B. Perlu investigasi."
            )
        else:
            insights.append(
                f"**⚖️ Performa Penjualan:** Kinerja penjualan **STABIL** (perubahan {pct_rev:+.1%}) "
                f"dibanding periode B."
            )

        # 2. Insight Pendorong GMV (Transaksi vs ATV)
        trx_A = kpi_A_gmv.get("Total Transaksi", 0)
        trx_B = kpi_B_gmv.get("Total Transaksi", 0)
        pct_trx = calc_pct_delta(trx_A, trx_B)

        atv_A = kpi_A_gmv.get("Rata-rata Nilai Transaksi (ATV)", 0)
        atv_B = kpi_B_gmv.get("Rata-rata Nilai Transaksi (ATV)", 0)
        pct_atv = calc_pct_delta(atv_A, atv_B)

        if abs(pct_trx) > abs(pct_atv):
            insights.append(
                f"**💸 Pendorong Penjualan:** Perubahan penjualan utama didorong oleh **Jumlah Transaksi** "
                f"(berubah **{pct_trx:+.1%}**), sementara ATV lebih stabil."
            )
        else:
            insights.append(
                f"**💸 Pendorong Penjualan:** Perubahan penjualan utama didorong oleh **Nilai Transaksi (ATV)** "
                f"(berubah **{pct_atv:+.1%}**), sementara jumlah transaksi lebih stabil."
            )

        # 3. Insight Profitabilitas
        profit_A = kpi_A_cogs[2]  # Total Profit A
        profit_B = kpi_B_cogs[2]  # Total Profit B
        pct_profit = calc_pct_delta(profit_A, profit_B)

        margin_A = kpi_A_cogs[3]  # Margin % A
        margin_B = kpi_B_cogs[3]  # Margin % B
        delta_margin = margin_A - margin_B  # Delta absolut

        if pct_profit > 0.01:
            insights.append(
                f"**💰 Kinerja Profit:** Kabar baik! **Total Profit NAIK** sebesar **{pct_profit:+.1%}**. "
                f"Efisiensi margin juga (membaik/memburuk) sebesar **{delta_margin:+.1f} poin**."
            )
        elif pct_profit < -0.01:
            insights.append(
                f"**⚠️ Kinerja Profit:** Hati-hati! **Total Profit TURUN** sebesar **{pct_profit:+.1%}**. "
                f"Efisiensi margin juga (membaik/memburuk) sebesar **{delta_margin:+.1f} poin**."
            )

    except Exception as e:
        print(f"Gagal generate insight comparison: {e}")
        insights.append(f"Gagal membuat perbandingan insight: {e}")

    return insights


# #################################################################
# --- FUNGSI BARU UNTUK INSIGHT DI TAB 7 (ULASAN) ---
# #################################################################
@st.cache_data(show_spinner=False)
def generate_review_insights(
    total_ulasan, avg_rating, nps_score, df_positive_topics, df_negative_topics
):
    """
    Menganalisis data ulasan yang sudah diproses dan menghasilkan insight
    dalam bahasa alami untuk Tab 7.
    """
    insights = []

    if total_ulasan == 0:
        return ["Belum ada data ulasan untuk dianalisis."]

    try:
        # 1. Insight Sentimen Umum
        insight_nps = "NETRAL"
        if nps_score > 20:
            insight_nps = "BAIK"
        elif nps_score < 0:
            insight_nps = "PERLU PERHATIAN"

        insights.append(
            f"**❤️ Sentimen Umum:** Anda menerima **{total_ulasan} ulasan** dengan rata-rata rating **{avg_rating:.1f} dari 5**. "
            f"Skor NPS Anda adalah **{nps_score:.1f}**, yang tergolong **{insight_nps}**."
        )

        # 2. Insight Kekuatan Terbesar (Top Positive)
        if not df_positive_topics.empty:
            top_positive = df_positive_topics.iloc[0]
            insights.append(
                f"**👍 Kekuatan Terbesar:** Pelanggan paling sering memuji tentang **{top_positive['Topik']}** "
                f"(disebut **{top_positive['Jumlah']} kali**). Pertahankan ini!"
            )

        # 3. Insight Keluhan Utama (Top Negative)
        if not df_negative_topics.empty:
            top_negative = df_negative_topics.iloc[0]
            insights.append(
                f"**👎 Keluhan Utama:** Area perbaikan paling mendesak adalah **{top_negative['Topik']}** "
                f"(disebut **{top_negative['Jumlah']} kali**). Ini adalah prioritas Anda."
            )

        # 4. Insight Rating Bintang 1
        if "Rating 1 (Sangat Buruk)" in df_negative_topics["Topik"].values:
            count_bintang_1 = df_negative_topics[
                df_negative_topics["Topik"] == "Rating 1 (Sangat Buruk)"
            ].iloc[0]["Jumlah"]
            if count_bintang_1 > 0:
                insights.append(
                    f"**🚨 Peringatan Detractor:** Ada **{count_bintang_1} ulasan bintang 1** "
                    f"yang perlu segera ditindaklanjuti."
                )

    except Exception as e:
        print(f"Gagal generate insight ulasan: {e}")
        insights.append(f"Gagal membuat insight ulasan: {e}")

    return insights


# #################################################################
# --- FUNGSI BARU UNTUK INSIGHT DI TAB 8 (PEMBELIAN) ---
# #################################################################
@st.cache_data(show_spinner=False)
def generate_purchase_insights(
    total_cost, cost_by_category, cost_by_supplier, top_items, fcp
):
    """
    Menganalisis data pembelian yang sudah diproses dan menghasilkan insight
    dalam bahasa alami untuk Tab 8.
    """
    insights = []

    # Tambahkan logika insight FCP
    if fcp > 0:
        insights.append(
            f"💰 **Food Cost Percentage (FCP):** Angka FCP saat ini adalah **{fcp:.1f}%**. Ini adalah metrik krusial yang perlu dipantau terhadap target ideal (biasanya 25%-35%)."
        )

    # ... (lanjutan logika insight pembelian lainnya) ...

    # Contoh: Insight top cost
    if not cost_by_category.empty:
        top_cat = cost_by_category.iloc[0]["Category"]
        insights.append(
            f"🛒 **Biaya Terbesar:** Kategori '{top_cat}' menyumbang biaya terbesar, memerlukan perhatian khusus untuk negosiasi atau substitusi."
        )

    if total_cost == 0 or top_items.empty:
        return ["Belum ada data pembelian untuk dianalisis."]

    try:
        # 1. Insight Total Biaya
        insights.append(
            f"**🛒 Total Biaya:** Total biaya pembelian Anda (yang tercatat > Rp 0) "
            f"pada periode ini adalah **{format_rupiah(total_cost)}**."
        )

        # 2. Insight Kategori Biaya Terbesar
        if not cost_by_category.empty:
            top_cat = cost_by_category.iloc[0]
            insights.append(
                f"**🍔 Kategori Terbesar:** Kategori pengeluaran terbesar Anda adalah **`{top_cat['Category']}`**, "
                f"menghabiskan **{format_rupiah(top_cat['Total'])}**."
            )

        # 3. Insight Supplier Terbesar
        if not cost_by_supplier.empty:
            top_supp = cost_by_supplier.iloc[0]
            insights.append(
                f"**🚚 Supplier Utama:** Supplier dengan pembelian terbesar adalah **`{top_supp['Supplier Name']}`**, "
                f"dengan total pembelian **{format_rupiah(top_supp['Total'])}**."
            )

        # 4. Insight Item Termahal
        if not top_items.empty:
            top_item = top_items.iloc[0]
            insights.append(
                f"**💸 Item Termahal:** Item tunggal yang paling banyak memakan biaya adalah **`{top_item['Product Name']}`**, "
                f"dengan total **{format_rupiah(top_item['Total'])}**. "
                f"Selalu periksa harga beli dan penggunaan (waste) item ini."
            )

    except Exception as e:
        print(f"Gagal generate insight pembelian: {e}")
        insights.append(f"Gagal membuat insight pembelian: {e}")

    return insights


# #################################################################
# --- FUNGSI BARU UNTUK INSIGHT DI TAB 9 (REKOMENDASI) ---
# #################################################################
@st.cache_data(show_spinner=False)
def generate_recommendation_insights(rules_df):
    """
    Menganalisis data 'rules_df' yang sudah diproses dan menghasilkan insight
    dalam bahasa alami untuk Tab 9.
    """
    insights = []

    # Pastikan data tidak kosong
    if rules_df is None or rules_df.empty:
        return ["Tidak ada aturan rekomendasi yang cukup kuat untuk dianalisis."]

    try:
        # 1. Insight 1: Potensi Sales Tertinggi (Expected Value)
        # Kita asumsikan rules_df sudah di-sort by expected_value, tapi kita sort lagi
        top_ev_rule = rules_df.nlargest(1, "expected_value").iloc[0]
        item_A_ev = top_ev_rule["antecedents"]
        item_B_ev = top_ev_rule["consequents"]
        ev = top_ev_rule["expected_value"]

        insights.append(
            f"**💸 Potensi Sales Tertinggi:** Aturan paling bernilai adalah **JIKA BELI `{item_A_ev}`, TAWARKAN `{item_B_ev}`**. "
            f"Setiap tawaran ini memiliki *expected value* (potensi sales) sebesar **{format_rupiah(ev)}**."
        )

        # 2. Insight 2: Pasangan Paling Setia (Confidence)
        top_conf_rule = rules_df.nlargest(1, "confidence").iloc[0]
        item_A_conf = top_conf_rule["antecedents"]
        item_B_conf = top_conf_rule["consequents"]
        conf = top_conf_rule["confidence"]

        insights.append(
            f"**🤝 Pasangan Paling Setia:** Aturan paling *pasti* (confidence tertinggi) adalah **JIKA BELI `{item_A_conf}`, TAWARKAN `{item_B_conf}`**. "
            f"**{conf:.1%}** pelanggan yang membeli `{item_A_conf}` juga membeli `{item_B_conf}`."
        )

        # 3. Insight 3: Koneksi Terkuat (Lift)
        top_lift_rule = rules_df.nlargest(1, "lift").iloc[0]
        item_A_lift = top_lift_rule["antecedents"]
        item_B_lift = top_lift_rule["consequents"]
        lift = top_lift_rule["lift"]

        insights.append(
            f"**🔗 Koneksi Terkuat (Lift):** Pasangan **`{item_A_lift}`** dan **`{item_B_lift}`** adalah yang paling unik. "
            f"Pelanggan **{lift:.1f}x lebih mungkin** membeli keduanya bersamaan daripada secara acak. Ini adalah peluang *cross-sell* yang kuat."
        )

        # 4. Insight 4: Aksi
        insights.append(
            "**💡 Aksi:** Gunakan filter **'JIKA Pelanggan Beli Menu Ini'** di atas untuk "
            "mengeksplorasi rekomendasi spesifik untuk item terlaris Anda."
        )

    except Exception as e:
        print(f"Gagal generate insight rekomendasi: {e}")
        insights.append(f"Gagal membuat insight: {e}")

    return insights


# ==============================================================================
#                      FUNGSI ANALISIS KECURANGAN
# ==============================================================================
@st.cache_data
def generate_pl_insights(df_pl):
    """Insight otomatis untuk P&L."""
    insights = []
    if df_pl is None or df_pl.empty:
        return ["Data P&L kosong."]

    # Filter Tahun Berjalan Saja untuk default
    df_curr = df_pl[df_pl["Year_Type"] == "Current Year"]

    total_rev = df_curr[df_curr["Category"] == "Revenue"]["Value"].sum()
    total_exp = df_curr[df_curr["Category"] == "Expense"]["Value"].sum()
    total_cogs = df_curr[df_curr["Category"] == "COGS"]["Value"].sum()
    net_profit = total_rev - total_cogs - total_exp

    # 1. Net Profit Margin
    npm = (net_profit / total_rev * 100) if total_rev else 0
    if npm > 15:
        insights.append(
            f"**💰 Profitabilitas Sehat:** Net Profit Margin Anda mencapai **{npm:.1f}%**, angka yang sangat sehat untuk F&B."
        )
    elif npm > 0:
        insights.append(
            f"**⚠️ Profit Tipis:** Bisnis mencetak laba (**{npm:.1f}% margin**), namun perlu efisiensi biaya."
        )
    else:
        insights.append(
            f"**🚨 Merugi:** Periode ini mencatat kerugian bersih sebesar **{format_rupiah(abs(net_profit))}**."
        )

    # 2. Expense Ratio
    exp_ratio = (total_exp / total_rev * 100) if total_rev else 0
    insights.append(
        f"**💸 Beban Operasional:** Biaya operasional memakan **{exp_ratio:.1f}%** dari omzet."
    )

    return insights


def get_fraud_analysis(df):
    """
    Menganalisa kecurangan dengan deteksi anomali berbasis Deviasi Standar.
    PERBAIKAN: Menggunakan 'Sales Type' sebagai kolom jenis transaksi.
    """
    col_waiter = "Waiter"
    # PERBAIKAN KEY ERROR: Menggunakan nama kolom yang konsisten ('Sales Type')
    col_type = "Sales Type"

    # --- 1. DATA CLEANING ---
    df_clean = df.copy()

    # Guardrail: Pastikan kolom ada sebelum operasi
    if col_type not in df_clean.columns or col_waiter not in df_clean.columns:
        # Mengembalikan pesan error yang spesifik
        return (
            None,
            f"Kolom tidak ditemukan. Pastikan ada '{col_waiter}' dan '{col_type}'.",
        )

    # Gabungkan semua variasi void menjadi satu nama 'Void'
    df_clean[col_type] = df_clean[col_type].replace(
        {"Void Sales": "Void", "Void sales": "Void", "VOID": "Void"}
    )

    cat_void = "Void"
    cat_nonsales = "Non Sales"
    # ------------------------

    # --- 2. BUAT TABEL PIVOT (DATA MENTAH) ---
    fraud_pivot = pd.crosstab(df_clean[col_waiter], df_clean[col_type])

    # --- 3. PENGAMANAN & INISIALISASI KOLOM ---
    if "Void Sales" in fraud_pivot.columns:
        fraud_pivot = fraud_pivot.drop(columns=["Void Sales"])

    if cat_void not in fraud_pivot.columns:
        fraud_pivot[cat_void] = 0
    if cat_nonsales not in fraud_pivot.columns:
        fraud_pivot[cat_nonsales] = 0

    # --- 4. HITUNG STATISTIK (DETEKSI ANOMALI: Mean + 2 * Std Dev) ---

    # 4.1. Analisa Void
    void_counts = fraud_pivot[cat_void]
    avg_void = void_counts.mean()
    std_void = void_counts.std()
    threshold_void = np.ceil(max(2.0, avg_void + 2 * std_void))

    suspects_void = fraud_pivot[void_counts > threshold_void].sort_values(
        by=cat_void, ascending=False
    )

    # 4.2. Analisa Non Sales
    nonsales_counts = fraud_pivot[cat_nonsales]
    avg_nonsales = nonsales_counts.mean()
    std_nonsales = nonsales_counts.std()
    threshold_nonsales = np.ceil(max(3.0, avg_nonsales + 2 * std_nonsales))

    suspects_nonsales = fraud_pivot[nonsales_counts > threshold_nonsales].sort_values(
        by=cat_nonsales, ascending=False
    )

    suspects_void_out = suspects_void[[cat_void]]
    suspects_nonsales_out = suspects_nonsales[[cat_nonsales]]

    return {
        "raw_data": fraud_pivot,
        "void": {
            "suspects": suspects_void_out,
            "avg": avg_void,
            "threshold": threshold_void,
            "std": std_void,
        },
        "nonsales": {
            "suspects": suspects_nonsales_out,
            "avg": avg_nonsales,
            "threshold": threshold_nonsales,
            "std": std_nonsales,
        },
    }, "Success"


def generate_fraud_insights(fraud_result):
    """Menghasilkan insight otomatis dari hasil analisis fraud."""
    insights = []

    if not fraud_result:
        return ["Analisis deteksi anomali gagal dijalankan."]

    res_void = fraud_result["void"]
    res_nonsales = fraud_result["nonsales"]

    # 1. Insight Void
    if not res_void["suspects"].empty:
        count = len(res_void["suspects"])
        names = ", ".join(res_void["suspects"].index.tolist()[:3])
        threshold = res_void["threshold"]
        insights.append(
            f"⚠️ **{count} Karyawan (cth: {names}) terindikasi anomali Void.** Jumlah Void mereka melebihi ambang batas {threshold:.0f} kali (Rata-rata Toko + 2 Deviasi Standar)."
        )
    else:
        insights.append(
            "✅ **Laporan Void:** Tidak ada anomali Void yang terdeteksi secara signifikan."
        )

    # 2. Insight Non Sales
    if not res_nonsales["suspects"].empty:
        count = len(res_nonsales["suspects"])
        names = ", ".join(res_nonsales["suspects"].index.tolist()[:3])
        threshold = res_nonsales["threshold"]
        insights.append(
            f"🟠 **{count} Karyawan (cth: {names}) terindikasi anomali Non-Sales.** Aktivitas Non-Sales mereka melebihi ambang batas {threshold:.0f} kali."
        )
    else:
        insights.append(
            "✅ **Laporan Non-Sales:** Tidak ada anomali Non-Sales yang terdeteksi secara signifikan."
        )

    return insights


def get_void_details(df):
    """
    Mengambil rincian baris data khusus transaksi VOID.
    PERBAIKAN: Filter diperluas untuk mencakup semua jenis 'Void'.
    """
    # Filter semua variasi Void
    df_void_details = df[
        df["Sales Type"].isin(["Void", "Void Sales", "Void sales", "VOID"])
    ].copy()

    if df_void_details.empty:
        return None

    # Daftar kemungkinan nama kolom untuk "Alasan"
    potential_reason_cols = [
        "Reason",
        "Void Reason",
        "Remark",
        "Notes",
        "Keterangan",
        "Alasan",
    ]

    found_reason_col = "Tidak Ditemukan"
    for col in potential_reason_cols:
        if col in df.columns:
            found_reason_col = col
            break

    # Pilih kolom yang mau ditampilkan
    target_cols = ["Sales Date In", "Time", "Waiter", "Total", "Net Sales"]

    if found_reason_col != "Tidak Ditemukan":
        target_cols.append(found_reason_col)

    final_cols = [c for c in target_cols if c in df.columns]

    return df_void_details[final_cols]


# ==============================================================================
#                      FUNGSI TAMPILAN STREAMLIT (HR TAB)
# ==============================================================================


def build_tab3_hr(filtered_waiter):
    """Menggambar semua elemen untuk Tab 3."""

    if filtered_waiter is None or filtered_waiter.empty:
        if filtered_waiter is not None and filtered_waiter.empty:
            st.warning(
                "Tidak ada data ditemukan di File Rekapitulasi untuk rentang waktu yang dipilih."
            )
        else:
            st.info(
                "Silakan upload file Laporan Rekapitulasi Detail (File 3) di sidebar untuk melihat analisis waiter."
            )
        return

    st.header("🧑‍🍳 Analisis Kinerja Waiter & Waktu Kunjungan")

    time_data = get_peak_time_analysis(filtered_waiter)
    waiter_data = get_waiter_performance(filtered_waiter)

    # --- BAGIAN 1: WAKTU KUNJUNGAN ---
    with st.expander("🕒 Analisis Waktu Kunjungan Pelanggan", expanded=True):
        st.subheader("🕒 Waktu Kunjungan Pelanggan")
        sort_order_time = [
            "Breakfast/Brunch (10-12)",
            "Lunch (12-17)",
            "Dinner (17-22)",
            "Luar Jam Buka",
        ]

        t_col1, t_col2 = st.columns(2)
        with t_col1:
            st.markdown("##### Berdasarkan Jumlah Transaksi")
            chart1 = create_vertical_bar_chart(
                time_data,
                "Waktu Kunjungan",
                "Jumlah_Transaksi",
                "Waktu Kunjungan",
                "Jumlah Transaksi",
                "O",
                sort_order=sort_order_time,
            )
            st.plotly_chart(chart1, use_container_width=True)
        with t_col2:
            st.markdown("##### Berdasarkan Total Penjualan")
            chart2 = create_vertical_bar_chart(
                time_data,
                "Waktu Kunjungan",
                "Total_Penjualan",
                "Waktu Kunjungan",
                "Total Penjualan (Rp)",
                "O",
                sort_order=sort_order_time,
            )
            st.plotly_chart(chart2, use_container_width=True)

        st.dataframe(
            time_data.set_index("Waktu Kunjungan").style.format(
                {
                    "Total_Penjualan": format_rupiah,
                    "Jumlah_Transaksi": format_angka_bulat,
                }
            ),
            use_container_width=True,
        )

    st.markdown("---")

    # --- BAGIAN 2: PERFORMA WAITER ---
    with st.expander("🏆 Performa Waiter Teratas (Top 10)", expanded=True):
        st.subheader("🏆 Performa Waiter Teratas (Top 10)")
        chart_waiter = create_horizontal_bar_chart(
            waiter_data,
            "Total_Penjualan",
            "Waiter",
            "Total Penjualan (Rp)",
            "Performa Waiter Teratas (by Penjualan)",
        )
        st.plotly_chart(chart_waiter, use_container_width=True)
        st.dataframe(
            waiter_data.set_index("Waiter").style.format(
                {
                    "Total_Penjualan": format_rupiah,
                    "Jumlah_Transaksi": format_angka_bulat,
                }
            ),
            use_container_width=True,
        )

    st.markdown("---")

    # --- BAGIAN 3: DETEKSI KECURANGAN ---
    st.header("🕵️ Deteksi Anomali & Potensi Kecurangan")
    st.caption(
        "Menganalisa pola transaksi 'Void Sales' dan 'Non Sales' yang tidak wajar."
    )

    fraud_result, status_msg = get_fraud_analysis(filtered_waiter)

    if fraud_result:
        tab_void, tab_nonsales, tab_data = st.tabs(
            ["🔴 Analisa Void", "🟠 Analisa Non-Sales", "📋 Data Mentah"]
        )

        # 1. ANALISA VOID
        with tab_void:
            res_void = fraud_result["void"]

            df_all_voids = filtered_waiter[
                filtered_waiter["Sales Type"].isin(
                    ["Void", "Void Sales", "Void sales", "VOID"]
                )
            ].copy()

            if not df_all_voids.empty:
                col_uang = (
                    "Net Sales" if "Net Sales" in df_all_voids.columns else "Total"
                )

                c1, c2, c3 = st.columns(3)
                c1.metric("Total Kejadian Void", f"{len(df_all_voids)} kali")

                if col_uang in df_all_voids.columns:
                    total_uang = df_all_voids[col_uang].sum()
                    c2.metric("Total Nilai Void", format_rupiah(total_uang))
                else:
                    c2.metric("Total Nilai Void", "Rp 0")

                c3.metric("Rata-rata Void Toko", f"{res_void['avg']:.1f} kali/waiter")

                st.markdown("---")

                st.subheader("📋 Daftar Transaksi Void")

                desired_cols = [
                    "Sales Date In",
                    "Time",
                    "Waiter",
                    "Qty",
                    "Net Sales",
                    "Total",
                    "Table",
                    "Section",
                ]

                final_cols = [c for c in desired_cols if c in df_all_voids.columns]
                df_display = df_all_voids[final_cols].copy()

                sort_columns = []
                if "Sales Date In" in df_display.columns:
                    sort_columns.append("Sales Date In")
                if "Time" in df_display.columns:
                    sort_columns.append("Time")

                if sort_columns:
                    df_display = df_display.sort_values(
                        by=sort_columns, ascending=False
                    )

                format_dict = {}
                col_uang_untuk_total = None  # Ditambahkan untuk Grand Total

                if "Net Sales" in final_cols:
                    format_dict["Net Sales"] = format_rupiah
                    col_uang_untuk_total = "Net Sales"
                elif "Total" in final_cols:
                    format_dict["Total"] = format_rupiah
                    col_uang_untuk_total = "Total"
                if "Qty" in final_cols:
                    format_dict["Qty"] = "{:.0f}"

                st.dataframe(
                    df_display.style.format(format_dict),
                    use_container_width=True,
                )

                csv = df_all_voids[final_cols].to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="📥 Download Data Void (CSV)",
                    data=csv,
                    file_name="laporan_void.csv",
                    mime="text/csv",
                )

            else:
                st.success("✅ Bersih! Tidak ada transaksi Void sama sekali.")

            st.markdown("---")

            if not res_void["suspects"].empty:
                st.error(
                    f"⚠️ HIGHLIGHT: {len(res_void['suspects'])} Karyawan dengan Void terbanyak (Threshold: {res_void['threshold']:.0f}x):"
                )
                st.table(res_void["suspects"].style.highlight_max(axis=0, color="pink"))

        # 2. ANALISA NON-SALES
        with tab_nonsales:
            res_nonsales = fraud_result["nonsales"]
            if not res_nonsales["suspects"].empty:
                st.error(
                    f"⚠️ HIGHLIGHT: {len(res_nonsales['suspects'])} Karyawan dengan Non-Sales terbanyak (Threshold: {res_nonsales['threshold']:.0f}x):"
                )
                st.table(
                    res_nonsales["suspects"].style.highlight_max(axis=0, color="orange")
                )
                st.caption(
                    f"Ambang Batas (Threshold): Lebih dari {res_nonsales['threshold']:.0f} kali."
                )
            else:
                st.success("✅ Bersih! Tidak ada anomali Non-Sales yang terdeteksi.")
                st.caption(
                    f"Rata-rata Non-Sales: {res_nonsales['avg']:.1f} kali/waiter."
                )

        # 3. DATA MENTAH
        with tab_data:
            st.dataframe(fraud_result["raw_data"], use_container_width=True)

    else:
        st.warning(f"Gagal menjalankan analisa: {status_msg}")

    st.markdown("---")

    # ============================================================
    # BAGIAN 4: INSIGHT OTOMATIS (Diperbarui)
    # ============================================================
    st.header("💡 Insight Otomatis (SDM, Waktu & Anomali)")  # Judul Diperbarui

    # Panggil fungsi 'pencari insight' kita
    insights_hr = generate_hr_insights(time_data, waiter_data)

    # Tambahkan Fraud Insight
    if fraud_result:
        insights_fraud = generate_fraud_insights(fraud_result)
        # Gabungkan semua insight
        all_insights = insights_hr + insights_fraud
    else:
        all_insights = insights_hr + [f"Gagal menghasilkan insight fraud: {status_msg}"]

    with st.expander(
        "Klik untuk melihat Temuan Kunci dari Data SDM, Waktu & Anomali", expanded=True
    ):
        if all_insights:
            for insight in all_insights:
                st.markdown(f"&bull; {insight}")
        else:
            st.info("Tidak ada insight otomatis yang dapat dibuat dari data ini.")


# Tab4


def build_tab4_comparison(data_gmv, data_cogs, data_waiter):
    """Menggambar Tab 4 (Perbandingan A/B) dengan tata letak A | Delta | B.
    (VERSI BARU DENGAN KOTAK INSIGHT DI BAWAH)
    """

    st.header("⚖️ Analisis Perbandingan Periodik (A vs B)")
    st.info(
        "Gunakan tab ini untuk membandingkan kinerja antara dua periode (A vs B). "
        "Filter global diabaikan di tab ini, namun **rentang waktu** ditentukan dari sini."
    )

    # --- Helper Lokal (Hanya untuk Tab 4) ---
    def slice_data_by_period(df, date_col_name, ref_date, comparison_type):
        if df is None or df.empty:
            return (
                pd.DataFrame(columns=df.columns if df is not None else None),
                "Tidak ada data",
            )
        caption_text = ""
        sliced_df = pd.DataFrame(columns=df.columns)
        if comparison_type == "Harian":
            caption_text = f"Periode: {ref_date.strftime('%d-%m-%Y')}"
            sliced_df = df[df[date_col_name].dt.date == ref_date]
        elif comparison_type == "Mingguan":
            start_date = ref_date
            end_date = start_date + pd.to_timedelta(6, unit="d")
            caption_text = f"Periode: {start_date.strftime('%d-%m-%Y')} s.d. {end_date.strftime('%d-%m-%Y')}"
            sliced_df = df[
                (df[date_col_name].dt.date >= start_date)
                & (df[date_col_name].dt.date <= end_date)
            ]
        elif comparison_type == "Bulanan":
            ref_month, ref_year = ref_date.month, ref_date.year
            caption_text = f"Periode: {ref_date.strftime('%B %Y')}"
            sliced_df = df[
                (df[date_col_name].dt.month == ref_month)
                & (df[date_col_name].dt.year == ref_year)
            ]
        elif comparison_type == "Tahunan":
            ref_year = ref_date.year
            caption_text = f"Periode: Tahun {ref_year}"
            sliced_df = df[df[date_col_name].dt.year == ref_year]
        return sliced_df, caption_text

    def get_profit_kpis(df_cogs_sliced):
        # Kita panggil fungsi analyze_profit yang sudah di-cache dan difilter
        profit_df = analyze_profit(df_cogs_sliced)
        if profit_df.empty:
            return 0, 0, 0, 0
        total_revenue = profit_df["Total Revenue (Rp)"].sum()
        total_cogs = profit_df["Total COGS (Rp)"].sum()
        total_profit = profit_df["Total Profit (Rp)"].sum()
        margin = (total_profit / total_revenue) * 100 if total_revenue > 0 else 0
        return total_revenue, total_cogs, total_profit, margin

    def get_waiter_kpis(df_waiter_sliced):
        if df_waiter_sliced is None or df_waiter_sliced.empty:
            return 0, 0, 0
        bill_df = (
            df_waiter_sliced.groupby("Bill Number")
            .agg(Total_Sales=("Total After Bill Discount", "sum"))
            .reset_index()
        )
        total_penjualan = bill_df["Total_Sales"].sum()
        total_transaksi = bill_df["Bill Number"].nunique()
        cleaned_waiters = df_waiter_sliced["Waiter"].fillna("Tidak Diketahui")
        unique_waiters_count = cleaned_waiters[
            cleaned_waiters != "Tidak Diketahui"
        ].nunique()
        avg_sales_per_waiter = (
            total_penjualan / unique_waiters_count if unique_waiters_count > 0 else 0
        )
        return total_penjualan, total_transaksi, avg_sales_per_waiter

    # --- Akhir Helper Lokal ---

    master_min_date = pd.Timestamp.now().date()
    master_max_date = pd.Timestamp.now().date()
    has_data = False

    try:
        if data_gmv is not None and not data_gmv.empty:
            master_min_date = data_gmv["Sales Date In"].min().date()
            master_max_date = data_gmv["Sales Date In"].max().date()
            has_data = True
        elif data_cogs is not None and not data_cogs.empty:
            master_min_date = data_cogs["Sales Date"].min().date()
            master_max_date = data_cogs["Sales Date"].max().date()
            has_data = True
        elif data_waiter is not None and not data_waiter.empty:
            master_min_date = data_waiter["Order Time"].min().date()
            master_max_date = data_waiter["Order Time"].max().date()
            has_data = True
    except Exception:
        pass

    if not has_data:
        st.warning(
            "Upload setidaknya satu file data untuk memulai analisis perbandingan."
        )
        return

    comparison_type = st.selectbox(
        "Pilih Tipe Perbandingan:", ["Harian", "Mingguan", "Bulanan", "Tahunan"]
    )
    st.markdown("---")

    filter_col_A, filter_col_Delta, filter_col_B = st.columns([0.35, 0.3, 0.35])

    default_A_date = master_max_date
    default_B_date = master_min_date
    try:
        if comparison_type == "Harian":
            default_B_date = default_A_date - pd.to_timedelta(1, unit="d")
        elif comparison_type == "Mingguan":
            default_B_date = default_A_date - pd.to_timedelta(7, unit="d")
        elif comparison_type == "Bulanan":
            default_B_date = default_A_date - pd.DateOffset(months=1)
            # Konversi Timestamp ke date
            default_B_date = default_B_date.date()
        elif comparison_type == "Tahunan":
            default_B_date = default_A_date - pd.DateOffset(years=1)
            # Konversi Timestamp ke date
            default_B_date = default_B_date.date()

        if pd.Timestamp(default_B_date) < pd.Timestamp(master_min_date):
            default_B_date = master_min_date
        elif isinstance(default_B_date, pd.Timestamp):
            default_B_date = default_B_date.date()
    except Exception:
        default_B_date = master_min_date

    with filter_col_A:
        st.subheader("Periode A (Saat Ini)")
        date_A = st.date_input(
            f"Tanggal Acuan (A)",
            value=default_A_date,
            min_value=master_min_date,
            max_value=master_max_date,
            key="comp_date_A",
            label_visibility="collapsed",
        )

    with filter_col_Delta:
        st.subheader("Perubahan")
        st.caption("A vs B")

    with filter_col_B:
        st.subheader("Periode B (Pembanding)")
        date_B = st.date_input(
            f"Tanggal Acuan (B)",
            value=default_B_date,
            min_value=master_min_date,
            max_value=master_max_date,
            key="comp_date_B",
            label_visibility="collapsed",
        )

    gmv_A, caption_A_gmv = slice_data_by_period(
        data_gmv, "Sales Date In", date_A, comparison_type
    )
    gmv_B, caption_B_gmv = slice_data_by_period(
        data_gmv, "Sales Date In", date_B, comparison_type
    )
    cogs_A, caption_A_cogs = slice_data_by_period(
        data_cogs, "Sales Date", date_A, comparison_type
    )
    cogs_B, caption_B_cogs = slice_data_by_period(
        data_cogs, "Sales Date", date_B, comparison_type
    )
    waiter_A, caption_A_waiter = slice_data_by_period(
        data_waiter, "Order Time", date_A, comparison_type
    )
    waiter_B, caption_B_waiter = slice_data_by_period(
        data_waiter, "Order Time", date_B, comparison_type
    )

    with filter_col_A:
        st.caption(
            caption_A_gmv
            if data_gmv is not None
            else (caption_A_cogs if data_cogs is not None else caption_A_waiter)
        )
    with filter_col_B:
        st.caption(
            caption_B_gmv
            if data_gmv is not None
            else (caption_B_cogs if data_cogs is not None else caption_B_waiter)
        )

    st.markdown("---")

    # Hitung semua KPI
    kpi_A_gmv = (
        calculate_sales_kpi(gmv_A)
        if data_gmv is not None
        else calculate_sales_kpi(None)
    )
    kpi_B_gmv = (
        calculate_sales_kpi(gmv_B)
        if data_gmv is not None
        else calculate_sales_kpi(None)
    )
    kpi_A_cogs = get_profit_kpis(cogs_A) if data_cogs is not None else (0, 0, 0, 0)
    kpi_B_cogs = get_profit_kpis(cogs_B) if data_cogs is not None else (0, 0, 0, 0)
    kpi_A_waiter = get_waiter_kpis(waiter_A) if data_waiter is not None else (0, 0, 0)
    kpi_B_waiter = get_waiter_kpis(waiter_B) if data_waiter is not None else (0, 0, 0)

    # Tampilkan UI Metrik (tidak berubah)
    if data_gmv is not None:
        st.markdown("##### 📊 Kinerja Penjualan (dari File 1: GMV)")
        with st.container(border=True):
            # ... (semua kode metrik A | Delta | B Anda) ...
            col_A, col_Delta, col_B = st.columns([0.35, 0.3, 0.35])
            # (Revenue)
            with col_A:
                st.metric(
                    "Total Pendapatan Kotor",
                    format_rupiah(kpi_A_gmv["Total Pendapatan Kotor"]),
                )
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_gmv["Total Pendapatan Kotor"],
                    kpi_B_gmv["Total Pendapatan Kotor"],
                    format_rupiah,
                    True,
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric(
                    "Total Pendapatan Kotor",
                    format_rupiah(kpi_B_gmv["Total Pendapatan Kotor"]),
                )
            # (Transaksi)
            with col_A:
                st.metric(
                    "Total Transaksi", format_angka_bulat(kpi_A_gmv["Total Transaksi"])
                )
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_gmv["Total Transaksi"],
                    kpi_B_gmv["Total Transaksi"],
                    format_angka_bulat,
                    True,
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric(
                    "Total Transaksi", format_angka_bulat(kpi_B_gmv["Total Transaksi"])
                )
            # (ATV)
            with col_A:
                st.metric(
                    "Rata-rata (ATV)",
                    format_rupiah(kpi_A_gmv["Rata-rata Nilai Transaksi (ATV)"]),
                )
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_gmv["Rata-rata Nilai Transaksi (ATV)"],
                    kpi_B_gmv["Rata-rata Nilai Transaksi (ATV)"],
                    format_rupiah,
                    True,
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric(
                    "Rata-rata (ATV)",
                    format_rupiah(kpi_B_gmv["Rata-rata Nilai Transaksi (ATV)"]),
                )
            # (IPB)
            with col_A:
                st.metric(
                    "Item per Transaksi (IPB)",
                    f"{kpi_A_gmv['Item per Transaksi (IPB)']:.2f}",
                )
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_gmv["Item per Transaksi (IPB)"],
                    kpi_B_gmv["Item per Transaksi (IPB)"],
                    format_persen,
                    True,
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric(
                    "Item per Transaksi (IPB)",
                    f"{kpi_B_gmv['Item per Transaksi (IPB)']:.2f}",
                )
            # (Diskon)
            with col_A:
                st.metric("Total Diskon", format_rupiah(kpi_A_gmv["Total Diskon"]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_gmv["Total Diskon"],
                    kpi_B_gmv["Total Diskon"],
                    format_rupiah,
                    False,
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total Diskon", format_rupiah(kpi_B_gmv["Total Diskon"]))
        st.markdown("---")

    if data_cogs is not None:
        st.markdown("##### 💰 Kinerja Profitabilitas (dari File 2: COGS)")
        with st.container(border=True):
            # ... (semua kode metrik COGS A | Delta | B Anda) ...
            col_A, col_Delta, col_B = st.columns([0.35, 0.3, 0.35])
            # (Revenue COGS)
            with col_A:
                st.metric("Total Revenue (COGS)", format_rupiah(kpi_A_cogs[0]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_cogs[0], kpi_B_cogs[0], format_rupiah, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total Revenue (COGS)", format_rupiah(kpi_B_cogs[0]))
            # (Total COGS)
            with col_A:
                st.metric("Total COGS", format_rupiah(kpi_A_cogs[1]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_cogs[1], kpi_B_cogs[1], format_rupiah, False
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total COGS", format_rupiah(kpi_B_cogs[1]))
            # (Total Profit)
            with col_A:
                st.metric("Total Profit", format_rupiah(kpi_A_cogs[2]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_cogs[2], kpi_B_cogs[2], format_rupiah, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total Profit", format_rupiah(kpi_B_cogs[2]))
            # (Margin Profit)
            with col_A:
                st.metric("Margin Profit (%)", format_persen(kpi_A_cogs[3]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_cogs[3], kpi_B_cogs[3], format_persen, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Margin Profit (%)", format_persen(kpi_B_cogs[3]))
        st.markdown("---")

    if data_waiter is not None:
        st.markdown("##### 🧑‍🍳 Kinerja SDM (dari File 3: Waiter)")
        with st.container(border=True):
            # ... (semua kode metrik Waiter A | Delta | B Anda) ...
            col_A, col_Delta, col_B = st.columns([0.35, 0.3, 0.35])
            # (Total Penjualan SDM)
            with col_A:
                st.metric("Total Penjualan (SDM)", format_rupiah(kpi_A_waiter[0]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_waiter[0], kpi_B_waiter[0], format_rupiah, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total Penjualan (SDM)", format_rupiah(kpi_B_waiter[0]))
            # (Total Transaksi SDM)
            with col_A:
                st.metric("Total Transaksi (SDM)", format_angka_bulat(kpi_A_waiter[1]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_waiter[1], kpi_B_waiter[1], format_angka_bulat, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total Transaksi (SDM)", format_angka_bulat(kpi_B_waiter[1]))
            # (Rata-rata/Waiter)
            with col_A:
                st.metric("Rata-rata / Waiter", format_rupiah(kpi_A_waiter[2]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_waiter[2], kpi_B_waiter[2], format_rupiah, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Rata-rata / Waiter", format_rupiah(kpi_B_waiter[2]))

    # #############################################################
    # --- BLOK INSIGHT BARU DITEMPATKAN DI SINI (PALING BAWAH) ---
    # #############################################################

    st.markdown("---")  # Tambahkan pemisah visual
    st.header("💡 Insight Otomatis (Analisis Perbandingan)")

    # Panggil fungsi 'pencari insight' kita
    # Kita gunakan semua KPI A dan B yang sudah dihitung di awal
    insights = generate_comparison_insights(
        kpi_A_gmv, kpi_B_gmv, kpi_A_cogs, kpi_B_cogs, kpi_A_waiter, kpi_B_waiter
    )

    # Tampilkan dalam expander baru
    with st.expander(
        "Klik untuk melihat Temuan Kunci dari Perbandingan A vs B", expanded=True
    ):
        if insights:
            for insight in insights:
                st.markdown(f"&bull; {insight}")  # Tampilkan sebagai daftar
        else:
            st.info("Tidak ada data yang cukup untuk membuat perbandingan insight.")

    # #############################################################
    # --- BATAS BLOK INSIGHT BARU ---
    # #############################################################


def build_tab5_forecast(data_gmv):
    """Menggambar Tab 5 (Peramalan Detail) menggunakan Prophet.
    (VERSI BARU DENGAN KOTAK INSIGHT DI BAWAH)
    """
    st.header("🔮 Peramalan Tren Penjualan Detail")
    st.info(
        "Tab ini menggunakan model `Prophet` untuk menganalisis data GMV Anda, "
        "mendeteksi pola mingguan, dan meramalkan penjualan di masa depan."
    )

    if data_gmv is None or data_gmv.empty:
        st.warning(
            "Silakan upload file Laporan GMV (File 1) di sidebar untuk melihat peramalan tren."
        )
        return

    # --- 1. Agregasi data penjualan per hari ---
    try:
        daily_sales = (
            data_gmv.groupby(data_gmv["Sales Date In"].dt.date)[
                "Total After Bill Discount"
            ]
            .sum()
            .reset_index()
        )
        daily_sales.rename(
            columns={
                "Sales Date In": "ds",
                "Total After Bill Discount": "y",
            },
            inplace=True,
        )
        daily_sales["ds"] = pd.to_datetime(daily_sales["ds"])
        daily_sales = daily_sales.sort_values(by="ds")

    except Exception as e:
        st.error(f"Gagal memproses data GMV untuk peramalan: {e}")
        return

    # --- 2. Cek apakah data cukup ---
    if len(daily_sales) < 15:
        st.warning(
            f"Data tidak cukup untuk peramalan detail. "
            f"Dibutuhkan minimal 15 hari data, Anda memiliki {len(daily_sales)} hari."
        )
        st.dataframe(daily_sales)
        return

    # --- 3. Buat Pilihan Periode Ramalan ---
    st.markdown("---")
    st.subheader("Pengaturan Peramalan")
    last_date = daily_sales["ds"].max()

    forecast_days = st.slider(
        "Pilih jumlah hari ke depan untuk diramal:",
        min_value=7,
        max_value=90,
        value=30,
        step=1,
    )
    st.info(
        f"Model akan dilatih pada data Anda (sampai {last_date.strftime('%d-%m-%Y')}) "
        f"dan meramalkan penjualan untuk **{forecast_days} hari ke depan**."
    )

    # --- 4. Latih Model Prophet ---
    try:
        model = Prophet(weekly_seasonality=True, daily_seasonality=False)
        model.fit(daily_sales)
        future_df = model.make_future_dataframe(periods=forecast_days)
        forecast_df = model.predict(future_df)

        # --- 5. Tampilkan Hasil ---
        st.markdown("---")
        st.subheader(f"📈 Grafik Peramalan {forecast_days} Hari ke Depan")

        fig1 = plot_plotly(model, forecast_df)
        fig1.update_layout(
            title="Peramalan Penjualan (Aktual vs. Prediksi)",
            xaxis_title="Tanggal",
            yaxis_title="Estimasi Penjualan (Rp)",
        )
        st.plotly_chart(fig1, use_container_width=True)

        with st.expander("Lihat Data Tabel Peramalan (Angka Detail)"):
            future_data_table = forecast_df[forecast_df["ds"] > last_date]
            future_data_table_display = future_data_table[
                ["ds", "yhat", "yhat_lower", "yhat_upper"]
            ].copy()

            future_data_table_display.rename(
                columns={
                    "ds": "Tanggal",
                    "yhat": "Estimasi Penjualan",
                    "yhat_lower": "Estimasi Terendah",
                    "yhat_upper": "Estimasi Tertinggi",
                },
                inplace=True,
            )

            st.write(
                "Tabel ini menunjukkan estimasi penjualan Anda beserta rentang kepercayaan (terendah/tertinggi)."
            )
            st.dataframe(
                future_data_table_display.style.format(
                    {
                        "Estimasi Penjualan": format_rupiah,
                        "Estimasi Terendah": format_rupiah,
                        "Estimasi Tertinggi": format_rupiah,
                    }
                ),
                use_container_width=True,
            )

        st.markdown("---")
        st.subheader("📊 Analisis Komponen Tren & Musiman")
        st.write("Grafik ini membedah ramalan Anda:")
        st.write(
            "1. **Trend**: Menunjukkan arah bisnis Anda secara jangka panjang (mengabaikan naik/turun harian)."
        )
        st.write(
            "2. **Weekly**: Menunjukkan pola mingguan. Hari apa yang paling kuat dan paling lemah?"
        )

        fig2 = plot_components_plotly(model, forecast_df)
        st.plotly_chart(fig2, use_container_width=True)

        # #############################################################
        # --- BLOK INSIGHT BARU DITEMPATKAN DI SINI (PALING BAWAH) ---
        # #############################################################

        st.markdown("---")  # Tambahkan pemisah visual
        st.header("💡 Insight Otomatis (Analisis Ramalan)")

        # Panggil fungsi 'pencari insight' kita
        insights = generate_forecast_insights(forecast_df, last_date)

        # Tampilkan dalam expander baru
        with st.expander(
            "Klik untuk melihat Temuan Kunci dari Model Ramalan", expanded=True
        ):
            if insights:
                for insight in insights:
                    st.markdown(f"&bull; {insight}")  # Tampilkan sebagai daftar
            else:
                st.info("Tidak ada insight otomatis yang dapat dibuat dari data ini.")

        # #############################################################
        # --- BATAS BLOK INSIGHT BARU ---
        # #############################################################

    except Exception as e:
        st.error(f"Terjadi kesalahan saat melatih model Prophet: {e}")
        st.exception(e)


def build_tab6_target(data_gmv):
    """Menggambar Tab 6 (Pencapaian Target) dengan perbandingan ramalan Prophet.
    (VERSI 2.0 DENGAN INSIGHT INTERAKTIF)
    """
    st.header("🎯 Pencapaian Target & Proyeksi (Dinamis)")

    if data_gmv is None or data_gmv.empty:
        st.warning(
            "Silakan upload file Laporan GMV (File 1) di sidebar untuk melihat analisis target."
        )
        return

    st.subheader("Pengaturan Analisis")

    day_map = {
        "Monday": "Senin",
        "Tuesday": "Selasa",
        "Wednesday": "Rabu",
        "Thursday": "Kamis",
        "Friday": "Jumat",
        "Saturday": "Sabtu",
        "Sunday": "Minggu",
    }
    day_sort_order = list(day_map.values())
    weekdays_def = ["Senin", "Selasa", "Rabu", "Kamis"]
    weekends_def = ["Jumat", "Sabtu", "Minggu"]

    try:
        data_gmv_copy = data_gmv.copy()
        data_gmv_copy["Bulan-Tahun"] = data_gmv_copy["Sales Date In"].dt.to_period("M")
        available_months = sorted(data_gmv_copy["Bulan-Tahun"].unique(), reverse=True)
        month_options = {
            period: period.strftime("%B %Y") for period in available_months
        }

        selected_month_str = st.selectbox(
            "Pilih Bulan yang Ingin Dianalisis:",
            options=month_options.values(),
            help="Tab ini akan menganalisis target untuk bulan yang Anda pilih.",
        )

        selected_period = [
            period
            for period, str_val in month_options.items()
            if str_val == selected_month_str
        ][0]

        active_month = selected_period.month
        active_year = selected_period.year
        active_month_name = selected_month_str

    except Exception as e:
        st.error(f"Gagal memproses data tanggal untuk filter bulan: {e}")
        return

    with st.container(border=True):
        target_juta = st.number_input(
            f"Target {active_month_name} (dalam Juta Rp)",
            min_value=1,
            value=500,
            step=10,
            help="Masukkan target dalam Juta Rupiah. Misal: ketik 500 untuk Rp 500.000.000",
            key=f"target_juta_{active_month_name}",
        )
        target_bulanan = target_juta * 1_000_000
        st.metric(label=f"Target Anda Diatur ke:", value=format_rupiah(target_bulanan))

    st.markdown("---")

    # Buat kamus untuk menampung semua KPI yang akan dipakai insight
    kpi_insight_dict = {}

    try:
        data_bulan_aktif = data_gmv[
            (data_gmv["Sales Date In"].dt.month == active_month)
            & (data_gmv["Sales Date In"].dt.year == active_year)
        ].copy()

        if data_bulan_aktif.empty:
            st.error(f"Tidak ada data penjualan ditemukan untuk {active_month_name}.")
            return

        latest_data_date_in_full_dataset = data_gmv["Sales Date In"].max()
        max_date_in_selected_month = data_bulan_aktif["Sales Date In"].max()
        total_days_in_month = pd.Period(
            f"{active_year}-{active_month}", freq="M"
        ).days_in_month

        is_latest_month = (active_year == latest_data_date_in_full_dataset.year) and (
            active_month == latest_data_date_in_full_dataset.month
        )

        if is_latest_month:
            hari_berjalan = max_date_in_selected_month.day
            st.info(
                f"Menganalisis bulan berjalan ({active_month_name}). Data terdeteksi sampai tanggal {hari_berjalan}."
            )
        else:
            hari_berjalan = total_days_in_month
            st.success(
                f"Menampilkan ulasan performa bulan lalu ({active_month_name}) yang telah selesai."
            )

        sisa_hari = total_days_in_month - hari_berjalan
        kpi_insight_dict["sisa_hari"] = sisa_hari  # -> Simpan untuk insight

        penjualan_saat_ini = data_bulan_aktif["Total After Bill Discount"].sum()
        kpi_insight_dict["penjualan_saat_ini"] = penjualan_saat_ini

        pencapaian_persen = (
            (penjualan_saat_ini / target_bulanan) if target_bulanan > 0 else 0
        )
        kpi_insight_dict["pencapaian_persen"] = pencapaian_persen  # -> Simpan

        rata_rata_harian_total = (
            penjualan_saat_ini / hari_berjalan if hari_berjalan > 0 else 0
        )
        sales_dibutuhkan = target_bulanan - penjualan_saat_ini

        if sales_dibutuhkan < 0:
            sales_dibutuhkan = 0

        data_bulan_aktif["Nama Hari"] = (
            data_bulan_aktif["Sales Date In"].dt.day_name().map(day_map)
        )
        data_bulan_aktif["Tipe Hari"] = data_bulan_aktif["Nama Hari"].apply(
            lambda x: "Weekend (Jum-Min)" if x in weekends_def else "Weekday (Sen-Kam)"
        )

        daily_sales_agg = (
            data_bulan_aktif.groupby(["Sales Date In", "Tipe Hari", "Nama Hari"])[
                "Total After Bill Discount"
            ]
            .sum()
            .reset_index()
        )

        avg_sales_weekday = daily_sales_agg[
            daily_sales_agg["Tipe Hari"] == "Weekday (Sen-Kam)"
        ]["Total After Bill Discount"].mean()
        avg_sales_weekend = daily_sales_agg[
            daily_sales_agg["Tipe Hari"] == "Weekend (Jum-Min)"
        ]["Total After Bill Discount"].mean()

        if pd.isna(avg_sales_weekday) or avg_sales_weekday < 0:
            avg_sales_weekday = 0
        if pd.isna(avg_sales_weekend) or avg_sales_weekend < 0:
            avg_sales_weekend = 0

        # Logika baru untuk weekend_weight (INI BLOK YANG SUDAH DIPERBAIKI)
        if avg_sales_weekday > 0:
            weekend_weight = avg_sales_weekend / avg_sales_weekday
        elif avg_sales_weekend > 0:
            # Kasus weekday = 0 tapi weekend > 0. Setel bobot sangat tinggi.
            weekend_weight = 1000.0
        else:
            # Kasus keduanya 0, anggap bobotnya sama.
            weekend_weight = 1.0

        kpi_insight_dict["avg_sales_weekday"] = avg_sales_weekday  # -> Simpan
        kpi_insight_dict["avg_sales_weekend"] = avg_sales_weekend  # -> Simpan

        proyeksi_akhir_bulan = penjualan_saat_ini
        rdr_weekday = 0
        rdr_weekend = 0
        proyeksi_prophet = 0

        if sisa_hari > 0:
            # Gunakan total_days_in_month untuk menghitung tanggal akhir bulan yang benar
            tanggal_akhir_bulan = max_date_in_selected_month.replace(
                day=total_days_in_month
            )
            tanggal_mulai_sisa = max_date_in_selected_month + pd.Timedelta(days=1)

            sisa_tanggal_df = pd.DataFrame(
                pd.date_range(start=tanggal_mulai_sisa, end=tanggal_akhir_bulan),
                columns=["Tanggal"],
            )
            sisa_tanggal_df["Nama Hari"] = (
                sisa_tanggal_df["Tanggal"].dt.day_name().map(day_map)
            )

            sisa_weekdays_count = sisa_tanggal_df["Nama Hari"].isin(weekdays_def).sum()
            sisa_weekends_count = sisa_tanggal_df["Nama Hari"].isin(weekends_def).sum()

            proyeksi_sales_sisa = (sisa_weekdays_count * avg_sales_weekday) + (
                sisa_weekends_count * avg_sales_weekend
            )
            proyeksi_akhir_bulan += proyeksi_sales_sisa

            pembagi = sisa_weekdays_count + (sisa_weekends_count * weekend_weight)
            if (
                pembagi > 0 and sales_dibutuhkan > 0
            ):  # <-- Tambahkan cek 'sales_dibutuhkan'
                rdr_weekday = sales_dibutuhkan / pembagi
                rdr_weekend = rdr_weekday * weekend_weight
            else:
                rdr_weekday = 0
                rdr_weekend = 0

            kpi_insight_dict["rdr_weekday"] = rdr_weekday  # -> Simpan
            kpi_insight_dict["rdr_weekend"] = rdr_weekend  # -> Simpan

            try:
                prophet_data = (
                    data_bulan_aktif.groupby(data_bulan_aktif["Sales Date In"].dt.date)[
                        "Total After Bill Discount"
                    ]
                    .sum()
                    .reset_index()
                )
                prophet_data.rename(
                    columns={"Sales Date In": "ds", "Total After Bill Discount": "y"},
                    inplace=True,
                )

                ramalan_sisa_hari = get_prophet_projection(prophet_data, sisa_hari)

                if ramalan_sisa_hari is not None:
                    proyeksi_prophet = penjualan_saat_ini + ramalan_sisa_hari
                else:
                    proyeksi_prophet = proyeksi_akhir_bulan  # Fallback

            except Exception as e_prophet:
                st.warning(f"Gagal memproses data untuk Prophet: {e_prophet}")
                proyeksi_prophet = proyeksi_akhir_bulan

        else:
            proyeksi_akhir_bulan = penjualan_saat_ini
            proyeksi_prophet = penjualan_saat_ini  # Set sama jika bulan selesai

        proyeksi_vs_target_persen = (
            (proyeksi_akhir_bulan / target_bulanan) if target_bulanan > 0 else 0
        )
        kekurangan_proyeksi = target_bulanan - proyeksi_akhir_bulan

        # Simpan sisa data untuk insight
        kpi_insight_dict["proyeksi_vs_target_persen"] = proyeksi_vs_target_persen
        kpi_insight_dict["proyeksi_akhir_bulan"] = proyeksi_akhir_bulan
        kpi_insight_dict["proyeksi_prophet"] = proyeksi_prophet
        kpi_insight_dict["target_bulanan"] = target_bulanan

        if proyeksi_vs_target_persen > 1.05:
            status_color = "normal"
        elif proyeksi_vs_target_persen >= 0.98:
            status_color = "off"
        else:
            status_color = "inverse"

    except Exception as e:
        st.error(f"Gagal mengkalkulasi KPI Target: {e}")
        st.exception(e)
        return

    # --- Sisa kode UI (METRIK DAN GRAFIK) tidak berubah ---

    st.subheader("📈 Gambaran Besar (Pencapaian)")

    col1, col2, col3 = st.columns(3)
    col1.metric(
        label=f"Pencapaian per {max_date_in_selected_month.strftime('%d-%m-%Y')}",
        value=f"{pencapaian_persen*100:,.1f}%",
        help=f"{format_rupiah(penjualan_saat_ini)} dari {format_rupiah(target_bulanan)}",
    )

    if sisa_hari > 0:
        col2.metric(
            label="Penjualan Dibutuhkan",
            value=format_rupiah(sales_dibutuhkan),
            help=f"Sisa penjualan yang harus dikejar dalam {sisa_hari} hari.",
        )
        col3.metric(
            label="Proyeksi Kekurangan (vs Cerdas)",
            value=format_rupiah(kekurangan_proyeksi),
            help="Selisih antara Target Bulanan dan Proyeksi Cerdas.",
        )
    else:
        col2.metric(
            label="Hasil Akhir Bulan (Selesai)",
            value=format_rupiah(penjualan_saat_ini),
            delta=f"{pencapaian_persen*100:,.1f}% dari Target",
            delta_color=status_color,
        )
        col3.metric(
            label="Selisih Target",
            value=format_rupiah(penjualan_saat_ini - target_bulanan),
            help="Hasil akhir penjualan dikurangi target bulanan.",
        )

    # #############################################################
    # --- BLOK INSIGHT BARU DITEMPATKAN DI SINI (PALING BAWAH) ---
    # #############################################################

    st.markdown("---")  # Tambahkan pemisah visual
    st.header("💡 Insight Otomatis (Analisis Target)")

    # Panggil fungsi 'pencari insight' kita
    # kpi_insight_dict sudah diisi di seluruh fungsi ini
    insights = generate_target_insights(kpi_insight_dict)

    # Tampilkan dalam expander baru
    with st.expander(
        "Klik untuk melihat Rangkuman & Rencana Aksi Target Anda",
        expanded=True,
    ):
        if insights:
            for insight_item in insights:
                # Gunakan tipe untuk menentukan st.success/warning/info/error
                if insight_item["type"] == "success":
                    st.success(insight_item["text"], icon="✅")
                elif insight_item["type"] == "warning":
                    st.warning(insight_item["text"], icon="⚠️")
                elif insight_item["type"] == "error":
                    st.error(insight_item["text"], icon="🚨")
                else:  # "info"
                    st.info(insight_item["text"], icon="💡")
        else:
            st.info("Tidak ada insight otomatis yang dapat dibuat dari data ini.")

    # #############################################################
    # --- BATAS BLOK INSIGHT BARU ---
    # #############################################################


def build_tab7_ulasan(data_ulasan):
    """Menggambar Tab 7 (Analisis Ulasan Pelanggan)."""
    st.header("❤️ Analisis Sentimen & Ulasan Pelanggan")

    if data_ulasan is None or data_ulasan.empty:
        st.warning(
            "Silakan upload file Laporan Ulasan (File 4) di sidebar untuk melihat analisis sentimen."
        )
        return

    # --- 1. Definisikan Kata Kunci (Dapat Disesuaikan) ---
    with st.expander("Pengaturan Kustomisasi Topik Ulasan (Opsional)"):
        st.info(
            "Masukkan kata kunci (dipisah koma) untuk melacak topik spesifik dalam ulasan."
        )
        col_k1, col_k2 = st.columns(2)
        keyword_makanan_str = col_k1.text_input(
            "Topik Makanan (cth: enak, lezat, porsi, hambar, asin)",
            "enak,lezat,porsi,hambar,asin,dingin,basi,mantap,nikmat,segar",
        )
        keyword_pelayanan_str = col_k1.text_input(
            "Topik Pelayanan (cth: ramah, cepat, lama, jutek, sopan)",
            "ramah,cepat,lama,jutek,sopan,membantu,lambat,pelayanan",
        )
        keyword_suasana_str = col_k2.text_input(
            "Topik Suasana (cth: nyaman, bersih, kotor, berisik, adem)",
            "nyaman,bersih,kotor,berisik,adem,dingin,panas,cozy,tempat,suasana",
        )
        keyword_harga_str = col_k2.text_input(
            "Topik Harga (cth: murah, mahal, worth it, promo)",
            "murah,mahal,worth it,promo,diskon,terjangkau,harga",
        )

        KEYWORDS = {
            "Makanan": [
                k.strip().lower() for k in keyword_makanan_str.split(",") if k.strip()
            ],
            "Pelayanan": [
                k.strip().lower() for k in keyword_pelayanan_str.split(",") if k.strip()
            ],
            "Suasana": [
                k.strip().lower() for k in keyword_suasana_str.split(",") if k.strip()
            ],
            "Harga": [
                k.strip().lower() for k in keyword_harga_str.split(",") if k.strip()
            ],
        }

    # Fungsi untuk kategorisasi
    @st.cache_data
    def find_topics(ulasan_text, keyword_dict):
        ulasan_text_lower = str(ulasan_text).lower()
        topics_found = []
        for topic, keys in keyword_dict.items():
            for key in keys:
                # Gunakan regex word boundary untuk pencarian yang lebih akurat
                if re.search(r"\b" + re.escape(key) + r"\b", ulasan_text_lower):
                    topics_found.append(topic)
                    break  # Hanya catat satu kali per topik
        if not topics_found:
            return "Lainnya"
        return ", ".join(topics_found)

    # --- 2. Hitung NPS dan Kategori ---
    df = data_ulasan.copy()
    df.dropna(subset=["Ulasan", "Rating_Clean"], inplace=True)

    def categorize_nps(rating):
        if rating >= 5:  # Asumsi rating 5 adalah promoter di skala 1-5
            return "Promoter"
        elif rating == 4:  # Asumsi rating 4 adalah passive
            return "Passive"
        else:  # 1-3
            return "Detractor"

    df["NPS_Category"] = df["Rating_Clean"].apply(categorize_nps)
    df["Topik"] = df["Ulasan"].apply(lambda x: find_topics(x, KEYWORDS))

    total_ulasan = len(df)
    avg_rating = df["Rating_Clean"].mean()

    if total_ulasan > 0:
        promoters_pct = (df["NPS_Category"] == "Promoter").sum() / total_ulasan
        detractors_pct = (df["NPS_Category"] == "Detractor").sum() / total_ulasan
        nps_score = (promoters_pct - detractors_pct) * 100
    else:
        nps_score = 0
        avg_rating = 0

    # --- 3. Tampilkan Metrik KPI ---
    st.subheader("📊 KPI Sentimen Pelanggan")
    kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
    kpi_col1.metric("Total Ulasan", f"{total_ulasan} ulasan")
    kpi_col2.metric("Rata-rata Rating", f"{avg_rating:.1f} / 5 ⭐")
    kpi_col3.metric("Net Promoter Score (NPS)", f"{nps_score:.1f}")

    # --- 4. Tampilkan Grafik Analisis Topik ---
    st.markdown("---")
    st.subheader("🗣️ Analisis Topik Ulasan")

    # Pisahkan ulasan positif (4-5) dan negatif (1-3)
    df_positive = df[df["Rating_Clean"] >= 4]
    df_negative = df[df["Rating_Clean"] <= 3]

    # Hitung topik
    positive_topics = (
        df_positive["Topik"].str.split(", ").explode().value_counts().reset_index()
    )
    negative_topics = (
        df_negative["Topik"].str.split(", ").explode().value_counts().reset_index()
    )
    positive_topics.columns = ["Topik", "Jumlah"]
    negative_topics.columns = ["Topik", "Jumlah"]

    # Tambahkan Peringkat Bintang 1 sebagai topik negatif (sesuai history)
    bintang_1_count = (df["Rating_Clean"] == 1).sum()
    if bintang_1_count > 0:
        new_row = pd.DataFrame(
            [{"Topik": "Rating 1 (Sangat Buruk)", "Jumlah": bintang_1_count}]
        )
        negative_topics = pd.concat([new_row, negative_topics], ignore_index=True)

    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        st.markdown("##### 👍 Topik Positif (Rating 4-5)")
        if not positive_topics.empty:
            chart_pos = create_horizontal_bar_chart(
                positive_topics,
                "Jumlah",
                "Topik",
                "Jumlah Sebutan",
                "Topik Ulasan Positif",
            )
            st.plotly_chart(chart_pos, use_container_width=True)
        else:
            st.info("Tidak ada topik positif yang terdeteksi.")

    with chart_col2:
        st.markdown("##### 👎 Topik Negatif (Rating 1-3)")
        if not negative_topics.empty:
            chart_neg = create_horizontal_bar_chart(
                negative_topics,
                "Jumlah",
                "Topik",
                "Jumlah Sebutan",
                "Topik Ulasan Negatif",
            )
            st.plotly_chart(chart_neg, use_container_width=True)
        else:
            st.info("Tidak ada topik negatif yang terdeteksi.")

    # --- 5. Tampilkan Data Mentah ---
    with st.expander("Lihat Data Mentah Ulasan"):
        st.dataframe(
            df[["Nama", "Rating_Clean", "Ulasan", "NPS_Category", "Topik"]],
            use_container_width=True,
        )

    # --- 6. Blok Insight (PALING BAWAH) ---
    st.markdown("---")
    st.header("💡 Insight Otomatis (Analisis Ulasan)")
    insights = generate_review_insights(
        total_ulasan, avg_rating, nps_score, positive_topics, negative_topics
    )
    with st.expander(
        "Klik untuk melihat Temuan Kunci dari Ulasan Pelanggan", expanded=True
    ):
        if insights:
            for insight in insights:
                st.markdown(f"&bull; {insight}")
        else:
            st.info("Tidak ada insight otomatis yang dapat dibuat dari data ini.")


def build_tab8_purchase(filtered_purchase, total_sales_revenue):
    """
    Menggambar Tab 8 (Analisis Laporan Pembelian).
    DIPERBAIKI: Menambahkan perhitungan dan tampilan Food Cost Percentage (FCP).
    """
    st.header("🛒 Analisis Biaya Pembelian (Purchase)")
    st.info(
        "Tab ini menganalisis Laporan Pembelian (File 5) untuk melacak "
        "pengeluaran berdasarkan kategori, supplier, dan item."
    )

    if filtered_purchase is None:
        st.warning(
            "Silakan upload file Laporan Pembelian (File 5) di sidebar "
            "untuk melihat analisis biaya."
        )
        return

    if filtered_purchase.empty:
        st.warning("Tidak ada data pembelian untuk filter yang dipilih.")
        return

    # --- 1. Analisis Data ---
    # Fungsi ini (analyze_purchase_data) sudah di-cache
    (
        total_cost,
        cost_by_category,
        cost_by_supplier,
        top_items,
        raw_data_filtered,
    ) = analyze_purchase_data(filtered_purchase)

    # --- 2. Perhitungan Food Cost Percentage (FCP) ---
    # Food Cost Percentage (FCP) = (Total Biaya Pembelian / Total Sales Revenue) * 100
    fcp = 0
    # Guardrail: Jika revenue > 0, hitung FCP.
    if total_sales_revenue is not None and total_sales_revenue > 0:
        fcp = (total_cost / total_sales_revenue) * 100

    # --- 3. Tampilkan Metrik KPI ---
    st.subheader("📊 KPI Biaya Pembelian & Food Cost")

    # 🌟 PERBAIKAN: Gunakan 3 kolom untuk menampilkan kedua input dan hasil FCP
    col_kpi_1, col_kpi_2, col_kpi_3 = st.columns(3)

    with col_kpi_1:
        st.metric(
            "Total Biaya Pembelian (Numerator)",
            format_rupiah(total_cost),
            help="Total Cost (Pembilang FCP). Jika 0, FCP akan 0%.",
        )

    with col_kpi_2:
        # 🌟 DEBUG METRIC: Tampilkan Total Sales Revenue sebagai input
        st.metric(
            "Total Sales Revenue (Denominator)",
            format_rupiah(total_sales_revenue if total_sales_revenue else 0),
            help="Total Net Sales dari GMV (Penyebut FCP). Jika 0, FCP akan 0%.",
        )

    with col_kpi_3:
        # Tampilkan Food Cost Percentage
        st.metric(
            "Food Cost Percentage (FCP)",
            f"{fcp:.1f}%",
            help="Dihitung dari (Total Pembelian / Total Sales Revenue) * 100.",
        )

    # --- 4. Tampilkan Grafik Analisis ---
    # ... (Sisa kode visualisasi tetap sama) ...
    st.markdown("---")
    st.subheader("📈 Analisis Rincian Biaya")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("##### Biaya per Kategori (Top 10)")
        if not cost_by_category.empty:
            top_10_cat = cost_by_category.nlargest(10, "Total")
            chart_cat = create_horizontal_bar_chart(
                top_10_cat,
                "Total",
                "Category",
                "Total Biaya (Rp)",
                "Top 10 Biaya per Kategori",
            )
            st.plotly_chart(chart_cat, use_container_width=True)
        else:
            st.info("Tidak ada data biaya per kategori.")

    with col2:
        st.markdown("##### Biaya per Supplier (Top 10)")
        if not cost_by_supplier.empty:
            top_10_supp = cost_by_supplier.nlargest(10, "Total")
            chart_supp = create_horizontal_bar_chart(
                top_10_supp,
                "Total",
                "Supplier Name",
                "Total Biaya (Rp)",
                "Top 10 Biaya per Supplier",
            )
            st.plotly_chart(chart_supp, use_container_width=True)
        else:
            st.info("Tidak ada data biaya per supplier.")

    st.markdown("---")
    st.subheader("💸 Top 20 Item dengan Biaya Tertinggi")
    if not top_items.empty:
        chart_items = create_horizontal_bar_chart(
            top_items,
            "Total",
            "Product Name",
            "Total Biaya (Rp)",
            "Top 20 Item Termahal",
        )
        st.plotly_chart(chart_items, use_container_width=True)
    else:
        st.info("Tidak ada data item termahal.")

    # --- 5. Tampilkan Data Mentah ---
    with st.expander("Lihat Rincian Data Pembelian (Sudah Difilter)"):
        st.dataframe(
            raw_data_filtered.style.format(
                {"Price": format_rupiah, "Total": format_rupiah}
            ),
            use_container_width=True,
        )

    # --- 6. Blok Insight ---
    st.markdown("---")
    st.header("💡 Insight Otomatis (Analisis Pembelian)")
    # Insight harus diperbarui untuk menyertakan FCP
    insights = generate_purchase_insights(
        total_cost, cost_by_category, cost_by_supplier, top_items, fcp
    )
    with st.expander(
        "Klik untuk melihat Temuan Kunci dari Biaya Pembelian", expanded=True
    ):
        if insights:
            for insight in insights:
                st.markdown(f"&bull; {insight}")
        else:
            st.info("Tidak ada insight otomatis yang dapat dibuat dari data ini.")


# #################################################################
# --- FUNGSI BARU UNTUK INSIGHT DI TAB 9 (REKOMENDASI) ---
# #################################################################
@st.cache_data
def generate_recommendation_insights(rules_df):
    """
    Menganalisis data 'rules_df' yang sudah diproses dan menghasilkan insight
    dalam bahasa alami untuk Tab 9.
    """
    insights = []

    # Pastikan data tidak kosong
    if rules_df is None or rules_df.empty:
        return ["Tidak ada aturan rekomendasi yang cukup kuat untuk dianalisis."]

    try:
        # 1. Insight 1: Potensi Sales Tertinggi (Expected Value)
        # Kita asumsikan rules_df sudah di-sort by expected_value, tapi kita sort lagi
        top_ev_rule = rules_df.nlargest(1, "expected_value").iloc[0]
        item_A_ev = top_ev_rule["antecedents"]
        item_B_ev = top_ev_rule["consequents"]
        ev = top_ev_rule["expected_value"]

        insights.append(
            f"**💸 Potensi Sales Tertinggi:** Aturan paling bernilai adalah **JIKA BELI `{item_A_ev}`, TAWARKAN `{item_B_ev}`**. "
            f"Setiap tawaran ini memiliki *expected value* (potensi sales) sebesar **{format_rupiah(ev)}**."
        )

        # 2. Insight 2: Pasangan Paling Setia (Confidence)
        top_conf_rule = rules_df.nlargest(1, "confidence").iloc[0]
        item_A_conf = top_conf_rule["antecedents"]
        item_B_conf = top_conf_rule["consequents"]
        conf = top_conf_rule["confidence"]

        insights.append(
            f"**🤝 Pasangan Paling Setia:** Aturan paling *pasti* (confidence tertinggi) adalah **JIKA BELI `{item_A_conf}`, TAWARKAN `{item_B_conf}`**. "
            f"**{conf:.1%}** pelanggan yang membeli `{item_A_conf}` juga membeli `{item_B_conf}`."
        )

        # 3. Insight 3: Koneksi Terkuat (Lift)
        top_lift_rule = rules_df.nlargest(1, "lift").iloc[0]
        item_A_lift = top_lift_rule["antecedents"]
        item_B_lift = top_lift_rule["consequents"]
        lift = top_lift_rule["lift"]

        insights.append(
            f"**🔗 Koneksi Terkuat (Lift):** Pasangan **`{item_A_lift}`** dan **`{item_B_lift}`** adalah yang paling unik. "
            f"Pelanggan **{lift:.1f}x lebih mungkin** membeli keduanya bersamaan daripada secara acak. Ini adalah peluang *cross-sell* yang kuat."
        )

        # 4. Insight 4: Aksi
        insights.append(
            "**💡 Aksi:** Gunakan filter **'JIKA Pelanggan Beli Menu Ini'** di atas untuk "
            "mengeksplorasi rekomendasi spesifik untuk item terlaris Anda."
        )

    except Exception as e:
        print(f"Gagal generate insight rekomendasi: {e}")
        insights.append(f"Gagal membuat insight: {e}")

    return insights


# #################################################################
# --- FUNGSI BARU UNTUK TAB 9 (REKOMENDASI) ---
# #################################################################


@st.cache_data
def get_market_basket_rules(df, min_support_threshold):
    """
    Menjalankan Market Basket Analysis (Apriori) pada data GMV.
    VERSI BARU: Menambahkan data harga dan filter 1 item.
    """
    try:

        # Filter Input Data (dari permintaan sebelumnya)
        filter_regex = "PACKAGE|REFILL OCHA"
        df_no_packages = df[
            ~df["Menu"].str.contains(filter_regex, na=False, case=False, regex=True)
        ]

        # Ambil data harga menu
        menu_prices = df_no_packages.groupby("Menu")["Price (Net)"].mean()

        # 1. Transformasi Data
        with st.spinner(
            f"Memproses {df_no_packages['Bill Number'].nunique()} transaksi..."
        ):
            menu_counts = df_no_packages["Menu"].value_counts()
            relevant_menus = menu_counts[menu_counts > 1].index
            df_filtered = df_no_packages[df_no_packages["Menu"].isin(relevant_menus)]
            transactions_list = (
                df_filtered.groupby("Bill Number")["Menu"].apply(list).values.tolist()
            )

            if not transactions_list:
                st.warning(
                    "Tidak ada transaksi yang cukup untuk dianalisis (setelah filter)."
                )
                return pd.DataFrame()

        # 2. Encode Transaksi
        with st.spinner("Meng-encode data (TransactionEncoder)..."):
            te = TransactionEncoder()
            te_ary = te.fit(transactions_list).transform(transactions_list)
            df_encoded = pd.DataFrame(te_ary, columns=te.columns_)

        # 3. Jalankan Algoritma Apriori
        with st.spinner(f"Menjalankan Apriori (support={min_support_threshold})..."):
            frequent_itemsets = apriori(
                df_encoded, min_support=min_support_threshold, use_colnames=True
            )

        if frequent_itemsets.empty:
            st.warning(
                "Tidak ditemukan pola item yang cukup kuat. Coba turunkan 'Minimal Support'."
            )
            return pd.DataFrame()

        # 4. Buat Aturan Asosiasi
        with st.spinner("Membangun aturan asosiasi..."):
            rules = association_rules(
                frequent_itemsets, metric="lift", min_threshold=1.0
            )

        if rules.empty:
            st.info("Tidak ada aturan asosiasi yang ditemukan dengan lift > 1.")
            return pd.DataFrame()

        # Filter Logika Bisnis (Antecedents)
        with st.spinner("Menerapkan filter logika bisnis (add-ons)..."):
            addon_keywords_list = [
                "UPGRADE",
                "ADDITIONAL",
                "ADD ON",
                "ADD-ON",
                "REFILL",
                "OCHA",
                "MINERAL WATER",
            ]

            def check_if_addon(item_set):
                try:
                    for item in item_set:
                        item_upper = str(item).upper()
                        for keyword in addon_keywords_list:
                            if keyword in item_upper:
                                return True
                    return False
                except Exception:
                    return False

            rules = rules[~rules["antecedents"].apply(check_if_addon)]

        if rules.empty:
            st.info(
                "Tidak ada aturan asosiasi yang tersisa setelah filter logika bisnis."
            )
            return pd.DataFrame()

        # #############################################################
        # --- PERBAIKAN BARU: Filter 'Jika Beli' Hanya 1 Item ---
        # Sesuai permintaan, kita hanya ingin aturan A -> B, bukan (A,C) -> B
        # 'antecedents' saat ini masih berupa frozenset, jadi kita bisa cek 'len()'
        # #############################################################

        rules = rules[rules["antecedents"].apply(len) == 1]

        if rules.empty:
            st.info("Tidak ada aturan (single item) yang tersisa setelah filter.")
            return pd.DataFrame()
        # #############################################################
        # --- BATAS PERBAIKAN ---
        # #############################################################

        # Hitung Expected Value
        with st.spinner("Menghitung Expected Value (Confidence x Harga)..."):

            # Ubah frozenset consequents menjadi string
            # Kita filter juga agar consequents HANYA 1 item (untuk logika bersih)
            rules = rules[rules["consequents"].apply(len) == 1]
            rules["consequents_str"] = rules["consequents"].apply(
                lambda x: next(iter(x))
            )

            # Gabungkan (merge) dengan data harga
            rules = rules.merge(
                menu_prices.rename("consequent_price"),
                left_on="consequents_str",
                right_index=True,
                how="left",
            )
            rules["consequent_price"] = rules["consequent_price"].fillna(0)
            rules["expected_value"] = rules["confidence"] * rules["consequent_price"]

        # 5. Bersihkan Hasil
        # Sekarang kita aman untuk konversi ke string, karena kita tahu len == 1
        rules["antecedents"] = rules["antecedents"].apply(lambda x: ", ".join(list(x)))
        rules["consequents"] = rules["consequents"].apply(lambda x: ", ".join(list(x)))

        # URUTKAN berdasarkan metrik baru kita!
        rules = rules.sort_values(["expected_value"], ascending=[False])

        return rules[
            [
                "antecedents",
                "consequents",
                "confidence",
                "lift",
                "consequent_price",  # Tambahkan kolom ini
                "expected_value",  # Tambahkan kolom ini
            ]
        ]
    except Exception as e:
        st.error(f"Gagal menjalankan Market Basket Analysis: {e}")
        st.exception(e)
        return pd.DataFrame()


def build_tab9_rekomendasi(filtered_gmv):
    """
    Menggambar Tab 9 (Rekomendasi Menu / Market Basket Analysis).
    VERSI FINAL: Hanya filter "JIKA BELI"
    """
    st.header("💡 Rekomendasi Menu Interaktif")
    st.info(
        "Pilih menu di filter 'JIKA Beli' untuk melihat semua item yang "
        "paling berpotensi meningkatkan sales jika ditawarkan bersamaan."
    )

    if filtered_gmv is None or filtered_gmv.empty:
        st.warning(
            "Silakan upload file Laporan GMV (File 1) di sidebar "
            "untuk melihat analisis rekomendasi."
        )
        return

    if filtered_gmv["Bill Number"].nunique() < 50:
        st.warning(
            f"Jumlah transaksi terlalu sedikit ({filtered_gmv['Bill Number'].nunique()}) untuk analisis yang akurat. Minimal 50 transaksi."
        )
        return

    # Panggil fungsi analisis yang sudah di-cache
    with st.spinner(
        "Menganalisis semua aturan asosiasi (bisa perlu waktu 1-2 menit)..."
    ):
        # Kita set support sangat rendah agar dapat semua kemungkinan data
        rules_df = get_market_basket_rules(filtered_gmv, 0.001)

    if rules_df.empty:
        st.error("Gagal menjalankan analisis. Tidak ada aturan yang ditemukan.")
        return

    # #############################################################
    # --- FILTER INTERAKTIF BARU (HANYA 1 FILTER) ---
    # #############################################################

    st.subheader("Filter Analisis Rekomendasi")

    # --- Buat Filter UI (Hanya 1) ---
    all_antecedents = sorted(rules_df["antecedents"].unique())

    filter_antecedents = st.multiselect(
        "JIKA Pelanggan Beli Menu Ini:",
        options=all_antecedents,
        placeholder="Ketik menu... (Kosongkan = Tampilkan Semua Aturan Teratas)",
    )

    # --- 3. Terapkan Filter ke DataFrame ---
    with st.spinner("Menerapkan filter..."):
        filtered_rules = rules_df.copy()

        # Filter 1: Berdasarkan 'Jika Beli' (Antecedents)
        if filter_antecedents:
            filtered_rules = filtered_rules[
                filtered_rules["antecedents"].isin(filter_antecedents)
            ]

    # --- 4. Tampilkan Hasil ---
    st.markdown("---")
    st.subheader(f"Hasil Rekomendasi (Total: {len(filtered_rules)} aturan)")

    # Penjelasan metrik
    col_a, col_b = st.columns(2)
    with col_a:
        st.success("**Apa itu Confidence (Kepercayaan)?**")
        st.write(
            "**Confidence 50%** pada (A -> B) berarti: "
            "Dari semua orang yang membeli **Menu A**, **50%** dari mereka "
            "**juga membeli Menu B**."
        )
    with col_b:
        st.error("**Apa itu Expected Value (Nilai Harapan)?**")
        st.write(
            "**Expected Value = Confidence x Harga Menu B.** "
            "Ini adalah metrik terbaik untuk menentukan prioritas. "
            "Semakin tinggi, semakin besar potensi sales."
        )

    st.markdown("---")

    if not filtered_rules.empty:

        # Urutkan lagi berdasarkan Expected Value (terpenting!)
        filtered_rules = filtered_rules.sort_values("expected_value", ascending=False)

        # Tampilkan DataFrame
        st.dataframe(
            filtered_rules.rename(
                columns={
                    "antecedents": "Jika Beli Menu Ini (A)",
                    "consequents": "Tawarkan Menu Ini (B)",
                    "confidence": "Confidence (A->B)",
                    "lift": "Lift",
                    "consequent_price": "Harga (Tawaran B)",
                    "expected_value": "Expected Value (Potensi Sales)",
                }
            )[
                [  # Tampilkan dalam urutan kolom yang logis
                    "Jika Beli Menu Ini (A)",
                    "Tawarkan Menu Ini (B)",
                    "Expected Value (Potensi Sales)",
                    "Confidence (A->B)",
                    "Harga (Tawaran B)",
                    "Lift",
                ]
            ].style.format(
                {
                    "Confidence (A->B)": "{:.2%}",
                    "Lift": "{:.2f}x",
                    "Harga (Tawaran B)": format_rupiah,
                    "Expected Value (Potensi Sales)": format_rupiah,
                }
            ),
            use_container_width=True,
        )
    else:
        st.info(
            "Tidak ada rekomendasi yang cocok dengan filter Anda. Coba perlebar filter Anda."
        )

    # #############################################################
    # --- TAMBAHAN BARU: BLOK INSIGHT ---
    # #############################################################
    st.markdown("---")  # Tambahkan pemisah visual
    st.header("💡 Insight Otomatis (Rekomendasi Menu)")

    # Panggil fungsi 'pencari insight' kita
    # --- PERBAIKAN DI SINI: Gunakan 'rules_df' untuk insight global, bukan 'filtered_rules' ---
    insights = generate_recommendation_insights(filtered_rules)

    # Tampilkan dalam expander baru
    with st.expander(
        "Klik untuk melihat Temuan Kunci dari Pola Belanja Pelanggan", expanded=True
    ):
        if insights:
            for insight in insights:
                st.markdown(f"&bull; {insight}")  # Tampilkan sebagai daftar
        else:
            st.info("Tidak ada insight otomatis yang dapat dibuat dari data ini.")


# #################################################################
# --- FUNGSI BARU UNTUK TAB 10 (ANALISIS PROMO) ---
# #################################################################

# [ ... KODE ANDA SEBELUMNYA ... ]


def build_tab10_promo(filtered_gmv, filtered_cogs):
    """
    Menggambar Tab 10 (Analisis & Simulasi Promo)
    VERSI 4.0: Menggunakan 'Price (Pricelist)' dari GMV dan 'COGS' dari COGS.
    """
    st.header("💸 Analisis & Simulator Profitabilitas Promo")
    st.info(
        "Gunakan alat ini untuk mensimulasikan profit dari skenario diskon atau promo."
    )

    # #############################################################
    # --- PERUBAHAN BESAR: GABUNGKAN DATA GMV & COGS ---
    # #############################################################

    # 1. Cek ketersediaan data
    if filtered_gmv is None or filtered_cogs is None:
        st.error(
            "Simulator Promo memerlukan File 1 (GMV) dan File 2 (COGS) untuk di-upload bersamaan."
        )
        return

    # 2. Ambil Harga Jual (Price Pricelist) dari GMV
    @st.cache_data
    def get_gmv_prices(df_gmv):
        # --- PERUBAHAN UTAMA DI SINI ---
        # Kita cek kolom 'Price (Pricelist)'
        if "Price (Pricelist)" not in df_gmv.columns:
            st.error("Kolom 'Price (Pricelist)' tidak ditemukan di File 1 (GMV).")
            st.info("Simulator akan menggunakan 'Price (Net)' sebagai gantinya.")
            # Fallback jika kolom tidak ada
            if "Price (Net)" not in df_gmv.columns:
                st.error(
                    "Kolom 'Price (Net)' juga tidak ditemukan. Gagal memuat harga."
                )
                return pd.DataFrame()

            gmv_data = df_gmv[df_gmv["Price (Net)"] > 0]
            prices_df = gmv_data.groupby("Menu")["Price (Net)"].mean().reset_index()
            prices_df.rename(columns={"Price (Net)": "Harga Jual"}, inplace=True)
            return prices_df

        else:
            # --- JALUR IDEAL (MENGGUNAKAN PRICE PRICELIST) ---
            gmv_data = df_gmv[df_gmv["Price (Pricelist)"] > 0]
            # Ambil harga rata-rata 'Price (Pricelist)'
            prices_df = (
                gmv_data.groupby("Menu")["Price (Pricelist)"].mean().reset_index()
            )
            # Kita tetap menamainya 'Harga Jual' agar sisa simulator berfungsi
            prices_df.rename(columns={"Price (Pricelist)": "Harga Jual"}, inplace=True)
            return prices_df
        # --- BATAS PERUBAHAN ---

    # 3. Ambil COGS dari file COGS
    @st.cache_data
    def get_cogs_costs(df_cogs):
        cogs_df = df_cogs[df_cogs["COGS"] > 0]
        cogs_unit_df = cogs_df.groupby("Menu")["COGS"].mean().reset_index()
        return cogs_unit_df

    with st.spinner("Menggabungkan data Harga Jual (GMV) dan COGS..."):
        prices_df = get_gmv_prices(filtered_gmv)
        cogs_df = get_cogs_costs(filtered_cogs)

        if prices_df.empty:
            st.error("Gagal mendapatkan data harga dari File GMV.")
            return

        # 4. Gabungkan (Merge) data
        profit_df = pd.merge(prices_df, cogs_df, on="Menu", how="inner")

        # 5. Hitung profitabilitas
        if not profit_df.empty:
            profit_df["Margin (Rp)"] = profit_df["Harga Jual"] - profit_df["COGS"]

    # --- AKHIR PERUBAHAN BESAR ---

    if profit_df is None or profit_df.empty:
        st.error(
            "Tidak ada menu yang cocok (nama menu yang sama) antara File GMV dan File COGS."
        )
        return

    menu_list = profit_df[profit_df["Harga Jual"] > 0]["Menu"].unique()
    if len(menu_list) == 0:
        st.warning("Tidak ada data menu yang valid untuk dianalisis.")
        return

    st.markdown("---")

    # #############################################################
    # --- SCENARIO 1: DISKON MENU TUNGGAL ---
    # #############################################################
    st.subheader("Scenario 1: Analisis Diskon Menu Tunggal")
    st.write(
        "Pilih satu menu, atur diskon, dan lihat potensi profit Anda "
        "jika Anda berhasil menjual sejumlah target kuantitas."
    )

    col1_s1, col2_s1 = st.columns([2, 1])
    with col1_s1:
        selected_menu_s1 = st.selectbox(
            "Pilih Menu untuk Didiskon:",
            menu_list,
            key="promo_menu_s1",
            placeholder="Ketik untuk mencari menu...",
        )
    with col2_s1:
        discount_percent = st.slider(
            "Persentase Diskon (%):", 0, 100, 15, key="promo_discount"
        )
    expected_qty_s1 = st.number_input(
        "Target Kuantitas Terjual (selama periode promo):",
        min_value=1,
        value=50,
        step=10,
        key="promo_qty_s1",
    )

    if selected_menu_s1:
        try:
            menu_data = profit_df[profit_df["Menu"] == selected_menu_s1].iloc[0]

            # 'Harga Jual' ini SEKARANG adalah 'Price (Pricelist)' dari GMV
            original_price = menu_data["Harga Jual"]
            cogs = menu_data["COGS"]
            original_margin_rp = menu_data["Margin (Rp)"]

            discount_rp = original_price * (discount_percent / 100)
            new_price = original_price - discount_rp
            new_margin_rp = new_price - cogs
            total_new_profit = new_margin_rp * expected_qty_s1
            total_original_profit = original_margin_rp * expected_qty_s1
            selisih_profit = total_new_profit - total_original_profit
            delta_color = "normal" if selisih_profit >= 0 else "inverse"

            st.markdown("##### 📈 Hasil Simulasi Diskon (Per Item)")
            with st.container(border=True):
                kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
                kpi_col1.metric(
                    "Harga Jual Promo (Nett)",
                    format_rupiah(new_price),
                    f"Normal: {format_rupiah(original_price)}",
                    delta_color="inverse",
                )
                kpi_col2.metric(
                    "COGS (Biaya Modal)",
                    format_rupiah(cogs),
                    "Biaya Tetap",
                    delta_color="off",
                )
                kpi_col3.metric(
                    "Profit Promo (per Item)",
                    format_rupiah(new_margin_rp),
                    f"Normal: {format_rupiah(original_margin_rp)}",
                    delta_color=(
                        "inverse" if new_margin_rp < original_margin_rp else "normal"
                    ),
                )
            st.metric(
                f"🎉 TOTAL PROFIT (jika {expected_qty_s1} terjual)",
                format_rupiah(total_new_profit),
                f"Selisih: {format_rupiah(selisih_profit)} vs Normal",
                delta_color=delta_color,
            )
            st.info(
                f"**Perbandingan Profit (vs Jual Normal @ {expected_qty_s1} porsi):**\n"
                f"* **Total Profit (Harga Normal):** {format_rupiah(total_original_profit)}\n"
                f"* **Total Profit (Harga Promo):** {format_rupiah(total_new_profit)}\n"
                f"* **Selisih Profit:** **{format_rupiah(selisih_profit)}**"
            )
        except Exception as e:
            st.error(f"Gagal menghitung simulasi untuk {selected_menu_s1}: {e}")

    st.markdown("---")

    # #############################################################
    # --- SCENARIO 2: PROMO BOGO / BUY A GET B ---
    # #############################################################
    st.subheader("Scenario 2: Analisis Promo Paket (Buy A, Get B)")
    st.write(
        "Simulasikan promo 'Buy One Get One' atau 'Beli A, Gratis B'. "
        "**Untuk BOGO (Buy 1 Get 1 Free)**, pilih menu yang *sama* di kedua kotak."
    )

    col1_s2, col2_s2 = st.columns(2)
    with col1_s2:
        menu_A = st.selectbox(
            "JIKA Pelanggan Membeli (Menu A):",
            menu_list,
            key="promo_menu_a",
            placeholder="Pilih menu berbayar...",
        )
    with col2_s2:
        menu_B = st.selectbox(
            "MAKA Mendapat Gratis (Menu B):",
            menu_list,
            key="promo_menu_b",
            placeholder="Pilih menu gratis...",
        )
    expected_deals_s2 = st.number_input(
        "Target Paket Promo Terjual:",
        min_value=1,
        value=30,
        step=5,
        key="promo_qty_s2",
    )

    if menu_A and menu_B:
        try:
            data_A = profit_df[profit_df["Menu"] == menu_A].iloc[0]
            price_A = data_A["Harga Jual"]
            cogs_A = data_A["COGS"]
            original_margin_A = data_A["Margin (Rp)"]

            data_B = profit_df[profit_df["Menu"] == menu_B].iloc[0]
            price_B_normal = data_B["Harga Jual"]
            cogs_B = data_B["COGS"]

            revenue_per_deal = price_A
            cost_per_deal = cogs_A + cogs_B
            profit_per_deal = revenue_per_deal - cost_per_deal
            total_promo_profit = profit_per_deal * expected_deals_s2
            total_normal_profit_A_only = original_margin_A * expected_deals_s2
            selisih_profit_s2 = total_promo_profit - total_normal_profit_A_only
            delta_color_s2 = "inverse"

            st.markdown("##### 📦 Rincian Item (Per Paket)")
            with st.container(border=True):
                kpi_item_a, kpi_item_b = st.columns(2)
                kpi_item_a.metric(
                    f"Menu A (Bayar): {menu_A}",
                    format_rupiah(price_A),
                    "Revenue Diterima",
                )
                kpi_item_b.metric(
                    f"Menu B (Gratis): {menu_B}",
                    format_rupiah(0),
                    f"Harga Normal: {format_rupiah(price_B_normal)}",
                    delta_color="inverse",
                )

            st.markdown("##### 📈 Hasil Profit (Per Paket)")
            with st.container(border=True):
                kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
                kpi_col1.metric(
                    "Total Revenue (Paket)",
                    format_rupiah(revenue_per_deal),
                    "Hanya dari Menu A",
                )
                kpi_col2.metric(
                    "Total COGS (Paket)",
                    format_rupiah(cost_per_deal),
                    f"COGS A + COGS B",
                    delta_color="inverse",
                )
                kpi_col3.metric(
                    "Profit Bersih (PakET)",
                    format_rupiah(profit_per_deal),
                    "Revenue - Total COGS",
                )

            st.metric(
                f"🎉 TOTAL PROFIT (jika {expected_deals_s2} paket terjual)",
                format_rupiah(total_promo_profit),
                f"Selisih: {format_rupiah(selisih_profit_s2)} vs Jual Normal (Menu A Saja)",
                delta_color=delta_color_s2,
            )
            st.info(
                f"**Perbandingan Profit (vs Jual Normal Menu A @ {expected_deals_s2} porsi):**\n"
                f"* **Total Profit (Hanya Jual Menu A):** {format_rupiah(total_normal_profit_A_only)}\n"
                f"* **Total Profit (Promo Paket):** {format_rupiah(total_promo_profit)}\n"
                f"* **Biaya/Kerugian Promo (Selisih):** **{format_rupiah(selisih_profit_s2)}**"
            )

            if menu_A == menu_B:
                st.warning(
                    f"**Insight BOGO:** Menjual **{expected_deals_s2} paket BOGO** "
                    f"`{menu_A}` akan memberikan total profit **{format_rupiah(total_promo_profit)}**. "
                    f"Ini adalah strategi untuk *volume* (menjual {expected_deals_s2*2} item) "
                    f"dengan mengorbankan margin."
                )

        except Exception as e:
            st.error(f"Gagal menghitung simulasi paket: {e}")


# #################################################################
# --- FUNGSI BARU UNTUK ANALISIS TAB 11 (TEMPLATE) ---
# #################################################################
@st.cache_data
def analyze_tab11_weekend_effect(df_gmv, df_kalender):
    """
    (PERBAIKAN LOGIKA) Menganalisis Efek Weekend vs Holiday.
    Hanya menganggap 'Libur Nasional' sebagai Holiday, Musim dianggap hari biasa.
    """
    if df_gmv is None or df_gmv.empty or df_kalender is None:
        return None, None

    try:
        # 1. Persiapan Data Harian
        df_sales = df_gmv.copy()

        # Pastikan nama kolom sesuai file Anda
        # (Ganti jika nama kolom di file Anda berbeda)
        col_tgl = "Sales Date In"
        col_revenue = "Total After Bill Discount"

        df_sales["Tanggal"] = pd.to_datetime(df_sales[col_tgl]).dt.date

        # Hitung Total Penjualan per Hari
        df_daily = df_sales.groupby("Tanggal")[col_revenue].sum().reset_index()
        df_daily["Tanggal"] = pd.to_datetime(df_daily["Tanggal"])

        # 2. Gabungkan dengan Kalender
        df_kalender["Tanggal"] = pd.to_datetime(df_kalender["Tanggal"])
        df_merged = pd.merge(df_daily, df_kalender, on="Tanggal", how="left")

        # 3. LOGIKA BARU (YANG LEBIH SPESIFIK)
        # ---------------------------------------------------------
        df_merged["Is_Weekend"] = df_merged["Tanggal"].dt.dayofweek >= 5

        # Isi NaN dengan 'Biasa'
        df_merged["Tipe_Event"] = df_merged["Tipe_Event"].fillna("Biasa")

        # KUNCI PERBAIKAN:
        # Hanya anggap Holiday jika Tipe_Event mengandung kata "Nasional" atau "Cuti"
        # Musim (Ramadan/Libur Sekolah) akan dianggap False (Hari Biasa)
        df_merged["Is_Holiday"] = df_merged["Tipe_Event"].str.contains(
            "Libur Nasional|Cuti Bersama|Hari Raya", case=False, regex=True
        )
        # ---------------------------------------------------------

        def get_category(row):
            # Prioritas 1: Tanggal Merah (Libur Nasional)
            if row["Is_Holiday"]:
                if row["Is_Weekend"]:
                    return "4. Weekend & Libur"  # Sabtu Pahing (Merah)
                else:
                    return "3. Weekday Libur"  # Harpitnas / Kejepit (Merah)

            # Prioritas 2: Hari Biasa (Termasuk Musim Libur Sekolah tapi bukan tgl merah)
            else:
                if row["Is_Weekend"]:
                    return "2. Weekend Biasa"  # Sabtu-Minggu Normal
                else:
                    return "1. Weekday Biasa"  # Senin-Jumat Normal

        df_merged["Kategori_Hari"] = df_merged.apply(get_category, axis=1)

        # 4. Hitung Rata-rata per Kategori
        df_result = (
            df_merged.groupby("Kategori_Hari")
            .agg(
                Rata_Rata_Omzet=(col_revenue, "mean"),
                Jumlah_Hari=("Tanggal", "count"),
                Total_Omzet=(col_revenue, "sum"),
            )
            .reset_index()
        )

        return df_result, df_merged

    except Exception as e:
        st.error(f"Error analisis Weekend vs Holiday: {e}")
        return None, None


# #################################################################
# --- FUNGSI BARU UNTUK INSIGHT DI TAB 11 (TEMPLATE) ---
# #################################################################
@st.cache_data(show_spinner=False)
def generate_tab11_insights(analysis_data):
    """
    (PERBAIKAN) Menganalisis data musiman yang sudah diproses
    dan menghasilkan insight.
    """
    insights = []

    if not analysis_data:
        return ["Tidak ada data untuk dianalisis."]

    try:
        df_impact = analysis_data["event_impact"]
        df_monthly = analysis_data["monthly_trend"]

        # 1. Insight Performa Hari Libur vs Hari Biasa
        try:
            sales_biasa = df_impact[df_impact["Tipe_Event"] == "Hari Biasa"][
                "Rata_Rata_Penjualan"
            ].values[0]
            sales_libur = df_impact[df_impact["Tipe_Event"] == "Libur Nasional"][
                "Rata_Rata_Penjualan"
            ].values[0]

            if sales_biasa > 0:
                persentase = ((sales_libur - sales_biasa) / sales_biasa) * 100
                if persentase > 0:
                    insights.append(
                        f"**📈 Performa Liburan:** Penjualan pada **Libur Nasional** rata-rata **{persentase:.1f}% lebih tinggi** "
                        f"(**{format_rupiah(sales_libur)}**) dibandingkan **Hari Biasa** (**{format_rupiah(sales_biasa)}**)."
                    )
                else:
                    insights.append(
                        f"**📉 Performa Liburan:** Penjualan pada **Libur Nasional** rata-rata **{abs(persentase):.1f}% lebih rendah** "
                        f"(**{format_rupiah(sales_libur)}**) dibandingkan **Hari Biasa** (**{format_rupiah(sales_biasa)}**)."
                    )
        except (IndexError, KeyError):
            pass  # Gagal membuat insight ini jika data tidak ada

        # 2. Insight High Season (Bulan Puncak)
        try:
            puncak = df_monthly.loc[df_monthly["Total_Penjualan"].idxmax()]
            insights.append(
                f"**☀️ High Season:** Bulan puncak penjualan Anda adalah **{puncak['Bulan_Tahun']}** "
                f"dengan total penjualan **{format_rupiah(puncak['Total_Penjualan'])}**."
            )
        except (ValueError, KeyError):
            pass

        # 3. Insight Low Season (Bulan Terendah)
        try:
            terendah = df_monthly.loc[df_monthly["Total_Penjualan"].idxmin()]
            insights.append(
                f"**❄️ Low Season:** Bulan penjualan terendah Anda adalah **{terendah['Bulan_Tahun']}** "
                f"dengan total penjualan **{format_rupiah(terendah['Total_Penjualan'])}**."
            )
        except (ValueError, KeyError):
            pass

        # 4. Insight Tipe Event Lain
        try:
            musim_liburan = df_impact[df_impact["Tipe_Event"] == "Musim Liburan"][
                "Rata_Rata_Penjualan"
            ].values[0]
            if musim_liburan > sales_biasa:
                insights.append(
                    f"**🏖️ Musim Liburan:** Periode 'Musim Liburan' (sesuai kalender Anda) menunjukkan penjualan harian rata-rata "
                    f"**{format_rupiah(musim_liburan)}**, yang lebih tinggi dari Hari Biasa."
                )
        except (IndexError, KeyError):
            pass  # Gagal jika tidak ada 'Musim Liburan'

    except Exception as e:
        print(f"Gagal generate insight Tab 11: {e}")
        insights.append(f"Gagal membuat insight: {e}")

    return insights


# #################################################################
# --- FUNGSI BARU UNTUK MEMBANGUN UI TAB 11
# #################################################################
# --- PASTIKAN FUNGSI HELPER INI ADA DI FILE ANDA (ATAU GUNAKAN YANG INI) ---
def format_rupiah(angka):
    """Format angka ke Rupiah Indonesia (Rp 1.000.000) dengan aman."""
    if pd.isna(angka):
        return "Rp 0"
    # Format jadi string angka dulu, lalu ganti koma jadi titik
    return f"Rp {angka:,.0f}".replace(",", ".")


# -----------------------------------------------------------------------------
# GANTI SELURUH FUNGSI build_tab11_new DENGAN KODE DI BAWAH INI
# -----------------------------------------------------------------------------
def build_tab11_musiman(df_gmv, df_kalender):
    """
    (FINAL COMPLETED - INSIGHT DI BAWAH) Tab 11: Grafik Total + Insight Rata-rata.
    """
    st.header("⚔️ Analisis: Weekend vs. Tanggal Merah")
    st.info(
        "Grafik batang menampilkan **TOTAL PENDAPATAN**, sedangkan Insight teks menganalisis **RATA-RATA HARIAN**."
    )

    if df_gmv is None or df_kalender is None:
        st.warning("Data belum lengkap.")
        return

    # Panggil fungsi analisis
    df_summary, df_raw = analyze_tab11_weekend_effect(df_gmv, df_kalender)

    if df_summary is not None and not df_summary.empty:

        # 1. Format Label Rupiah (Untuk Grafik Total)
        df_summary["Label_Total"] = (
            df_summary["Total_Omzet"].fillna(0).apply(format_rupiah)
        )
        df_summary["Label_Rata"] = (
            df_summary["Rata_Rata_Omzet"].fillna(0).apply(format_rupiah)
        )

        # ============================================================
        # 1. GRAFIK UTAMA: BAR CHART (TOTAL OMZET)
        # ============================================================
        st.subheader("1. Total Kontribusi Omzet per Jenis Hari")
        fig = px.bar(
            df_summary,
            x="Kategori_Hari",
            y="Total_Omzet",
            color="Kategori_Hari",
            text="Label_Total",
            # PENTING: Gunakan LIST [...] agar urutan customdata[0], [1] terkunci pasti
            hover_data=["Label_Rata", "Jumlah_Hari"],
            color_discrete_map={
                "1. Weekday Biasa": "#bdc3c7",
                "2. Weekend Biasa": "#3498db",
                "3. Weekday Libur": "#e67e22",
                "4. Weekend & Libur": "#e74c3c",
            },
        )
        # Format Tooltip
        fig.update_traces(
            texttemplate="%{text}",
            textposition="outside",
            hovertemplate=(
                "<b>%{x}</b><br>"
                + "💰 Total Omzet: <b>%{text}</b><br>"
                + "📊 Rata-rata/Hari: <b>%{customdata[0]}</b><br>"
                + "📅 Jumlah Hari: %{customdata[1]}"
                + "<extra></extra>"
            ),
        )
        fig.update_layout(
            showlegend=False,
            yaxis_title="Total Omzet (Akumulasi)",
            yaxis=dict(showticklabels=False),
            margin=dict(t=50, b=50),
        )
        st.plotly_chart(fig, use_container_width=True)

        # ============================================================
        # 2. TABEL RINCIAN EVENT
        # ============================================================
        st.markdown("---")
        st.subheader("2. Rincian Hari Libur Spesifik")
        if "Nama_Event" in df_raw.columns:
            df_libur = df_raw[
                df_raw["Kategori_Hari"].str.contains("Libur", na=False)
            ].copy()

            if not df_libur.empty:
                df_libur["Tanggal_Str"] = df_libur["Tanggal"].dt.strftime("%d %B %Y")
                df_libur["Omzet"] = df_libur["Total After Bill Discount"].apply(
                    format_rupiah
                )

                # Urutkan dulu baru pilih kolom
                df_libur_sorted = df_libur.sort_values("Tanggal", ascending=True)
                tabel_show = df_libur_sorted[
                    ["Tanggal_Str", "Nama_Event", "Kategori_Hari", "Omzet"]
                ]

                st.dataframe(
                    tabel_show.rename(
                        columns={
                            "Tanggal_Str": "Tanggal",
                            "Nama_Event": "Nama Liburan",
                            "Kategori_Hari": "Jenis Hari",
                        }
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info("Tidak ada Hari Libur Nasional di periode yang dipilih.")

        # ============================================================
        # 3. GRAFIK DETAIL: SCATTER PLOT
        # ============================================================
        st.markdown("---")
        st.subheader("3. Sebaran Data Harian (Detail)")
        with st.expander("Klik untuk melihat grafik sebaran detail", expanded=True):
            try:
                df_raw["Hover_Rupiah"] = df_raw["Total After Bill Discount"].apply(
                    format_rupiah
                )
                df_raw["Hover_Tanggal"] = df_raw["Tanggal"].dt.strftime("%d-%m-%Y")
                df_raw["Hover_Event"] = df_raw["Nama_Event"].fillna("-")
            except:
                df_raw["Hover_Rupiah"] = df_raw["Total After Bill Discount"].astype(str)
                df_raw["Hover_Event"] = "-"

            fig_box = px.box(
                df_raw,
                x="Kategori_Hari",
                y="Total After Bill Discount",
                color="Kategori_Hari",
                points=False,
                title="Distribusi Penjualan Harian",
            )
            fig_points = px.strip(
                df_raw,
                x="Kategori_Hari",
                y="Total After Bill Discount",
                color="Kategori_Hari",
                hover_data=["Hover_Rupiah", "Hover_Tanggal", "Hover_Event"],
            )
            for trace in fig_points.data:
                trace.hovertemplate = (
                    "<b>%{customdata[2]}</b><br>"
                    + "Tanggal: %{customdata[1]}<br>"
                    + "Omzet: <b>%{customdata[0]}</b><br>"
                    + "<extra></extra>"
                )
                trace.marker.size = 7
                trace.marker.opacity = 0.7
                trace.marker.line.width = 1
                trace.marker.line.color = "white"
                trace.showlegend = False
                fig_box.add_trace(trace)

            fig_box.update_layout(
                showlegend=False,
                yaxis_title="Omzet Harian",
                yaxis=dict(tickprefix="Rp "),
            )
            st.plotly_chart(fig_box, use_container_width=True)

        # ============================================================
        # 4. INSIGHT & KESIMPULAN (PINDAH KE PALING BAWAH)
        # ============================================================
        st.markdown("---")
        st.header("💡 Insight & Temuan Penting")
        st.write(
            "*Insight di bawah ini dihitung berdasarkan Rata-rata Penjualan per Hari (Apple-to-Apple).*"
        )

        # Persiapan Data untuk Insight
        data_avg = dict(zip(df_summary["Kategori_Hari"], df_summary["Rata_Rata_Omzet"]))

        avg_wd_biasa = data_avg.get("1. Weekday Biasa", 0)
        avg_wd_libur = data_avg.get("3. Weekday Libur", 0)
        avg_weekend = data_avg.get("2. Weekend Biasa", 0)
        avg_weekend_libur = data_avg.get("4. Weekend & Libur", 0)

        col_ins1, col_ins2 = st.columns(2)

        with col_ins1:
            st.subheader("📅 Efek 'Tanggal Merah' di Hari Kerja")
            if avg_wd_libur > 0 and avg_wd_biasa > 0:
                kenaikan = ((avg_wd_libur - avg_wd_biasa) / avg_wd_biasa) * 100
                selisih = avg_wd_libur - avg_wd_biasa

                if kenaikan > 0:
                    st.success(
                        f"**Positif:** Jika hari kerja (Senin-Jumat) adalah Tanggal Merah, "
                        f"omzet harian Anda NAIK rata-rata **{kenaikan:.1f}%** "
                        f"(+ {format_rupiah(selisih)}) dibandingkan hari kerja biasa."
                    )
                else:
                    st.warning(
                        f"**Negatif:** Tanggal merah di hari kerja justru menurunkan omzet rata-rata "
                        f"sebesar **{abs(kenaikan):.1f}%**. (Mungkin target pasar Anda adalah karyawan kantor?)"
                    )
            elif avg_wd_libur == 0:
                st.info(
                    "Belum ada data 'Tanggal Merah di Hari Kerja' pada periode yang dipilih."
                )

        with col_ins2:
            st.subheader("🆚 Battle: Weekend vs. Libur")
            if avg_wd_libur > 0 and avg_weekend > 0:
                if avg_wd_libur > avg_weekend:
                    selisih = avg_wd_libur - avg_weekend
                    st.success(
                        f"**Pemenang: Tanggal Merah!**\n"
                        f"Ternyata, 'Tanggal Merah di Hari Kerja' lebih ramai daripada 'Weekend Biasa'. "
                        f"Selisih rata-rata: **{format_rupiah(selisih)}**."
                    )
                else:
                    st.info(
                        f"**Pemenang: Weekend Biasa.**\n"
                        f"Sabtu-Minggu biasa ternyata masih lebih kuat performanya dibandingkan Tanggal Merah yang jatuh di hari kerja."
                    )
            else:
                st.write("Data pembanding belum lengkap.")

        # --- RINGKASAN ANGKA (METRIC) ---
        st.markdown("##### Ringkasan Rata-rata Omzet per Hari:")
        col_a, col_b, col_c, col_d = st.columns(4)
        col_a.metric("Senin-Jumat (Biasa)", format_rupiah(avg_wd_biasa))
        col_b.metric("Sabtu-Minggu (Biasa)", format_rupiah(avg_weekend))
        col_c.metric(
            "Senin-Jumat (LIBUR)",
            format_rupiah(avg_wd_libur),
            delta="vs Biasa",
            delta_color="normal",
        )
        col_d.metric("Sabtu-Minggu (LIBUR)", format_rupiah(avg_weekend_libur))

    else:
        st.warning("Tidak ada data hasil analisis.")


# Pastikan baris ini ada di bagian paling atas file Anda (bersama import lainnya)
import plotly.express as px


def build_tab_unique_features(df_gmv, df_cogs):
    """
    (VERSI INTERAKTIF: FILTER TABEL BERDASARKAN KUADRAN)
    """
    st.header("🧪 Lab Strategi & Eksperimen")
    st.caption(
        "Analisis ini menggunakan **SEMUA DATA** (tidak terpengaruh filter tanggal) untuk akurasi strategi."
    )

    if df_gmv is None or df_cogs is None:
        st.warning("⚠️ Harap upload **File 1 (GMV)** dan **File 2 (COGS)** di sidebar.")
        return

    try:
        # --- 1. PERSIAPAN DATA ---
        # A. Cari Kolom Harga
        col_harga_fix = "Price (Pricelist)"
        if col_harga_fix not in df_gmv.columns:
            col_harga_fix = "Price (Net)"

        # B. Agregasi Sales
        df_sales_prep = df_gmv.copy()
        df_sales_prep["Menu_Match"] = (
            df_sales_prep["Menu"].astype(str).str.strip().str.upper()
        )

        df_sales_summary = (
            df_sales_prep.groupby("Menu_Match")
            .agg(
                Menu_Asli=("Menu", "first"),
                Total_Qty=("Qty", "sum"),
                Total_Revenue=("Total After Bill Discount", "sum"),
                Harga_Jual_Fix=(col_harga_fix, "mean"),
            )
            .reset_index()
        )

        # C. Agregasi COGS
        df_cogs_prep = df_cogs.copy()
        df_cogs_prep["Menu_Match"] = (
            df_cogs_prep["Menu"].astype(str).str.strip().str.upper()
        )

        df_cogs_summary = (
            df_cogs_prep.groupby("Menu_Match")
            .agg(COGS_Satuan=("COGS", "mean"))
            .reset_index()
        )

        # D. Gabungkan
        df_master = pd.merge(
            df_sales_summary, df_cogs_summary, on="Menu_Match", how="inner"
        )

        if df_master.empty:
            st.error("❌ **DATA TIDAK COCOK!** Tidak ada nama menu yang sama.")
            return

        # E. Hitung Profit
        df_master["Menu"] = df_master["Menu_Asli"]
        df_master["Harga_Jual"] = df_master["Harga_Jual_Fix"]
        df_master["Total_COGS"] = df_master["Total_Qty"] * df_master["COGS_Satuan"]
        df_master["Gross_Profit"] = df_master["Total_Revenue"] - df_master["Total_COGS"]

        # Filter data valid
        df_master = df_master[df_master["Total_Qty"] > 0]

        # =================================================================
        # FITUR 1: MATRIX MENU ENGINEERING
        # =================================================================
        st.subheader("1. Matrix Menu Engineering")

        avg_qty = df_master["Total_Qty"].mean()
        avg_profit = df_master["Gross_Profit"].mean()

        def get_quadrant(row):
            high_qty = row["Total_Qty"] >= avg_qty
            high_profit = row["Gross_Profit"] >= avg_profit

            if high_qty and high_profit:
                return "⭐ STARS (Laris & Untung)"
            elif high_qty and not high_profit:
                return "🐴 PLOWHORSE (Laris, Margin Tipis)"
            elif not high_qty and high_profit:
                return "❓ PUZZLE (Jarang Laku, Untung Besar)"
            else:
                return "🐕 DOGS (Kurang Laku, Rugi)"

        df_master["Kuadran"] = df_master.apply(get_quadrant, axis=1)

        # Format Data untuk Tooltip & Tampilan
        df_master["Hover_Profit"] = df_master["Gross_Profit"].apply(format_rupiah)
        df_master["Hover_Revenue"] = df_master["Total_Revenue"].apply(format_rupiah)
        df_master["Hover_Qty"] = df_master["Total_Qty"].apply(lambda x: f"{x:,.0f}")

        # --- GRAFIK ---
        fig_matrix = px.scatter(
            df_master,
            x="Total_Qty",
            y="Gross_Profit",
            color="Kuadran",
            hover_name="Menu",
            hover_data={
                "Kuadran": False,
                "Total_Qty": False,
                "Gross_Profit": False,
                "Hover_Qty": True,
                "Hover_Profit": True,
                "Hover_Revenue": True,
            },
            size="Total_Revenue",
            size_max=40,
            title="Peta Kekuatan Menu (Stars vs Dogs)",
            color_discrete_map={
                "⭐ STARS (Laris & Untung)": "#2ecc71",
                "🐴 PLOWHORSE (Laris, Margin Tipis)": "#f1c40f",
                "❓ PUZZLE (Jarang Laku, Untung Besar)": "#3498db",
                "🐕 DOGS (Kurang Laku, Rugi)": "#e74c3c",
            },
        )
        fig_matrix.update_traces(
            hovertemplate="<b>%{hovertext}</b><br>📦 Terjual: %{customdata[0]}<br>💰 Profit: %{customdata[1]}<br>💵 Omzet: %{customdata[2]}<extra></extra>"
        )
        fig_matrix.add_hline(
            y=avg_profit, line_dash="dot", annotation_text="Rata-rata Profit"
        )
        fig_matrix.add_vline(
            x=avg_qty, line_dash="dot", annotation_text="Rata-rata Qty"
        )
        fig_matrix.update_layout(
            yaxis=dict(tickprefix="Rp "),
            xaxis_title="Jumlah Terjual (Qty)",
            yaxis_title="Total Profit (Gross)",
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1
            ),
        )

        st.plotly_chart(fig_matrix, use_container_width=True)

        # --- INTERAKSI: FILTER TABEL DI BAWAH ---
        st.markdown("---")
        st.subheader("📋 Detail Menu per Kuadran")
        st.write(
            "Pilih kategori di bawah ini untuk melihat daftar menu secara spesifik."
        )

        # Pilihan Kuadran
        pilihan_kuadran = st.selectbox(
            "Tampilkan Data:",
            [
                "Semua Menu",
                "⭐ STARS (Laris & Untung)",
                "🐴 PLOWHORSE (Laris, Margin Tipis)",
                "❓ PUZZLE (Jarang Laku, Untung Besar)",
                "🐕 DOGS (Kurang Laku, Rugi)",
            ],
        )

        # Filter Dataframe
        if pilihan_kuadran == "Semua Menu":
            df_tampil = df_master
        else:
            df_tampil = df_master[df_master["Kuadran"] == pilihan_kuadran]

        # Rapikan Kolom untuk Tampilan
        df_table_view = df_tampil[
            [
                "Menu",
                "Kuadran",
                "Total_Qty",
                "Harga_Jual",
                "COGS_Satuan",
                "Gross_Profit",
            ]
        ].copy()
        df_table_view = df_table_view.sort_values("Gross_Profit", ascending=False)

        # Format Angka di Tabel
        df_table_view["Total_Qty"] = df_table_view["Total_Qty"].apply(
            lambda x: f"{x:,.0f}"
        )
        df_table_view["Harga_Jual"] = df_table_view["Harga_Jual"].apply(format_rupiah)
        df_table_view["COGS_Satuan"] = df_table_view["COGS_Satuan"].apply(format_rupiah)
        df_table_view["Gross_Profit"] = df_table_view["Gross_Profit"].apply(
            format_rupiah
        )

        st.dataframe(df_table_view, use_container_width=True, hide_index=True)

        # Insight Cepat sesuai Pilihan
        if "STARS" in pilihan_kuadran:
            st.success(
                "💡 **Tips STARS:** Menu ini sudah sempurna. Pastikan stok tidak pernah kosong dan kualitas rasa konsisten."
            )
        elif "DOGS" in pilihan_kuadran:
            st.error(
                "💡 **Tips DOGS:** Menu ini membebani operasional. Coba ganti resep, ubah nama, atau hapus dari buku menu."
            )
        elif "PLOWHORSE" in pilihan_kuadran:
            st.warning(
                "💡 **Tips PLOWHORSE:** Ini menu populer tapi untungnya tipis. Coba naikkan harga sedikit atau kurangi porsi bahan mahal."
            )
        elif "PUZZLE" in pilihan_kuadran:
            st.info(
                "💡 **Tips PUZZLE:** Untungnya besar tapi jarang dibeli. Perbaiki foto di buku menu atau minta waiter menawarkannya aktif."
            )

        # =================================================================
        # FITUR 2: SIMULATOR PROFIT
        # =================================================================
        st.markdown("---")
        st.subheader("2. 🔮 Simulator Profit")

        menu_list = sorted(df_master["Menu"].unique())
        selected_menu_sim = st.selectbox("Pilih Menu:", options=menu_list)

        if selected_menu_sim:
            item_data = df_master[df_master["Menu"] == selected_menu_sim].iloc[0]

            curr_price = item_data["Harga_Jual"]
            curr_cogs = item_data["COGS_Satuan"]
            curr_qty = item_data["Total_Qty"]
            curr_profit_total = item_data["Gross_Profit"]

            st.markdown(
                f"**Status Saat Ini:** Harga List `{format_rupiah(curr_price)}` | HPP `{format_rupiah(curr_cogs)}` | Profit Total `{format_rupiah(curr_profit_total)}`"
            )

            col1, col2, col3 = st.columns(3)
            with col1:
                pct_price = st.slider("Ubah Harga (%)", -100, 100, 0)
                new_price = curr_price * (1 + pct_price / 100)
                st.metric("Harga Baru", format_rupiah(new_price))
            with col2:
                pct_cogs = st.slider("Ubah HPP (%)", -100, 100, 0)
                new_cogs = curr_cogs * (1 + pct_cogs / 100)
                st.metric("HPP Baru", format_rupiah(new_cogs))
            with col3:
                pct_qty = st.slider("Dampak Qty (%)", -100, 100, 0)
                new_qty = curr_qty * (1 + pct_qty / 100)
                st.metric("Qty Baru", f"{new_qty:,.0f}")

            new_profit_total = (new_price - new_cogs) * new_qty
            delta = new_profit_total - curr_profit_total

            st.metric(
                "Estimasi Profit Total Baru",
                format_rupiah(new_profit_total),
                delta=format_rupiah(delta),
            )

            if delta > 0:
                st.success("✅ **PROFIT NAIK!** Skenario ini menguntungkan.")
            elif delta < 0:
                st.error("⚠️ **PROFIT TURUN!** Hati-hati dengan skenario ini.")
            else:
                st.info("Profit tetap sama.")

    except Exception as e:
        st.error(f"Terjadi error di Lab Strategi: {e}")


def build_tab13_pl(df_pl):
    """
    Membangun Tab 13: Analisis P&L dengan visualisasi yang Rapi & Profesional.
    """
    st.header("📉 Laporan Laba Rugi (Profit & Loss)")
    st.caption(
        "Analisis kinerja keuangan komprehensif dengan visualisasi data akuntansi."
    )

    if df_pl is None or df_pl.empty:
        st.warning(
            "⚠️ Data P&L belum tersedia atau tidak ada angka valid. Silakan upload **File 6** di sidebar."
        )
        return

    # --- 1. Filter Periode (Tahun) ---
    # Otomatis pilih tahun yang datanya paling banyak jika Current Year kosong
    if "Year_Type" in df_pl.columns:
        years = sorted(
            df_pl["Year_Type"].unique(), reverse=True
        )  # Current Year biasanya di awal
        selected_year_type = st.selectbox("📅 Pilih Periode Laporan:", years, index=0)
        df_filtered = df_pl[df_pl["Year_Type"] == selected_year_type]
    else:
        df_filtered = df_pl.copy()

    if df_filtered.empty:
        st.warning("Tidak ada data untuk periode yang dipilih.")
        return

    # --- 2. Kalkulasi Metrik Utama ---
    total_revenue = df_filtered[df_filtered["Category"] == "Revenue"]["Value"].sum()
    total_cogs = df_filtered[df_filtered["Category"] == "COGS"]["Value"].sum()
    total_expense = df_filtered[df_filtered["Category"] == "Expense"]["Value"].sum()

    gross_profit = total_revenue - total_cogs
    net_profit = gross_profit - total_expense

    # Hitung Persentase (Margin)
    gp_margin = (gross_profit / total_revenue * 100) if total_revenue > 0 else 0
    np_margin = (net_profit / total_revenue * 100) if total_revenue > 0 else 0
    cogs_pct = (total_cogs / total_revenue * 100) if total_revenue > 0 else 0
    opex_pct = (total_expense / total_revenue * 100) if total_revenue > 0 else 0

    # --- 3. Struktur Tab ---
    tab_utama, tab_rincian = st.tabs(["📊 Dashboard Utama", "💸 Rincian Biaya"])

    # =================================================================
    # TAB 1: DASHBOARD UTAMA
    # =================================================================
    with tab_utama:
        # A. KPI Cards (Baris Atas - Style Clean)
        st.subheader("Ringkasan Kinerja")
        kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)

        with kpi1:
            st.metric("Total Revenue", format_rupiah(total_revenue), help="Total Omzet")
        with kpi2:
            st.metric(
                "HPP (COGS)",
                format_rupiah(total_cogs),
                f"{cogs_pct:.1f}%",
                delta_color="inverse",
            )
        with kpi3:
            st.metric("Gross Profit", format_rupiah(gross_profit), f"{gp_margin:.1f}%")
        with kpi4:
            st.metric(
                "Beban (OpEx)",
                format_rupiah(total_expense),
                f"{opex_pct:.1f}%",
                delta_color="inverse",
            )
        with kpi5:
            st.metric(
                "Net Profit",
                format_rupiah(net_profit),
                f"{np_margin:.1f}%",
                delta_color="normal" if net_profit > 0 else "inverse",
            )

        st.markdown("---")

        # B. Grafik Utama (Waterfall & Trend)
        col_chart1, col_chart2 = st.columns(2)

        # 1. Waterfall Chart (Aliran Profit)
        with col_chart1:
            st.markdown("##### 💧 Aliran Profitabilitas (Waterfall)")

            measures = ["relative", "relative", "total", "relative", "total"]
            x_labels = ["Revenue", "HPP", "Gross Profit", "Beban", "Net Profit"]
            y_values = [
                total_revenue,
                -total_cogs,
                gross_profit,
                -total_expense,
                net_profit,
            ]

            # Format teks label (singkat agar rapi di chart, detail di hover)
            text_labels = [format_rupiah(v) for v in y_values]

            fig_waterfall = go.Figure(
                go.Waterfall(
                    name="Flow",
                    orientation="v",
                    measure=measures,
                    x=x_labels,
                    y=y_values,
                    text=text_labels,
                    textposition="outside",
                    connector={"line": {"color": "#bdc3c7"}},
                    decreasing={"marker": {"color": "#e74c3c"}},  # Merah Soft
                    increasing={"marker": {"color": "#2ecc71"}},  # Hijau Soft
                    totals={"marker": {"color": "#3498db"}},  # Biru Soft
                )
            )

            fig_waterfall.update_layout(
                title="",
                showlegend=False,
                waterfallgap=0.3,
                template="plotly_white",
                height=400,
                margin=dict(t=20, b=20, l=20, r=20),
                yaxis=dict(showgrid=True, gridcolor="#f5f5f5"),
            )
            st.plotly_chart(fig_waterfall, use_container_width=True)

        # 2. Trend Chart (Omzet vs Profit)
        with col_chart2:
            st.markdown("##### 📈 Tren Bulanan: Omzet vs Profit")

            # Siapkan Data Bulanan
            df_monthly = (
                df_filtered.groupby("Date")
                .apply(
                    lambda x: pd.Series(
                        {
                            "Revenue": x[x["Category"] == "Revenue"]["Value"].sum(),
                            "Net_Profit": x[x["Category"] == "Revenue"]["Value"].sum()
                            - x[x["Category"] == "COGS"]["Value"].sum()
                            - x[x["Category"] == "Expense"]["Value"].sum(),
                        }
                    )
                )
                .reset_index()
            )

            # Format label untuk tooltip
            df_monthly["Label_Rev"] = df_monthly["Revenue"].apply(format_rupiah)
            df_monthly["Label_Profit"] = df_monthly["Net_Profit"].apply(format_rupiah)

            fig_trend = go.Figure()

            # Bar Revenue
            fig_trend.add_trace(
                go.Bar(
                    x=df_monthly["Date"],
                    y=df_monthly["Revenue"],
                    name="Revenue",
                    marker_color="#3498db",
                    opacity=0.6,
                    customdata=df_monthly["Label_Rev"],
                    hovertemplate="Bulan: %{x|%b %Y}<br>Revenue: <b>%{customdata}</b><extra></extra>",
                )
            )

            # Line Net Profit
            fig_trend.add_trace(
                go.Scatter(
                    x=df_monthly["Date"],
                    y=df_monthly["Net_Profit"],
                    name="Net Profit",
                    mode="lines+markers",
                    line=dict(color="#27ae60", width=3),
                    marker=dict(size=8, color="#27ae60"),
                    customdata=df_monthly["Label_Profit"],
                    hovertemplate="Bulan: %{x|%b %Y}<br>Net Profit: <b>%{customdata}</b><extra></extra>",
                )
            )

            fig_trend.update_layout(
                title="",
                template="plotly_white",
                height=400,
                legend=dict(orientation="h", y=1.1, x=0.5, xanchor="center"),
                margin=dict(t=40, b=20, l=20, r=20),
                hovermode="x unified",
                yaxis=dict(showgrid=True, gridcolor="#f5f5f5"),
            )
            st.plotly_chart(fig_trend, use_container_width=True)

    # =================================================================
    # TAB 2: RINCIAN BIAYA & STRUKTUR
    # =================================================================
    with tab_rincian:
        st.subheader("Bedah Biaya & Struktur Keuangan")

        col_d1, col_d2 = st.columns([1, 1.5])

        # 1. Donut Chart: Struktur Biaya
        with col_d1:
            st.markdown("##### 🍩 Struktur Pengeluaran")

            # Data Donut (Hanya nilai positif untuk visualisasi)
            val_profit = max(0, net_profit)
            labels = ["HPP (COGS)", "Beban (OpEx)", "Net Profit"]
            values = [total_cogs, total_expense, val_profit]
            colors = ["#e74c3c", "#f1c40f", "#2ecc71"]  # Merah, Kuning, Hijau

            fig_donut = go.Figure(
                data=[
                    go.Pie(
                        labels=labels,
                        values=values,
                        hole=0.5,
                        marker=dict(colors=colors),
                        textinfo="percent",  # Tampilkan persen di dalam donut
                        textfont_size=14,
                        hovertemplate="<b>%{label}</b><br>Nilai: %{value:,.0f}<br>Porsi: %{percent}<extra></extra>",
                    )
                ]
            )

            fig_donut.update_layout(
                showlegend=True,
                legend=dict(orientation="h", y=-0.1),
                margin=dict(t=20, b=20, l=20, r=20),
                height=350,
            )
            st.plotly_chart(fig_donut, use_container_width=True)

        # 2. Bar Chart: Top 10 Expenses
        with col_d2:
            st.markdown("##### 💸 10 Pos Biaya Terbesar")

            df_exp = df_filtered[df_filtered["Category"] == "Expense"]

            if not df_exp.empty:
                top_exp = (
                    df_exp.groupby("Description")["Value"]
                    .sum()
                    .reset_index()
                    .sort_values("Value", ascending=False)
                    .head(10)
                )
                top_exp["Label_Value"] = top_exp["Value"].apply(format_rupiah)

                # Gunakan grafik Horizontal Bar yang bersih
                fig_bar = go.Figure(
                    go.Bar(
                        x=top_exp["Value"],
                        y=top_exp["Description"],
                        orientation="h",
                        text=top_exp["Label_Value"],  # Tampilkan Rupiah di ujung bar
                        textposition="auto",
                        marker_color="#e74c3c",
                        opacity=0.8,
                        customdata=top_exp["Label_Value"],
                        hovertemplate="<b>%{y}</b><br>Biaya: %{customdata}<extra></extra>",
                    )
                )

                fig_bar.update_layout(
                    yaxis=dict(autorange="reversed"),  # Urutkan dari atas ke bawah
                    xaxis=dict(
                        showgrid=False, showticklabels=False
                    ),  # Sembunyikan axis X agar bersih
                    template="plotly_white",
                    margin=dict(t=20, b=20, l=20, r=20),
                    height=350,
                )
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.info("Data rincian biaya tidak ditemukan.")

        # 3. Tabel Rincian (Expandable)
        with st.expander("📋 Lihat Tabel Rincian Biaya Lengkap"):
            if not df_exp.empty:
                detail_table = (
                    df_exp.groupby("Description")["Value"]
                    .sum()
                    .reset_index()
                    .sort_values("Value", ascending=False)
                )
                detail_table["Value"] = detail_table["Value"].apply(format_rupiah)
                st.dataframe(
                    detail_table.rename(
                        columns={
                            "Description": "Nama Akun Biaya",
                            "Value": "Total Biaya",
                        }
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

    # --- Insight Otomatis ---
    st.markdown("---")
    st.header("💡 Insight Cerdas")

    insights = generate_pl_insights(df_pl)
    with st.expander("Klik untuk melihat Analisis Kinerja Keuangan", expanded=True):
        if insights:
            for i in insights:
                st.markdown(f"&bull; {i}")
        else:
            st.info("Belum ada data yang cukup untuk analisis otomatis.")


# #################################################################
# --- [PINDAHKAN FUNGSI HELPER LOTTIE/WELCOME/FOOTER KE SINI] ---
# --- (Definisikan SEBELUM 'main()' yang memanggilnya) ---
# #################################################################

import json  # Pastikan ini ada di atas file (atau di bagian import utama)


# Fungsi ini membaca file dari komputer Anda
def load_lottiefile(filepath: str):
    """Memuat Lottie JSON dari file lokal."""
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"PERINGATAN: File Lottie tidak ditemukan di {filepath}")
        return None
    except Exception as e:
        print(f"Error saat memuat file Lottie: {e}")
        return None


# --- Di dalam fungsi build_welcome_screen ---
def build_welcome_screen():
    """
    Menampilkan halaman selamat datang dengan 3 animasi berdampingan
    dalam 3 kolom yang sama rata.
    """

    st.title("🚀 Selamat Datang di Dashboard Analisis F&B!")
    st.subheader(
        "**Ubah data mentah Anda menjadi insight yang dapat ditindaklanjuti.**"
    )
    st.markdown("---")

    # --- [BLOK ANIMASI BARU: 3 KOLOM SAMA RATA] ---

    col_anim1, col_anim2, col_anim3 = st.columns(3)

    with col_anim1:
        lottie_json_main = load_lottiefile("pic2.json")
        if lottie_json_main:
            st_lottie(
                lottie_json_main,
                speed=1,
                loop=True,
                quality="high",
                height=300,  # Atur tinggi agar seragam
                width="100%",
                key="lottie_main_col",
            )
        else:
            st.warning("Gagal memuat 'pic2.json'.", icon="⚠️")

    with col_anim2:
        lottie_json_processing = load_lottiefile("pic1.json")
        if lottie_json_processing:
            st_lottie(
                lottie_json_processing,
                speed=1,
                loop=True,
                quality="medium",
                height=300,  # Atur tinggi agar seragam
                width="100%",
                key="lottie_processing_col",
            )
        else:
            st.warning("Gagal memuat 'pic1.json'.", icon="⚠️")

    with col_anim3:
        lottie_json_charts = load_lottiefile("pic3.json")
        if lottie_json_charts:
            st_lottie(
                lottie_json_charts,
                speed=1,
                loop=True,
                quality="medium",
                height=300,  # Atur tinggi agar seragam
                width="100%",
                key="lottie_charts_col",
            )
        else:
            st.warning("Gagal memuat 'pic3.json'.", icon="⚠️")

    # --- AKHIR BLOK ANIMASI ---

    st.markdown("---")

    # 2. Call to Action (CTA) Utama
    st.header("👇 Mulai Analisis Anda dalam 3 Langkah Mudah")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(
            """
            <div style="padding: 20px; border: 1px solid #444; border-radius: 10px; height: 100%; text-align: center;">
            <h2 style="text-align: center;">1. Buka Sidebar</h2>
            <p style="text-align: center; font-size: 3rem;">⬅️</p>
            <p style="text-align: center;">Buka panel di sebelah kiri (klik tanda > jika tersembunyi).</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col2:
        st.markdown(
            """
            <div style="padding: 20px; border: 1px solid #444; border-radius: 10px; height: 100%; text-align: center;">
            <h2 style="text-align: center;">2. Upload Data</h2>
            <p style="text-align: center; font-size: 3rem;">📤</p>
            <p style="text-align: center;">Upload <b>Laporan GMV (File 1)</b> Anda untuk memulai. File lain bersifat opsional.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col3:
        st.markdown(
            """
            <div style="padding: 20px; border: 1px solid #444; border-radius: 10px; height: 100%; text-align: center;">
            <h2 style="text-align: center;">3. Lihat Insight</h2>
            <p style="text-align: center; font-size: 3rem;">💡</p>
            <p style="text-align: center;">Dashboard akan otomatis memproses data dan menampilkan analisis lengkap.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.info(
        "**Tips:** Anda juga dapat mencentang 'Gunakan data terakhir dari database' "
        "jika Anda sudah pernah menyimpan data sebelumnya.",
        icon="ℹ️",
    )

    st.markdown("---")

    # 3. Penjelasan Fitur (YANG SUDAH DIPERBARUI)
    st.header("✨ Apa yang Bisa Anda Lakukan?")
    st.markdown(
        """
    * **📊 Analisis Penjualan (GMV):** Lihat performa penjualan, menu terlaris, dan jam sibuk.
    * **💰 Analisis COGS & Profit:** Temukan menu mana yang paling *profitabel*, bukan hanya paling laku.
    * **🧑‍🍳 Kinerja SDM:** Lacak performa waiter, Cek Anomali Transaksi dan lihat siapa top sales Anda.
    * **🛒 Biaya Pembelian:** Kontrol pengeluaran dengan menganalisis biaya per supplier dan per item.
    * **⚖️ A/B Comparison:** Bandingkan kinerja 'Minggu Ini' vs 'Minggu Lalu' secara berdampingan.
    * **🎯 Pelacakan Target:** Masukkan target bulanan Anda dan lihat proyeksi pencapaian secara *real-time*.
    * **🔮 Forecast (AI):** Gunakan AI (Prophet) untuk meramalkan tren penjualan Anda 30 hari ke depan.
    * **❤️ Ulasan Pelanggan:** Pahami sentimen pelanggan dan temukan keluhan utama.
    * **💡 Rekomendasi:** Dapatkan rekomendasi *cross-sell* (JIKA beli A, TAWARKAN B).
    * **💸 Simulator Promo:** Hitung profitabilitas skenario diskon atau BOGO.
    * **✨ Analisis Musiman:** Pahami pola penjualan saat Hari Libur Nasional vs Hari Biasa.
    * **🍔 Strategi Menu:** Temukan menu 'Juara' di setiap musim untuk optimasi stok.
    * **🧪 Lab Strategi:** Tentukan nasib menu (Stars/Dogs) dan simulasikan perubahan harga.
    """
    )
    st.markdown("---")


def build_footer():
    "Membangun footer aplikasi."
    st.markdown(
        """
        <div id="custom-footer">
            <div>
                Data Driven F&B Analyst Dashboard © 2025
            </div>
            <div class="links">
                Developer © ronihidayat
                <a href="https://api.whatsapp.com/message/542JTLNDT3HCO1?autoload=1&app_absent=0" target="_blank">Contact Me</a>
                <a href="https://www.linkedin.com/in/roni-hidayat0692/" target="_blank">LinkedIn</a>
                <a href="https://github.com/RONI1920" target="_blank">GitHub</a>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def main():
    """Fungsi utama untuk menjalankan aplikasi Streamlit."""

    # Inisialisasi Database
    init_db()
    load_css("style.css")

    # #############################################################
    # --- 1. INISIALISASI SESSION STATE ---
    # #############################################################
    # ... (Biarkan kode session state Anda yang panjang itu, tidak perlu diubah) ...
    # Pastikan state untuk P&L ada:
    if "save_pl_flag" not in st.session_state:
        st.session_state.save_pl_flag = False
    if "pl_saved_status" not in st.session_state:
        st.session_state.pl_saved_status = False

    # #############################################################
    # --- 2. SIDEBAR ---
    # #############################################################
    (
        gmv_file,
        cogs_file,
        waiter_file,
        ulasan_file,
        purchase_file,
        pl_file,  # <--- Pastikan ini diterima dari sidebar
        use_db,
    ) = build_sidebar()

    # #############################################################
    # --- 3. PEMUATAN DATA ---
    # #############################################################

    # 1. Load GMV
    data_gmv, file_company, file_period, file_branch = load_data_gmv(gmv_file, use_db)

    # 2. Inisialisasi variabel data_pl DULU biar tidak error "Not Defined"
    data_pl = None  # <--- (PENTING! JANGAN DIHAPUS)

    with st.spinner("Memuat data pendukung..."):
        data_cogs = load_cogs_data(cogs_file, use_db)
        data_waiter = load_data_waiter(waiter_file, use_db)
        data_ulasan = load_data_ulasan(ulasan_file, use_db)
        data_purchase = load_data_purchase(purchase_file, use_db)

        # 3. Load Data P&L (Isi variabel data_pl di sini)
        data_pl = load_pl_data(pl_file, use_db)  # <--- (PENTING! PEMANGGILAN FUNGSI)

        # Muat data kalender
        data_kalender = load_kalender_data("kalender/kalender_event1.csv")

    # #############################################################
    # --- 4. LOGIKA PENYIMPANAN DATA ---
    # #############################################################

    def clear_db_cache_and_rerun():
        load_dataframe_from_db.clear()
        st.session_state.use_db = True
        st.rerun()

    # ... (Kode simpan GMV, COGS, Waiter, Ulasan, Purchase tetap sama) ...

    # Logika Simpan P&L (BARU)
    if st.session_state.save_pl_flag:
        if data_pl is not None:
            save_dataframe_smart_append(data_pl, "pl_data", "Date")
            st.session_state.pl_saved_status = True
            st.session_state.save_pl_flag = False
            clear_db_cache_and_rerun()

    # #############################################################
    # --- 5. FILTER GLOBAL ---
    # #############################################################
    (
        filtered_gmv,
        filtered_cogs,
        filtered_waiter,
        filtered_purchase,
        filtered_pl,
    ) = build_global_filters(
        data_gmv,
        data_cogs,
        data_waiter,
        data_purchase,
        data_pl,
    )

    # #################################################################
    # --- 6. RENDER KONTEN UTAMA ---
    # #################################################################

    # Hitung Total Revenue untuk FCP
    total_sales_revenue = 0
    if filtered_gmv is not None and not filtered_gmv.empty:
        # ... (Logika hitung revenue tetap sama) ...
        revenue_col = None
        if "Total Nett Sales" in filtered_gmv.columns:
            revenue_col = "Total Nett Sales"
        elif "Net Sales" in filtered_gmv.columns:
            revenue_col = "Net Sales"
        elif "Total Gross Sales" in filtered_gmv.columns:
            revenue_col = "Total Gross Sales"

        if revenue_col:
            filtered_gmv[revenue_col] = pd.to_numeric(
                filtered_gmv[revenue_col], errors="coerce"
            ).fillna(0)
            total_sales_revenue = filtered_gmv[revenue_col].sum()

    # --- LOGIKA PENENTU (Welcome Screen vs Dashboard) ---
    # Sekarang data_pl sudah didefinisikan di atas, jadi aman dipakai di sini.
    all_data_is_missing = (
        data_gmv is None
        and data_cogs is None
        and data_waiter is None
        and data_ulasan is None
        and data_purchase is None
        and data_pl
        is None  # <--- Tidak akan error lagi karena data_pl sudah ada nilainya (None atau DataFrame)
    )

    # OPSI A: TIDAK ADA DATA SAMA SEKALI -> Tampilkan Welcome Screen
    if all_data_is_missing:
        build_welcome_screen()
        build_footer()
        st.stop()

    # OPSI B: ADA DATA -> Tampilkan Dashboard

    # ... (Kode Header Dinamis tetap sama) ...
    # Tampilkan Header standar jika GMV kosong tapi file lain ada
    if not all_data_is_missing and data_gmv is None:
        st.title("Dashboard Analisis Data F&B")
        st.info(
            "Mode analisis parsial. Beberapa fitur mungkin nonaktif karena data GMV tidak ada."
        )

    # #############################################################
    # --- 7. NAVIGASI HALAMAN ---
    # #############################################################
    st.divider()

    page_options = [
        "📊 Penjualan (GMV)",
        "💰 COGS & Profit",
        "🧑‍🍳 SDM & Waktu Sibuk",
        "🛒 Pembelian",
        "⚖️ A/B Comparison",
        "🎯 Target",
        "🔮 Forecast (AI)",
        "❤️ Ulasan",
        "💡 Rekomendasi",
        "💸 Analisis Promo",
        "✨ Analisis Musiman Tahunan",
        "🧪 Lab Strategi",
        "📉 Laporan Laba Rugi (P&L)",  # <--- Menu Baru
    ]

    page = st.selectbox("Pilih Halaman Analisis:", page_options, key="nav_select")
    st.divider()

    # #############################################################
    # --- 8. RENDER HALAMAN ---
    # #############################################################

    if page == "📊 Penjualan (GMV)":
        build_tab1_sales(filtered_gmv)
    elif page == "💰 COGS & Profit":
        build_tab2_cogs(filtered_cogs)
    elif page == "🧑‍🍳 SDM & Waktu Sibuk":
        build_tab3_hr(filtered_waiter)
    elif page == "🛒 Pembelian":
        build_tab8_purchase(filtered_purchase, total_sales_revenue)
    elif page == "⚖️ A/B Comparison":
        build_tab4_comparison(filtered_gmv, filtered_cogs, filtered_waiter)
    elif page == "🎯 Target":
        build_tab6_target(filtered_gmv)
    elif page == "🔮 Forecast (AI)":
        build_tab5_forecast(filtered_gmv)
    elif page == "❤️ Ulasan":
        build_tab7_ulasan(data_ulasan)
    elif page == "💡 Rekomendasi":
        build_tab9_rekomendasi(filtered_gmv)
    elif page == "💸 Analisis Promo":
        build_tab10_promo(filtered_gmv, filtered_cogs)
    elif page == "✨ Analisis Musiman Tahunan":
        build_tab11_musiman(filtered_gmv, data_kalender)
    elif page == "🧪 Lab Strategi":
        build_tab_unique_features(data_gmv, data_cogs)
    elif page == "📉 Laporan Laba Rugi (P&L)":
        build_tab13_pl(data_pl)  # <--- Panggil Fungsi UI Baru

    # --- 9. FOOTER & SKRIP LAINNYA ---
    build_footer()

    st.markdown(
        """
        <script>
        function hideAllNotices() {
            const notices = document.querySelectorAll(
                '.stAlert[data-baseweb="alert"]:not(.fading-out)'
            );

            notices.forEach(function(notice) {
                if (notice.className.includes("stAlert-success") || notice.className.includes("stAlert-info") || notice.className.includes("stAlert-warning")) {
                    notice.classList.add('fading-out');

                    setTimeout(() => {
                        notice.style.transition = 'opacity 0.5s ease-out';
                        notice.style.opacity = '0';
                        
                        setTimeout(() => {
                            notice.style.display = 'none';
                        }, 500);
                    }, 5000); 
                }
            });
        }
        setTimeout(hideAllNotices, 1000); 
        </script>
        """,
        unsafe_allow_html=True,
    )
    # --- Akhir dari fungsi main() ---


# #################################################################
# --- ENTRY POINT APLIKASI ---
# #################################################################

if __name__ == "__main__":
    main()
