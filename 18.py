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


# --- CSS KUSTOM UNTUK TAB ---
st.markdown(
    """
<style>
.stTabs [data-baseweb="tab-list"] button {
    font-size: 1.5rem;  /* Ukuran font tab sudah besar */
    padding: 10px 15px;
}
</style>
""",
    unsafe_allow_html=True,
)
# -----------------------------------------------------------------

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

DB_FILE = "fnb_analyst_data.db"


def get_db_connection():
    """Membuat koneksi ke database SQLite."""
    conn = sqlite3.connect(DB_FILE)
    return conn


def init_db():
    """
    Membuat skema tabel database yang BENAR jika belum ada.
    Ini adalah pondasi agar 'Smart Append' berfungsi.
    """
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        # 1. Skema Tabel GMV
        # Disesuaikan dengan fungsi load_data_gmv
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS gmv_data (
            "Sales Date In" DATETIME,
            "Sales Date Out" DATETIME,
            "Bill Number" TEXT,
            "Menu" TEXT,
            "Qty" REAL,
            "Price (Net)" REAL,
            "Service Charge" REAL,
            "Tax" REAL,
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

        # 2. Skema Tabel COGS
        # Disesuaikan dengan fungsi load_cogs_data
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
        # Disesuaikan dengan fungsi load_data_waiter
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS waiter_data (
            "Bill Number" TEXT,
            "Waiter" TEXT,
            "Order Time" DATETIME,
            "Total After Bill Discount" REAL,
            "Branch" TEXT
        );
        """
        )

        # 4. Skema Tabel Ulasan
        # Disesuaikan dengan fungsi load_data_ulasan
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
        # Disesuaikan dengan fungsi load_data_purchase
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
        # 'replace' akan drop tabel jika ada dan membuatnya kembali
        df.to_sql(table_name, conn, if_exists="replace", index=False)
        conn.commit()
        conn.close()
        # Kita buat pesan suksesnya lebih 'silent' agar tidak bentrok
        # st.success(f"Data {table_name} berhasil disimpan ke database!")
        print(f"Data {table_name} (mode replace) berhasil disimpan.")
    except Exception as e:
        st.error(f"Gagal menyimpan data {table_name} ke DB: {e}")


def save_dataframe_smart_append(df, table_name, date_col_name):
    """
    Menyimpan DataFrame ke DB dengan strategi "Smart Append".
    Menghapus data duplikat berdasarkan rentang tanggal, lalu menambahkan data baru.
    """
    if df is None or df.empty:
        st.error(f"Dataframe untuk {table_name} kosong, tidak ada yang disimpan.")
        return

    # 1. Pastikan kolom tanggal ada
    if date_col_name not in df.columns:
        st.error(
            f"Kolom tanggal '{date_col_name}' tidak ditemukan di data. Gagal menyimpan."
        )
        return

    # 2. Konversi ke datetime (jika belum) dan cari rentang tanggal
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

        # 3. HAPUS data lama yang tumpang tindih
        # Ini adalah bagian "Smart"-nya
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
        df.to_sql(table_name, conn, if_exists="append", index=False)

        conn.commit()
        st.success(
            f"Sukses! {len(df)} baris data baru untuk {table_name} berhasil disimpan ke database."
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


@st.cache_data(show_spinner=False)  # Tambahkan ini untuk cache
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
def get_menu_performance(df):
    """
    Menganalisis performa menu dan kategori.
    VERSI HYBRID: Mengembalikan data untuk Top/Bottom 10 DAN data mentah
    untuk drill-down interaktif.
    """

    # --- Filter awal (tetap sama) ---
    menu_sales = df[~df["Menu"].str.contains("PACKAGE", na=False, case=False)]
    menu_sales = menu_sales[menu_sales["Price (Net)"] > 0]

    # #################################################################
    # --- PERBAIKAN FINAL: FILTER GANDA (Jaring Pengaman) ---

    # Definisikan filter regex kita SEKALI
    filter_regex = r"ADD[ -]?ON|ADDITIONAL|Level"

    # Filter Kolom 1: 'Menu Category' (untuk Minuman "Additional")
    if "Menu Category" in df.columns:
        menu_sales = menu_sales[
            ~menu_sales["Menu Category"].str.contains(
                filter_regex, na=False, case=False, regex=True
            )
        ]

    # Definisikan nama kolom detail (ini sudah benar dari history kita)
    NAMA_KOLOM_DETAIL = "Menu Category Detail"

    # Filter Kolom 2: 'Menu Kategori Detail' (untuk Makanan "ADD ON Nori")
    if NAMA_KOLOM_DETAIL in df.columns:
        menu_sales = menu_sales[
            ~menu_sales[NAMA_KOLOM_DETAIL].str.contains(
                filter_regex, na=False, case=False, regex=True
            )
        ]

    # --- BATAS PERBAIKAN ---
    # #################################################################

    # #################################################################
    # <-- PERBAIKAN DI SINI: Filter Nama Item Spesifik (Ocha, dll.) -->
    #
    # Filter item spesifik yang tidak ingin dianggap sebagai menu utama
    # Kita gunakan regex dengan | (ATAU)
    filter_regex_items = r"Ocha|Refill|Mineral Water"

    if "Menu" in menu_sales.columns:
        menu_sales = menu_sales[
            ~menu_sales["Menu"].str.contains(
                filter_regex_items, na=False, case=False, regex=True
            )
        ]
    # <-- BATAS PERBAIKAN -->
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
    # menu_sales di sini sudah bersih dari 'Ocha' dan 'Mineral Water'
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
@st.cache_data(show_spinner=False)
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

    # --- BLOK 1: Logika Pemuatan DB (DIPERBAIKI) ---
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

            # --- PERBAIKAN: Kembalikan 'DB_MODE' agar header dinamis ---
            return df, "DB_MODE", "DB_MODE", "DB_MODE"

    # --- BLOK 2: Logika Asli (jika file di-upload) ---
    if uploaded_file is None:
        return None, None, None, None

    # ... (Sisa fungsi ini tidak perlu diubah) ...
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
            st.error(
                f"Format file {uploaded_file.name} tidak didukung. Harap upload .xlsx atau .csv"
            )
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

    # PERBAIKAN TAMBAHAN: Pastikan semua kolom tanggal di-load
    if "Order Time" in df_data.columns:
        df_data["Order Time"] = pd.to_datetime(df_data["Order Time"], errors="coerce")

    df_data["Company"] = company_name
    df_data["Period"] = period_str
    df_data["Branch"] = branch_name_header

    if "Bill Number" not in df_data.columns or "Sales Date In" not in df_data.columns:
        st.warning(
            "Peringatan: Kolom 'Bill Number' atau 'Sales Date In' tidak ditemukan setelah pembersihan."
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

    # ... (Sisa fungsi ini tidak perlu diubah, biarkan apa adanya) ...
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


@st.cache_data
def load_cogs_data(uploaded_file, use_db=False):
    """Memuat dan membersihkan data COGS (File 2)."""

    # --- BLOK BARU: Logika Pemuatan DB ---
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
            return df  # Kembalikan data dari DB

    # --- Logika Asli (jika file di-upload) ---
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

    # Cek kolom Branch
    if "Branch" not in df.columns:
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

    # Normalisasi kolom cabang jika ada
    if "Branch" in df.columns:
        df["Branch"] = df["Branch"].astype(str).str.strip().str.title()

    return df


@st.cache_data
def load_data_waiter(uploaded_file, use_db=False):
    """Memuat dan membersihkan data Waiter (File 3)."""

    # --- BLOK BARU: Logika Pemuatan DB ---
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
            return df  # Kembalikan data dari DB

    # --- Logika Asli (jika file di-upload) ---
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

    # Cek kolom Branch
    if "Branch" not in df.columns:
        st.warning(
            "Peringatan: Kolom 'Branch' tidak ditemukan di File Waiter. Filter cabang mungkin tidak berfungsi."
        )

    df["Order Time"] = pd.to_datetime(df["Order Time"], errors="coerce")
    df["Total After Bill Discount"] = pd.to_numeric(
        df["Total After Bill Discount"], errors="coerce"
    ).fillna(0)
    df.dropna(subset=["Bill Number", "Order Time"], inplace=True)

    # Normalisasi kolom cabang jika ada
    if "Branch" in df.columns:
        df["Branch"] = df["Branch"].astype(str).str.strip().str.title()

    return df


@st.cache_data
def load_data_ulasan(uploaded_file, use_db=False):
    """Memuat dan membersihkan data Ulasan (File 4)."""

    # --- BLOK BARU: Logika Pemuatan DB ---
    if use_db and uploaded_file is None:
        with st.spinner("Memuat data Ulasan dari database..."):
            numeric_config = {"Rating_Clean": "int"}
            df = load_dataframe_from_db(
                "ulasan_data", numeric_cols_config=numeric_config
            )
            if df is None:
                st.info("Database Ulasan kosong. Silakan upload file baru.", icon="ℹ️")
                return None
            return df  # Kembalikan data dari DB

    # --- Logika Asli (jika file di-upload) ---
    if uploaded_file is None:
        return None

    df = None
    try:
        # --- LOGIKA DIPERBARUI UNTUK EXCEL & CSV ---
        if uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        elif uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, encoding="latin1")
        else:
            st.error(
                "Format file Ulasan (File 4) tidak didukung. Harap upload .xlsx atau .csv"
            )
            return None
        # --- BATAS PERUBAHAN ---

    except Exception as e:
        st.error(f"Error membaca file Ulasan (File 4): {e}")
        st.error(
            "Pastikan file adalah .csv atau .xlsx dengan kolom: Nama, Rating, Ulasan"
        )
        return None

    # --- Pembersihan Data Ulasan ---
    if "Rating" not in df.columns or "Ulasan" not in df.columns:
        st.error("File Ulasan (File 4) harus memiliki kolom 'Rating' dan 'Ulasan'.")
        return None

    # Membersihkan kolom Rating (cth: "5 bintang" -> 5)
    # Menggunakan regex untuk mengekstrak angka pertama yang ditemukan

    # --- PERBAIKAN DARI ERROR SEBELUMNYA ---
    # Menggunakan .str.extract()
    df["Rating_Clean"] = (
        df["Rating"].astype(str).str.extract(r"(\d+)").fillna(0).astype(int)
    )
    # --- BATAS PERBAIKAN ---

    # Menghapus baris yang tidak memiliki ulasan atau rating bersih
    df.dropna(subset=["Ulasan"], inplace=True)
    df = df[df["Rating_Clean"] > 0]  # Hanya ambil yg punya rating valid

    df["Ulasan"] = df["Ulasan"].astype(str)

    return df


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

    def uncheck_db_on_upload():
        """
        Callback untuk menonaktifkan DB JIKA file baru di-upload.
        DAN me-reset status 'tersimpan'.
        """
        st.session_state.use_db = False

        # Reset semua status saat ada file baru
        st.session_state.gmv_saved_status = False
        st.session_state.cogs_saved_status = False
        st.session_state.waiter_saved_status = False
        st.session_state.ulasan_saved_status = False
        st.session_state.purchase_saved_status = False

    # --- BATAS FUNGSI CALLBACK ---

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
            st.session_state.use_db = True

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
            on_change=uncheck_db_on_upload,
        )
        if gmv_file is not None:
            if st.session_state.gmv_saved_status:
                # --- PERBAIKAN DI SINI ---
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
            on_change=uncheck_db_on_upload,
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
            on_change=uncheck_db_on_upload,
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
            on_change=uncheck_db_on_upload,
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
            on_change=uncheck_db_on_upload,
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

    # Kembalikan file DAN status checkbox
    return (
        gmv_file,
        cogs_file,
        waiter_file,
        ulasan_file,
        purchase_file,
        st.session_state.use_db,  # Kita tetap kembalikan var KONTROL kita
    )


def build_global_filters(data_gmv, data_cogs, data_waiter, data_purchase):
    """Menggambar filter global dan mengembalikan data yang sudah difilter."""

    # Inisialisasi data yang difilter sebagai data asli
    filtered_gmv = data_gmv
    filtered_cogs = data_cogs
    filtered_waiter = data_waiter
    filtered_purchase = data_purchase

    # #############################################################
    # --- PERBAIKAN 1: Inisialisasi Variabel Scope Aman ---
    # #############################################################
    master_min_date = pd.Timestamp.now().date()
    master_max_date = pd.Timestamp.now().date()
    filter_source_file = None  # <--- Perbaikan 1 (A)
    selected_branches = []
    selected_date = master_max_date  # <--- Perbaikan 2 (D)

    try:
        if data_gmv is not None and not data_gmv.empty:
            master_min_date = data_gmv["Sales Date In"].min().date()
            master_max_date = data_gmv["Sales Date In"].max().date()
            filter_source_file = "GMV"
        elif data_cogs is not None and not data_cogs.empty:
            master_min_date = data_cogs["Sales Date"].min().date()
            master_max_date = data_cogs["Sales Date"].max().date()
            filter_source_file = "COGS"
        elif data_waiter is not None and not data_waiter.empty:
            master_min_date = data_waiter["Order Time"].min().date()
            master_max_date = data_waiter["Order Time"].max().date()
            filter_source_file = "Waiter"
        elif data_purchase is not None and not data_purchase.empty:
            master_min_date = data_purchase["Purchase Date"].min().date()
            master_max_date = data_purchase["Purchase Date"].max().date()
            filter_source_file = "Purchase"
    except Exception as e:
        st.error(f"Gagal membaca rentang tanggal: {e}")

    # #############################################################
    # --- BLOK FILTER CABANG & TANGGAL (DIMULAI DARI IF INI) ---
    # #############################################################
    if filter_source_file:  # <--- Perbaikan 1 (B)
        st.subheader("Filter Analisis Global")
        # st.info(
        #     f"Filter global saat ini menggunakan rentang tanggal dari file: **{filter_source_file}**"
        # )

        # --- MODIFIKASI: BLOK FILTER CABANG BARU ---

        # Branch Filter hanya jika ada data GMV (sebagai master list)
        if data_gmv is not None and "Branch" in data_gmv.columns:
            all_branches = sorted(data_gmv["Branch"].unique())
            selected_branches = st.multiselect(
                "Pilih Cabang (Branch):",
                options=all_branches,
                default=all_branches,
                key="branch_filter",
            )

            # Terapkan filter branch ke GMV
            filtered_gmv = data_gmv[data_gmv["Branch"].isin(selected_branches)]

            # >>> TAMBAHAN UNTUK DATA LAIN (Filter Cabang) <<<
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

        # --- BATAS MODIFIKASI BLOK FILTER CABANG ---

        # Terapkan filter radio ke semua tab
        filter_type = st.radio(
            "Pilih rentang waktu (Tab 1, 2, 3, 8):",
            ["Semua Periode", "Harian", "Mingguan", "Bulanan", "Tahunan"],
            horizontal=True,
            key="filter_type_global",
        )

        # --- Terapkan filter tanggal ke 'filtered_...' (Lanjutan) ---
        if filter_type == "Harian":
            selected_date = st.date_input(  # <--- Perbaikan 2 (A)
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
            if data_cogs is not None:
                filtered_cogs = (
                    filtered_cogs[  # Gunakan filtered_cogs (sudah difilter cabang)
                        filtered_cogs["Sales Date"].dt.date == selected_date
                    ]
                )
            if data_waiter is not None:
                filtered_waiter = filtered_waiter[  # Gunakan filtered_waiter
                    filtered_waiter["Order Time"].dt.date == selected_date
                ]
            if data_purchase is not None:
                filtered_purchase = filtered_purchase[  # Gunakan filtered_purchase
                    filtered_purchase["Purchase Date"].dt.date == selected_date
                ]

        elif filter_type == "Mingguan":
            default_start_mingguan = master_max_date - pd.to_timedelta(6, unit="d")
            if default_start_mingguan < master_min_date:
                default_start_mingguan = master_min_date

            selected_start_date = st.date_input(
                "Pilih Tanggal Mulai (periode 7 hari)",
                value=default_start_mingguan,
                min_value=master_min_date,
                max_value=master_max_date,
                key="global_week_start",
            )
            start_date = selected_start_date
            end_date = start_date + pd.to_timedelta(6, unit="d")
            if end_date > master_max_date:
                end_date = master_max_date
            st.info(
                f"Menampilkan data dari {start_date.strftime('%d-%m-%Y')} s.d. {end_date.strftime('%d-%m-%Y')}"
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

        elif filter_type == "Bulanan":
            selected_date = st.date_input(  # <--- Perbaikan 2 (B)
                "Pilih tanggal dalam bulan",
                value=master_max_date,
                min_value=master_min_date,
                max_value=master_max_date,
                key="global_month",
            )
            selected_month = selected_date.month
            selected_year = selected_date.year
            st.info(f"Menampilkan data untuk bulan {selected_date.strftime('%B %Y')}")

            if filtered_gmv is not None:
                filtered_gmv = filtered_gmv[
                    (filtered_gmv["Sales Date In"].dt.month == selected_month)
                    & (filtered_gmv["Sales Date In"].dt.year == selected_year)
                ]
            if filtered_cogs is not None:
                filtered_cogs = filtered_cogs[
                    (filtered_cogs["Sales Date"].dt.month == selected_month)
                    & (filtered_cogs["Sales Date"].dt.year == selected_year)
                ]
            if filtered_waiter is not None:
                filtered_waiter = filtered_waiter[
                    (filtered_waiter["Order Time"].dt.month == selected_month)
                    & (filtered_waiter["Order Time"].dt.year == selected_year)
                ]
            if filtered_purchase is not None:
                filtered_purchase = filtered_purchase[
                    (filtered_purchase["Purchase Date"].dt.month == selected_month)
                    & (filtered_purchase["Purchase Date"].dt.year == selected_year)
                ]

        elif filter_type == "Tahunan":
            selected_date = st.date_input(  # <--- Perbaikan 2 (C)
                "Pilih tanggal dalam tahun",
                value=master_max_date,
                min_value=master_min_date,
                max_value=master_max_date,
                key="global_year",
            )
            selected_year = selected_date.year
            st.info(f"Menampilkan data untuk tahun {selected_year}")

            if filtered_gmv is not None:
                filtered_gmv = filtered_gmv[
                    filtered_gmv["Sales Date In"].dt.year == selected_year
                ]
            if filtered_cogs is not None:
                filtered_cogs = filtered_cogs[
                    filtered_cogs["Sales Date"].dt.year == selected_year
                ]
            if filtered_waiter is not None:
                filtered_waiter = filtered_waiter[
                    filtered_waiter["Order Time"].dt.year == selected_year
                ]
            if filtered_purchase is not None:
                filtered_purchase = filtered_purchase[
                    filtered_purchase["Purchase Date"].dt.year == selected_year
                ]

        # Jika filter_type == "Semua Periode", kita tidak melakukan apa-apa.

    elif (
        data_gmv is None
        and data_cogs is None
        and data_waiter is None
        and data_purchase is None
    ):
        st.markdown("---")

    # Kembalikan data yang SUDAH DIFILTER
    return filtered_gmv, filtered_cogs, filtered_waiter, filtered_purchase


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
            ) = get_menu_performance(filtered_gmv)

            # === 2. TAMPILKAN KPI UTAMA (Expander Asli Anda) ===
            with st.expander(
                "📈 KPI Kinerja Penjualan (Revenue, ATV, IPB)", expanded=True
            ):
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

                    # INI ADALAH GRAFIK ALTAIR KUSTOM ANDA, JADI KITA BIARKAN
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
                    # --- DIGANTI ---
                    st.plotly_chart(chart, use_container_width=True)
                with col12:
                    chart = create_horizontal_bar_chart(
                        top_grossing,
                        "Total Nett Sales",
                        "Menu",
                        "Total Nett Sales (Rp)",
                        "Menu Pendapatan Tertinggi (by Nett Sales)",
                    )
                    # --- DIGANTI ---
                    st.plotly_chart(chart, use_container_width=True)

            st.markdown("---")

            # === 5. URUTAN BARU 4: KPI OPERASIONAL ===
            st.header("⚙️ Analisis Operasional & Pelanggan")

            with st.expander(
                "⚙️ KPI Operasional (Jam Sibuk, Hari Sibuk, Durasi Makan)", expanded=True
            ):
                # Panggil fungsi kpi operasional di sini
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
                    # --- DIGANTI ---
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
                    # --- DIGANTI ---
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
                    # --- DIGANTI ---
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
                    # --- DIGANTI ---
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

            # #############################################################
            # --- BATAS BLOK INSIGHT BARU ---
            # #############################################################

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
                f"** busiest Waktu Tersibuk (Transaksi):** Sesi **{peak_time_trx['Waktu Kunjungan']}** "
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
    total_cost, cost_by_category, cost_by_supplier, top_items
):
    """
    Menganalisis data pembelian yang sudah diproses dan menghasilkan insight
    dalam bahasa alami untuk Tab 8.
    """
    insights = []

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


def build_tab3_hr(filtered_waiter):
    """Menggambar semua elemen untuk Tab 3.
    (VERSI BARU DENGAN KOTAK INSIGHT DI BAWAH)
    """
    if filtered_waiter is not None:
        if not filtered_waiter.empty:
            st.header("🧑‍🍳 Analisis Kinerja Waiter & Waktu Kunjungan")

            # 1. Hitung data analisis (sudah di-cache)
            time_data = get_peak_time_analysis(filtered_waiter)
            waiter_data = get_waiter_performance(filtered_waiter)

            # 2. Expander Waktu Kunjungan (tidak berubah)
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
                        x_type="O",
                        sort_order=sort_order_time,
                    )
                    # --- DIGANTI ---
                    st.plotly_chart(chart1, use_container_width=True)
                with t_col2:
                    st.markdown("##### Berdasarkan Total Penjualan")
                    chart2 = create_vertical_bar_chart(
                        time_data,
                        "Waktu Kunjungan",
                        "Total_Penjualan",
                        "Waktu Kunjungan",
                        "Total Penjualan (Rp)",
                        x_type="O",
                        sort_order=sort_order_time,
                    )
                    # --- DIGANTI ---
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

            # 3. Expander Performa Waiter (tidak berubah)
            with st.expander("🏆 Performa Waiter Teratas (Top 10)", expanded=True):
                st.subheader("🏆 Performa Waiter Teratas (Top 10)")
                chart_waiter = create_horizontal_bar_chart(
                    waiter_data,
                    "Total_Penjualan",
                    "Waiter",
                    "Total Penjualan (Rp)",
                    "Performa Waiter Teratas (by Penjualan)",
                )
                # --- DIGANTI ---
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

            # #############################################################
            # --- BLOK INSIGHT BARU DITEMPATKAN DI SINI (PALING BAWAH) ---
            # #############################################################

            st.markdown("---")  # Tambahkan pemisah visual
            st.header("💡 Insight Otomatis (SDM & Waktu)")

            # Panggil fungsi 'pencari insight' kita
            # Kita gunakan data yang sudah dihitung di awal tab ini
            insights = generate_hr_insights(time_data, waiter_data)

            # Tampilkan dalam expander baru
            with st.expander(
                "Klik untuk melihat Temuan Kunci dari Data SDM & Waktu", expanded=True
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

        elif filtered_waiter is not None and filtered_waiter.empty:
            st.warning(
                "Tidak ada data ditemukan di File Rekapitulasi untuk rentang waktu yang dipilih."
            )
    else:
        st.info(
            "Silakan upload file Laporan Rekapitulasi Detail (File 3) di sidebar untuk melihat analisis waiter."
        )


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
            default_B_date = default_A_date - pd.DateOffset(months=1).date()
        elif comparison_type == "Tahunan":
            default_B_date = default_A_date - pd.DateOffset(years=1).date()

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

        penjualan_saat_ini = data_bulan_aktif["Total After Bill Discount"].sum()
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

        if pd.isna(avg_sales_weekday) or avg_sales_weekday == 0:
            avg_sales_weekday = 1
        if pd.isna(avg_sales_weekend) or avg_sales_weekend == 0:
            # Jika weekend 0, setel ke weekday agar tidak terjadi weight 0
            avg_sales_weekend = avg_sales_weekday

        kpi_insight_dict["avg_sales_weekday"] = avg_sales_weekday  # -> Simpan
        kpi_insight_dict["avg_sales_weekend"] = avg_sales_weekend  # -> Simpan

        weekend_weight = avg_sales_weekend / avg_sales_weekday

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
            if pembagi > 0:
                rdr_weekday = sales_dibutuhkan / pembagi
                rdr_weekend = rdr_weekday * weekend_weight

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
        value=f"{pencapaian_persen*100:,.1f}%",  # <-- Perbaikan kecil di sini
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
            delta=f"{pencapaian_persen*100:,.1f}% dari Target",  # <-- Perbaikan kecil
            delta_color=status_color,
        )
        col3.metric(
            label="Selisih Target",
            value=format_rupiah(penjualan_saat_ini - target_bulanan),
            help="Hasil akhir penjualan dikurangi target bulanan.",
        )

    st.subheader("🔮 Perbandingan Model Ramalan")
    st.write(
        "Membandingkan Proyeksi Cerdas (berbasis rata-rata) dengan Proyeksi Prophet (berbasis tren & musiman)."
    )

    with st.container(border=True):
        col_p1, col_p2 = st.columns(2)

        if target_bulanan > 0:
            proyeksi_prophet_persen = proyeksi_prophet / target_bulanan
        else:
            proyeksi_prophet_persen = 0

        if proyeksi_prophet_persen > 1.05:
            prophet_color = "normal"
        elif proyeksi_prophet_persen >= 0.98:
            prophet_color = "off"
        else:
            prophet_color = "inverse"

        col_p1.metric(
            label="Proyeksi Akhir Bulan (Cerdas)",
            value=format_rupiah(proyeksi_akhir_bulan),
            delta=f"{proyeksi_vs_target_persen * 100:,.1f}% dari Target",
            delta_color=status_color,
            help="Estimasi berdasarkan performa rata-rata weekday & weekend Anda.",
        )
        col_p2.metric(
            label="Proyeksi Akhir Bulan (Model Prophet)",
            value=format_rupiah(proyeksi_prophet),
            delta=f"{proyeksi_prophet_persen * 100:,.1f}% dari Target",
            delta_color=prophet_color,
            help="Estimasi dari model AI (Prophet) berdasarkan tren & pola mingguan.",
        )

    st.markdown("---")
    st.subheader("💡 Rencana Aksi & Diagnostik")

    col_a, col_b = st.columns(2)
    if sisa_hari > 0:

        # #############################################################
        # --- PERBAIKAN SINTAKS DI SINI ---
        # #############################################################

        col_a.metric(
            label="Target Harian Sisa (Weekday)",
            value=format_rupiah(rdr_weekday),
            delta=f"vs Rata-rata: {format_rupiah(avg_sales_weekday)}",
            delta_color=(
                "inverse"
                if rdr_weekday > avg_sales_weekday
                else ("normal" if rdr_weekday > 0 else "off")
            ),
            help="Penjualan harian (Sen-Kam) yang dibutuhkan untuk mencapai target.",
        )

        # #############################################################
        # --- BATAS PERBAIKAN ---
        # #############################################################

        with col_b:
            col_b.metric(
                label="Target Harian Sisa (Weekend)",
                value=format_rupiah(rdr_weekend),
                delta=f"vs Rata-rata: {format_rupiah(avg_sales_weekend)}",
                delta_color=(
                    "inverse"
                    if rdr_weekend > avg_sales_weekend
                    else ("normal" if rdr_weekend > 0 else "off")
                ),
                help="Penjualan harian (Jum-Min) yang dibutuhkan untuk mencapai target.",
            )
    else:
        # Jika bulan sudah selesai
        col_a.metric("Rata-rata Weekday (Final)", format_rupiah(avg_sales_weekday))

    st.markdown("---")
    st.subheader(f"📈 Tren Penjualan Harian - {active_month_name}")

    # Buat grafik Plotly untuk tren harian
    fig_daily_trend = px.line(
        daily_sales_agg,
        x="Sales Date In",
        y="Total After Bill Discount",
        color="Tipe Hari",
        title=f"Tren Penjualan Harian - {active_month_name}",
        labels={
            "Sales Date In": "Tanggal",
            "Total After Bill Discount": "Total Penjualan (Rp)",
            "Tipe Hari": "Tipe Hari",
        },
        template="plotly_white",
        markers=True,
    )
    fig_daily_trend.update_layout(
        title_x=0.01,
        yaxis_tickformat=".2s",
        legend_title_text="",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig_daily_trend.update_traces(
        hovertemplate="<b>%{x|%d %B}</b><br>Penjualan: %{y:,.0f}<extra></extra>"
    )
    st.plotly_chart(fig_daily_trend, use_container_width=True)

    # Grafik Rata-rata Tipe Hari
    avg_sales_df = (
        daily_sales_agg.groupby("Tipe Hari")["Total After Bill Discount"]
        .mean()
        .reset_index()
    )
    fig_avg_type = create_vertical_bar_chart(
        avg_sales_df,
        "Tipe Hari",
        "Total After Bill Discount",
        "Tipe Hari",
        "Rata-rata Penjualan (Rp)",
        x_type="O",
        sort_order=["Weekday (Sen-Kam)", "Weekend (Jum-Min)"],
    )
    st.plotly_chart(fig_avg_type, use_container_width=True)

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


def build_tab8_purchase(filtered_purchase):
    """Menggambar Tab 8 (Analisis Laporan Pembelian)."""
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

    # --- 2. Tampilkan Metrik KPI ---
    st.subheader("📊 KPI Biaya Pembelian")
    st.metric(
        "Total Biaya Pembelian (Tercatat)",
        format_rupiah(total_cost),
        help="Total dari kolom 'Total' di mana nilainya > 0.",
    )

    # --- 3. Tampilkan Grafik Analisis ---
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

    # --- 4. Tampilkan Data Mentah ---
    with st.expander("Lihat Rincian Data Pembelian (Sudah Difilter)"):
        st.dataframe(
            raw_data_filtered.style.format(
                {"Price": format_rupiah, "Total": format_rupiah}
            ),
            use_container_width=True,
        )

    # --- 5. Blok Insight (PALING BAWAH) ---
    st.markdown("---")
    st.header("💡 Insight Otomatis (Analisis Pembelian)")
    insights = generate_purchase_insights(
        total_cost, cost_by_category, cost_by_supplier, top_items
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
# --- FUNGSI BARU UNTUK TAB 9 (REKOMENDASI) ---
# #################################################################


@st.cache_data(show_spinner=False)
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
    # --- PERBAIKAN: Ganti 'rules_df' menjadi 'filtered_rules' ---
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
    # #############################################################
    # --- BATAS TAMBAHAN ---
    # #############################################################


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


# --- CSS KUSTOM UNTUK TAB ---
st.markdown(
    """
<style>
.stTabs [data-baseweb="tab-list"] button {
    font-size: 1.5rem;  /* Ukuran font tab sudah besar */
    padding: 10px 15px;
}
</style>
""",
    unsafe_allow_html=True,
)
# -----------------------------------------------------------------

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

DB_FILE = "fnb_analyst_data.db"


def get_db_connection():
    """Membuat koneksi ke database SQLite."""
    conn = sqlite3.connect(DB_FILE)
    return conn


def init_db():
    """
    Membuat skema tabel database yang BENAR jika belum ada.
    Ini adalah pondasi agar 'Smart Append' berfungsi.
    """
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        # 1. Skema Tabel GMV
        # Disesuaikan dengan fungsi load_data_gmv
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS gmv_data (
            "Sales Date In" DATETIME,
            "Sales Date Out" DATETIME,
            "Bill Number" TEXT,
            "Menu" TEXT,
            "Qty" REAL,
            "Price (Net)" REAL,
            "Service Charge" REAL,
            "Tax" REAL,
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

        # 2. Skema Tabel COGS
        # Disesuaikan dengan fungsi load_cogs_data
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
        # Disesuaikan dengan fungsi load_data_waiter
        cursor.execute(
            """
        CREATE TABLE IF NOT EXISTS waiter_data (
            "Bill Number" TEXT,
            "Waiter" TEXT,
            "Order Time" DATETIME,
            "Total After Bill Discount" REAL,
            "Branch" TEXT
        );
        """
        )

        # 4. Skema Tabel Ulasan
        # Disesuaikan dengan fungsi load_data_ulasan
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
        # Disesuaikan dengan fungsi load_data_purchase
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
        # 'replace' akan drop tabel jika ada dan membuatnya kembali
        df.to_sql(table_name, conn, if_exists="replace", index=False)
        conn.commit()
        conn.close()
        # Kita buat pesan suksesnya lebih 'silent' agar tidak bentrok
        # st.success(f"Data {table_name} berhasil disimpan ke database!")
        print(f"Data {table_name} (mode replace) berhasil disimpan.")
    except Exception as e:
        st.error(f"Gagal menyimpan data {table_name} ke DB: {e}")


def save_dataframe_smart_append(df, table_name, date_col_name):
    """
    Menyimpan DataFrame ke DB dengan strategi "Smart Append".
    Menghapus data duplikat berdasarkan rentang tanggal, lalu menambahkan data baru.
    """
    if df is None or df.empty:
        st.error(f"Dataframe untuk {table_name} kosong, tidak ada yang disimpan.")
        return

    # 1. Pastikan kolom tanggal ada
    if date_col_name not in df.columns:
        st.error(
            f"Kolom tanggal '{date_col_name}' tidak ditemukan di data. Gagal menyimpan."
        )
        return

    # 2. Konversi ke datetime (jika belum) dan cari rentang tanggal
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

        # 3. HAPUS data lama yang tumpang tindih
        # Ini adalah bagian "Smart"-nya
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
        df.to_sql(table_name, conn, if_exists="append", index=False)

        conn.commit()
        st.success(
            f"Sukses! {len(df)} baris data baru untuk {table_name} berhasil disimpan ke database."
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


@st.cache_data(show_spinner=False)  # Tambahkan ini untuk cache
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
def get_menu_performance(df):
    """
    Menganalisis performa menu dan kategori.
    VERSI HYBRID: Mengembalikan data untuk Top/Bottom 10 DAN data mentah
    untuk drill-down interaktif.
    """

    # --- Filter awal (tetap sama) ---
    menu_sales = df[~df["Menu"].str.contains("PACKAGE", na=False, case=False)]
    menu_sales = menu_sales[menu_sales["Price (Net)"] > 0]

    # #################################################################
    # --- PERBAIKAN FINAL: FILTER GANDA (Jaring Pengaman) ---

    # Definisikan filter regex kita SEKALI
    filter_regex = r"ADD[ -]?ON|ADDITIONAL|Level"

    # Filter Kolom 1: 'Menu Category' (untuk Minuman "Additional")
    if "Menu Category" in df.columns:
        menu_sales = menu_sales[
            ~menu_sales["Menu Category"].str.contains(
                filter_regex, na=False, case=False, regex=True
            )
        ]

    # Definisikan nama kolom detail (ini sudah benar dari history kita)
    NAMA_KOLOM_DETAIL = "Menu Category Detail"

    # Filter Kolom 2: 'Menu Kategori Detail' (untuk Makanan "ADD ON Nori")
    if NAMA_KOLOM_DETAIL in df.columns:
        menu_sales = menu_sales[
            ~menu_sales[NAMA_KOLOM_DETAIL].str.contains(
                filter_regex, na=False, case=False, regex=True
            )
        ]

    # --- BATAS PERBAIKAN ---
    # #################################################################

    # #################################################################
    # <-- PERBAIKAN DI SINI: Filter Nama Item Spesifik (Ocha, dll.) -->
    #
    # Filter item spesifik yang tidak ingin dianggap sebagai menu utama
    # Kita gunakan regex dengan | (ATAU)
    filter_regex_items = r"Ocha|Refill|Mineral Water"

    if "Menu" in menu_sales.columns:
        menu_sales = menu_sales[
            ~menu_sales["Menu"].str.contains(
                filter_regex_items, na=False, case=False, regex=True
            )
        ]
    # <-- BATAS PERBAIKAN -->
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
    # menu_sales di sini sudah bersih dari 'Ocha' dan 'Mineral Water'
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
@st.cache_data(show_spinner=False)
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

    # --- BLOK 1: Logika Pemuatan DB (DIPERBAIKI) ---
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

            # --- PERBAIKAN: Kembalikan 'DB_MODE' agar header dinamis ---
            return df, "DB_MODE", "DB_MODE", "DB_MODE"

    # --- BLOK 2: Logika Asli (jika file di-upload) ---
    if uploaded_file is None:
        return None, None, None, None

    # ... (Sisa fungsi ini tidak perlu diubah) ...
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
            st.error(
                f"Format file {uploaded_file.name} tidak didukung. Harap upload .xlsx atau .csv"
            )
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

    # PERBAIKAN TAMBAHAN: Pastikan semua kolom tanggal di-load
    if "Order Time" in df_data.columns:
        df_data["Order Time"] = pd.to_datetime(df_data["Order Time"], errors="coerce")

    df_data["Company"] = company_name
    df_data["Period"] = period_str
    df_data["Branch"] = branch_name_header

    if "Bill Number" not in df_data.columns or "Sales Date In" not in df_data.columns:
        st.warning(
            "Peringatan: Kolom 'Bill Number' atau 'Sales Date In' tidak ditemukan setelah pembersihan."
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

    # ... (Sisa fungsi ini tidak perlu diubah, biarkan apa adanya) ...
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


@st.cache_data
def load_cogs_data(uploaded_file, use_db=False):
    """Memuat dan membersihkan data COGS (File 2)."""

    # --- BLOK BARU: Logika Pemuatan DB ---
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
            return df  # Kembalikan data dari DB

    # --- Logika Asli (jika file di-upload) ---
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

    # Cek kolom Branch
    if "Branch" not in df.columns:
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

    # Normalisasi kolom cabang jika ada
    if "Branch" in df.columns:
        df["Branch"] = df["Branch"].astype(str).str.strip().str.title()

    return df


@st.cache_data
def load_data_waiter(uploaded_file, use_db=False):
    """Memuat dan membersihkan data Waiter (File 3)."""

    # --- BLOK BARU: Logika Pemuatan DB ---
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
            return df  # Kembalikan data dari DB

    # --- Logika Asli (jika file di-upload) ---
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

    # Cek kolom Branch
    if "Branch" not in df.columns:
        st.warning(
            "Peringatan: Kolom 'Branch' tidak ditemukan di File Waiter. Filter cabang mungkin tidak berfungsi."
        )

    df["Order Time"] = pd.to_datetime(df["Order Time"], errors="coerce")
    df["Total After Bill Discount"] = pd.to_numeric(
        df["Total After Bill Discount"], errors="coerce"
    ).fillna(0)
    df.dropna(subset=["Bill Number", "Order Time"], inplace=True)

    # Normalisasi kolom cabang jika ada
    if "Branch" in df.columns:
        df["Branch"] = df["Branch"].astype(str).str.strip().str.title()

    return df


@st.cache_data
def load_data_ulasan(uploaded_file, use_db=False):
    """Memuat dan membersihkan data Ulasan (File 4)."""

    # --- BLOK BARU: Logika Pemuatan DB ---
    if use_db and uploaded_file is None:
        with st.spinner("Memuat data Ulasan dari database..."):
            numeric_config = {"Rating_Clean": "int"}
            df = load_dataframe_from_db(
                "ulasan_data", numeric_cols_config=numeric_config
            )
            if df is None:
                st.info("Database Ulasan kosong. Silakan upload file baru.", icon="ℹ️")
                return None
            return df  # Kembalikan data dari DB

    # --- Logika Asli (jika file di-upload) ---
    if uploaded_file is None:
        return None

    df = None
    try:
        # --- LOGIKA DIPERBARUI UNTUK EXCEL & CSV ---
        if uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        elif uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, encoding="latin1")
        else:
            st.error(
                "Format file Ulasan (File 4) tidak didukung. Harap upload .xlsx atau .csv"
            )
            return None
        # --- BATAS PERUBAHAN ---

    except Exception as e:
        st.error(f"Error membaca file Ulasan (File 4): {e}")
        st.error(
            "Pastikan file adalah .csv atau .xlsx dengan kolom: Nama, Rating, Ulasan"
        )
        return None

    # --- Pembersihan Data Ulasan ---
    if "Rating" not in df.columns or "Ulasan" not in df.columns:
        st.error("File Ulasan (File 4) harus memiliki kolom 'Rating' dan 'Ulasan'.")
        return None

    # Membersihkan kolom Rating (cth: "5 bintang" -> 5)
    # Menggunakan regex untuk mengekstrak angka pertama yang ditemukan

    # --- PERBAIKAN DARI ERROR SEBELUMNYA ---
    # Menggunakan .str.extract()
    df["Rating_Clean"] = (
        df["Rating"].astype(str).str.extract(r"(\d+)").fillna(0).astype(int)
    )
    # --- BATAS PERBAIKAN ---

    # Menghapus baris yang tidak memiliki ulasan atau rating bersih
    df.dropna(subset=["Ulasan"], inplace=True)
    df = df[df["Rating_Clean"] > 0]  # Hanya ambil yg punya rating valid

    df["Ulasan"] = df["Ulasan"].astype(str)

    return df


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

    def uncheck_db_on_upload():
        """
        Callback untuk menonaktifkan DB JIKA file baru di-upload.
        DAN me-reset status 'tersimpan'.
        """
        st.session_state.use_db = False

        # Reset semua status saat ada file baru
        st.session_state.gmv_saved_status = False
        st.session_state.cogs_saved_status = False
        st.session_state.waiter_saved_status = False
        st.session_state.ulasan_saved_status = False
        st.session_state.purchase_saved_status = False

    # --- BATAS FUNGSI CALLBACK ---

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
            st.session_state.use_db = True

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
            on_change=uncheck_db_on_upload,
        )
        if gmv_file is not None:
            if st.session_state.gmv_saved_status:
                # --- PERBAIKAN DI SINI ---
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
            on_change=uncheck_db_on_upload,
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
            on_change=uncheck_db_on_upload,
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
            on_change=uncheck_db_on_upload,
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
            on_change=uncheck_db_on_upload,
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

    # Kembalikan file DAN status checkbox
    return (
        gmv_file,
        cogs_file,
        waiter_file,
        ulasan_file,
        purchase_file,
        st.session_state.use_db,  # Kita tetap kembalikan var KONTROL kita
    )


def build_global_filters(data_gmv, data_cogs, data_waiter, data_purchase):
    """Menggambar filter global dan mengembalikan data yang sudah difilter."""

    # Inisialisasi data yang difilter sebagai data asli
    filtered_gmv = data_gmv
    filtered_cogs = data_cogs
    filtered_waiter = data_waiter
    filtered_purchase = data_purchase

    # #############################################################
    # --- PERBAIKAN 1: Inisialisasi Variabel Scope Aman ---
    # #############################################################
    master_min_date = pd.Timestamp.now().date()
    master_max_date = pd.Timestamp.now().date()
    filter_source_file = None  # <--- Perbaikan 1 (A)
    selected_branches = []
    selected_date = master_max_date  # <--- Perbaikan 2 (D)

    try:
        if data_gmv is not None and not data_gmv.empty:
            master_min_date = data_gmv["Sales Date In"].min().date()
            master_max_date = data_gmv["Sales Date In"].max().date()
            filter_source_file = "GMV"
        elif data_cogs is not None and not data_cogs.empty:
            master_min_date = data_cogs["Sales Date"].min().date()
            master_max_date = data_cogs["Sales Date"].max().date()
            filter_source_file = "COGS"
        elif data_waiter is not None and not data_waiter.empty:
            master_min_date = data_waiter["Order Time"].min().date()
            master_max_date = data_waiter["Order Time"].max().date()
            filter_source_file = "Waiter"
        elif data_purchase is not None and not data_purchase.empty:
            master_min_date = data_purchase["Purchase Date"].min().date()
            master_max_date = data_purchase["Purchase Date"].max().date()
            filter_source_file = "Purchase"
    except Exception as e:
        st.error(f"Gagal membaca rentang tanggal: {e}")

    # #############################################################
    # --- BLOK FILTER CABANG & TANGGAL (DIMULAI DARI IF INI) ---
    # #############################################################
    if filter_source_file:  # <--- Perbaikan 1 (B)
        st.subheader("Filter Analisis Global")
        st.info(
            f"Filter global saat ini menggunakan rentang tanggal dari file: **{filter_source_file}**"
        )

        # --- MODIFIKASI: BLOK FILTER CABANG BARU ---

        # Branch Filter hanya jika ada data GMV (sebagai master list)
        if data_gmv is not None and "Branch" in data_gmv.columns:
            all_branches = sorted(data_gmv["Branch"].unique())
            selected_branches = st.multiselect(
                "Pilih Cabang (Branch):",
                options=all_branches,
                default=all_branches,
                key="branch_filter",
            )

            # Terapkan filter branch ke GMV
            filtered_gmv = data_gmv[data_gmv["Branch"].isin(selected_branches)]

            # >>> TAMBAHAN UNTUK DATA LAIN (Filter Cabang) <<<
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

        # --- BATAS MODIFIKASI BLOK FILTER CABANG ---

        # Terapkan filter radio ke semua tab
        filter_type = st.radio(
            "Pilih rentang waktu (Tab 1, 2, 3, 8):",
            ["Semua Periode", "Harian", "Mingguan", "Bulanan", "Tahunan"],
            horizontal=True,
            key="filter_type_global",
        )

        # --- Terapkan filter tanggal ke 'filtered_...' (Lanjutan) ---
        if filter_type == "Harian":
            selected_date = st.date_input(  # <--- Perbaikan 2 (A)
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
            if data_cogs is not None:
                filtered_cogs = (
                    filtered_cogs[  # Gunakan filtered_cogs (sudah difilter cabang)
                        filtered_cogs["Sales Date"].dt.date == selected_date
                    ]
                )
            if data_waiter is not None:
                filtered_waiter = filtered_waiter[  # Gunakan filtered_waiter
                    filtered_waiter["Order Time"].dt.date == selected_date
                ]
            if data_purchase is not None:
                filtered_purchase = filtered_purchase[  # Gunakan filtered_purchase
                    filtered_purchase["Purchase Date"].dt.date == selected_date
                ]

        elif filter_type == "Mingguan":
            default_start_mingguan = master_max_date - pd.to_timedelta(6, unit="d")
            if default_start_mingguan < master_min_date:
                default_start_mingguan = master_min_date

            selected_start_date = st.date_input(
                "Pilih Tanggal Mulai (periode 7 hari)",
                value=default_start_mingguan,
                min_value=master_min_date,
                max_value=master_max_date,
                key="global_week_start",
            )
            start_date = selected_start_date
            end_date = start_date + pd.to_timedelta(6, unit="d")
            if end_date > master_max_date:
                end_date = master_max_date
            st.info(
                f"Menampilkan data dari {start_date.strftime('%d-%m-%Y')} s.d. {end_date.strftime('%d-%m-%Y')}"
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

        elif filter_type == "Bulanan":
            selected_date = st.date_input(  # <--- Perbaikan 2 (B)
                "Pilih tanggal dalam bulan",
                value=master_max_date,
                min_value=master_min_date,
                max_value=master_max_date,
                key="global_month",
            )
            selected_month = selected_date.month
            selected_year = selected_date.year
            st.info(f"Menampilkan data untuk bulan {selected_date.strftime('%B %Y')}")

            if filtered_gmv is not None:
                filtered_gmv = filtered_gmv[
                    (filtered_gmv["Sales Date In"].dt.month == selected_month)
                    & (filtered_gmv["Sales Date In"].dt.year == selected_year)
                ]
            if filtered_cogs is not None:
                filtered_cogs = filtered_cogs[
                    (filtered_cogs["Sales Date"].dt.month == selected_month)
                    & (filtered_cogs["Sales Date"].dt.year == selected_year)
                ]
            if filtered_waiter is not None:
                filtered_waiter = filtered_waiter[
                    (filtered_waiter["Order Time"].dt.month == selected_month)
                    & (filtered_waiter["Order Time"].dt.year == selected_year)
                ]
            if filtered_purchase is not None:
                filtered_purchase = filtered_purchase[
                    (filtered_purchase["Purchase Date"].dt.month == selected_month)
                    & (filtered_purchase["Purchase Date"].dt.year == selected_year)
                ]

        elif filter_type == "Tahunan":
            selected_date = st.date_input(  # <--- Perbaikan 2 (C)
                "Pilih tanggal dalam tahun",
                value=master_max_date,
                min_value=master_min_date,
                max_value=master_max_date,
                key="global_year",
            )
            selected_year = selected_date.year
            st.info(f"Menampilkan data untuk tahun {selected_year}")

            if filtered_gmv is not None:
                filtered_gmv = filtered_gmv[
                    filtered_gmv["Sales Date In"].dt.year == selected_year
                ]
            if filtered_cogs is not None:
                filtered_cogs = filtered_cogs[
                    filtered_cogs["Sales Date"].dt.year == selected_year
                ]
            if filtered_waiter is not None:
                filtered_waiter = filtered_waiter[
                    filtered_waiter["Order Time"].dt.year == selected_year
                ]
            if filtered_purchase is not None:
                filtered_purchase = filtered_purchase[
                    filtered_purchase["Purchase Date"].dt.year == selected_year
                ]

        # Jika filter_type == "Semua Periode", kita tidak melakukan apa-apa.

    elif (
        data_gmv is None
        and data_cogs is None
        and data_waiter is None
        and data_purchase is None
    ):
        st.markdown("---")

    # Kembalikan data yang SUDAH DIFILTER
    return filtered_gmv, filtered_cogs, filtered_waiter, filtered_purchase


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
            ) = get_menu_performance(filtered_gmv)

            # === 2. TAMPILKAN KPI UTAMA (Expander Asli Anda) ===
            with st.expander(
                "📈 KPI Kinerja Penjualan (Revenue, ATV, IPB)", expanded=True
            ):
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

                    # INI ADALAH GRAFIK ALTAIR KUSTOM ANDA, JADI KITA BIARKAN
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
                    # --- DIGANTI ---
                    st.plotly_chart(chart, use_container_width=True)
                with col12:
                    chart = create_horizontal_bar_chart(
                        top_grossing,
                        "Total Nett Sales",
                        "Menu",
                        "Total Nett Sales (Rp)",
                        "Menu Pendapatan Tertinggi (by Nett Sales)",
                    )
                    # --- DIGANTI ---
                    st.plotly_chart(chart, use_container_width=True)

            st.markdown("---")

            # === 5. URUTAN BARU 4: KPI OPERASIONAL ===
            st.header("⚙️ Analisis Operasional & Pelanggan")

            with st.expander(
                "⚙️ KPI Operasional (Jam Sibuk, Hari Sibuk, Durasi Makan)", expanded=True
            ):
                # Panggil fungsi kpi operasional di sini
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
                    # --- DIGANTI ---
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
                    # --- DIGANTI ---
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
                    # --- DIGANTI ---
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
                    # --- DIGANTI ---
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

            # #############################################################
            # --- BATAS BLOK INSIGHT BARU ---
            # #############################################################

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
                f"** busiest Waktu Tersibuk (Transaksi):** Sesi **{peak_time_trx['Waktu Kunjungan']}** "
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
    total_cost, cost_by_category, cost_by_supplier, top_items
):
    """
    Menganalisis data pembelian yang sudah diproses dan menghasilkan insight
    dalam bahasa alami untuk Tab 8.
    """
    insights = []

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


def build_tab3_hr(filtered_waiter):
    """Menggambar semua elemen untuk Tab 3.
    (VERSI BARU DENGAN KOTAK INSIGHT DI BAWAH)
    """
    if filtered_waiter is not None:
        if not filtered_waiter.empty:
            st.header("🧑‍🍳 Analisis Kinerja Waiter & Waktu Kunjungan")

            # 1. Hitung data analisis (sudah di-cache)
            time_data = get_peak_time_analysis(filtered_waiter)
            waiter_data = get_waiter_performance(filtered_waiter)

            # 2. Expander Waktu Kunjungan (tidak berubah)
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
                        x_type="O",
                        sort_order=sort_order_time,
                    )
                    # --- DIGANTI ---
                    st.plotly_chart(chart1, use_container_width=True)
                with t_col2:
                    st.markdown("##### Berdasarkan Total Penjualan")
                    chart2 = create_vertical_bar_chart(
                        time_data,
                        "Waktu Kunjungan",
                        "Total_Penjualan",
                        "Waktu Kunjungan",
                        "Total Penjualan (Rp)",
                        x_type="O",
                        sort_order=sort_order_time,
                    )
                    # --- DIGANTI ---
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

            # 3. Expander Performa Waiter (tidak berubah)
            with st.expander("🏆 Performa Waiter Teratas (Top 10)", expanded=True):
                st.subheader("🏆 Performa Waiter Teratas (Top 10)")
                chart_waiter = create_horizontal_bar_chart(
                    waiter_data,
                    "Total_Penjualan",
                    "Waiter",
                    "Total Penjualan (Rp)",
                    "Performa Waiter Teratas (by Penjualan)",
                )
                # --- DIGANTI ---
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

            # #############################################################
            # --- BLOK INSIGHT BARU DITEMPATKAN DI SINI (PALING BAWAH) ---
            # #############################################################

            st.markdown("---")  # Tambahkan pemisah visual
            st.header("💡 Insight Otomatis (SDM & Waktu)")

            # Panggil fungsi 'pencari insight' kita
            # Kita gunakan data yang sudah dihitung di awal tab ini
            insights = generate_hr_insights(time_data, waiter_data)

            # Tampilkan dalam expander baru
            with st.expander(
                "Klik untuk melihat Temuan Kunci dari Data SDM & Waktu", expanded=True
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

        elif filtered_waiter is not None and filtered_waiter.empty:
            st.warning(
                "Tidak ada data ditemukan di File Rekapitulasi untuk rentang waktu yang dipilih."
            )
    else:
        st.info(
            "Silakan upload file Laporan Rekapitulasi Detail (File 3) di sidebar untuk melihat analisis waiter."
        )


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
            default_B_date = default_A_date - pd.DateOffset(months=1).date()
        elif comparison_type == "Tahunan":
            default_B_date = default_A_date - pd.DateOffset(years=1).date()

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

        penjualan_saat_ini = data_bulan_aktif["Total After Bill Discount"].sum()
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

        if pd.isna(avg_sales_weekday) or avg_sales_weekday == 0:
            avg_sales_weekday = 1
        if pd.isna(avg_sales_weekend) or avg_sales_weekend == 0:
            # Jika weekend 0, setel ke weekday agar tidak terjadi weight 0
            avg_sales_weekend = avg_sales_weekday

        kpi_insight_dict["avg_sales_weekday"] = avg_sales_weekday  # -> Simpan
        kpi_insight_dict["avg_sales_weekend"] = avg_sales_weekend  # -> Simpan

        weekend_weight = avg_sales_weekend / avg_sales_weekday

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
            if pembagi > 0:
                rdr_weekday = sales_dibutuhkan / pembagi
                rdr_weekend = rdr_weekday * weekend_weight

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
        value=f"{pencapaian_persen*100:,.1f}%",  # <-- Perbaikan kecil di sini
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
            delta=f"{pencapaian_persen*100:,.1f}% dari Target",  # <-- Perbaikan kecil
            delta_color=status_color,
        )
        col3.metric(
            label="Selisih Target",
            value=format_rupiah(penjualan_saat_ini - target_bulanan),
            help="Hasil akhir penjualan dikurangi target bulanan.",
        )

    st.subheader("🔮 Perbandingan Model Ramalan")
    st.write(
        "Membandingkan Proyeksi Cerdas (berbasis rata-rata) dengan Proyeksi Prophet (berbasis tren & musiman)."
    )

    with st.container(border=True):
        col_p1, col_p2 = st.columns(2)

        if target_bulanan > 0:
            proyeksi_prophet_persen = proyeksi_prophet / target_bulanan
        else:
            proyeksi_prophet_persen = 0

        if proyeksi_prophet_persen > 1.05:
            prophet_color = "normal"
        elif proyeksi_prophet_persen >= 0.98:
            prophet_color = "off"
        else:
            prophet_color = "inverse"

        col_p1.metric(
            label="Proyeksi Akhir Bulan (Cerdas)",
            value=format_rupiah(proyeksi_akhir_bulan),
            delta=f"{proyeksi_vs_target_persen * 100:,.1f}% dari Target",
            delta_color=status_color,
            help="Estimasi berdasarkan performa rata-rata weekday & weekend Anda.",
        )
        col_p2.metric(
            label="Proyeksi Akhir Bulan (Model Prophet)",
            value=format_rupiah(proyeksi_prophet),
            delta=f"{proyeksi_prophet_persen * 100:,.1f}% dari Target",
            delta_color=prophet_color,
            help="Estimasi dari model AI (Prophet) berdasarkan tren & pola mingguan.",
        )

    st.markdown("---")
    st.subheader("💡 Rencana Aksi & Diagnostik")

    col_a, col_b = st.columns(2)
    if sisa_hari > 0:

        # #############################################################
        # --- PERBAIKAN SINTAKS DI SINI ---
        # #############################################################

        col_a.metric(
            label="Target Harian Sisa (Weekday)",
            value=format_rupiah(rdr_weekday),
            delta=f"vs Rata-rata: {format_rupiah(avg_sales_weekday)}",
            delta_color=(
                "inverse"
                if rdr_weekday > avg_sales_weekday
                else ("normal" if rdr_weekday > 0 else "off")
            ),
            help="Penjualan harian (Sen-Kam) yang dibutuhkan untuk mencapai target.",
        )

        # #############################################################
        # --- BATAS PERBAIKAN ---
        # #############################################################

        with col_b:
            col_b.metric(
                label="Target Harian Sisa (Weekend)",
                value=format_rupiah(rdr_weekend),
                delta=f"vs Rata-rata: {format_rupiah(avg_sales_weekend)}",
                delta_color=(
                    "inverse"
                    if rdr_weekend > avg_sales_weekend
                    else ("normal" if rdr_weekend > 0 else "off")
                ),
                help="Penjualan harian (Jum-Min) yang dibutuhkan untuk mencapai target.",
            )
    else:
        # Jika bulan sudah selesai
        col_a.metric("Rata-rata Weekday (Final)", format_rupiah(avg_sales_weekday))

    st.markdown("---")
    st.subheader(f"📈 Tren Penjualan Harian - {active_month_name}")

    # Buat grafik Plotly untuk tren harian
    fig_daily_trend = px.line(
        daily_sales_agg,
        x="Sales Date In",
        y="Total After Bill Discount",
        color="Tipe Hari",
        title=f"Tren Penjualan Harian - {active_month_name}",
        labels={
            "Sales Date In": "Tanggal",
            "Total After Bill Discount": "Total Penjualan (Rp)",
            "Tipe Hari": "Tipe Hari",
        },
        template="plotly_white",
        markers=True,
    )
    fig_daily_trend.update_layout(
        title_x=0.01,
        yaxis_tickformat=".2s",
        legend_title_text="",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig_daily_trend.update_traces(
        hovertemplate="<b>%{x|%d %B}</b><br>Penjualan: %{y:,.0f}<extra></extra>"
    )
    st.plotly_chart(fig_daily_trend, use_container_width=True)

    # Grafik Rata-rata Tipe Hari
    avg_sales_df = (
        daily_sales_agg.groupby("Tipe Hari")["Total After Bill Discount"]
        .mean()
        .reset_index()
    )
    fig_avg_type = create_vertical_bar_chart(
        avg_sales_df,
        "Tipe Hari",
        "Total After Bill Discount",
        "Tipe Hari",
        "Rata-rata Penjualan (Rp)",
        x_type="O",
        sort_order=["Weekday (Sen-Kam)", "Weekend (Jum-Min)"],
    )
    st.plotly_chart(fig_avg_type, use_container_width=True)

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


def build_tab8_purchase(filtered_purchase):
    """Menggambar Tab 8 (Analisis Laporan Pembelian)."""
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

    # --- 2. Tampilkan Metrik KPI ---
    st.subheader("📊 KPI Biaya Pembelian")
    st.metric(
        "Total Biaya Pembelian (Tercatat)",
        format_rupiah(total_cost),
        help="Total dari kolom 'Total' di mana nilainya > 0.",
    )

    # --- 3. Tampilkan Grafik Analisis ---
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

    # --- 4. Tampilkan Data Mentah ---
    with st.expander("Lihat Rincian Data Pembelian (Sudah Difilter)"):
        st.dataframe(
            raw_data_filtered.style.format(
                {"Price": format_rupiah, "Total": format_rupiah}
            ),
            use_container_width=True,
        )

    # --- 5. Blok Insight (PALING BAWAH) ---
    st.markdown("---")
    st.header("💡 Insight Otomatis (Analisis Pembelian)")
    insights = generate_purchase_insights(
        total_cost, cost_by_category, cost_by_supplier, top_items
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
# --- FUNGSI BARU UNTUK TAB 9 (REKOMENDASI) ---
# #################################################################


@st.cache_data(show_spinner=False)
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
    # --- PERBAIKAN: Ganti 'rules_df' menjadi 'filtered_rules' ---
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

    # --- BATAS TAMBAHAN ---


# #################################################################
# --- FUNGSI BARU UNTUK TAB 10 (ANALISIS PROMO) ---
# #################################################################


def build_tab10_promo(filtered_cogs):
    """
    Menggambar Tab 10 (Analisis & Simulasi Promo)
    VERSI DIPERBARUI: Menampilkan rincian harga per item di Skenario 2.
    """
    st.header("💸 Analisis & Simulator Profitabilitas Promo")
    st.info(
        "Gunakan alat ini untuk mensimulasikan profit dari skenario diskon atau promo "
        "berdasarkan data COGS Anda."
    )

    # Langkah Kunci: Dapatkan data profitabilitas dari fungsi cache yang ada
    profit_df = analyze_profit(filtered_cogs)

    if profit_df is None or profit_df.empty:
        st.error(
            "Data COGS tidak ditemukan. Harap upload File COGS (File 2) "
            "untuk menggunakan simulator promo ini."
        )
        return

    # Buat daftar menu yang valid (yang memiliki harga jual)
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

    # --- Input UI untuk Scenario 1 ---
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

    # --- Perhitungan Scenario 1 ---
    if selected_menu_s1:
        try:
            # Ambil data asli dari profit_df
            menu_data = profit_df[profit_df["Menu"] == selected_menu_s1].iloc[0]
            original_price = menu_data["Harga Jual"]
            cogs = menu_data["COGS"]
            original_margin_rp = menu_data["Margin (Rp)"]

            # Hitung nilai promo
            discount_rp = original_price * (discount_percent / 100)
            new_price = original_price - discount_rp
            new_margin_rp = new_price - cogs

            # Hitung profit harga normal & selisihnya
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

            # Metrik Total Profit
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

    # --- Input UI untuk Scenario 2 ---
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

    # --- Perhitungan Scenario 2 ---
    if menu_A and menu_B:
        try:
            # Ambil data A
            data_A = profit_df[profit_df["Menu"] == menu_A].iloc[0]
            price_A = data_A["Harga Jual"]
            cogs_A = data_A["COGS"]
            original_margin_A = data_A["Margin (Rp)"]

            # Ambil data B
            data_B = profit_df[profit_df["Menu"] == menu_B].iloc[0]
            price_B_normal = data_B[
                "Harga Jual"
            ]  # <-- [PERMINTAAN] Ambil harga normal B
            cogs_B = data_B["COGS"]

            # Hitung nilai per paket
            revenue_per_deal = price_A
            cost_per_deal = cogs_A + cogs_B
            profit_per_deal = revenue_per_deal - cost_per_deal

            # Hitung perbandingan
            total_promo_profit = profit_per_deal * expected_deals_s2
            total_normal_profit_A_only = original_margin_A * expected_deals_s2
            selisih_profit_s2 = total_promo_profit - total_normal_profit_A_only
            delta_color_s2 = "inverse"

            # <-- [PERUBAHAN TATA LETAK METRIK] -->
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
                    format_rupiah(0),  # Harga jualnya 0
                    f"Harga Normal: {format_rupiah(price_B_normal)}",  # Delta-nya harga normal
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
                    "Profit Bersih (Paket)",
                    format_rupiah(profit_per_deal),
                    "Revenue - Total COGS",
                )

            # Metrik total di luar container
            st.metric(
                f"🎉 TOTAL PROFIT (jika {expected_deals_s2} paket terjual)",
                format_rupiah(total_promo_profit),
                f"Selisih: {format_rupiah(selisih_profit_s2)} vs Jual Normal (Menu A Saja)",
                delta_color=delta_color_s2,
            )
            # <-- [BATAS PERUBAHAN TATA LETAK] -->

            st.info(
                f"**Perbandingan Profit (vs Jual Normal Menu A @ {expected_deals_s2} porsi):**\n"
                f"* **Total Profit (Hanya Jual Menu A):** {format_rupiah(total_normal_profit_A_only)}\n"
                f"* **Total Profit (Promo Paket):** {format_rupiah(total_promo_profit)}\n"
                f"* **Biaya/Kerugian Promo (Selisih):** **{format_rupiah(selisih_profit_s2)}**"
            )

            # Insight khusus BOGO
            if menu_A == menu_B:
                st.warning(
                    f"**Insight BOGO:** Menjual **{expected_deals_s2} paket BOGO** "
                    f"`{menu_A}` akan memberikan total profit **{format_rupiah(total_promo_profit)}**. "
                    f"Ini adalah strategi untuk *volume* (menjual {expected_deals_s2*2} item) "
                    f"dengan mengorbankan margin."
                )

        except Exception as e:
            st.error(f"Gagal menghitung simulasi paket: {e}")


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


# #################################################################
# --- BAGIAN 4: FUNGSI UTAMA (MAIN) ---
# #################################################################
def main():
    """Fungsi utama untuk menjalankan aplikasi Streamlit."""

    # Inisialisasi Database
    init_db()

    # Muat CSS
    load_css("style.css")

    # #############################################################
    # --- 1. INISIALISASI SESSION STATE ---
    # #############################################################
    if "save_gmv_flag" not in st.session_state:
        st.session_state.save_gmv_flag = False
    if "save_cogs_flag" not in st.session_state:
        st.session_state.save_cogs_flag = False
    if "save_waiter_flag" not in st.session_state:
        st.session_state.save_waiter_flag = False
    if "save_ulasan_flag" not in st.session_state:
        st.session_state.save_ulasan_flag = False
    if "save_purchase_flag" not in st.session_state:
        st.session_state.save_purchase_flag = False
    if "gmv_saved_status" not in st.session_state:
        st.session_state.gmv_saved_status = False
    if "cogs_saved_status" not in st.session_state:
        st.session_state.cogs_saved_status = False
    if "waiter_saved_status" not in st.session_state:
        st.session_state.waiter_saved_status = False
    if "ulasan_saved_status" not in st.session_state:
        st.session_state.ulasan_saved_status = False
    if "purchase_saved_status" not in st.session_state:
        st.session_state.purchase_saved_status = False
    if "use_db" not in st.session_state:
        st.session_state.use_db = True

    # #############################################################
    # --- 2. SIDEBAR ---
    # #############################################################
    (
        gmv_file,
        cogs_file,
        waiter_file,
        ulasan_file,
        purchase_file,
        use_db,
    ) = build_sidebar()

    # #############################################################
    # --- 3. PEMUATAN DATA ---
    # #############################################################
    data_gmv, file_company, file_period, file_branch = load_data_gmv(gmv_file, use_db)

    with st.spinner("Memuat data pendukung..."):
        data_cogs = load_cogs_data(cogs_file, use_db)
        data_waiter = load_data_waiter(waiter_file, use_db)
        data_ulasan = load_data_ulasan(ulasan_file, use_db)
        data_purchase = load_data_purchase(purchase_file, use_db)

    # #############################################################
    # --- 4. LOGIKA PENYIMPANAN DATA ---
    # #############################################################

    def clear_db_cache_and_rerun():
        load_dataframe_from_db.clear()
        st.session_state.use_db = True
        st.rerun()

    if st.session_state.save_gmv_flag:
        if data_gmv is not None:
            save_dataframe_smart_append(data_gmv, "gmv_data", "Sales Date In")
            st.session_state.gmv_saved_status = True
            st.session_state.save_gmv_flag = False
            clear_db_cache_and_rerun()

    if st.session_state.save_cogs_flag:
        if data_cogs is not None:
            save_dataframe_smart_append(data_cogs, "cogs_data", "Sales Date")
            st.session_state.cogs_saved_status = True
            st.session_state.save_cogs_flag = False
            clear_db_cache_and_rerun()

    if st.session_state.save_waiter_flag:
        if data_waiter is not None:
            save_dataframe_smart_append(data_waiter, "waiter_data", "Order Time")
            st.session_state.waiter_saved_status = True
            st.session_state.save_waiter_flag = False
            clear_db_cache_and_rerun()

    if st.session_state.save_ulasan_flag:
        if data_ulasan is not None:
            save_dataframe_to_db(data_ulasan, "ulasan_data")
            st.session_state.ulasan_saved_status = True
            st.session_state.save_ulasan_flag = False
            clear_db_cache_and_rerun()

    if st.session_state.save_purchase_flag:
        if data_purchase is not None:
            save_dataframe_smart_append(data_purchase, "purchase_data", "Purchase Date")
            st.session_state.purchase_saved_status = True
            st.session_state.save_purchase_flag = False
            clear_db_cache_and_rerun()

    # #############################################################
    # --- 5. FILTER GLOBAL ---
    # #############################################################
    (
        filtered_gmv,
        filtered_cogs,
        filtered_waiter,
        filtered_purchase,
    ) = build_global_filters(data_gmv, data_cogs, data_waiter, data_purchase)

    # #############################################################
    # --- 6. RENDER KONTEN UTAMA (HEADER + TABS ATAU WELCOME) ---
    # #############################################################

    # --- OPSI A: TIDAK ADA DATA SAMA SEKALI (Welcome Screen) ---
    if data_gmv is None:

        build_welcome_screen()  # Panggil fungsi selamat datang
        build_footer()  # Tampilkan footer juga di halaman utama

        st.stop()  # Hentikan eksekusi script di sini

    # --- OPSI B: ADA DATA (Tampilkan Dashboard) ---

    # --- 6.A. Tampilkan Header Dinamis ---
    if use_db == False and file_company != "DB_MODE":
        st.title(f"Analisis Data: {file_company}")
        st.subheader(f"Cabang: {file_branch} | Periode Data: {file_period}")

    elif use_db == True and filtered_gmv is not None and not filtered_gmv.empty:
        company_name = filtered_gmv["Company"].iloc[0]
        unique_branches = filtered_gmv["Branch"].unique()
        if len(unique_branches) == 1:
            branch_name_header = unique_branches[0]
        else:
            branch_name_header = "Banyak Cabang"

        min_date_str = filtered_gmv["Sales Date In"].min().strftime("%d-%m-%Y")
        max_date_str = filtered_gmv["Sales Date In"].max().strftime("%d-%m-%Y")
        period_str = (
            f"{min_date_str} s.d. {max_date_str}"
            if min_date_str != max_date_str
            else min_date_str
        )

        st.title(f"Analisis Data: {company_name}")
        st.subheader(f"Cabang: {branch_name_header} | Periode Data: {period_str}")

    else:  # Fallback (data ada tapi filter kosong)
        st.title("Dashboard Analisis Data F&B")
        st.warning(
            "Tidak ada data yang ditemukan untuk filter/periode yang Anda pilih. Silakan sesuaikan filter Anda."
        )

    # #############################################################
    # --- 7. NAVIGASI HALAMAN (VERSI st.selectbox / DROPDOWN) ---
    # #############################################################
    st.divider()

    page_options = [
        "📊 Penjualan (GMV)",
        "💰 COGS & Profit",
        "🧑‍🍳 SDM & Waktu",
        "🛒 Pembelian",
        "⚖️ A/B Comparison",
        "🎯 Target",
        "🔮 Forecast (AI)",
        "❤️ Ulasan",
        "💡 Rekomendasi",
        "💸 Analisis Promo",
    ]

    # <-- [PERUBAHAN DI SINI] Menggunakan st.selectbox untuk dropdown list
    page = st.selectbox("Pilih Halaman Analisis:", page_options, key="nav_select")

    st.divider()

    # #############################################################
    # --- 8. RENDER HALAMAN (VERSI if/elif) ---
    # #############################################################

    if page == "📊 Penjualan (GMV)":
        build_tab1_sales(filtered_gmv)
    elif page == "💰 COGS & Profit":
        build_tab2_cogs(filtered_cogs)
    elif page == "🧑‍🍳 SDM & Waktu":
        build_tab3_hr(filtered_waiter)
    elif page == "🛒 Pembelian":
        build_tab8_purchase(filtered_purchase)
    elif page == "⚖️ A/B Comparison":
        build_tab4_comparison(data_gmv, data_cogs, data_waiter)
    elif page == "🎯 Target":
        build_tab6_target(data_gmv)
    elif page == "🔮 Forecast (AI)":
        build_tab5_forecast(data_gmv)
    elif page == "❤️ Ulasan":
        build_tab7_ulasan(data_ulasan)
    elif page == "💡 Rekomendasi":
        build_tab9_rekomendasi(filtered_gmv)
    elif page == "💸 Analisis Promo":
        build_tab10_promo(filtered_cogs)

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


import json  # Pastikan ini ada di atas file
import streamlit as st
from streamlit_lottie import st_lottie  # Pastikan ini ada di atas file


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
    st.subheader("Ubah data mentah Anda menjadi insight yang dapat ditindaklanjuti.")
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

    # 3. Penjelasan Fitur (Sekarang dalam satu kolom penuh)
    st.header("✨ Apa yang Bisa Anda Lakukan?")
    st.markdown(
        """
    * **📊 Analisis Penjualan (GMV):** Lihat performa penjualan, menu terlaris, dan jam sibuk.
    * **💰 Analisis COGS & Profit:** Temukan menu mana yang paling *profitabel*, bukan hanya paling laku.
    * **🧑‍🍳 Kinerja SDM:** Lacak performa waiter dan lihat siapa top sales Anda.
    * **🛒 Biaya Pembelian:** Kontrol pengeluaran dengan menganalisis biaya per supplier dan per item.
    * **⚖️ A/B Comparison:** Bandingkan kinerja 'Minggu Ini' vs 'Minggu Lalu' secara berdampingan.
    * **🎯 Pelacakan Target:** Masukkan target bulanan Anda dan lihat proyeksi pencapaian secara *real-time*.
    * **🔮 Forecast (AI):** Gunakan AI (Prophet) untuk meramalkan tren penjualan Anda 30 hari ke depan.
    * **❤️ Ulasan Pelanggan:** Pahami sentimen pelanggan dan temukan keluhan utama (cth: 'makanan dingin', 'lama').
    * **💡 Rekomendasi:** Dapatkan rekomendasi *cross-sell* (JIKA beli A, TAWARKAN B) berdasarkan data transaksi nyata.
    * **💸 Simulator Promo:** Hitung profitabilitas skenario diskon atau BOGO sebelum Anda menjalankannya.
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


# #################################################################
# --- ENTRY POINT APLIKASI ---
# #################################################################

if __name__ == "__main__":
    main()
