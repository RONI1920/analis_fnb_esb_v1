import time
import pandas as pd
import streamlit as st
import openpyxl  # Diperlukan agar pandas bisa membaca file .xlsx
import altair as alt  # Library untuk grafik yang lebih baik
import numpy as np  # Diperlukan untuk kalkulasi margin

# --- Konfigurasi Halaman Streamlit ---
st.set_page_config(
    page_title="Data Driven Analyst Dashboard KPI - Milky Way Lippo Mall Puri",
    layout="wide",
    initial_sidebar_state="expanded",
)


# --- FUNGSI UNTUK MEMUAT CSS EKSTERNAL ---
def load_css(file_name):
    """Membaca file CSS dan menerapkannya ke aplikasi."""
    try:
        with open(file_name) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.error(
            f"File CSS '{file_name}' tidak ditemukan. Pastikan file ada di folder yang sama dengan skrip Python ini."
        )


# --- Terapkan CSS dari file eksternal ---
load_css("style.css")
# -------------------------------------------

# --- CSS KUSTOM UNTUK MEMPERBESAR TULISAN TAB (DIPERTAHANKAN UNTUK KONSISTENSI) ---
# Namun karena spesifisitas inline CSS di Streamlit, kita pertahankan:
st.markdown(
    """
<style>
/* Target kelas tab Streamlit untuk memastikan ukuran font lebih besar */
.stTabs [data-baseweb="tab-list"] button {
    font-size: 1.25rem; /* Ukuran font lebih besar */
    padding: 10px 15px; /* Tambah padding agar area klik nyaman */
}
</style>
""",
    unsafe_allow_html=True,
)
# -----------------------------------------------------------------


# --- Fungsi Bantuan (Helpers) ---


def format_rupiah(amount):
    """Format angka menjadi string Rupiah DENGAN titik pemisah ribuan."""
    return f"Rp {amount:,.0f}".replace(",", ".")


def format_angka_bulat(number):
    """Format angka kuantitas menjadi bulat tanpa desimal."""
    return f"{number:.0f}"


def format_persen(number):
    """Format angka menjadi string persentase."""
    return f"{number:,.1f}%"


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


# --- Fungsi Pemuatan & Pembersihan Data ---


@st.cache_data  # Cache data agar tidak perlu dimuat ulang setiap interaksi
def load_data_gmv(uploaded_file):
    """Memuat dan membersihkan data GMV dari file yang di-upload (Header baris 10)."""
    df = None
    try:
        if uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, header=9, engine="openpyxl")
        elif uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, header=9, encoding="latin1")
        else:
            st.error("Format file tidak didukung. Harap upload .xlsx atau .csv")
            return None
    except Exception as e:
        st.error(f"Error saat membaca file GMV: {e}")
        st.error(
            "Pastikan file GMV yang di-upload valid dan header berada di baris ke-10."
        )
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
        else:
            st.warning(f"Peringatan (File GMV): Kolom '{col}' tidak ditemukan.")

    # Kolom krusial untuk operasional
    if "Sales Date In" in df.columns:
        df["Sales Date In"] = pd.to_datetime(df["Sales Date In"], errors="coerce")
    if "Sales Date Out" in df.columns:
        df["Sales Date Out"] = pd.to_datetime(df["Sales Date Out"], errors="coerce")

    df.dropna(subset=["Bill Number", "Sales Date In"], inplace=True)
    return df


@st.cache_data
def load_cogs_data(uploaded_file):
    """Memuat data COGS dari file Excel atau CSV (Header baris 13)."""
    df = None
    try:
        # File COGS Anda headernya di baris 13 -> header=12
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, header=12, encoding="latin1")
        elif uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, header=12, engine="openpyxl")
        else:
            st.error("Format file COGS tidak didukung. Harap upload .csv atau .xlsx")
            return None

    except Exception as e:
        st.error(f"Error saat membaca file COGS: {e}")
        st.error("Pastikan header (Menu, Price, COGS Total) ada di BARIS 13 file Anda.")
        return None

    # Mapping kolom dari file COGS
    column_mapping = {
        "Price": "Harga Jual",
        "COGS Total": "COGS",
    }
    df.rename(columns=column_mapping, inplace=True)

    # --- PERUBAHAN DI SINI ---
    # Verifikasi kolom-kolom penting (Sekarang menggunakan nama baru)
    required_cols = ["Menu", "Harga Jual", "COGS", "Qty", "Total", "Sales Date"]
    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        st.error(
            f"File COGS harus memiliki kolom: 'Menu', 'Price' (Harga Jual), 'COGS Total' (COGS), 'Qty', 'Total', dan 'Sales Date'."
        )
        st.error(f"Kolom yang HILANG setelah rename: {missing_cols}")
        st.error(f"Kolom yang DITEMUKAN: {list(df.columns)}")
        return None

    # Konversi Tipe Data
    df["Menu"] = df["Menu"].astype(str)
    df["Harga Jual"] = pd.to_numeric(df["Harga Jual"], errors="coerce").fillna(0)
    df["COGS"] = pd.to_numeric(df["COGS"], errors="coerce").fillna(0)
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
    df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0)
    # Tambahkan konversi tanggal untuk file COGS
    df["Sales Date"] = pd.to_datetime(df["Sales Date"], errors="coerce")

    return df


# --- FUNGSI BARU UNTUK FILE KE-3 ---
@st.cache_data
def load_data_waiter(uploaded_file):
    """Memuat data Rekapitulasi Detail (Waiter & Waktu). Header baris 12."""
    df = None
    try:
        # File Rekapitulasi Anda headernya di baris 12 -> header=11
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, header=11, encoding="latin1")
        elif uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, header=11, engine="openpyxl")
        else:
            st.error("Format file tidak didukung. Harap upload .csv atau .xlsx")
            return None
    except Exception as e:
        st.error(f"Error saat membaca file Rekapitulasi: {e}")
        st.error(
            "Pastikan header (Bill Number, Waiter, Order Time) ada di BARIS 12 file Anda."
        )
        return None

    # Kolom penting
    required_cols = ["Bill Number", "Waiter", "Order Time", "Total After Bill Discount"]
    if not all(col in df.columns for col in required_cols):
        st.error(
            f"File Rekapitulasi Detail harus memiliki kolom: {', '.join(required_cols)}"
        )
        st.error(f"Kolom ditemukan: {list(df.columns)}")
        return None

    # Konversi
    df["Order Time"] = pd.to_datetime(df["Order Time"], errors="coerce")
    df["Total After Bill Discount"] = pd.to_numeric(
        df["Total After Bill Discount"], errors="coerce"
    ).fillna(0)

    # Hapus baris di mana data penting tidak ada
    df.dropna(subset=["Bill Number", "Order Time"], inplace=True)
    return df


# --- AKHIR FUNGSI BARU ---


# --- Fungsi Analisis KPI (Penjualan) ---


def calculate_sales_kpi(df):
    """Menghitung KPI Penjualan Utama."""
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


def get_payment_analysis(df):
    """Menganalisis penjualan berdasarkan metode pembayaran."""
    # Agregasi di level bill dulu
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


def get_visit_purpose_analysis(df):
    """Menganalisis penjualan berdasarkan tujuan kunjungan."""
    # Agregasi di level bill dulu
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


def get_menu_performance(df):
    """Menganalisis performa menu dan kategori."""
    menu_sales = df[~df["Menu"].str.contains("PACKAGE", na=False, case=False)]
    menu_sales = menu_sales[menu_sales["Price (Net)"] > 0]

    # Inisialisasi variabel kategori
    top_selling_categories = pd.DataFrame(columns=["Menu Category", "Qty"])
    top_grossing_categories = pd.DataFrame(
        columns=["Menu Category", "Total Nett Sales"]
    )

    if "Menu Category" in df.columns:
        menu_sales_cat = menu_sales[
            ~menu_sales["Menu Category"].str.contains(
                "ADDITIONAL", na=False, case=False
            )
        ]
        menu_sales_cat = menu_sales_cat[
            ~menu_sales_cat["Menu Category"].str.contains(
                "ADD-ONS", na=False, case=False
            )
        ]

        top_selling_categories = (
            menu_sales_cat.groupby("Menu Category")["Qty"]
            .sum()
            .nlargest(10)
            .sort_values(ascending=False)
        )
        top_grossing_categories = (
            menu_sales_cat.groupby("Menu Category")["Total Nett Sales"]
            .sum()
            .nlargest(10)
            .sort_values(ascending=False)
        )

    top_selling_items = menu_sales.groupby("Menu")["Qty"].sum().nlargest(10)
    top_grossing_items = (
        menu_sales.groupby("Menu")["Total Nett Sales"].sum().nlargest(10)
    )
    all_menu_sales = menu_sales.groupby("Menu")["Qty"].sum()
    bottom_selling_items = all_menu_sales.nsmallest(10).sort_values(ascending=True)
    all_menu_revenue = menu_sales.groupby("Menu")["Total Nett Sales"].sum()
    bottom_grossing_items = all_menu_revenue.nsmallest(10).sort_values(ascending=True)

    return (
        top_selling_items.reset_index(),
        top_grossing_items.reset_index(),
        top_selling_categories.reset_index(),  # Variabel ini sekarang ada
        top_grossing_categories.reset_index(),  # Variabel ini sekarang ada
        bottom_selling_items.reset_index(),
        bottom_grossing_items.reset_index(),
    )


def get_operational_kpi(df):
    """Menghitung KPI Operasional (Jam Sibuk, Hari Sibuk, Durasi)."""
    # Pastikan 'Visit Purpose' dan 'Sales Date Out' ada
    if "Visit Purpose" not in df.columns or "Sales Date Out" not in df.columns:
        st.warning(
            "Kolom 'Visit Purpose' / 'Sales Date Out' tidak ada. KPI Durasi Makan tidak dapat dihitung."
        )
        avg_dining_time = 0.0  # Beri nilai default
    else:
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


# --- FUNGSI ANALISIS COGS & PROFIT (HANYA PAKAI FILE COGS) ---


@st.cache_data
def analyze_profit(df_cogs):
    """
    Menganalisis profitabilitas HANYA dari file COGS (file GMV tidak dipakai).
    Logika ini menjumlahkan (SUM) semua transaksi berdasarkan Menu.
    """

    # File COGS Anda sudah berisi Qty, Harga Jual (per unit), dan COGS (per unit).
    # Kita hanya perlu menghitung profit per baris, lalu menjumlahkannya berdasarkan Menu.

    profit_df = df_cogs.copy()

    # 1. Hitung metrik profit per baris (per transaksi)
    #    Ini adalah Harga Jual per unit - COGS per unit
    profit_df["Margin (Rp)"] = profit_df["Harga Jual"] - profit_df["COGS"]

    # 2. Hitung total profit/revenue/cogs per BARIS TRANSAKSI
    #    (dikalikan Qty di baris itu)
    profit_df["Total Revenue (Rp)"] = profit_df[
        "Total"
    ]  # Ambil dari kolom "Total" yang sudah di-sum
    profit_df["Total COGS (Rp)"] = profit_df["COGS"] * profit_df["Qty"]
    profit_df["Total Profit (Rp)"] = profit_df["Margin (Rp)"] * profit_df["Qty"]

    # 3. Sekarang, Agregasi (JUMLAHKAN) berdasarkan Menu
    #    Ini adalah langkah yang Anda inginkan (menjumlahkan)
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

    # 4. Hitung ulang Metrik Agregat (Margin % dan Margin Rp per unit)
    #    Hindari pembagian dengan nol jika Qty = 0
    agg_df["Margin (Rp)"] = np.where(
        agg_df["Qty"] > 0, agg_df["Total_Profit_Rp"] / agg_df["Qty"], 0
    )
    agg_df["Margin (%)"] = np.where(
        agg_df["Total_Revenue_Rp"] > 0,
        (agg_df["Total_Profit_Rp"] / agg_df["Total_Revenue_Rp"]) * 100,
        0,
    )

    # 5. Ambil Harga Jual & COGS per unit (rata-rata) dari data asli
    #    Kita ambil 'mean' (rata-rata) untuk jaga-jaga jika ada harga yang beda-beda
    unit_costs = (
        profit_df[profit_df["Harga Jual"] > 0]
        .groupby("Menu")
        .agg(Harga_Jual_Unit=("Harga Jual", "mean"), COGS_Unit=("COGS", "mean"))
        .reset_index()
    )

    final_df = pd.merge(agg_df, unit_costs, on="Menu", how="left")

    # 6. Ganti nama kolom agar sesuai dengan sisa program
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

    # 7. Pilih kolom final
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

    # Isi NaN (jika ada) dengan 0
    final_df.fillna(0, inplace=True)

    return final_df[final_cols].sort_values(by="Total Profit (Rp)", ascending=False)


# --- FUNGSI BARU UNTUK ANALISIS WAITER & WAKTU (FILE KE-3) ---


def get_peak_time_analysis(df):
    """Menganalisis transaksi berdasarkan waktu (Breakfast, Lunch, Dinner)."""

    # 1. Agregasi per Bill Number dulu, karena file detail per menu
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

    # 2. Definisikan kategori waktu (SESUAI PERMINTAAN BARU: 10:00 - 22:00)
    conditions = [
        (bill_df["Hour"] >= 10) & (bill_df["Hour"] < 12),  # 10:00 - 11:59
        (bill_df["Hour"] >= 12) & (bill_df["Hour"] < 17),  # 12:00 - 16:59
        (bill_df["Hour"] >= 17) & (bill_df["Hour"] < 22),  # 17:00 - 21:59
    ]
    choices = [
        "Breakfast/Brunch (10-12)",
        "Lunch (12-17)",
        "Dinner (17-22)",
    ]

    bill_df["Waktu Kunjungan"] = np.select(conditions, choices, default="Luar Jam Buka")

    # 3. Agregasi berdasarkan kategori waktu
    time_analysis = (
        bill_df.groupby("Waktu Kunjungan")
        .agg(
            Jumlah_Transaksi=("Bill Number", "nunique"),
            Total_Penjualan=("Total_Sales", "sum"),
        )
        .reset_index()
    )

    # Urutkan
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


def get_waiter_performance(df):
    """Menganalisis performa waiter (Top 10)."""

    # 1. Agregasi per Bill Number untuk mendapatkan total bill dan waiternya
    # PENTING: File rekapitulasi ada di level menu, kita harus agregasi per bill
    bill_df = (
        df.groupby("Bill Number")
        .agg(
            Waiter=("Waiter", "first"), Total_Sales=("Total After Bill Discount", "sum")
        )
        .reset_index()
    )

    # Ganti NaN/None Waiter dengan 'Tidak Diketahui'
    bill_df["Waiter"] = bill_df["Waiter"].fillna("Tidak Diketahui")

    # 2. Agregasi berdasarkan Waiter
    waiter_perf = (
        bill_df.groupby("Waiter")
        .agg(
            Total_Penjualan=("Total_Sales", "sum"),
            Jumlah_Transaksi=("Bill Number", "nunique"),
        )
        .reset_index()
    )

    # 3. Ambil Top 10, urutkan dari tertinggi
    top_10_waiters = waiter_perf.nlargest(10, "Total_Penjualan")

    return top_10_waiters


# --- Fungsi Grafik Altair (Warna-warni) ---


def create_horizontal_bar_chart(data, x_col, y_col, x_title, y_title, sort_order="-x"):
    """Membuat grafik batang horizontal Altair yang berwarna-warni."""
    chart = (
        alt.Chart(data)
        .mark_bar()
        .encode(
            x=alt.X(f"{x_col}:Q", title=x_title, axis=alt.Axis(format="~s")),
            y=alt.Y(f"{y_col}:N", title=y_title, sort=sort_order),
            color=alt.Color(f"{y_col}:N", title=y_title, legend=None),
            tooltip=[
                alt.Tooltip(y_col, title=y_title),
                alt.Tooltip(x_col, title=x_title, format=",.0f"),
            ],
        )
        .interactive()
    )
    return chart


def create_vertical_bar_chart(
    data, x_col, y_col, x_title, y_title, x_type="N", sort_order=None
):
    """Membuat grafik batang vertikal Altair yang berwarna-warni."""
    chart = (
        alt.Chart(data)
        .mark_bar()
        .encode(
            x=alt.X(f"{x_col}:{x_type}", title=x_title, sort=sort_order),
            y=alt.Y(f"{y_col}:Q", title=y_title),
            color=alt.Color(f"{x_col}:{x_type}", title=x_title, legend=None),
            tooltip=[
                alt.Tooltip(x_col, title=x_title),
                alt.Tooltip(y_col, title=y_title, format=",.0f"),
            ],
        )
        .interactive()
    )
    return chart


# --- Tampilan Utama Dashboard (Main UI) ---

st.title("📈 Data Driven Analyst Dashboard KPI - Milky Way Lippo Mall Puri")

# 1. Widget File Uploader di Sidebar
with st.sidebar:
    # Mengganti logo dengan teks yang lebih besar dan terpisah baris
    # MENGHAPUS INLINE STYLE AGAR CSS EKSTERNAL DARI style.css BERFUNGSI
    st.markdown(
        """
        <div style='text-align: center; margin-bottom: 20px;'>
            <h2>DATA DRIVEN</h2>
            <h2>SPECIALYST FNB</h2>
            <p style='font-size: 0.9rem; color: #aaaaaa; margin-top: 5px;'>Developer: @ronihidayat</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.header("Upload Data")

    # Uploader 1: Laporan GMV (File Operasional)
    gmv_file = st.file_uploader(
        "1. Upload Laporan GMV (Operasional)",
        type=["xlsx", "csv"],
    )
    st.info("ℹ️ Laporan GMV asli (header di baris ke-10).")

    # Uploader 2: Laporan COGS (File Master)
    cogs_file = st.file_uploader(
        "2. Upload Laporan COGS (Menu COGS Report)",
        type=["xlsx", "csv"],  # Izinkan CSV untuk COGS
    )
    st.info("ℹ️ File COGS (header di baris ke-13).")

    # --- PENAMBAHAN FILE UPLOADER KE-3 ---
    waiter_file = st.file_uploader(
        "3. Upload Laporan Waiter (Rekapitulasi Detail)",
        type=["xlsx", "csv"],
    )
    st.info("ℹ️ Laporan Rekapitulasi Detail (header di baris ke-12).")

# --- PINDAHKAN LOGIKA LOADING & FILTER KE GLOBAL ---

# 1. Load semua data secara global
data_gmv = None
data_cogs = None
data_waiter = None

if gmv_file:
    data_gmv = load_data_gmv(gmv_file)
if cogs_file:
    data_cogs = load_cogs_data(cogs_file)
if waiter_file:
    data_waiter = load_data_waiter(waiter_file)

# 2. Inisialisasi variabel filter
filtered_gmv = data_gmv
filtered_cogs = data_cogs
filtered_waiter = data_waiter

# 3. Buat Filter Global (Hanya jika file GMV ada, sebagai master tanggal)
if data_gmv is not None:
    st.subheader("Filter Analisis Global")

    try:
        # Gunakan 'Sales Date In' dari file GMV sebagai master tanggal
        min_date = data_gmv["Sales Date In"].min().date()
        max_date = data_gmv["Sales Date In"].max().date()
    except Exception as e:
        st.error(
            "Gagal membaca rentang tanggal dari file GMV. Menggunakan tanggal hari ini."
        )
        min_date = pd.Timestamp.now().date()
        max_date = pd.Timestamp.now().date()

    filter_type = st.radio(
        "Pilih rentang waktu analisis untuk semua tab:",
        [
            "Semua Periode",
            "Harian",
            "Mingguan",
            "Bulanan",
            "Tahunan",
        ],  # TAMBAH "Tahunan"
        horizontal=True,
        key="filter_type_global",
    )

    # Terapkan filter ke SEMUA dataframe yang ada
    if filter_type == "Harian":
        selected_date = st.date_input(
            "Pilih Tanggal",
            value=max_date,
            min_value=min_date,
            max_value=max_date,
            key="global_date",
        )

        filtered_gmv = data_gmv[data_gmv["Sales Date In"].dt.date == selected_date]
        if data_cogs is not None:
            filtered_cogs = data_cogs[data_cogs["Sales Date"].dt.date == selected_date]
        if data_waiter is not None:
            filtered_waiter = data_waiter[
                data_waiter["Order Time"].dt.date == selected_date
            ]

    elif filter_type == "Mingguan":
        selected_date = st.date_input(
            "Pilih tanggal dalam minggu",
            value=max_date,
            min_value=min_date,
            max_value=max_date,
            key="global_week",
        )
        start_of_week = selected_date - pd.to_timedelta(
            selected_date.weekday(), unit="d"
        )
        end_of_week = start_of_week + pd.to_timedelta(6, unit="d")
        st.info(
            f"Menampilkan data dari {start_of_week.strftime('%d-%m-%Y')} s.d. {end_of_week.strftime('%d-%m-%Y')}"
        )

        filtered_gmv = data_gmv[
            (data_gmv["Sales Date In"].dt.date >= start_of_week)
            & (data_gmv["Sales Date In"].dt.date <= end_of_week)
        ]
        if data_cogs is not None:
            filtered_cogs = data_cogs[
                (data_cogs["Sales Date"].dt.date >= start_of_week)
                & (data_cogs["Sales Date"].dt.date <= end_of_week)
            ]
        if data_waiter is not None:
            filtered_waiter = data_waiter[
                (data_waiter["Order Time"].dt.date >= start_of_week)
                & (data_waiter["Order Time"].dt.date <= end_of_week)
            ]

    elif filter_type == "Bulanan":
        selected_date = st.date_input(
            "Pilih tanggal dalam bulan",
            value=max_date,
            min_value=min_date,
            max_value=max_date,
            key="global_month",
        )
        selected_month = selected_date.month
        selected_year = selected_date.year
        st.info(f"Menampilkan data untuk bulan {selected_date.strftime('%B %Y')}")

        filtered_gmv = data_gmv[
            (data_gmv["Sales Date In"].dt.month == selected_month)
            & (data_gmv["Sales Date In"].dt.year == selected_year)
        ]
        if data_cogs is not None:
            filtered_cogs = data_cogs[
                (data_cogs["Sales Date"].dt.month == selected_month)
                & (data_cogs["Sales Date"].dt.year == selected_year)
            ]
        if data_waiter is not None:
            filtered_waiter = data_waiter[
                (data_waiter["Order Time"].dt.month == selected_month)
                & (data_waiter["Order Time"].dt.year == selected_year)
            ]

    elif filter_type == "Tahunan":  # LOGIKA BARU "Tahunan"
        selected_date = st.date_input(
            "Pilih tanggal dalam tahun",
            value=max_date,
            min_value=min_date,
            max_value=max_date,
            key="global_year",
        )
        selected_year = selected_date.year
        st.info(f"Menampilkan data untuk tahun {selected_year}")

        filtered_gmv = data_gmv[data_gmv["Sales Date In"].dt.year == selected_year]
        if data_cogs is not None:
            filtered_cogs = data_cogs[data_cogs["Sales Date"].dt.year == selected_year]
        if data_waiter is not None:
            filtered_waiter = data_waiter[
                data_waiter["Order Time"].dt.year == selected_year
            ]

    # (Jika "Semua Periode", kita tidak melakukan apa-apa, data = data.copy())

    st.markdown("---")  # Pemisah setelah filter

elif gmv_file is None:
    st.info(
        "ℹ️ Upload file Laporan GMV (file 1) untuk mengaktifkan filter Harian/Mingguan/Bulanan/Tahunan."
    )
    st.markdown("---")

# --- AKHIR DARI LOGIKA FILTER GLOBAL ---


# --- Buat Tabulasi ---
tab1, tab2, tab3 = st.tabs(
    [
        "📊 Analisis Penjualan (KPI)",
        "💰 Analisis COGS & Profit",
        "🧑‍🍳 Analisis SDM & Waktu",  # NAMA TAB KETIGA DIGANTI
    ]
)

# --- ISI TAB 1: ANALISIS PENJUALAN (KPI) ---
with tab1:
    # Gunakan data yang sudah di-load dan di-filter secara global
    if filtered_gmv is not None:
        if not filtered_gmv.empty:
            start_date = filtered_gmv["Sales Date In"].min().strftime("%d-%m-%Y")
            end_date = filtered_gmv["Sales Date In"].max().strftime("%d-%m-%Y")
            st.subheader(f"Periode Analisis: {start_date} s.d. {end_date}")

            # --- 1. Ringkasan Kinerja Penjualan ---
            with st.expander(
                "📈 KPI Kinerja Penjualan (Revenue, ATV, IPB)", expanded=True
            ):
                st.header("📊 KPI Kinerja Penjualan")
                # Gunakan 'filtered_gmv'
                kpi = calculate_sales_kpi(filtered_gmv)

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

                # Expander untuk Service, Tax, Diskon (Data dari GMV)
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

            # PERUBAHAN JS untuk menyembunyikan notifikasi setelah 3 detik
            st.markdown(
                """
                <div id='menu_info_notice'>
                ℹ️ Menu 'PACKAGE', 'ADDITIONAL', 'ADD-ONS', dan item gratis (Price = 0) telah disaring dari analisis performa menu.
                </div>
                <script>
                // Mencari elemen st.info/st.warning terdekat
                const menuInfo = document.getElementById('menu_info_notice');
                if (menuInfo) {
                    // Temukan container Streamlit terdekat yang menampung pesan info
                    let parentElement = menuInfo.closest('.stAlert[data-baseweb="alert"]');
                    if (parentElement) {
                        setTimeout(() => {
                            parentElement.style.transition = 'opacity 0.5s ease-out';
                            parentElement.style.opacity = '0';
                            // Hapus elemen sepenuhnya setelah transisi
                            setTimeout(() => {
                                parentElement.style.display = 'none';
                            }, 500); // 500ms setelah opacity 0
                        }, 3000); // Tunda 3 detik sebelum mulai transisi
                    }
                }
                </script>
                """,
                unsafe_allow_html=True,
            )

            # --- 2. Kinerja Menu ---
            # PERUBAHAN 4: Menambahkan st.expander untuk bagian KPI Kinerja Menu (Teratas)
            with st.expander(
                "🏆 KPI Kinerja Menu (Top 10 Selling/Grossing)", expanded=True
            ):
                st.header("🍽️ KPI Kinerja Menu")
                (
                    top_selling,
                    top_grossing,
                    top_sell_cat,
                    top_gross_cat,
                    bottom_selling,
                    bottom_grossing,
                ) = get_menu_performance(
                    filtered_gmv
                )  # Gunakan 'filtered_gmv'

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
                        top_selling, "Qty", "Menu", "Kuantitas Terjual", "Menu"
                    )
                    st.altair_chart(chart, use_container_width=True)
                with col12:
                    chart = create_horizontal_bar_chart(
                        top_grossing,
                        "Total Nett Sales",
                        "Menu",
                        "Total Nett Sales (Rp)",
                        "Menu",
                    )
                    st.altair_chart(chart, use_container_width=True)

            # PERUBAHAN 5: Menambahkan st.expander untuk bagian KPI Kinerja Menu (Terbawah)
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

                col15, col16 = st.columns(2)
                with col15:
                    chart = create_horizontal_bar_chart(
                        bottom_selling,
                        "Qty",
                        "Menu",
                        "Kuantitas Terjual",
                        "Menu",
                        sort_order="x",
                    )
                    st.altair_chart(chart, use_container_width=True)
                with col16:
                    chart = create_horizontal_bar_chart(
                        bottom_grossing,
                        "Total Nett Sales",
                        "Menu",
                        "Total Nett Sales (Rp)",
                        "Menu",
                        sort_order="x",
                    )
                    st.altair_chart(chart, use_container_width=True)

            # PERUBAHAN 6: Menambahkan st.expander untuk bagian Kategori Menu
            if "Menu Category" in filtered_gmv.columns and not top_sell_cat.empty:
                st.markdown("---")
                with st.expander("🍰 Analisis Kategori Menu"):
                    st.subheader("🍰 Performa Kategori Menu")
                    col17, col18 = st.columns(2)
                    with col17:
                        st.markdown("##### Kategori Terlaris (by Kuantitas)")
                        chart = create_horizontal_bar_chart(
                            top_sell_cat,
                            "Qty",
                            "Menu Category",
                            "Kuantitas Terjual",
                            "Kategori Menu",
                        )
                        st.altair_chart(chart, use_container_width=True)
                    with col18:
                        st.markdown(
                            "##### Kategori Pendapatan Tertinggi (by Nett Sales)"
                        )
                        chart = create_horizontal_bar_chart(
                            top_gross_cat,
                            "Total Nett Sales",
                            "Menu Category",
                            "Total Nett Sales (Rp)",
                            "Kategori Menu",
                        )
                        st.altair_chart(chart, use_container_width=True)

            st.markdown("---")

            # PERUBAHAN 3: Menambahkan st.expander untuk bagian Metode Pembayaran & Tipe Kunjungan
            with st.expander(
                "💳 Analisis Transaksi (Pembayaran & Kunjungan)", expanded=True
            ):
                col7, col8 = st.columns(2)
                # Cek jika kolom ada
                if "Payment Method" in filtered_gmv.columns:
                    with col7:
                        st.subheader("💳 Penjualan per Metode Pembayaran")
                        # Gunakan 'filtered_gmv'
                        payment_data = get_payment_analysis(filtered_gmv)
                        chart = create_horizontal_bar_chart(
                            payment_data,
                            "Total_Penjualan",
                            "Cleaned_Payment",
                            "Total Penjualan (Rp)",
                            "Metode Pembayaran",
                        )
                        st.altair_chart(chart, use_container_width=True)
                else:
                    with col7:
                        st.warning(
                            "Kolom 'Payment Method' tidak ditemukan di file GMV."
                        )

                if "Visit Purpose" in filtered_gmv.columns:
                    with col8:
                        st.subheader("🏪 Penjualan per Tipe Kunjungan")
                        # Gunakan 'filtered_gmv'
                        visit_data = get_visit_purpose_analysis(filtered_gmv)
                        chart = create_horizontal_bar_chart(
                            visit_data,
                            "Total After Bill Discount",
                            "Visit Purpose",
                            "Total Penjualan (Rp)",
                            "Tipe Kunjungan",
                        )
                        st.altair_chart(chart, use_container_width=True)
                else:
                    with col8:
                        st.warning("Kolom 'Visit Purpose' tidak ditemukan di file GMV.")

            st.markdown("---")

            # --- 3. Kinerja Operasional ---
            # PERUBAHAN 7: Menambahkan st.expander untuk bagian KPI Operasional
            with st.expander(
                "⚙️ KPI Operasional (Jam Sibuk, Hari Sibuk, Durasi Makan)", expanded=True
            ):
                st.header("⚙️ KPI Operasional & Pelanggan")
                # Gunakan 'filtered_gmv'
                avg_time, peak_hours, peak_days_of_week = get_operational_kpi(
                    filtered_gmv
                )
                st.metric("⏱️ Rata-rata Durasi Makan (Dine In)", f"{avg_time:.1f} menit")

                col19, col20 = st.columns(2)
                with col19:
                    st.subheader("🕒 Jam Sibuk (Berdasarkan Transaksi)")
                    chart = create_vertical_bar_chart(
                        peak_hours,
                        "Hour",
                        "Bill Number",
                        "Jam",
                        "Jumlah Transaksi",
                        x_type="O",
                    )
                    st.altair_chart(chart, use_container_width=True)
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
                    st.altair_chart(chart, use_container_width=True)

            if st.checkbox("Tampilkan Data Mentah GMV (Sudah Dibersihkan)"):
                st.dataframe(filtered_gmv)

        elif filtered_gmv is not None and filtered_gmv.empty:
            st.warning(
                "Tidak ada data ditemukan di File GMV untuk rentang waktu yang dipilih."
            )
    else:
        st.info(
            "Silakan upload file Laporan GMV di sidebar untuk melihat analisis penjualan."
        )


# --- ISI TAB 2: ANALISIS COGS & PROFIT ---
with tab2:
    # Gunakan data yang sudah di-load dan di-filter secara global
    if filtered_cogs is not None:
        if not filtered_cogs.empty:
            st.header("💰 Analisis Profitabilitas Menu (COGS)")

            # Lakukan analisis profit HANYA dari file cogs yang sudah difilter
            profit_df = analyze_profit(filtered_cogs)

            # PERUBAHAN 8: Menambahkan st.expander untuk Ringkasan Profitabilitas
            with st.expander(
                "💰 Ringkasan Profitabilitas (Total Revenue, COGS, Profit)",
                expanded=True,
            ):
                # Tampilkan KPI Profitabilitas
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

            # PERUBAHAN 9: Menambahkan st.expander untuk Rincian Profitabilitas per Menu
            with st.expander("📝 Rincian Profitabilitas per Menu (Tabel)"):
                # Tampilkan tabel data profit
                st.subheader("Rincian Profitabilitas per Menu")
                st.info(
                    "Data ini dijumlahkan (agregasi) HANYA dari file Laporan COGS (sesuai filter waktu yang dipilih)."
                )

                # Format tabel untuk tampilan
                formatted_df = profit_df.copy()
                formatted_df["Harga Jual"] = formatted_df["Harga Jual"].apply(
                    format_rupiah
                )
                formatted_df["COGS"] = formatted_df["COGS"].apply(format_rupiah)
                formatted_df["Margin (Rp)"] = formatted_df["Margin (Rp)"].apply(
                    format_rupiah
                )
                formatted_df["Margin (%)"] = formatted_df["Margin (%)"].apply(
                    format_persen
                )
                formatted_df["Total Revenue (Rp)"] = formatted_df[
                    "Total Revenue (Rp)"
                ].apply(format_rupiah)
                formatted_df["Total COGS (Rp)"] = formatted_df["Total COGS (Rp)"].apply(
                    format_rupiah
                )
                formatted_df["Total Profit (Rp)"] = formatted_df[
                    "Total Profit (Rp)"
                ].apply(format_rupiah)

                st.dataframe(formatted_df.set_index("Menu"))

            st.markdown("---")

            # PERUBAHAN 10: Menambahkan st.expander untuk Analisis Performa Profit Menu
            with st.expander(
                "📊 Analisis Performa Profit Menu (Grafik Top & Bottom 10)",
                expanded=True,
            ):
                st.subheader("Analisis Performa Profit Menu")

                # Ambil Top 10 dan Bottom 10
                top_10_profit = profit_df.nlargest(10, "Total Profit (Rp)")
                bottom_10_profit = profit_df[profit_df["Qty"] > 0].nsmallest(
                    10, "Total Profit (Rp)"
                )  # Hanya yg terjual
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
                        "Menu",
                    )
                    st.altair_chart(chart, use_container_width=True)
                with p_col6:
                    st.markdown("##### 📈 Menu Margin Tertinggi (by %)")
                    chart = create_horizontal_bar_chart(
                        top_10_margin_pct, "Margin (%)", "Menu", "Margin (%)", "Menu"
                    )
                    st.altair_chart(chart, use_container_width=True)

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
                        "Menu",
                        sort_order="x",
                    )
                    st.altair_chart(chart, use_container_width=True)
                with p_col8:
                    st.markdown("##### 📉 Menu Margin Terendah (by %)")
                    chart = create_horizontal_bar_chart(
                        bottom_10_margin_pct,
                        "Margin (%)",
                        "Menu",
                        "Margin (%)",
                        "Menu",
                        sort_order="x",
                    )
                    st.altair_chart(chart, use_container_width=True)

        elif filtered_cogs is not None and filtered_cogs.empty:
            st.warning(
                "Tidak ada data ditemukan di File COGS untuk rentang waktu yang dipilih."
            )
    else:
        st.info(
            "Silakan upload file Laporan COGS di sidebar untuk melihat analisis profitabilitas."
        )


# --- ISI TAB 3: ANALISIS WAITER & WAKTU ---
with tab3:
    # Gunakan data yang sudah di-load dan di-filter secara global
    if filtered_waiter is not None:
        if not filtered_waiter.empty:
            st.header("🧑‍🍳 Analisis Kinerja Waiter & Waktu Kunjungan")

            # --- FILTER SUDAH DIPINDAHKAN KE ATAS ---

            # 1. Analisis Waktu
            # PERUBAHAN 11: Menambahkan st.expander untuk Analisis Waktu Kunjungan
            with st.expander("🕒 Analisis Waktu Kunjungan Pelanggan", expanded=True):
                st.subheader("🕒 Waktu Kunjungan Pelanggan")
                # Gunakan data yang sudah difilter
                time_data = get_peak_time_analysis(filtered_waiter)

                t_col1, t_col2 = st.columns(2)
                with t_col1:
                    st.markdown("##### Berdasarkan Jumlah Transaksi")
                    # Tentukan urutan manual untuk grafik
                    sort_order_time = [
                        "Breakfast/Brunch (10-12)",
                        "Lunch (12-17)",
                        "Dinner (17-22)",
                        "Luar Jam Buka",
                    ]
                    chart1 = create_vertical_bar_chart(
                        time_data,
                        "Waktu Kunjungan",
                        "Jumlah_Transaksi",
                        "Waktu Kunjungan",
                        "Jumlah Transaksi",
                        x_type="O",
                        sort_order=sort_order_time,
                    )
                    st.altair_chart(chart1, use_container_width=True)
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
                    st.altair_chart(chart2, use_container_width=True)

                st.dataframe(
                    time_data.set_index("Waktu Kunjungan").style.format(
                        {
                            "Total_Penjualan": format_rupiah,
                            "Jumlah_Transaksi": format_angka_bulat,
                        }
                    )
                )

            st.markdown("---")

            # 2. Analisis Waiter
            # PERUBAHAN 12: Menambahkan st.expander untuk Analisis Waiter
            with st.expander("🏆 Performa Waiter Teratas (Top 10)", expanded=True):
                st.subheader("🏆 Performa Waiter Teratas (Top 10)")
                # Gunakan data yang sudah difilter
                waiter_data = get_waiter_performance(filtered_waiter)

                chart_waiter = create_horizontal_bar_chart(
                    waiter_data,
                    "Total_Penjualan",
                    "Waiter",
                    "Total Penjualan (Rp)",
                    "Waiter",
                )
                st.altair_chart(chart_waiter, use_container_width=True)

                st.dataframe(
                    waiter_data.set_index("Waiter").style.format(
                        {
                            "Total_Penjualan": format_rupiah,
                            "Jumlah_Transaksi": format_angka_bulat,
                        }
                    )
                )

        elif filtered_waiter is not None and filtered_waiter.empty:
            st.warning(
                "Tidak ada data ditemukan di File Rekapitulasi untuk rentang waktu yang dipilih."
            )
    else:
        st.info(
            "Silakan upload file Laporan Rekapitulasi Detail (file ke-3) di sidebar untuk melihat analisis waiter."
        )

# --- BAGIAN BARU: FOOTER TETAP (STICKY FOOTER) ---
st.markdown(
    """
    <div id="custom-footer">
        <div class="copyright">
            © Copyright Roni Hidayat Data Driven Speacialist Food and Beverage - 2025
        </div>
        <div class="links">
            <a href="#">Hubungi Kami</a>
            <a href="#">Kebijakan Privasi</a>
            <a href="#">Tentang Kami</a>
            <a href="#">Kerjasama</a>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)
# --- AKHIR FOOTER ---
