import time
import pandas as pd
import streamlit as st
import openpyxl  # Diperlukan agar pandas bisa membaca file .xlsx
import altair as alt  # Library untuk grafik yang lebih baik
import numpy as np  # Diperlukan untuk kalkulasi margin
from prophet import Prophet
from prophet.plot import plot_plotly, plot_components_plotly

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
            f"File CSS '{file_name}' tidak ditemukan. Pastikan file ada di folder yang sama."
        )


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


def get_menu_performance(df):
    """Menganalisis performa menu dan kategori."""
    menu_sales = df[~df["Menu"].str.contains("PACKAGE", na=False, case=False)]
    menu_sales = menu_sales[menu_sales["Price (Net)"] > 0]

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
        top_selling_categories.reset_index(),
        top_grossing_categories.reset_index(),
        bottom_selling_items.reset_index(),
        bottom_grossing_items.reset_index(),
    )


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
        # Dulu ada st.warning, sekarang kita diamkan saja.
        # KPI Durasi Makan akan otomatis 0.0
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
    profit_df["Margin (Rp)"] = profit_df["Harga Jual"] - profit_df["COGS"]
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


def create_horizontal_bar_chart(data, x_col, y_col, x_title, y_title, sort_order="-x"):
    """Membuat grafik batang horizontal Altair."""
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
    """Membuat grafik batang vertikal Altair."""
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


# --- FUNGSI HELPER UNTUK TAB 4, SUDAH BENAR ---
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
# --- BAGIAN 2: FUNGSI PEMUATAN DATA (DENGAN CACHE) ---
# #################################################################


@st.cache_data
def load_data_gmv(uploaded_file):
    """Memuat dan membersihkan data GMV (File 1)."""
    if uploaded_file is None:
        return None
    df = None
    try:
        if uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, header=9, engine="openpyxl")
        elif uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, header=9, encoding="latin1")
        else:
            st.error("Format file 1 tidak didukung. Harap upload .xlsx atau .csv")
            return None
    except Exception as e:
        st.error(f"Error membaca file GMV (File 1): {e}")
        st.error("Pastikan header file GMV ada di baris 10.")
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
            st.warning(f"Peringatan (File 1): Kolom '{col}' tidak ditemukan.")

    if "Sales Date In" in df.columns:
        df["Sales Date In"] = pd.to_datetime(df["Sales Date In"], errors="coerce")
    if "Sales Date Out" in df.columns:
        df["Sales Date Out"] = pd.to_datetime(df["Sales Date Out"], errors="coerce")

    df.dropna(subset=["Bill Number", "Sales Date In"], inplace=True)
    return df


@st.cache_data
def load_cogs_data(uploaded_file):
    """Memuat dan membersihkan data COGS (File 2)."""
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
    return df


@st.cache_data
def load_data_waiter(uploaded_file):
    """Memuat dan membersihkan data Waiter (File 3)."""
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

    df["Order Time"] = pd.to_datetime(df["Order Time"], errors="coerce")
    df["Total After Bill Discount"] = pd.to_numeric(
        df["Total After Bill Discount"], errors="coerce"
    ).fillna(0)
    df.dropna(subset=["Bill Number", "Order Time"], inplace=True)
    return df


# #################################################################
# --- BAGIAN 3: FUNGSI PEMBANGUN UI (INTERFACE) ---
# #################################################################


def build_sidebar():
    """Menggambar sidebar dan mengembalikan file yang di-upload."""
    with st.sidebar:
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

        gmv_file = st.file_uploader(
            "1. Upload Laporan GMV (Operasional)", type=["xlsx", "csv"]
        )
        st.info("ℹ️ Laporan GMV asli (header di baris ke-10).")

        cogs_file = st.file_uploader(
            "2. Upload Laporan COGS (Menu COGS Report)", type=["xlsx", "csv"]
        )
        st.info("ℹ️ File COGS (header di baris ke-13).")

        waiter_file = st.file_uploader(
            "3. Upload Laporan Waiter (Rekapitulasi Detail)", type=["xlsx", "csv"]
        )
        st.info("ℹ️ Laporan Rekapitulasi Detail (header di baris ke-12).")

    return gmv_file, cogs_file, waiter_file


# --- PERBAIKAN FILTER MINGGUAN ADA DI SINI ---
def build_global_filters(data_gmv, data_cogs, data_waiter):
    """Menggambar filter global dan mengembalikan data yang sudah difilter."""

    # Inisialisasi data yang difilter sebagai data asli
    filtered_gmv = data_gmv
    filtered_cogs = data_cogs
    filtered_waiter = data_waiter

    # Tentukan master tanggal
    master_min_date = pd.Timestamp.now().date()
    master_max_date = pd.Timestamp.now().date()
    filter_source_file = None

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
    except Exception as e:
        st.error(f"Gagal membaca rentang tanggal: {e}")

    if filter_source_file:
        st.subheader("Filter Analisis Global")
        st.info(
            f"Filter global saat ini menggunakan rentang tanggal dari file: **{filter_source_file}**"
        )

        filter_type = st.radio(
            "Pilih rentang waktu analisis untuk Tab 1, 2, dan 3:",
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
            if data_gmv is not None:
                filtered_gmv = data_gmv[
                    data_gmv["Sales Date In"].dt.date == selected_date
                ]
            if data_cogs is not None:
                filtered_cogs = data_cogs[
                    data_cogs["Sales Date"].dt.date == selected_date
                ]
            if data_waiter is not None:
                filtered_waiter = data_waiter[
                    data_waiter["Order Time"].dt.date == selected_date
                ]

        # --- LOGIKA MINGGUAN DIPERBAIKI DI SINI ---
        elif filter_type == "Mingguan":

            # 1. Tentukan default value yang lebih baik (7 hari terakhir)
            default_start_mingguan = master_max_date - pd.to_timedelta(6, unit="d")
            # Pastikan default tidak lebih awal dari data
            if default_start_mingguan < master_min_date:
                default_start_mingguan = master_min_date

            # 2. Ubah label agar lebih jelas
            selected_start_date = st.date_input(
                "Pilih Tanggal Mulai (periode 7 hari)",  # <-- Label diubah
                value=default_start_mingguan,  # <-- Default diubah
                min_value=master_min_date,
                max_value=master_max_date,
                key="global_week_start",  # <-- Key diubah
            )

            # 3. Logika baru: Tanggal mulai = yg dipilih; Tanggal akhir = mulai + 6 hari
            start_date = selected_start_date
            end_date = start_date + pd.to_timedelta(6, unit="d")

            # 4. Pastikan tanggal akhir tidak melebihi data maksimum
            if end_date > master_max_date:
                end_date = master_max_date

            st.info(
                f"Menampilkan data dari {start_date.strftime('%d-%m-%Y')} s.d. {end_date.strftime('%d-%m-%Y')}"
            )

            # 5. Terapkan filter
            if data_gmv is not None:
                filtered_gmv = data_gmv[
                    (data_gmv["Sales Date In"].dt.date >= start_date)
                    & (data_gmv["Sales Date In"].dt.date <= end_date)
                ]
            if data_cogs is not None:
                filtered_cogs = data_cogs[
                    (data_cogs["Sales Date"].dt.date >= start_date)
                    & (data_cogs["Sales Date"].dt.date <= end_date)
                ]
            if data_waiter is not None:
                filtered_waiter = data_waiter[
                    (data_waiter["Order Time"].dt.date >= start_date)
                    & (data_waiter["Order Time"].dt.date <= end_date)
                ]
        # --- AKHIR PERBAIKAN MINGGUAN ---

        elif filter_type == "Bulanan":
            selected_date = st.date_input(
                "Pilih tanggal dalam bulan",
                value=master_max_date,
                min_value=master_min_date,
                max_value=master_max_date,
                key="global_month",
            )
            selected_month = selected_date.month
            selected_year = selected_date.year
            st.info(f"Menampilkan data untuk bulan {selected_date.strftime('%B %Y')}")

            if data_gmv is not None:
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

        elif filter_type == "Tahunan":
            selected_date = st.date_input(
                "Pilih tanggal dalam tahun",
                value=master_max_date,
                min_value=master_min_date,
                max_value=master_max_date,
                key="global_year",
            )
            selected_year = selected_date.year
            st.info(f"Menampilkan data untuk tahun {selected_year}")

            if data_gmv is not None:
                filtered_gmv = data_gmv[
                    data_gmv["Sales Date In"].dt.year == selected_year
                ]
            if data_cogs is not None:
                filtered_cogs = data_cogs[
                    data_cogs["Sales Date"].dt.year == selected_year
                ]
            if data_waiter is not None:
                filtered_waiter = data_waiter[
                    data_waiter["Order Time"].dt.year == selected_year
                ]

        st.markdown("---")

    elif data_gmv is None and data_cogs is None and data_waiter is None:
        st.info("ℹ️ Upload file di sidebar untuk memulai analisis.")
        st.markdown("---")

    return filtered_gmv, filtered_cogs, filtered_waiter


def build_tab1_sales(filtered_gmv):
    """Menggambar semua elemen untuk Tab 1."""
    if filtered_gmv is not None:
        if not filtered_gmv.empty:
            start_date = filtered_gmv["Sales Date In"].min().strftime("%d-%m-%Y")
            end_date = filtered_gmv["Sales Date In"].max().strftime("%d-%m-%Y")
            st.subheader(f"Periode Analisis: {start_date} s.d. {end_date}")

            with st.expander(
                "📈 KPI Kinerja Penjualan (Revenue, ATV, IPB)", expanded=True
            ):
                st.header("📊 KPI Kinerja Penjualan")
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
                ) = get_menu_performance(filtered_gmv)

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

            with st.expander(
                "💳 Analisis Transaksi (Pembayaran & Kunjungan)", expanded=True
            ):
                col7, col8 = st.columns(2)
                if "Payment Method" in filtered_gmv.columns:
                    with col7:
                        st.subheader("💳 Penjualan per Metode Pembayaran")
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
                        st.warning("Kolom 'Payment Method' tidak ditemukan di File 1.")

                if "Visit Purpose" in filtered_gmv.columns:
                    with col8:
                        st.subheader("🏪 Penjualan per Tipe Kunjungan")
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
                        st.warning("Kolom 'Visit Purpose' tidak ditemukan di File 1.")

            st.markdown("---")

            with st.expander(
                "⚙️ KPI Operasional (Jam Sibuk, Hari Sibuk, Durasi Makan)", expanded=True
            ):
                st.header("⚙️ KPI Operasional & Pelanggan")
                avg_time, peak_hours, peak_days_of_week = get_operational_kpi(
                    filtered_gmv
                )

                # --- PERUBAHAN DI SINI ---
                # Hanya tampilkan metrik jika nilainya valid (lebih dari 0)
                if avg_time > 0:
                    st.metric(
                        "⏱️ Rata-rata Durasi Makan (Dine In)", f"{avg_time:.1f} menit"
                    )
                # --- AKHIR PERUBAHAN ---

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

        elif filtered_gmv is not None and filtered_gmv.empty:
            st.warning(
                "Tidak ada data ditemukan di File GMV untuk rentang waktu yang dipilih."
            )
    else:
        st.info(
            "Silakan upload file Laporan GMV (File 1) di sidebar untuk melihat analisis penjualan."
        )


def build_tab2_cogs(filtered_cogs):
    """Menggambar semua elemen untuk Tab 2."""
    if filtered_cogs is not None:
        if not filtered_cogs.empty:
            st.header("💰 Analisis Profitabilitas Menu (COGS)")
            profit_df = analyze_profit(filtered_cogs)

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

            with st.expander("📝 Rincian Profitabilitas per Menu (Tabel)"):
                st.subheader("Rincian Profitabilitas per Menu")
                st.info(
                    "Data ini dijumlahkan (agregasi) HANYA dari file Laporan COGS (sesuai filter waktu yang dipilih)."
                )

                # Format tabel untuk tampilan
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
            "Silakan upload file Laporan COGS (File 2) di sidebar untuk melihat analisis profitabilitas."
        )


def build_tab3_hr(filtered_waiter):
    """Menggambar semua elemen untuk Tab 3."""
    if filtered_waiter is not None:
        if not filtered_waiter.empty:
            st.header("🧑‍🍳 Analisis Kinerja Waiter & Waktu Kunjungan")

            with st.expander("🕒 Analisis Waktu Kunjungan Pelanggan", expanded=True):
                st.subheader("🕒 Waktu Kunjungan Pelanggan")
                time_data = get_peak_time_analysis(filtered_waiter)
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
                    ),
                    use_container_width=True,
                )

            st.markdown("---")

            with st.expander("🏆 Performa Waiter Teratas (Top 10)", expanded=True):
                st.subheader("🏆 Performa Waiter Teratas (Top 10)")
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
                    ),
                    use_container_width=True,
                )

        elif filtered_waiter is not None and filtered_waiter.empty:
            st.warning(
                "Tidak ada data ditemukan di File Rekapitulasi untuk rentang waktu yang dipilih."
            )
    else:
        st.info(
            "Silakan upload file Laporan Rekapitulasi Detail (File 3) di sidebar untuk melihat analisis waiter."
        )


# --- FUNGSI TAB 4 YANG DIPERBARUI ---
def build_tab4_comparison(data_gmv, data_cogs, data_waiter):
    """Menggambar Tab 4 (Perbandingan A/B) dengan tata letak A | Delta | B."""

    st.header("⚖️ Analisis Perbandingan Periodik (A vs B)")
    st.info(
        "Gunakan tab ini untuk membandingkan kinerja antara dua periode (A vs B). "
        "Filter global diabaikan di tab ini."
    )

    # --- Helper Lokal (Hanya untuk Tab 4) ---

    def slice_data_by_period(df, date_col_name, ref_date, comparison_type):
        """
        Memotong dataframe berdasarkan tipe perbandingan dan tanggal acuan.
        Mengembalikan DataFrame yang diiris DAN string caption.
        """
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
            # Logika baru: 7 hari ke DEPAN dari tanggal yang dipilih
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
        """Helper lokal untuk menghitung KPI profit dari irisan data COGS."""
        profit_df = analyze_profit(df_cogs_sliced)  # Memanggil fungsi helper global
        if profit_df.empty:
            return 0, 0, 0, 0
        total_revenue = profit_df["Total Revenue (Rp)"].sum()
        total_cogs = profit_df["Total COGS (Rp)"].sum()
        total_profit = profit_df["Total Profit (Rp)"].sum()
        margin = (total_profit / total_revenue) * 100 if total_revenue > 0 else 0
        return total_revenue, total_cogs, total_profit, margin

    # --- FUNGSI HELPER WAITER INI DIPERBARUI ---
    def get_waiter_kpis(df_waiter_sliced):
        """Helper lokal untuk menghitung KPI dari irisan data Waiter."""
        if df_waiter_sliced is None or df_waiter_sliced.empty:
            return 0, 0, 0  # <--- PERUBAHAN: Kembalikan 3 nilai

        # 1. Hitung total penjualan dan transaksi (seperti sebelumnya)
        bill_df = (
            df_waiter_sliced.groupby("Bill Number")
            .agg(Total_Sales=("Total After Bill Discount", "sum"))
            .reset_index()
        )
        total_penjualan = bill_df["Total_Sales"].sum()
        total_transaksi = bill_df["Bill Number"].nunique()

        # 2. Hitung jumlah waiter unik (BARU)
        #    Kita hitung dari data asli (df_waiter_sliced) sebelum di-group
        #    dan filter "Tidak Diketahui" (yang berasal dari NaN/None)
        cleaned_waiters = df_waiter_sliced["Waiter"].fillna("Tidak Diketahui")
        unique_waiters_count = cleaned_waiters[
            cleaned_waiters != "Tidak Diketahui"
        ].nunique()

        # 3. Hitung rata-rata penjualan per waiter (BARU)
        avg_sales_per_waiter = (
            total_penjualan / unique_waiters_count if unique_waiters_count > 0 else 0
        )

        return total_penjualan, total_transaksi, avg_sales_per_waiter  # <--- PERUBAHAN

    # --- Akhir Helper Lokal ---

    # Tentukan Master Date Range
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

    # --- UI UTAMA: Filter dan 3 Kolom ---

    comparison_type = st.selectbox(
        "Pilih Tipe Perbandingan:", ["Harian", "Mingguan", "Bulanan", "Tahunan"]
    )
    st.markdown("---")

    # PERBAIKAN TATA LETAK: Ganti nama variabel kolom filter
    filter_col_A, filter_col_Delta, filter_col_B = st.columns([0.4, 0.2, 0.4])

    # Tentukan default pintar untuk Periode B
    default_A_date = master_max_date
    default_B_date = master_min_date
    try:
        if comparison_type == "Harian":
            default_B_date = default_A_date - pd.to_timedelta(1, unit="d")
        elif comparison_type == "Mingguan":
            default_B_date = default_A_date - pd.to_timedelta(
                7, unit="d"
            )  # Defaultnya 7 hari sebelumnya
        elif comparison_type == "Bulanan":
            default_B_date = default_A_date - pd.DateOffset(months=1)
        elif comparison_type == "Tahunan":
            default_B_date = default_A_date - pd.DateOffset(years=1)

        # Pastikan default B tidak lebih awal dari data
        if pd.Timestamp(default_B_date) < pd.Timestamp(master_min_date):
            default_B_date = master_min_date
        elif isinstance(default_B_date, pd.Timestamp):
            default_B_date = default_B_date.date()
    except Exception:
        default_B_date = master_min_date

    # --- Setup Judul Kolom dan Date Picker ---
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

    # --- Slicing Data ---
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

    # Tampilkan caption
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

    st.markdown("---")  # Pemisah tebal setelah setup

    # --- Hitung semua KPI ---
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

    # PERUBAHAN: Disesuaikan untuk 3 nilai
    kpi_A_waiter = get_waiter_kpis(waiter_A) if data_waiter is not None else (0, 0, 0)
    kpi_B_waiter = get_waiter_kpis(waiter_B) if data_waiter is not None else (0, 0, 0)

    # --- Tampilkan Metrik Secara Berdampingan ---

    if data_gmv is not None:
        st.markdown("##### 📊 Kinerja Penjualan (dari File 1: GMV)")
        with st.container(border=True):
            # PERBAIKAN: Buat kolom BARU di sini
            col_A, col_Delta, col_B = st.columns([0.4, 0.2, 0.4])

            # Baris 1: Total Pendapatan
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

            # Baris 2: Total Transaksi
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

            # Baris 3: ATV
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

            # Baris 4: Item per Transaksi (IPB)
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

            # Baris 5: Total Diskon (Lower is better)
            with col_A:
                st.metric("Total Diskon", format_rupiah(kpi_A_gmv["Total Diskon"]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_gmv["Total Diskon"],
                    kpi_B_gmv["Total Diskon"],
                    format_rupiah,
                    False,
                )  # False = lebih rendah lebih baik
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total Diskon", format_rupiah(kpi_B_gmv["Total Diskon"]))

        st.markdown("---")  # Pemisah antar grup

    if data_cogs is not None:
        st.markdown("##### 💰 Kinerja Profitabilitas (dari File 2: COGS)")
        with st.container(border=True):
            # PERBAIKAN: Buat kolom BARU di sini
            col_A, col_Delta, col_B = st.columns([0.4, 0.2, 0.4])

            # Baris 1: Total Revenue (COGS)
            with col_A:
                st.metric("Total Revenue (COGS)", format_rupiah(kpi_A_cogs[0]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_cogs[0], kpi_B_cogs[0], format_rupiah, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total Revenue (COGS)", format_rupiah(kpi_B_cogs[0]))

            # Baris 2: Total COGS (Lower is better)
            with col_A:
                st.metric("Total COGS", format_rupiah(kpi_A_cogs[1]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_cogs[1], kpi_B_cogs[1], format_rupiah, False
                )  # False = lebih rendah lebih baik
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total COGS", format_rupiah(kpi_B_cogs[1]))

            # Baris 3: Total Profit
            with col_A:
                st.metric("Total Profit", format_rupiah(kpi_A_cogs[2]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_cogs[2], kpi_B_cogs[2], format_rupiah, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total Profit", format_rupiah(kpi_B_cogs[2]))

            # Baris 4: Margin (%)
            with col_A:
                st.metric("Margin Profit (%)", format_persen(kpi_A_cogs[3]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_cogs[3], kpi_B_cogs[3], format_persen, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Margin Profit (%)", format_persen(kpi_B_cogs[3]))

        st.markdown("---")  # Pemisah antar grup

    if data_waiter is not None:
        st.markdown("##### 🧑‍🍳 Kinerja SDM (dari File 3: Waiter)")
        with st.container(border=True):
            # PERBAIKAN: Buat kolom BARU di sini
            col_A, col_Delta, col_B = st.columns([0.4, 0.2, 0.4])

            # Baris 1: Total Penjualan (SDM)
            with col_A:
                st.metric("Total Penjualan (SDM)", format_rupiah(kpi_A_waiter[0]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_waiter[0], kpi_B_waiter[0], format_rupiah, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total Penjualan (SDM)", format_rupiah(kpi_B_waiter[0]))

            # Baris 2: Total Transaksi (SDM)
            with col_A:
                st.metric("Total Transaksi (SDM)", format_angka_bulat(kpi_A_waiter[1]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_waiter[1], kpi_B_waiter[1], format_angka_bulat, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total Transaksi (SDM)", format_angka_bulat(kpi_B_waiter[1]))

            # --- BARIS 3: KPI BARU (Rata-rata per Waiter) ---
            with col_A:
                st.metric("Rata-rata / Waiter", format_rupiah(kpi_A_waiter[2]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_waiter[2], kpi_B_waiter[2], format_rupiah, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Rata-rata / Waiter", format_rupiah(kpi_B_waiter[2]))


# #################################################################
# --- BAGIAN 3.5: FUNGSI PEMBANGUN UI (TAB 5 - VERSI CANGGIH) ---
# #################################################################


def build_tab5_forecast(data_gmv):
    """Menggambar Tab 5 (Peramalan Detail) menggunakan Prophet."""
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
        # Ganti nama kolom agar sesuai standar Prophet: 'ds' (tanggal) dan 'y' (nilai)
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
    # Prophet butuh setidaknya 2 siklus penuh (2 minggu) untuk deteksi musiman
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
    # Tentukan tanggal terakhir data
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
        # Inisialisasi model
        # Kita set 'daily_seasonality=False' karena data kita harian (bukan per jam)
        # 'weekly_seasonality=True' sangat penting untuk F&B (pola weekend)
        model = Prophet(weekly_seasonality=True, daily_seasonality=False)

        # Latih model
        model.fit(daily_sales)

        # Buat dataframe 'masa depan'
        future_df = model.make_future_dataframe(periods=forecast_days)

        # Buat peramalan
        forecast_df = model.predict(future_df)

        # --- 5. Tampilkan Hasil ---
        st.markdown("---")
        st.subheader(f"📈 Grafik Peramalan {forecast_days} Hari ke Depan")

        # Gunakan plot_plotly untuk grafik interaktif
        fig1 = plot_plotly(model, forecast_df)
        fig1.update_layout(
            title="Peramalan Penjualan (Aktual vs. Prediksi)",
            xaxis_title="Tanggal",
            yaxis_title="Estimasi Penjualan (Rp)",
        )
        st.plotly_chart(fig1, use_container_width=True)

        with st.expander("Lihat Data Tabel Peramalan (Angka Detail)"):
            # Tampilkan hanya data ramalan di masa depan
            future_data_table = forecast_df[forecast_df["ds"] > last_date]

            # Format kolom untuk dibaca
            future_data_table_display = future_data_table[
                ["ds", "yhat", "yhat_lower", "yhat_upper"]
            ].copy()

            # 'yhat' = Nilai Ramalan
            # 'yhat_lower' = Batas Bawah (Optimis)
            # 'yhat_upper' = Batas Atas (Pesimis)

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

        # Buat plot komponen (juga interaktif)
        fig2 = plot_components_plotly(model, forecast_df)
        st.plotly_chart(fig2, use_container_width=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan saat melatih model Prophet: {e}")
        st.exception(e)


# #################################################################
# --- BAGIAN 3.6: FUNGSI PEMBANGUN UI (TAB 6 - DENGAN RAMALAN PROPHET) ---
# #################################################################


def build_tab6_target(data_gmv):
    """Menggambar Tab 6 (Pencapaian Target) dengan perbandingan ramalan Prophet."""
    st.header("🎯 Pencapaian Target & Proyeksi (Dinamis)")

    if data_gmv is None or data_gmv.empty:
        st.warning(
            "Silakan upload file Laporan GMV (File 1) di sidebar untuk melihat analisis target."
        )
        return

    # --- 1. Dapatkan Input Target dari User & Setup Tanggal ---
    st.subheader("Pengaturan Analisis")

    # Helper map hari
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

    # Input Target Tampilan Baru
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

    # --- 2. Kalkulasi Data Bulan yang Dipilih ---
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
        total_hari_bulan_ini = max_date_in_selected_month.days_in_month

        is_latest_month = (active_year == latest_data_date_in_full_dataset.year) and (
            active_month == latest_data_date_in_full_dataset.month
        )

        if is_latest_month:
            hari_berjalan = max_date_in_selected_month.day
            st.info(
                f"Menganalisis bulan berjalan ({active_month_name}). Data terdeteksi sampai tanggal {hari_berjalan}."
            )
        else:
            hari_berjalan = total_hari_bulan_ini
            st.success(
                f"Menampilkan ulasan performa bulan lalu ({active_month_name}) yang telah selesai."
            )

        sisa_hari = total_hari_bulan_ini - hari_berjalan

        penjualan_saat_ini = data_bulan_aktif["Total After Bill Discount"].sum()
        pencapaian_persen = (penjualan_saat_ini / target_bulanan) * 100
        rata_rata_harian_total = penjualan_saat_ini / hari_berjalan
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
            avg_sales_weekend = avg_sales_weekday

        weekend_weight = avg_sales_weekend / avg_sales_weekday

        proyeksi_akhir_bulan = penjualan_saat_ini
        rdr_weekday = 0
        rdr_weekend = 0
        proyeksi_prophet = 0  # Inisialisasi

        if sisa_hari > 0:
            tanggal_mulai_sisa = max_date_in_selected_month + pd.Timedelta(days=1)
            tanggal_akhir_bulan = max_date_in_selected_month.replace(
                day=total_hari_bulan_ini
            )

            sisa_tanggal_df = pd.DataFrame(
                pd.date_range(start=tanggal_mulai_sisa, end=tanggal_akhir_bulan),
                columns=["Tanggal"],
            )
            sisa_tanggal_df["Nama Hari"] = (
                sisa_tanggal_df["Tanggal"].dt.day_name().map(day_map)
            )

            sisa_weekdays_count = sisa_tanggal_df["Nama Hari"].isin(weekdays_def).sum()
            sisa_weekends_count = sisa_tanggal_df["Nama Hari"].isin(weekends_def).sum()

            # 1. Kalkulasi Proyeksi (Ramalan) Cerdas (Weighted)
            proyeksi_sales_sisa = (sisa_weekdays_count * avg_sales_weekday) + (
                sisa_weekends_count * avg_sales_weekend
            )
            proyeksi_akhir_bulan += proyeksi_sales_sisa

            # 2. Kalkulasi Target Harian Dinamis (RDR Dinamis)
            pembagi = sisa_weekdays_count + (sisa_weekends_count * weekend_weight)
            if pembagi > 0:
                rdr_weekday = sales_dibutuhkan / pembagi
                rdr_weekend = rdr_weekday * weekend_weight

            # 3. [BARU] Kalkulasi Ramalan Prophet
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

                if len(prophet_data) >= 7:  # Butuh min 7 hari
                    with st.spinner("Menjalankan model ramalan Prophet..."):
                        model_prophet = Prophet(
                            weekly_seasonality=True,
                            daily_seasonality=False,
                            changepoint_prior_scale=0.1,
                        )
                        model_prophet.fit(prophet_data)
                        future_df_prophet = model_prophet.make_future_dataframe(
                            periods=sisa_hari
                        )
                        forecast_df_prophet = model_prophet.predict(future_df_prophet)
                        ramalan_sisa_hari = forecast_df_prophet.iloc[-sisa_hari:][
                            "yhat"
                        ].sum()
                        proyeksi_prophet = penjualan_saat_ini + ramalan_sisa_hari

                else:
                    proyeksi_prophet = proyeksi_akhir_bulan  # Fallback
            except Exception as e_prophet:
                st.warning(
                    f"Gagal menjalankan ramalan Prophet (data mungkin kurang): {e_prophet}"
                )
                proyeksi_prophet = proyeksi_akhir_bulan  # Fallback

        else:  # Bulan sudah selesai
            proyeksi_akhir_bulan = penjualan_saat_ini
            proyeksi_prophet = penjualan_saat_ini

        proyeksi_vs_target_persen = proyeksi_akhir_bulan / target_bulanan
        kekurangan_proyeksi = target_bulanan - proyeksi_akhir_bulan

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

    # --- 4. Tampilkan Metrik KPI ---
    st.subheader("📈 Gambaran Besar (Pencapaian)")

    col1, col2, col3 = st.columns(3)
    col1.metric(
        label=f"Pencapaian per {max_date_in_selected_month.strftime('%d-%m-%Y')}",
        value=f"{pencapaian_persen:,.1f}%",
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
            delta=f"{pencapaian_persen:,.1f}% dari Target",
            delta_color=status_color,
        )
        col3.metric(
            label="Selisih Target",
            value=format_rupiah(penjualan_saat_ini - target_bulanan),
            help="Hasil akhir penjualan dikurangi target bulanan.",
        )

    # --- [BLOK BARU] HASIL RAMALAN DI BAWAH ---
    st.subheader("🔮 Perbandingan Model Ramalan")
    st.write(
        "Membandingkan Proyeksi Cerdas (berbasis rata-rata) dengan Proyeksi Prophet (berbasis tren & musiman)."
    )

    with st.container(border=True):
        col_p1, col_p2 = st.columns(2)

        # Hitung status untuk Proyeksi Prophet
        proyeksi_prophet_persen = proyeksi_prophet / target_bulanan
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
    # --- AKHIR BLOK BARU ---

    st.markdown("---")
    st.subheader("💡 Rencana Aksi & Diagnostik")

    col_a, col_b = st.columns(2)
    if sisa_hari > 0:
        col_a.metric(
            label="Target Harian Sisa (Weekday)",
            value=format_rupiah(rdr_weekday),
            help="Target penjualan harian baru Anda untuk hari Sen, Sel, Rab, Kam.",
        )
        col_b.metric(
            label="Target Harian Sisa (Weekend)",
            value=format_rupiah(rdr_weekend),
            help="Target penjualan harian baru Anda untuk hari Jum, Sab, Min.",
        )
    else:
        col_a.info("Bulan ini telah selesai. Tidak ada target sisa hari.")

    col_d, col_e = st.columns(2)
    col_d.metric(
        label="Rata-rata Weekday (Sen-Kam)",
        value=format_rupiah(avg_sales_weekday),
        delta=format_rupiah(avg_sales_weekday - rdr_weekday) if sisa_hari > 0 else None,
        delta_color="normal" if (avg_sales_weekday - rdr_weekday) > 0 else "inverse",
    )
    col_e.metric(
        label="Rata-rata Weekend (Jum-Min)",
        value=format_rupiah(avg_sales_weekend),
        delta=format_rupiah(avg_sales_weekend - rdr_weekend) if sisa_hari > 0 else None,
        delta_color="normal" if (avg_sales_weekend - rdr_weekend) > 0 else "inverse",
    )

    # --- 5. Tampilkan Grafik Tren Harian vs Target ---
    try:
        # --- Grafik 1: Proyeksi vs. Target Dinamis ---
        st.subheader("Grafik 1: Tren Penjualan & Target Sisa Bulan")

        actual_data_df = (
            daily_sales_agg.groupby(daily_sales_agg["Sales Date In"].dt.date)[
                "Total After Bill Discount"
            ]
            .sum()
            .reset_index()
        )
        actual_data_df.rename(
            columns={
                "Sales Date In": "Tanggal",
                "Total After Bill Discount": "Penjualan",
            },
            inplace=True,
        )
        actual_data_df["Tipe"] = "1 - Penjualan Aktual"
        actual_data_df["Tanggal"] = pd.to_datetime(actual_data_df["Tanggal"])

        plot_df_1 = actual_data_df

        if sisa_hari > 0:
            sisa_tanggal_df["Penjualan"] = sisa_tanggal_df["Nama Hari"].apply(
                lambda x: rdr_weekend if x in weekends_def else rdr_weekday
            )
            sisa_tanggal_df["Tipe"] = "2 - Target Dinamis"
            sisa_tanggal_df.rename(columns={"Tanggal": "Tanggal"}, inplace=True)
            plot_df_1 = pd.concat(
                [actual_data_df, sisa_tanggal_df[["Tanggal", "Penjualan", "Tipe"]]]
            )
            st.write(
                "Grafik ini menunjukkan penjualan aktual Anda (biru) dilanjutkan dengan target dinamis (merah putus-putus) yang harus Anda capai di sisa hari."
            )
        else:
            st.write(
                "Grafik ini menunjukkan penjualan aktual Anda (biru) selama bulan yang telah selesai."
            )

        line_chart = (
            alt.Chart(plot_df_1)
            .mark_line(point=True)
            .encode(
                x=alt.X("Tanggal:T", title="Tanggal"),
                y=alt.Y("Penjualan:Q", title="Penjualan (Rp)"),
                color=alt.Color("Tipe:N", title="Legenda", sort="ascending"),
                strokeDash=alt.StrokeDash("Tipe:N", legend=None, sort="ascending"),
                tooltip=[
                    alt.Tooltip("Tanggal", format="%d-%m-%Y"),
                    alt.Tooltip("Penjualan", format=",.0f"),
                    "Tipe",
                ],
            )
        )
        st.altair_chart(line_chart.interactive(), use_container_width=True)

        # --- Grafik 2: Diagnostik Performa Rata-rata vs. Target ---
        st.subheader("Grafik 2: Diagnostik Performa Rata-rata vs. Target")

        daily_sales_with_dayname = (
            data_bulan_aktif.groupby(
                [data_bulan_aktif["Sales Date In"].dt.date, "Nama Hari"]
            )["Total After Bill Discount"]
            .sum()
            .reset_index()
        )

        avg_sales_by_day_raw = (
            daily_sales_with_dayname.groupby("Nama Hari")["Total After Bill Discount"]
            .mean()
            .reset_index()
        )

        master_days_df = pd.DataFrame({"Nama Hari": day_sort_order})

        avg_sales_by_day = pd.merge(
            master_days_df, avg_sales_by_day_raw, on="Nama Hari", how="left"
        ).fillna(0)

        avg_sales_by_day.rename(
            columns={"Total After Bill Discount": "Jumlah"}, inplace=True
        )
        avg_sales_by_day["Tipe"] = "1 - Rata-rata Aktual"

        target_per_hari_df = pd.DataFrame(
            {
                "Nama Hari": day_sort_order,
                "Jumlah": [
                    rdr_weekday if day in weekdays_def else rdr_weekend
                    for day in day_sort_order
                ],
                "Tipe": "2 - Target Harian Baru (RDR)",
            }
        )

        plot_df_2 = pd.concat([avg_sales_by_day, target_per_hari_df])

        base = alt.Chart(plot_df_2).encode(
            x=alt.X("Nama Hari:N", title="Hari", sort=day_sort_order),
            y=alt.Y("Jumlah:Q", title="Penjualan (Rp)"),
        )

        bar_chart = (
            base.transform_filter(alt.datum["Tipe"] == "1 - Rata-rata Aktual")
            .mark_bar()
            .encode(
                color=alt.Color("Nama Hari:N", title="Hari"),
                tooltip=[
                    "Nama Hari",
                    alt.Tooltip("Jumlah", format=",.0f", title="Rata-rata Aktual"),
                ],
            )
        )

        line_chart_target = (
            base.transform_filter(alt.datum["Tipe"] == "2 - Target Harian Baru (RDR)")
            .mark_line(point=True, color="red", strokeDash=[5, 5])
            .encode(
                tooltip=[
                    "Nama Hari",
                    alt.Tooltip("Jumlah", format=",.0f", title="Target Harian Baru"),
                ]
            )
        )

        if sisa_hari > 0:
            st.write(
                "Grafik ini membandingkan performa rata-rata Anda per hari (Bars) dengan Target Harian Dinamis (garis merah) yang baru."
            )
        else:
            st.write(
                "Grafik ini menampilkan performa rata-rata Anda per hari (Bars) untuk bulan yang telah selesai."
            )

        st.altair_chart(
            (bar_chart + line_chart_target).interactive(), use_container_width=True
        )

    except Exception as e:
        st.error(f"Gagal membuat grafik tren harian: {e}")
        st.exception(e)


def build_footer():
    """Menggambar footer."""
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


# #################################################################
# --- BAGIAN 4: EKSEKUSI APLIKASI UTAMA ---
# #################################################################


def main():
    """Fungsi utama untuk menjalankan aplikasi Streamlit."""
    st.title("📈 Data Driven Analyst Dashboard KPI - Milky Way Lippo Mall Puri")

    # 1. Gambar Sidebar dan dapatkan file
    gmv_file, cogs_file, waiter_file = build_sidebar()

    # 2. Muat data (akan menggunakan cache jika file sama)
    data_gmv = load_data_gmv(gmv_file)
    data_cogs = load_cogs_data(cogs_file)
    data_waiter = load_data_waiter(waiter_file)

    # 3. Gambar Filter Global dan dapatkan data yang sudah difilter
    filtered_gmv, filtered_cogs, filtered_waiter = build_global_filters(
        data_gmv, data_cogs, data_waiter
    )

    # 4. Buat struktur Tab
    tab_titles = [
        "📊 Analisis Penjualan (KPI)",
        "💰 Analisis COGS & Profit",
        "🧑‍🍳 Analisis SDM & Waktu",
        "⚖️ Analisis Perbandingan Periodik",
        "🔮 Peramalan Tren",
        "🎯 Target & Proyeksi",  # <-- PERUBAHAN 1: Tambah judul Tab 6
    ]
    # <-- PERUBAHAN 2: Tambah 'tab6'
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(tab_titles)

    # 5. Isi setiap Tab dengan data yang relevan
    with tab1:
        build_tab1_sales(filtered_gmv)

    with tab2:
        build_tab2_cogs(filtered_cogs)

    with tab3:
        build_tab3_hr(filtered_waiter)

    with tab4:
        # Tab 4 menggunakan data *asli* (bukan filtered) untuk perbandingan internal
        build_tab4_comparison(data_gmv, data_cogs, data_waiter)

    with tab5:
        # Tab 5 menggunakan data *asli* (bukan filtered) untuk peramalan tren
        build_tab5_forecast(data_gmv)

    # <-- PERUBAHAN 3: Tambah blok 'with tab6'
    with tab6:
        # Tab 6 juga menggunakan data asli (bukan filtered)
        build_tab6_target(data_gmv)

    # 6. Gambar Footer
    build_footer()

    # --- PERUBAHAN DI SINI: SEMUA NOTIFIKASI HILANG SETELAH 5 DETIK ---
    st.markdown(
        """
        <script>
        function hideAllNotices() {
            // PERUBAHAN: Menghapus :not([data-testid="stError"]) agar SEMUA notifikasi ikut hilang
            const notices = document.querySelectorAll(
                '.stAlert[data-baseweb="alert"]:not(.fading-out)'
            );

            notices.forEach(function(notice) {
                // Tandai sebagai 'sedang diproses' agar tidak dipilih lagi
                notice.classList.add('fading-out');

                // PERUBAHAN: Waktu tunggu diubah menjadi 5000ms (5 detik)
                setTimeout(() => {
                    // Mulai transisi fade-out
                    notice.style.transition = 'opacity 0.5s ease-out';
                    notice.style.opacity = '0';
                    
                    // Setelah transisi selesai, sembunyikan elemen
                    setTimeout(() => {
                        notice.style.display = 'none';
                    }, 500); // 500ms = durasi transisi
                }, 5000); // <-- UBAH JADI 5000 (5 DETIK)
            });
        }

        // Jalankan fungsi ini setiap 1 detik untuk menangkap notifikasi baru
        setInterval(hideAllNotices, 1000);
        </script>
        """,
        unsafe_allow_html=True,
    )
    # --- AKHIR PERUBAHAN ---


if __name__ == "__main__":
    # Muat CSS eksternal (jika ada)
    load_css("style.css")
    # Jalankan aplikasi utama
    main()
