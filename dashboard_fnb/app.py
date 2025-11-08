### app.py ###
import time
import pandas as pd
import streamlit as st
import openpyxl
import altair as alt
import numpy as np
from prophet import Prophet
from prophet.plot import plot_plotly, plot_components_plotly
import re
import sqlite3  # <-- Library baru untuk database

# --- Konfigurasi Halaman ---
st.set_page_config(
    page_title="Data Driven Analyst Specialyst FnB",
    layout="wide",
    initial_sidebar_state="expanded",
)


# --- FUNGSI CSS DAN FORMATTING (Sama seperti kode asli) ---
def load_css(file_name):
    try:
        with open(file_name) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.error(f"File CSS '{file_name}' tidak ditemukan.")


st.markdown(
    """
<style>
.stTabs [data-baseweb="tab-list"] button {
    font-size: 1.5rem; padding: 10px 15px;
}
</style>
""",
    unsafe_allow_html=True,
)


# --- SEMUA FUNGSI HELPER (Sama seperti kode asli) ---
# (calculate_sales_kpi, get_payment_analysis, get_menu_performance,
#  get_operational_kpi, analyze_profit, get_peak_time_analysis,
#  get_waiter_performance, create_horizontal_bar_chart,
#  create_vertical_bar_chart, calculate_delta, format_rupiah, dll.)
# ... (Salin semua fungsi helper Anda dari kode asli ke sini) ...
def format_rupiah(amount):
    return f"Rp {amount:,.0f}".replace(",", ".")


def format_angka_bulat(number):
    return f"{number:.0f}"


def format_persen(number):
    return f"{number:,.1f}%"


def clean_payment_method(method_str):
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
    menu_sales = df[~df["Menu"].str.contains("PACKAGE", na=False, case=False)]
    menu_sales = menu_sales[menu_sales["Price (Net)"] > 0]
    top_selling_categories = pd.DataFrame(columns=["Menu Category", "Qty"])
    top_grossing_categories = pd.DataFrame(
        columns=["Menu Category", "Total Nett Sales"]
    )
    menu_sales_cat_df = pd.DataFrame()
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
        menu_sales_cat_df = menu_sales_cat.copy()
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
        menu_sales_cat_df,
    )


def get_operational_kpi(df):
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
    chart = (
        alt.Chart(data)
        .mark_bar()
        .encode(
            x=alt.X(
                f"{x_col}:Q", title=x_title, axis=alt.Axis(orient="top", format="~s")
            ),
            y=alt.Y(
                f"{y_col}:N",
                title=y_title,
                sort=sort_order,
                axis=alt.Axis(labelLimit=300),
            ),
            color=alt.Color(f"{y_col}:N", title=y_title, legend=None),
            tooltip=[
                alt.Tooltip(y_col, title=y_title),
                alt.Tooltip(x_col, title=x_title, format=",.0f"),
            ],
        )
        .properties(padding={"top": 5})
    )
    return chart


def create_vertical_bar_chart(
    data, x_col, y_col, x_title, y_title, x_type="N", sort_order=None
):
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
        .properties(padding={"top": 20})
    )
    return chart


def calculate_delta(value_A, value_B, formatter_func, higher_is_better=True):
    delta_abs = value_A - value_B
    if formatter_func == format_rupiah:
        delta_abs_formatted = formatter_func(delta_abs)
    elif formatter_func == format_angka_bulat:
        delta_abs_formatted = formatter_func(delta_abs)
    else:
        delta_abs_formatted = f"{delta_abs:,.2f}"
    delta_pct_str = ""
    delta_color = "off"
    if value_B != 0:
        delta_pct = (delta_abs / value_B) * 100
        arrow = "🔼" if delta_pct > 0 else "🔽"
        delta_pct_str = f"{arrow} {delta_pct:.1f}%"
        if delta_abs > 0:
            delta_color = "normal" if higher_is_better else "inverse"
        elif delta_abs < 0:
            delta_color = "inverse" if higher_is_better else "normal"
    elif value_A != 0:
        delta_pct_str = "🔼 100% +"
        delta_color = "normal" if higher_is_better else "inverse"
    else:
        delta_pct_str = "-"
    return delta_abs_formatted, delta_pct_str, delta_color


# #################################################################
# --- BAGIAN 2: FUNGSI PEMUATAN DATA (BARU) ---
# #################################################################

DB_NAME = "fnb_data.db"


@st.cache_resource
def get_db_connection():
    """Membuat koneksi ke database SQLite."""
    try:
        conn = sqlite3.connect(DB_NAME, check_same_thread=False)
        return conn
    except sqlite3.Error as e:
        st.error(f"Error koneksi ke database '{DB_NAME}': {e}")
        st.error(
            "Pastikan Anda sudah menjalankan 'python data_loader.py' terlebih dahulu."
        )
        return None


@st.cache_data
def load_data_from_db(_conn, table_name, date_col, start_date=None, end_date=None):
    """
    Fungsi generik untuk memuat data dari tabel dengan filter tanggal.
    Parameter _conn ada untuk trik caching Streamlit.
    """
    if _conn is None:
        return pd.DataFrame()

    query = f"SELECT * FROM {table_name}"
    params = ()

    # Jika start_date dan end_date diberikan, tambahkan filter WHERE
    if start_date and end_date:
        # Konversi date object ke string timestamp
        start_str = pd.to_datetime(start_date).strftime("%Y-%m-%d 00:00:00")
        end_str = pd.to_datetime(end_date).strftime("%Y-%m-%d 23:59:59")

        query += f' WHERE "{date_col}" BETWEEN ? AND ?'
        params = (start_str, end_str)

    try:
        df = pd.read_sql(query, _conn, params=params)

        # Konversi kolom tanggal kembali ke datetime object
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col])

        # Penanganan khusus untuk kolom tanggal lain jika ada
        if table_name == "gmv" and "Sales Date Out" in df.columns:
            df["Sales Date Out"] = pd.to_datetime(df["Sales Date Out"], errors="coerce")

        return df

    except pd.errors.DatabaseError as e:
        if "no such table" in str(e):
            st.error(f"Tabel '{table_name}' tidak ditemukan di database.")
            st.error("Pastikan 'data_loader.py' sudah dijalankan dan berhasil.")
        else:
            st.error(f"Error mengambil data dari {table_name}: {e}")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Error data {table_name}: {e}")
        return pd.DataFrame()


@st.cache_data
def get_full_date_range(_conn):
    """Mendapatkan rentang tanggal min/max dari semua tabel."""
    if _conn is None:
        return pd.Timestamp.now().date(), pd.Timestamp.now().date()

    tables_info = {"gmv": "Sales Date In", "cogs": "Sales Date", "waiter": "Order Time"}
    min_dates = []
    max_dates = []

    for table, date_col in tables_info.items():
        try:
            min_q = f'SELECT MIN("{date_col}") FROM {table}'
            max_q = f'SELECT MAX("{date_col}") FROM {table}'
            min_d = pd.read_sql(min_q, _conn).iloc[0, 0]
            max_d = pd.read_sql(max_q, _conn).iloc[0, 0]
            if min_d:
                min_dates.append(pd.to_datetime(min_d))
            if max_d:
                max_dates.append(pd.to_datetime(max_d))
        except Exception:
            pass  # Abaikan jika tabel tidak ada

    if not min_dates or not max_dates:
        return pd.Timestamp.now().date(), pd.Timestamp.now().date()

    return min(min_dates).date(), max(max_dates).date()


# #################################################################
# --- BAGIAN 3: FUNGSI PEMBANGUN UI (INTERFACE) ---
# #################################################################


def build_sidebar():
    """Menggambar sidebar (TANPA file uploader)."""
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
        st.header("Database Mode")
        st.info("ℹ️ Data dimuat langsung dari database. Filter ada di halaman utama.")
        st.warning(
            "Untuk memperbarui data, jalankan ulang skrip `data_loader.py` lalu refresh halaman ini."
        )


def build_global_filters(master_min_date, master_max_date):
    """Menggambar filter global dan mengembalikan data yang sudah difilter."""

    start_date_filter = master_min_date
    end_date_filter = master_max_date

    if master_min_date == master_max_date:
        st.info("Filter tidak tersedia. Hanya ada satu hari data di database.")
        st.markdown("---")
        return start_date_filter, end_date_filter

    st.subheader("Filter Analisis Global")
    st.info(f"Filter global saat ini menggunakan data dari database.")

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
        start_date_filter = selected_date
        end_date_filter = selected_date

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
        start_date_filter = selected_start_date
        end_date_filter = start_date_filter + pd.to_timedelta(6, unit="d")
        if end_date_filter > master_max_date:
            end_date_filter = master_max_date
        st.info(
            f"Menampilkan data dari {start_date_filter.strftime('%d-%m-%Y')} s.d. {end_date_filter.strftime('%d-%m-%Y')}"
        )

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
        start_date_filter = pd.Timestamp(
            year=selected_year, month=selected_month, day=1
        ).date()
        end_of_month = pd.Timestamp(start_date_filter) + pd.offsets.MonthEnd(0)
        end_date_filter = end_of_month.date()

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
        start_date_filter = pd.Timestamp(year=selected_year, month=1, day=1).date()
        end_date_filter = pd.Timestamp(year=selected_year, month=12, day=31).date()

    st.markdown("---")

    # Kembalikan tanggal yang dipilih
    if filter_type == "Semua Periode":
        return master_min_date, master_max_date

    return start_date_filter, end_date_filter


# --- SEMUA FUNGSI BUILD TAB (Sama seperti kode asli) ---
# (build_tab1_sales, build_tab2_cogs, build_tab3_hr,
#  build_tab4_comparison, build_tab5_forecast, build_tab6_target,
#  build_tab7_pelanggan, build_footer)
# ... (Salin semua fungsi build_tabX Anda dari kode asli ke sini) ...
def build_tab1_sales(filtered_gmv):
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
            (
                top_selling,
                top_grossing,
                top_sell_cat,
                top_gross_cat,
                bottom_selling,
                bottom_grossing,
                menu_sales_cat_df,
            ) = get_menu_performance(filtered_gmv)
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
                        help="Centang untuk menampilkan semua kategori",
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
                            title=title_grafik_atas, height=chart_height_kategori
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
            st.markdown("---")
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
                        "Metode Pembayaran",
                    )
                    st.altair_chart(chart, use_container_width=True)
                st.markdown("---")
                if "Visit Purpose" in filtered_gmv.columns:
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
        elif filtered_gmv is not None and filtered_gmv.empty:
            st.warning(
                "Tidak ada data ditemukan di File GMV untuk rentang waktu yang dipilih."
            )
    else:
        st.info("Data GMV tidak ditemukan di database.")


def build_tab2_cogs(filtered_cogs):
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
                st.info("Data ini dijumlahkan (agregasi) HANYA dari file Laporan COGS.")
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
        st.info("Data COGS tidak ditemukan di database.")


def build_tab3_hr(filtered_waiter):
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
        st.info("Data Waiter tidak ditemukan di database.")


def build_tab4_comparison(
    data_gmv, data_cogs, data_waiter, master_min_date, master_max_date
):
    st.header("⚖️ Analisis Perbandingan Periodik (A vs B)")
    st.info(
        "Gunakan tab ini untuk membandingkan kinerja antara dua periode (A vs B). Filter global diabaikan di tab ini."
    )

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

    has_data = (
        (data_gmv is not None) or (data_cogs is not None) or (data_waiter is not None)
    )
    if not has_data:
        st.warning(
            "Data tidak ditemukan di database untuk memulai analisis perbandingan."
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
        elif comparison_type == "Tahunan":
            default_B_date = default_A_date - pd.DateOffset(years=1)
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
    if data_gmv is not None:
        st.markdown("##### 📊 Kinerja Penjualan (dari File 1: GMV)")
        with st.container(border=True):
            col_A, col_Delta, col_B = st.columns([0.35, 0.3, 0.35])
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
            col_A, col_Delta, col_B = st.columns([0.35, 0.3, 0.35])
            with col_A:
                st.metric("Total Revenue (COGS)", format_rupiah(kpi_A_cogs[0]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_cogs[0], kpi_B_cogs[0], format_rupiah, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total Revenue (COGS)", format_rupiah(kpi_B_cogs[0]))
            with col_A:
                st.metric("Total COGS", format_rupiah(kpi_A_cogs[1]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_cogs[1], kpi_B_cogs[1], format_rupiah, False
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total COGS", format_rupiah(kpi_B_cogs[1]))
            with col_A:
                st.metric("Total Profit", format_rupiah(kpi_A_cogs[2]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_cogs[2], kpi_B_cogs[2], format_rupiah, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total Profit", format_rupiah(kpi_B_cogs[2]))
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
            col_A, col_Delta, col_B = st.columns([0.35, 0.3, 0.35])
            with col_A:
                st.metric("Total Penjualan (SDM)", format_rupiah(kpi_A_waiter[0]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_waiter[0], kpi_B_waiter[0], format_rupiah, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total Penjualan (SDM)", format_rupiah(kpi_B_waiter[0]))
            with col_A:
                st.metric("Total Transaksi (SDM)", format_angka_bulat(kpi_A_waiter[1]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_waiter[1], kpi_B_waiter[1], format_angka_bulat, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Total Transaksi (SDM)", format_angka_bulat(kpi_B_waiter[1]))
            with col_A:
                st.metric("Rata-rata / Waiter", format_rupiah(kpi_A_waiter[2]))
            with col_Delta:
                d_val, d_str, d_col = calculate_delta(
                    kpi_A_waiter[2], kpi_B_waiter[2], format_rupiah, True
                )
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric("Rata-rata / Waiter", format_rupiah(kpi_B_waiter[2]))


def build_tab5_forecast(data_gmv):
    st.header("🔮 Peramalan Tren Penjualan Detail")
    st.info("Tab ini menggunakan model `Prophet` untuk menganalisis data GMV Anda.")
    if data_gmv is None or data_gmv.empty:
        st.warning("Data GMV tidak ditemukan di database untuk melihat peramalan tren.")
        return
    try:
        daily_sales = (
            data_gmv.groupby(data_gmv["Sales Date In"].dt.date)[
                "Total After Bill Discount"
            ]
            .sum()
            .reset_index()
        )
        daily_sales.rename(
            columns={"Sales Date In": "ds", "Total After Bill Discount": "y"},
            inplace=True,
        )
        daily_sales["ds"] = pd.to_datetime(daily_sales["ds"])
        daily_sales = daily_sales.sort_values(by="ds")
    except Exception as e:
        st.error(f"Gagal memproses data GMV untuk peramalan: {e}")
        return
    if len(daily_sales) < 15:
        st.warning(
            f"Data tidak cukup untuk peramalan detail. Dibutuhkan minimal 15 hari data, Anda memiliki {len(daily_sales)} hari."
        )
        st.dataframe(daily_sales)
        return
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
        f"Model akan dilatih pada data Anda (sampai {last_date.strftime('%d-%m-%Y')}) dan meramalkan penjualan untuk **{forecast_days} hari ke depan**."
    )
    try:
        with st.spinner("Melatih model Prophet... Ini mungkin perlu beberapa detik..."):
            model = Prophet(weekly_seasonality=True, daily_seasonality=False)
            model.fit(daily_sales)
            future_df = model.make_future_dataframe(periods=forecast_days)
            forecast_df = model.predict(future_df)
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
        st.write("1. **Trend**: Menunjukkan arah bisnis Anda secara jangka panjang.")
        st.write(
            "2. **Weekly**: Menunjukkan pola mingguan. Hari apa yang paling kuat dan paling lemah?"
        )
        fig2 = plot_components_plotly(model, forecast_df)
        st.plotly_chart(fig2, use_container_width=True)
    except Exception as e:
        st.error(f"Terjadi kesalahan saat melatih model Prophet: {e}")
        st.exception(e)


def build_tab6_target(data_gmv):
    st.header("🎯 Pencapaian Target & Proyeksi (Dinamis)")
    if data_gmv is None or data_gmv.empty:
        st.warning(
            "Data GMV tidak ditemukan di database untuk melihat analisis target."
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
            "Pilih Bulan yang Ingin Dianalisis:", options=month_options.values()
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
            key=f"target_juta_{active_month_name}",
        )
        target_bulanan = target_juta * 1_000_000
        st.metric(label=f"Target Anda Diatur ke:", value=format_rupiah(target_bulanan))
    st.markdown("---")
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
        proyeksi_prophet = 0
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
            proyeksi_sales_sisa = (sisa_weekdays_count * avg_sales_weekday) + (
                sisa_weekends_count * avg_sales_weekend
            )
            proyeksi_akhir_bulan += proyeksi_sales_sisa
            pembagi = sisa_weekdays_count + (sisa_weekends_count * weekend_weight)
            if pembagi > 0:
                rdr_weekday = sales_dibutuhkan / pembagi
                rdr_weekend = rdr_weekday * weekend_weight
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
                if len(prophet_data) >= 7:
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
                    proyeksi_prophet = proyeksi_akhir_bulan
            except Exception as e_prophet:
                st.warning(f"Gagal menjalankan ramalan Prophet: {e_prophet}")
                proyeksi_prophet = proyeksi_akhir_bulan
        else:
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
    st.subheader("🔮 Perbandingan Model Ramalan")
    with st.container(border=True):
        col_p1, col_p2 = st.columns(2)
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
    try:
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
                "Grafik ini menunjukkan penjualan aktual Anda (biru) dilanjutkan dengan target dinamis (merah putus-putus)."
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
            .properties(padding={"top": 20})
        )
        st.altair_chart(line_chart, use_container_width=True)
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
        base = (
            alt.Chart(plot_df_2)
            .encode(
                x=alt.X("Nama Hari:N", title="Hari", sort=day_sort_order),
                y=alt.Y("Jumlah:Q", title="Penjualan (Rp)"),
            )
            .properties(padding={"top": 20})
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
                "Grafik ini membandingkan performa rata-rata Anda per hari (Bars) dengan Target Harian Dinamis (garis merah)."
            )
        else:
            st.write(
                "Grafik ini menampilkan performa rata-rata Anda per hari (Bars) untuk bulan yang telah selesai."
            )
        st.altair_chart((bar_chart + line_chart_target), use_container_width=True)
    except Exception as e:
        st.error(f"Gagal membuat grafik tren harian: {e}")
        st.exception(e)


def build_tab7_pelanggan(data_ulasan):
    st.header("❤️ Analisis Ulasan & Sentimen Pelanggan")
    if data_ulasan is None or data_ulasan.empty:
        st.warning("Data Ulasan tidak ditemukan di database.")
        return
    st.subheader("📊 KPI Kuantitatif Pelanggan")
    total_ulasan = len(data_ulasan)
    avg_rating = data_ulasan["Rating_Clean"].mean()
    promoters = len(data_ulasan[data_ulasan["Rating_Clean"] == 5])
    detractors = len(data_ulasan[data_ulasan["Rating_Clean"] <= 3])
    nps_score = 0
    if total_ulasan > 0:
        nps_score = (promoters / total_ulasan) * 100 - (detractors / total_ulasan) * 100
    kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
    kpi_col1.metric("💬 Total Ulasan Diterima", f"{total_ulasan} Ulasan")
    kpi_col2.metric("⭐ Rata-rata Rating", f"{avg_rating:.1f} / 5.0")
    kpi_col3.metric("📈 Estimasi NPS", f"{nps_score:.1f}")
    st.markdown("---")
    st.subheader("Grafik Distribusi Rating")
    rating_counts = (
        data_ulasan["Rating_Clean"]
        .value_counts()
        .reset_index()
        .rename(columns={"Rating_Clean": "Rating", "count": "Jumlah"})
    )
    all_ratings = pd.DataFrame({"Rating": [1, 2, 3, 4, 5]})
    rating_counts = pd.merge(
        all_ratings, rating_counts, on="Rating", how="left"
    ).fillna(0)
    chart_dist = create_vertical_bar_chart(
        rating_counts,
        "Rating",
        "Jumlah",
        "Rating Bintang",
        "Jumlah Ulasan",
        x_type="O",
        sort_order=[1, 2, 3, 4, 5],
    )
    chart_dist = chart_dist.configure_axisX(labelAngle=0)
    st.altair_chart(chart_dist, use_container_width=True)
    st.markdown("---")
    st.subheader("💭 Analisis Topik Kualitatif")
    POSITIVE_KEYWORDS = {
        "Rasa Enak": ["enak", "nagih", "pas bumbunya", "oke banget"],
        "Fasilitas/Suasana": [
            "nyaman",
            "estetik",
            "bagus",
            "bersih",
            "VIP",
            "karaoke",
            "lift",
        ],
        "Porsi": ["porsi besar", "kenyang", "lumayan besar"],
        "Pelayanan": ["ramah", "cepat", "baik"],
    }
    NEGATIVE_KEYWORDS = {
        "Harga": ["mahal", "overpriced", "pricey", "kemahalan", "ga sebanding"],
        "Rasa Kurang": ["biasa aja", "kurang", "B aja", "ga extraordinary"],
        "Porsi": ["porsi kecil", "sedikit", "ga rugi"],
        "Pelayanan": ["lama", "lambat", "tidak ramah"],
    }

    def hitung_keyword(df_ulasan, keyword_dict):
        results = {}
        for category, keywords in keyword_dict.items():
            count = 0
            for keyword in keywords:
                count += df_ulasan["Ulasan"].str.contains(keyword, case=False).sum()
            results[category] = count
        return (
            pd.DataFrame.from_dict(results, orient="index", columns=["Jumlah"])
            .reset_index()
            .rename(columns={"index": "Topik"})
            .sort_values(by="Jumlah", ascending=False)
        )

    df_positive_topics = hitung_keyword(data_ulasan, POSITIVE_KEYWORDS)
    df_negative_topics = hitung_keyword(data_ulasan, NEGATIVE_KEYWORDS)
    col_pos, col_neg = st.columns(2)
    with col_pos:
        st.markdown("##### 👍 Topik Positif yang Paling Sering Disebut")
        chart_pos = create_horizontal_bar_chart(
            df_positive_topics, "Jumlah", "Topik", "Jumlah Penyebutan", "Topik Positif"
        )
        st.altair_chart(chart_pos, use_container_width=True)
    with col_neg:
        st.markdown("##### 👎 Topik Negatif (Keluhan) yang Paling Sering Disebut")
        chart_neg = create_horizontal_bar_chart(
            df_negative_topics, "Jumlah", "Topik", "Jumlah Penyebutan", "Topik Negatif"
        )
        st.altair_chart(chart_neg, use_container_width=True)
    st.markdown("---")
    st.subheader("🔍 Telusuri Ulasan Pelanggan")
    filter_rating = st.selectbox(
        "Tampilkan ulasan untuk rating:",
        options=["Semua", 5, 4, 3, 2, 1],
        format_func=lambda x: f"{x} Bintang" if isinstance(x, int) else "Semua Ulasan",
    )
    if filter_rating == "Semua":
        filtered_reviews = data_ulasan
    else:
        filtered_reviews = data_ulasan[data_ulasan["Rating_Clean"] == filter_rating]
    st.write(f"Menampilkan {len(filtered_reviews)} dari {total_ulasan} ulasan:")
    with st.container(height=400):
        for _, row in filtered_reviews.iterrows():
            with st.expander(f"**{row['Nama']}** - {row['Rating']}"):
                st.write(row["Ulasan"])


def build_footer():
    st.markdown(
        """<div id="custom-footer"><div class="copyright">© Copyright Roni Hidayat Data Driven Speacialist Food and Beverage - 2025</div><div class="links"><a href="#">Hubungi Kami</a><a href="#">Kebijakan Privasi</a><a href="#">Tentang Kami</a><a href="#">Kerjasama</a></div></div>""",
        unsafe_allow_html=True,
    )


# #################################################################
# --- BAGIAN 4: EKSEKUSI APLIKASI UTAMA ---
# #################################################################


def main():
    """Fungsi utama untuk menjalankan aplikasi Streamlit."""
    st.title("📈 Data Driven Specialyst FnB Analyst")
    build_sidebar()

    # 1. Buat koneksi database
    conn = get_db_connection()
    if conn is None:
        return  # Hentikan aplikasi jika koneksi gagal

    # 2. Dapatkan rentang tanggal master
    master_min_date, master_max_date = get_full_date_range(conn)

    # 3. Muat SEMUA data sekali (untuk Tab 4, 5, 6)
    #    Kita gunakan koneksi (conn) sebagai argumen pertama untuk caching
    data_gmv = load_data_from_db(
        conn, "gmv", "Sales Date In", master_min_date, master_max_date
    )
    data_cogs = load_data_from_db(
        conn, "cogs", "Sales Date", master_min_date, master_max_date
    )
    data_waiter = load_data_from_db(
        conn, "waiter", "Order Time", master_min_date, master_max_date
    )
    data_ulasan = load_data_from_db(
        conn, "ulasan", "Rating_Clean"
    )  # Ulasan tdk punya tgl

    # 4. Gambar filter global dan dapatkan tanggal yang difilter
    start_date, end_date = build_global_filters(master_min_date, master_max_date)

    # 5. Muat data yang DIFILTER (untuk Tab 1, 2, 3)
    #    Ini akan menggunakan cache jika rentang tanggal sama
    filtered_gmv = load_data_from_db(conn, "gmv", "Sales Date In", start_date, end_date)
    filtered_cogs = load_data_from_db(conn, "cogs", "Sales Date", start_date, end_date)
    filtered_waiter = load_data_from_db(
        conn, "waiter", "Order Time", start_date, end_date
    )

    # 6. Buat Tab
    tab_titles = [
        "📊 KPI Penjualan",
        "💰 COGS & Profit",
        "🧑‍🍳 SDM & Waktu",
        "⚖️ A/B Periodik",
        "🔮 Ramalan Tren",
        "🎯 Target",
        "❤️ Ulasan Pelanggan",
    ]
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(tab_titles)

    # 7. Isi setiap Tab
    with tab1:
        build_tab1_sales(filtered_gmv)
    with tab2:
        build_tab2_cogs(filtered_cogs)
    with tab3:
        build_tab3_hr(filtered_waiter)
    with tab4:
        # Tab 4 menggunakan data LENGKAP
        build_tab4_comparison(
            data_gmv, data_cogs, data_waiter, master_min_date, master_max_date
        )
    with tab5:
        build_tab5_forecast(data_gmv)
    with tab6:
        build_tab6_target(data_gmv)
    with tab7:
        build_tab7_pelanggan(data_ulasan)

    build_footer()

    # ... (Script auto-hide notifikasi Anda) ...


if __name__ == "__main__":
    load_css("style.css")
    main()
