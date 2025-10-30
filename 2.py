import time
import pandas as pd
import streamlit as st
import openpyxl  # Diperlukan agar pandas bisa membaca file .xlsx
import altair as alt  # Library untuk grafik yang lebih baik

# --- Konfigurasi Halaman Streamlit ---
st.set_page_config(
    page_title="Dashboard KPI Milky Way",
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

# --- Fungsi Bantuan (Helpers) ---


def format_rupiah(amount):
    """Format angka menjadi string Rupiah DENGAN titik pemisah ribuan."""
    # Menggunakan f-string formatting dengan pemisah ribuan (,) dan 0 desimal (.0f)
    # Kemudian mengganti koma (,) dengan titik (.) untuk format Rupiah
    return f"Rp {amount:,.0f}".replace(",", ".")


def format_angka_bulat(number):
    """Format angka kuantitas menjadi bulat tanpa desimal."""
    return f"{number:.0f}"


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
def load_data(uploaded_file):
    """Memuat dan membersihkan data dari file yang di-upload."""
    df = None
    try:
        if uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, header=9, engine="openpyxl")
        elif uploaded_file.name.endswith(".csv"):
            # Coba dengan encoding latin1 jika utf-8 gagal
            df = pd.read_csv(uploaded_file, header=9, encoding="latin1")
        else:
            st.error("Format file tidak didukung. Harap upload .xlsx atau .csv")
            return None

    except Exception as e:
        st.error(f"Error saat membaca file: {e}")
        st.error(
            "Pastikan file yang di-upload adalah file Excel/CSV yang valid dan header berada di baris ke-10."
        )
        return None

    # Daftar kolom numerik yang diharapkan
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

    # Konversi kolom numerik, tangani error 'coerce' (ubah jadi NaN jika gagal)
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            st.warning(f"Peringatan: Kolom '{col}' tidak ditemukan di file Anda.")

    # Konversi kolom tanggal, tangani error 'coerce'
    df["Sales Date In"] = pd.to_datetime(df["Sales Date In"], errors="coerce")
    df["Sales Date Out"] = pd.to_datetime(df["Sales Date Out"], errors="coerce")

    # Hapus baris di mana data penting (Bill Number, Sales Date In) tidak ada
    df.dropna(subset=["Bill Number", "Sales Date In"], inplace=True)

    return df


# --- Fungsi Analisis KPI ---


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
    # Agregasi data per Bill Number terlebih dahulu
    bill_data = (
        df.groupby("Bill Number")
        .agg(
            Bill_Revenue=("Total After Bill Discount", "sum"),
            Payment_Method=(
                "Payment Method",
                "first",
            ),  # Ambil metode pembayaran pertama
        )
        .reset_index()
    )
    # Bersihkan nama metode pembayaran
    bill_data["Cleaned_Payment"] = bill_data["Payment_Method"].apply(
        clean_payment_method
    )
    # Analisis berdasarkan metode pembayaran yang sudah bersih
    payment_analysis = (
        bill_data.groupby("Cleaned_Payment")["Bill_Revenue"]
        .agg(Total_Penjualan="sum", Jumlah_Transaksi="count")
        .sort_values(by="Total_Penjualan", ascending=False)
    )
    return payment_analysis.reset_index()


def get_visit_purpose_analysis(df):
    """Menganalisis penjualan berdasarkan tujuan kunjungan."""
    # Perlu agregasi per bill number agar tidak double count
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
    # Filter 1: Hapus item 'PACKAGE' atau gratis
    menu_sales = df[~df["Menu"].str.contains("PACKAGE", na=False, case=False)]
    menu_sales = menu_sales[menu_sales["Price (Net)"] > 0]

    # Filter 2: Hapus kategori "Additional" dan "Add-ons"
    menu_sales = menu_sales[
        ~menu_sales["Menu Category"].str.contains("ADDITIONAL", na=False, case=False)
    ]
    menu_sales = menu_sales[
        ~menu_sales["Menu Category"].str.contains("ADD-ONS", na=False, case=False)
    ]

    # Analisis Top 10
    top_selling_items = menu_sales.groupby("Menu")["Qty"].sum().nlargest(10)
    top_grossing_items = (
        menu_sales.groupby("Menu")["Total Nett Sales"].sum().nlargest(10)
    )
    top_selling_categories = (
        menu_sales.groupby("Menu Category")["Qty"]
        .sum()
        .nlargest(10)
        .sort_values(ascending=False)
    )
    top_grossing_categories = (
        menu_sales.groupby("Menu Category")["Total Nett Sales"]
        .sum()
        .nlargest(10)
        .sort_values(ascending=False)
    )

    # Analisis Bottom 10 (Menu Kurang Laku)
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
    # Rata-rata Durasi Makan
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
    # Filter durasi yang aneh (misal > 8 jam atau < 1 menit)
    bill_times_filtered = bill_times[
        (bill_times["Duration_minutes"] > 1) & (bill_times["Duration_minutes"] < 480)
    ]
    avg_dining_time = bill_times_filtered["Duration_minutes"].mean()

    # Jam Sibuk (berdasarkan jam 'Sales Date In')
    df_hourly = df.copy()
    df_hourly["Hour"] = df_hourly["Sales Date In"].dt.hour
    peak_hours = (
        df_hourly.groupby("Hour")["Bill Number"].nunique().sort_values(ascending=False)
    )

    # Hari Sibuk (Senin - Minggu)
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
    # Pastikan tidak ada NaN di 'Sales Date In' sebelum ambil day_name
    df_daily.dropna(subset=["Sales Date In"], inplace=True)
    df_daily["Day Name"] = df_daily["Sales Date In"].dt.day_name().map(day_map)
    peak_days_of_week = (
        df_daily.groupby("Day Name")["Bill Number"].nunique().reset_index()
    )

    return avg_dining_time, peak_hours.reset_index(), peak_days_of_week


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

st.title("📈 Dashboard KPI - Milky Way Lippo Mall Puri")

# 1. Widget File Uploader di Sidebar
with st.sidebar:
    st.image(
        "https://img.freepik.com/premium-vector/cute-milk-tea-bubble-tea-logo_51197-248.jpg",
        width=150,
    )  # Ganti logo
    st.header("Upload Data")
    uploaded_file = st.file_uploader(
        "Upload Laporan GMV (.xlsx atau .csv)",
        type=["xlsx", "csv"],
        label_visibility="collapsed",
    )
    st.info("ℹ️ Pastikan header data Anda berada di baris ke-10.")

# 2. Kondisi: Jika file sudah di-upload
if uploaded_file is not None:
    data = load_data(uploaded_file)

    if data is not None and not data.empty:
        st.success(f"File '{uploaded_file.name}' berhasil di-upload dan diproses.")

        start_date = data["Sales Date In"].min().strftime("%d-%m-%Y")
        end_date = data["Sales Date In"].max().strftime("%d-%m-%Y")
        st.subheader(f"Periode Analisis: {start_date} s.d. {end_date}")

        # --- 1. Ringkasan Kinerja Penjualan ---
        st.header("📊 KPI Kinerja Penjualan")
        kpi = calculate_sales_kpi(data)

        col1, col2, col3 = st.columns(3)
        col1.metric(
            "💰 Total Pendapatan Kotor", format_rupiah(kpi["Total Pendapatan Kotor"])
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
        col5.metric("📦 Total Item Terjual", f"{kpi['Total Item Terjual']:,.0f} Items")
        col6.metric(
            "🛍️ Item per Transaksi (IPB)", f"{kpi['Item per Transaksi (IPB)']:.2f}"
        )

        # --- TAMBAHAN: Expander untuk KPI Rinci ---
        with st.expander("Lihat Rincian Pendapatan (Diskon, Service, Pajak)"):
            exp_col1, exp_col2, exp_col3 = st.columns(3)
            exp_col1.metric("📉 Total Diskon", format_rupiah(kpi["Total Diskon"]))
            exp_col2.metric(
                "🛎️ Total Service Charge", format_rupiah(kpi["Total Service Charge"])
            )
            exp_col3.metric("🧾 Total Pajak", format_rupiah(kpi["Total Pajak"]))
        # --- Akhir Tambahan ---

        st.markdown("---")

        col7, col8 = st.columns(2)
        with col7:
            st.subheader("💳 Penjualan per Metode Pembayaran")
            payment_data = get_payment_analysis(data)
            chart = create_horizontal_bar_chart(
                payment_data,
                "Total_Penjualan",
                "Cleaned_Payment",
                "Total Penjualan (Rp)",
                "Metode Pembayaran",
            )
            st.altair_chart(chart, use_container_width=True)

        with col8:
            st.subheader("🏪 Penjualan per Tipe Kunjungan")
            visit_data = get_visit_purpose_analysis(data)
            chart = create_horizontal_bar_chart(
                visit_data,
                "Total After Bill Discount",
                "Visit Purpose",
                "Total Penjualan (Rp)",
                "Tipe Kunjungan",
            )
            st.altair_chart(chart, use_container_width=True)

        st.markdown("---")

        # --- 2. Kinerja Menu ---
        st.header("🍽️ KPI Kinerja Menu")

        (
            top_selling,
            top_grossing,
            top_sell_cat,
            top_gross_cat,
            bottom_selling,
            bottom_grossing,
        ) = get_menu_performance(data)

        st.info(
            "ℹ️ Menu 'PACKAGE', 'ADDITIONAL', 'ADD-ONS', dan item gratis (Price = 0) telah disaring dari analisis performa menu."
        )

        st.subheader("🏆 Performa Menu Teratas (Top 10)")
        col9, col10 = st.columns(2)
        with col9:
            st.markdown("##### Menu Terlaris (by Kuantitas)")
            st.dataframe(
                top_selling.set_index("Menu").style.format({"Qty": format_angka_bulat})
            )

        with col10:
            st.markdown("##### Menu Pendapatan Tertinggi (by Nett Sales)")
            st.dataframe(
                top_grossing.set_index("Menu").style.format(
                    {"Total Nett Sales": format_rupiah}
                )
            )

        # Grafik untuk Top 10
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

        # Grafik untuk Bottom 10
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

        st.markdown("---")
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
            st.markdown("##### Kategori Pendapatan Tertinggi (by Nett Sales)")
            chart = create_horizontal_bar_chart(
                top_gross_cat,
                "Total Nett Sales",
                "Menu Category",
                "Total Nett Sales (Rp)",
                "Kategori Menu",
            )
            st.altair_chart(chart, use_container_width=True)

        st.markdown("---")

        # --- 3. Kinerja Operasional ---
        st.header("⚙️ KPI Operasional & Pelanggan")

        avg_time, peak_hours, peak_days_of_week = get_operational_kpi(data)

        st.metric("⏱️ Rata-rata Durasi Makan (Dine In)", f"{avg_time:.1f} menit")

        col19, col20 = st.columns(2)
        with col19:
            st.subheader("🕒 Jam Sibuk (Berdasarkan Transaksi)")
            chart = create_vertical_bar_chart(
                peak_hours,  # Tampilkan semua jam
                "Hour",
                "Bill Number",
                "Jam",
                "Jumlah Transaksi",
                x_type="O",  # Perlakukan jam sebagai Kategori Ordinal
            )
            st.altair_chart(chart, use_container_width=True)

        with col20:
            st.subheader("🗓️ Hari Sibuk (Berdasarkan Transaksi)")
            # BARU: Urutkan berdasarkan hari (Senin-Minggu)
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

        if st.checkbox("Tampilkan Data Mentah (Sudah Dibersihkan)"):
            st.dataframe(data)

    elif data is not None and data.empty:
        st.error(
            "File berhasil dibaca, namun tidak ada data yang valid setelah diproses. Periksa kembali isi file Anda."
        )

# 4. Kondisi: Jika belum ada file di-upload
else:
    st.info(
        "Silakan upload file Laporan GMV Anda di sidebar kiri untuk memulai analisis."
    )
