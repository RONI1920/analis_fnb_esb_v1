import time
import pandas as pd
import streamlit as st
import altair as alt
import numpy as np

try:
    from prophet.plot import plot_plotly, plot_components_plotly
except ImportError:
    st.warning(
        "Gagal mengimpor 'prophet'. Tab Forecasting tidak akan berfungsi. Install dengan: pip install prophet"
    )
    plot_plotly, plot_components_plotly = None, None


# --- 1. IMPORT FILE LOKAL ANDA ---
try:
    from data_manager import (
        init_db,
        save_dataframe_to_db,
        load_data_gmv,
        load_cogs_data,
        load_data_waiter,
        load_data_ulasan,
        load_data_purchase,
    )
    from analysis import (
        format_rupiah,
        format_angka_bulat,
        format_persen,
        create_horizontal_bar_chart,
        create_vertical_bar_chart,
        calculate_delta,
        calculate_sales_kpi,
        get_payment_analysis,
        get_visit_purpose_analysis,
        get_menu_performance,
        get_operational_kpi,
        analyze_profit,
        get_peak_time_analysis,
        get_waiter_performance,
        run_prophet_forecast,
        calculate_target_kpis,
        get_top_phrases,
        analyze_purchase_data,
        run_market_basket_analysis,
        get_summary_kpi_sales,
        get_summary_kpi_cogs,
        get_summary_kpi_hr,
        get_summary_kpi_reviews,
        get_summary_kpi_cost,
        run_promo_simulator,
    )
except ImportError as e:
    st.error(
        f"Gagal mengimpor file lokal. Pastikan 'data_manager.py' dan 'analysis.py' ada di folder yang sama."
    )
    st.error(f"Detail Error: {e}")
    st.stop()
# --- BATAS IMPORT ---


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
        with open(file_name) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        # Jangan tampilkan error jika file CSS tidak ada, anggap opsional
        pass


# --- CSS KUSTOM UNTUK TAB ---
st.markdown(
    """
<style>
.stTabs [data-baseweb="tab-list"] button {
    font-size: 1.5rem;  /* Ukuran font tab sudah besar */
    padding: 10px 15px;
}
/* CSS untuk Footer */
#custom-footer {
    border-top: 1px solid #ddd;
    padding: 20px;
    margin-top: 50px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    color: #888;
    font-size: 0.9rem;
}
#custom-footer .links a {
    margin-left: 15px;
    color: #888;
    text-decoration: none;
}
</style>
""",
    unsafe_allow_html=True,
)
# -----------------------------------------------------------------


# #################################################################
# --- BAGIAN 3: FUNGSI PEMBANGUN UI (INTERFACE) ---
# #################################################################


def build_sidebar():
    """Menggambar sidebar dan mengembalikan file yang di-upload."""

    def uncheck_db_on_upload():
        """Callback untuk menonaktifkan DB jika file baru di-upload."""
        st.session_state.use_db = False

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

        st.header("Mode Pemuatan Data")
        if "use_db" not in st.session_state:
            st.session_state.use_db = True

        use_db = st.checkbox(
            "Gunakan data terakhir dari database",
            key="use_db",
            help="Centang untuk memuat data terakhir. Hapus centang untuk mengabaikan database. Meng-upload file baru akan otomatis mengabaikan database.",
        )
        st.markdown("---")

        st.header("Upload Data")
        tipe_file_standar = [
            "xlsx",
            "csv",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "text/csv",
        ]

        gmv_file = st.file_uploader(
            "1. Upload Laporan GMV (Operasional)",
            type=tipe_file_standar,
            key="uploader_gmv",
            on_change=uncheck_db_on_upload,
        )
        if gmv_file is not None:
            if st.button("Simpan Data GMV ini ke Database", key="save_gmv"):
                st.session_state.save_gmv_flag = True
        st.info("ℹ️ Laporan GMV asli (header di baris ke-10).")

        cogs_file = st.file_uploader(
            "2. Upload Laporan COGS (Menu COGS Report)",
            type=tipe_file_standar,
            key="uploader_cogs",
            on_change=uncheck_db_on_upload,
        )
        if cogs_file is not None:
            if st.button("Simpan Data COGS ini ke Database", key="save_cogs"):
                st.session_state.save_cogs_flag = True
        st.info("ℹ️ File COGS (header di baris ke-13).")

        waiter_file = st.file_uploader(
            "3. Upload Sales Recapitulation Detail  (Rekapitulasi Detail)",
            type=tipe_file_standar,
            key="uploader_waiter",
            on_change=uncheck_db_on_upload,
        )
        if waiter_file is not None:
            if st.button("Simpan Data Waiter ini ke Database", key="save_waiter"):
                st.session_state.save_waiter_flag = True
        st.info("ℹ️ Laporan Rekapitulasi Detail (header di baris ke-12).")

        ulasan_file = st.file_uploader(
            "4. Upload Laporan Ulasan Pelanggan",
            type=tipe_file_standar,
            key="uploader_ulasan",
            on_change=uncheck_db_on_upload,
        )
        if ulasan_file is not None:
            if st.button("Simpan Data Ulasan ini ke Database", key="save_ulasan"):
                st.session_state.save_ulasan_flag = True
        st.info("ℹ️ File .csv atau .xlsx berisi kolom: Nama, Rating, Ulasan.")

        purchase_file = st.file_uploader(
            "5. Upload Laporan Pembelian (Purchase Recapitulation)",
            type=tipe_file_standar,
            key="uploader_purchase",
            on_change=uncheck_db_on_upload,
        )
        if purchase_file is not None:
            if st.button("Simpan Data Pembelian ini ke Database", key="save_purchase"):
                st.session_state.save_purchase_flag = True
        st.info("ℹ️ Laporan Pembelian (header di baris ke-12).")

    return (
        gmv_file,
        cogs_file,
        waiter_file,
        ulasan_file,
        purchase_file,
        st.session_state.use_db,
    )


def build_global_filters(data_gmv, data_cogs, data_waiter, data_purchase):
    """Menggambar filter global dan mengembalikan data yang sudah difilter."""
    filtered_gmv = data_gmv
    filtered_cogs = data_cogs
    filtered_waiter = data_waiter
    filtered_purchase = data_purchase

    # Kontainer untuk filter global
    filter_container = st.container()

    with filter_container:
        st.markdown("### 🌎 Filter Global")

        # --- Persiapan Tanggal ---
        min_date = pd.to_datetime("2020-01-01")
        max_date = pd.to_datetime("today")
        default_start = pd.to_datetime("today") - pd.DateOffset(months=1)

        # Coba dapatkan rentang tanggal dari data_gmv jika ada
        if (
            data_gmv is not None
            and not data_gmv.empty
            and "Tanggal" in data_gmv.columns
        ):
            try:
                min_date_data = data_gmv["Tanggal"].min()
                max_date_data = data_gmv["Tanggal"].max()

                # Hanya update jika tanggal valid
                if pd.notna(min_date_data) and pd.notna(max_date_data):
                    min_date = min_date_data
                    max_date = max_date_data
                    default_start = max_date - pd.DateOffset(months=1)

                    # Pastikan default_start tidak lebih awal dari min_date
                    if default_start < min_date:
                        default_start = min_date

            except Exception as e:
                # Biarkan default jika ada masalah (misal: data tanggal tidak valid)
                st.warning(
                    f"Tidak dapat membaca rentang tanggal: {e}. Menggunakan default."
                )
                pass

        # --- Tampilkan Filter ---
        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input(
                "Tanggal Mulai",
                value=default_start,
                min_value=min_date,
                max_value=max_date,
                key="global_start_date",
            )
        with col2:
            end_date = st.date_input(
                "Tanggal Selesai",
                value=max_date,
                min_value=min_date,
                max_value=max_date,
                key="global_end_date",
            )

        # Konversi ke datetime (untuk filtering)
        start_date = pd.to_datetime(start_date)
        end_date = pd.to_datetime(end_date)

        if start_date > end_date:
            st.error("Tanggal Mulai tidak boleh lebih akhir dari Tanggal Selesai.")
            st.stop()

        # --- Terapkan Filter ---
        try:
            if data_gmv is not None and "Tanggal" in data_gmv.columns:
                filtered_gmv = data_gmv[
                    (data_gmv["Tanggal"] >= start_date)
                    & (data_gmv["Tanggal"] <= end_date)
                ]

            # Data COGS adalah 'master' (diagregasi), tidak perlu difilter tgl
            # filtered_cogs = data_cogs

            if data_waiter is not None and "Tanggal" in data_waiter.columns:
                filtered_waiter = data_waiter[
                    (data_waiter["Tanggal"] >= start_date)
                    & (data_waiter["Tanggal"] <= end_date)
                ]

            if data_purchase is not None and "Tanggal" in data_purchase.columns:
                filtered_purchase = data_purchase[
                    (data_purchase["Tanggal"] >= start_date)
                    & (data_purchase["Tanggal"] <= end_date)
                ]
        except Exception as e:
            st.error(f"Error saat menerapkan filter tanggal: {e}")

    # Mengembalikan data yang sudah difilter
    # Perhatikan: kita mengembalikan data_cogs asli (bukan filtered_cogs)
    return filtered_gmv, data_cogs, filtered_waiter, filtered_purchase


# #################################################################
# --- BAGIAN 4: FUNGSI PEMBANGUN TAB ---
# #################################################################


def build_summary_tab(data_gmv, data_cogs, data_waiter, data_ulasan):
    """Membangun tab Ringkasan KPI."""
    with st.spinner("Menghitung KPI Ringkasan..."):
        st.header("Ringkasan Kinerja Bisnis")

        col1, col2, col3, col4, col5 = st.columns(5)

        with col1:
            get_summary_kpi_sales(data_gmv)

        with col2:
            # --- PERBAIKAN ---
            # Memanggil fungsi dengan DUA argumen, sesuai definisi di analysis.py
            get_summary_kpi_cogs(data_gmv, data_cogs)

        with col3:
            get_summary_kpi_hr(data_waiter)

        with col4:
            # Data ulasan tidak difilter berdasarkan tanggal
            get_summary_kpi_reviews(data_ulasan)

        with col5:
            get_summary_kpi_cost(data_gmv)  # Asumsi data_gmv punya info diskon/biaya

        st.markdown("---")
        # Tambahkan visualisasi ringkasan (misal: tren GMV vs COGS)
        st.subheader("Tren GMV vs Profit")
        if data_gmv is not None and data_cogs is not None:
            try:
                # Asumsi fungsi analyze_profit mengembalikan chart
                profit_chart = analyze_profit(data_gmv, data_cogs, "Harian")
                if profit_chart:
                    st.altair_chart(profit_chart, use_container_width=True)
            except Exception as e:
                st.warning(f"Gagal membuat chart tren profit: {e}")
        else:
            st.info("Upload data GMV dan COGS untuk melihat tren profit.")


def build_sales_tab(data_gmv):
    """Membangun tab Analisis Sales."""
    st.header("Analisis Mendalam: Penjualan (Sales)")

    if data_gmv is None or data_gmv.empty:
        st.warning("Data GMV tidak tersedia untuk analisis sales.")
        return

    calculate_sales_kpi(data_gmv)
    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Analisis Metode Pembayaran")
        get_payment_analysis(data_gmv)

    with col2:
        st.subheader("Analisis Tujuan Kunjungan")
        get_visit_purpose_analysis(data_gmv)

    st.markdown("---")
    st.subheader("Performa Menu")
    get_menu_performance(data_gmv)

    st.markdown("---")
    st.subheader("Analisis Waktu Ramai (Peak Time)")
    get_peak_time_analysis(data_gmv)


def build_cogs_tab(data_gmv, data_cogs):
    """Membangun tab Analisis COGS & Profitabilitas."""
    st.header("Analisis Mendalam: COGS & Profitabilitas")

    if data_gmv is None or data_cogs is None or data_gmv.empty or data_cogs.empty:
        st.warning("Data GMV dan COGS diperlukan untuk analisis profitabilitas.")
        return

    # Opsi untuk granulasi (Harian, Mingguan, Bulanan)
    granularity = st.radio(
        "Pilih Granularitas Tren Profit",
        ["Harian", "Mingguan", "Bulanan"],
        index=2,  # Default Bulanan
        horizontal=True,
    )

    st.subheader(f"Tren Profit (Granularitas {granularity})")
    try:
        profit_chart = analyze_profit(data_gmv, data_cogs, granularity)
        if profit_chart:
            st.altair_chart(profit_chart, use_container_width=True)
    except Exception as e:
        st.error(f"Gagal membuat chart profit: {e}")

    st.markdown("---")
    st.subheader("Analisis KPI Target")
    calculate_target_kpis(data_gmv, data_cogs)


def build_hr_tab(data_waiter):
    """Membangun tab Analisis Kinerja HR (Waiter)."""
    st.header("Analisis Mendalam: Kinerja Karyawan (Waiter)")

    if data_waiter is None or data_waiter.empty:
        st.warning("Data Rekapitulasi Detail (Waiter) tidak tersedia.")
        return

    st.subheader("Performa Waiter")
    get_waiter_performance(data_waiter)


def build_reviews_tab(data_ulasan):
    """Membangun tab Analisis Ulasan Pelanggan."""
    st.header("Analisis Mendalam: Ulasan Pelanggan")

    if data_ulasan is None or data_ulasan.empty:
        st.warning("Data Ulasan Pelanggan tidak tersedia.")
        return

    st.subheader("Frasa yang Sering Muncul (Top Phrases)")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### Ulasan Positif 👍")
        get_top_phrases(data_ulasan, sentiment="positif")
    with col2:
        st.markdown("#### Ulasan Negatif 👎")
        get_top_phrases(data_ulasan, sentiment="negatif")


def build_purchase_tab(data_purchase):
    """Membangun tab Analisis Pembelian (Purchase) dan Market Basket."""
    st.header("Analisis Mendalam: Pembelian & Keranjang Belanja")

    if data_purchase is None or data_purchase.empty:
        st.warning("Data Pembelian (Purchase Recapitulation) tidak tersedia.")
        return

    st.subheader("Analisis Data Pembelian")
    analyze_purchase_data(data_purchase)

    st.markdown("---")
    st.subheader("Analisis Keranjang Belanja (Market Basket Analysis)")

    # Input untuk parameter Market Basket
    col1, col2 = st.columns(2)
    with col1:
        min_support = st.slider("Minimum Support", 0.01, 0.1, 0.02, 0.005)
    with col2:
        min_threshold = st.slider(
            "Minimum Threshold (Confidence/Lift)", 0.1, 1.0, 0.5, 0.1
        )

    try:
        run_market_basket_analysis(data_purchase, min_support, min_threshold)
    except Exception as e:
        st.error(f"Gagal menjalankan Market Basket Analysis: {e}")


def build_forecasting_tab(data_gmv):
    """Membangun tab Forecasting menggunakan Prophet."""
    st.header("Forecasting Penjualan (Prophet)")

    if plot_plotly is None:
        st.error("Gagal memuat library 'prophet'. Tab ini tidak dapat berfungsi.")
        return

    if data_gmv is None or data_gmv.empty:
        st.warning("Data GMV diperlukan untuk forecasting.")
        return

    st.markdown(
        """
        Forecasting ini menggunakan data historis **lengkap** (bukan data yang difilter) 
        untuk memprediksi tren penjualan di masa depan.
        """
    )

    periods = st.slider("Jumlah hari ke depan untuk diprediksi", 7, 90, 30)

    try:
        with st.spinner(
            "Melakukan forecasting... Ini mungkin perlu waktu beberapa saat."
        ):
            model, forecast = run_prophet_forecast(data_gmv, periods)

        if model is None or forecast is None:
            st.error("Gagal menjalankan forecast. Data tidak cukup atau error.")
            return

        st.subheader(f"Prediksi Penjualan {periods} Hari ke Depan")
        fig_forecast = plot_plotly(model, forecast)
        fig_forecast.update_layout(title="Prediksi Penjualan (Forecast)")
        st.plotly_chart(fig_forecast, use_container_width=True)

        st.subheader("Komponen Tren & Musiman")
        fig_components = plot_components_plotly(model, forecast)
        fig_components.update_layout(title="Komponen Forecast")
        st.plotly_chart(fig_components, use_container_width=True)

    except Exception as e:
        st.error(f"Gagal menjalankan forecasting: {e}")


def build_simulation_tab(data_gmv, data_cogs):
    """Membangun tab Simulasi Promo."""
    st.header("Simulasi & Skenario")

    if data_gmv is None or data_cogs is None or data_gmv.empty or data_cogs.empty:
        st.warning("Data GMV dan COGS diperlukan untuk simulasi.")
        return

    st.subheader("Simulator Dampak Promo")

    col1, col2, col3 = st.columns(3)
    with col1:
        persen_diskon = st.slider("Persentase Diskon (%)", 0, 50, 10)
    with col2:
        target_kenaikan_qty = st.slider("Target Kenaikan Kuantitas (%)", 0, 100, 20)
    with col3:
        asumsi_cogs_tetap = st.checkbox("Asumsikan COGS per unit tetap", value=True)

    try:
        run_promo_simulator(
            data_gmv, data_cogs, persen_diskon, target_kenaikan_qty, asumsi_cogs_tetap
        )
    except Exception as e:
        st.error(f"Gagal menjalankan simulasi: {e}")


def build_footer():
    """Membangun footer aplikasi."""
    st.markdown(
        """
        <div id="custom-footer">
            <div>
                Data Driven F&B Analyst Dashboard © 2025
            </div>
            <div class="links">
                Developer: @ronihidayat
                <a href="https://www.linkedin.com/in/ronihidayat/" target="_blank">LinkedIn</a>
                <a href="https://github.com/ronihidayat" target="_blank">GitHub</a>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# #################################################################
# --- BAGIAN 5: FUNGSI UTAMA (MAIN) ---
# #################################################################


def main():
    """Fungsi utama untuk menjalankan aplikasi Streamlit."""

    # 1. Inisialisasi Database (membuat tabel jika belum ada)
    try:
        init_db()
    except Exception as e:
        st.sidebar.error(f"Gagal koneksi ke DB: {e}")
        st.sidebar.warning("Mode Database mungkin tidak berfungsi.")

    # 2. Gambar Sidebar dan dapatkan file yang di-upload
    (
        gmv_file,
        cogs_file,
        waiter_file,
        ulasan_file,
        purchase_file,
        use_db,
    ) = build_sidebar()

    # 3. Inisialisasi variabel data
    data_gmv, data_cogs, data_waiter, data_ulasan, data_purchase = (
        None,
        None,
        None,
        None,
        None,
    )

    # 4. Proses Pemuatan Data (Data Loading)
    with st.spinner("Memuat data..."):
        if use_db:
            st.sidebar.info("Memuat data dari Database...")
            try:
                data_gmv = load_data_gmv(None, from_db=True)
                data_cogs = load_cogs_data(None, from_db=True)
                data_waiter = load_data_waiter(None, from_db=True)
                data_ulasan = load_data_ulasan(None, from_db=True)
                data_purchase = load_data_purchase(None, from_db=True)
            except Exception as e:
                st.sidebar.error(f"Gagal memuat dari DB: {e}. Upload file manual.")
                st.session_state.use_db = False  # Nonaktifkan DB jika gagal

        # Jika tidak pakai DB, atau jika ada file yg di-upload (use_db=False)
        if not st.session_state.use_db:
            if gmv_file:
                data_gmv = load_data_gmv(gmv_file, from_db=False)
            if cogs_file:
                data_cogs = load_cogs_data(cogs_file, from_db=False)
            if waiter_file:
                data_waiter = load_data_waiter(waiter_file, from_db=False)
            if ulasan_file:
                data_ulasan = load_data_ulasan(ulasan_file, from_db=False)
            if purchase_file:
                data_purchase = load_data_purchase(purchase_file, from_db=False)

    # 5. Proses Penyimpanan Data (jika tombol save diklik)
    if st.session_state.get("save_gmv_flag", False) and data_gmv is not None:
        save_dataframe_to_db(data_gmv, "gmv_data")
        st.session_state.save_gmv_flag = False
        st.session_state.use_db = True  # Otomatis setel ulang untuk pakai DB
        time.sleep(1)
        st.rerun()

    if st.session_state.get("save_cogs_flag", False) and data_cogs is not None:
        save_dataframe_to_db(data_cogs, "cogs_data")
        st.session_state.save_cogs_flag = False
        st.session_state.use_db = True
        time.sleep(1)
        st.rerun()

    if st.session_state.get("save_waiter_flag", False) and data_waiter is not None:
        save_dataframe_to_db(data_waiter, "waiter_data")
        st.session_state.save_waiter_flag = False
        st.session_state.use_db = True
        time.sleep(1)
        st.rerun()

    if st.session_state.get("save_ulasan_flag", False) and data_ulasan is not None:
        save_dataframe_to_db(data_ulasan, "ulasan_data")
        st.session_state.save_ulasan_flag = False
        st.session_state.use_db = True
        time.sleep(1)
        st.rerun()

    if st.session_state.get("save_purchase_flag", False) and data_purchase is not None:
        save_dataframe_to_db(data_purchase, "purchase_data")
        st.session_state.save_purchase_flag = False
        st.session_state.use_db = True
        time.sleep(1)
        st.rerun()

    # 6. Peringatan jika tidak ada data sama sekali
    if (
        data_gmv is None
        and data_cogs is None
        and data_waiter is None
        and data_ulasan is None
        and data_purchase is None
    ):
        st.warning(
            "👋 Selamat Datang! Silakan upload file data di sidebar kiri "
            "atau centang 'Gunakan data terakhir dari database' untuk memulai analisis."
        )
        st.stop()

    # 7. Gambar Filter Global dan dapatkan data yang sudah difilter
    (filtered_gmv, filtered_cogs, filtered_waiter, filtered_purchase) = (
        build_global_filters(data_gmv, data_cogs, data_waiter, data_purchase)
    )

    # --- 8. BAGIAN DEBUGGER (BARU) ---
    # Gunakan ini untuk memeriksa apakah data Anda dimuat dengan benar
    # dan apakah filter Anda berfungsi.

    with st.sidebar.expander("ℹ️ Data Debugger (Cek Data Asli)"):
        st.markdown("### Data Asli (dari data_manager.py)")
        st.write("Data GMV:")
        if data_gmv is not None:
            st.dataframe(data_gmv.head(5))
        else:
            st.write("None (File Gagal Dimuat)")

        st.write("Data COGS (Master):")
        if data_cogs is not None:
            st.dataframe(data_cogs.head(5))
        else:
            st.write("None (File Gagal Dimuat)")

        st.write("Data Waiter (Agregat):")
        if data_waiter is not None:
            st.dataframe(data_waiter.head(5))
        else:
            st.write("None (File Gagal Dimuat)")

        st.write("Data Ulasan:")
        if data_ulasan is not None:
            st.dataframe(data_ulasan.head(5))
        else:
            st.write("None (File Gagal Dimuat)")

        st.write("Data Pembelian:")
        if data_purchase is not None:
            st.dataframe(data_purchase.head(5))
        else:
            st.write("None (File Gagal Dimuat)")

    with st.sidebar.expander("ℹ️ Data Debugger (Cek Data Terfilter)"):
        st.markdown("### Data Setelah Filter Tanggal Global")
        st.write("Data GMV (Terfilter):")
        if filtered_gmv is not None and not filtered_gmv.empty:
            st.dataframe(filtered_gmv.head(5))
        else:
            st.write("None (Data Asli Kosong atau Terfilter Habis)")

        st.write("Data Waiter (Terfilter):")
        if filtered_waiter is not None and not filtered_waiter.empty:
            st.dataframe(filtered_waiter.head(5))
        else:
            st.write("None (Data Asli Kosong atau Terfilter Habis)")

        st.write("Data Pembelian (Terfilter):")
        if filtered_purchase is not None and not filtered_purchase.empty:
            st.dataframe(filtered_purchase.head(5))
        else:
            st.write("None (Data Asli Kosong atau Terfilter Habis)")

    # --- 9. Buat Tampilan Tab Utama (sebelumnya langkah 8) ---
    (
        tab_summary,
        tab_sales,
        tab_cogs,
        tab_hr,
        tab_reviews,
        tab_purchase,
        tab_forecast,
        tab_sim,
    ) = st.tabs(
        [
            "📊 Ringkasan",
            "💰 Analisis Sales",
            "📈 Analisis COGS & Profit",
            "👥 Analisis HR",
            "⭐ Analisis Ulasan",
            "🛒 Analisis Pembelian",
            "🔮 Forecasting",
            "💡 Simulasi",
        ]
    )

    with tab_summary:
        build_summary_tab(filtered_gmv, filtered_cogs, filtered_waiter, data_ulasan)

    with tab_sales:
        build_sales_tab(filtered_gmv)

    with tab_cogs:
        build_cogs_tab(filtered_gmv, filtered_cogs)

    with tab_hr:
        build_hr_tab(filtered_waiter)

    with tab_reviews:
        # Ulasan biasanya tidak difilter berdasarkan tanggal GMV
        build_reviews_tab(data_ulasan)

    with tab_purchase:
        build_purchase_tab(filtered_purchase)

    with tab_forecast:
        # Forecasting paling baik menggunakan SEMUA data historis (data_gmv, bukan filtered_gmv)
        build_forecasting_tab(data_gmv)

    with tab_sim:
        build_simulation_tab(filtered_gmv, filtered_cogs)

    # 9. Gambar Footer
    build_footer()


# #################################################################
# --- BAGIAN 6: ENTRY POINT APLIKASI ---
# #################################################################

if __name__ == "__main__":
    main()
