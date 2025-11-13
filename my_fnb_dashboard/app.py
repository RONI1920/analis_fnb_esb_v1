# app.py
import streamlit as st
import pandas as pd  # Diperlukan untuk helper 'clear_db_cache_and_rerun'
import time

# --- 1. Impor modul, bukan fungsi individu ---
# Impor "Si Pemuat Data" & "Si Desainer UI"
import data_loader
import ui_components

# --- Konfigurasi Halaman (Harus jadi perintah Streamlit pertama) ---
st.set_page_config(
    page_title="Data Driven Analyst Specialyst FnB",
    layout="wide",
    initial_sidebar_state="expanded",
)


def main():
    """Fungsi utama untuk menjalankan aplikasi Streamlit."""

    # --- 2. Inisialisasi Awal ---
    data_loader.init_db()  # Pastikan database & tabel ada
    ui_components.load_css("style.css")  # Muat CSS kustom

    # --- 3. Inisialisasi Session State (Penting untuk alur kerja) ---
    # Inisialisasi 'flag' (Pemicu tombol)
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

    # Inisialisasi 'status' (Pesan sukses)
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

    # Inisialisasi Variabel KONTROL 'use_db'
    if "use_db" not in st.session_state:
        st.session_state.use_db = True

    # --- 4. Bangun Sidebar & Dapatkan Input ---
    # Panggil fungsi dari ui_components.py
    (
        gmv_file,
        cogs_file,
        waiter_file,
        ulasan_file,
        purchase_file,
        use_db,
    ) = ui_components.build_sidebar()

    # --- 5. Muat Data (dari File atau DB) ---
    # Panggil fungsi dari data_loader.py
    data_gmv, file_company, file_period, file_branch = data_loader.load_data_gmv(
        gmv_file, use_db
    )

    with st.spinner("Memuat data pendukung..."):
        data_cogs = data_loader.load_cogs_data(cogs_file, use_db)
        data_waiter = data_loader.load_data_waiter(waiter_file, use_db)
        data_ulasan = data_loader.load_data_ulasan(ulasan_file, use_db)
        data_purchase = data_loader.load_data_purchase(purchase_file, use_db)

    # --- 6. Logika Penyimpanan Data (Smart Append) ---
    def clear_db_cache_and_rerun():
        """Helper untuk membersihkan cache DB dan refresh."""
        # HANYA HAPUS CACHE FUNGSI DB
        data_loader.load_dataframe_from_db.clear()
        st.session_state.use_db = True  # Paksa kembali ke mode DB
        st.rerun()

    if st.session_state.save_gmv_flag:
        if data_gmv is not None:
            data_loader.save_dataframe_smart_append(
                data_gmv, "gmv_data", "Sales Date In"
            )
            st.session_state.gmv_saved_status = True
            st.session_state.save_gmv_flag = False
            clear_db_cache_and_rerun()

    if st.session_state.save_cogs_flag:
        if data_cogs is not None:
            data_loader.save_dataframe_smart_append(
                data_cogs, "cogs_data", "Sales Date"
            )
            st.session_state.cogs_saved_status = True
            st.session_state.save_cogs_flag = False
            clear_db_cache_and_rerun()

    if st.session_state.save_waiter_flag:
        if data_waiter is not None:
            data_loader.save_dataframe_smart_append(
                data_waiter, "waiter_data", "Order Time"
            )
            st.session_state.waiter_saved_status = True
            st.session_state.save_waiter_flag = False
            clear_db_cache_and_rerun()

    if st.session_state.save_ulasan_flag:
        if data_ulasan is not None:
            data_loader.save_dataframe_to_db(data_ulasan, "ulasan_data")
            st.session_state.ulasan_saved_status = True
            st.session_state.save_ulasan_flag = False
            clear_db_cache_and_rerun()

    if st.session_state.save_purchase_flag:
        if data_purchase is not None:
            data_loader.save_dataframe_smart_append(
                data_purchase, "purchase_data", "Purchase Date"
            )
            st.session_state.purchase_saved_status = True
            st.session_state.save_purchase_flag = False
            clear_db_cache_and_rerun()

    # --- 7. Filter Global ---
    # Panggil fungsi dari ui_components.py
    (
        filtered_gmv,
        filtered_cogs,
        filtered_waiter,
        filtered_purchase,
    ) = ui_components.build_global_filters(
        data_gmv, data_cogs, data_waiter, data_purchase
    )

    # --- 8. Tampilkan Header Dinamis ---
    if use_db == False and file_company != "DB_MODE":
        # Mode 1: Upload file baru, gunakan info header file
        st.title(f"Analisis Data: {file_company}")
        st.subheader(f"Cabang: {file_branch} | Periode Data: {file_period}")
    elif use_db == True and filtered_gmv is not None and not filtered_gmv.empty:
        # Mode 2: Pakai DB, ambil info dari data yang difilter
        company_name = filtered_gmv["Company"].iloc[0]
        branch_name_header = filtered_gmv["Branch"].iloc[0]
        min_date_str = filtered_gmv["Sales Date In"].min().strftime("%d-%m-%Y")
        max_date_str = filtered_gmv["Sales Date In"].max().strftime("%d-%m-%Y")
        period_str = (
            min_date_str
            if min_date_str == max_date_str
            else f"{min_date_str} s.d. {max_date_str}"
        )

        st.title(f"Analisis Data: {company_name}")
        st.subheader(f"Cabang: {branch_name_header} | Periode Data: {period_str}")
    else:
        # Mode 3: Fallback
        st.title("Dashboard Analisis Data F&B")
        if data_gmv is not None and (filtered_gmv is None or filtered_gmv.empty):
            st.warning("Tidak ada data GMV untuk filter/periode yang Anda pilih.")

    # --- 9. Navigasi Halaman (Tab Utama) ---
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
    ]
    page = st.selectbox("Pilih Halaman Analisis:", page_options)
    st.divider()

    # --- 10. Render Halaman yang Dipilih ---
    # Panggil fungsi builder dari ui_components.py
    if page == "📊 Penjualan (GMV)":
        ui_components.build_tab1_sales(filtered_gmv)
    elif page == "💰 COGS & Profit":
        ui_components.build_tab2_cogs(filtered_cogs)
    elif page == "🧑‍🍳 SDM & Waktu":
        ui_components.build_tab3_hr(filtered_waiter)
    elif page == "🛒 Pembelian":
        ui_components.build_tab8_purchase(filtered_purchase)
    elif page == "⚖️ A/B Comparison":
        ui_components.build_tab4_comparison(data_gmv, data_cogs, data_waiter)
    elif page == "🎯 Target":
        ui_components.build_tab6_target(data_gmv)
    elif page == "🔮 Forecast (AI)":
        ui_components.build_tab5_forecast(data_gmv)
    elif page == "❤️ Ulasan":
        ui_components.build_tab7_ulasan(data_ulasan)
    elif page == "💡 Rekomendasi":
        ui_components.build_tab9_rekomendasi(filtered_gmv)

    # --- 11. Footer ---
    ui_components.build_footer()

    # Skrip JS untuk fading notifikasi (bisa juga dipindah ke ui_components)
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
                    }, 5000); // 5 detik
                }
            });
        }
        setTimeout(hideAllNotices, 1000); 
        </script>
        """,
        unsafe_allow_html=True,
    )


# --- Entry Point Aplikasi ---
if __name__ == "__main__":
    main()
