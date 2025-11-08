### view.py ###
import streamlit as st
import altair as alt
import pandas as pd

# --- BAGIAN 1: FUNGSI FORMATTING (UNTUK TAMPILAN) ---


def format_rupiah(amount):
    """Format angka menjadi string Rupiah DENGAN titik pemisah ribuan."""
    return f"Rp {amount:,.0f}".replace(",", ".")


def format_angka_bulat(number):
    """Format angka kuantitas menjadi bulat tanpa desimal."""
    return f"{number:.0f}"


def format_persen(number):
    """Format angka menjadi string persentase."""
    return f"{number:,.1f}%"


def load_css(file_name):
    """Membaca file CSS dan menerapkannya ke aplikasi."""
    try:
        with open(file_name) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.error(
            f"File CSS '{file_name}' tidak ditemukan. Pastikan file ada di folder yang sama."
        )


# --- BAGIAN 2: FUNGSI GRAFIK (HELPER TAMPILAN) ---


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


# --- BAGIAN 3: PEMBANGUN UI (VIEW) ---


def build_sidebar():
    """Menggambar sidebar. *TANPA* file uploader."""
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
        st.info("ℹ️ Data dimuat langsung dari Database. Filter ada di halaman utama.")


def build_global_filters(master_min_date, master_max_date):
    """Menggambar filter global dan mengembalikan data yang sudah difilter."""

    st.subheader("Filter Analisis Global")
    st.info(f"Filter global saat ini menggunakan data dari database.")

    filter_type = st.radio(
        "Pilih rentang waktu analisis untuk Tab 1, 2, dan 3:",
        ["Semua Periode", "Harian", "Mingguan", "Bulanan", "Tahunan"],
        horizontal=True,
        key="filter_type_global",
    )

    start_date_filter = master_min_date
    end_date_filter = master_max_date

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

        # Atur start dan end date ke awal dan akhir bulan
        start_date_filter = pd.Timestamp(
            year=selected_year, month=selected_month, day=1
        ).date()
        # Cari hari terakhir di bulan itu
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
    return start_date_filter, end_date_filter


# --- FUNGSI PEMBANGUN TAB (HANYA UI) ---
# Fungsi-fungsi ini sekarang menerima data yang sudah diolah sebagai argumen


def build_tab1_sales(kpi, menu_data, ops_data, payment_data, visit_data, filtered_gmv):

    (
        top_selling,
        top_grossing,
        top_sell_cat,
        top_gross_cat,
        bottom_selling,
        bottom_grossing,
        menu_sales_cat_df,
    ) = menu_data

    (avg_time, peak_hours, peak_days_of_week) = ops_data

    if filtered_gmv is not None:
        if not filtered_gmv.empty:
            start_date = filtered_gmv["Sales Date In"].min().strftime("%d-%m-%Y")
            end_date = filtered_gmv["Sales Date In"].max().strftime("%d-%m-%Y")
            st.subheader(f"Periode Analisis: {start_date} s.d. {end_date}")

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
            st.header("🍽️ Analisis Menu & Kategori")

            # (Sisanya kode UI Anda untuk Tab 1...)
            # ... (UI untuk drill-down kategori) ...
            # ... (UI untuk Top/Bottom 10 Menu) ...

            # Contoh pemanggilan UI:
            if "Menu Category" in filtered_gmv.columns and not menu_sales_cat_df.empty:
                with st.expander(
                    "🍰 Analisis Kategori Menu Interaktif (Klik untuk Detail)",
                    expanded=True,
                ):
                    # ... (Semua logika UI drill-down Anda di sini) ...
                    st.write("UI Drill-Down Kategori (disederhanakan untuk contoh)")

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

            # ... (Sisa UI Tab 1) ...

        elif filtered_gmv is not None and filtered_gmv.empty:
            st.warning(
                "Tidak ada data ditemukan di File GMV untuk rentang waktu yang dipilih."
            )
    else:
        st.info("Data GMV tidak dimuat.")


def build_tab2_cogs(profit_df):
    if profit_df is not None:
        if not profit_df.empty:
            st.header("💰 Analisis Profitabilitas Menu (COGS)")

            with st.expander("💰 Ringkasan Profitabilitas", expanded=True):
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

            # ... (Sisa UI Tab 2 Anda di sini) ...
            with st.expander("📝 Rincian Profitabilitas per Menu (Tabel)"):
                # ... (UI Tabel) ...
                pass

            with st.expander(
                "📊 Analisis Performa Profit Menu (Grafik Top & Bottom 10)",
                expanded=True,
            ):
                # ... (UI Grafik) ...
                pass

        elif profit_df is not None and profit_df.empty:
            st.warning(
                "Tidak ada data ditemukan di File COGS untuk rentang waktu yang dipilih."
            )
    else:
        st.info("Data COGS tidak dimuat.")


def build_tab3_hr(time_data, waiter_data, filtered_waiter):
    if filtered_waiter is not None:
        if not filtered_waiter.empty:
            st.header("🧑‍🍳 Analisis Kinerja Waiter & Waktu Kunjungan")

            with st.expander("🕒 Analisis Waktu Kunjungan Pelanggan", expanded=True):
                # ... (UI Tab 3 Anda di sini) ...
                pass

            with st.expander("🏆 Performa Waiter Teratas (Top 10)", expanded=True):
                # ... (UI Tab 3 Anda di sini) ...
                pass

        elif filtered_waiter is not None and filtered_waiter.empty:
            st.warning(
                "Tidak ada data ditemukan di File Rekapitulasi untuk rentang waktu yang dipilih."
            )
    else:
        st.info("Data Waiter tidak dimuat.")


# ... (Salin sisa fungsi build_tabX Anda ke sini, pastikan mereka hanya menerima data) ...


def build_tab4_comparison(
    kpi_A_gmv,
    kpi_B_gmv,
    kpi_A_cogs,
    kpi_B_cogs,
    kpi_A_waiter,
    kpi_B_waiter,
    captions,
    deltas,
):
    """Menggambar Tab 4 (Perbandingan A/B)"""

    st.header("⚖️ Analisis Perbandingan Periodik (A vs B)")

    # UI Filter ada di app.py (controller), di sini kita hanya tampilkan

    filter_col_A, filter_col_Delta, filter_col_B = st.columns([0.35, 0.3, 0.35])

    with filter_col_A:
        st.subheader("Periode A (Saat Ini)")
        st.caption(captions["A"])
    with filter_col_Delta:
        st.subheader("Perubahan")
        st.caption("A vs B")
    with filter_col_B:
        st.subheader("Periode B (Pembanding)")
        st.caption(captions["B"])

    st.markdown("---")

    # Tampilkan Metrik
    if kpi_A_gmv:
        st.markdown("##### 📊 Kinerja Penjualan (dari File 1: GMV)")
        with st.container(border=True):
            col_A, col_Delta, col_B = st.columns([0.35, 0.3, 0.35])

            # Baris 1: Total Pendapatan
            d_val, d_str, d_col = deltas["gmv_revenue"]
            with col_A:
                st.metric(
                    "Total Pendapatan Kotor",
                    format_rupiah(kpi_A_gmv["Total Pendapatan Kotor"]),
                )
            with col_Delta:
                st.metric("Perubahan", d_val, d_str, delta_color=d_col)
            with col_B:
                st.metric(
                    "Total Pendapatan Kotor",
                    format_rupiah(kpi_B_gmv["Total Pendapatan Kotor"]),
                )

            # ... (Salin sisa UI metrik Tab 4 Anda ke sini) ...


def build_tab5_forecast(forecast_data):
    """Menggambar Tab 5 (Peramalan Detail)"""
    st.header("🔮 Peramalan Tren Penjualan Detail")

    if forecast_data is None:
        st.warning(
            "Data GMV tidak cukup (kurang dari 15 hari) atau tidak ada untuk peramalan."
        )
        return

    fig1, fig2, future_data_table_display, forecast_days = forecast_data

    st.subheader(f"📈 Grafik Peramalan {forecast_days} Hari ke Depan")
    st.plotly_chart(fig1, use_container_width=True)

    with st.expander("Lihat Data Tabel Peramalan (Angka Detail)"):
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
    st.plotly_chart(fig2, use_container_width=True)


# ... (Salin sisa fungsi build_tab6 dan build_tab7 Anda ke sini) ...
# ... Pastikan mereka HANYA menggambar UI ...


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
