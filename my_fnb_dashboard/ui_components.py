# ui_components.py
import time
import pandas as pd
import streamlit as st
import altair as alt
import numpy as np
import re
import os
import plotly.express as px
import plotly.graph_objects as go
from prophet.plot import plot_plotly, plot_components_plotly
from prophet import Prophet  # Diperlukan untuk build_tab5_forecast

# Impor dari file modular kita
from utils import format_rupiah, format_persen, format_angka_bulat

# Impor SEMUA fungsi analisis
import analysis


# --- FUNGSI UNTUK MEMUAT CSS EKSTERNAL ---
def load_css(file_name):
    """Membaca file CSS dan menerapkannya ke aplikasi."""
    try:
        if os.path.exists(file_name):
            with open(file_name) as f:
                st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
        else:
            pass  # Jangan tampilkan error jika style.css tidak ada
    except Exception as e:
        st.error(f"Gagal memuat CSS: {e}")


# --- FUNGSI PEMBUAT GRAFIK (PLOTLY) ---


def create_horizontal_bar_chart(data, x_col, y_col, x_title, y_title, sort_order="-x"):
    """
    Membuat grafik batang horizontal Plotly Express yang profesional.
    """
    sort_ascending = False if sort_order == "-x" else True
    data_sorted = data.sort_values(by=x_col, ascending=sort_ascending)

    fig = px.bar(
        data_sorted,
        x=x_col,
        y=y_col,
        orientation="h",
        labels={x_col: x_title, y_col: ""},
        title=y_title,
        color=x_col,
        color_continuous_scale=px.colors.sequential.Blues,
        template="plotly_white",
        text=x_col,
    )
    fig.update_layout(
        xaxis_title=x_title,
        yaxis_title="",
        xaxis_side="top",
        coloraxis_showscale=False,
        title_x=0.01,
        title_font_size=18,
        margin=dict(l=0, r=20, t=60, b=20),
        yaxis=(
            {"categoryorder": "total ascending"}
            if sort_ascending
            else {"categoryorder": "total descending"}
        ),
    )
    fig.update_traces(
        texttemplate="%{x:.2s}",
        textposition="outside",
        hovertemplate=f"<b>%{{y}}</b><br>{x_title}: %{{x:,.0f}}<extra></extra>",
    )
    fig.update_yaxes(showgrid=False)
    fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor="#E5E5E5")
    return fig


def create_vertical_bar_chart(
    data, x_col, y_col, x_title, y_title, x_type="N", sort_order=None
):
    """
    Membuat grafik batang vertikal Plotly Express yang profesional.
    """
    category_orders = {}
    if sort_order:
        category_orders[x_col] = sort_order

    fig = px.bar(
        data,
        x=x_col,
        y=y_col,
        title=f"{y_title} vs {x_title}",
        labels={x_col: x_title, y_col: y_title},
        color=x_col,
        template="plotly_white",
        category_orders=category_orders,
        text=y_col,
    )
    fig.update_layout(
        xaxis_title=x_title,
        yaxis_title=y_title,
        showlegend=False,
        title_x=0.01,
        title_font_size=18,
        margin=dict(l=0, r=0, t=60, b=0),
        yaxis_tickformat=".2s",
    )
    fig.update_traces(
        texttemplate="%{y:.2s}",
        textposition="outside",
        hovertemplate=f"<b>%{{x}}</b><br>{y_title}: %{{y:,.0f}}<extra></extra>",
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor="#E5E5E5")
    return fig


# --- FUNGSI HELPER UI ---


def calculate_delta(value_A, value_B, formatter_func, higher_is_better=True):
    """
    Menghitung delta antara A dan B, mengembalikan string dan warna.
    """
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


# --- KOMPONEN UI UTAMA ---


def build_sidebar():
    """Menggambar sidebar dan mengembalikan file yang di-upload."""

    def on_checkbox_change():
        st.session_state.use_db = st.session_state.use_db_widget_key

    def uncheck_db_on_upload():
        st.session_state.use_db = False
        st.session_state.gmv_saved_status = False
        st.session_state.cogs_saved_status = False
        st.session_state.waiter_saved_status = False
        st.session_state.ulasan_saved_status = False
        st.session_state.purchase_saved_status = False

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
        if "use_db" not in st.session_state:
            st.session_state.use_db = True

        use_db = st.checkbox(
            "Gunakan data terakhir dari database",
            value=st.session_state.use_db,
            key="use_db_widget_key",
            on_change=on_checkbox_change,
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

        # Uploader 1: GMV
        gmv_file = st.file_uploader(
            "1. Upload Laporan GMV (Operasional)",
            type=tipe_file_standar,
            key="uploader_gmv",
            on_change=uncheck_db_on_upload,
        )
        if gmv_file is not None:
            if st.session_state.gmv_saved_status:
                st.success("Data GMV berhasil disimpan!", icon="✅")
            else:
                if st.button("Simpan Data GMV ini ke Database", key="save_gmv"):
                    st.session_state.save_gmv_flag = True
        st.info("ℹ️ Laporan GMV asli (header di baris ke-10).")

        # Uploader 2: COGS
        cogs_file = st.file_uploader(
            "2. Upload Laporan COGS (Menu COGS Report)",
            type=tipe_file_standar,
            key="uploader_cogs",
            on_change=uncheck_db_on_upload,
        )
        if cogs_file is not None:
            if st.session_state.cogs_saved_status:
                st.success("Data COGS berhasil disimpan!", icon="✅")
            else:
                if st.button("Simpan Data COGS ini ke Database", key="save_cogs"):
                    st.session_state.save_cogs_flag = True
        st.info("ℹ️ File COGS (header di baris ke-13).")

        # Uploader 3: Waiter
        waiter_file = st.file_uploader(
            "3. Upload Sales Recapitulation Detail (Rekapitulasi Detail)",
            type=tipe_file_standar,
            key="uploader_waiter",
            on_change=uncheck_db_on_upload,
        )
        if waiter_file is not None:
            if st.session_state.waiter_saved_status:
                st.success("Data Waiter berhasil disimpan!", icon="✅")
            else:
                if st.button("Simpan Data Waiter ini ke Database", key="save_waiter"):
                    st.session_state.save_waiter_flag = True
        st.info("ℹ️ Laporan Rekapitulasi Detail (header di baris ke-12).")

        # Uploader 4: Ulasan
        ulasan_file = st.file_uploader(
            "4. Upload Laporan Ulasan Pelanggan",
            type=tipe_file_standar,
            key="uploader_ulasan",
            on_change=uncheck_db_on_upload,
        )
        if ulasan_file is not None:
            if st.session_state.ulasan_saved_status:
                st.success("Data Ulasan berhasil disimpan!", icon="✅")
            else:
                if st.button("Simpan Data Ulasan ini ke Database", key="save_ulasan"):
                    st.session_state.save_ulasan_flag = True
        st.info("ℹ️ File .csv atau .xlsx berisi kolom: Nama, Rating, Ulasan.")

        # Uploader 5: Purchase
        purchase_file = st.file_uploader(
            "5. Upload Laporan Pembelian (Purchase Recapitulation)",
            type=tipe_file_standar,
            key="uploader_purchase",
            on_change=uncheck_db_on_upload,
        )
        if purchase_file is not None:
            if st.session_state.purchase_saved_status:
                st.success("Data Pembelian berhasil disimpan!", icon="✅")
            else:
                if st.button(
                    "Simpan Data Pembelian ini ke Database", key="save_purchase"
                ):
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

    master_min_date = pd.Timestamp.now().date()
    master_max_date = pd.Timestamp.now().date()
    filter_source_file = None
    selected_branches = []

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

    if filter_source_file:
        st.subheader("Filter Analisis Global")
        st.info(
            f"Filter global saat ini menggunakan rentang tanggal dari file: **{filter_source_file}**"
        )

        if data_gmv is not None and "Branch" in data_gmv.columns:
            all_branches = sorted(data_gmv["Branch"].unique())
            selected_branches = st.multiselect(
                "Pilih Cabang (Branch):",
                options=all_branches,
                default=all_branches,
                key="branch_filter",
            )
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

        filter_type = st.radio(
            "Pilih rentang waktu (Tab 1, 2, 3, 8):",
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
            if data_cogs is not None:
                filtered_cogs = filtered_cogs[
                    filtered_cogs["Sales Date"].dt.date == selected_date
                ]
            if data_waiter is not None:
                filtered_waiter = filtered_waiter[
                    filtered_waiter["Order Time"].dt.date == selected_date
                ]
            if data_purchase is not None:
                filtered_purchase = filtered_purchase[
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
            selected_date = st.date_input(
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

    elif (
        data_gmv is None
        and data_cogs is None
        and data_waiter is None
        and data_purchase is None
    ):
        st.markdown("---")

    return filtered_gmv, filtered_cogs, filtered_waiter, filtered_purchase


# --- FUNGSI BUILDER TAB ---


def build_tab1_sales(filtered_gmv):
    """
    Menggambar semua elemen untuk Tab 1.
    """
    if filtered_gmv is not None:
        if not filtered_gmv.empty:
            start_date = filtered_gmv["Sales Date In"].min().strftime("%d-%m-%Y")
            end_date = filtered_gmv["Sales Date In"].max().strftime("%d-%m-%Y")
            st.subheader(f"Periode Analisis: {start_date} s.d. {end_date}")

            # === 1. HITUNG SEMUA DATA (PANGGIL analysis.py) ===
            kpi = analysis.calculate_sales_kpi(filtered_gmv)
            (
                top_selling,
                top_grossing,
                top_sell_cat,
                top_gross_cat,
                bottom_selling,
                bottom_grossing,
                menu_sales_cat_df,
            ) = analysis.get_menu_performance(filtered_gmv)
            avg_time, peak_hours, peak_days_of_week = analysis.get_operational_kpi(
                filtered_gmv
            )

            # === 2. TAMPILKAN KPI UTAMA ===
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

            # === 3. ANALISIS MENU & KATEGORI ===
            st.header("🍽️ Analisis Menu & Kategori")
            if "Menu Category" in filtered_gmv.columns and not menu_sales_cat_df.empty:
                with st.expander(
                    "🍰 Analisis Kategori Menu Interaktif (Klik untuk Detail)",
                    expanded=True,
                ):
                    # ... (Logika UI Altair Anda untuk drill-down tetap di sini) ...
                    st.subheader("Grafik Kuantiti Terjual (Drill-Down)")
                    st.info(
                        "Klik pada salah satu batang di Grafik Kategori untuk detail item."
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
                    )
                    st.markdown("---")

                    data_untuk_grafik_atas = (
                        data_kategori_sorted
                        if show_all_categories
                        else data_kategori_display_default
                    )
                    title_grafik_atas = (
                        "Total Penjualan per Kategori (Semua)"
                        if show_all_categories
                        else f"Total Penjualan per Kategori (Top {N_TOP} & Lainnya)"
                    )

                    data_menu_item = (
                        menu_sales_cat_df.groupby(["Menu Category", "Menu"])["Qty"]
                        .sum()
                        .reset_index()
                    )
                    selection_kategori = alt.selection_point(fields=["Menu Category"])

                    # ... (Logika Altair Chart Anda) ...
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
                            title="Detail Penjualan per Menu Item",
                            height=chart_height_detail,
                        )
                    )

                    combined_chart = alt.vconcat(
                        chart_kategori, chart_detail, spacing=40
                    ).resolve_scale(y="independent")
                    st.altair_chart(combined_chart, use_container_width=True)

            # === 4. TOP & BOTTOM MENU ===
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

            # === 5. OPERASIONAL & PELANGGAN ===
            st.header("⚙️ Analisis Operasional & Pelanggan")
            with st.expander(
                "⚙️ KPI Operasional (Jam Sibuk, Hari Sibuk, Durasi Makan)", expanded=True
            ):
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
                        peak_hours_formatted,
                        "Jam_Label",
                        "Bill Number",
                        "Jam",
                        "Jumlah Transaksi",
                        x_type="O",
                        sort_order=hour_sort_order,
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

            with st.expander(
                "💳 Analisis Transaksi (Pembayaran & Kunjungan)", expanded=True
            ):
                if "Payment Method" in filtered_gmv.columns:
                    st.subheader("💳 Penjualan per Metode Pembayaran")
                    payment_data = analysis.get_payment_analysis(
                        filtered_gmv
                    )  # Panggil analysis.py
                    chart = create_horizontal_bar_chart(
                        payment_data,
                        "Total_Penjualan",
                        "Cleaned_Payment",
                        "Total Penjualan (Rp)",
                        "Penjualan per Metode Pembayaran",
                    )
                    st.plotly_chart(chart, use_container_width=True)

                st.markdown("---")

                if "Visit Purpose" in filtered_gmv.columns:
                    st.subheader("🏪 Penjualan per Tipe Kunjungan")
                    visit_data = analysis.get_visit_purpose_analysis(
                        filtered_gmv
                    )  # Panggil analysis.py
                    chart = create_horizontal_bar_chart(
                        visit_data,
                        "Total After Bill Discount",
                        "Visit Purpose",
                        "Total Penjualan (Rp)",
                        "Penjualan per Tipe Kunjungan",
                    )
                    st.plotly_chart(chart, use_container_width=True)

            # === 6. BLOK INSIGHT ===
            st.markdown("---")
            st.header("💡 Insight Otomatis (Ringkasan)")
            insights = analysis.generate_gmv_insights(
                kpi, top_selling, bottom_selling, peak_hours, peak_days_of_week
            )
            with st.expander(
                "Klik untuk melihat Temuan Kunci dari Data Penjualan Anda",
                expanded=True,
            ):
                if insights:
                    for insight in insights:
                        st.markdown(f"&bull; {insight}")
                else:
                    st.info("Tidak ada insight otomatis yang dapat dibuat.")

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

            # 1. Panggil fungsi analisis (sudah di-cache)
            profit_df = analysis.analyze_profit(filtered_cogs)

            # 2. Expander Ringkasan Profitabilitas
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

            # 3. Expander Rincian Profitabilitas
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

            # 4. Expander Analisis Performa
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
                    st.plotly_chart(chart, use_container_width=True)

            # 5. Blok Insight
            st.markdown("---")
            st.header("💡 Insight Otomatis (COGS & Profit)")
            insights = analysis.generate_cogs_insights(profit_df)
            with st.expander(
                "Klik untuk melihat Temuan Kunci dari Data Profit Anda", expanded=True
            ):
                if insights:
                    for insight in insights:
                        st.markdown(f"&bull; {insight}")
                else:
                    st.info("Tidak ada insight otomatis yang dapat dibuat.")

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

            # 1. Hitung data analisis (panggil analysis.py)
            time_data = analysis.get_peak_time_analysis(filtered_waiter)
            waiter_data = analysis.get_waiter_performance(filtered_waiter)

            # 2. Expander Waktu Kunjungan
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

            # 3. Expander Performa Waiter
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

            # 4. Blok Insight
            st.markdown("---")
            st.header("💡 Insight Otomatis (SDM & Waktu)")
            insights = analysis.generate_hr_insights(time_data, waiter_data)
            with st.expander(
                "Klik untuk melihat Temuan Kunci dari Data SDM & Waktu", expanded=True
            ):
                if insights:
                    for insight in insights:
                        st.markdown(f"&bull; {insight}")
                else:
                    st.info("Tidak ada insight otomatis yang dapat dibuat.")

        elif filtered_waiter is not None and filtered_waiter.empty:
            st.warning(
                "Tidak ada data ditemukan di File Rekapitulasi untuk rentang waktu yang dipilih."
            )
    else:
        st.info(
            "Silakan upload file Laporan Rekapitulasi Detail (File 3) di sidebar untuk melihat analisis waiter."
        )


def build_tab4_comparison(data_gmv, data_cogs, data_waiter):
    """Menggambar Tab 4 (Perbandingan A/B)."""

    st.header("⚖️ Analisis Perbandingan Periodik (A vs B)")
    st.info(
        "Gunakan tab ini untuk membandingkan kinerja antara dua periode (A vs B). Filter global diabaikan di tab ini."
    )

    # --- CATATAN ---
    # Logika helper ini SEHARUSNYA ada di analysis.py
    # Namun, untuk menjaga struktur kode asli Anda, kita biarkan di sini.
    # Versi ideal: Pindahkan 3 fungsi helper ini ke analysis.py

    @st.cache_data
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

    @st.cache_data
    def get_profit_kpis(df_cogs_sliced):
        profit_df = analysis.analyze_profit(df_cogs_sliced)  # Panggil analysis.py
        if profit_df.empty:
            return 0, 0, 0, 0
        total_revenue = profit_df["Total Revenue (Rp)"].sum()
        total_cogs = profit_df["Total COGS (Rp)"].sum()
        total_profit = profit_df["Total Profit (Rp)"].sum()
        margin = (total_profit / total_revenue) * 100 if total_revenue > 0 else 0
        return total_revenue, total_cogs, total_profit, margin

    @st.cache_data
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

    # ... (Sisa Logika UI untuk Tab 4) ...
    filter_col_A, filter_col_Delta, filter_col_B = st.columns([0.35, 0.3, 0.35])
    # ... (Logika penentuan default_A_date dan default_B_date) ...
    default_A_date = master_max_date
    default_B_date = master_min_date
    try:
        if comparison_type == "Harian":
            default_B_date = default_A_date - pd.to_timedelta(1, unit="d")
        elif comparison_type == "Mingguan":
            default_B_date = default_A_date - pd.to_timedelta(7, unit="d")
        elif comparison_type == "Bulanan":
            default_B_date = (
                default_A_date - pd.DateOffset(months=1)
            ).date()  # Perbaikan .date()
        elif comparison_type == "Tahunan":
            default_B_date = (
                default_A_date - pd.DateOffset(years=1)
            ).date()  # Perbaikan .date()

        if pd.Timestamp(default_B_date) < pd.Timestamp(master_min_date):
            default_B_date = master_min_date
        elif isinstance(default_B_date, pd.Timestamp):
            default_B_date = default_B_date.date()
    except Exception:
        default_B_date = master_min_date

    with filter_col_A:
        st.subheader("Periode A (Saat Ini)")
        date_A = st.date_input(
            "Tanggal Acuan (A)",
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
            "Tanggal Acuan (B)",
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
        analysis.calculate_sales_kpi(gmv_A)
        if data_gmv is not None
        else analysis.calculate_sales_kpi(None)
    )
    kpi_B_gmv = (
        analysis.calculate_sales_kpi(gmv_B)
        if data_gmv is not None
        else analysis.calculate_sales_kpi(None)
    )
    kpi_A_cogs = get_profit_kpis(cogs_A) if data_cogs is not None else (0, 0, 0, 0)
    kpi_B_cogs = get_profit_kpis(cogs_B) if data_cogs is not None else (0, 0, 0, 0)
    kpi_A_waiter = get_waiter_kpis(waiter_A) if data_waiter is not None else (0, 0, 0)
    kpi_B_waiter = get_waiter_kpis(waiter_B) if data_waiter is not None else (0, 0, 0)

    # Tampilkan UI Metrik
    if data_gmv is not None:
        st.markdown("##### 📊 Kinerja Penjualan (dari File 1: GMV)")
        with st.container(border=True):
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

    # Blok Insight
    st.markdown("---")
    st.header("💡 Insight Otomatis (Analisis Perbandingan)")
    insights = analysis.generate_comparison_insights(
        kpi_A_gmv, kpi_B_gmv, kpi_A_cogs, kpi_B_cogs, kpi_A_waiter, kpi_B_waiter
    )
    with st.expander(
        "Klik untuk melihat Temuan Kunci dari Perbandingan A vs B", expanded=True
    ):
        if insights:
            for insight in insights:
                st.markdown(f"&bull; {insight}")
        else:
            st.info("Tidak ada data yang cukup untuk membuat perbandingan insight.")


def build_tab5_forecast(data_gmv):
    """Menggambar Tab 5 (Peramalan Detail) menggunakan Prophet."""

    st.header("🔮 Peramalan Tren Penjualan Detail")
    st.info(
        "Tab ini menggunakan model `Prophet` untuk menganalisis data GMV Anda, mendeteksi pola mingguan, dan meramalkan penjualan di masa depan."
    )

    if data_gmv is None or data_gmv.empty:
        st.warning(
            "Silakan upload file Laporan GMV (File 1) di sidebar untuk melihat peramalan tren."
        )
        return

    # --- 1. Agregasi data (Logika UI/Persiapan) ---
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

    # --- 2. Pengaturan UI ---
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

    # --- 3. Latih Model (Logika Analisis di dalam UI) ---
    # CATATAN: Idealnya, blok try/except ini ada di analysis.py
    try:
        model = Prophet(weekly_seasonality=True, daily_seasonality=False)
        model.fit(daily_sales)
        future_df = model.make_future_dataframe(periods=forecast_days)
        forecast_df = model.predict(future_df)

        # --- 4. Tampilkan Hasil (UI) ---
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
        fig2 = plot_components_plotly(model, forecast_df)
        st.plotly_chart(fig2, use_container_width=True)

        # --- 5. Blok Insight (Panggil analysis.py) ---
        st.markdown("---")
        st.header("💡 Insight Otomatis (Analisis Ramalan)")
        insights = analysis.generate_forecast_insights(forecast_df, last_date)
        with st.expander(
            "Klik untuk melihat Temuan Kunci dari Model Ramalan", expanded=True
        ):
            if insights:
                for insight in insights:
                    st.markdown(f"&bull; {insight}")
            else:
                st.info("Tidak ada insight otomatis yang dapat dibuat.")

    except Exception as e:
        st.error(f"Terjadi kesalahan saat melatih model Prophet: {e}")
        st.exception(e)


def build_tab6_target(data_gmv):
    """Menggambar Tab 6 (Pencapaian Target)."""

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
            key=f"target_juta_{active_month_name}",
        )
        target_bulanan = target_juta * 1_000_000
        st.metric(label=f"Target Anda Diatur ke:", value=format_rupiah(target_bulanan))
    st.markdown("---")

    # --- CATATAN: Logika Analisis Berat di dalam Fungsi UI ---
    # Idealnya, semua di dalam try/except ini dipindahkan ke analysis.py
    # dan hanya mengembalikan kpi_insight_dict

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
        kpi_insight_dict["sisa_hari"] = sisa_hari

        penjualan_saat_ini = data_bulan_aktif["Total After Bill Discount"].sum()
        kpi_insight_dict["penjualan_saat_ini"] = penjualan_saat_ini

        pencapaian_persen = (
            (penjualan_saat_ini / target_bulanan) if target_bulanan > 0 else 0
        )
        kpi_insight_dict["pencapaian_persen"] = pencapaian_persen

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

        kpi_insight_dict["avg_sales_weekday"] = avg_sales_weekday
        kpi_insight_dict["avg_sales_weekend"] = avg_sales_weekend

        weekend_weight = avg_sales_weekend / avg_sales_weekday
        proyeksi_akhir_bulan = penjualan_saat_ini
        rdr_weekday = 0
        rdr_weekend = 0
        proyeksi_prophet = 0

        if sisa_hari > 0:
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

            kpi_insight_dict["rdr_weekday"] = rdr_weekday
            kpi_insight_dict["rdr_weekend"] = rdr_weekend

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

                # Panggil analysis.py
                ramalan_sisa_hari = analysis.get_prophet_projection(
                    prophet_data, sisa_hari
                )

                if ramalan_sisa_hari is not None:
                    proyeksi_prophet = penjualan_saat_ini + ramalan_sisa_hari
                else:
                    proyeksi_prophet = proyeksi_akhir_bulan
            except Exception as e_prophet:
                st.warning(f"Gagal memproses data untuk Prophet: {e_prophet}")
                proyeksi_prophet = proyeksi_akhir_bulan
        else:
            proyeksi_akhir_bulan = penjualan_saat_ini
            proyeksi_prophet = penjualan_saat_ini

        proyeksi_vs_target_persen = (
            (proyeksi_akhir_bulan / target_bulanan) if target_bulanan > 0 else 0
        )
        kekurangan_proyeksi = target_bulanan - proyeksi_akhir_bulan

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

    # --- Mulai UI untuk Tab 6 ---
    st.subheader("📈 Gambaran Besar (Pencapaian)")
    col1, col2, col3 = st.columns(3)
    col1.metric(
        label=f"Pencapaian per {max_date_in_selected_month.strftime('%d-%m-%Y')}",
        value=f"{pencapaian_persen*100:,.1f}%",
        help=f"{format_rupiah(penjualan_saat_ini)} dari {format_rupiah(target_bulanan)}",
    )
    if sisa_hari > 0:
        col2.metric(
            "Penjualan Dibutuhkan",
            format_rupiah(sales_dibutuhkan),
            help=f"Sisa {sisa_hari} hari.",
        )
        col3.metric(
            "Proyeksi Kekurangan (vs Cerdas)", format_rupiah(kekurangan_proyeksi)
        )
    else:
        col2.metric(
            "Hasil Akhir Bulan (Selesai)",
            format_rupiah(penjualan_saat_ini),
            delta=f"{pencapaian_persen*100:,.1f}% dari Target",
            delta_color=status_color,
        )
        col3.metric(
            "Selisih Target", format_rupiah(penjualan_saat_ini - target_bulanan)
        )

    st.subheader("🔮 Perbandingan Model Ramalan")
    with st.container(border=True):
        col_p1, col_p2 = st.columns(2)
        proyeksi_prophet_persen = (
            (proyeksi_prophet / target_bulanan) if target_bulanan > 0 else 0
        )
        if proyeksi_prophet_persen > 1.05:
            prophet_color = "normal"
        elif proyeksi_prophet_persen >= 0.98:
            prophet_color = "off"
        else:
            prophet_color = "inverse"

        col_p1.metric(
            "Proyeksi Akhir Bulan (Cerdas)",
            format_rupiah(proyeksi_akhir_bulan),
            delta=f"{proyeksi_vs_target_persen * 100:,.1f}% dari Target",
            delta_color=status_color,
        )
        col_p2.metric(
            "Proyeksi Akhir Bulan (Model Prophet)",
            format_rupiah(proyeksi_prophet),
            delta=f"{proyeksi_prophet_persen * 100:,.1f}% dari Target",
            delta_color=prophet_color,
        )

    st.markdown("---")
    st.subheader("💡 Rencana Aksi & Diagnostik")
    col_a, col_b = st.columns(2)
    if sisa_hari > 0:
        col_a.metric(
            label="Target Harian Sisa (Weekday)",
            value=format_rupiah(rdr_weekday),
            delta=f"vs Rata-rata: {format_rupiah(avg_sales_weekday)}",
            delta_color=(
                "inverse"
                if rdr_weekday > avg_sales_weekday
                else ("normal" if rdr_weekday > 0 else "off")
            ),
        )
        col_b.metric(
            label="Target Harian Sisa (Weekend)",
            value=format_rupiah(rdr_weekend),
            delta=f"vs Rata-rata: {format_rupiah(avg_sales_weekend)}",
            delta_color=(
                "inverse"
                if rdr_weekend > avg_sales_weekend
                else ("normal" if rdr_weekend > 0 else "off")
            ),
        )
    else:
        col_a.metric("Rata-rata Weekday (Final)", format_rupiah(avg_sales_weekday))
        col_b.metric("Rata-rata Weekend (Final)", format_rupiah(avg_sales_weekend))

    st.markdown("---")
    st.subheader(f"📈 Tren Penjualan Harian - {active_month_name}")
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
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig_daily_trend.update_traces(
        hovertemplate="<b>%{x|%d %B}</b><br>Penjualan: %{y:,.0f}<extra></extra>"
    )
    st.plotly_chart(fig_daily_trend, use_container_width=True)

    # Blok Insight
    st.markdown("---")
    st.header("💡 Insight Otomatis (Analisis Target)")
    insights = analysis.generate_target_insights(kpi_insight_dict)
    with st.expander(
        "Klik untuk melihat Rangkuman & Rencana Aksi Target Anda", expanded=True
    ):
        if insights:
            for insight_item in insights:
                if insight_item["type"] == "success":
                    st.success(insight_item["text"], icon="✅")
                elif insight_item["type"] == "warning":
                    st.warning(insight_item["text"], icon="⚠️")
                elif insight_item["type"] == "error":
                    st.error(insight_item["text"], icon="🚨")
                else:
                    st.info(insight_item["text"], icon="💡")
        else:
            st.info("Tidak ada insight otomatis yang dapat dibuat.")


def build_tab7_ulasan(data_ulasan):
    """Menggambar Tab 7 (Analisis Ulasan Pelanggan)."""
    st.header("❤️ Analisis Sentimen & Ulasan Pelanggan")

    if data_ulasan is None or data_ulasan.empty:
        st.warning("Silakan upload file Laporan Ulasan (File 4) di sidebar.")
        return

    # --- 1. Pengaturan UI Kata Kunci ---
    with st.expander("Pengaturan Kustomisasi Topik Ulasan (Opsional)"):
        col_k1, col_k2 = st.columns(2)
        keyword_makanan_str = col_k1.text_input(
            "Topik Makanan",
            "enak,lezat,porsi,hambar,asin,dingin,basi,mantap,nikmat,segar",
        )
        keyword_pelayanan_str = col_k1.text_input(
            "Topik Pelayanan", "ramah,cepat,lama,jutek,sopan,membantu,lambat,pelayanan"
        )
        keyword_suasana_str = col_k2.text_input(
            "Topik Suasana",
            "nyaman,bersih,kotor,berisik,adem,dingin,panas,cozy,tempat,suasana",
        )
        keyword_harga_str = col_k2.text_input(
            "Topik Harga", "murah,mahal,worth it,promo,diskon,terjangkau,harga"
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

    # --- 2. Logika Analisis (Idealnya di analysis.py) ---
    @st.cache_data
    def find_topics(ulasan_text, keyword_dict):
        ulasan_text_lower = str(ulasan_text).lower()
        topics_found = []
        for topic, keys in keyword_dict.items():
            for key in keys:
                if re.search(r"\b" + re.escape(key) + r"\b", ulasan_text_lower):
                    topics_found.append(topic)
                    break
        if not topics_found:
            return "Lainnya"
        return ", ".join(topics_found)

    def categorize_nps(rating):
        if rating >= 5:
            return "Promoter"
        elif rating == 4:
            return "Passive"
        else:
            return "Detractor"

    df = data_ulasan.copy()
    df.dropna(subset=["Ulasan", "Rating_Clean"], inplace=True)
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

    # --- 3. Tampilkan Metrik KPI (UI) ---
    st.subheader("📊 KPI Sentimen Pelanggan")
    kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
    kpi_col1.metric("Total Ulasan", f"{total_ulasan} ulasan")
    kpi_col2.metric("Rata-rata Rating", f"{avg_rating:.1f} / 5 ⭐")
    kpi_col3.metric("Net Promoter Score (NPS)", f"{nps_score:.1f}")
    st.markdown("---")

    # --- 4. Tampilkan Grafik Analisis (UI) ---
    st.subheader("🗣️ Analisis Topik Ulasan")
    df_positive = df[df["Rating_Clean"] >= 4]
    df_negative = df[df["Rating_Clean"] <= 3]
    positive_topics = (
        df_positive["Topik"].str.split(", ").explode().value_counts().reset_index()
    )
    negative_topics = (
        df_negative["Topik"].str.split(", ").explode().value_counts().reset_index()
    )
    positive_topics.columns = ["Topik", "Jumlah"]
    negative_topics.columns = ["Topik", "Jumlah"]
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
            st.info("Tidak ada topik positif.")
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
            st.info("Tidak ada topik negatif.")

    # --- 5. Tampilkan Data Mentah (UI) ---
    with st.expander("Lihat Data Mentah Ulasan"):
        st.dataframe(
            df[["Nama", "Rating_Clean", "Ulasan", "NPS_Category", "Topik"]],
            use_container_width=True,
        )

    # --- 6. Blok Insight (Panggil analysis.py) ---
    st.markdown("---")
    st.header("💡 Insight Otomatis (Analisis Ulasan)")
    insights = analysis.generate_review_insights(
        total_ulasan, avg_rating, nps_score, positive_topics, negative_topics
    )
    with st.expander(
        "Klik untuk melihat Temuan Kunci dari Ulasan Pelanggan", expanded=True
    ):
        if insights:
            for insight in insights:
                st.markdown(f"&bull; {insight}")
        else:
            st.info("Tidak ada insight otomatis yang dapat dibuat.")


def build_tab8_purchase(filtered_purchase):
    """Menggambar Tab 8 (Analisis Laporan Pembelian)."""
    st.header("🛒 Analisis Biaya Pembelian (Purchase)")
    st.info(
        "Tab ini menganalisis Laporan Pembelian (File 5) untuk melacak pengeluaran."
    )

    if filtered_purchase is None:
        st.warning("Silakan upload file Laporan Pembelian (File 5) di sidebar.")
        return
    if filtered_purchase.empty:
        st.warning("Tidak ada data pembelian untuk filter yang dipilih.")
        return

    # --- 1. Analisis Data (Panggil analysis.py) ---
    (
        total_cost,
        cost_by_category,
        cost_by_supplier,
        top_items,
        raw_data_filtered,
    ) = analysis.analyze_purchase_data(filtered_purchase)

    # --- 2. Tampilkan Metrik KPI (UI) ---
    st.subheader("📊 KPI Biaya Pembelian")
    st.metric(
        "Total Biaya Pembelian (Tercatat)",
        format_rupiah(total_cost),
        help="Total dari kolom 'Total' di mana nilainya > 0.",
    )
    st.markdown("---")

    # --- 3. Tampilkan Grafik Analisis (UI) ---
    st.subheader("📈 Analisis Rincian Biaya")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("##### Biaya per Kategori (Top 10)")
        if not cost_by_category.empty:
            chart_cat = create_horizontal_bar_chart(
                cost_by_category.nlargest(10, "Total"),
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
            chart_supp = create_horizontal_bar_chart(
                cost_by_supplier.nlargest(10, "Total"),
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

    # --- 4. Tampilkan Data Mentah (UI) ---
    with st.expander("Lihat Rincian Data Pembelian (Sudah Difilter)"):
        st.dataframe(
            raw_data_filtered.style.format(
                {"Price": format_rupiah, "Total": format_rupiah}
            ),
            use_container_width=True,
        )

    # --- 5. Blok Insight (Panggil analysis.py) ---
    st.markdown("---")
    st.header("💡 Insight Otomatis (Analisis Pembelian)")
    insights = analysis.generate_purchase_insights(
        total_cost, cost_by_category, cost_by_supplier, top_items
    )
    with st.expander(
        "Klik untuk melihat Temuan Kunci dari Biaya Pembelian", expanded=True
    ):
        if insights:
            for insight in insights:
                st.markdown(f"&bull; {insight}")
        else:
            st.info("Tidak ada insight otomatis yang dapat dibuat.")


def build_tab9_rekomendasi(filtered_gmv):
    """
    Menggambar Tab 9 (Rekomendasi Menu / Market Basket Analysis).
    """
    st.header("💡 Rekomendasi Menu Interaktif")
    st.info(
        "Pilih menu di filter 'JIKA Beli' untuk melihat item yang paling berpotensi meningkatkan sales."
    )

    if filtered_gmv is None or filtered_gmv.empty:
        st.warning("Silakan upload file Laporan GMV (File 1) di sidebar.")
        return
    if filtered_gmv["Bill Number"].nunique() < 50:
        st.warning(
            f"Jumlah transaksi terlalu sedikit ({filtered_gmv['Bill Number'].nunique()}) untuk analisis. Minimal 50 transaksi."
        )
        return

    # --- 1. Analisis Data (Panggil analysis.py) ---
    with st.spinner(
        "Menganalisis semua aturan asosiasi (bisa perlu waktu 1-2 menit)..."
    ):
        rules_df = analysis.get_market_basket_rules(filtered_gmv, 0.001)

    if rules_df.empty:
        st.error("Gagal menjalankan analisis. Tidak ada aturan yang ditemukan.")
        return

    # --- 2. Filter Interaktif (UI) ---
    st.subheader("Filter Analisis Rekomendasi")
    all_antecedents = sorted(rules_df["antecedents"].unique())
    filter_antecedents = st.multiselect(
        "JIKA Pelanggan Beli Menu Ini:",
        options=all_antecedents,
        placeholder="Ketik menu... (Kosongkan = Tampilkan Semua Aturan Teratas)",
    )

    # --- 3. Terapkan Filter (Logika UI) ---
    with st.spinner("Menerapkan filter..."):
        filtered_rules = rules_df.copy()
        if filter_antecedents:
            filtered_rules = filtered_rules[
                filtered_rules["antecedents"].isin(filter_antecedents)
            ]

    # --- 4. Tampilkan Hasil (UI) ---
    st.markdown("---")
    st.subheader(f"Hasil Rekomendasi (Total: {len(filtered_rules)} aturan)")

    col_a, col_b = st.columns(2)
    with col_a:
        st.success("**Apa itu Confidence (Kepercayaan)?**")
        st.write(
            "**Confidence 50%** (A -> B) berarti: Dari semua orang yang membeli **Menu A**, **50%** dari mereka **juga membeli Menu B**."
        )
    with col_b:
        st.error("**Apa itu Expected Value (Nilai Harapan)?**")
        st.write(
            "**Expected Value = Confidence x Harga Menu B.** Ini adalah metrik terbaik untuk prioritas. Semakin tinggi, potensi sales semakin besar."
        )

    st.markdown("---")
    if not filtered_rules.empty:
        filtered_rules = filtered_rules.sort_values("expected_value", ascending=False)
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
                [
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
        st.info("Tidak ada rekomendasi yang cocok dengan filter Anda.")

    # --- 5. Blok Insight (Panggil analysis.py) ---
    st.markdown("---")
    st.header("💡 Insight Otomatis (Rekomendasi Menu)")
    insights = analysis.generate_recommendation_insights(filtered_rules)
    with st.expander(
        "Klik untuk melihat Temuan Kunci dari Pola Belanja Pelanggan", expanded=True
    ):
        if insights:
            for insight in insights:
                st.markdown(f"&bull; {insight}")
        else:
            st.info("Tidak ada insight otomatis yang dapat dibuat.")


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
