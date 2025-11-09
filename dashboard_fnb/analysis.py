import pandas as pd
import streamlit as st
import altair as alt
import numpy as np

try:
    from prophet import Prophet
except ImportError:
    st.warning(
        "Gagal mengimpor 'prophet'. Tab Forecasting tidak akan berfungsi. Install dengan: pip install prophet"
    )
    Prophet = None
try:
    from mlxtend.frequent_patterns import apriori, association_rules
except ImportError:
    st.warning(
        "Gagal mengimpor 'mlxtend'. Tab Pembelian (Market Basket) tidak akan berfungsi. Install dengan: pip install mlxtend"
    )
    apriori, association_rules = None, None

import re  # Diperlukan for get_top_phrases

# --- Fungsi Utility (Format) ---


def format_rupiah(angka):
    if pd.isna(angka):
        return "Rp 0"
    return f"Rp {angka:,.0f}".replace(",", ".")


def format_angka_bulat(angka):
    if pd.isna(angka):
        return "0"
    return f"{angka:,.0f}".replace(",", ".")


def format_persen(angka):
    if pd.isna(angka):
        return "0,0%"
    return f"{angka:,.1f}%".replace(".", ",")


def create_horizontal_bar_chart(df, x, y, tooltip, title):
    """Membuat chart bar horizontal Altair."""
    if df is None or df.empty:
        st.info(f"Data tidak cukup untuk chart '{title}'.")
        return None
    try:
        chart = (
            alt.Chart(df)
            .mark_bar()
            .encode(
                x=alt.X(x, title="Jumlah"),
                y=alt.Y(y, title="Nama", sort="-x"),
                tooltip=tooltip,
            )
            .properties(title=title)
            .interactive()
        )
        return chart
    except Exception as e:
        st.warning(f"Gagal membuat chart '{title}': {e}")
        return None


def create_vertical_bar_chart(df, x, y, x_title, y_title, tooltip, title):
    """Membuat chart bar vertikal Altair."""
    if df is None or df.empty:
        st.info(f"Data tidak cukup untuk chart '{title}'.")
        return None
    try:
        chart = (
            alt.Chart(df)
            .mark_bar()
            .encode(
                x=alt.X(x, title=x_title), y=alt.Y(y, title=y_title), tooltip=tooltip
            )
            .properties(title=title)
            .interactive()
        )
        return chart
    except Exception as e:
        st.warning(f"Gagal membuat chart '{title}': {e}")
        return None


def calculate_delta(value, comparison_value):
    if comparison_value == 0 or pd.isna(comparison_value) or pd.isna(value):
        return 0
    return ((value - comparison_value) / comparison_value) * 100


# --- Fungsi Ringkasan (Summary) ---


def get_summary_kpi_sales(data_gmv):
    if data_gmv is None or data_gmv.empty:
        st.metric(label="Total Sales (GMV)", value="Rp 0")
        return

    total_sales = data_gmv["Total"].sum()
    st.metric(label="Total Sales (GMV)", value=format_rupiah(total_sales))


def get_summary_kpi_cogs(data_gmv, data_cogs):
    if data_gmv is None or data_cogs is None or data_gmv.empty or data_cogs.empty:
        st.metric(label="Total COGS", value="Rp 0")
        return

    try:
        # data_cogs sudah menjadi master list (Nama Menu, COGS) dari data_manager

        # Gabungkan data GMV (yang punya Qty) dengan master COGS
        df_merged = pd.merge(
            data_gmv[["Nama Menu", "Qty", "Total"]],
            data_cogs,  # data_cogs adalah master_cogs
            on="Nama Menu",
            how="left",
        )

        # Hitung total COGS
        df_merged["COGS"] = df_merged["COGS"].fillna(
            0
        )  # Asumsi 0 jika tidak ada di master
        df_merged["Total COGS"] = df_merged["Qty"] * df_merged["COGS"]
        total_cogs = df_merged["Total COGS"].sum()

        # Hitung Persentase COGS
        total_sales = data_gmv["Total"].sum()
        if total_sales > 0:
            persen_cogs = (total_cogs / total_sales) * 100
            st.metric(
                label="Total COGS",
                value=format_rupiah(total_cogs),
                delta=f"{persen_cogs:.1f}% dari Sales",
                delta_color="inverse",  # Merah berarti 'buruk' (biaya tinggi)
            )
        else:
            st.metric(label="Total COGS", value=format_rupiah(total_cogs))

    except Exception as e:
        st.metric(label="Total COGS", value="Rp 0")
        st.caption(f"Gagal hitung COGS: {e}")


def get_summary_kpi_hr(data_waiter):
    if data_waiter is None or data_waiter.empty:
        st.metric(label="Total Waiter Aktif", value="0")
        return

    total_waiter = data_waiter["Nama Waiter"].nunique()
    st.metric(label="Total Waiter Aktif", value=format_angka_bulat(total_waiter))


def get_summary_kpi_reviews(data_ulasan):
    if data_ulasan is None or data_ulasan.empty:
        st.metric(label="Avg. Rating", value="0.0/5")
        return

    avg_rating = data_ulasan["Rating"].mean()
    st.metric(label="Avg. Rating", value=f"{avg_rating:.1f}/5")


def get_summary_kpi_cost(data_gmv):
    if data_gmv is None or data_gmv.empty or "Diskon" not in data_gmv.columns:
        st.metric(label="Total Diskon", value="Rp 0")
        return

    total_diskon = data_gmv["Diskon"].sum()
    st.metric(label="Total Diskon", value=format_rupiah(total_diskon))


# --- Fungsi Tab Sales ---


def calculate_sales_kpi(data_gmv):
    if data_gmv is None or data_gmv.empty:
        st.info("Data GMV tidak cukup untuk KPI Sales.")
        return

    try:
        # Hitung KPI
        total_sales = data_gmv["Total"].sum()
        total_transactions = data_gmv["Nomor Transaksi"].nunique()
        aov = total_sales / total_transactions if total_transactions > 0 else 0
        total_qty = data_gmv["Qty"].sum()

        # Tampilkan dalam kolom
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Sales (GMV)", format_rupiah(total_sales))
        col2.metric("Total Transaksi", format_angka_bulat(total_transactions))
        col3.metric("Avg. Order Value (AOV)", format_rupiah(aov))
        col4.metric("Total Item Terjual", format_angka_bulat(total_qty))

    except Exception as e:
        st.error(f"Gagal menghitung KPI Sales: {e}")


def get_payment_analysis(data_gmv):
    if "Metode Pembayaran" not in data_gmv.columns:
        st.info("Kolom 'Metode Pembayaran' tidak ditemukan di file GMV.")
        return

    # Kita hanya ingin menghitung total per transaksi, bukan per item
    df_trans = data_gmv.drop_duplicates(subset=["Nomor Transaksi"])
    payment_data = df_trans.groupby("Metode Pembayaran")["Total"].count().reset_index()

    chart = create_horizontal_bar_chart(
        payment_data,
        "Total",  # Ini sekarang adalah 'count'
        "Metode Pembayaran",
        ["Metode Pembayaran", "Total"],
        "Jumlah Transaksi per Metode Pembayaran",
    )
    if chart:
        st.altair_chart(chart, use_container_width=True)


def get_visit_purpose_analysis(data_gmv):
    if "Tujuan Kunjungan" not in data_gmv.columns:
        st.info("Kolom 'Tujuan Kunjungan' tidak ditemukan di file GMV.")
        return

    # Kita hanya ingin menghitung total per transaksi, bukan per item
    df_trans = data_gmv.drop_duplicates(subset=["Nomor Transaksi"])
    visit_data = df_trans.groupby("Tujuan Kunjungan")["Total"].count().reset_index()

    chart = create_horizontal_bar_chart(
        visit_data,
        "Total",  # Ini adalah 'count'
        "Tujuan Kunjungan",
        ["Tujuan Kunjungan", "Total"],
        "Jumlah Transaksi per Tujuan Kunjungan",
    )
    if chart:
        st.altair_chart(chart, use_container_width=True)


def get_menu_performance(data_gmv):
    if data_gmv is None or data_gmv.empty:
        st.info("Data GMV tidak cukup untuk analisis performa menu.")
        return

    try:
        df_agg = (
            data_gmv.groupby("Nama Menu")
            .agg(Total=("Total", "sum"), Qty=("Qty", "sum"))
            .reset_index()
        )

        df_agg = df_agg[(df_agg["Qty"] > 0) & (df_agg["Total"] > 0)]
        df_agg["Harga Rata-rata"] = df_agg["Total"] / df_agg["Qty"]

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("#### 10 Menu Terlaris (Berdasarkan Total Sales)")
            df_top_sales = df_agg.nlargest(10, "Total")
            chart_sales = create_horizontal_bar_chart(
                df_top_sales,
                "Total",
                "Nama Menu",
                [
                    "Nama Menu",
                    "Total",
                    "Qty",
                    alt.Tooltip("Harga Rata-rata", format="Rp,.0f"),
                ],
                "Top 10 Menu (Sales)",
            )
            if chart_sales:
                st.altair_chart(chart_sales, use_container_width=True)

        with col2:
            st.markdown("#### 10 Menu Terlaris (Berdasarkan Kuantitas)")
            df_top_qty = df_agg.nlargest(10, "Qty")
            chart_qty = create_horizontal_bar_chart(
                df_top_qty,
                "Qty",
                "Nama Menu",
                [
                    "Nama Menu",
                    "Qty",
                    "Total",
                    alt.Tooltip("Harga Rata-rata", format="Rp,.0f"),
                ],
                "Top 10 Menu (Kuantitas)",
            )
            if chart_qty:
                st.altair_chart(chart_qty, use_container_width=True)

    except Exception as e:
        st.error(f"Gagal menganalisis performa menu: {e}")


def get_operational_kpi(data_gmv):
    st.info("Fungsi get_operational_kpi belum diimplementasi.")


def get_peak_time_analysis(data_gmv):
    if data_gmv is None or data_gmv.empty:
        st.info("Data GMV tidak cukup untuk analisis jam ramai.")
        return

    try:
        df_peak = data_gmv.copy()

        if (df_peak["Tanggal"].dt.hour == 0).all():
            st.info(
                "Data GMV tidak memiliki informasi jam (waktu). Analisis jam ramai tidak dapat dilakukan."
            )
            return

        col1, col2 = st.columns(2)

        # Analisis per Jam
        with col1:
            df_peak["Jam"] = df_peak["Tanggal"].dt.hour
            df_agg_hour = (
                df_peak.groupby("Jam").agg(Total=("Total", "sum")).reset_index()
            )

            chart_hour = create_vertical_bar_chart(
                df_agg_hour,
                "Jam:O",  # :O untuk Ordinal (kategorikal)
                "Total",
                "Jam dalam Sehari",
                "Total Sales",
                ["Jam", alt.Tooltip("Total", format="Rp,.0f")],
                "Total Sales per Jam",
            )
            if chart_hour:
                st.altair_chart(chart_hour, use_container_width=True)

        # Analisis per Hari
        with col2:
            df_peak["Hari"] = df_peak["Tanggal"].dt.day_name()
            hari_order = [
                "Monday",
                "Tuesday",
                "Wednesday",
                "Thursday",
                "Friday",
                "Saturday",
                "Sunday",
            ]

            df_agg_day = (
                df_peak.groupby("Hari")
                .agg(Total=("Total", "sum"))
                .reindex(hari_order)
                .reset_index()
            )

            # --- PERBAIKAN: Buat chart ini secara manual ---
            try:
                chart_day = (
                    alt.Chart(df_agg_day)
                    .mark_bar()
                    .encode(
                        x=alt.X("Hari", sort=hari_order, title="Hari dalam Seminggu"),
                        y=alt.Y("Total", title="Total Sales"),
                        tooltip=["Hari", alt.Tooltip("Total", format="Rp,.0f")],
                    )
                    .properties(title="Total Sales per Hari")
                    .interactive()
                )

                st.altair_chart(chart_day, use_container_width=True)

            except Exception as e:
                st.warning(f"Gagal membuat chart Sales per Hari: {e}")
            # --- Akhir Perbaikan ---

    except Exception as e:
        st.error(f"Gagal menganalisis jam ramai: {e}")


# --- Fungsi Tab COGS ---


def analyze_profit(data_gmv, data_cogs, granularity):
    if data_gmv is None or data_cogs is None or data_gmv.empty or data_cogs.empty:
        st.info("Data GMV dan COGS tidak cukup untuk analisis profit.")
        return None

    try:
        # data_cogs sudah menjadi master list (Nama Menu, COGS) dari data_manager

        df_merged = pd.merge(
            data_gmv[["Tanggal", "Nama Menu", "Qty", "Total"]],
            data_cogs,  # data_cogs adalah master_cogs
            on="Nama Menu",
            how="left",
        )

        df_merged["COGS"] = df_merged["COGS"].fillna(0)
        df_merged["Total COGS"] = df_merged["Qty"] * df_merged["COGS"]
        df_merged["Total Profit"] = df_merged["Total"] - df_merged["Total COGS"]

        if granularity == "Harian":
            rule = "D"
            date_format = "%Y-%m-%d"
        elif granularity == "Mingguan":
            rule = "W-MON"
            date_format = "%Y-%m-%d (Minggu ke %U)"
        else:  # Bulanan (default)
            rule = "MS"
            date_format = "%Y-%m (Bulan)"

        df_agg = df_merged.set_index("Tanggal")
        df_resampled = (
            df_agg.resample(rule)[["Total", "Total COGS", "Total Profit"]]
            .sum()
            .reset_index()
        )

        df_long = df_resampled.melt("Tanggal", var_name="Metrik", value_name="Jumlah")

        df_long["Metrik"] = df_long["Metrik"].replace(
            {
                "Total": "Total Sales",
                "Total COGS": "Total COGS",
                "Total Profit": "Total Profit",
            }
        )

        chart = (
            alt.Chart(df_long)
            .mark_line(point=True)
            .encode(
                x=alt.X("Tanggal", title="Tanggal", axis=alt.Axis(format=date_format)),
                y=alt.Y("Jumlah", title="Jumlah (Rp)"),
                color=alt.Color("Metrik", title="Metrik"),
                tooltip=[
                    alt.Tooltip("Tanggal", title="Tanggal", format=date_format),
                    "Metrik",
                    alt.Tooltip("Jumlah", title="Jumlah (Rp)", format=",.0f"),
                ],
            )
            .properties(title=f"Tren Profitabilitas ({granularity})")
            .interactive()
        )

        return chart

    except Exception as e:
        st.error(f"Gagal menganalisis profit: {e}")
        return None


def calculate_target_kpis(data_gmv, data_cogs):
    if data_gmv is None or data_cogs is None or data_gmv.empty or data_cogs.empty:
        st.info("Data GMV dan COGS tidak cukup untuk KPI Profitabilitas.")
        return

    try:
        # data_cogs sudah menjadi master list (Nama Menu, COGS) dari data_manager

        df_merged = pd.merge(
            data_gmv[["Nama Menu", "Qty", "Total"]],
            data_cogs,  # data_cogs adalah master_cogs
            on="Nama Menu",
            how="left",
        )

        df_merged["COGS"] = df_merged["COGS"].fillna(0)
        df_merged["Total COGS"] = df_merged["Qty"] * df_merged["COGS"]

        total_sales = df_merged["Total"].sum()
        total_cogs = df_merged["Total COGS"].sum()
        total_profit = total_sales - total_cogs

        profit_margin = (total_profit / total_sales) * 100 if total_sales > 0 else 0
        cogs_percent = (total_cogs / total_sales) * 100 if total_sales > 0 else 0

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Profit", format_rupiah(total_profit))
        col2.metric("Profit Margin", format_persen(profit_margin))
        col3.metric("Total COGS", format_rupiah(total_cogs))
        col4.metric("COGS (%)", format_persen(cogs_percent))

    except Exception as e:
        st.error(f"Gagal menghitung KPI Target: {e}")


# --- Fungsi Tab HR ---


def get_waiter_performance(data_waiter):
    if data_waiter is None or data_waiter.empty:
        st.info("Data Waiter tidak cukup untuk analisis performa.")
        return

    try:
        # data_waiter sudah diagregasi per hari per waiter dari data_manager
        df_agg = (
            data_waiter.groupby("Nama Waiter").agg(Total=("Total", "sum")).reset_index()
        )

        st.markdown("#### 10 Waiter Performa Terbaik (Berdasarkan Total Sales)")
        df_top_sales = df_agg.nlargest(10, "Total")
        chart_sales = create_horizontal_bar_chart(
            df_top_sales,
            "Total",
            "Nama Waiter",
            ["Nama Waiter", alt.Tooltip("Total", format="Rp,.0f")],
            "Top 10 Waiter (Sales)",
        )
        if chart_sales:
            st.altair_chart(chart_sales, use_container_width=True)

    except Exception as e:
        st.error(f"Gagal menganalisis performa waiter: {e}")


# --- Fungsi Tab Reviews ---


def get_top_phrases(data_ulasan, sentiment):
    if data_ulasan is None or data_ulasan.empty:
        st.info("Data Ulasan tidak cukup untuk analisis frasa.")
        return

    try:
        if sentiment == "positif":
            df_filtered = data_ulasan[data_ulasan["Rating"] > 3]
            st.caption(
                f"Menganalisis {len(df_filtered)} ulasan positif (Rating > 3)..."
            )
        else:
            df_filtered = data_ulasan[data_ulasan["Rating"] < 3]
            st.caption(
                f"Menganalisis {len(df_filtered)} ulasan negatif (Rating < 3)..."
            )

        if df_filtered.empty:
            st.warning(f"Tidak ada ulasan {sentiment} untuk dianalisis.")
            return

        stop_words = set(
            [
                "yg",
                "ya",
                "utk",
                "dgn",
                "aja",
                "ga",
                "gak",
                "gk",
                "enggak",
                "nya",
                "di",
                "ke",
                "dari",
                "dan",
                "atau",
                "tapi",
                "saya",
                "aku",
                "kamu",
                "dia",
                "kita",
                "kami",
                "anda",
                "ini",
                "itu",
                "ada",
                "adalah",
                "sama",
                "saja",
                "juga",
                "buat",
                "untuk",
                "pada",
                "karena",
                "dengan",
                "yang",
                "sudah",
                "belum",
                "bisa",
                "tidak",
                "gak",
                "ga",
                "banget",
                "bangetnya",
                "sekali",
                "lagi",
                "telah",
                "akan",
                "lebih",
                "kurang",
                "saat",
                "ketika",
                "nan",
                "waktu",
                "sih",
                "kok",
                "deh",
                "aja",
                "cuma",
                "cuman",
                "memang",
                "pake",
                "pakai",
                "biar",
                "jadi",
                "rasanya",
                "saya",
                "makan",
                "minum",
                "pesan",
                "disini",
                "di",
                "sini",
                "situ",
                "sana",
            ]
        )

        all_text = " ".join(df_filtered["Ulasan"].str.lower().dropna())
        all_text = re.sub(r"[^\w\s]", "", all_text)

        words = all_text.split()
        word_counts = {}

        for word in words:
            if word not in stop_words and len(word) > 2:
                word_counts[word] = word_counts.get(word, 0) + 1

        if not word_counts:
            st.warning("Tidak ada kata kunci yang signifikan ditemukan.")
            return

        df_counts = pd.DataFrame(list(word_counts.items()), columns=["Kata", "Jumlah"])
        df_counts = df_counts.nlargest(15, "Jumlah").reset_index(drop=True)

        st.dataframe(df_counts)

    except Exception as e:
        st.error(f"Gagal menganalisis frasa: {e}")


# --- Fungsi Tab Purchase ---


def analyze_purchase_data(data_purchase):
    if data_purchase is None or data_purchase.empty:
        st.info("Data Pembelian tidak cukup untuk dianalisis.")
        return

    try:
        # Analisis 1: Top 10 Supplier berdasarkan Total Pembelian
        if (
            "Total" in data_purchase.columns
            and "Nama Supplier" in data_purchase.columns
        ):
            st.markdown("#### Top 10 Supplier (Berdasarkan Total Pembelian)")
            df_supplier = (
                data_purchase.groupby("Nama Supplier")
                .agg(Total=("Total", "sum"))
                .nlargest(10, "Total")
                .reset_index()
            )
            chart_supplier = create_horizontal_bar_chart(
                df_supplier,
                "Total",
                "Nama Supplier",
                ["Nama Supplier", alt.Tooltip("Total", format="Rp,.0f")],
                "Top 10 Supplier",
            )
            if chart_supplier:
                st.altair_chart(chart_supplier, use_container_width=True)
        else:
            st.info(
                "Kolom 'Total' atau 'Nama Supplier' tidak ada untuk analisis supplier."
            )

        # Analisis 2: Top 10 Item berdasarkan Qty Pembelian
        if "Qty" in data_purchase.columns and "Nama" in data_purchase.columns:
            st.markdown("#### Top 10 Item Dibeli (Berdasarkan Kuantitas)")
            df_item = (
                data_purchase.groupby("Nama")
                .agg(Qty=("Qty", "sum"))
                .nlargest(10, "Qty")
                .reset_index()
            )
            chart_item = create_horizontal_bar_chart(
                df_item, "Qty", "Nama", ["Nama", "Qty"], "Top 10 Item Dibeli"
            )
            if chart_item:
                st.altair_chart(chart_item, use_container_width=True)
        else:
            st.info("Kolom 'Qty' atau 'Nama' (item) tidak ada untuk analisis item.")

    except Exception as e:
        st.error(f"Gagal menganalisis data pembelian: {e}")


def run_market_basket_analysis(data_purchase, min_support, min_threshold):
    if apriori is None:
        st.error("Gagal impor 'mlxtend'. Install dengan: pip install mlxtend")
        return

    if data_purchase is None or data_purchase.empty:
        st.info("Data Pembelian tidak cukup untuk Market Basket Analysis.")
        return

    if "Nomor" not in data_purchase.columns or "Nama" not in data_purchase.columns:
        st.info(
            "Market Basket Analysis memerlukan kolom 'Nomor' (PO) dan 'Nama' (Item)."
        )
        return

    try:
        # 1. Buat data transaksi (basket)
        # Kita anggap 'Nomor' adalah nomor PO/transaksi
        basket = (
            data_purchase.groupby(["Nomor", "Nama"])["Qty"]
            .sum()
            .unstack()
            .reset_index()
            .fillna(0)
            .set_index("Nomor")
        )

        # 2. Ubah Qty menjadi 0 (tidak ada) atau 1 (ada)
        def encode_units(x):
            if x <= 0:
                return 0
            if x >= 1:
                return 1

        basket_sets = basket.applymap(encode_units)

        # Hapus kolom yang jarang dibeli (opsional, tapi mempercepat)
        basket_sets = basket_sets.loc[:, basket_sets.sum() > 1]

        if basket_sets.empty:
            st.warning(
                "Tidak ada item yang cukup sering dibeli bersamaan untuk dianalisis."
            )
            return

        # 3. Jalankan Apriori
        frequent_itemsets = apriori(
            basket_sets, min_support=min_support, use_colnames=True
        )

        if frequent_itemsets.empty:
            st.warning(
                f"Tidak ada itemset yang memenuhi minimum support {min_support}."
            )
            return

        # 4. Jalankan Association Rules
        rules = association_rules(
            frequent_itemsets, metric="lift", min_threshold=min_threshold
        )

        if rules.empty:
            st.warning(
                f"Tidak ada aturan asosiasi yang memenuhi minimum threshold {min_threshold}."
            )
            return

        # 5. Tampilkan hasil
        st.markdown("#### Aturan Asosiasi (Item yang Sering Dibeli Bersamaan)")
        st.dataframe(
            rules[
                ["antecedents", "consequents", "support", "confidence", "lift"]
            ].sort_values("lift", ascending=False)
        )

    except Exception as e:
        st.error(f"Gagal menjalankan Market Basket Analysis: {e}")


# --- Fungsi Tab Forecasting ---


def run_prophet_forecast(data_gmv, periods):
    if Prophet is None:
        st.error("Gagal mengimpor 'prophet'. Forecasting tidak dapat dijalankan.")
        return None, None

    if data_gmv is None or data_gmv.empty:
        st.warning("Data GMV tidak cukup untuk forecast.")
        return None, None

    if "Tanggal" not in data_gmv.columns or "Total" not in data_gmv.columns:
        st.warning("Data GMV harus punya kolom 'Tanggal' dan 'Total' untuk forecast.")
        return None, None

    # Prophet perlu 'ds' dan 'y'
    df_prophet = (
        data_gmv.groupby(data_gmv["Tanggal"].dt.date)
        .agg(y=("Total", "sum"))
        .reset_index()
    )
    df_prophet = df_prophet.rename(columns={"Tanggal": "ds"})

    if len(df_prophet) < 2:
        st.warning("Data tidak cukup untuk forecasting (kurang dari 2 data point).")
        return None, None

    try:
        m = Prophet()
        m.fit(df_prophet)
        future = m.make_future_dataframe(periods=periods)
        forecast = m.predict(future)

        return m, forecast
    except Exception as e:
        st.error(f"Error saat menjalankan Prophet: {e}")
        return None, None


# --- Fungsi Tab Simulasi ---


def run_promo_simulator(
    data_gmv, data_cogs, persen_diskon, target_kenaikan_qty, asumsi_cogs_tetap
):
    if data_gmv is None or data_cogs is None or data_gmv.empty or data_cogs.empty:
        st.info("Data GMV dan COGS tidak cukup untuk simulasi.")
        return

    try:
        # --- 1. Hitung Skenario Saat Ini (Current) ---
        master_cogs = data_cogs  # Sudah jadi master list
        df_merged = pd.merge(
            data_gmv[["Nama Menu", "Qty", "Total"]],
            master_cogs,
            on="Nama Menu",
            how="left",
        )
        df_merged["COGS"] = df_merged["COGS"].fillna(0)
        df_merged["Total COGS"] = df_merged["Qty"] * df_merged["COGS"]

        # KPI Saat Ini
        total_sales_current = df_merged["Total"].sum()
        total_cogs_current = df_merged["Total COGS"].sum()
        total_profit_current = total_sales_current - total_cogs_current
        total_qty_current = df_merged["Qty"].sum()

        if total_qty_current == 0 or total_sales_current == 0:
            st.warning(
                "Data saat ini tidak memiliki penjualan (Qty=0 atau Sales=0). Simulasi dibatalkan."
            )
            return

        avg_price_current = total_sales_current / total_qty_current
        avg_cogs_current = total_cogs_current / total_qty_current
        profit_margin_current = (total_profit_current / total_sales_current) * 100

        # --- 2. Hitung Skenario Simulasi (Simulated) ---

        new_qty = total_qty_current * (1 + target_kenaikan_qty / 100)
        new_price_unit = avg_price_current * (1 - persen_diskon / 100)

        new_cogs_unit = avg_cogs_current  # Asumsi COGS per unit tetap

        new_sales = new_qty * new_price_unit
        new_cogs = new_qty * new_cogs_unit
        new_profit = new_sales - new_cogs
        new_profit_margin = (new_profit / new_sales) * 100 if new_sales > 0 else 0

        # Hitung perbedaan
        delta_sales = new_sales - total_sales_current
        delta_profit = new_profit - total_profit_current
        delta_margin = new_profit_margin - profit_margin_current

        # --- 3. Tampilkan Hasil ---
        st.markdown("---")
        st.subheader("Hasil Simulasi")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("#### Skenario Saat Ini")
            st.metric("Total Sales", format_rupiah(total_sales_current))
            st.metric("Total Profit", format_rupiah(total_profit_current))
            st.metric("Profit Margin", format_persen(profit_margin_current))
            st.metric("Total Item Terjual", format_angka_bulat(total_qty_current))

        with col2:
            st.markdown(
                f"#### Simulasi: Diskon {persen_diskon}% & Kenaikan Qty {target_kenaikan_qty}%"
            )
            st.metric(
                "Total Sales Baru",
                format_rupiah(new_sales),
                delta=format_rupiah(delta_sales),
            )
            st.metric(
                "Total Profit Baru",
                format_rupiah(new_profit),
                delta=format_rupiah(delta_profit),
            )
            st.metric(
                "Profit Margin Baru",
                format_persen(new_profit_margin),
                delta=f"{delta_margin:,.1f} pts",
            )
            st.metric("Total Item Terjual Baru", format_angka_bulat(new_qty))

        if delta_profit > 0:
            st.success(
                f"**Kesimpulan:** Skenario ini diprediksi **MENGUNTUNGKAN**, dengan estimasi kenaikan profit sebesar **{format_rupiah(delta_profit)}**."
            )
        else:
            st.error(
                f"**Kesimpulan:** Skenario ini diprediksi **MERUGIKAN**, dengan estimasi kerugian profit sebesar **{format_rupiah(delta_profit)}**."
            )

    except Exception as e:
        st.error(f"Gagal menjalankan simulasi: {e}")
