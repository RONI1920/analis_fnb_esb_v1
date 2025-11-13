# analysis.py
import pandas as pd
import numpy as np
import streamlit as st
import re
from prophet import Prophet
from mlxtend.frequent_patterns import apriori, association_rules
from mlxtend.preprocessing import TransactionEncoder

# Impor helper dari utils.py
from utils import format_rupiah, format_persen, format_angka_bulat, clean_payment_method


@st.cache_data(show_spinner=False)
def get_prophet_projection(prophet_data, sisa_hari):
    """
    Melatih model Prophet yang di-cache dan mengembalikan nilai ramalan.
    """
    try:
        if len(prophet_data) < 7:
            st.warning("Data bulan ini < 7 hari, ramalan Prophet tidak akurat.")
            return None
        with st.spinner("Menjalankan model ramalan Prophet (pertama kali)..."):
            model_prophet = Prophet(
                weekly_seasonality=True,
                daily_seasonality=False,
                changepoint_prior_scale=0.1,
            )
            model_prophet.fit(prophet_data)
            future_df_prophet = model_prophet.make_future_dataframe(periods=sisa_hari)
            forecast_df_prophet = model_prophet.predict(future_df_prophet)
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
    # Fungsi clean_payment_method diimpor dari utils.py di atas

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
    """
    menu_sales = df[~df["Menu"].str.contains("PACKAGE", na=False, case=False)]
    menu_sales = menu_sales[menu_sales["Price (Net)"] > 0]

    filter_regex = r"ADD[ -]?ON|ADDITIONAL|Level"
    if "Menu Category" in df.columns:
        menu_sales = menu_sales[
            ~menu_sales["Menu Category"].str.contains(
                filter_regex, na=False, case=False, regex=True
            )
        ]
    NAMA_KOLOM_DETAIL = "Menu Category Detail"
    if NAMA_KOLOM_DETAIL in df.columns:
        menu_sales = menu_sales[
            ~menu_sales[NAMA_KOLOM_DETAIL].str.contains(
                filter_regex, na=False, case=False, regex=True
            )
        ]

    filter_regex_items = r"Ocha|Refill|Mineral Water"
    if "Menu" in menu_sales.columns:
        menu_sales = menu_sales[
            ~menu_sales["Menu"].str.contains(
                filter_regex_items, na=False, case=False, regex=True
            )
        ]

    top_selling_categories = pd.DataFrame(columns=["Menu Category", "Qty"])
    top_grossing_categories = pd.DataFrame(
        columns=["Menu Category", "Total Nett Sales"]
    )
    menu_sales_cat_df = pd.DataFrame()

    if "Menu Category" in df.columns:
        menu_sales_cat_df = menu_sales.copy()
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
    filter_regex_all = r"ADDITIONAL|ADD[ -]?ON|New Add-ons|Level"

    if "Menu Category" in profit_df.columns:
        profit_df = profit_df[
            ~profit_df["Menu Category"].str.contains(
                filter_regex_all, na=False, case=False, regex=True
            )
        ]
    if "Menu" in profit_df.columns:
        profit_df = profit_df[
            ~profit_df["Menu"].str.contains(
                filter_regex_all, na=False, case=False, regex=True
            )
        ]

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


@st.cache_data
def analyze_purchase_data(df):
    """
    Menganalisis data pembelian yang sudah difilter.
    """
    if df is None or df.empty:
        return 0, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df_with_cost = df[df["Total"] > 0].copy()
    total_cost = df_with_cost["Total"].sum()

    cost_by_category = (
        df_with_cost.groupby("Category")["Total"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )
    cost_by_supplier = (
        df_with_cost.groupby("Supplier Name")["Total"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )
    top_items = (
        df_with_cost.groupby("Product Name")["Total"]
        .sum()
        .nlargest(20)
        .sort_values(ascending=False)
        .reset_index()
    )
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


@st.cache_data(show_spinner=False)
def get_market_basket_rules(df, min_support_threshold):
    """
    Menjalankan Market Basket Analysis (Apriori) pada data GMV.
    """
    try:
        filter_regex = "PACKAGE|REFILL OCHA"
        df_no_packages = df[
            ~df["Menu"].str.contains(filter_regex, na=False, case=False, regex=True)
        ]
        menu_prices = df_no_packages.groupby("Menu")["Price (Net)"].mean()

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
                st.warning("Tidak ada transaksi yang cukup untuk dianalisis.")
                return pd.DataFrame()

        with st.spinner("Meng-encode data (TransactionEncoder)..."):
            te = TransactionEncoder()
            te_ary = te.fit(transactions_list).transform(transactions_list)
            df_encoded = pd.DataFrame(te_ary, columns=te.columns_)

        with st.spinner(f"Menjalankan Apriori (support={min_support_threshold})..."):
            frequent_itemsets = apriori(
                df_encoded, min_support=min_support_threshold, use_colnames=True
            )
            if frequent_itemsets.empty:
                st.warning(
                    "Tidak ditemukan pola item yang kuat. Coba turunkan 'Minimal Support'."
                )
                return pd.DataFrame()

        with st.spinner("Membangun aturan asosiasi..."):
            rules = association_rules(
                frequent_itemsets, metric="lift", min_threshold=1.0
            )
            if rules.empty:
                st.info("Tidak ada aturan asosiasi yang ditemukan dengan lift > 1.")
                return pd.DataFrame()

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
                st.info("Tidak ada aturan sisa setelah filter logika bisnis.")
                return pd.DataFrame()

        rules = rules[rules["antecedents"].apply(len) == 1]
        if rules.empty:
            st.info("Tidak ada aturan (single item) yang tersisa setelah filter.")
            return pd.DataFrame()

        with st.spinner("Menghitung Expected Value (Confidence x Harga)..."):
            rules = rules[rules["consequents"].apply(len) == 1]
            rules["consequents_str"] = rules["consequents"].apply(
                lambda x: next(iter(x))
            )
            rules = rules.merge(
                menu_prices.rename("consequent_price"),
                left_on="consequents_str",
                right_index=True,
                how="left",
            )
            rules["consequent_price"] = rules["consequent_price"].fillna(0)
            rules["expected_value"] = rules["confidence"] * rules["consequent_price"]

        rules["antecedents"] = rules["antecedents"].apply(lambda x: ", ".join(list(x)))
        rules["consequents"] = rules["consequents"].apply(lambda x: ", ".join(list(x)))
        rules = rules.sort_values(["expected_value"], ascending=[False])

        return rules[
            [
                "antecedents",
                "consequents",
                "confidence",
                "lift",
                "consequent_price",
                "expected_value",
            ]
        ]
    except Exception as e:
        st.error(f"Gagal menjalankan Market Basket Analysis: {e}")
        st.exception(e)
        return pd.DataFrame()


# --- FUNGSI GENERATE INSIGHTS ---
# (Semua fungsi generate_..._insights() masuk ke sini)


@st.cache_data(show_spinner=False)
def generate_gmv_insights(
    kpi, top_selling, bottom_selling, peak_hours, peak_days_of_week
):
    insights = []
    try:
        if not top_selling.empty:
            top_item = top_selling.iloc[0]["Menu"]
            top_qty = top_selling.iloc[0]["Qty"]
            insights.append(
                f"**🚀 Menu Paling Laris:** `{top_item}` (terjual **{top_qty:,.0f} porsi**)."
            )
    except Exception as e:
        print(f"Gagal insight top_selling: {e}")
    try:
        if not bottom_selling.empty:
            bottom_item = bottom_selling.iloc[0]["Menu"]
            bottom_qty = bottom_selling.iloc[0]["Qty"]
            insights.append(
                f"**📉 Menu Jarang Laku:** `{bottom_item}` (hanya terjual **{bottom_qty:,.0f} porsi**)."
            )
    except Exception as e:
        print(f"Gagal insight bottom_selling: {e}")
    try:
        if not peak_days_of_week.empty:
            peak_day_data = peak_days_of_week.sort_values(
                by="Bill Number", ascending=False
            ).iloc[0]
            insights.append(
                f"**🗓️ Hari Paling Ramai:** **{peak_day_data['Day Name']}** ({peak_day_data['Bill Number']:,.0f} transaksi)."
            )
    except Exception as e:
        print(f"Gagal insight peak_days: {e}")
    try:
        if not peak_hours.empty:
            peak_hour_data = peak_hours.sort_values(
                by="Bill Number", ascending=False
            ).iloc[0]
            insights.append(
                f"**🕒 Jam Paling Ramai:** Pukul **{peak_hour_data['Hour']}:00** ({peak_hour_data['Bill Number']:,.0f} transaksi)."
            )
    except Exception as e:
        print(f"Gagal insight peak_hours: {e}")
    try:
        atv = kpi.get("Rata-rata Nilai Transaksi (ATV)", 0)
        ipb = kpi.get("Item per Transaksi (IPB)", 0)
        insights.append(
            f"**💸 Pola Belanja:** Rata-rata pelanggan menghabiskan **{format_rupiah(atv)}** ({ipb:.2f} item/transaksi)."
        )
    except Exception as e:
        print(f"Gagal insight kpi: {e}")
    insights.append(
        "**💡 Catatan Profit:** Cek tab **'💰 COGS & Profit'** untuk menu paling *profit*. Tab ini hanya menganalisis *penjualan*."
    )
    return insights


@st.cache_data(show_spinner=False)
def generate_cogs_insights(profit_df):
    insights = []
    if profit_df is None or profit_df.empty:
        return ["Tidak ada data profit untuk dianalisis."]
    try:
        total_revenue = profit_df["Total Revenue (Rp)"].sum()
        total_profit = profit_df["Total Profit (Rp)"].sum()
        avg_margin_percent = (
            (total_profit / total_revenue) * 100 if total_revenue > 0 else 0
        )
        insights.append(
            f"**💰 Profitabilitas Umum:** Profit kotor **{format_rupiah(total_profit)}** dari revenue **{format_rupiah(total_revenue)}** (Margin **{avg_margin_percent:,.1f}%**)."
        )

        df_valid = profit_df[profit_df["Qty"] > 0].copy()
        if not df_valid.empty:
            top_profit_item = df_valid.nlargest(1, "Total Profit (Rp)").iloc[0]
            insights.append(
                f"**🏆 Bintang Profit (Rp):** `{top_profit_item['Menu']}` (Menghasilkan **{format_rupiah(top_profit_item['Total Profit (Rp)'])}**)."
            )

            top_margin_item = df_valid.nlargest(1, "Margin (%)").iloc[0]
            insights.append(
                f"**📈 Efisiensi Terbaik (%):** `{top_margin_item['Menu']}` (Margin **{format_persen(top_margin_item['Margin (%)'])}**)."
            )

            bottom_profit_item = df_valid.nsmallest(1, "Total Profit (Rp)").iloc[0]
            insights.append(
                f"**💸 Perlu Perhatian (Rp):** `{bottom_profit_item['Menu']}` (Hanya profit **{format_rupiah(bottom_profit_item['Total Profit (Rp)'])}**)."
            )
    except Exception as e:
        print(f"Gagal generate insight COGS: {e}")
        insights.append(f"Gagal membuat insight: {e}")
    return insights


@st.cache_data(show_spinner=False)
def generate_hr_insights(time_data, waiter_data):
    insights = []
    if (time_data is None or time_data.empty) and (
        waiter_data is None or waiter_data.empty
    ):
        return ["Tidak ada data SDM & Waktu untuk dianalisis."]
    try:
        if time_data is not None and not time_data.empty:
            peak_time_sales = time_data.nlargest(1, "Total_Penjualan").iloc[0]
            insights.append(
                f"**🕒 Waktu Emas (Penjualan):** Sesi **{peak_time_sales['Waktu Kunjungan']}** (Menyumbang **{format_rupiah(peak_time_sales['Total_Penjualan'])}**)."
            )
            peak_time_trx = time_data.nlargest(1, "Jumlah_Transaksi").iloc[0]
            insights.append(
                f"** busiest Waktu Tersibuk (Transaksi):** Sesi **{peak_time_trx['Waktu Kunjungan']}** (**{format_angka_bulat(peak_time_trx['Jumlah_Transaksi'])}** transaksi)."
            )
    except Exception as e:
        print(f"Gagal generate insight time_data: {e}")
    try:
        if waiter_data is not None and not waiter_data.empty:
            top_waiter = waiter_data.nlargest(1, "Total_Penjualan").iloc[0]
            insights.append(
                f"**🏆 Waiter Performa Terbaik:** **{top_waiter['Waiter']}** (Sales **{format_rupiah(top_waiter['Total_Penjualan'])}**)."
            )
            avg_sales_per_waiter = waiter_data["Total_Penjualan"].mean()
            insights.append(
                f"**🧑‍🍳 Performa Tim (Top 10):** Rata-rata sales per waiter adalah **{format_rupiah(avg_sales_per_waiter)}**."
            )
    except Exception as e:
        print(f"Gagal generate insight waiter_data: {e}")
    return insights


@st.cache_data(show_spinner=False)
def generate_target_insights(kpi_dict):
    insights = []
    try:
        sisa_hari = kpi_dict.get("sisa_hari", 0)
        if sisa_hari <= 0:
            pencapaian_persen = kpi_dict.get("pencapaian_persen", 0)
            if pencapaian_persen >= 1:
                insights.append(
                    {
                        "type": "success",
                        "text": f"**TARGET TERCAPAI:** Selamat! Bulan ditutup dengan **{pencapaian_persen*100:,.1f}%** dari target.",
                    }
                )
            else:
                insights.append(
                    {
                        "type": "warning",
                        "text": f"**LAPORAN FINAL:** Bulan ditutup dengan **{pencapaian_persen*100:,.1f}%** dari target.",
                    }
                )
            return insights

        proyeksi_pct = kpi_dict.get("proyeksi_vs_target_persen", 0)
        rdr_weekday = kpi_dict.get("rdr_weekday", 0)
        avg_sales_weekday = kpi_dict.get("avg_sales_weekday", 0)
        rdr_weekend = kpi_dict.get("rdr_weekend", 0)
        avg_sales_weekend = kpi_dict.get("avg_sales_weekend", 0)

        if proyeksi_pct > 1.05:
            insights.append(
                {
                    "type": "success",
                    "text": f"**SANGAT ON TRACK:** Proyeksi cerdas Anda **{proyeksi_pct*100:,.1f}%** dari target.",
                }
            )
        elif proyeksi_pct >= 0.98:
            insights.append(
                {
                    "type": "info",
                    "text": f"**ON TRACK:** Proyeksi cerdas Anda **{proyeksi_pct*100:,.1f}%** dari target.",
                }
            )
        else:
            insights.append(
                {
                    "type": "error",
                    "text": f"**OFF TRACK:** Proyeksi cerdas Anda **{proyeksi_pct*100:,.1f}%** dari target. Rencana aksi diperlukan.",
                }
            )

        delta_weekday = rdr_weekday - avg_sales_weekday
        if delta_weekday > 0:
            insights.append(
                {
                    "type": "warning",
                    "text": f"**FOKUS WEEKDAY:** Target harian (Sen-Kam) **{format_rupiah(rdr_weekday)}**, perlu tambahan **{format_rupiah(delta_weekday)}** dari rata-rata.",
                }
            )

        delta_weekend = rdr_weekend - avg_sales_weekend
        if delta_weekend > 0 and rdr_weekend > 0:
            insights.append(
                {
                    "type": "warning",
                    "text": f"**FOKUS WEEKEND:** Target harian (Jum-Min) **{format_rupiah(rdr_weekend)}**, perlu tambahan **{format_rupiah(delta_weekend)}** dari rata-rata.",
                }
            )

    except Exception as e:
        print(f"Gagal generate insight target: {e}")
        insights.append({"type": "error", "text": f"Gagal membuat insight target: {e}"})
    return insights


@st.cache_data(show_spinner=False)
def generate_forecast_insights(forecast_df, last_date):
    insights = []
    if forecast_df is None or forecast_df.empty:
        return ["Tidak ada data ramalan untuk dianalisis."]
    try:
        trend_now = forecast_df[forecast_df["ds"] <= last_date]["trend"].iloc[-1]
        trend_future_end = forecast_df["trend"].iloc[-1]
        trend_pct = (trend_future_end - trend_now) / trend_now if trend_now != 0 else 0

        if trend_pct > 0.01:
            insights.append(
                f"**📈 Tren Jangka Panjang:** Model mendeteksi **tren NAIK** positif ({trend_pct:+.1%})."
            )
        elif trend_pct < -0.01:
            insights.append(
                f"**📉 Tren Jangka Panjang:** Model mendeteksi **tren TURUN** ({trend_pct:+.1%})."
            )
        else:
            insights.append(
                f"**⚖️ Tren Jangka Panjang:** Model mendeteksi tren **STABIL**."
            )

        future_df = forecast_df[forecast_df["ds"] > last_date]
        if len(future_df) >= 7:
            next_7_days_sales = future_df.iloc[:7]["yhat"].sum()
            insights.append(
                f"**🔮 Ramalan 7 Hari:** Model memprediksi penjualan **{format_rupiah(next_7_days_sales)}** untuk 7 hari ke depan."
            )

        insights.append(
            f"**🗓️ Pola Mingguan:** Cek grafik **'Analisis Komponen'** di atas untuk melihat hari terkuat/terlemah."
        )
    except Exception as e:
        print(f"Gagal generate insight forecast: {e}")
    return insights


@st.cache_data(show_spinner=False)
def generate_comparison_insights(
    kpi_A_gmv, kpi_B_gmv, kpi_A_cogs, kpi_B_cogs, kpi_A_waiter, kpi_B_waiter
):
    insights = []

    def calc_pct_delta(a, b):
        if b is None or b == 0:
            return 1.0 if (a is not None and a != 0) else 0.0
        if a is None:
            a = 0
        return (a - b) / b

    try:
        rev_A = kpi_A_gmv.get("Total Pendapatan Kotor", 0)
        rev_B = kpi_B_gmv.get("Total Pendapatan Kotor", 0)
        pct_rev = calc_pct_delta(rev_A, rev_B)

        if pct_rev > 0.01:
            insights.append(
                f"**📈 Performa Penjualan:** Kinerja penjualan **NAIK** signifikan (**{pct_rev:+.1%}**) dibanding periode B."
            )
        elif pct_rev < -0.01:
            insights.append(
                f"**📉 Performa Penjualan:** Kinerja penjualan **TURUN** (**{pct_rev:+.1%}**) dibanding periode B."
            )
        else:
            insights.append(f"**⚖️ Performa Penjualan:** Kinerja penjualan **STABIL**.")

        trx_A = kpi_A_gmv.get("Total Transaksi", 0)
        trx_B = kpi_B_gmv.get("Total Transaksi", 0)
        pct_trx = calc_pct_delta(trx_A, trx_B)
        atv_A = kpi_A_gmv.get("Rata-rata Nilai Transaksi (ATV)", 0)
        atv_B = kpi_B_gmv.get("Rata-rata Nilai Transaksi (ATV)", 0)
        pct_atv = calc_pct_delta(atv_A, atv_B)

        if abs(pct_trx) > abs(pct_atv):
            insights.append(
                f"**💸 Pendorong Penjualan:** Perubahan utama didorong oleh **Jumlah Transaksi** (berubah **{pct_trx:+.1%}**)."
            )
        else:
            insights.append(
                f"**💸 Pendorong Penjualan:** Perubahan utama didorong oleh **Nilai Transaksi (ATV)** (berubah **{pct_atv:+.1%}**)."
            )

        profit_A = kpi_A_cogs[2]
        profit_B = kpi_B_cogs[2]
        pct_profit = calc_pct_delta(profit_A, profit_B)
        margin_A = kpi_A_cogs[3]
        margin_B = kpi_B_cogs[3]
        delta_margin = margin_A - margin_B

        if pct_profit > 0.01:
            insights.append(
                f"**💰 Kinerja Profit:** **Total Profit NAIK** (**{pct_profit:+.1%}**) dengan perubahan margin **{delta_margin:+.1f} poin**."
            )
        elif pct_profit < -0.01:
            insights.append(
                f"**⚠️ Kinerja Profit:** **Total Profit TURUN** (**{pct_profit:+.1%}**) dengan perubahan margin **{delta_margin:+.1f} poin**."
            )
    except Exception as e:
        print(f"Gagal generate insight comparison: {e}")
    return insights


@st.cache_data(show_spinner=False)
def generate_review_insights(
    total_ulasan, avg_rating, nps_score, df_positive_topics, df_negative_topics
):
    insights = []
    if total_ulasan == 0:
        return ["Belum ada data ulasan."]
    try:
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
                f"**👍 Kekuatan Terbesar:** Pelanggan paling sering memuji **{top_positive['Topik']}** "
                f"(disebut **{top_positive['Jumlah']} kali**)."
            )

        # 3. Insight Keluhan Utama (Top Negative)
        if not df_negative_topics.empty:
            top_negative = df_negative_topics.iloc[0]
            insights.append(
                f"**👎 Keluhan Utama:** Area perbaikan paling mendesak adalah **{top_negative['Topik']}** "
                f"(disebut **{top_negative['Jumlah']} kali**)."
            )

        # 4. Insight Rating Bintang 1 (diasumsikan 'Rating 1 (Sangat Buruk)' ditambahkan di UI)
        try:
            if "Rating 1 (Sangat Buruk)" in df_negative_topics["Topik"].values:
                count_bintang_1 = df_negative_topics[
                    df_negative_topics["Topik"] == "Rating 1 (Sangat Buruk)"
                ].iloc[0]["Jumlah"]
                if count_bintang_1 > 0:
                    insights.append(
                        f"**🚨 Peringatan Detractor:** Ada **{count_bintang_1} ulasan bintang 1** "
                        f"yang perlu segera ditindaklanjuti."
                    )
        except Exception:
            pass  # Lewati jika topik tidak ada

    except Exception as e:
        print(f"Gagal generate insight ulasan: {e}")
        insights.append(f"Gagal membuat insight ulasan: {e}")
    return insights


@st.cache_data(show_spinner=False)
def generate_purchase_insights(
    total_cost, cost_by_category, cost_by_supplier, top_items
):
    """
    Menganalisis data pembelian dan menghasilkan insight.
    """
    insights = []
    if total_cost == 0 or top_items.empty:
        return ["Belum ada data pembelian untuk dianalisis."]
    try:
        insights.append(
            f"**🛒 Total Biaya:** Total biaya pembelian (tercatat > Rp 0) "
            f"adalah **{format_rupiah(total_cost)}**."
        )
        if not cost_by_category.empty:
            top_cat = cost_by_category.iloc[0]
            insights.append(
                f"**🍔 Kategori Terbesar:** Kategori **`{top_cat['Category']}`** "
                f"menghabiskan **{format_rupiah(top_cat['Total'])}**."
            )
        if not cost_by_supplier.empty:
            top_supp = cost_by_supplier.iloc[0]
            insights.append(
                f"**🚚 Supplier Utama:** **`{top_supp['Supplier Name']}`** "
                f"dengan total pembelian **{format_rupiah(top_supp['Total'])}**."
            )
        if not top_items.empty:
            top_item = top_items.iloc[0]
            insights.append(
                f"**💸 Item Termahal:** **`{top_item['Product Name']}`** "
                f"memakan biaya **{format_rupiah(top_item['Total'])}**. Periksa harga beli & waste."
            )
    except Exception as e:
        print(f"Gagal generate insight pembelian: {e}")
        insights.append(f"Gagal membuat insight pembelian: {e}")
    return insights


@st.cache_data(show_spinner=False)
def generate_recommendation_insights(rules_df):
    """
    Menganalisis data aturan rekomendasi dan menghasilkan insight.
    """
    insights = []
    if rules_df is None or rules_df.empty:
        return ["Tidak ada aturan rekomendasi yang cukup kuat untuk dianalisis."]
    try:
        # 1. Potensi Sales Tertinggi
        top_ev_rule = rules_df.nlargest(1, "expected_value").iloc[0]
        insights.append(
            f"**💸 Potensi Sales Tertinggi:** JIKA BELI **`{top_ev_rule['antecedents']}`**, "
            f"TAWARKAN **`{top_ev_rule['consequents']}`**. "
            f"(Potensi sales: **{format_rupiah(top_ev_rule['expected_value'])}**)."
        )

        # 2. Pasangan Paling Setia (Confidence)
        top_conf_rule = rules_df.nlargest(1, "confidence").iloc[0]
        insights.append(
            f"**🤝 Pasangan Paling Setia:** JIKA BELI **`{top_conf_rule['antecedents']}`**, "
            f"TAWARKAN **`{top_conf_rule['consequents']}`**. "
            f"(**{top_conf_rule['confidence']:.1%}** pelanggan juga membeli B)."
        )

        # 3. Koneksi Terkuat (Lift)
        top_lift_rule = rules_df.nlargest(1, "lift").iloc[0]
        insights.append(
            f"**🔗 Koneksi Terkuat (Lift):** Pasangan **`{top_lift_rule['antecedents']}`** "
            f"& **`{top_lift_rule['consequents']}`** "
            f"(**{top_lift_rule['lift']:.1f}x lebih mungkin** dibeli bersamaan)."
        )

        insights.append(
            "**💡 Aksi:** Gunakan filter di atas untuk mengeksplorasi rekomendasi spesifik."
        )
    except Exception as e:
        print(f"Gagal generate insight rekomendasi: {e}")
        insights.append(f"Gagal membuat insight: {e}")
    return insights
