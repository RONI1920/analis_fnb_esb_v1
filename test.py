import streamlit as st
import pandas as pd
import altair as alt

# --- Konfigurasi Halaman ---
st.set_page_config(layout="wide")
st.title("Dashboard Interaktif Penjualan Menu 📊")

# --- 1. Widget File Uploader ---
st.header("1. Unggah Laporan Anda")
st.write("Silakan unggah file CSV atau Excel laporan Anda.")

# Membuat file uploader
# Kita izinkan file .csv dan .xlsx
uploaded_file = st.file_uploader("Pilih file laporan...", type=["csv", "xlsx"])

# --- 2. Logika Utama (Hanya berjalan jika file diunggah) ---
if uploaded_file is not None:
    st.header("2. Analisis Data")

    try:
        # --- 3. Memuat Data ---
        df = None

        # Membaca file berdasarkan tipenya
        if uploaded_file.name.endswith(".csv"):
            st.info(f"Membaca file CSV: `{uploaded_file.name}`")
            # Membaca file CSV, dengan asumsi 9 baris header
            df = pd.read_csv(uploaded_file, skiprows=9)

        elif uploaded_file.name.endswith(".xlsx"):
            st.info(f"Membaca file Excel: `{uploaded_file.name}`")
            # Membaca file Excel, dengan asumsi 9 baris header dan di sheet pertama
            df = pd.read_excel(uploaded_file, skiprows=9, sheet_name=0)

        if df is None:
            st.error("Gagal membaca file. Pastikan formatnya benar.")
        else:
            # --- 4. Pembersihan & Persiapan Data ---

            required_cols = ["Menu Category", "Menu", "Qty"]
            if not all(col in df.columns for col in required_cols):
                st.error(
                    f"Error: File Anda tidak memiliki kolom yang diperlukan: {required_cols}"
                )
            else:
                # Mengubah 'Qty' menjadi angka
                df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce")

                # Menghapus baris dimana 'Qty' bukan angka atau 0
                df = df.dropna(subset=["Qty"])
                df = df[df["Qty"] > 0]
                df["Qty"] = df["Qty"].astype(int)

                # --- 5. Agregasi Data ---
                data_kategori = df.groupby("Menu Category")["Qty"].sum().reset_index()
                data_menu_item = (
                    df.groupby(["Menu Category", "Menu"])["Qty"].sum().reset_index()
                )

                # --- 6. Membuat Grafik Interaktif (Altair) ---
                st.write(
                    "Klik pada salah satu batang di **Grafik Kategori** untuk melihat detailnya di **Grafik Menu Item**."
                )

                selection_kategori = alt.selection_point(fields=["Menu Category"])

                # Grafik 1: Kategori Utama
                chart_kategori = (
                    alt.Chart(data_kategori)
                    .mark_bar()
                    .encode(
                        x=alt.X("Menu Category:N", title="Kategori Menu"),
                        y=alt.Y("Qty:Q", title="Total Kuantiti Terjual"),
                        tooltip=["Menu Category", "Qty"],
                        color=alt.condition(
                            selection_kategori,
                            alt.value("orange"),  # Warna saat diklik
                            alt.value("steelblue"),
                        ),  # Warna default
                    )
                    .add_params(selection_kategori)
                    .properties(title="Total Penjualan per Kategori Menu")
                )

                # Grafik 2: Detail Menu Item
                chart_detail = (
                    alt.Chart(data_menu_item)
                    .mark_bar()
                    .encode(
                        x=alt.X("Menu:N", title="Menu Item", sort="-y"),
                        y=alt.Y("Qty:Q", title="Total Kuantiti Terjual"),
                        tooltip=["Menu Category", "Menu", "Qty"],
                    )
                    .transform_filter(selection_kategori)  # Filter berdasarkan seleksi
                    .properties(
                        title="Detail Penjualan per Menu Item (Berdasarkan Kategori Dipilih)"
                    )
                )

                # --- 7. Tampilkan Grafik ---
                combined_chart = alt.vconcat(chart_kategori, chart_detail)
                st.altair_chart(combined_chart, use_container_width=True)

                # Tampilkan data mentah jika diperlukan
                with st.expander("Lihat data yang telah diproses"):
                    st.write("Data Kategori:")
                    st.dataframe(data_kategori)
                    st.write("Data Menu Item:")
                    st.dataframe(data_menu_item)

    except Exception as e:
        st.error(f"Terjadi error saat memproses file: {e}")
        st.write("Pastikan file Anda adalah laporan yang benar.")
        st.write(
            "Catatan: Kode ini mengasumsikan 9 baris pertama adalah header/metadata."
        )

else:
    # Tampilkan pesan ini jika belum ada file yang diunggah
    st.info("Silakan unggah file untuk memulai analisis.")
