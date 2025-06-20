import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import plotly.express as px

st.set_page_config(page_title="🎨 Dashboard Kepatuhan Pajak Daerah", layout="wide")
st.title("🎯 Dashboard Kepatuhan Pajak Daerah")

with st.expander("📌 Panduan Penggunaan Dashboard", expanded=False):
    st.markdown("""
    **Format Excel yang Didukung:**
    - Kolom wajib: `Nama OP`, `TMT`, `STATUS`, `Nm Unit`, 'Klasifikasi'
    - Kolom pembayaran bulan bisa dalam format `Jan-24`, `01/2024`, atau `2024-01-01`

    **Langkah-langkah:**
    1. Upload file Excel sesuai format.
    2. Pilih sheet & tahun pajak.
    3. Gunakan filter UPPPD / Klasifikasi / Status jika diperlukan.
    4. Lihat grafik & download hasilnya.

    ⚠️ Jika ada error, periksa apakah kolom sudah sesuai.
    """)

uploaded_file = st.file_uploader("📁 Upload File Excel", type=["xlsx"])
tahun_pajak = st.number_input("📅 Pilih Tahun Pajak", min_value=2000, max_value=2100, value=2024)

def normalisasi_kolom(df):
    kolom_alias = {
        'tmt': 'TMT', 't.m.t': 'TMT', 'tgl mulai': 'TMT',
        'nama OP': 'Nama Op', 'nama op': 'Nama Op',
        'nm unit': 'Nm Unit', 'unit': 'Nm Unit',
        'kategori': 'KLASIFIKASI', 'klasifikasi': 'KLASIFIKASI',
        'klasifikasi hiburan': 'KLASIFIKASI', 'jenis': 'KLASIFIKASI',
        'status': 'STATUS', 'Status': 'STATUS',
        'nama wp': 'Nama WP', 'nama WP': 'Nama WP', 'Nama WP': 'Nama WP', 'WP': 'Nama WP', 
        'Wajib Pajak': 'Nama WP', 'wajib pajak': 'Nama WP'
    }
    df.columns = [str(col).strip().lower().replace('.', '').replace('_', ' ') for col in df.columns]
    df.columns = [kolom_alias.get(col, col) for col in df.columns]
    return df

def konversi_kolom_bulan(df):
    new_cols = []
    for col in df.columns:
        try:
            # coba ubah ke datetime dan ambil bulan (to_period)
            dt = pd.to_datetime(col, errors='raise')
            # ubah ke awal bulan (format bulanan)
            dt_month = dt.to_period('M').to_timestamp()
            new_cols.append(dt_month)
        except:
            new_cols.append(col)
    df.columns = new_cols
    return df

def hitung_kepatuhan(df, tahun_pajak):
    df['TMT'] = pd.to_datetime(df['TMT'], errors='coerce')
    payment_cols = [col for col in df.columns if isinstance(col, datetime) and col.year == tahun_pajak]

    total_pembayaran = df[payment_cols].sum(axis=1)

    def hitung_bulan_aktif(tmt):
        if pd.isna(tmt): return 0
        if tmt.year < tahun_pajak: return 12
        elif tmt.year > tahun_pajak: return 0
        else: return 12 - tmt.month + 1

    bulan_aktif = df['TMT'].apply(hitung_bulan_aktif)
    bulan_pembayaran = df[payment_cols].gt(0).sum(axis=1)
    rata_rata_pembayaran = total_pembayaran / bulan_pembayaran.replace(0, 1)
    kepatuhan_persen = bulan_pembayaran / bulan_aktif.replace(0, 1) * 100

    def klasifikasi(row):
        if row['bulan_aktif'] == 0 and row['bulan_pembayaran'] == 0:
            return "Belum Aktif"
        elif row['bulan_pembayaran'] == row['bulan_aktif']:
            return "Patuh"
        elif row['bulan_aktif'] - row['bulan_pembayaran'] <= 3:
            return "Kurang Patuh"
        else:
            return "Tidak Patuh"

    df["Total Pembayaran"] = total_pembayaran
    df["bulan_aktif"] = bulan_aktif
    df["bulan_pembayaran"] = bulan_pembayaran
    df["Rata-rata Pembayaran"] = rata_rata_pembayaran
    df["Kepatuhan (%)"] = kepatuhan_persen
    df["Klasifikasi Kepatuhan"] = df.apply(klasifikasi, axis=1)

    return df, payment_cols

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    selected_sheet = st.selectbox("📄 Pilih Nama Sheet", sheet_names)
    df_input = pd.read_excel(xls, sheet_name=selected_sheet)
    df_input = normalisasi_kolom(df_input)
    df_input = konversi_kolom_bulan(df_input)

    required_cols = ["TMT", "STATUS", "KLASIFIKASI", "Nm Unit"]
    missing_cols = [col for col in required_cols if col not in df_input.columns]

    if missing_cols:
        st.error(f"❌ Kolom wajib hilang: {', '.join(missing_cols)}. Harap periksa file Anda.")
    else:
        df_output, payment_cols = hitung_kepatuhan(df_input.copy(), tahun_pajak)

        st.markdown("### 🔍 Filter Data")
        filter_col1, filter_col2, filter_col3 = st.columns(3)
        
        with filter_col1:
            selected_unit = st.selectbox("🏢 UPPPD", ["Semua"] + sorted(df_output["Nm Unit"].dropna().unique().tolist()))
        
        with filter_col2:
            selected_klasifikasi = st.selectbox("📂 Klasifikasi Pajak", ["Semua"] + sorted(df_output["KLASIFIKASI"].dropna().unique().tolist()))
        
        with filter_col3:
            selected_status = st.selectbox("📌 Status OP", ["Semua"] + sorted(df_output["STATUS"].dropna().unique().tolist()))
        
        # Terapkan filter
        if selected_unit != "Semua":
            df_output = df_output[df_output["Nm Unit"] == selected_unit]
        
        if selected_klasifikasi != "Semua":
            df_output = df_output[df_output["KLASIFIKASI"] == selected_klasifikasi]
        
        if selected_status != "Semua":
            df_output = df_output[df_output["STATUS"] == selected_status]()

        st.success("✅ Data berhasil diproses dan difilter!")
        st.dataframe(df_output.head(50), use_container_width=True)



        output = BytesIO()
        df_output.to_excel(output, index=False)
        st.download_button("⬇️ Download Hasil Excel", data=output.getvalue(), file_name="hasil_dashboard.xlsx")

        st.subheader("🥧 Pie Chart Kepatuhan WP")

        if "Klasifikasi Kepatuhan" in df_output.columns and not df_output.empty:
            pie_data = df_output["Klasifikasi Kepatuhan"].value_counts().reset_index()
            pie_data.columns = ["Klasifikasi", "Jumlah"]

            fig_pie = px.pie(
                pie_data,
                names="Klasifikasi",
                values="Jumlah",
                title="Distribusi Kepatuhan WP",
                color_discrete_sequence=px.colors.qualitative.Pastel,
                hole=0.3  # bisa dihapus kalau mau pie chart penuh
            )
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("📭 Tidak ada data kepatuhan yang tersedia untuk ditampilkan dalam pie chart.")


        st.subheader("📈 Tren Pembayaran Pajak per Bulan")
        if payment_cols:
            bulanan = df_output[payment_cols].sum().reset_index()
            bulanan.columns = ["Bulan", "Total Pembayaran"]
            bulanan["Bulan"] = pd.to_datetime(bulanan["Bulan"], errors="coerce")
            bulanan = bulanan[bulanan["Bulan"].dt.year == tahun_pajak]
            
            if bulanan.empty:
                st.warning(f"📭 Tidak ada data pembayaran untuk tahun {tahun_pajak}.")
            else:
                bulanan["BulanSort"] = bulanan["Bulan"]
                bulanan["BulanLabel"] = bulanan["Bulan"].dt.strftime('%b %Y')
                bulanan = bulanan.sort_values("BulanSort")
        
                fig_line = px.line(
                    bulanan,
                    x="Bulan",
                    y="Total Pembayaran",
                    title="Total Pembayaran Pajak per Bulan",
                    markers=True
                )
                st.plotly_chart(fig_line, use_container_width=True)
        else:
            st.warning("📭 Tidak ditemukan kolom pembayaran murni yang valid.")


        st.subheader("🏅 Top 5 Objek Pajak Berdasarkan Total Pembayaran")
        top_wp = (
            df_output[["Nama Op", "Total Pembayaran", "Nm Unit", "KLASIFIKASI"]]
            .groupby(["Nama Op", "Nm Unit", "KLASIFIKASI"], as_index=False)
            .sum()
            .sort_values("Total Pembayaran", ascending=False)
            .head(5)
        )
        st.dataframe(top_wp.style.format({"Total Pembayaran": "Rp{:,.0f}"}), use_container_width=True)
