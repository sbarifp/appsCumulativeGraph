import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib import cm
import numpy as np
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📊 MACC Curve Generator")

# --- Upload file ---
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("Pilih worksheet", xls.sheet_names)
    df = pd.read_excel(xls, sheet_name=sheet)

    st.markdown("---")
    st.subheader("🧩 Pilih Kolom untuk Setiap Komponen")

    col_project = st.selectbox("Kolom Nama Proyek", df.columns)
    col_x = st.selectbox("Kolom Sumbu X (AF / Pengurangan Emisi)", df.columns)
    col_middle = st.selectbox("Kolom Nilai Tengah Batang (NPV)", df.columns)
    col_top = st.selectbox("Kolom Nilai Ujung Batang (MAC)", df.columns)

    st.markdown("---")
    st.subheader("⚙️ Parameter Tambahan")
    mac_min = st.number_input("Batas bawah MAC", value=-2500)
    mac_max = st.number_input("Batas atas MAC", value=3000)

    x_label = st.text_input("Label Sumbu X", value="Avoided Emission (Ton CO2e/Tahun)")
    y_label = st.text_input("Label Sumbu Y", value="Abatement Cost (USD/Ton CO2e)")

    if st.button("Buat Grafik") or 'fig1' not in st.session_state:
        df_clean = df[[col_project, col_x, col_middle, col_top]].dropna()
        df_clean.columns = ['Project', 'AF', 'NPV', 'MAC']

        df_clean['AF'] = pd.to_numeric(df_clean['AF'], errors='coerce')
        df_clean['NPV'] = pd.to_numeric(df_clean['NPV'], errors='coerce')
        df_clean['MAC'] = pd.to_numeric(df_clean['MAC'], errors='coerce')
        df_clean = df_clean.dropna().sort_values(by='MAC').reset_index(drop=True)

        df_clean['MAC_clipped'] = df_clean['MAC'].clip(lower=mac_min, upper=mac_max)

        unique_projects = df_clean['Project'].unique()
        colors = cm.tab20.colors
        project_to_color = {proj: colors[i % len(colors)] for i, proj in enumerate(unique_projects)}

        def plot_macc(df_subset, title, y_min, y_max, use_clipped=False, ax=None, x_label="", y_label="", suppress_output=False):
            df_subset = df_subset.reset_index(drop=True)
            x_start = [0]
            for val in df_subset['AF'].tolist()[:-1]:
                x_start.append(x_start[-1] + val)

            if ax is None:
                fig, ax = plt.subplots(figsize=(12, 4))
                is_standalone = True
            else:
                is_standalone = False

            ax.set_ylim(y_min, y_max)

            for i, row in df_subset.iterrows():
                start = x_start[i]
                width = row['AF']
                mid_x = start + width / 2
                color = project_to_color.get(row['Project'], 'gray')

                height = row['MAC_clipped'] if use_clipped else row['MAC']

                ax.bar(x=start, height=height, width=width, align='edge', bottom=0, color=color, label=row['Project'])

                if y_min < height < y_max:
                    ax.text(mid_x, height + (100 if height >= 0 else -100), f"{row['MAC']:.2f}",
                            ha='center', va='bottom' if height >= 0 else 'top', fontsize=5)
                    ax.text(mid_x, height/2, f"{row['NPV']:,.0f}", ha='center', va='center',
                            fontsize=5, color='white' if abs(height) > 500 else 'black')

                ax.text(mid_x, 0, f"{row['AF']:.2f}", ha='center', va='bottom', fontsize=5)

            for spine in ax.spines.values():
                spine.set_visible(False)
            ax.spines['left'].set_visible(True)
            ax.axhline(0, color='black', linewidth=0.8)
            ax.set_ylabel(y_label, fontsize=9)
            ax.set_xlabel(x_label, fontsize=9)
            ax.set_xticks([])
            ax.yaxis.set_major_formatter(ticker.StrMethodFormatter('{x:,.0f}'))
            ax.set_title(title, fontsize=10)

            handles, labels = ax.get_legend_handles_labels()
            unique = dict(zip(labels, handles))
            ax.legend(unique.values(), unique.keys(), fontsize=5, bbox_to_anchor=(1.01, 1), loc='upper left')

            if is_standalone and not suppress_output:
                st.pyplot(fig)

            return ax.figure

        df_extreme = df_clean[~df_clean['MAC'].between(mac_min, mac_max)]
        ymin, ymax = 0, 1
        if not df_extreme.empty:
            ymax = df_extreme['MAC'].max() * 1.1
            ymin = df_extreme['MAC'].min() * 1.1 if df_extreme['MAC'].min() < 0 else 0

        fig_gab = plt.figure(figsize=(14, 8), constrained_layout=True)
        ax1, ax2 = fig_gab.subplots(2, 1)
        plot_macc(df_clean, "MACC Curve - Bagian 1 (Gabungan)", mac_min, mac_max, use_clipped=True, ax=ax1, x_label=x_label, y_label=y_label, suppress_output=True)
        if not df_extreme.empty:
            plot_macc(df_extreme, "MACC Curve - Bagian 2 (Gabungan)", ymin, ymax, use_clipped=False, ax=ax2, x_label=x_label, y_label=y_label, suppress_output=True)

        st.session_state['fig_gabungan'] = fig_gab
        st.session_state['fig1'] = plot_macc(df_clean, "MACC Curve - Bagian 1", mac_min, mac_max, use_clipped=True, x_label=x_label, y_label=y_label, suppress_output=True)
        if not df_extreme.empty:
            st.session_state['fig2'] = plot_macc(df_extreme, "MACC Curve - Bagian 2", ymin, ymax, use_clipped=False, x_label=x_label, y_label=y_label, suppress_output=True)

    st.subheader("🖼️ Grafik Gabungan (1 Halaman)")
    if 'fig_gabungan' in st.session_state:
        st.pyplot(st.session_state['fig_gabungan'])
        buffer = BytesIO()
        st.session_state['fig_gabungan'].savefig(buffer, format="png", bbox_inches="tight")
        st.download_button("⬇️ Download PNG Gabungan", buffer.getvalue(), file_name="MACC_Gabungan.png", mime="image/png")

        buffer_pdf = BytesIO()
        st.session_state['fig_gabungan'].savefig(buffer_pdf, format="pdf", bbox_inches="tight")
        st.download_button("⬇️ Download PDF Gabungan", buffer_pdf.getvalue(), file_name="MACC_Gabungan.pdf", mime="application/pdf")

    st.subheader("📈 Grafik MACC Bagian 1 (Terpisah)")
    if 'fig1' in st.session_state:
        fig1 = st.session_state['fig1']
        st.pyplot(fig1)
        buf1 = BytesIO()
        fig1.savefig(buf1, format="png", bbox_inches="tight")
        st.download_button("⬇️ Download PNG Bagian 1", buf1.getvalue(), file_name="MACC_Bagian_1.png", mime="image/png")

        buf1_pdf = BytesIO()
        fig1.savefig(buf1_pdf, format="pdf", bbox_inches="tight")
        st.download_button("⬇️ Download PDF Bagian 1", buf1_pdf.getvalue(), file_name="MACC_Bagian_1.pdf", mime="application/pdf")

    st.subheader("📈 Grafik MACC Bagian 2 (Terpisah)")
    if 'fig2' in st.session_state:
        fig2 = st.session_state['fig2']
        st.pyplot(fig2)
        buf2 = BytesIO()
        fig2.savefig(buf2, format="png", bbox_inches="tight")
        st.download_button("⬇️ Download PNG Bagian 2", buf2.getvalue(), file_name="MACC_Bagian_2.png", mime="image/png")

        buf2_pdf = BytesIO()
        fig2.savefig(buf2_pdf, format="pdf", bbox_inches="tight")
        st.download_button("⬇️ Download PDF Bagian 2", buf2_pdf.getvalue(), file_name="MACC_Bagian_2.pdf", mime="application/pdf")
