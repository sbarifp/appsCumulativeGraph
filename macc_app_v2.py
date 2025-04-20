import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib import cm
import numpy as np
from io import BytesIO
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

st.set_page_config(layout="wide")
st.title("📊 Curve Generator")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("Pilih worksheet", xls.sheet_names)

    st.markdown("---")
    st.subheader("📌 Tentukan Baris Header (Nama Kolom)")
    header_row = st.number_input("Gunakan baris ke-berapa sebagai header?", min_value=1, value=2)
    df = pd.read_excel(xls, sheet_name=sheet, header=int(header_row) - 1)

    st.markdown("---")
    st.subheader("🧩 Pilih Kolom untuk Setiap Komponen (termasuk Benchmark)")

    col1, col2 = st.columns(2)

    with col1:
        col_project = st.selectbox("Kolom Nama Proyek", df.columns)
        col_x = st.selectbox("Kolom Sumbu X", df.columns)

        bcol1, bcol2 = st.columns([1, 2])
        benchmark_col = bcol1.selectbox("Benchmark", ["(Tidak Ada)"] + list(df.columns))
        benchmark_label = bcol2.text_input("Label", value="/tCO2e Benchmark")

    with col2:
        col_middle = st.selectbox("Kolom Nilai Tengah Batang", df.columns)
        col_top = st.selectbox("Kolom Nilai Ujung Batang", df.columns)

    benchmark_value = None
    if benchmark_col != "(Tidak Ada)":
        benchmark_value = df[benchmark_col].dropna().values[0]
        st.success(f"Garis benchmark akan ditampilkan di posisi: {benchmark_value:.2f}")

    st.markdown("---")
    st.subheader("⚙️ Parameter Tambahan")
    mac_min = st.number_input("Batas bawah", value=-2500)
    mac_max = st.number_input("Batas atas", value=3000)
    x_label = st.text_input("Label Sumbu X", value="Avoided Emission (tCO2e per Tahun)")
    y_label = st.text_input("Label Sumbu Y", value="Abatement Cost (USD per tCO2e)")

    df_clean = df[[col_project, col_x, col_middle, col_top]].dropna()
    df_clean.columns = ['Project', 'Nilai_Xaxis', 'Nilai_Tengah', 'Nilai_Ujung']
    df_clean['Nilai_Xaxis'] = pd.to_numeric(df_clean['Nilai_Xaxis'], errors='coerce')
    df_clean['Nilai_Tengah'] = pd.to_numeric(df_clean['Nilai_Tengah'], errors='coerce')
    df_clean['Nilai_Ujung'] = pd.to_numeric(df_clean['Nilai_Ujung'], errors='coerce')
    df_clean = df_clean.dropna().sort_values(by='Nilai_Ujung').reset_index(drop=True)
    df_clean['MAC_clipped'] = df_clean['Nilai_Ujung'].clip(lower=mac_min, upper=mac_max)

    unique_projects = df_clean['Project'].unique()

    st.markdown("---")
    st.subheader("🎨 Pilih Skema Warna Batang")
    colormaps_available = ['PuBuGn', 'cool', 'Blues', 'plasma', 'viridis', 'cividis']
    selected_colormap = st.selectbox("Pilih Skema Warna", colormaps_available)

    def preview_colormap(cmap_name, n=10):
        gradient = np.linspace(0.1, 0.9, n).reshape(1, n)
        fig, ax = plt.subplots(figsize=(5, 0.4))
        cmap = cm.get_cmap(cmap_name)
        ax.imshow(gradient, aspect='auto', cmap=cmap)
        ax.set_axis_off()
        st.pyplot(fig)

    st.markdown("🔍 Pratinjau Skema Warna Terpilih:")
    preview_colormap(selected_colormap, n=len(unique_projects))

    cmap_base = cm.get_cmap(selected_colormap)
    color_vals = np.linspace(0.1, 0.9, len(unique_projects))
    project_to_color = {proj: cmap_base(val) for proj, val in zip(unique_projects, color_vals)}

    if st.button("Buat Grafik") or 'fig1' not in st.session_state:

        def plot_macc(df_subset, title, y_min, y_max, use_clipped=False, ax=None,
                      x_label="", y_label="", suppress_output=False,
                      benchmark_value=None, benchmark_label=""):

            df_subset = df_subset.reset_index(drop=True)
            x_start = [0]
            for val in df_subset['Nilai_Xaxis'].tolist()[:-1]:
                x_start.append(x_start[-1] + val)

            if ax is None:
                fig, ax = plt.subplots(figsize=(12, 4))
                is_standalone = True
            else:
                is_standalone = False

            ax.set_ylim(y_min, y_max)

            for i, row in df_subset.iterrows():
                start = x_start[i]
                width = row['Nilai_Xaxis']
                mid_x = start + width / 2
                color = project_to_color.get(row['Project'], 'gray')
                height = row['MAC_clipped'] if use_clipped else row['Nilai_Ujung']

                ax.bar(x=start, height=height, width=width, align='edge', bottom=0, color=color, label=row['Project'])

                offset = 0.05 * abs(height) if abs(height) > 1 else 1
                pos_text = height + offset if height >= 0 else height - offset

                if y_min < height < y_max:
                    ax.text(mid_x, pos_text,
                            f"{row['Nilai_Ujung']:,.0f}".replace(",", "."),
                            ha='center', va='bottom' if height >= 0 else 'top', fontsize=5)
                    ax.text(mid_x, height / 2,
                            f"{row['Nilai_Tengah']:,.0f}".replace(",", "."),
                            ha='center', va='center',
                            fontsize=5, color='white' if abs(height) > 500 else 'black')

                ax.text(mid_x, 0, f"{row['Nilai_Xaxis']:,.0f}".replace(",", "."), ha='center', va='bottom', fontsize=5)

            if benchmark_value is not None:
                formatted_label = benchmark_label.replace("CO2e", "CO\u2082e").replace("CO2", "CO\u2082")
                ax.axhline(benchmark_value, color='cyan', linestyle='--', linewidth=1)
                ax.text(
                    x=ax.get_xlim()[0], y=y_max,
                    s=f"${benchmark_value:.0f}{formatted_label}",
                    color='cyan', fontsize=8, fontweight='bold',
                    ha='left', va='top'
                )

            for spine in ax.spines.values():
                spine.set_visible(False)
            ax.spines['left'].set_visible(True)
            ax.axhline(0, color='black', linewidth=0.8)
            ax.set_ylabel(y_label, fontsize=9)
            ax.set_xlabel(x_label, fontsize=9)
            ax.set_xticks([])
            ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f"{int(x):,}".replace(",", ".")))
            ax.set_title(title, fontsize=10)

            handles, labels = ax.get_legend_handles_labels()
            unique = dict(zip(labels, handles))
            ax.legend(unique.values(), unique.keys(), fontsize=5, bbox_to_anchor=(1.01, 1), loc='upper left')

            if is_standalone and not suppress_output:
                st.pyplot(fig)

            return ax.figure

        df_extreme = df_clean[~df_clean['Nilai_Ujung'].between(mac_min, mac_max)]
        ymin, ymax = 0, 1
        if not df_extreme.empty:
            ymax = df_extreme['Nilai_Ujung'].max() * 1.1
            ymin = df_extreme['Nilai_Ujung'].min() * 1.1 if df_extreme['Nilai_Ujung'].min() < 0 else 0

        fig_gab = plt.figure(figsize=(14, 8), constrained_layout=True)
        ax1, ax2 = fig_gab.subplots(2, 1)
        plot_macc(df_clean, "MACC Curve - Bagian 1 (Gabungan)", mac_min, mac_max,
                  use_clipped=True, ax=ax1, x_label=x_label, y_label=y_label,
                  suppress_output=True, benchmark_value=benchmark_value, benchmark_label=benchmark_label)
        if not df_extreme.empty:
            plot_macc(df_extreme, "MACC Curve - Bagian 2 (Gabungan)", ymin, ymax,
                      use_clipped=False, ax=ax2, x_label=x_label, y_label=y_label,
                      suppress_output=True, benchmark_value=benchmark_value, benchmark_label=benchmark_label)

        st.session_state['fig_gabungan'] = fig_gab
        st.session_state['fig1'] = plot_macc(df_clean, "MACC Curve - Bagian 1",
                                             mac_min, mac_max, use_clipped=True,
                                             x_label=x_label, y_label=y_label,
                                             suppress_output=True, benchmark_value=benchmark_value, benchmark_label=benchmark_label)
        if not df_extreme.empty:
            st.session_state['fig2'] = plot_macc(df_extreme, "MACC Curve - Bagian 2",
                                                 ymin, ymax, use_clipped=False,
                                                 x_label=x_label, y_label=y_label,
                                                 suppress_output=True, benchmark_value=benchmark_value, benchmark_label=benchmark_label)

    # ⬇️ Tambahkan jarak antar section
    st.markdown("<br><br>", unsafe_allow_html=True)

    st.subheader("🖼️ Grafik Gabungan (1 Halaman)")
    if 'fig_gabungan' in st.session_state:
        st.pyplot(st.session_state['fig_gabungan'])
        buffer = BytesIO()
        st.session_state['fig_gabungan'].savefig(buffer, format="png", bbox_inches="tight")
        st.download_button("⬇️ Download PNG Gabungan", buffer.getvalue(), file_name="MACC_Gabungan.png", mime="image/png")

        buffer_pdf = BytesIO()
        st.session_state['fig_gabungan'].savefig(buffer_pdf, format="pdf", bbox_inches="tight")
        st.download_button("⬇️ Download PDF Gabungan", buffer_pdf.getvalue(), file_name="MACC_Gabungan.pdf", mime="application/pdf")
        st.markdown("<br>", unsafe_allow_html=True)

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
        st.markdown("<br>", unsafe_allow_html=True)

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
