import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib import cm
import matplotlib.colors as mcolors
import numpy as np
from io import BytesIO
from PIL import Image

Image.MAX_IMAGE_PIXELS = None

st.set_page_config(layout="wide")
st.title("📊 MACC Curve Generator")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("Pilih worksheet", xls.sheet_names)

    st.markdown("---")
    st.subheader("📌 Tentukan Baris Header (Nama Kolom)")
    header_row = st.number_input("Gunakan baris ke-berapa sebagai header?", min_value=1, value=2)

    df = pd.read_excel(xls, sheet_name=sheet, header=int(header_row) - 1)

    st.markdown("---")
    st.subheader("🧩 Pilih Kolom untuk Setiap Komponen")

    col1, col2 = st.columns(2)

    with col1:
        col_project = st.selectbox("Kolom Nama Proyek", df.columns)
        col_x = st.selectbox("Kolom Sumbu X", df.columns)

        bcol1, bcol2 = st.columns([1, 2])
        benchmark_col = bcol1.selectbox("Kolom Benchmark (opsional)", ["(Tidak Ada)"] + list(df.columns))
        benchmark_label = bcol2.text_input("Label Benchmark", value="/tCO2e Benchmark")

    with col2:
        col_middle = st.selectbox("Kolom Nilai Tengah Batang", df.columns)
        col_top = st.selectbox("Kolom Nilai Ujung Batang", df.columns)

    benchmark_value = None
    if benchmark_col != "(Tidak Ada)":
        benchmark_series = df[benchmark_col].dropna()
        if not benchmark_series.empty:
            benchmark_value = benchmark_series.values[0]
            st.success(f"Garis benchmark akan ditampilkan di posisi: {benchmark_value:.2f}")

    st.markdown("---")
    st.subheader("⚙️ Pengaturan Tambahan")

    custom_scale = st.checkbox("Gunakan Skala Manual pada Sumbu Y")

    if custom_scale:
        y_min = st.number_input("Minimum Y-axis", value=-2500)
        y_max = st.number_input("Maksimum Y-axis", value=3000)
    else:
        y_min = None
        y_max = None

    x_label = st.text_input("Label Sumbu X", value="Avoided Emission (tCO2e per Tahun)")
    y_label = st.text_input("Label Sumbu Y", value="Abatement Cost (USD per tCO2e)")

    st.markdown("---")
    st.subheader("🎨 Pilih Skema Warna Batang")
    colormaps_available = ['PuBuGn', 'cool', 'Blues', 'plasma', 'viridis', 'cividis']
    selected_colormap = st.selectbox("Pilih Skema Warna", colormaps_available)

    def preview_colormap(cmap_name, n=20):
        gradient = np.linspace(0, 1, n).reshape(1, n)
        fig, ax = plt.subplots(figsize=(5, 0.4))
        cmap = cm.get_cmap(cmap_name)
        ax.imshow(gradient, aspect='auto', cmap=cmap)
        ax.set_axis_off()
        st.pyplot(fig)

    st.markdown("🔍 Pratinjau Skema Warna Terpilih:")
    preview_colormap(selected_colormap)

    def plot_macc(df_subset, title, y_min=None, y_max=None, ax=None,
                  x_label="", y_label="", suppress_output=False,
                  benchmark_value=None, benchmark_label=""):

        df_subset = df_subset.reset_index(drop=True)
        x_start = [0]
        for val in df_subset['AF'].tolist()[:-1]:
            x_start.append(x_start[-1] + val)

        fig_width = max(10, len(df_subset) * 0.4)
        if ax is None:
            fig, ax = plt.subplots(figsize=(fig_width, 6))
            is_standalone = True
        else:
            is_standalone = False

        if y_min is not None and y_max is not None:
            if y_min >= y_max:
                y_min, y_max = min(y_min, y_max - 1), max(y_min + 1, y_max)
            ax.set_ylim(bottom=y_min, top=y_max)
        else:
            all_values = df_subset['MAC']
            margin = (all_values.max() - all_values.min()) * 0.1
            ax.set_ylim(bottom=all_values.min() - margin, top=all_values.max() + margin)

        unique_projects = df_subset['Project'].unique()
        cmap = cm.get_cmap(selected_colormap, len(unique_projects))
        project_to_color = {proj: cmap(i) for i, proj in enumerate(unique_projects)}

        for i, row in df_subset.iterrows():
            start = x_start[i]
            width = row['AF']
            mid_x = start + width / 2
            color = project_to_color.get(row['Project'], 'gray')
            height = row['MAC']

            ax.bar(x=start, height=height, width=width, align='edge', bottom=0, color=color, label=row['Project'])

            offset = 0.05 * abs(height) if abs(height) > 1 else 1
            pos_text = height + offset if height >= 0 else height - offset

            ax.text(mid_x, pos_text,
                    f"{row['MAC']:,.0f}".replace(",", "."),
                    ha='center', va='bottom' if height >= 0 else 'top', fontsize=6)

            ax.text(mid_x, height / 2,
                    f"{row['NPV']:,.0f}".replace(",", "."),
                    ha='center', va='center',
                    fontsize=6, color='white' if abs(height) > 500 else 'black')

            ax.text(mid_x, 0,
                    f"{row['AF']:,.0f}".replace(",", "."),
                    ha='center', va='bottom', fontsize=6)

        if benchmark_value is not None:
            formatted_label = benchmark_label.replace("CO2e", "CO₂e").replace("CO2", "CO₂")
            ax.axhline(benchmark_value, color='cyan', linestyle='--', linewidth=1)

            ax.text(
                0.03, 0.98,
                f"${benchmark_value:.0f}{formatted_label}",
                transform=ax.transAxes,
                fontsize=5,
                color='black',
                verticalalignment='top',
                horizontalalignment='left',
                bbox=dict(facecolor='white', edgecolor='none', pad=0)
            )

            for i in range(4):
                ax.text(
                    0.01 + i * 0.005, 0.973,
                    "-",
                    transform=ax.transAxes,
                    fontsize=8,
                    color='cyan',
                    ha='center', va='center'
                )

        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.spines['left'].set_visible(True)
        ax.axhline(0, color='black', linewidth=0.8)
        ax.set_ylabel(y_label, fontsize=10)
        ax.set_xlabel(x_label, fontsize=10)
        ax.set_xticks([])
        ax.yaxis.set_major_formatter(ticker.StrMethodFormatter('{x:,.0f}'))
        ax.set_title(title, fontsize=12)

        handles, labels = ax.get_legend_handles_labels()
        unique = dict(zip(labels, handles))
        ax.legend(unique.values(), unique.keys(), fontsize=6, bbox_to_anchor=(1.01, 1), loc='upper left')

        if is_standalone and not suppress_output:
            st.pyplot(fig)

        return ax.figure

    if st.button("Buat Grafik"):
        df_clean = df[[col_project, col_x, col_middle, col_top]].dropna()
        df_clean.columns = ['Project', 'AF', 'NPV', 'MAC']

        df_clean['AF'] = pd.to_numeric(df_clean['AF'], errors='coerce')
        df_clean['NPV'] = pd.to_numeric(df_clean['NPV'], errors='coerce')
        df_clean['MAC'] = pd.to_numeric(df_clean['MAC'], errors='coerce')
        df_clean = df_clean.dropna().sort_values(by='MAC').reset_index(drop=True)

        fig = plot_macc(df_clean, "MACC Curve", y_min, y_max,
                        x_label=x_label, y_label=y_label,
                        benchmark_value=benchmark_value, benchmark_label=benchmark_label,
                        suppress_output=True)
        st.session_state['fig'] = fig

    st.markdown("<br><br>", unsafe_allow_html=True)

    if 'fig' in st.session_state:
        st.subheader("📈 Grafik MACC Curve (Satu Halaman)")
        st.pyplot(st.session_state['fig'])

        buffer = BytesIO()
        st.session_state['fig'].savefig(buffer, format="png", bbox_inches="tight")
        st.download_button("⬇️ Download PNG", buffer.getvalue(), file_name="MACC_Curve.png", mime="image/png")

        buffer_pdf = BytesIO()
        st.session_state['fig'].savefig(buffer_pdf, format="pdf", bbox_inches="tight")
        st.download_button("⬇️ Download PDF", buffer_pdf.getvalue(), file_name="MACC_Curve.pdf", mime="application/pdf")
