import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pycountry
from datetime import datetime, date
import plotly.io as pio
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Image, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import HRFlowable
import io
import base64
import tempfile
import os
from preprocessing import (
    comprehensive_clean, extract_city_comprehensive,
    clean_extracted_cities_df, CITIES_INDONESIA
)


# ========================
# FUNGSI BANTU
# ========================
def get_iso3(name):
    try:
        return pycountry.countries.lookup(name).alpha_3
    except:
        return name  # fallback kalau tidak ketemu

def safe_write_image(fig, format="png", width=800, height=600, scale=2):
    """
    Wrapper aman untuk pio.to_image.
    Jika Kaleido error karena Chrome/Chromium tidak ada,
    otomatis jalankan plotly_get_chrome untuk download Chromium portable.
    """
    import subprocess, sys
    try:
        return pio.to_image(fig, format=format, width=width, height=height, scale=scale)
    except RuntimeError as e:
        if "Kaleido requires Google Chrome" in str(e):
            print("⚠️ Chrome/Chromium tidak ditemukan. Menjalankan plotly_get_chrome...")
            subprocess.check_call([sys.executable, "-m", "plotly.io._utils", "plotly_get_chrome"])
            # coba ulangi export setelah Chromium terpasang
            return pio.to_image(fig, format=format, width=width, height=height, scale=scale)
        else:
            raise

def create_pdf_report(figures_dict, date_range, total_records, filtered_records):
    """
    Membuat laporan PDF dari semua chart yang ada di dashboard
    """
    # Buffer untuk menyimpan PDF
    buffer = io.BytesIO()
    
    # Buat dokumen PDF
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=72,
        leftMargin=72,
        topMargin=72,
        bottomMargin=18
    )
    
    # Style untuk dokumen
    styles = getSampleStyleSheet()

     # Style untuk header utama dengan emoji
    header_style = ParagraphStyle(
        'CustomHeader',
        parent=styles['Heading1'],
        fontSize=28,
        spaceAfter=20,
        spaceBefore=10,
        alignment=TA_CENTER,
        textColor='#3EADB3',
        fontName='Helvetica-Bold'
    )
    
    # Style untuk subtitle
    subtitle_style = ParagraphStyle(
        'SubtitleStyle',
        parent=styles['Normal'],
        fontSize=14,
        spaceAfter=20,
        alignment=TA_CENTER,
        textColor='#9E6B2B',
        fontName='Helvetica-Oblique'
    )

    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=30,
        alignment=TA_CENTER,
        textColor='#1f4e79'
    )
    
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=16,
        spaceAfter=12,
        spaceBefore=20,
        textColor='#2c5aa0'
    )
    
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=11,
        spaceAfter=12,
    )
    
    # Konten PDF
    story = []
    
    # Judul laporan
    story.append(Paragraph("Dashboard BBKK Surabaya", header_style))
    story.append(Paragraph("Analisis Ekspor OMKABA BBKK Surabaya", subtitle_style))
    story.append(Spacer(1, 30))

    story.append(HRFlowable(width="100%", thickness=1, lineCap='round', color='#e9ecef'))
    story.append(Spacer(1, 20))
    
    # Informasi laporan
    info_text = f"""
    <b>Tanggal Laporan:</b> {datetime.now().strftime('%d %B %Y, %H:%M WIB')}<br/>
    <b>Periode Data:</b> {date_range}<br/>
    <b>Total Records:</b> {filtered_records:,} dari {total_records:,} data
    """
    story.append(Paragraph(info_text, normal_style))
    story.append(Spacer(1, 30))
    
    # Tambahkan setiap chart ke PDF
    for title, fig in figures_dict.items():
        try:
            # Konversi plotly figure ke gambar
            if "Sebaran Kota Perusahaan Eksportir" in title:
                # Extract lat, lon, kota dari fig scatter_mapbox
                lats = fig.data[0].lat
                lons = fig.data[0].lon
                hovertext = getattr(fig.data[0], "hovertext", ["Kota"]*len(lats))
            
                df_geo = pd.DataFrame({
                    "lat": lats,
                    "lon": lons,
                    "Kota": hovertext,
                })
                
                fig_geo = px.scatter_geo(
                    df_geo,
                    lat="lat", lon="lon",
                    hover_name="Kota",
                    title="Sebaran Kota Perusahaan Eksportir"
                )
                img_bytes = safe_write_image(fig_geo, format="png", width=800, height=600, scale=2)
            else:
                img_bytes = safe_write_image(fig, format="png", width=800, height=600, scale=2)
            
            img_buffer = io.BytesIO(img_bytes)

            # Tambahkan judul chart
            story.append(Paragraph(title, heading_style))
            story.append(Spacer(1, 12))
            
            # Tambahkan gambar ke PDF langsung dari buffer
            img = Image(img_buffer, width=6.5*inch, height=4.875*inch)
            story.append(img)
            story.append(Spacer(1, 20))

            # Page break kecuali untuk chart terakhir
            if list(figures_dict.keys()).index(title) < len(figures_dict) - 1:
                story.append(PageBreak())
        
        except Exception as e:
            # Jika ada error, tambahkan pesan error
            error_msg = f"Error saat memproses {title}: {str(e)}"
            story.append(Paragraph(error_msg, normal_style))
            story.append(Spacer(1, 20))

            if "mapbox" in str(e).lower():
                story.append(Paragraph(f"{title}: Tidak bisa diexport ke PDF karena membutuhkan Mapbox API.", normal_style))
            else:
                story.append(Paragraph(f"Error saat memproses {title}: {str(e)}", normal_style))
            story.append(Spacer(1, 20))
    
    # Footer
    story.append(Spacer(1, 30))
    footer_text = "Laporan ini dibuat secara otomatis oleh Dashboard BBKK Surabaya"
    footer_style = ParagraphStyle(
        'Footer',
        parent=styles['Normal'],
        fontSize=9,
        alignment=TA_CENTER,
        textColor='gray'
    )
    story.append(Paragraph(footer_text, footer_style))
    
    # Build PDF
    doc.build(story)
    
    # Return buffer
    buffer.seek(0)
    return buffer


def get_download_link(buffer, filename):
    """
    Membuat link download untuk file PDF
    """
    b64 = base64.b64encode(buffer.getvalue()).decode()
    return f'<a href="data:application/pdf;base64,{b64}" download="{filename}">📄 Download PDF Report</a>'

# ========================
# FUNGSI EXPORT EXCEL
# ========================
def create_excel_report(df_filtered, date_range, total_records, filtered_records):
    """
    Membuat laporan Excel dari data yang sudah difilter dan dipreprocessing
    """
    # Buffer untuk menyimpan Excel
    buffer = io.BytesIO()
    
    # Buat Excel writer
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        # Sheet 1: Data Utama (hanya kolom yang relevan)
        relevant_columns = [
            'Diterbitkan Tanggal',
            'Nama Exportir/Importir', 
            'Alamat Perusahaan',
            'Kota',
            'Jenis Komoditi',
            'Negara Tujuan'
        ]
        
        # Filter hanya kolom yang ada di dataframe
        available_columns = [col for col in relevant_columns if col in df_filtered.columns]
        df_export = df_filtered[available_columns].copy()
        
        # Format tanggal untuk Excel
        if 'Diterbitkan Tanggal' in df_export.columns:
            df_export['Diterbitkan Tanggal'] = df_export['Diterbitkan Tanggal'].dt.strftime('%d-%m-%Y')
        
        # Urutkan berdasarkan tanggal
        df_export = df_export.sort_values('Diterbitkan Tanggal', ascending=False)
        
        # Tulis ke sheet
        df_export.to_excel(writer, sheet_name='Data_Ekspor', index=False)
        
        # Sheet 2: Summary Komoditas
        if 'Jenis Komoditi' in df_filtered.columns:
            commodity_summary = df_filtered['Jenis Komoditi'].value_counts().reset_index()
            commodity_summary.columns = ['Jenis Komoditi', 'Jumlah']
            commodity_summary['Persentase (%)'] = (commodity_summary['Jumlah'] / commodity_summary['Jumlah'].sum() * 100).round(2)
            commodity_summary.to_excel(writer, sheet_name='Summary_Komoditas', index=False)
        
        # Sheet 3: Summary Negara Tujuan
        if 'Negara Tujuan' in df_filtered.columns:
            country_summary = df_filtered['Negara Tujuan'].value_counts().reset_index()
            country_summary.columns = ['Negara Tujuan', 'Jumlah']
            country_summary['Persentase (%)'] = (country_summary['Jumlah'] / country_summary['Jumlah'].sum() * 100).round(2)
            country_summary.to_excel(writer, sheet_name='Summary_Negara', index=False)
        
        # Sheet 4: Summary Kota
        if 'Kota' in df_filtered.columns:
            city_summary = df_filtered['Kota'].value_counts().reset_index()
            city_summary.columns = ['Kota', 'Jumlah']
            city_summary['Persentase (%)'] = (city_summary['Jumlah'] / city_summary['Jumlah'].sum() * 100).round(2)
            city_summary.to_excel(writer, sheet_name='Summary_Kota', index=False)
        
        # Sheet 5: Summary Perusahaan
        if 'Nama Exportir/Importir' in df_filtered.columns:
            company_summary = df_filtered['Nama Exportir/Importir'].value_counts().reset_index()
            company_summary.columns = ['Nama Perusahaan', 'Jumlah']
            company_summary['Persentase (%)'] = (company_summary['Jumlah'] / company_summary['Jumlah'].sum() * 100).round(2)
            company_summary.to_excel(writer, sheet_name='Summary_Perusahaan', index=False)
        
        # Sheet 6: Info Laporan
        info_data = {
            'Keterangan': [
                'Tanggal Generate',
                'Periode Data',
                'Total Records Asli',
                'Records Setelah Filter',
                'Persentase Data Ditampilkan (%)'
            ],
            'Nilai': [
                datetime.now().strftime('%d %B %Y, %H:%M WIB'),
                date_range,
                f"{total_records:,}",
                f"{filtered_records:,}",
                f"{(filtered_records/total_records*100):.2f}%" if total_records > 0 else "0%"
            ]
        }
        info_df = pd.DataFrame(info_data)
        info_df.to_excel(writer, sheet_name='Info_Laporan', index=False)
    
    # Return buffer
    buffer.seek(0)
    return buffer



# ========================
# DASHBOARD STREAMLIT
# ========================
st.set_page_config(page_title="Dashboard Ekspor", layout="wide")

st.markdown("""
<div style="text-align: center; padding: 20px 0; background: linear-gradient(90deg, #f8f9fa 0%, #e9ecef 50%, #f8f9fa 100%); border-radius: 10px; margin-bottom: 30px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
    <div style="font-size: 48px; margin-bottom: 10px;">
        ✈️ 🚢 📦 🌍
    </div>
    <h1 style="color: #3EADB3; margin: 0; font-family: 'Arial', sans-serif; font-weight: bold; text-shadow: 1px 1px 2px rgba(0,0,0,0.1);">
        📊 Dashboard BBKK Surabaya
    </h1>
    <p style="color: #9E6B2B; margin: 10px 0 0 0; font-style: italic; font-size: 20px;">
        Analisis Ekspor OMKABA BBKK Surabaya
    </p>
    <div style="margin-top: 15px; font-size: 24px;">
        📈 📋 🗺️
    </div>
</div>
""", unsafe_allow_html=True)

# Upload file
uploaded_file = st.file_uploader("Upload file CSV/Excel", type=["csv", "xlsx"])

# baca file koordinat kota bawaan (selalu ada di project)
df_city = pd.read_excel("koordinat_kota.xlsx")

if uploaded_file is not None:
    # Baca file
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    
    # Preprocessing
    if "Nama Exportir/Importir" in df.columns:
        df["Nama Exportir/Importir"] = df["Nama Exportir/Importir"].apply(comprehensive_clean)

    if "Alamat Perusahaan" in df.columns:
        df["Kota"] = df["Alamat Perusahaan"].apply(
            lambda x: extract_city_comprehensive(x, CITIES_INDONESIA)
        )
    
    # Konversi kolom tanggal
    df["Diterbitkan Tanggal"] = pd.to_datetime(df["Diterbitkan Tanggal"], dayfirst=True, errors="coerce")
    
    # Remove rows with NaT (Not a Time) values in date column
    df = df.dropna(subset=["Diterbitkan Tanggal"])
    
    st.success("✅ File berhasil diupload!")
    
    # ========================
    # FILTER TANGGAL
    # ========================
    st.sidebar.header("🗓️ Filter Tanggal")
    
    # Deteksi batas minimum dan maksimum tanggal
    if df["Diterbitkan Tanggal"].empty:
        st.error("❌ Tidak ada data valid pada kolom tanggal setelah preprocessing!")
        st.stop()
    
    min_date = df["Diterbitkan Tanggal"].min().date()
    max_date = df["Diterbitkan Tanggal"].max().date()

    if pd.isna(min_date) or pd.isna(max_date):
        st.sidebar.warning("⚠️ Rentang tanggal tidak dapat ditentukan (semua NaT)")
    else:
        st.sidebar.info(f"📅 Rentang data tersedia:\n{min_date.strftime('%d-%m-%y')} - {max_date.strftime('%d-%m-%y')}")
    
    # Checkbox untuk pilih semua rentang tanggal
    select_all_dates = st.sidebar.checkbox("Pilih Semua Rentang Tanggal", value=True)
    
    if select_all_dates:
        # Gunakan seluruh rentang tanggal
        start_date = min_date
        end_date = max_date
        st.sidebar.success("✅ Menampilkan semua data")
    else:
        # Date picker untuk rentang tanggal
        col_start, col_end = st.sidebar.columns(2)
        
        with col_start:
            start_date = st.date_input(
                "Tanggal Mulai",
                value=min_date,
                min_value=min_date,
                max_value=max_date
            )
        
        with col_end:
            end_date = st.date_input(
                "Tanggal Akhir",
                value=max_date,
                min_value=min_date,
                max_value=max_date
            )
        
        # Validasi rentang tanggal
        if start_date > end_date:
            st.sidebar.error("⚠️ Tanggal mulai tidak boleh lebih besar dari tanggal akhir!")
            st.stop()
    
    # Filter data berdasarkan tanggal
    df_filtered = df[
        (df["Diterbitkan Tanggal"].dt.date >= start_date) & 
        (df["Diterbitkan Tanggal"].dt.date <= end_date)
    ]
    
    # Tampilkan informasi filter
    total_records = len(df)
    filtered_records = len(df_filtered)
    
    st.sidebar.metric(
        label="📊 Data Ditampilkan",
        value=f"{filtered_records:,}",
        delta=f"{filtered_records - total_records:,} dari total {total_records:,}"
    )
    
    # Jika tidak ada data setelah filter
    if df_filtered.empty:
        st.warning("⚠️ Tidak ada data dalam rentang tanggal yang dipilih. Silakan pilih rentang tanggal lain.")
        st.stop()

    # ========================
    # TOMBOL EXPORT PDF
    # ========================
    st.sidebar.header("📄 Export Laporan")
    
    if st.sidebar.button("🔄 Generate PDF Report", type="primary"):
        with st.spinner("Sedang membuat laporan PDF..."):
            try:
                # Dictionary untuk menyimpan semua figures
                figures_dict = {}
                
                # Date range string
                if start_date == min_date and end_date == max_date:
                    date_range = f"{min_date.strftime('%d-%m-%Y')} - {max_date.strftime('%d-%m-%Y')} (Semua Data)"
                else:
                    date_range = f"{start_date.strftime('%d-%m-%Y')} - {end_date.strftime('%d-%m-%Y')}"

                # ========================
                # 1. PIE CHART KOMODITAS
                # ========================
                comodity = df_filtered['Jenis Komoditi'].value_counts().reset_index()
                comodity.columns = ["Jenis Komoditi", "Jumlah"]

                total = comodity["Jumlah"].sum()
                comodity["Persentase"] = comodity["Jumlah"] / total * 100
                comodity["Label"] = comodity.apply(
                    lambda x: f"{x['Jenis Komoditi']} ({x['Jumlah']} unit, {x['Persentase']:.2f}%)", axis=1
                )

                n_categories = len(comodity)
                base_size = 390
                additional_size = n_categories * 15

                fig1 = go.Figure(
                    data=[
                        go.Pie(
                            labels=comodity["Jenis Komoditi"],
                            values=comodity["Jumlah"],
                            hole=0.35,
                            textinfo="label+percent+value",
                            texttemplate="<b>%{label}</b><br>%{value} unit<br>(%{percent})",
                            hovertemplate="<b>%{label}</b><br>Jumlah: %{value}<br>Persen: %{percent}<extra></extra>",
                            marker=dict(
                                colors=px.colors.qualitative.Plotly, 
                                line=dict(color="white", width=2)
                            ),
                            textposition="outside",
                            textfont=dict(size=12, family="Arial", color="black"),
                            insidetextorientation='radial',
                            outsidetextfont=dict(size=11),
                            pull=0.03,
                        )
                    ]
                )
            
                fig1.update_layout(
                    showlegend=False,
                    annotations=[],
                    uniformtext_minsize=10,
                    uniformtext_mode='hide',
                    
                    margin=dict(l=5, r=5, t=90, b=35),  # Margin sangat ketat
                    width=base_size + additional_size,  # Ukuran dinamis
                    height=base_size + additional_size,
                    autosize=False,
                    
                    # Hilangkan padding dan spacing yang tidak perlu
                    paper_bgcolor='white',
                    plot_bgcolor='white',
                    
                    # Hilangkan axis dan grid yang tidak diperlukan
                    xaxis=dict(visible=False, showgrid=False, zeroline=False),
                    yaxis=dict(visible=False, showgrid=False, zeroline=False),
                )
            
                fig1.update_traces(
                    marker_line_color='white',
                    marker_line_width=2,
                    textfont_size=11,
                    textposition='outside',
                    textfont_color='black',
                    insidetextfont=dict(size=11, color='white', family='Arial'),
                    outsidetextfont=dict(size=10, family='Arial'),
                    pull=[0.15 if i == 0 else 0.05 for i in range(len(comodity))],
                    rotation=30,
                    hole=0.4
                )
                
                # Atur layout agar lebih compact
                fig1.update_layout(
                    autosize=True,  # Auto-size untuk menyesuaikan konten
                )
                
                figures_dict["1. Distribusi Komoditas"] = fig1

                # ========================
                # 2. BAR CHART NEGARA
                # ========================
                top_10_country = df_filtered['Negara Tujuan'].value_counts().head(10).reset_index()
                top_10_country.columns = ["Negara", "Jumlah"]

                fig2 = px.bar(
                    top_10_country,
                    x="Negara",
                    y="Jumlah",
                    text="Jumlah",
                    color="Jumlah",
                    color_continuous_scale=["#e69795", "#bc656d", "#b03031"],
                    title="Top 10 Negara Tujuan"
                )

                fig2.update_traces(texttemplate='%{text}', textposition='outside')
                fig2.update_layout(
                    xaxis_title="Negara", 
                    yaxis_title="Jumlah",
                    width=800,
                    height=600
                )
                figures_dict["2. Top 10 Negara Tujuan"] = fig2

                # ========================
                # 3. BAR CHART KOTA PERUSAHAAN EKSPOR
                # ========================
                top_10_city = df_filtered['Kota'].value_counts().head(10).reset_index()
                top_10_city.columns = ["Kota", "Jumlah"]
                
                fig = px.bar(
                    top_10_city,
                    x="Kota",
                    y="Jumlah",
                    text="Jumlah", 
                    color="Kota", 
                    color_discrete_sequence=px.colors.qualitative.Plotly,
                    title="Top 10 Kota Perusahaan Eksportir"
                )
                
                fig.update_traces(texttemplate='%{text}', textposition='outside', showlegend=False)
                fig.update_layout(
                    xaxis_title="Kota",
                    yaxis_title="Jumlah",
                    uniformtext_minsize=8,
                    uniformtext_mode='hide',
                    showlegend=False,
                    width=800,
                    height=600
                )
                
                figures_dict["3. Top 10 Kota Perusahaan Eksportir"] = fig

                # ========================
                # 4. BAR CHART PERUSAHAAN EXPORTIR
                # ========================
                top_10_company = df_filtered['Nama Exportir/Importir'].value_counts().head(10).reset_index()
                top_10_company.columns = ["Perusahaan", "Jumlah"]
                
                fig4 = px.bar(
                    top_10_company,
                    x="Perusahaan",
                    y="Jumlah",
                    text="Jumlah", 
                    color="Jumlah",
                    color_continuous_scale="plasma",
                    title="Top 10 Perusahaan Ekspor Paling Banyak"
                )
                
                fig4.update_traces(texttemplate='%{text}', textposition='outside')
                fig4.update_layout(
                    xaxis_title="Perusahaan",
                    yaxis_title="Jumlah",
                    width=800,
                    height=600
                )
                figures_dict["4. Top 10 Perusahaan Ekspor"] = fig4

                # ========================
                # 5. LINE CHART TIMELINE
                # ========================
                timeline = df_filtered.groupby("Diterbitkan Tanggal").size().reset_index(name="Jumlah")
                timeline['Tanggal_Display'] = timeline["Diterbitkan Tanggal"].dt.strftime('%d-%m-%Y')
                timeline['Bulan_Tahun'] = timeline["Diterbitkan Tanggal"].dt.strftime('%b %Y')

                fig5 = px.line(
                    timeline,
                    x="Diterbitkan Tanggal",
                    y="Jumlah",
                    markers=True,
                    title="Tren Jumlah Ekspor",
                    hover_data={'Diterbitkan Tanggal': False}
                )

                fig5.update_traces(
                    line=dict(color="royalblue", width=2),
                    marker=dict(size=8, color="orange"),
                    hovertemplate="<b>Tanggal:</b> %{customdata[0]}<br><b>Jumlah:</b> %{y}<extra></extra>",
                    customdata=timeline[['Tanggal_Display']].values
                )

                fig5.update_layout(
                    xaxis=dict(
                        title="Tanggal", 
                        showgrid=True, 
                        gridcolor="lightgrey",
                        tickformat="%b %Y",
                        dtick="M1",  # Tick setiap bulan
                        tickangle=45  # Rotate label agar tidak overlap
                    ),
                    yaxis=dict(title="Jumlah", showgrid=True, gridcolor="lightgrey"),
                    plot_bgcolor="white",
                    hovermode="x unified",
                    width=800,
                    height=600,
                    margin=dict(b=100)  # Tambah margin bawah untuk rotated labels
                )

                monthly_ticks = timeline.groupby(timeline["Diterbitkan Tanggal"].dt.to_period("M")).first()

                fig5.update_xaxes(
                    tickvals=monthly_ticks["Diterbitkan Tanggal"],
                    ticktext=monthly_ticks["Bulan_Tahun"],
                    tickangle=45
                )
                
                figures_dict["5. Tren Jumlah Ekspor"] = fig5

                # ========================
                # 6. MAP NEGARA TUJUAN
                # ========================
                country_counts = df_filtered['Negara Tujuan'].value_counts().reset_index()
                country_counts.columns = ["Negara", "Jumlah"]
                country_counts["ISO3"] = country_counts["Negara"].apply(get_iso3)

                top_labels = country_counts.head(10)

                fig6 = px.choropleth(
                    country_counts,
                    locations="Negara",
                    locationmode="country names",
                    color="Jumlah",
                    hover_name="Negara",
                    color_continuous_scale=["#e69795", "#bc656d", "#b03031"],
                    title="Sebaran Negara Tujuan"
                )

                for _, row in top_labels.iterrows():
                    fig6.add_trace(go.Scattergeo(
                        locationmode="country names",
                        locations=[row["Negara"]],
                        text=row["ISO3"],
                        mode="text",
                        showlegend=False,
                        textfont=dict(size=9, color="black"),
                        hoverinfo="skip"
                    ))

                fig6.update_layout(
                    geo=dict(showframe=False, showcoastlines=True, projection_type='natural earth'),
                    width=800,
                    height=600
                )
                figures_dict["6. Peta Sebaran Negara Tujuan"] = fig6

                # Generate PDF
                pdf_buffer = create_pdf_report(figures_dict, date_range, total_records, filtered_records)
                
                # Nama file PDF
                filename = f"Dashboard_Ekspor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                
                # Simpan ke session state untuk download
                st.session_state.pdf_buffer = pdf_buffer
                st.session_state.pdf_filename = filename
                
                st.sidebar.success("✅ PDF Report berhasil dibuat!")
                
            except Exception as e:
                st.sidebar.error(f"❌ Error saat membuat PDF: {str(e)}")
    
    # Tombol Export Excel
    if st.sidebar.button("📊 Generate Excel Report", type="primary"):
        with st.spinner("Sedang membuat laporan Excel..."):
            try:
                # Date range string
                if start_date == min_date and end_date == max_date:
                    date_range = f"{min_date.strftime('%d-%m-%Y')} - {max_date.strftime('%d-%m-%Y')} (Semua Data)"
                else:
                    date_range = f"{start_date.strftime('%d-%m-%Y')} - {end_date.strftime('%d-%m-%Y')}"
                
                # Generate Excel
                excel_buffer = create_excel_report(df_filtered, date_range, total_records, filtered_records)
                
                # Nama file Excel
                excel_filename = f"Data_Ekspor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                
                # Simpan ke session state untuk download
                st.session_state.excel_buffer = excel_buffer
                st.session_state.excel_filename = excel_filename
                
                st.sidebar.success("✅ Excel Report berhasil dibuat!")
                
            except Exception as e:
                st.sidebar.error(f"❌ Error saat membuat Excel: {str(e)}")

    # Tombol download jika PDF sudah dibuat
    if 'pdf_buffer' in st.session_state:
        st.sidebar.download_button(
            label="📥 Download PDF Report",
            data=st.session_state.pdf_buffer.getvalue(),
            file_name=st.session_state.pdf_filename,
            mime="application/pdf",
            type="secondary"
        )

    # Tombol download Excel jika sudah dibuat (tambahkan setelah tombol download PDF)
    if 'excel_buffer' in st.session_state:
        st.sidebar.download_button(
            label="📥 Download Excel Report",
            data=st.session_state.excel_buffer.getvalue(),
            file_name=st.session_state.excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="secondary"
        )

    # ========================
    # REGENERATE FIGURES FOR DISPLAY
    # ========================
    
    # ========================
    # 1. PIE CHART KOMODITAS
    # ========================
    comodity = df_filtered['Jenis Komoditi'].value_counts().reset_index()
    comodity.columns = ["Jenis Komoditi", "Jumlah"]

    total = comodity["Jumlah"].sum()
    comodity["Persentase"] = comodity["Jumlah"] / total * 100
    comodity["Label"] = comodity.apply(
        lambda x: f"{x['Jenis Komoditi']} ({x['Jumlah']} unit, {x['Persentase']:.2f}%)", axis=1
    )

    n_categories = len(comodity)
    base_size = 390
    additional_size = n_categories * 15  # Tambahan size berdasarkan jumlah kategori

    fig1 = go.Figure(
        data=[
            go.Pie(
                labels=comodity["Jenis Komoditi"],
                values=comodity["Jumlah"],
                hole=0.35,
                textinfo="label+percent+value",
                texttemplate="<b>%{label}</b><br>%{value} unit<br>(%{percent})",
                hovertemplate="<b>%{label}</b><br>Jumlah: %{value}<br>Persen: %{percent}<extra></extra>",
                marker=dict(
                    colors=px.colors.qualitative.Plotly, 
                    line=dict(color="white", width=2)
                ),
                textposition="outside",
                textfont=dict(size=12, family="Arial", color="black"),
                insidetextorientation='radial',
                outsidetextfont=dict(size=11),
                pull=0.03,
            )
        ]
    )

    fig1.update_layout(
        showlegend=False,
        annotations=[],
        uniformtext_minsize=10,
        uniformtext_mode='hide',
        
        margin=dict(l=5, r=5, t=90, b=35),  # Margin sangat ketat
        width=base_size + additional_size,  # Ukuran dinamis
        height=base_size + additional_size,
        autosize=False,
        
        # Hilangkan padding dan spacing yang tidak perlu
        paper_bgcolor='white',
        plot_bgcolor='white',
        
        # Hilangkan axis dan grid yang tidak diperlukan
        xaxis=dict(visible=False, showgrid=False, zeroline=False),
        yaxis=dict(visible=False, showgrid=False, zeroline=False),
    )

    fig1.update_traces(
        marker_line_color='white',
        marker_line_width=2,
        textfont_size=11,
        textposition='outside',
        textfont_color='black',
        insidetextfont=dict(size=11, color='white', family='Arial'),
        outsidetextfont=dict(size=10, family='Arial'),
        pull=[0.15 if i == 0 else 0.05 for i in range(len(comodity))],
        rotation=30,
        hole=0.4
    )
    
    # Atur layout agar lebih compact
    fig1.update_layout(
        autosize=True,  # Auto-size untuk menyesuaikan konten
        title={
        'text': "Persentase Jenis Komoditi",
        'y': 0.97,
        'x': 0.0,
        'xanchor': 'left',
        'yanchor': 'top'
    },
    title_font=dict(size=20, color="black", family="Arial Black"),
    title_font_color="black"
    )

    # ========================
    # 2. BAR CHART NEGARA
    # ========================
    top_10_country = df_filtered['Negara Tujuan'].value_counts().head(10).reset_index()
    top_10_country.columns = ["Negara", "Jumlah"]

    fig2 = px.bar(
        top_10_country,
        x="Negara",
        y="Jumlah",
        text="Jumlah",
        color="Jumlah",
        color_continuous_scale=["#e69795", "#bc656d", "#b03031"],
        title="Top 10 Negara Tujuan"
    )

    fig2.update_traces(texttemplate='%{text}', textposition='outside')

    max_value = max(top_10_country["Jumlah"])
    y_range_max = max_value * 1.2
    
    fig2.update_layout(xaxis_title="Negara", yaxis_title="Jumlah", yaxis=dict(range=[0, y_range_max]))

    # ========================
    # 3. MAP KOTA PERUSAHAAN EKSPOR
    # ========================
    city_counts = df_filtered['Kota'].value_counts().reset_index()
    city_counts.columns = ['Kota', 'Jumlah']
    
    df_map = pd.merge(df_city, city_counts, on='Kota', how='inner')
    
    if not df_map.empty:
        df_map["Kota_Label"] = df_map["Kota"] + " (" + df_map["Jumlah"].astype(str) + ")"
        
        fig3 = px.scatter_mapbox(
            df_map,
            lat="lat",
            lon="lon",
            hover_name="Kota",
            hover_data={"Jumlah": True, "lat": False, "lon": False},
            color="Kota_Label",   
            zoom=4,
            title="Sebaran Kota Perusahaan Eksportir"
        )
        
        fig3.update_traces(marker=dict(size=12, opacity=0.8))
        
        fig3.update_layout(
            mapbox_style="carto-positron",
            legend_title="Kota (Jumlah)",
            margin={"r":0,"t":50,"l":0,"b":0},
            title=dict(
                text="Sebaran Kota Perusahaan Eksportir",
                x=0.0,  # pojok kiri
                xanchor="left"
            )
        )
    else:
        # Jika tidak ada data kota yang cocok
        fig3 = go.Figure()
        fig3.add_annotation(text="Tidak ada data kota yang tersedia", 
                          xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        fig3.update_layout(title="Sebaran Kota Perusahaan Eksportir")

    # ========================
    # 4. BAR CHART PERUSAHAAN EXPORTIR
    # ========================
    top_10_company = df_filtered['Nama Exportir/Importir'].value_counts().head(10).reset_index()
    top_10_company.columns = ["Perusahaan", "Jumlah"]
    
    fig4 = px.bar(
        top_10_company,
        x="Perusahaan",
        y="Jumlah",
        text="Jumlah", 
        color="Jumlah",
        color_continuous_scale="plasma",
        title="Top 10 Perusahaan Ekspor Paling Banyak"
    )
    
    fig4.update_traces(texttemplate='%{text}', textposition='outside')

    max_value = max(top_10_company["Jumlah"])
    y_range_max = max_value * 1.2
    
    fig4.update_layout(
        xaxis_title="Perusahaan",
        yaxis_title="Jumlah",
        yaxis=dict(range=[0, y_range_max])
    )

    # ========================
    # 5. LINE CHART TIMELINE
    # ========================
    timeline = df_filtered.groupby("Diterbitkan Tanggal").size().reset_index(name="Jumlah")

    fig5 = px.line(
        timeline,
        x="Diterbitkan Tanggal",
        y="Jumlah",
        markers=True,
        title="Tren Jumlah Ekspor"
    )

    fig5.update_traces(
        line=dict(color="royalblue", width=2),
        marker=dict(size=8, color="orange"),
        hovertemplate="<b>Tanggal:</b> %{x|%d %B %Y}<br><b>Jumlah:</b> %{y}<extra></extra>"
    )

    fig5.update_layout(
        xaxis=dict(title="Tanggal", showgrid=True, gridcolor="lightgrey", tickformat="%b %Y"),
        yaxis=dict(title="Jumlah", showgrid=True, gridcolor="lightgrey"),
        plot_bgcolor="white",
        hovermode="x unified"
    )

    # ========================
    # 6. MAP NEGARA TUJUAN
    # ========================
    country_counts = df_filtered['Negara Tujuan'].value_counts().reset_index()
    country_counts.columns = ["Negara", "Jumlah"]
    country_counts["ISO3"] = country_counts["Negara"].apply(get_iso3)

    top_labels = country_counts.head(10)

    fig6 = px.choropleth(
        country_counts,
        locations="Negara",
        locationmode="country names",
        color="Jumlah",
        hover_name="Negara",
        color_continuous_scale=["#e69795", "#bc656d", "#b03031"],
        title="Sebaran Negara Tujuan"
    )

    for _, row in top_labels.iterrows():
        fig6.add_trace(go.Scattergeo(
            locationmode="country names",
            locations=[row["Negara"]],
            text=row["ISO3"],
            mode="text",
            showlegend=False,
            textfont=dict(size=9, color="black"),
            hoverinfo="skip"
        ))

    fig6.update_layout(geo=dict(showframe=False, showcoastlines=True, projection_type='natural earth'))

    # ========================
    # TAMPILKAN DI STREAMLIT
    # ========================
    st.plotly_chart(fig1, use_container_width=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(fig2, use_container_width=True)
    with col2:
        st.plotly_chart(fig4, use_container_width=True)

    st.plotly_chart(fig3, use_container_width=True)
    st.plotly_chart(fig6, use_container_width=True)
    st.plotly_chart(fig5, use_container_width=True)
  
else:
    st.info("📥 Silakan upload file CSV/Excel untuk memulai.")