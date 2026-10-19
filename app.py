import streamlit as st
import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import io

# Konfigurasi Halaman Utama
st.set_page_config(page_title="Asisten e-Kinerja & Bukti Dukung", layout="wide")
st.title("🚀 Aplikasi Pengisi e-Kinerja & Generator Bukti Dukung")
st.write("Masukkan rekap data harian Anda di bawah ini untuk otomatis menghasilkan laporan bulanan.")

# Inisialisasi Database Sementara di Session State
if 'data_kinerja' not in st.session_state:
    st.session_state.data_kinerja = pd.DataFrame(columns=[
        'Tanggal', 'Jenis Aktivitas', 'Jumlah (Volume)', 'Detail Keterangan/Instansi'
    ])

# ==================== FORM INPUT HARIAN ====================
st.header("📝 Input Aktivitas Harian")
col1, col2, col3, col4 = st.columns(4)

with col1:
    tgl = st.date_input("Tanggal Kegiatan", datetime.now())
with col2:
    jenis_act = st.selectbox("Jenis Aktivitas (Sesuai RHK)", [
        "Verifikasi Akun Non-Penyedia (PP, PPK, dll)",
        "Verifikasi Akun Penyedia SPSE",
        "Verifikasi Akun RUP PA/KPA",
        "Layanan Helpdesk & Penanganan Kendala",
        "Tugas Tambahan (Direktif Pimpinan)"
    ])
with col3:
    volume = st.number_input("Jumlah (Volume)", min_value=1, value=1, step=1)
with col4:
    detail = st.text_input("Detail / Keterangan", placeholder="Misal: Dinas X, PT Y, kendala login")

if st.button("➕ Tambah ke Rekapitulasi"):
    new_data = pd.DataFrame([{
        'Tanggal': tgl.strftime('%Y-%m-%d'),
        'Jenis Aktivitas': jenis_act,
        'Jumlah (Volume)': int(volume),
        'Detail Keterangan/Instansi': detail
    }])
    st.session_state.data_kinerja = pd.concat([st.session_state.data_kinerja, new_data], ignore_index=True)
    st.success("Data berhasil ditambahkan ke tabel rekap harian!")

st.markdown("---")

# ==================== TABEL REKAP & DOWNLOAD ====================
st.header("📊 Tabel Rekapitulasi Kinerja Bulanan")

if not st.session_state.data_kinerja.empty:
    # Menampilkan tabel di Dashboard
    st.dataframe(st.session_state.data_kinerja, use_container_width=True)
    
    # Tombol Reset Data
    if st.button("⚠️ Reset Seluruh Data Tabel"):
        st.session_state.data_kinerja = pd.DataFrame(columns=['Tanggal', 'Jenis Aktivitas', 'Jumlah (Volume)', 'Detail Keterangan/Instansi'])
        st.rerun()

    st.markdown("### 📥 Unduh Dokumen Bukti Dukung")
    btn_col1, btn_col2 = st.columns(2)
    
    # 1. GENERATOR EXCEL (.xlsx)
    with btn_col1:
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
            st.session_state.data_kinerja.to_excel(writer, sheet_name='Rekap_Bulanan', index=False)
        st.download_button(
            label="📊 Unduh Format Excel (.xlsx)",
            data=buffer_excel.getvalue(),
            file_name=f"Rekap_Bukti_Dukung_{datetime.now().strftime('%B_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    # 2. GENERATOR PDF (.pdf) menggunakan ReportLab
    with btn_col2:
        def generate_pdf(df):
            pdf_buffer = io.BytesIO()
            doc = SimpleDocTemplate(pdf_buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
            story = []
            
            # Pengaturan Gaya Teks
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=14, leading=16, alignment=1, spaceAfter=20)
            text_style = ParagraphStyle('TextStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=12)
            
            # Judul Dokumen PDF
            story.append(Paragraph("LAPORAN BULANAN PELAYANAN DAN LOG DATA SPSE", title_style))
            story.append(Paragraph(f"Periode Laporan: {datetime.now().strftime('%B %Y')}", text_style))
            story.append(Spacer(1, 15))
            
            # Menyiapkan Data Tabel untuk PDF
            table_data = [[Paragraph(f"<b>{col}</b>", text_style) for col in df.columns]]
            for index, row in df.iterrows():
                table_data.append([Paragraph(str(val), text_style) for val in row])
            
            # Membuat Objek Tabel ReportLab
            col_widths = [70, 200, 60, 210]  # Total lebar agar muat di kertas letter
            t = Table(table_data, colWidths=col_widths, repeatRows=1)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('TOPPADDING', (0,0), (-1,-1), 6),
                ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ]))
            story.append(t)
            
            # Render PDF
            doc.build(story)
            pdf_buffer.seek(0)
            return pdf_buffer.getvalue()

        pdf_data = generate_pdf(st.session_state.data_kinerja)
        st.download_button(
            label="📄 Unduh Bukti Dukung Cetak (.pdf)",
            data=pdf_data,
            file_name=f"Laporan_Bukti_Dukung_{datetime.now().strftime('%B_%Y')}.pdf",
            mime="application/pdf"
        )
else:
    st.info("Belum ada data aktivitas yang dimasukkan. Silakan isi form di atas terlebih dahulu.")
