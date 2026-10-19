import streamlit as st
import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import io
import re

# Konfigurasi Halaman Utama
st.set_page_config(page_title="Asisten e-Kinerja & Bukti Dukung", layout="wide")
st.title("🚀 Aplikasi Pengisi e-Kinerja & Generator Bukti Dukung")
st.write("Masukkan rekap data harian Anda di bawah ini untuk otomatis menghasilkan laporan bulanan.")

# Inisialisasi Database Sementara di Session State
if 'data_kinerja' not in st.session_state:
    st.session_state.data_kinerja = pd.DataFrame(columns=[
        'Tanggal', 'Jenis Aktivitas', 'Jumlah (Volume)', 'Detail Keterangan/Instansi'
    ])

# Buat Tab agar tampilan aplikasi rapi
tab1, tab2 = st.tabs(["🤖 Tempel dari Claude (Otomatis)", "📝 Isi Formulir (Manual)"])

# ==================== TAB 1: INPUT OTOMATIS DARI CLAUDE ====================
with tab1:
    st.header("🤖 Tempel Teks Rekomendasi Claude")
    st.write("Salin draf laporan harian dari Claude/Antigravity dan tempel di bawah ini.")
    
    tgl_auto = st.date_input("Tanggal Kegiatan (Otomatis)", datetime.now(), key="tgl_auto")
    teks_ai = st.text_area(
        "Kotak Teks Claude:", 
        height=150, 
        placeholder="Contoh:\nVerifikasi Akun Non-Penyedia (PPK) | 3 Akun | Dinas Kesehatan\nVerifikasi Akun Penyedia | 2 Akun | PT Abadi"
    )
    
    if st.button("⚡ Ekstrak Teks Menjadi Baris Tabel"):
        if teks_ai:
            baris_data = teks_ai.strip().split('\n')
            berhasil_tambah = 0
            
            for baris in baris_data:
                # Menghilangkan karakter pembatas tabel markdown jika pengguna menyalin mentah-mentah
                if '|' in baris and '---' not in baris and 'Jenis Aktivitas' not in baris:
                    parts = [p.strip() for p in baris.split('|') if p.strip()]
                    if len(parts) >= 2:
                        act = parts[0]
                        
                        # Ekstrak angka volume dari teks (misal: "3 Akun" -> 3)
                        vol_match = re.search(r'\d+', parts[1])
                        vol = int(vol_match.group()) if vol_match else 1
                        
                        det = parts[2] if len(parts) > 2 else "-"
                        
                        new_row = pd.DataFrame([{
                            'Tanggal': tgl_auto.strftime('%Y-%m-%d'),
                            'Jenis Aktivitas': act,
                            'Jumlah (Volume)': vol,
                            'Detail Keterangan/Instansi': det
                        }])
                        st.session_state.data_kinerja = pd.concat([st.session_state.data_kinerja, new_row], ignore_index=True)
                        berhasil_tambah += 1
            
            if berhasil_tambah > 0:
                st.success(f"Berhasil mengekstrak {berhasil_tambah} aktivitas ke dalam tabel!")
                st.rerun()
            else:
                st.warning("Format teks tidak dikenali. Pastikan teks menggunakan pembatas garis tegak ( | ) seperti format tabel dari Claude.")
        else:
            st.error("Kotak teks masih kosong!")

# ==================== TAB 2: FORM INPUT MANUAL ====================
with tab2:
    st.header("📝 Input Aktivitas Manual")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        tgl = st.date_input("Tanggal Kegiatan", datetime.now(), key="tgl_manual")
    with col2:
        jenis_act = st.selectbox("Jenis Aktivitas (Sesuai RHK)", [
            "Verifikasi Akun Non-Penyedia (PP, PPK, dll)",
            "Verifikasi Akun Penyedia SPSE",
            "Verifikasi Akun RUP PA/KPA",
            "Layanan Helpdesk & Penanganan Kendala",
            "Tugas Tambahan (Direktif Pimpinan)"
        ])
    with col3:
        volume = st.number_input("Jumlah (Volume)", min_value=1, value=1, step=1, key="vol_manual")
    with col4:
        detail = st.text_input("Detail / Keterangan", placeholder="Misal: Dinas X, PT Y", key="det_manual")

    if st.button("➕ Tambah Manual ke Tabel"):
        new_data = pd.DataFrame([{
            'Tanggal': tgl.strftime('%Y-%m-%d'),
            'Jenis Aktivitas': jenis_act,
            'Jumlah (Volume)': int(volume),
            'Detail Keterangan/Instansi': detail
        }])
        st.session_state.data_kinerja = pd.concat([st.session_state.data_kinerja, new_data], ignore_index=True)
        st.success("Data manual berhasil ditambahkan!")
        st.rerun()

st.markdown("---")

# ==================== TABEL REKAP & DOWNLOAD ====================
st.header("📊 Tabel Rekapitulasi Kinerja Bulanan")

if not st.session_state.data_kinerja.empty:
    st.dataframe(st.session_state.data_kinerja, use_container_width=True)
    
    if st.button("⚠️ Reset Seluruh Data Tabel"):
        st.session_state.data_kinerja = pd.DataFrame(columns=['Tanggal', 'Jenis Aktivitas', 'Jumlah (Volume)', 'Detail Keterangan/Instansi'])
        st.rerun()

    st.markdown("### 📥 Unduh Dokumen Bukti Dukung")
    btn_col1, btn_col2 = st.columns(2)
    
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
        
    with btn_col2:
        def generate_pdf(df):
            pdf_buffer = io.BytesIO()
            doc = SimpleDocTemplate(pdf_buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
            story = []
            
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=14, leading=16, alignment=1, spaceAfter=20)
            text_style = ParagraphStyle('TextStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=12)
            
            story.append(Paragraph("LAPORAN BULANAN PELAYANAN DAN LOG DATA SPSE", title_style))
            story.append(Paragraph(f"Periode Laporan: {datetime.now().strftime('%B %Y')}", text_style))
            story.append(Spacer(1, 15))
            
            table_data = [[Paragraph(f"<b>{col}</b>", text_style) for col in df.columns]]
            for index, row in df.iterrows():
                table_data.append([Paragraph(str(val), text_style) for val in row])
            
            col_widths = [70, 220, 70, 180]
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
    st.info("Belum ada data aktivitas yang dimasukkan. Silakan tempel teks dari Claude atau isi form manual.")
