import streamlit as st
from fpdf import FPDF
from docx import Document
from docx.shared import Pt, RGBColor
from io import BytesIO

st.markdown("""
    <style>
    .stApp { 
        background-image: radial-gradient(#d1d1d1 0.5px, transparent 0.5px);
        background-size: 20px 20px;
        background-color: #f8f9fa;
    }
    h1 { color: #b22222; text-align: center; border-bottom: 3px solid #DAA520; padding-bottom: 10px; }
    .stTextInput, .stTextArea, .stSelectbox { border-left: 5px solid #b22222; }
    div.stButton > button:first-child { 
        background-color: #b22222; color: white; border: none; font-weight: bold; width: 100%;
    }
    </style>
""", unsafe_allow_html=True)

col1, col2, col3 = st.columns([2, 1, 2])
with col2:
    st.image("image.png", width=130)
st.title("PORTAL PERLINDUNGAN KONSUMEN")
st.subheader("Formulir Layanan Surat Resmi OJK JEMBER")

def cetak_baris(pdf, label, nilai):
    pdf.cell(40, 8, txt=label, ln=0)
    pdf.cell(5, 8, txt=":", ln=0)
    pdf.cell(0, 8, txt=nilai, ln=1)

def bersihkan_teks(teks):
    if not teks:
        return ""
    return teks.encode('latin-1', 'replace').decode('latin-1')

# --- INPUT DATA ---
kronologis, no_hp, email, isi_pernyataan = "", "", "", ""
tipe_surat = st.selectbox("Pilih Jenis Layanan:", ["Surat Pengaduan", "Surat Pernyataan"])
nama = st.text_input("Nama Lengkap")
nik = st.text_input("NIK / Nomor Identitas")
if nik and (len(nik) != 16 or not nik.isdigit()):
    st.warning("NIK harus 16 digit angka.")
alamat = st.text_area("Alamat")

if tipe_surat == "Surat Pengaduan":
    no_hp = st.text_input("No. HP")
    email = st.text_input("Email")
elif tipe_surat == "Surat Pernyataan":
    pt_dituju = st.text_input("Nama PT yang Dituju")
   

kota_ttd = st.text_input("Kota", value="Jember")
tanggal_ttd = st.date_input("Tanggal")

TEKS_PERNYATAAN = (
    "menyatakan dengan sesungguhnya bahwa permasalahan yang saya ajukan melalui "
    "Aplikasi Portal Perlindungan Konsumen (APPK) tidak sedang dalam proses atau "
    "pernah diputus oleh lembaga arbitrase atau peradilan atau lembaga mediasi lainnya "
    "termasuk lembaga alternatif penyelesaian Sengketa dan belum pernah difasilitasi oleh OJK."
)
TEKS_PENUTUP = "Demikian surat pernyataan ini dibuat dengan sadar dan tanpa paksaan dari pihak manapun."

# --- GENERATOR PDF ---
def buat_pdf(tipe_surat):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_margins(20, 20, 20)

    if tipe_surat == "Surat Pengaduan":
        # Kepada Yth
        pdf.set_font("Arial", size=12)
        pdf.cell(0, 8, txt="Kepada Yth.", ln=True)
        pdf.cell(0, 8, txt=bersihkan_teks(pt_dituju), ln=True)
        pdf.cell(0, 8, txt="di Tempat", ln=True)
        pdf.ln(8)

        # Judul
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, txt="SURAT PERNYATAAN", ln=True, align='C')
        pdf.ln(8)

        # Pembuka
        pdf.set_font("Arial", size=12)
        pdf.cell(0, 8, txt="Yang bertanda tangan di bawah ini:", ln=True)
        pdf.ln(2)
        cetak_baris(pdf, "Nama", bersihkan_teks(nama))
        cetak_baris(pdf, "NIK", bersihkan_teks(nik))
        cetak_baris(pdf, "Alamat", bersihkan_teks(alamat))
        pdf.ln(6)

        # Isi pernyataan
        pdf.multi_cell(0, 8, txt=bersihkan_teks(TEKS_PERNYATAAN))
        pdf.ln(4)

        # Permasalahan tambahan
        if isi_pernyataan:
            pdf.multi_cell(0, 8, txt=bersihkan_teks(isi_pernyataan))
            pdf.ln(4)

        # Penutup (paragraf baru)
        pdf.multi_cell(0, 8, txt=bersihkan_teks(TEKS_PENUTUP))

    else:
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, txt="SURAT PENGADUAN", ln=True, align='C')
        pdf.set_fill_color(178, 34, 34)
        pdf.cell(0, 1, "", ln=True, fill=True)
        pdf.ln(10)
        pdf.set_font("Arial", size=12)
        pdf.cell(0, 8, txt="Saya yang bertanda tangan di bawah ini:", ln=True)
        cetak_baris(pdf, "Nama", bersihkan_teks(nama))
        cetak_baris(pdf, "NIK", bersihkan_teks(nik))
        cetak_baris(pdf, "Alamat", bersihkan_teks(alamat))
        cetak_baris(pdf, "No. HP", bersihkan_teks(no_hp))
        cetak_baris(pdf, "Email", bersihkan_teks(email))
        pdf.ln(5)
        pdf.cell(0, 8, txt="Kronologis permasalahan:", ln=True)
        pdf.multi_cell(0, 8, txt=bersihkan_teks(kronologis))

    # Tanda tangan
    pdf.ln(20)
    pdf.cell(110)
    pdf.cell(0, 8, txt=bersihkan_teks(f"{kota_ttd}, {tanggal_ttd.strftime('%d %B %Y')}"), ln=True)
    pdf.ln(20)
    pdf.cell(110)
    pdf.cell(0, 8, txt=bersihkan_teks(f"({nama})"), ln=True)

    pdf.output("surat_hasil.pdf")
    with open("surat_hasil.pdf", "rb") as f:
        return f.read()

# --- GENERATOR WORD ---
def buat_word(tipe_surat):
    doc = Document()

    if tipe_surat == "Surat Pernyataan":
        # Kepada Yth
        doc.add_paragraph("Kepada Yth.")
        doc.add_paragraph(pt_dituju)
        doc.add_paragraph("di Tempat")
        doc.add_paragraph("")

        # Judul
        judul = doc.add_paragraph("SURAT PERNYATAAN")
        judul.alignment = 1  # center
        judul.runs[0].bold = True
        judul.runs[0].font.size = Pt(14)
        doc.add_paragraph("")

        # Pembuka
        doc.add_paragraph("Yang bertanda tangan di bawah ini:")
        doc.add_paragraph(f"Nama               : {nama}")
        doc.add_paragraph(f"NIK                    : {nik}")
        doc.add_paragraph(f"Alamat             : {alamat}")
        doc.add_paragraph("")

        # Isi pernyataan
        doc.add_paragraph(TEKS_PERNYATAAN)
        if isi_pernyataan:
            doc.add_paragraph(isi_pernyataan)

        # Penutup paragraf baru
        doc.add_paragraph(TEKS_PENUTUP)

    else:
        doc.add_heading("SURAT PENGADUAN", level=1)
        doc.add_paragraph("Saya yang bertanda tangan di bawah ini:")
        doc.add_paragraph(f"Nama       : {nama}")
        doc.add_paragraph(f"NIK            : {nik}")
        doc.add_paragraph(f"Alamat     : {alamat}")
        doc.add_paragraph(f"No. HP    : {no_hp}")
        doc.add_paragraph(f"Email        : {email}")
        doc.add_paragraph("")
        doc.add_paragraph("Kronologis permasalahan:")
        doc.add_paragraph(kronologis)

    # Tanda tangan
    doc.add_paragraph("")
    doc.add_paragraph("")
    ttd = doc.add_paragraph(f"                                                  {kota_ttd}, {tanggal_ttd.strftime('%d %B %Y')}")
    doc.add_paragraph("")
    doc.add_paragraph("")
    doc.add_paragraph(f"                                                  ({nama})")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- TOMBOL PROSES ---
if st.button("PROSES & CETAK SURAT"):
    if len(nik) != 16 or not nik.isdigit():
        st.error("⚠️ NIK harus berupa 16 digit angka!")
    else:
        pdf_bytes = buat_pdf(tipe_surat)
        word_buffer = buat_word(tipe_surat)

        col_pdf, col_word = st.columns(2)
        with col_pdf:
            st.download_button("📄 Download PDF", pdf_bytes, file_name="surat_hasil.pdf")
        with col_word:
            st.download_button("📝 Download Word", word_buffer, file_name="surat_hasil.docx")

st.caption("© 2026 Otoritas Jasa Keuangan & Pemerintah Kabupaten Jember | Layanan Terintegrasi")
