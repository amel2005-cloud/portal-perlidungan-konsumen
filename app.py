import streamlit as st
from fpdf import FPDF
st.markdown("""
    <style>
    /* Efek Grid Kertas */
    .stApp { 
        background-image: radial-gradient(#d1d1d1 0.5px, transparent 0.5px);
        background-size: 20px 20px;
        background-color: #f8f9fa;
    }
    
    /* Header/Judul */
    h1 { color: #b22222; text-align: center; border-bottom: 3px solid #DAA520; padding-bottom: 10px; }
    
    /* Box Formulir */
    .stTextInput, .stTextArea, .stSelectbox { border-left: 5px solid #b22222; }
    
    /* Tombol */
    div.stButton > button:first-child { 
        background-color: #b22222; color: white; border: none; font-weight: bold; width: 100%;
    }
    </style>
""", unsafe_allow_html=True)


# --- HEADER PORTAL ---
col1, col2, col3 = st.columns([2, 1, 2])
with col2:
    st.image("image.png", width=130)
st.title("PORTAL PERLINDUNGAN KONSUMEN")
st.markdown("---")
st.subheader("Formulir Layanan Surat Resmi OJK JEMBER")

# --- FUNGSI TATA LETAK ---
def cetak_baris(pdf, label, nilai):
    pdf.cell(40, 8, txt=label, ln=0)
    pdf.cell(5, 8, txt=":", ln=0)
    pdf.cell(0, 8, txt=nilai, ln=1)

# --- INPUT DATA ---
kronologis, no_hp, email, isi_pernyataan = "", "", "", ""
tipe_surat = st.selectbox("Pilih Jenis Layanan:", ["Surat Pengaduan", "Surat Pernyataan"])
nama = st.text_input("Nama Lengkap")
nik = st.text_input("NIK / Nomor Identitas")
alamat = st.text_area("Alamat")

if tipe_surat == "Surat Pengaduan":
    no_hp = st.text_input("No. HP")
    email = st.text_input("Email")
    kronologis = st.text_area("Tuliskan kronologis permasalahan:")
elif tipe_surat == "Surat Pernyataan":
    isi_pernyataan = st.text_area("Tuliskan permasalahan yang diajukan:")

kota_ttd = st.text_input("Kota", value="Jember")
tanggal_ttd = st.date_input("Tanggal")

# --- GENERATOR PDF ---
if st.button("PROSES & CETAK SURAT"):
    pdf = FPDF()
    pdf.add_page()
    
    # Elemen Header PDF
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, txt=tipe_surat.upper(), ln=True, align='C')
    pdf.set_fill_color(178, 34, 34) # Warna Merah OJK
    pdf.cell(0, 1, "", ln=True, fill=True)
    pdf.ln(10)
    
    pdf.set_font("Arial", size=12)
    
    if tipe_surat == "Surat Pengaduan":
        pdf.cell(0, 8, txt="Saya yang bertanda tangan di bawah ini :", ln=True)
        cetak_baris(pdf, "Nama", nama)
        cetak_baris(pdf, "NIK", nik)
        cetak_baris(pdf, "Alamat", alamat)
        cetak_baris(pdf, "No. HP", no_hp)
        cetak_baris(pdf, "Email", email)
        pdf.ln(5)
        pdf.cell(0, 8, txt="Kronologis permasalahan:", ln=True)
        pdf.multi_cell(0, 8, txt=kronologis)
    else:
        pdf.multi_cell(0, 8, txt="Sehubungan dengan pengajuan penyelesaian pengaduan melalui Aplikasi Portal Perlindungan Konsumen (APPK), dengan ini saya:")
        pdf.ln(2)
        cetak_baris(pdf, "Nama", nama)
        cetak_baris(pdf, "Nomor Identitas", nik)
        cetak_baris(pdf, "Alamat", alamat)
        pdf.ln(5)
        pdf.multi_cell(0, 8, txt=f"{isi_pernyataan}")

    # Tanda Tangan
    pdf.ln(20)
    pdf.cell(110)
    pdf.cell(0, 8, txt=f"{kota_ttd}, {tanggal_ttd.strftime('%d %B %Y')}", ln=True)
    pdf.ln(20)
    pdf.cell(110)
    pdf.cell(0, 8, txt=f"({nama})", ln=True)
    
    pdf.output("surat_hasil.pdf")
    with open("surat_hasil.pdf", "rb") as f:
        st.download_button("DOWNLOAD DOKUMEN", f, file_name="surat_hasil.pdf")

        st.markdown("---")
st.caption("© 2026 Otoritas Jasa Keuangan & Pemerintah Kabupaten Jember | Layanan Terintegrasi")
