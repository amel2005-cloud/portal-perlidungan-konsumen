import streamlit as st
from fpdf import FPDF
from docx import Document
from docx.shared import Pt
from io import BytesIO

# ===================== INISIALISASI SESSION STATE =====================
if "halaman" not in st.session_state:
    st.session_state.halaman = "pembukaan"

# ===================== CSS GLOBAL =====================
st.markdown("""
<style>
.stApp {
    background-image: radial-gradient(#d1d1d1 0.5px, transparent 0.5px);
    background-size: 20px 20px;
    background-color: #f8f9fa;
}
.stTextInput, .stTextArea, .stSelectbox { border-left: 5px solid #b22222; }
div.stButton > button:first-child {
    background-color: #b22222; color: white; border: none;
    font-weight: bold; width: 100%;
}
</style>
""", unsafe_allow_html=True)

# ===================== HALAMAN PEMBUKAAN =====================
if st.session_state.halaman == "pembukaan":
    st.markdown("""
    <style>
    .stApp { background: #8B0000 !important; background-image: none !important; }

    .splash-wrap {
        min-height: 88vh;
        display: flex; flex-direction: column;
        align-items: center; justify-content: center;
        text-align: center;
        padding: 40px 24px;
        position: relative;
        font-family: Arial, sans-serif;
    }
    .splash-wrap::before {
        content:'';
        position:fixed; inset:0;
        background: repeating-linear-gradient(
            45deg, transparent, transparent 30px,
            rgba(255,255,255,0.02) 30px, rgba(255,255,255,0.02) 60px
        );
        pointer-events:none;
    }
    .gold-top {
        position:fixed; top:0; left:0; right:0; height:5px;
        background: linear-gradient(90deg, #b8860b, #FFD700, #b8860b);
        z-index:999;
    }
    .gold-bot {
        position:fixed; bottom:0; left:0; right:0; height:5px;
        background: linear-gradient(90deg, #b8860b, #FFD700, #b8860b);
        z-index:999;
    }

    .logo-ring {
        width:110px; height:110px; border-radius:50%;
        background:white; border:3px solid #FFD700;
        display:flex; align-items:center; justify-content:center;
        margin:0 auto 28px;
        animation: popIn 0.6s ease 0.2s both;
    }
    .logo-ring span {
        font-size:16px; font-weight:900; color:#8B0000;
        line-height:1.25; letter-spacing:1px;
    }

    @keyframes popIn {
        from { transform:scale(0.5); opacity:0; }
        to   { transform:scale(1);   opacity:1; }
    }

    .inst-label {
        font-size:11px; letter-spacing:3px; text-transform:uppercase;
        color:rgba(255,255,255,0.5); margin-bottom:8px;
        animation: fadeUp 0.5s ease 0.7s both;
    }
    .splash-title {
        font-size:32px; font-weight:900; color:white;
        line-height:1.2; margin-bottom:6px;
        text-shadow:0 2px 12px rgba(0,0,0,0.4);
        animation: fadeUp 0.5s ease 0.9s both;
    }
    .splash-sub {
        font-size:14px; color:rgba(255,255,255,0.7);
        font-style:italic; margin-bottom:36px;
        animation: fadeUp 0.5s ease 1.1s both;
    }

    .typed-container {
        min-height: 44px;
        display:flex; align-items:center; justify-content:center;
    }
    .typed-text {
        font-size:32px; font-weight:900; color:white;
        text-shadow:0 2px 12px rgba(0,0,0,0.4);
        overflow:hidden; white-space:nowrap;
        border-right:3px solid #FFD700;
        width:0;
        animation:
            typing 1.8s steps(30,end) 1.2s forwards,
            blink 0.7s step-end infinite 1.2s;
    }
    @keyframes typing {
        from { width:0 }
        to   { width:100% }
    }
    @keyframes blink {
        0%,100%{ border-color:#FFD700 }
        50%    { border-color:transparent }
    }

    .sub-anim {
        font-size:14px; color:rgba(255,255,255,0.7);
        font-style:italic; margin-bottom:36px;
        opacity:0;
        animation: fadeUp 0.6s ease 3.2s forwards;
    }
    .badge-row {
        display:flex; gap:10px; flex-wrap:wrap; justify-content:center;
        margin-bottom:36px;
        opacity:0; animation: fadeUp 0.5s ease 3.5s forwards;
    }
    .badge {
        background:rgba(255,255,255,0.1);
        border:1px solid rgba(255,215,0,0.35);
        color:rgba(255,255,255,0.85);
        font-size:11.5px; padding:5px 14px; border-radius:20px;
        display:inline-flex; align-items:center; gap:6px;
    }
    .badge-dot { width:6px;height:6px;border-radius:50%;background:#FFD700;display:inline-block; }

    .appk-card {
        background:rgba(255,255,255,0.1);
        border:1px solid rgba(255,215,0,0.45);
        border-radius:12px; padding:20px 28px;
        max-width:420px; margin:0 auto 32px;
        opacity:0; animation: fadeUp 0.5s ease 3.8s forwards;
        font-family:Arial,sans-serif;
    }
    .appk-card-title {
        font-size:10px; letter-spacing:2px; text-transform:uppercase;
        color:rgba(255,255,255,0.5); margin-bottom:6px;
    }
    .appk-card-text {
        font-size:14px; color:rgba(255,255,255,0.85); margin-bottom:14px; line-height:1.5;
    }
    .appk-btn {
        display:inline-block;
        background:#FFD700; color:#5a0000;
        font-weight:800; font-size:13px;
        padding:10px 22px; border-radius:7px;
        text-decoration:none; letter-spacing:0.3px;
        transition:background 0.2s;
    }
    .appk-btn:hover { background:#ffe033; }

    @keyframes fadeUp {
        from { transform:translateY(12px); opacity:0; }
        to   { transform:translateY(0);    opacity:1; }
    }
    </style>

    <div class="gold-top"></div>
    <div class="gold-bot"></div>

    <div class="splash-wrap">
        <div class="logo-ring"><span>OJK<br>JEMBER</span></div>

        <div class="inst-label">Otoritas Jasa Keuangan &bull; Kantor Jember</div>

        <div class="typed-container">
            <span class="typed-text">Portal Perlindungan Konsumen</span>
        </div>

        <div class="sub-anim">Layanan Surat Resmi Terintegrasi &mdash; OJK Jember</div>

        <div class="badge-row">
            <span class="badge"><span class="badge-dot"></span>Resmi &amp; Terpercaya</span>
            <span class="badge"><span class="badge-dot"></span>Pengaduan Terverifikasi</span>
            <span class="badge"><span class="badge-dot"></span>Layanan Gratis</span>
        </div>

        <div class="appk-card">
            <div class="appk-card-title">Layanan Pengaduan Online OJK</div>
            <div class="appk-card-text">
                Akses langsung ke <strong style="color:#FFD700;">APPK (Aplikasi Portal Perlindungan Konsumen)</strong>
                untuk mengajukan pengaduan secara online ke OJK Pusat.
            </div>
            <a class="appk-btn" href="https://kontak157.ojk.go.id/appkpublicportal/" target="_blank">
                Buka APPK OJK &#8599;
            </a>
        </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("📄 Lanjut ke Formulir Surat Resmi"):
            st.session_state.halaman = "formulir"
            st.rerun()

    st.caption("© 2026 Otoritas Jasa Keuangan | Layanan Perlindungan Konsumen")

# ===================== HALAMAN FORMULIR =====================
elif st.session_state.halaman == "formulir":

    # Header mini di atas formulir
    st.markdown("""
    <style>
    .mini-header {
        background: linear-gradient(135deg, #8B0000, #b22222);
        border-radius: 10px;
        padding: 16px 24px;
        display: flex; align-items: center; gap: 16px;
        margin-bottom: 4px;
        border-bottom: 3px solid #FFD700;
        font-family: Arial, sans-serif;
    }
    .mini-logo {
        width:48px; height:48px; border-radius:50%; background:white;
        border:2px solid #FFD700;
        display:flex; align-items:center; justify-content:center;
        font-size:10px; font-weight:900; color:#8B0000;
        line-height:1.2; text-align:center; flex-shrink:0;
    }
    .mini-title { font-size:16px; font-weight:900; color:white; }
    .mini-sub   { font-size:11px; color:rgba(255,255,255,0.65); margin-top:2px; }
    .info-ribbon {
        background:#fffbf0; border-left:4px solid #FFD700;
        padding:11px 16px; border-radius:0 7px 7px 0;
        font-size:12.5px; color:#555; margin-bottom:8px;
        border:1px solid #f0e0a0; border-left:4px solid #FFD700;
        font-family:Arial,sans-serif; line-height:1.55;
    }
    </style>

    <div class="mini-header">
        <div class="mini-logo">OJK<br>JBR</div>
        <div>
            <div class="mini-title">Portal Perlindungan Konsumen</div>
            <div class="mini-sub">Formulir Layanan Surat Resmi &mdash; OJK Jember</div>
        </div>
    </div>
    <div class="info-ribbon">
        &#x2139;&#xFE0F;&nbsp; <strong>Petunjuk:</strong> Pastikan NIK sesuai KTP. Isi semua kolom dengan benar.
        Surat yang dihasilkan bersifat resmi dan dapat diajukan ke OJK Jember.
        &nbsp;|&nbsp; <a href="https://kontak157.ojk.go.id/appkpublicportal/" target="_blank"
        style="color:#8B0000;font-weight:700;">Pengaduan Online via APPK OJK &#8599;</a>
    </div>
    """, unsafe_allow_html=True)

    col_back, _ = st.columns([1, 4])
    with col_back:
        if st.button("← Kembali"):
            st.session_state.halaman = "pembukaan"
            st.rerun()

    # ---- INPUT DATA ----
    kronologis, no_hp, email, pt_dituju = "", "", "", ""
    tipe_surat = st.selectbox("Pilih Jenis Layanan:", ["Surat Pengaduan", "Surat Pernyataan"])
    nama   = st.text_input("Nama Lengkap")
    nik    = st.text_input("NIK / Nomor Identitas")
    if nik and (len(nik) != 16 or not nik.isdigit()):
        st.warning("NIK harus 16 digit angka.")
    alamat = st.text_area("Alamat")

    if tipe_surat == "Surat Pengaduan":
        pt_dituju  = st.text_input("Nama PT yang Dituju")
        no_hp      = st.text_input("No. HP")
        email      = st.text_input("Email")
        kronologis = st.text_area("Tuliskan kronologis permasalahan:")

    kota_ttd   = st.text_input("Kota", value="Jember")
    tanggal_ttd = st.date_input("Tanggal")

    # ---- FUNGSI PDF & WORD ----
    TEKS_PERNYATAAN = (
        "menyatakan dengan sesungguhnya bahwa permasalahan yang saya ajukan melalui "
        "Aplikasi Portal Perlindungan Konsumen (APPK) tidak sedang dalam proses atau "
        "pernah diputus oleh lembaga arbitrase atau peradilan atau lembaga mediasi lainnya "
        "termasuk lembaga alternatif penyelesaian Sengketa dan belum pernah difasilitasi oleh OJK."
    )
    TEKS_PENUTUP = "Demikian surat pernyataan ini dibuat dengan sadar dan tanpa paksaan dari pihak manapun."

    def cetak_baris(pdf, label, nilai):
        pdf.cell(40, 8, txt=label, ln=0)
        pdf.cell(5, 8, txt=":", ln=0)
        pdf.cell(0, 8, txt=nilai, ln=1)

    def bersihkan_teks(teks):
        if not teks: return ""
        return teks.encode('latin-1', 'replace').decode('latin-1')

    def buat_pdf():
        pdf = FPDF()
        pdf.add_page()
        pdf.set_margins(20, 20, 20)
        pdf.set_font("Arial", size=12)
        if tipe_surat == "Surat Pernyataan":
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, txt="SURAT PERNYATAAN", ln=True, align='C')
            pdf.ln(6)
            pdf.set_font("Arial", size=12)
            pdf.cell(0, 8, txt="Yang bertanda tangan di bawah ini:", ln=True)
            cetak_baris(pdf, "Nama", bersihkan_teks(nama))
            cetak_baris(pdf, "NIK", bersihkan_teks(nik))
            cetak_baris(pdf, "Alamat", bersihkan_teks(alamat))
            pdf.ln(6)
            pdf.multi_cell(0, 8, txt=bersihkan_teks(TEKS_PERNYATAAN))
            pdf.ln(6)
            pdf.multi_cell(0, 8, txt=bersihkan_teks(TEKS_PENUTUP))
        else:
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, txt="SURAT PENGADUAN", ln=True, align='C')
            pdf.ln(6)
            pdf.set_font("Arial", size=12)
            pdf.cell(0, 8, txt=bersihkan_teks(f"Nama PT yang dituju : {pt_dituju}"), ln=True)
            pdf.ln(5)
            pdf.cell(0, 8, txt="Yang bertanda tangan di bawah ini:", ln=True)
