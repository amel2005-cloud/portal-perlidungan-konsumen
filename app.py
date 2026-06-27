import streamlit as st
import streamlit.components.v1 as components from fpdf 
import FPDFfrom docx 
import Documentfrom docx.shared 
import Ptfrom io 
import BytesIO

if "halaman" not in st.session_state:st.session_state.halaman = "pembukaan"

===================== HALAMAN PEMBUKAAN =====================

if st.session_state.halaman == "pembukaan":

st.markdown("""
<style>
.stApp {
    background: #8B0000 !important;
    background-image: none !important;
}
section[data-testid="stSidebar"] { display: none; }
header[data-testid="stHeader"] { background: transparent; }
div.stButton > button:first-child {
    background-color: #FFD700 !important;
    color: #5a0000 !important;
    border: none !important;
    font-weight: 800 !important;
    font-size: 16px !important;
    padding: 14px 0 !important;
    border-radius: 8px !important;
    width: 100% !important;
    letter-spacing: 0.3px;
}
div.stButton > button:first-child:hover {
    background-color: #ffe033 !important;
}
</style>
""", unsafe_allow_html=True)

components.html("""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
* { margin:0; padding:0; box-sizing:border-box; }
body {
    background: #8B0000;
    font-family: Arial, sans-serif;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    text-align: center;
    padding: 40px 24px;
    position: relative;
    overflow: hidden;
}
body::before {
    content:'';
    position:fixed; inset:0;
    background: repeating-linear-gradient(
        45deg,transparent,transparent 30px,
        rgba(255,255,255,0.02) 30px,rgba(255,255,255,0.02) 60px
    );
    pointer-events:none;
}
.gold-top {
    position:fixed;top:0;left:0;right:0;height:5px;
    background:linear-gradient(90deg,#b8860b,#FFD700,#b8860b);
    z-index:999;
}
.gold-bot {
    position:fixed;bottom:0;left:0;right:0;height:5px;
    background:linear-gradient(90deg,#b8860b,#FFD700,#b8860b);
    z-index:999;
}
.logo-ring {
    width:120px;height:120px;border-radius:50%;
    background:white;border:3px solid #FFD700;
    display:flex;align-items:center;justify-content:center;
    margin:0 auto 24px;
    animation:popIn 0.6s ease 0.2s both;
}
.logo-ring span {
    font-size:16px;font-weight:900;color:#8B0000;
    line-height:1.3;letter-spacing:1px;
}
@keyframes popIn {
    from{transform:scale(0.4);opacity:0}
    to{transform:scale(1);opacity:1}
}
@keyframes fadeUp {
    from{transform:translateY(14px);opacity:0}
    to{transform:translateY(0);opacity:1}
}
.inst-label {
    font-size:11px;letter-spacing:3px;text-transform:uppercase;
    color:rgba(255,255,255,0.5);margin-bottom:16px;
    animation:fadeUp 0.5s ease 0.8s both;
}
.typed-container {
    min-height:52px;display:flex;
    align-items:center;justify-content:center;
    margin-bottom:12px;
}
#typed-out {
    font-size:30px;font-weight:900;color:white;
    text-shadow:0 2px 12px rgba(0,0,0,0.4);
    letter-spacing:0.5px;
}
.cursor-blink {
    display:inline-block;width:3px;height:1.15em;
    background:#FFD700;vertical-align:middle;margin-left:4px;
    animation:blink 0.7s step-end infinite;
}
@keyframes blink{0%,100%{opacity:1}50%{opacity:0}}
.sub-anim {
    font-size:14px;color:rgba(255,255,255,0.7);
    font-style:italic;margin-bottom:28px;
    opacity:0;transition:opacity 0.7s ease;
}
.badge-row {
    display:flex;gap:10px;flex-wrap:wrap;justify-content:center;
    margin-bottom:28px;
    opacity:0;transition:opacity 0.7s ease;
}
.badge {
    background:rgba(255,255,255,0.1);
    border:1px solid rgba(255,215,0,0.35);
    color:rgba(255,255,255,0.85);
    font-size:11.5px;padding:6px 15px;border-radius:20px;
    display:inline-flex;align-items:center;gap:7px;
}
.badge-dot {
    width:7px;height:7px;border-radius:50%;
    background:#FFD700;display:inline-block;flex-shrink:0;
}
.appk-card {
    background:rgba(255,255,255,0.1);
    border:1px solid rgba(255,215,0,0.45);
    border-radius:12px;padding:22px 30px;
    max-width:420px;margin:0 auto 20px;
    opacity:0;transition:opacity 0.7s ease;
}
.appk-card-title {
    font-size:10px;letter-spacing:2.5px;text-transform:uppercase;
    color:rgba(255,255,255,0.5);margin-bottom:10px;
}
.appk-card-text {
    font-size:13.5px;color:rgba(255,255,255,0.85);
    margin-bottom:16px;line-height:1.6;
}
.appk-btn {
    display:inline-block;background:#FFD700;color:#5a0000;
    font-weight:800;font-size:13px;
    padding:11px 24px;border-radius:7px;
    text-decoration:none;letter-spacing:0.3px;
}
.appk-btn:hover { background:#ffe033; }
.footer-txt {
    font-size:11px;color:rgba(255,255,255,0.3);
    margin-top:12px;
}
</style>
</head>
<body>
<div class="gold-top"></div>
<div class="gold-bot"></div>

<div class="logo-ring"><span>OJK<br>JEMBER</span></div>
<div class="inst-label">Otoritas Jasa Keuangan &bull; Kantor Jember</div>

<div class="typed-container">
    <span id="typed-out"></span><span class="cursor-blink" id="cur"></span>
</div>

<div class="sub-anim" id="sub">Layanan Surat Resmi Terintegrasi &mdash; OJK Jember</div>

<div class="badge-row" id="badges">
    <span class="badge"><span class="badge-dot"></span>Resmi &amp; Terpercaya</span>
    <span class="badge"><span class="badge-dot"></span>Pengaduan Terverifikasi</span>
    <span class="badge"><span class="badge-dot"></span>Layanan Gratis</span>
</div>

<div class="appk-card" id="appk">
    <div class="appk-card-title">Layanan Pengaduan Online OJK</div>
    <div class="appk-card-text">
        Akses langsung ke <strong style="color:#FFD700;">APPK (Aplikasi Portal Perlindungan Konsumen)</strong>
        untuk mengajukan pengaduan secara online ke OJK Pusat.
    </div>
    <a class="appk-btn" href="https://kontak157.ojk.go.id/appkpublicportal/" target="_blank">
        Buka APPK OJK &#8599;
    </a>
</div>

<div class="footer-txt">&#169; 2026 Otoritas Jasa Keuangan | Layanan Perlindungan Konsumen</div>

<script>
(function(){
    var teks   = "Portal Perlindungan Konsumen";
    var el     = document.getElementById("typed-out");
    var cur    = document.getElementById("cur");
    var sub    = document.getElementById("sub");
    var badges = document.getElementById("badges");
    var appk   = document.getElementById("appk");
    var i = 0;
    function ketik(){
        if(i < teks.length){
            el.textContent += teks[i];
            i++;
            setTimeout(ketik, 60);
        } else {
            setTimeout(function(){ cur.style.display="none"; }, 800);
            setTimeout(function(){ sub.style.opacity="1"; }, 400);
            setTimeout(function(){ badges.style.opacity="1"; }, 800);
            setTimeout(function(){ appk.style.opacity="1"; }, 1200);
        }
    }
    setTimeout(ketik, 1000);
})();
</script>
</body>
</html>
""", height=700, scrolling=False)

col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    if st.button("📄 Lanjut ke Formulir Surat Resmi"):
        st.session_state.halaman = "formulir"
        st.rerun()

===================== HALAMAN FORMULIR =====================

elif st.session_state.halaman == "formulir":

st.markdown("""
<style>
.stApp {
    background-image: radial-gradient(#d1d1d1 0.5px, transparent 0.5px);
    background-size: 20px 20px;
    background-color: #f8f9fa;
}
.stTextInput > div > div > input,
.stTextArea > div > div > textarea {
    border-left: 4px solid #b22222 !important;
    border-radius: 0 6px 6px 0 !important;
}
div.stButton > button:first-child {
    background-color: #b22222 !important;
    color: white !important;
    border: none !important;
    font-weight: bold !important;
    width: 100% !important;
    padding: 12px 0 !important;
    font-size: 15px !important;
    border-radius: 8px !important;
}
div.stButton > button:first-child:hover {
    background-color: #8B0000 !important;
}
.mini-header {
    background: linear-gradient(135deg, #8B0000, #b22222);
    border-radius: 10px;
    padding: 16px 24px;
    display: flex; align-items: center; gap: 16px;
    margin-bottom: 12px;
    border-bottom: 3px solid #FFD700;
    font-family: Arial, sans-serif;
}
.mini-logo {
    width:50px;height:50px;border-radius:50%;
    background:white;border:2px solid #FFD700;
    display:flex;align-items:center;justify-content:center;
    font-size:10px;font-weight:900;color:#8B0000;
    line-height:1.2;text-align:center;flex-shrink:0;
}
.mini-title { font-size:16px;font-weight:900;color:white; }
.mini-sub   { font-size:11px;color:rgba(255,255,255,0.65);margin-top:2px; }
.info-ribbon {
    background:#fffbf0;
    border-left:4px solid #FFD700;
    padding:11px 16px;border-radius:0 7px 7px 0;
    font-size:12.5px;color:#555;margin-bottom:16px;
    border:1px solid #f0e0a0;border-left:4px solid #FFD700;
    font-family:Arial,sans-serif;line-height:1.55;
}
.form-section-header {
    background:#f8f0f0;border-left:4px solid #b22222;
    padding:10px 16px;border-radius:0 6px 6px 0;
    font-size:13px;font-weight:700;color:#7a0000;
    margin-bottom:8px;margin-top:12px;
    font-family:Arial,sans-serif;
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
    &#x2139;&#xFE0F;&nbsp;
    <strong>Petunjuk Pengisian:</strong> Pastikan NIK sesuai KTP (16 digit).
    Isi semua kolom dengan benar dan lengkap.
    &nbsp;|&nbsp;
    <a href="https://kontak157.ojk.go.id/appkpublicportal/" target="_blank"
    style="color:#8B0000;font-weight:700;text-decoration:none;">
    Pengaduan Online via APPK OJK &#8599;</a>
</div>
""", unsafe_allow_html=True)

col_back, _ = st.columns([1, 5])
with col_back:
    if st.button("← Kembali"):
        st.session_state.halaman = "pembukaan"
        st.rerun()

kronologis, no_hp, email, pt_dituju = "", "", "", ""

st.markdown('<div class="form-section-header">⚙️ Jenis Layanan</div>', unsafe_allow_html=True)
tipe_surat = st.selectbox("Jenis Layanan", ["Surat Pengaduan", "Surat Pernyataan"], label_visibility="collapsed")

st.markdown('<div class="form-section-header">👤 Data Diri Pemohon</div>', unsafe_allow_html=True)
nama   = st.text_input("Nama Lengkap")
nik    = st.text_input("NIK / Nomor Identitas (16 digit)")
if nik and (len(nik) != 16 or not nik.isdigit()):
    st.warning("⚠️ NIK harus berupa 16 digit angka.")
alamat = st.text_area("Alamat Lengkap")

if tipe_surat == "Surat Pengaduan":
    st.markdown('<div class="form-section-header">🏢 Data Pengaduan</div>', unsafe_allow_html=True)
    pt_dituju  = st.text_input("Nama PT / Lembaga yang Dituju")
    no_hp      = st.text_input("No. HP / WhatsApp")
    email      = st.text_input("Alamat Email")
    kronologis = st.text_area("Kronologis Permasalahan", height=180,
                               placeholder="Tuliskan kronologis secara jelas dan runtut...")

st.markdown('<div class="form-section-header">📅 Tanggal & Tempat</div>', unsafe_allow_html=True)
col_kota, col_tgl = st.columns(2)
with col_kota:
    kota_ttd    = st.text_input("Kota", value="Jember")
with col_tgl:
    tanggal_ttd = st.date_input("Tanggal Surat")

st.markdown("<br>", unsafe_allow_html=True)

TEKS_PERNYATAAN = (
    "menyatakan dengan sesungguhnya bahwa permasalahan yang saya ajukan melalui "
    "Aplikasi Portal Perlindungan Konsumen (APPK) tidak sedang dalam proses atau "
    "pernah diputus oleh lembaga arbitrase atau peradilan atau lembaga mediasi lainnya "
    "termasuk lembaga alternatif penyelesaian Sengketa dan belum pernah difasilitasi oleh OJK."
)
TEKS_PENUTUP = "Demikian surat pernyataan ini dibuat dengan sadar dan tanpa paksaan dari pihak manapun."

def cetak_baris(pdf, label, nilai):
    pdf.cell(40, 8, txt=label, ln=0)
    pdf.cell(5,  8, txt=":", ln=0)
    pdf.cell(0,  8, txt=nilai, ln=1)

def bersihkan(teks):
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
        cetak_baris(pdf, "Nama",   bersihkan(nama))
        cetak_baris(pdf, "NIK",    bersihkan(nik))
        cetak_baris(pdf, "Alamat", bersihkan(alamat))
        pdf.ln(6)
        pdf.multi_cell(0, 8, txt=bersihkan(TEKS_PERNYATAAN))
        pdf.ln(6)
        pdf.multi_cell(0, 8, txt=bersihkan(TEKS_PENUTUP))
    else:
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, txt="SURAT PENGADUAN", ln=True, align='C')
        pdf.ln(6)
        pdf.set_font("Arial", size=12)
        pdf.cell(0, 8, txt=bersihkan(f"Kepada Yth. {pt_dituju}"), ln=True)
        pdf.ln(4)
        pdf.cell(0, 8, txt="Yang bertanda tangan di bawah ini:", ln=True)
        cetak_baris(pdf, "Nama",   bersihkan(nama))
        cetak_baris(pdf, "NIK",    bersihkan(nik))
        cetak_baris(pdf, "Alamat", bersihkan(alamat))
        cetak_baris(pdf, "No. HP", bersihkan(no_hp))
        cetak_baris(pdf, "Email",  bersihkan(email))
        pdf.ln(5)
        pdf.cell(0, 8, txt="Kronologis permasalahan:", ln=True)
        pdf.multi_cell(0, 8, txt=bersihkan(kronologis))
    pdf.ln(20)
    pdf.cell(110)
    pdf.cell(0, 8, txt=bersihkan(f"{kota_ttd}, {tanggal_ttd.strftime('%d %B %Y')}"), ln=True)
    pdf.ln(20)
    pdf.cell(110)
    pdf.cell(0, 8, txt=bersihkan(f"({nama})"), ln=True)
    pdf.output("surat_hasil.pdf")
    with open("surat_hasil.pdf", "rb") as f:
        return f.read()

def buat_word():
    doc = Document()
    if tipe_surat == "Surat Pernyataan":
        judul = doc.add_paragraph("SURAT PERNYATAAN")
        judul.alignment = 1
        judul.runs[0].bold = True
        judul.runs[0].font.size = Pt(14)
        doc.add_paragraph("")
        doc.add_paragraph("Yang bertanda tangan di bawah ini:")
        doc.add_paragraph(f"Nama      : {nama}")
        doc.add_paragraph(f"NIK           : {nik}")
        doc.add_paragraph(f"Alamat    : {alamat}")
        doc.add_paragraph("")
        doc.add_paragraph(TEKS_PERNYATAAN)
        doc.add_paragraph("")
        doc.add_paragraph(TEKS_PENUTUP)
    else:
        judul = doc.add_paragraph("SURAT PENGADUAN")
        judul.alignment = 1
        judul.runs[0].bold = True
        judul.runs[0].font.size = Pt(14)
        doc.add_paragraph("")
        doc.add_paragraph(f"Kepada Yth. {pt_dituju}")
        doc.add_paragraph("")
        doc.add_paragraph("Yang bertanda tangan di bawah ini:")
        doc.add_paragraph(f"Nama      : {nama}")
        doc.add_paragraph(f"NIK           : {nik}")
        doc.add_paragraph(f"Alamat    : {alamat}")
        doc.add_paragraph(f"No. HP   : {no_hp}")
        doc.add_paragraph(f"Email       : {email}")
        doc.add_paragraph("")
        doc.add_paragraph("Kronologis permasalahan:")
        doc.add_paragraph(kronologis)
    doc.add_paragraph("")
    doc.add_paragraph("")
    doc.add_paragraph(
        f"                                                  "
        f"{kota_ttd}, {tanggal_ttd.strftime('%d %B %Y')}"
    )
    doc.add_paragraph("")
    doc.add_paragraph("")
    doc.add_paragraph(f"                                                  ({nama})")
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

if st.button("🖨️ PROSES & CETAK SURAT"):
    if not nama.strip():
        st.error("⚠️ Nama lengkap tidak boleh kosong.")
    elif len(nik) != 16 or not nik.isdigit():
        st.error("⚠️ NIK harus berupa 16 digit angka.")
    elif not alamat.strip():
        st.error("⚠️ Alamat tidak boleh kosong.")
    elif tipe_surat == "Surat Pengaduan" and not pt_dituju.strip():
        st.error("⚠️ Nama PT / lembaga yang dituju tidak boleh kosong.")
    elif tipe_surat == "Surat Pengaduan" and not kronologis.strip():
        st.error("⚠️ Kronologis permasalahan tidak boleh kosong.")
    else:
        with st.spinner("Membuat surat..."):
            pdf_bytes   = buat_pdf()
            word_buffer = buat_word()
        st.success("✅ Surat berhasil dibuat! Silakan unduh di bawah.")
        col_pdf, col_word = st.columns(2)
        with col_pdf:
            st.download_button(
                "📄 Download PDF", pdf_bytes,
                file_name="surat_ojk_jember.pdf",
                mime="application/pdf"
            )
        with col_word:
            st.download_button(
                "📝 Download Word", word_buffer,
                file_name="surat_ojk_jember.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

st.markdown("<br>", unsafe_allow_html=True)
st.caption("© 2026 Otoritas Jasa Keuangan | Layanan Perlindungan Konsumen")
