import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import re
import datetime
import plotly.express as px
import qrcode
from io import BytesIO
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import base64
from PIL import Image
import numpy as np
import random
from streamlit_drawable_canvas import st_canvas

# ------------------------------------------------------------
# KONFIGURASI HALAMAN & CUSTOM CSS MODERN
# ------------------------------------------------------------
st.set_page_config(
    page_title="Sistem Absensi Digital Badung", 
    page_icon="🏛️", 
    layout="wide"
)

st.markdown("""
    <style>
    div[data-testid="InputInstructions"] {
        display: none !important;
    }

    .hero-header-banner {
        background-color: #0E3B2E;
        border-top: 3px solid #D4AF37;
        border-bottom: 3px solid #D4AF37;
        padding: 18px 30px;
        border-radius: 6px;
        box-shadow: 0 6px 16px rgba(0,0,0,0.3);
        display: flex;
        align-items: center;
        gap: 25px;
        margin-bottom: 15px;
    }

    .hero-logo-ring-outer {
        width: 95px;
        height: 95px;
        border-radius: 50%;
        border: 3px solid #E29D29;
        box-shadow: 0 0 0 2px #0E3B2E, 0 0 0 5px #B27300;
        display: flex;
        align-items: center;
        justify-content: center;
        background-color: #0E3B2E;
        flex-shrink: 0;
    }

    .hero-logo-ring-outer img {
        width: 62px;
        height: auto;
        filter: drop-shadow(0px 2px 4px rgba(0,0,0,0.4));
    }

    .hero-title-main {
        font-family: 'Montserrat', 'Arial Black', sans-serif;
        font-weight: 900;
        font-size: 23px;
        color: #E29D29;
        letter-spacing: 1px;
        line-height: 1.15;
        text-transform: uppercase;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.5);
    }

    .hero-title-sub {
        font-family: 'Arial', sans-serif;
        font-weight: 600;
        font-size: 15px;
        color: #D4AF37;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        margin-top: 3px;
    }

    .agenda-card {
        background: #F4F6F4;
        border-left: 5px solid #1B5E20;
        padding: 12px 20px;
        border-radius: 6px;
        margin-bottom: 25px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 10px;
    }

    .agenda-title {
        font-size: 15px;
        font-weight: 700;
        color: #1B5E20;
    }

    .status-badge {
        background: #1B5E20;
        color: #FFFFFF;
        padding: 4px 12px;
        border-radius: 15px;
        font-size: 12px;
        font-weight: 600;
        display: inline-flex;
        align-items: center;
        gap: 6px;
    }

    .status-dot {
        width: 8px;
        height: 8px;
        background-color: #00E676;
        border-radius: 50%;
        box-shadow: 0 0 6px #00E676;
    }

    h1 { color: #1B5E20 !important; }
    h2, h3 { color: #C67D0A !important; }
    div[data-testid="stMetricValue"] { color: #1B5E20 !important; font-weight: bold; }

    div[data-testid="stFormSubmitButton"] > button {
        background: linear-gradient(135deg, #D87A00 0%, #B25900 100%) !important;
        color: white !important;
        border-radius: 10px !important;
        border: none !important;
        font-weight: 700 !important;
        width: 100%;
        font-size: 16px !important;
        padding: 12px 24px !important;
        box-shadow: 0 4px 14px 0 rgba(178, 89, 0, 0.35) !important;
        transition: all 0.3s ease-in-out !important;
        letter-spacing: 0.5px !important;
    }
    div[data-testid="stFormSubmitButton"] > button:hover {
        background: linear-gradient(135deg, #1B5E20 0%, #0B3C11 100%) !important;
        box-shadow: 0 6px 20px 0 rgba(27, 94, 32, 0.4) !important;
        transform: translateY(-2px) !important;
    }
    </style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------
# SESSION STATES & KONEKSI GOOGLE SHEETS
# ------------------------------------------------------------
conn = st.connection("gsheets", type=GSheetsConnection)

if "nama_event" not in st.session_state:
    st.session_state.nama_event = "Kegiatan Sosialisasi & Bimbingan Teknis"

if "jenis_presensi" not in st.session_state:
    st.session_state.jenis_presensi = "Presensi Zoom"

if "target_sheet_name" not in st.session_state:
    st.session_state.target_sheet_name = "Sheet1"

def load_data_from_gsheets():
    sheet_name = st.session_state.target_sheet_name
    try:
        df = conn.read(worksheet=sheet_name, ttl=0)
        return df
    except Exception:
        return pd.DataFrame(columns=[
            "Waktu Absen", "Jam Absen", "Jenis Presensi", "Nama Peserta (EYD)", "NIP", "Jabatan (EYD)", "OPD / Instansi", "Email", "Tanda Tangan", "Keterangan Status"
        ])

def save_data_to_gsheets(new_row_df, existing_df):
    sheet_name = st.session_state.target_sheet_name
    cols_to_save = [c for c in existing_df.columns if c != "_ttd_bytes"]
    clean_existing = existing_df[cols_to_save] if not existing_df.empty else pd.DataFrame(columns=cols_to_save)
    clean_new = new_row_df[cols_to_save]
    
    updated_df = pd.concat([clean_existing, clean_new], ignore_index=True)
    conn.update(worksheet=sheet_name, data=updated_df)
    st.cache_data.clear()

# Muat data dari Google Sheets berdasarkan sheet aktif
st.session_state.df_absensi = load_data_from_gsheets()

if "cutoff_enabled" not in st.session_state:
    st.session_state.cutoff_enabled = False

if "cutoff_time" not in st.session_state:
    st.session_state.cutoff_time = datetime.time(12, 0)

if "max_peserta_enabled" not in st.session_state:
    st.session_state.max_peserta_enabled = False

if "max_peserta_limit" not in st.session_state:
    st.session_state.max_peserta_limit = 100

if "is_admin_logged_in" not in st.session_state:
    st.session_state.is_admin_logged_in = False

ADMIN_PIN = "4869"

if "random_nip_placeholder" not in st.session_state:
    y = random.randint(1975, 2002)
    m = f"{random.randint(1, 12):02d}"
    d = f"{random.randint(1, 28):02d}"
    y_pns = y + random.randint(20, 26)
    m_pns = f"{random.randint(1, 12):02d}"
    gender = random.choice([1, 2])
    seq = f"{random.randint(1, 999):03d}"
    st.session_state.random_nip_placeholder = f"Contoh: {y}{m}{d}{y_pns}{m_pns}{gender}{seq}"

# ------------------------------------------------------------
# HEADER BANNER MURNI CSS
# ------------------------------------------------------------
def get_image_base64(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

local_logo_badung = Path(__file__).parent / "badung.png"
if local_logo_badung.exists():
    LOGO_BADUNG = f"data:image/png;base64,{get_image_base64(local_logo_badung)}"
else:
    LOGO_BADUNG = "https://upload.wikimedia.org/wikipedia/commons/1/1a/Lambang_Kabupaten_Badung.portal.png"

st.markdown(f"""
    <div class="hero-header-banner">
        <div class="hero-logo-ring-outer">
            <img src="{LOGO_BADUNG}" alt="Logo Pemkab Badung">
        </div>
        <div>
            <div class="hero-title-main">PEMERINTAH<br>KABUPATEN BADUNG</div>
            <div class="hero-title-sub">BAGIAN ORGANISASI</div>
        </div>
    </div>
""", unsafe_allow_html=True)

badge_label = "PRESENSI ZOOM AKTIF" if st.session_state.jenis_presensi == "Presensi Zoom" else "PRESENSI TATAP MUKA AKTIF"

st.markdown(f"""
    <div class="agenda-card">
        <div class="agenda-title">📌 Agenda: {st.session_state.nama_event} ({st.session_state.jenis_presensi})</div>
        <div class="status-badge">
            <span class="status-dot"></span> {badge_label}
        </div>
    </div>
""", unsafe_allow_html=True)

# ------------------------------------------------------------
# NAVIGASI SIDEBAR
# ------------------------------------------------------------
st.sidebar.title("📌 Menu Navigasi")
menu_pilihan = st.sidebar.radio(
    "Pilih Halaman:",
    ["📝 Form Absensi & Dashboard", "📱 Scan QR Code", "🔐 Panel Admin (Khusus Panitia)"]
)

if menu_pilihan != "🔐 Panel Admin (Khusus Panitia)":
    st.session_state.is_admin_logged_in = False

# ------------------------------------------------------------
# MASTER DATA OPD & HELPER
# ------------------------------------------------------------
LIST_OPD = [
    "Inspektorat", "Sekretariat Daerah", "Sekretariat DPRD", "Badan Kesatuan Bangsa Dan Politik",
    "Badan Penanggulangan Bencana Daerah", "Badan Perencanaan Pembangunan Daerah",
    "Badan Pengelola Keuangan dan Aset Daerah", "Badan Pendapatan Daerah",
    "Badan Kepegawaian dan Pengembangan Sumber Daya Manusia", "Badan Riset dan Inovasi Daerah",
    "Dinas Pendidikan, Kepemudaan dan Olah Raga", "Dinas Kesehatan",
    "Dinas Pekerjaan Umum dan Penataan Ruang", "Dinas Perumahan Rakyat dan Kawasan Permukiman",
    "Dinas Kebakaran dan Penyelamatan", "Dinas Sosial", "Dinas Lingkungan Hidup dan Kebersihan",
    "Dinas Kependudukan dan Pencatatan Sipil", "Dinas Pemberdayaan Masyarakat dan Desa",
    "Dinas Pengendalian Penduduk, Keluarga Berencana, Pemberdayaan Perempuan dan Perlindungan Anak",
    "Dinas Perhubungan", "Dinas Komunikasi dan Informatika",
    "Dinas Koperasi, Usaha Kecil Menengah dan Perdagangan",
    "Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu", "Dinas Kebudayaan",
    "Dinas Kearsipan dan Perpustakaan", "Dinas Perikanan", "Dinas Pariwisata",
    "Dinas Pertanian dan Pangan", "Dinas Perindustrian dan Tenaga Kerja",
    "Satuan Polisi Pamong Praja", "Kecamatan Kuta", "Kecamatan Kuta Utara",
    "Kecamatan Kuta Selatan", "Kecamatan Mengwi", "Kecamatan Abiansemal", "Kecamatan Petang",
    "RSD Mangusada", "Lainnya / Instansi luar"
]

KAMUS_GELAR = {
    "SKOM": "S.Kom", "S.KOM": "S.Kom", "SE": "S.E", "S.E.": "S.E",
    "SH": "S.H", "S.H.": "S.H", "ST": "S.T", "S.T.": "S.T",
    "SSTP": "S.STP", "MSI": "M.Si", "M.SI": "M.Si",
    "MAP": "M.A.P", "M.AP": "M.A.P", "MM": "M.M", "MH": "M.H",
    "SSOS": "S.Sos", "SPD": "S.Pd", "DR": "dr.", "DRG": "drg."
}

TYPO_DOMAINS = {
    "gmai.com": "gmail.com", "gamil.com": "gmail.com", "gmai.co": "gmail.com",
    "yaho.com": "yahoo.com", "yahoo.co": "yahoo.com", "yaho.co.id": "yahoo.co.id"
}

def format_eyd_nama(nama: str) -> str:
    if not nama: return ""
    nama_clean = re.sub(r'\s+', ' ', str(nama)).strip()
    parts = nama_clean.split(",")
    nama_words = parts[0].split(" ")
    nama_formatted = []
    for w in nama_words:
        w_up = w.upper().replace(".", "")
        if w_up in KAMUS_GELAR: nama_formatted.append(KAMUS_GELAR[w_up])
        else: nama_formatted.append(w.capitalize())
    hasil = " ".join(nama_formatted)
    if len(parts) > 1:
        gelar_list = [KAMUS_GELAR.get(g.strip().replace(".", "").upper(), g.strip().upper()) for g in parts[1:]]
        hasil += ", " + ", ".join(gelar_list)
    return hasil

def format_eyd_jabatan(jabatan: str) -> str:
    if not jabatan: return ""
    clean_text = re.sub(r'\s+', ' ', str(jabatan)).strip()
    kata_kecil = {"dan", "atau", "pada", "di", "ke", "dari", "untuk", "yang"}
    words = clean_text.split(" ")
    formatted = []
    for idx, w in enumerate(words):
        w_lower = w.lower()
        if idx > 0 and w_lower in kata_kecil:
            formatted.append(w_lower)
        else:
            formatted.append(w.capitalize())
    return " ".join(formatted)

def format_nama_instansi(instansi: str) -> str:
    if not instansi: return ""
    clean_text = re.sub(r'\s+', ' ', str(instansi)).strip()
    kata_kecil = {"dan", "atau", "pada", "di", "ke", "dari", "untuk", "yang"}
    words = clean_text.split(" ")
    formatted_words = []
    for idx, word in enumerate(words):
        word_lower = word.lower()
        if idx > 0 and word_lower in kata_kecil:
            formatted_words.append(word_lower)
        else:
            formatted_words.append(word.capitalize())
    return " ".join(formatted_words)

def clean_cmp(text: str) -> str:
    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).lower() if text else ""

def validate_nip(nip_str: str) -> tuple[bool, str]:
    nip_clean = re.sub(r'\D', '', str(nip_str).strip())
    if not nip_clean: return True, ""
    if len(nip_clean) != 18: return False, f"Jumlah digit NIP harus 18 angka (saat ini {len(nip_clean)} digit)."
    return True, nip_clean

def validate_email(email_str: str) -> tuple[bool, str, str]:
    if not email_str: return False, "", "Alamat email wajib diisi."
    clean_e = re.sub(r'\s+', '', str(email_str)).lower()
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    if not re.match(pattern, clean_e): return False, clean_e, "Format email tidak valid."
    domain = clean_e.split("@")[1]
    if domain in TYPO_DOMAINS:
        sug_e = clean_e.split('@')[0] + "@" + TYPO_DOMAINS[domain]
        return False, clean_e, f"Apakah maksud Anda '{sug_e}'? Domain '{domain}' terdeteksi typo."
    return True, clean_e, "Valid"

def generate_qr_code(url_text: str):
    qr = qrcode.QRCode(box_size=8, border=2)
    qr.add_data(url_text)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def export_formatted_excel(df: pd.DataFrame, nama_event: str) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rekap Presensi"
    ws.views.sheetView[0].showGridLines = True

    font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Arial", size=10, color="000000")
    
    fill_header = PatternFill(start_color="1B5E20", end_color="1B5E20", fill_type="solid")
    fill_dup = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='D0D0D0'),
        right=Side(style='thin', color='D0D0D0'),
        top=Side(style='thin', color='D0D0D0'),
        bottom=Side(style='thin', color='D0D0D0')
    )

    ws.append(["BAGIAN ORGANISASI KABUPATEN BADUNG"])
    ws.append([f"REKAPITULASI PRESENSI: {nama_event.upper()}"])
    ws.append([f"Tanggal Ekspor: {datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')} WITA"])
    ws.append([])

    export_cols = [c for c in df.columns if c != "_ttd_bytes"]
    headers = ["No"] + export_cols
    ws.append(headers)
    
    header_row_idx = 5
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=header_row_idx, column=col_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for i in range(len(df)):
        r_idx = i + 1
        row_num = header_row_idx + r_idx
        
        row_values = [r_idx] + [df.iloc[i][col] for col in export_cols]
        ws.append(row_values)
        
        status_val = str(df.iloc[i].get("Keterangan Status", ""))
        is_dup = (status_val == "DUPLIKAT / ABSEN GANDA")

        for c_idx in range(1, len(row_values) + 1):
            cell = ws.cell(row=row_num, column=c_idx)
            cell.font = font_data
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", horizontal="center" if c_idx in [1, 2, 3, 5] else "left")
            if is_dup:
                cell.fill = fill_dup

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# ------------------------------------------------------------
# HALAMAN 1: FORM ABSENSI & DASHBOARD
# ------------------------------------------------------------
if menu_pilihan == "📝 Form Absensi & Dashboard":
    st.subheader(f"Form Presensi Kehadiran — {st.session_state.nama_event}")
    
    now_time = datetime.datetime.now().time()
    
    is_closed_by_time = st.session_state.cutoff_enabled and (now_time > st.session_state.cutoff_time)
    
    total_valid_peserta = len(st.session_state.df_absensi[st.session_state.df_absensi["Keterangan Status"] == "UTAMA / VALID"]) if not st.session_state.df_absensi.empty and "Keterangan Status" in st.session_state.df_absensi.columns else 0
    is_closed_by_quota = st.session_state.max_peserta_enabled and (total_valid_peserta >= st.session_state.max_peserta_limit)
    
    if is_closed_by_time:
        st.error(f"⛔ Absensi telah DITUTUP (Batas Waktu Jam: {st.session_state.cutoff_time.strftime('%H:%M')} WITA). Silakan hubungi panitia.")
    elif is_closed_by_quota:
        st.error(f"⛔ Absensi telah DITUTUP (Kuota Peserta Penuh: {total_valid_peserta}/{st.session_state.max_peserta_limit} Peserta). Silakan hubungi panitia.")
    else:
        is_zoom_mode = (st.session_state.jenis_presensi == "Presensi Zoom")
        
        col1, col2 = st.columns(2)
        with col1:
            in_nama = st.text_input("Nama Lengkap & Gelar (EYD):", placeholder="Contoh: aditya putra, se")
            in_nip = st.text_input("NIP (18 Digit Angka):", placeholder=st.session_state.random_nip_placeholder)
            in_jabatan = st.text_input("Jabatan (EYD):", placeholder="Contoh: analis kebijakan ahli muda")
        
        with col2:
            in_opd = st.selectbox("OPD / Instansi Peserta:", LIST_OPD)
            if in_opd == "Lainnya / Instansi luar":
                manual_opd = st.text_input("Tuliskan Nama Instansi / OPD Anda:", placeholder="Contoh: kementerian hukum dan ham")
                in_opd = format_nama_instansi(manual_opd) if manual_opd.strip() else "Instansi Lainnya"
            
            if is_zoom_mode:
                in_email = st.text_input("Alamat Email Valid:", placeholder="Contoh: peserta@gmail.com")
            else:
                in_email = "-"

        ttd_status = "Tanda Tangan Daring (Zoom)"
        ttd_bytes_data = None
        
        if not is_zoom_mode:
            st.markdown("### ✍️ Tanda Tangan Digital Peserta")
            st.caption("Silakan bubuhkan tanda tangan Anda pada area kotak di bawah ini menggunakan jari atau stylus HP:")
            
            if "canvas_key" not in st.session_state:
                st.session_state.canvas_key = "canvas_ttd_0"

            canvas_result = st_canvas(
                fill_color="rgba(255, 255, 255, 0)",
                stroke_width=2,
                stroke_color="#000000",
                background_color="#FFFFFF",
                height=160,
                width=420,
                drawing_mode="freedraw",
                update_streamlit=True,
                return_image_data=True,
                key=st.session_state.canvas_key
            )
            
            if st.button("🗑️ Hapus / Ulangi Tanda Tangan"):
                st.session_state.canvas_key = f"canvas_ttd_{datetime.datetime.now().timestamp()}"
                st.rerun()

            if canvas_result.image_data is not None:
                img_array = canvas_result.image_data.astype(np.uint8)
                if np.any(img_array[:, :, 3] > 0):
                    ttd_status = "Ada TTD Digital"
                    pil_img = Image.fromarray(img_array)
                    buf = BytesIO()
                    pil_img.save(buf, format="PNG")
                    ttd_bytes_data = buf.getvalue()

        if st.button("🚀 Kirim Presensi Kehadiran", type="primary"):
            nip_valid, nip_msg = validate_nip(in_nip)
            
            if is_zoom_mode:
                email_valid, email_clean, email_msg = validate_email(in_email)
            else:
                email_valid, email_clean, email_msg = True, "-", ""
            
            if not in_nama.strip():
                st.error("⚠️ Nama Peserta wajib diisi!")
            elif not nip_valid:
                st.error(f"⚠️ Validasi NIP Gagal: {nip_msg}")
            elif not email_valid:
                st.error(f"⚠️ Validasi Email Gagal: {email_msg}")
            else:
                nama_eyd = format_eyd_nama(in_nama)
                jabatan_eyd = format_eyd_jabatan(in_jabatan)
                waktu_now = datetime.datetime.now()
                waktu_str = waktu_now.strftime("%Y-%m-%d %H:%M:%S")
                jam_str = waktu_now.strftime("%H:%M")
                
                df_ex = st.session_state.df_absensi
                is_dup = False
                if not df_ex.empty and "NIP" in df_ex.columns:
                    em_c = clean_cmp(email_clean)
                    nm_c = clean_cmp(nama_eyd)
                    nip_c = clean_cmp(in_nip)
                    for _, r in df_ex.iterrows():
                        if (is_zoom_mode and em_c and "Email" in r and em_c == clean_cmp(r["Email"])) or \
                           (nm_c and "Nama Peserta (EYD)" in r and nm_c == clean_cmp(r["Nama Peserta (EYD)"])) or \
                           (nip_c and len(nip_c) >= 18 and "NIP" in r and nip_c == clean_cmp(r["NIP"])):
                            is_dup = True
                            break
                            
                status_abs = "DUPLIKAT / ABSEN GANDA" if is_dup else "UTAMA / VALID"
                
                new_row = {
                    "Waktu Absen": waktu_str,
                    "Jam Absen": jam_str,
                    "Jenis Presensi": st.session_state.jenis_presensi,
                    "Nama Peserta (EYD)": nama_eyd,
                    "NIP": re.sub(r'\D', '', in_nip.strip()),
                    "Jabatan (EYD)": jabatan_eyd,
                    "OPD / Instansi": in_opd,
                    "Email": email_clean,
                    "Tanda Tangan": ttd_status,
                    "Keterangan Status": status_abs
                }
                
                new_row_df = pd.DataFrame([new_row])
                
                with st.spinner(f"Menyimpan presensi ke Google Sheets (Tab: {st.session_state.target_sheet_name})..."):
                    save_data_to_gsheets(new_row_df, df_ex)
                
                if is_dup:
                    st.warning(f"⚠️ Presensi diterima, namun ditandai sebagai **{status_abs}**.")
                else:
                    st.success(f"✅ Presensi Berhasil Disimpan & Terintegrasi Google Sheets! Terima kasih, **{nama_eyd}**.")
                
                st.rerun()

    # DASHBOARD MONITORING
    st.markdown("---")
    st.subheader("📊 Dashboard Monitoring Real-Time")
    df_live = st.session_state.df_absensi
    
    if not df_live.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Kehadiran", len(df_live))
        c2.metric("Data Valid (Utama)", len(df_live[df_live["Keterangan Status"] == "UTAMA / VALID"]) if "Keterangan Status" in df_live.columns else len(df_live))
        c3.metric("Absen Ganda", len(df_live[df_live["Keterangan Status"] == "DUPLIKAT / ABSEN GANDA"]) if "Keterangan Status" in df_live.columns else 0)
        
        g1, g2 = st.columns(2)
        GREEN_GOLD_PALETTE = ['#1B5E20', '#C67D0A', '#2E7D32', '#D87A00', '#4CAF50', '#FFB74D', '#0B3C11', '#B25900']
        
        with g1:
            if "OPD / Instansi" in df_live.columns:
                fig_p = px.pie(
                    df_live, 
                    names="OPD / Instansi", 
                    hole=0.4, 
                    title="Sebaran OPD Peserta",
                    color_discrete_sequence=GREEN_GOLD_PALETTE
                )
                st.plotly_chart(fig_p, use_container_width=True)
            
        with g2:
            if "Jam Absen" in df_live.columns:
                df_t = df_live.groupby("Jam Absen").size().reset_index(name="Jumlah")
                fig_b = px.bar(
                    df_t, 
                    x="Jam Absen", 
                    y="Jumlah", 
                    title="Tren Puncak Jam Kehadiran", 
                    color="Jumlah",
                    color_continuous_scale=['#81C784', '#1B5E20', '#C67D0A', '#B25900']
                )
                st.plotly_chart(fig_b, use_container_width=True)
    else:
        st.info("Dashboard interaktif akan otomatis aktif setelah ada presensi pertama masuk.")

# ------------------------------------------------------------
# HALAMAN 2: SCAN QR CODE
# ------------------------------------------------------------
elif menu_pilihan == "📱 Scan QR Code":
    st.header("📱 Scan QR Code untuk Absensi Digital")
    st.write("Tampilkan QR Code ini pada layar Zoom / Proyektor Aula agar peserta dapat memindai dari HP.")
    
    app_url = st.text_input("URL Link Absensi Ini:", value="https://absensi-badung.streamlit.app")
    qr_bytes = generate_qr_code(app_url)
    
    st.image(qr_bytes, caption="Pindai QR Code untuk Membuka Form Presensi", width=280)
    
    st.download_button(
        label="📥 Unduh Gambar QR Code (.png)",
        data=qr_bytes,
        file_name="QR_Code_Absensi_Badung.png",
        mime="image/png"
    )

# ------------------------------------------------------------
# HALAMAN 3: PANEL ADMIN
# ------------------------------------------------------------
elif menu_pilihan == "🔐 Panel Admin (Khusus Panitia)":
    st.header("🔐 Panel Otentikasi Admin")
    
    if not st.session_state.is_admin_logged_in:
        with st.form("admin_login_form"):
            pin_input = st.text_input("Masukkan PIN Keamanan Admin:", type="password")
            submit_login = st.form_submit_button("🔑 Login Admin", type="primary")
            
            if submit_login:
                if pin_input == ADMIN_PIN:
                    st.session_state.is_admin_logged_in = True
                    st.rerun()
                else:
                    st.error("❌ PIN Admin Salah!")
                    
    else:
        col_title, col_logout = st.columns([4, 1])
        with col_title:
            st.success("🔓 Akses Admin Diterima")
        with col_logout:
            if st.button("🚪 Logout Admin", use_container_width=True):
                st.session_state.is_admin_logged_in = False
                st.rerun()

        st.subheader("✏️ Pengaturan Nama Acara & Tab Google Sheets")
        
        col_ev1, col_ev2, col_ev3 = st.columns([2, 1, 1])
        with col_ev1:
            new_event_title = st.text_input("Judul Acara Saat Ini:", value=st.session_state.nama_event)
        with col_ev2:
            new_jenis_presensi = st.selectbox(
                "Pilih Jenis Presensi:", 
                ["Presensi Zoom", "Presensi Biasa (Offline / Aula)"],
                index=0 if st.session_state.jenis_presensi == "Presensi Zoom" else 1
            )
        with col_ev3:
            new_sheet_name = st.text_input("Nama Tab di Google Sheets:", value=st.session_state.target_sheet_name)
            
        if st.button("💾 Simpan Pengaturan Acara & Tab Google Sheets"):
            st.session_state.nama_event = new_event_title
            st.session_state.jenis_presensi = new_jenis_presensi
            st.session_state.target_sheet_name = new_sheet_name
            
            # Reload data dari tab baru
            try:
                st.session_state.df_absensi = load_data_from_gsheets()
                st.success(f"✅ Berhasil terhubung ke Tab Google Sheets: '{new_sheet_name}'!")
            except Exception as e:
                st.warning(f"⚠️ Tab '{new_sheet_name}' tidak ditemukan di Google Sheets. Pastikan Anda telah membuat tab tersebut di Google Sheets.")
                
            st.rerun()

        st.markdown("---")
        st.subheader("⚙️ Pengaturan Batas Waktu & Kuota Peserta")
        
        c_timer1, c_timer2 = st.columns(2)
        with c_timer1:
            set_enabled = st.checkbox("Aktifkan Batas Waktu (Jam)", value=st.session_state.cutoff_enabled)
        with c_timer2:
            set_time = st.time_input("Jam Absensi Ditutup:", value=st.session_state.cutoff_time)
            
        c_quota1, c_quota2 = st.columns(2)
        with c_quota1:
            set_max_enabled = st.checkbox("Aktifkan Batas Maksimal Kuota Peserta", value=st.session_state.max_peserta_enabled)
        with c_quota2:
            set_max_limit = st.number_input("Maksimal Kuota Peserta (Orang):", min_value=1, value=int(st.session_state.max_peserta_limit), step=1)
            
        if st.button("💾 Simpan Pengaturan Batas & Kuota"):
            st.session_state.cutoff_enabled = set_enabled
            st.session_state.cutoff_time = set_time
            st.session_state.max_peserta_enabled = set_max_enabled
            st.session_state.max_peserta_limit = set_max_limit
            st.success("✅ Pengaturan batas waktu & kuota berhasil disimpan!")
            st.rerun()
            
        st.markdown("---")
        st.subheader("📊 Rekapitulasi Data & Ekspor Laporan")
        df_res = st.session_state.df_absensi
        
        if not df_res.empty:
            view_cols = [c for c in df_res.columns if c != "_ttd_bytes"]
            df_view = df_res[view_cols]

            def style_table(row):
                if "Keterangan Status" in row and row["Keterangan Status"] == "DUPLIKAT / ABSEN GANDA":
                    return ['background-color: #ffc7ce; color: #9c0006; font-weight: bold;'] * len(row)
                return ['background-color: #c6efce; color: #006100;'] * len(row)
                
            st.dataframe(df_view.style.apply(style_table, axis=1), use_container_width=True)
            
            col_exp1, col_exp2 = st.columns(2)
            with col_exp1:
                excel_data = export_formatted_excel(df_res, st.session_state.nama_event)
                st.download_button(
                    label="📊 Download Rekap Excel Terformat (.xlsx)",
                    data=excel_data,
                    file_name=f"REKAP_PRESENSI_{st.session_state.nama_event.replace(' ', '_').upper()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col_exp2:
                st.download_button(
                    label="📄 Download Rekap Standard (.csv)",
                    data=df_view.to_csv(index=False).encode('utf-8'),
                    file_name="REKAP_PRESENSI_FULL.csv",
                    mime="text/csv"
                )
            
        else:
            st.info("Belum ada data presensi yang masuk pada tab ini.")