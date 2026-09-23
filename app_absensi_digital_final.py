import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import re
import datetime
import time
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
# KONFIGURASI HALAMAN & CUSTOM CSS MODERN (ENHANCED UI & CARD)
# ------------------------------------------------------------
st.set_page_config(
    page_title="Sistem Absensi Digital Badung", 
    page_icon="🏛️", 
    layout="wide"
)

st.markdown("""
    <style>
    /* 1. Latar Belakang Utama Aplikasi (Soft Light Grayish Green) */
    .stApp {
        background-color: #F4F6F4 !important;
    }

    /* Sembunyikan instruksi bawaan Streamlit */
    div[data-testid="InputInstructions"] {
        display: none !important;
    }

    /* 2. Banner Header Utama */
    .hero-header-banner {
        background-color: #0E3B2E;
        border-top: 3px solid #D4AF37;
        border-bottom: 3px solid #D4AF37;
        padding: 18px 30px;
        border-radius: 8px;
        box-shadow: 0 6px 18px rgba(0,0,0,0.15);
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

    /* 3. Card Agenda / Event Info */
    .agenda-card {
        background: #FFFFFF;
        border-left: 5px solid #1B5E20;
        padding: 14px 22px;
        border-radius: 8px;
        margin-bottom: 15px;
        box-shadow: 0 3px 12px rgba(0,0,0,0.05);
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
        padding: 5px 14px;
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

    /* 4. Broadcast Announcement Banner */
    .broadcast-banner {
        background: linear-gradient(135deg, #FFF8E1 0%, #FFECB3 100%);
        border: 1px solid #FFE082;
        border-left: 5px solid #FF8F00;
        padding: 12px 18px;
        border-radius: 8px;
        margin-bottom: 20px;
        font-size: 14px;
        color: #6D4C41;
        font-weight: 600;
        display: flex;
        align-items: center;
        gap: 10px;
    }

    /* 5. Floating Card Container Form Input Peserta */
    div[data-testid="stForm"] {
        background-color: #FFFFFF !important;
        border-radius: 12px !important;
        padding: 30px !important;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08) !important;
        border: 1px solid #E0E0E0 !important;
        margin-bottom: 25px !important;
    }

    /* 6. Menebalkan & Memperjelas Teks Label Input */
    div[data-testid="stWidgetLabel"] label {
        color: #0E3B2E !important;
        font-weight: 700 !important;
        font-size: 14px !important;
    }

    /* 7. Frame & Focus State Input Box */
    .stTextInput input, .stSelectbox div[data-baseweb="select"] {
        border: 1px solid #CCCCCC !important;
        border-radius: 8px !important;
        background-color: #FAFAFA !important;
    }

    .stTextInput input:focus, .stSelectbox div[data-baseweb="select"]:focus-within {
        border-color: #1B5E20 !important;
        box-shadow: 0 0 0 2px rgba(27, 94, 32, 0.2) !important;
        background-color: #FFFFFF !important;
    }

    h1 { color: #1B5E20 !important; }
    h2, h3 { color: #C67D0A !important; }
    div[data-testid="stMetricValue"] { color: #1B5E20 !important; font-weight: bold; }

    /* 8. Tombol Submit Modern */
    div[data-testid="stFormSubmitButton"] > button, div.stButton > button[kind="primary"] {
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
    div[data-testid="stFormSubmitButton"] > button:hover, div.stButton > button[kind="primary"]:hover {
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

if "pending_duplicate_data" not in st.session_state:
    st.session_state.pending_duplicate_data = None

if "preview_data" not in st.session_state:
    st.session_state.preview_data = None

# SESSION STATE UNTUK PENYIMPANAN DRAFT DATA INPUT PESERTA
if "draft_nama" not in st.session_state: st.session_state.draft_nama = ""
if "draft_nip" not in st.session_state: st.session_state.draft_nip = ""
if "draft_jabatan" not in st.session_state: st.session_state.draft_jabatan = ""
if "draft_opd" not in st.session_state: st.session_state.draft_opd = "Inspektorat"
if "draft_manual_opd" not in st.session_state: st.session_state.draft_manual_opd = ""
if "draft_email" not in st.session_state: st.session_state.draft_email = ""

# SESSION STATE UNTUK BROADCAST NOTICE & PIN ADMIN DYNAMIC
if "broadcast_text" not in st.session_state: st.session_state.broadcast_text = ""
if "admin_pin" not in st.session_state: st.session_state.admin_pin = "4869"

if "last_submit_time" not in st.session_state:
    st.session_state.last_submit_time = 0

LOCAL_CSV_FILE = Path(__file__).parent / "LOCAL_PRESENSI_BACKUP.csv"

# DISET KE 60 DETIK UNTUK MENGANTISIPASI TRAFIK MEMBLUDAK (> 2.000 PESERTA ZOOM)
@st.cache_data(ttl=60)
def fetch_gsheets_data(sheet_name):
    return conn.read(worksheet=sheet_name, ttl=0)

def load_data_from_gsheets():
    sheet_name = st.session_state.target_sheet_name
    try:
        df = fetch_gsheets_data(sheet_name)
        if LOCAL_CSV_FILE.exists():
            df_local = pd.read_csv(LOCAL_CSV_FILE, dtype=str)
            df = pd.concat([df, df_local], ignore_index=True).drop_duplicates(subset=["NIP", "Waktu Absen"], keep="last")
        return df
    except Exception:
        if LOCAL_CSV_FILE.exists():
            return pd.read_csv(LOCAL_CSV_FILE, dtype=str)
        return pd.DataFrame(columns=[
            "Waktu Absen", "Jam Absen", "Jenis Presensi", "Nama Peserta (EYD)", "NIP", "Jabatan (EYD)", "OPD / Instansi", "Email", "Tanda Tangan", "Keterangan Status"
        ])

def save_data_to_gsheets(new_row_df, existing_df):
    sheet_name = st.session_state.target_sheet_name
    cols_to_save = [c for c in existing_df.columns if c != "_ttd_bytes"]
    clean_existing = existing_df[cols_to_save] if not existing_df.empty else pd.DataFrame(columns=cols_to_save)
    clean_new = new_row_df[cols_to_save]
    
    updated_df = pd.concat([clean_existing, clean_new], ignore_index=True)
    
    try:
        clean_new.to_csv(LOCAL_CSV_FILE, mode='a', header=not LOCAL_CSV_FILE.exists(), index=False)
    except Exception:
        pass
        
    try:
        conn.update(worksheet=sheet_name, data=updated_df)
    except Exception:
        st.warning("⚠️ Koneksi Google Sheets sedang padat. Data Anda diselamatkan di cadangan lokal server.")
        
    st.cache_data.clear()

def replace_duplicate_data(df_existing, nip_target, email_target, new_row_dict):
    nip_c = clean_cmp(nip_target)
    em_c = clean_cmp(email_target)
    
    def is_match(row):
        r_nip = clean_cmp(row.get("NIP", ""))
        r_em = clean_cmp(row.get("Email", ""))
        match_nip = (len(nip_c) >= 18 and r_nip == nip_c)
        match_em = (em_c != "-" and em_c != "" and r_em == em_c)
        return match_nip or match_em

    df_filtered = df_existing[~df_existing.apply(is_match, axis=1)]
    new_df = pd.DataFrame([new_row_dict])
    df_updated = pd.concat([df_filtered, new_df], ignore_index=True)
    
    try:
        conn.update(worksheet=st.session_state.target_sheet_name, data=df_updated)
    except Exception:
        st.warning("⚠️ Gagal memperbarui Google Sheets. Menggunakan cadangan lokal.")
        
    st.cache_data.clear()
    return df_updated

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

# BROADCAST BANNER NOTICE
if st.session_state.broadcast_text.strip():
    st.markdown(f"""
        <div class="broadcast-banner">
            <span style="font-size: 18px;">📢</span>
            <div><strong>Pengumuman Panitia:</strong> {st.session_state.broadcast_text}</div>
        </div>
    """, unsafe_allow_html=True)

# ------------------------------------------------------------
# NAVIGASI SIDEBAR
# ------------------------------------------------------------
st.sidebar.title("📌 Menu Navigasi")
menu_pilihan = st.sidebar.radio(
    "Pilih Halaman:",
    ["📝 Form Absensi Peserta", "📱 Scan QR Code", "🔐 Panel Admin (Khusus Panitia)"]
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
# HALAMAN 1: FORM ABSENSI PESERTA (MURNI TANPA DASHBOARD PUBLIC)
# ------------------------------------------------------------
if menu_pilihan in ["📝 Form Absensi Peserta", "📝 Form Absensi & Dashboard"]:
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
        
        # 1. KOTAK KONFIRMASI DATA DUPLIKAT
        if st.session_state.pending_duplicate_data is not None:
            pending_info = st.session_state.pending_duplicate_data
            old_rec = pending_info["old_record"]
            new_rec = pending_info["new_row"]
            
            st.warning("⚠️ **PERHATIAN: DATA PRESENSI ANDA TERDETEKSI SUDAH ADA!**")
            
            st.markdown(f"""
            Sistem menemukan bahwa NIP / Email Anda sudah terdaftar pada:
            * **Waktu Presensi Awal**: `{old_rec.get('Waktu Absen', '-')}`
            * **Nama Terdaftar**: `{old_rec.get('Nama Peserta (EYD)', '-')}`
            * **OPD / Instansi**: `{old_rec.get('OPD / Instansi', '-')}`

            Apakah Anda ingin **menghapus data presensi lama** dan menggantinya dengan data baru ini?
            """)
            
            col_confirm1, col_confirm2 = st.columns(2)
            
            with col_confirm1:
                if st.button("🔴 Ya, Hapus Data Lama & Perbarui", type="primary", use_container_width=True):
                    with st.spinner("Menghapus data lama & memperbarui presensi Anda..."):
                        updated_df = replace_duplicate_data(
                            pending_info["df_ex"], 
                            new_rec["NIP"], 
                            new_rec["Email"], 
                            new_rec
                        )
                        st.session_state.df_absensi = updated_df
                        st.session_state.pending_duplicate_data = None
                        st.session_state.draft_nama = ""
                        st.session_state.draft_nip = ""
                        st.session_state.draft_jabatan = ""
                        st.session_state.draft_email = ""
                        st.session_state.draft_manual_opd = ""
                    st.success("✅ Data presensi lama berhasil dihapus dan diperbarui!")
                    st.rerun()
                    
            with col_confirm2:
                if st.button("⚪ Batal (Simpan Data Yang Sudah Ada)", use_container_width=True):
                    st.session_state.pending_duplicate_data = None
                    st.info("Proses dibatalkan. Data presensi pertama Anda tetap aman tersimpan.")
                    st.rerun()

        # 2. LAYAR PRATINJAU / PREVIEW DATA PESERTA
        elif st.session_state.preview_data is not None:
            p_data = st.session_state.preview_data
            
            st.markdown("""
                <style>
                .preview-card {
                    background: #FFFFFF;
                    border-radius: 14px;
                    padding: 28px 32px;
                    border: 1px solid #E0E0E0;
                    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.06);
                    margin-bottom: 25px;
                }
                .preview-header {
                    display: flex;
                    align-items: center;
                    gap: 12px;
                    padding-bottom: 16px;
                    border-bottom: 2px solid #E29D29;
                    margin-bottom: 20px;
                }
                .preview-header-title {
                    font-size: 18px;
                    font-weight: 800;
                    color: #0E3B2E;
                    letter-spacing: 0.5px;
                }
                .preview-grid {
                    display: grid;
                    grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
                    gap: 18px;
                }
                .preview-item {
                    background-color: #F8F9FA;
                    padding: 12px 16px;
                    border-radius: 8px;
                    border-left: 4px solid #1B5E20;
                }
                .preview-label {
                    font-size: 12px;
                    font-weight: 700;
                    color: #666666;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                    margin-bottom: 4px;
                }
                .preview-value {
                    font-size: 15px;
                    font-weight: 700;
                    color: #111111;
                    word-break: break-word;
                }
                </style>
            """, unsafe_allow_html=True)

            nip_display = p_data['NIP'] if p_data['NIP'] else "-"
            
            st.markdown(f"""
                <div class="preview-card">
                    <div class="preview-header">
                        <span style="font-size: 24px;">📋</span>
                        <div class="preview-header-title">KONFIRMASI DATA PRESENSI PESERTA</div>
                    </div>
                    <div style="font-size: 13px; color: #555; margin-bottom: 20px;">
                        Silakan periksa kembali kebenaran data Anda sebelum dikirim ke sistem:
                    </div>
                    <div class="preview-grid">
                        <div class="preview-item">
                            <div class="preview-label">👤 Nama Lengkap & Gelar</div>
                            <div class="preview-value">{p_data['Nama Peserta (EYD)']}</div>
                        </div>
                        <div class="preview-item">
                            <div class="preview-label">🪪 NIP Peserta</div>
                            <div class="preview-value">{nip_display}</div>
                        </div>
                        <div class="preview-item">
                            <div class="preview-label">💼 Jabatan</div>
                            <div class="preview-value">{p_data['Jabatan (EYD)']}</div>
                        </div>
                        <div class="preview-item">
                            <div class="preview-label">🏛️ OPD / Instansi</div>
                            <div class="preview-value">{p_data['OPD / Instansi']}</div>
                        </div>
                        <div class="preview-item" style="grid-column: 1 / -1;">
                            <div class="preview-label">✉️ Alamat Email</div>
                            <div class="preview-value">{p_data['Email']}</div>
                        </div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
            col_prev1, col_prev2 = st.columns(2)
            
            with col_prev1:
                if st.button("✏️ Ubah / Perbaiki Data", use_container_width=True):
                    st.session_state.preview_data = None
                    st.rerun()
                    
            with col_prev2:
                if st.button("✅ Sudah Benar & Kirim Presensi", type="primary", use_container_width=True):
                    df_ex = load_data_from_gsheets()
                    duplicate_found = False
                    old_record = None

                    if not df_ex.empty and "NIP" in df_ex.columns:
                        nip_c = clean_cmp(p_data["NIP"])
                        em_c = clean_cmp(p_data["Email"])
                        
                        for idx, r in df_ex.iterrows():
                            r_nip = clean_cmp(r.get("NIP", ""))
                            r_em = clean_cmp(r.get("Email", ""))
                            
                            if (len(nip_c) >= 18 and r_nip == nip_c) or (em_c != "-" and em_c != "" and r_em == em_c):
                                duplicate_found = True
                                old_record = r
                                break

                    if duplicate_found:
                        st.session_state.pending_duplicate_data = {
                            "new_row": p_data,
                            "old_record": old_record,
                            "df_ex": df_ex
                        }
                        st.session_state.preview_data = None
                        st.rerun()
                    else:
                        new_row_df = pd.DataFrame([p_data])
                        with st.spinner("Menyimpan presensi ke database..."):
                            save_data_to_gsheets(new_row_df, df_ex)
                        st.success(f"✅ Presensi Berhasil Disimpan! Terima kasih, **{p_data['Nama Peserta (EYD)']}**.")
                        st.session_state.preview_data = None
                        st.session_state.draft_nama = ""
                        st.session_state.draft_nip = ""
                        st.session_state.draft_jabatan = ""
                        st.session_state.draft_email = ""
                        st.session_state.draft_manual_opd = ""
                        st.rerun()

        # 3. FORM ISIAN UTAMA
        else:
            with st.form("form_presensi_peserta"):
                st.subheader(f"Form Presensi Kehadiran — {st.session_state.nama_event}")
                st.markdown("---")
                
                col1, col2 = st.columns(2)
                with col1:
                    in_nama = st.text_input("Nama Lengkap & Gelar (EYD):", value=st.session_state.draft_nama, placeholder="Contoh: aditya putra, se")
                    in_nip = st.text_input("NIP (18 Digit Angka):", value=st.session_state.draft_nip, placeholder=st.session_state.random_nip_placeholder)
                    in_jabatan = st.text_input("Jabatan (EYD):", value=st.session_state.draft_jabatan, placeholder="Contoh: analis kebijakan ahli muda")
                
                with col2:
                    default_opd_idx = LIST_OPD.index(st.session_state.draft_opd) if st.session_state.draft_opd in LIST_OPD else 0
                    in_opd = st.selectbox("OPD / Instansi Peserta:", LIST_OPD, index=default_opd_idx)
                    
                    manual_opd = st.text_input(
                        "Tuliskan Nama Instansi / OPD Anda (Khusus jika memilih 'Lainnya / Instansi luar'):",
                        value=st.session_state.draft_manual_opd,
                        placeholder="Contoh: Kementerian Hukum dan HAM / Universitas Udayana"
                    )
                    
                    if in_opd == "Lainnya / Instansi luar":
                        in_opd_final = format_nama_instansi(manual_opd) if manual_opd.strip() else "Instansi Lainnya"
                    else:
                        in_opd_final = in_opd
                    
                    if is_zoom_mode:
                        in_email = st.text_input("Alamat Email Valid:", value=st.session_state.draft_email, placeholder="Contoh: peserta@gmail.com")
                    else:
                        in_email = "-"

                ttd_status = "Tanda Tangan Daring (Zoom)"
                
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
                    
                    if canvas_result.image_data is not None:
                        img_array = canvas_result.image_data.astype(np.uint8)
                        if np.any(img_array[:, :, 3] > 0):
                            ttd_status = "Ada TTD Digital"

                submit_button = st.form_submit_button("👁️ Pratinjau & Periksa Data")

            if submit_button:
                current_time = time.time()
                if current_time - st.session_state.last_submit_time < 3:
                    st.warning("⏳ Sistem sedang memproses antrean. Harap tunggu 3 detik.")
                else:
                    st.session_state.last_submit_time = current_time
                    
                    st.session_state.draft_nama = in_nama
                    st.session_state.draft_nip = in_nip
                    st.session_state.draft_jabatan = in_jabatan
                    st.session_state.draft_opd = in_opd
                    st.session_state.draft_manual_opd = manual_opd
                    st.session_state.draft_email = in_email
                    
                    nip_valid, nip_msg = validate_nip(in_nip)
                    email_valid, email_clean, email_msg = (validate_email(in_email) if is_zoom_mode else (True, "-", ""))
                    
                    # FITUR PERINGATAN DINI OPD (OPSI 3)
                    if in_opd != "Lainnya / Instansi luar" and manual_opd.strip() != "":
                        st.warning(f"⚠️ **Peringatan Pilihan OPD**: Anda memilih OPD **'{in_opd}'**, tetapi juga mengisikan teks **'{manual_opd}'**. Sistem akan mencatat Anda sebagai peserta dari **'{in_opd}'**. Jika Anda berasal dari instansi luar, silakan ubah pilihan OPD di atas menjadi **'Lainnya / Instansi luar'**.")
                    
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
                        
                        st.session_state.preview_data = {
                            "Waktu Absen": waktu_now.strftime("%Y-%m-%d %H:%M:%S"),
                            "Jam Absen": waktu_now.strftime("%H:%M"),
                            "Jenis Presensi": st.session_state.jenis_presensi,
                            "Nama Peserta (EYD)": nama_eyd,
                            "NIP": re.sub(r'\D', '', in_nip.strip()),
                            "Jabatan (EYD)": jabatan_eyd,
                            "OPD / Instansi": in_opd_final,
                            "Email": email_clean,
                            "Tanda Tangan": ttd_status,
                            "Keterangan Status": "UTAMA / VALID"
                        }
                        st.rerun()

# ------------------------------------------------------------
# HALAMAN 2: SCAN QR CODE
# ------------------------------------------------------------
elif menu_pilihan == "📱 Scan QR Code":
    st.header("📱 Scan QR Code untuk Absensi Digital")
    st.write("Tampilkan QR Code ini pada layar Zoom / Proyektor Aula agar peserta dapat memindai dari HP.")
    
    app_url = st.text_input("URL Link Absensi Ini:", value="https://absensi-digital-badung.streamlit.app")
    qr_bytes = generate_qr_code(app_url)
    
    st.image(qr_bytes, caption="Pindai QR Code untuk Membuka Form Presensi", width=280)
    
    st.download_button(
        label="📥 Unduh Gambar QR Code (.png)",
        data=qr_bytes,
        file_name="QR_Code_Absensi_Badung.png",
        mime="image/png"
    )

# ------------------------------------------------------------
# HALAMAN 3: PANEL ADMIN (TERMASUK DASHBOARD MONITORING EXCLUSIVE)
# ------------------------------------------------------------
elif menu_pilihan == "🔐 Panel Admin (Khusus Panitia)":
    st.header("🔐 Panel Otentikasi Admin")
    
    if not st.session_state.is_admin_logged_in:
        with st.form("admin_login_form"):
            pin_input = st.text_input("Masukkan PIN Keamanan Admin:", type="password")
            submit_login = st.form_submit_button("🔑 Login Admin")
            
            if submit_login:
                if pin_input == st.session_state.admin_pin:
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

        # DASHBOARD MONITORING REAL-TIME (DIPINDAHKAN KHUSUS KE PANEL ADMIN)
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

        st.markdown("---")

        # 1. PENGATURAN ACARA & BROADCAST NOTICE
        st.subheader("✏️ Pengaturan Acara & Pengumuman Broadcast Peserta")
        
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
            
        new_broadcast = st.text_input(
            "📢 Teks Pengumuman Broadcast (Muncul di Atas Form Peserta):",
            value=st.session_state.broadcast_text,
            placeholder="Contoh: Presensi diperpanjang hingga pukul 13.00 WITA. Harap pastikan Email Anda valid."
        )

        if st.button("💾 Simpan Pengaturan Acara & Pengumuman", type="primary"):
            st.session_state.nama_event = new_event_title
            st.session_state.jenis_presensi = new_jenis_presensi
            st.session_state.target_sheet_name = new_sheet_name
            st.session_state.broadcast_text = new_broadcast
            
            try:
                st.session_state.df_absensi = load_data_from_gsheets()
                st.success(f"✅ Berhasil diperbarui! Connected to tab '{new_sheet_name}'.")
            except Exception:
                st.warning(f"⚠️ Tab '{new_sheet_name}' tidak ditemukan di Google Sheets.")
                
            st.rerun()

        st.markdown("---")
        
        # 2. PENGATURAN PIN ADMIN DYNAMIC
        col_pin1, col_pin2 = st.columns(2)
        with col_pin1:
            st.subheader("🔑 Ubah PIN Keamanan Admin")
            new_pin_val = st.text_input("PIN Baru Admin (Min 4 Karakter):", type="password", placeholder="Ketik PIN baru...")
            if st.button("💾 Perbarui PIN Admin"):
                if len(new_pin_val.strip()) >= 4:
                    st.session_state.admin_pin = new_pin_val.strip()
                    st.success("✅ PIN Admin berhasil diperbarui!")
                else:
                    st.error("⚠️ PIN minimal harus 4 karakter!")

        with col_pin2:
            st.subheader("🔄 Refresh Data Real-Time")
            st.caption("Gunakan tombol ini untuk memaksa pembacaan data terbaru jika caching sedang aktif:")
            if st.button("🔄 Paksa Perbarui Data (Clear Cache)", use_container_width=True):
                st.cache_data.clear()
                st.session_state.df_absensi = load_data_from_gsheets()
                st.success("✅ Cache dibersihkan! Data terbaru berhasil dimuat.")
                st.rerun()

        st.markdown("---")
        # 3. PENGATURAN BATAS WAKTU & KUOTA
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
            
        if st.button("💾 Simpan Pengaturan Batas & Kuota", type="primary"):
            st.session_state.cutoff_enabled = set_enabled
            st.session_state.cutoff_time = set_time
            st.session_state.max_peserta_enabled = set_max_enabled
            st.session_state.max_peserta_limit = set_max_limit
            st.success("✅ Pengaturan batas waktu & kuota berhasil disimpan!")
            st.rerun()
            
        st.markdown("---")
        # 4. REKAPITULASI & FITUR PENCARIAN PESERTA
        st.subheader("📊 Rekapitulasi Data Peserta & Pencarian Instant")
        df_res = st.session_state.df_absensi
        
        if not df_res.empty:
            view_cols = [c for c in df_res.columns if c != "_ttd_bytes"]
            df_view = df_res[view_cols]

            search_query = st.text_input("🔍 Cari Peserta Instant (Nama / NIP / OPD):", placeholder="Ketik Nama, NIP, atau OPD...")
            if search_query.strip():
                df_view = df_view[
                    df_view["Nama Peserta (EYD)"].astype(str).str.contains(search_query, case=False, na=False) |
                    df_view["NIP"].astype(str).str.contains(search_query, case=False, na=False) |
                    df_view["OPD / Instansi"].astype(str).str.contains(search_query, case=False, na=False)
                ]

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
