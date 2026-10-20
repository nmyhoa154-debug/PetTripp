from pathlib import Path
from urllib.parse import quote
from datetime import date, datetime
import re
import sqlite3
import json
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "data" / "pettrip.db"
DB_PATH.parent.mkdir(parents=True, exist_ok=True)
def khoi_tao_database():
    conn = sqlite3.connect(DB_PATH)

    try:
        cursor = conn.cursor()

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS chuyen_di (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ten_chuyen_di TEXT,
                diem_den TEXT NOT NULL,
                loai_pet TEXT,
                size_pet TEXT,
                ten_pet TEXT,
                ngay_di TEXT,
                ngay_ve TEXT,
                ngan_sach REAL,
                uu_tien TEXT,
                ghi_chu TEXT,
                ngay_tao TEXT,
                noi_dung TEXT
            )
        """)

        conn.commit()
    finally:
        conn.close()


khoi_tao_database()


# =====================================================
# 1. CẤU HÌNH
# =====================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "data" / "PetTrip.xlsx"
LOGO_FILE = BASE_DIR / "assets" / "logo.png"
DB_FILE = DB_PATH

NAVY = "#18324B"
NAVY_LIGHT = "#294D6B"
CREAM = "#F6F0E5"
CREAM_LIGHT = "#FFFCF6"
BEIGE = "#E7D8C2"
BROWN = "#8A6D4B"

st.set_page_config(
    page_title="PetTrip | Du lịch cùng thú cưng",
    page_icon="🐾",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =====================================================
# 2. GIAO DIỆN KEM - XANH NAVY
# =====================================================

st.markdown(
    f"""
    <style>
    .stApp {{
        background-color: {CREAM};
        color: {NAVY};
    }}

    [data-testid="stSidebar"] {{
        background-color: #EDE2D0;
        border-right: 1px solid {BEIGE};
    }}

    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] label {{
        color: {NAVY};
    }}

    h1, h2, h3, h4, p, label {{
        color: {NAVY};
    }}

    [data-testid="stMetric"] {{
        background: {CREAM_LIGHT};
        border: 1px solid {BEIGE};
        padding: 16px;
        border-radius: 14px;
    }}

    [data-testid="stMetricLabel"],
    [data-testid="stMetricValue"] {{
        color: {NAVY};
    }}

    div[data-testid="stVerticalBlockBorderWrapper"] {{
        background: {CREAM_LIGHT};
        border-color: {BEIGE} !important;
        border-radius: 16px;
    }}

    .hero {{
        background: linear-gradient(120deg, #E8D9C1, #F8F1E6);
        border: 1px solid {BEIGE};
        border-radius: 22px;
        padding: 34px;
        margin-bottom: 24px;
    }}

    .hero h1 {{
        color: {NAVY};
        font-size: 34px;
        margin-bottom: 12px;
    }}

    .hero p {{
        color: {NAVY};
        font-size: 17px;
        line-height: 1.8;
    }}

    .section-note {{
        color: {NAVY_LIGHT};
        font-size: 16px;
        line-height: 1.7;
    }}

    .pet-panel {{
        background: {CREAM_LIGHT};
        border: 1px solid {BEIGE};
        padding: 20px;
        border-radius: 16px;
        margin: 10px 0 20px 0;
    }}

    .pet-panel h3 {{
        color: {NAVY};
        margin-top: 0;
    }}

    .pet-panel p {{
        color: {NAVY_LIGHT};
    }}

    .place-name {{
        color: {NAVY};
        font-size: 21px;
        font-weight: 700;
        margin-bottom: 10px;
    }}

    .place-tag {{
        background: #EDE2D0;
        color: {NAVY};
        border-radius: 8px;
        padding: 5px 9px;
        display: inline-block;
        margin: 3px 3px 3px 0;
        font-size: 13px;
    }}

    .footer {{
        text-align: center;
        color: {NAVY_LIGHT};
        padding: 20px 0;
    }}

    .stButton > button,
    .stLinkButton > a,
    .stDownloadButton > button {{
        background-color: {NAVY} !important;
        color: #FFF8EC !important;
        border: 1px solid {NAVY} !important;
        border-radius: 10px;
        font-weight: 600;
    }}

    .stButton > button:hover,
    .stLinkButton > a:hover,
    .stDownloadButton > button:hover {{
        background-color: {NAVY_LIGHT} !important;
        color: #FFFFFF !important;
        border-color: {NAVY_LIGHT} !important;
    }}

    .stButton > button p,
    .stLinkButton > a p,
    .stDownloadButton > button p,
    .stButton > button span,
    .stLinkButton > a span,
    .stDownloadButton > button span {{
        color: #FFF8EC !important;
    }}

    input, textarea {{
        background-color: {CREAM_LIGHT} !important;
        color: {NAVY} !important;
    }}

    [data-baseweb="select"] > div {{
        background-color: {CREAM_LIGHT};
        color: {NAVY};
        border-color: {BEIGE};
    }}

    hr {{
        border-color: {BEIGE};
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# =====================================================
# 3. ĐỌC FILE EXCEL
# =====================================================

@st.cache_data
def load_excel(file_path, modified_time):
    excel = pd.ExcelFile(file_path)
    return {
        sheet: pd.read_excel(file_path, sheet_name=sheet)
        for sheet in excel.sheet_names
    }


def get_data():
    if not DATA_FILE.exists():
        return None, (
            "Không tìm thấy data/PetTrip.xlsx. "
            "Hãy kiểm tra vị trí file Excel."
        )

    try:
        return load_excel(
            str(DATA_FILE),
            DATA_FILE.stat().st_mtime
        ), None
    except Exception as error:
        return None, f"Không đọc được file Excel: {error}"


du_lieu, loi_du_lieu = get_data()


# =====================================================
# 4. HÀM XỬ LÝ DỮ LIỆU
# =====================================================

def hop_le(value):
    if value is None:
        return False

    try:
        if pd.isna(value):
            return False
    except (TypeError, ValueError):
        pass

    return str(value).strip().casefold() not in [
        "", "nan", "none", "nat", "null"
    ]


def hien_thi(value, mac_dinh="Chưa cập nhật"):
    if hop_le(value):
        return str(value).strip()
    return mac_dinh


def tim_cot(df, cac_ten):
    columns = list(df.columns)

    for col in columns:
        ten = str(col).strip().casefold()
        for wanted in cac_ten:
            if ten == wanted.casefold():
                return col

    for col in columns:
        ten = str(col).strip().casefold()
        for wanted in cac_ten:
            if len(wanted) >= 4 and wanted.casefold() in ten:
                return col

    return None


def thong_tin_cot(df):
    return {
        "id": tim_cot(df, ["ID"]),
        "ten": tim_cot(df, ["Tên", "Tên địa điểm", "Tên cơ sở"]),
        "thanh_pho": tim_cot(df, ["Thành phố"]),
        "dia_chi": tim_cot(df, ["Địa chỉ"]),
        "gia_min": tim_cot(
            df, ["GiaThapNhat", "GiaTu", "Giá thấp nhất"]
        ),
        "gia_max": tim_cot(
            df, ["GiaCaoNhat", "GiáCaoNhất", "GiaDen", "Giá cao nhất"]
        ),
        "loai": tim_cot(df, ["Giống loài"]),
        "size": tim_cot(df, ["Size (kg)", "Size"]),
        "sdt": tim_cot(df, ["SDT", "Số điện thoại"]),
        "gio": tim_cot(
            df, ["GioHoatDong", "GioMoCua", "Giờ hoạt động"]
        ),
        "dich_vu": tim_cot(df, ["DichVu", "Dịch vụ"]),
        "cap_cuu": tim_cot(df, ["CapCuu", "Cấp cứu"]),
    }


def link_map(ten, dia_chi="", thanh_pho=""):
    query = " ".join(
        part for part in [ten, dia_chi, thanh_pho]
        if hop_le(part)
    )
    return (
        "https://www.google.com/maps/search/?api=1&query="
        + quote(query)
    )


def chuan_hoa(value):
    return re.sub(r"\s+", " ", str(value).strip().casefold())


def loai_phu_hop(value, loai_chon):
    text = chuan_hoa(value)

    if not text or text in ["nan", "none"]:
        return False

    if "không công bố" in text or "chưa rõ" in text:
        return False

    if loai_chon == "Chó":
        return "chó" in text or "cho" in text

    return "mèo" in text or "meo" in text


def size_phu_hop(value, size_chon):
    """
    So khớp nhãn size trong Excel.
    Không tự suy luận size nếu dữ liệu không ghi rõ.
    """
    text = chuan_hoa(value)

    if not text or text in ["nan", "none"]:
        return False

    if "không giới hạn" in text:
        return True

    pattern = rf"(?<![a-z]){re.escape(size_chon)}\s*\("
    return re.search(pattern, text, flags=re.IGNORECASE) is not None


def la_sheet_thu_y(ten_sheet):
    ten = ten_sheet.casefold()
    return "thú y" in ten or "thu y" in ten


def la_sheet_khach_san(ten_sheet):
    ten = ten_sheet.casefold()
    return "khách sạn" in ten or "khach san" in ten


def ten_danh_muc_hien_thi(ten_sheet, index=None):
    ten = chuan_hoa(ten_sheet)

    danh_muc = [
        ("khách sạn", "Khách sạn"),
        ("khach san", "Khách sạn"),
        ("nhà hàng", "Nhà hàng"),
        ("nha hang", "Nhà hàng"),
        ("khu vui chơi", "Khu vui chơi"),
        ("cafe", "Cà phê"),
        ("cà phê", "Cà phê"),
        ("ca phe", "Cà phê"),
        ("thú y", "Thú y"),
        ("thu y", "Thú y"),
    ]

    for tu_khoa, ten_hien_thi in danh_muc:
        if tu_khoa in ten:
            return ten_hien_thi

    ten_mac_dinh = [
        "Khách sạn",
        "Nhà hàng",
        "Khu vui chơi",
        "Cà phê",
        "Thú y",
    ]

    if (
        re.fullmatch(r"sheet\s*\d+", ten, flags=re.IGNORECASE)
        and index is not None
        and index < len(ten_mac_dinh)
    ):
        return ten_mac_dinh[index]

    return (
        "Địa điểm " + str(index + 1)
        if index is not None
        else "Địa điểm"
    )


# =====================================================
# 5. CƠ SỞ DỮ LIỆU SQLITE
# =====================================================

def khoi_tao_database():
    DB_FILE.parent.mkdir(parents=True, exist_ok=True)

    with sqlite3.connect(DB_FILE) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS checklist (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                loai_pet TEXT NOT NULL,
                ten_viec TEXT NOT NULL,
                da_tich INTEGER DEFAULT 0,
                UNIQUE(loai_pet, ten_viec)
            )
        """)

        conn.execute("""
            CREATE TABLE IF NOT EXISTS chuyen_di (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ten_chuyen_di TEXT,
                diem_den TEXT NOT NULL,
                loai_pet TEXT,
                size_pet TEXT,
                ten_pet TEXT,
                ngay_di TEXT,
                ngay_ve TEXT,
                ngan_sach REAL,
                uu_tien TEXT,
                ghi_chu TEXT,
                ngay_tao TEXT,
                noi_dung TEXT
            )
        """)
        # Bổ sung cột cho database cũ nếu trước đây đã tạo bảng thiếu cột.
        columns = {
            row[1] for row in conn.execute("PRAGMA table_info(chuyen_di)").fetchall()
        }
        if "noi_dung" not in columns:
            conn.execute("ALTER TABLE chuyen_di ADD COLUMN noi_dung TEXT")

        conn.commit()


def doc_checklist(loai_pet):
    with sqlite3.connect(DB_FILE) as conn:
        rows = conn.execute(
            """
            SELECT ten_viec, da_tich
            FROM checklist
            WHERE loai_pet = ?
            ORDER BY id
            """,
            (loai_pet,)
        ).fetchall()

    return {ten: bool(da_tich) for ten, da_tich in rows}


def luu_checklist(loai_pet, ten_viec, da_tich):
    with sqlite3.connect(DB_FILE) as conn:
        conn.execute(
            """
            INSERT INTO checklist (loai_pet, ten_viec, da_tich)
            VALUES (?, ?, ?)
            ON CONFLICT(loai_pet, ten_viec)
            DO UPDATE SET da_tich = excluded.da_tich
            """,
            (loai_pet, ten_viec, int(da_tich))
        )
        conn.commit()


def dat_lai_checklist(loai_pet):
    with sqlite3.connect(DB_FILE) as conn:
        conn.execute(
            "DELETE FROM checklist WHERE loai_pet = ?",
            (loai_pet,)
        )
        conn.commit()



def luu_chuyen_di(ten, diem_den, ngay_di, ngay_ve, noi_dung):
    # Luôn ghi vào cùng database mà trang "Chuyến đi đã lưu" đọc.
    conn = sqlite3.connect(DB_FILE)

    try:
        cursor = conn.cursor()

        cursor.execute(
            """
            INSERT INTO chuyen_di (
                ten_chuyen_di,
                diem_den,
                ngay_di,
                ngay_ve,
                loai_pet,
                size_pet,
                ten_pet,
                ngan_sach,
                ghi_chu,
                ngay_tao,
                noi_dung
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                ten,
                diem_den,
                ngay_di,
                ngay_ve,
                noi_dung.get("thong_tin", {}).get("loai_pet", ""),
                noi_dung.get("thong_tin", {}).get("size_pet", ""),
                noi_dung.get("thong_tin", {}).get("ten_pet", ""),
                noi_dung.get("thong_tin", {}).get("ngan_sach", 0),
                noi_dung.get("thong_tin", {}).get("ghi_chu", ""),
                datetime.now().isoformat(timespec="seconds"),
                json.dumps(noi_dung, ensure_ascii=False),
            ),
        )

        conn.commit()

    finally:
        conn.close()



def doc_chuyen_di():
    """Đọc các chuyến đi đã lưu từ đúng database của ứng dụng."""
    with sqlite3.connect(DB_FILE) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT * FROM chuyen_di ORDER BY id DESC"
        ).fetchall()

    result = []
    for row in rows:
        item = dict(row)
        try:
            noi_dung = item.get("noi_dung") or "{}"
            item["noi_dung"] = json.loads(noi_dung) if isinstance(noi_dung, str) else noi_dung
        except (TypeError, json.JSONDecodeError):
            item["noi_dung"] = {}
        item["ten"] = item.get("ten_chuyen_di") or "Chuyến đi"
        result.append(item)
    return result


def xoa_chuyen_di(chuyen_di_id):
    with sqlite3.connect(DB_FILE) as conn:
        conn.execute(
            "DELETE FROM chuyen_di WHERE id = ?",
            (int(chuyen_di_id),)
        )
        conn.commit()


khoi_tao_database()
# =====================================================
# 6. HỒ SƠ THÚ CƯNG
# =====================================================

def hien_thi_ho_so():
    st.markdown(
        """
        <div class="pet-panel">
            <h3>🐾 Hồ sơ người bạn đồng hành</h3>
            <p>
                Hãy cho PetTrip biết về bé để tìm địa điểm phù hợp
                với loài và kích thước của bé.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)

    with col1:
        loai_pet = st.selectbox(
            "🐶🐱 Thú cưng của bạn",
            ["Chó", "Mèo"],
            key="ho_so_loai",
        )

    with col2:
        size_pet = st.selectbox(
            "📏 Kích thước của bé",
            ["S", "M", "L", "XL"],
            format_func=lambda x: {
                "S": "S (1-5kg)",
                "M": "M (5-10kg)",
                "L": "L (10-15kg)",
                "XL": "XL (15kg trở lên)",
            }[x],
            key="ho_so_size",
        )

    ten_pet = st.text_input(
        "💗 Tên thú cưng",
        placeholder="Ví dụ: Bông, Milo, Mochi...",
        key="ho_so_ten",
    )

    return loai_pet, size_pet, ten_pet


# =====================================================
# 7. THANH ĐIỀU HƯỚNG
# =====================================================

with st.sidebar:
    if LOGO_FILE.exists():
        st.image(str(LOGO_FILE), width=180)
    else:
        st.markdown(
            f"<h1 style='text-align:center;color:{NAVY}'>🐾</h1>",
            unsafe_allow_html=True,
        )

    st.markdown(
        f"<h2 style='text-align:center;color:{NAVY}'>PetTrip</h2>",
        unsafe_allow_html=True,
    )
    st.caption("Nền tảng du lịch cùng thú cưng")
    st.divider()

    trang = st.radio(
        "KHÁM PHÁ PETTRIP",
        [
            "🏠 Trang chủ",
            "🔎 Khám phá địa điểm",
            "🐾 Hồ sơ thú cưng",
            "🗺️ Lập kế hoạch chuyến đi",
            "🧳 Chuyến đi đã lưu",
            "✅ Checklist chuẩn bị",
        ],
    )

    st.divider()
    st.caption("Hành trình vui vẻ bắt đầu từ đây!")


# =====================================================
# 8. TRANG CHỦ
# =====================================================

if trang == "🏠 Trang chủ":
    st.markdown(
        """
        <div class="hero">
            <h1>🐾 PetTrip</h1>
            <p>
                Nền tảng du lịch cùng thú cưng giúp dễ dàng tìm kiếm, so sánh và lựa chọn địa điểm phù hợp với từng thú cưng, nhu cầu và ngân sách dựa trên đánh giá thực tế từ cộng đồng.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if loi_du_lieu:
        st.error(loi_du_lieu)

    elif du_lieu:
        tong_dia_diem = sum(len(df) for df in du_lieu.values())

        st.markdown("### Khám phá PetTrip")

        c1, c2, c3 = st.columns(3)

        with c1:
            st.metric("Danh mục địa điểm", len(du_lieu))

        with c2:
            st.metric("Địa điểm trong dữ liệu", tong_dia_diem)

        with c3:
            st.metric("Người bạn đồng hành", "Chó & Mèo")

        st.markdown(
            """
            <div class="pet-panel">
                <h3>✨ Bắt đầu hành trình của bạn</h3>
                <p>
                    Vào mục <b>Khám phá địa điểm</b> để tìm nơi phù hợp,
                    chọn <b>Lập kế hoạch chuyến đi</b> để tạo lịch trình,
                    hoặc mở <b>Checklist chuẩn bị</b> để kiểm tra đồ dùng
                    trước khi xuất phát.
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("### 💛 Giới thiệu PetTrip")
    st.markdown(
        """
        PetTrip là nền tảng hỗ trợ người nuôi thú cưng lên kế hoạch cho những chuyến đi cùng chó, mèo. 
        Bạn có thể khám phá và so sánh địa điểm, xem thông tin phù hợp với nhu cầu và ngân sách, 
        lập lịch trình, lưu chuyến đi và chuẩn bị đồ dùng cần thiết cho bé.

        **PetTrip- Nền tảng du lịch cùng thú cưng giúp dễ dàng tìm kiếm, so sánh và lựa chọn địa điểm phù hợp với từng thú cưng,nhu cầu và ngân sách dựa trên đánh giá thực tế từ cộng đồng**
        """
    )
    st.markdown(
        '<div class="footer">PetTrip • Cùng bé đi muôn nơi 🐾</div>',
        unsafe_allow_html=True,
    )


# =====================================================
# 9. KHÁM PHÁ ĐỊA ĐIỂM
# =====================================================

elif trang == "🔎 Khám phá địa điểm":
    st.title("🔎 Khám phá địa điểm")
    st.markdown(
        '<p class="section-note">'
        'Tìm địa điểm phù hợp cho chuyến đi cùng bé.'
        '</p>',
        unsafe_allow_html=True,
    )

    if loi_du_lieu:
        st.error(loi_du_lieu)

    else:
        danh_sach_sheet = list(du_lieu.keys())

        ten_hien_thi = {
            sheet: ten_danh_muc_hien_thi(sheet, i)
            for i, sheet in enumerate(danh_sach_sheet)
        }

        cac_danh_muc = list(dict.fromkeys(ten_hien_thi.values()))

        chon_danh_muc = st.selectbox(
            "📂 Chọn danh mục",
            ["Tất cả danh mục"] + cac_danh_muc,
        )

        col1, col2 = st.columns(2)

        with col1:
            tu_khoa = st.text_input(
                "🔍 Tìm theo tên hoặc địa chỉ",
                placeholder="Ví dụ: Đà Lạt, quận 1...",
            )

        with col2:
            loc_loai = st.selectbox(
                "🐾 Loại thú cưng",
                ["Tất cả", "Chó", "Mèo"],
            )

        so_ket_qua = 0

        for i, (sheet, df) in enumerate(du_lieu.items()):
            ten_dm = ten_hien_thi[sheet]

            if (
                chon_danh_muc != "Tất cả danh mục"
                and ten_dm != chon_danh_muc
            ):
                continue

            if df.empty:
                continue

            cot = thong_tin_cot(df)
            df_loc = df.copy()

            if tu_khoa.strip():
                cot_tim = [
                    col
                    for col in [
                        cot["ten"],
                        cot["dia_chi"],
                        cot["thanh_pho"],
                    ]
                    if col is not None
                ]

                if cot_tim:
                    mask = df_loc[cot_tim].astype(str).apply(
                        lambda row: row.str.contains(
                            tu_khoa.strip(),
                            case=False,
                            na=False,
                            regex=False,
                        ).any(),
                        axis=1,
                    )
                    df_loc = df_loc[mask]

            if loc_loai != "Tất cả" and cot["loai"]:
                df_loc = df_loc[
                    df_loc[cot["loai"]].apply(
                        lambda x: loai_phu_hop(x, loc_loai)
                    )
                ]

            if df_loc.empty:
                continue

            st.subheader(f"📍 {ten_dm}")
            st.caption(f"Tìm thấy {len(df_loc)} địa điểm")

            for _, row in df_loc.iterrows():
                ten = (
                    hien_thi(row[cot["ten"]], "Địa điểm")
                    if cot["ten"]
                    else "Địa điểm"
                )

                dia_chi = (
                    hien_thi(row[cot["dia_chi"]], "")
                    if cot["dia_chi"]
                    else ""
                )

                thanh_pho = (
                    hien_thi(row[cot["thanh_pho"]], "")
                    if cot["thanh_pho"]
                    else ""
                )

                with st.container(border=True):
                    st.markdown(
                        f'<div class="place-name">📍 {ten}</div>',
                        unsafe_allow_html=True,
                    )

                    if dia_chi:
                        st.write(f"**Địa chỉ:** {dia_chi}")

                    if thanh_pho:
                        st.write(f"**Thành phố:** {thanh_pho}")

                    if cot["gia_min"] and hop_le(row[cot["gia_min"]]):
                        gia = hien_thi(row[cot["gia_min"]])

                        if cot["gia_max"] and hop_le(row[cot["gia_max"]]):
                            gia += " - " + hien_thi(row[cot["gia_max"]])

                        st.write(f"**Giá tham khảo:** {gia}")

                    if cot["gio"] and hop_le(row[cot["gio"]]):
                        st.write(
                            f"**Giờ hoạt động:** "
                            f"{hien_thi(row[cot['gio']])}"
                        )

                    if cot["sdt"] and hop_le(row[cot["sdt"]]):
                        st.write(
                            f"**Liên hệ:** {hien_thi(row[cot['sdt']])}"
                        )

                    if cot["loai"] and hop_le(row[cot["loai"]]):
                        st.write(
                            f"**Loài phù hợp:** "
                            f"{hien_thi(row[cot['loai']])}"
                        )

                    if cot["size"] and hop_le(row[cot["size"]]):
                        st.write(
                            f"**Kích thước phù hợp:** "
                            f"{hien_thi(row[cot['size']])}"
                        )

                    if cot["dich_vu"] and hop_le(row[cot["dich_vu"]]):
                        st.write(
                            f"**Dịch vụ:** {hien_thi(row[cot['dich_vu']])}"
                        )

                    if cot["cap_cuu"] and hop_le(row[cot["cap_cuu"]]):
                        st.write(
                            f"**Cấp cứu:** {hien_thi(row[cot['cap_cuu']])}"
                        )

                    st.link_button(
                        "🗺️ Xem trên Google Maps",
                        link_map(ten, dia_chi, thanh_pho),
                    )

                so_ket_qua += 1

            st.divider()

        if so_ket_qua == 0:
            st.info(
                "Chưa tìm thấy địa điểm phù hợp. "
                "Thử đổi từ khóa hoặc chọn danh mục khác nhé."
            )


# =====================================================
# 10. HỒ SƠ THÚ CƯNG
# =====================================================

elif trang == "🐾 Hồ sơ thú cưng":
    st.title("🐾 Hồ sơ thú cưng")
    st.write(
        "Cập nhật thông tin của bé để sử dụng khi lập kế hoạch chuyến đi."
    )

    loai_pet, size_pet, ten_pet = hien_thi_ho_so()

    if st.button("💾 Lưu hồ sơ thú cưng", use_container_width=True):
        st.session_state["pet_loai"] = loai_pet
        st.session_state["pet_size"] = size_pet
        st.session_state["pet_ten"] = ten_pet.strip()

        st.success("Đã lưu hồ sơ trong phiên sử dụng hiện tại!")

    st.caption(
        "Hồ sơ được giữ trong phiên sử dụng Streamlit. "
        "Danh sách checklist và chuyến đi sẽ được lưu riêng trong SQLite."
    )
    # =====================================================
# 11. LẬP KẾ HOẠCH CHUYẾN ĐI
# =====================================================

elif trang == "🗺️ Lập kế hoạch chuyến đi":
    st.title("🗺️ Lập kế hoạch chuyến đi")
    st.write(
        "Tạo hành trình riêng cho bạn và người bạn bốn chân."
    )

    if loi_du_lieu:
        st.error(loi_du_lieu)

    else:
        st.subheader("🐾 1. Thông tin chuyến đi")

        with st.form("form_ke_hoach"):
            ten_chuyen_di = st.text_input(
                "Tên chuyến đi",
                placeholder="Ví dụ: Cuối tuần cùng Bông ở Đà Lạt",
            )

            col1, col2 = st.columns(2)

            with col1:
                ten_pet = st.text_input(
                    "Tên thú cưng",
                    value=st.session_state.get("pet_ten", ""),
                )

                loai_pet = st.selectbox(
                    "Thú cưng",
                    ["Chó", "Mèo"],
                    index=(
                        1
                        if st.session_state.get("pet_loai") == "Mèo"
                        else 0
                    ),
                )

                size_pet = st.selectbox(
                    "Kích thước",
                    ["S", "M", "L", "XL"],
                    index=(
                        ["S", "M", "L", "XL"].index(
                            st.session_state.get("pet_size", "S")
                        )
                        if st.session_state.get("pet_size", "S")
                        in ["S", "M", "L", "XL"]
                        else 0
                    ),
                )

            with col2:
                diem_den = st.text_input(
                    "Điểm đến",
                    placeholder="Ví dụ: Đà Lạt",
                )

                ngay_di = st.date_input(
                    "Ngày bắt đầu",
                    value=date.today(),
                )

                ngay_ve = st.date_input(
                    "Ngày kết thúc",
                    value=date.today(),
                )

            ngan_sach = st.number_input(
                "Ngân sách dự kiến (VNĐ)",
                min_value=0,
                value=1000000,
                step=100000,
            )

            ghi_chu = st.text_area(
                "Ghi chú",
                placeholder=(
                    "Ví dụ: Ưu tiên nơi có không gian ngoài trời, "
                    "có chỗ đậu xe..."
                ),
            )

            tao_ke_hoach = st.form_submit_button(
                "🔎 Tìm địa điểm phù hợp",
                use_container_width=True,
            )

        if tao_ke_hoach:
            if not diem_den.strip():
                st.warning("M nhập điểm đến trước nha.")

            elif ngay_ve < ngay_di:
                st.warning(
                    "Ngày kết thúc không được trước ngày bắt đầu."
                )

            else:
                st.session_state["ke_hoach_hien_tai"] = {
                    "ten_chuyen_di": (
                        ten_chuyen_di.strip()
                        or f"Chuyến đi {diem_den.strip()}"
                    ),
                    "ten_pet": ten_pet.strip(),
                    "loai_pet": loai_pet,
                    "size_pet": size_pet,
                    "diem_den": diem_den.strip(),
                    "ngay_di": ngay_di.isoformat(),
                    "ngay_ve": ngay_ve.isoformat(),
                    "ngan_sach": int(ngan_sach),
                    "ghi_chu": ghi_chu.strip(),
                }

                st.session_state.pop("ket_qua_ke_hoach", None)

                ket_qua = []

                for sheet, df in du_lieu.items():
                    if df.empty:
                        continue

                    cot = thong_tin_cot(df)
                    df_loc = df.copy()

                    # Lọc theo điểm đến
                    cot_dia_diem = [
                        col
                        for col in [
                            cot["ten"],
                            cot["dia_chi"],
                            cot["thanh_pho"],
                        ]
                        if col is not None
                    ]

                    if cot_dia_diem:
                        mask_dia_diem = df_loc[cot_dia_diem].astype(
                            str
                        ).apply(
                            lambda row: row.str.contains(
                                diem_den.strip(),
                                case=False,
                                na=False,
                                regex=False,
                            ).any(),
                            axis=1,
                        )

                        df_loc = df_loc[mask_dia_diem]

                    # Lọc theo loài nếu bảng có cột loài
                    if cot["loai"]:
                        df_loc = df_loc[
                            df_loc[cot["loai"]].apply(
                                lambda x: loai_phu_hop(x, loai_pet)
                            )
                        ]

                                      # Lọc theo kích thước nếu bảng có cột tương ứng
                    cot_size = cot.get("size")

                    if cot_size and cot_size in df_loc.columns:
                        df_loc = df_loc[
                            df_loc[cot_size].apply(
                                lambda x: size_phu_hop(x, size_pet)
                            )
                        ]

                    
                    # Duyệt qua từng dòng dữ liệu
                    for _, row in df_loc.iterrows():
                        ten = (
                            hien_thi(row[cot["ten"]], "Địa điểm")
                            if cot["ten"] and cot["ten"] in df_loc.columns
                            else "Địa điểm"
                        )

                        dia_chi = (
                            hien_thi(row[cot["dia_chi"]], "")
                            if cot["dia_chi"] and cot["dia_chi"] in df_loc.columns
                            else ""
                        )

                        thanh_pho = (
                            hien_thi(row[cot["thanh_pho"]], "")
                            if cot["thanh_pho"] and cot["thanh_pho"] in df_loc.columns
                            else ""
                        )

                        ket_qua.append({
                            "danh_muc": ten_danh_muc_hien_thi(
                                sheet,
                                list(du_lieu.keys()).index(sheet),
                            ),
                            "ten": ten,
                            "dia_chi": dia_chi,
                            "thanh_pho": thanh_pho,
                            "map": link_map(
                                ten,
                                dia_chi,
                                thanh_pho,
                            ),
                        })

                st.session_state["ket_qua_ke_hoach"] = ket_qua
                st.session_state["dia_diem_duoc_chon"] = [
                    f"{i}. {item.get('ten', 'Địa điểm')} — {item.get('dia_chi', '')}".strip(" —")
                    for i, item in enumerate(ket_qua[:15], start=1)
                ]

                st.success(
                    f"Đã tìm thấy {len(ket_qua)} địa điểm phù hợp."
                )

                
                # Hiển thị kết quả sau khi tìm kiếm
        ke_hoach = st.session_state.get("ke_hoach_hien_tai")
        ket_qua = st.session_state.get("ket_qua_ke_hoach")

        if ke_hoach and ket_qua is not None:
            st.divider()
            st.subheader("📍 2. Địa điểm gợi ý")

            st.write(f"**Điểm đến:** {ke_hoach['diem_den']}")
            st.write(
                f"**Thời gian:** {ke_hoach['ngay_di']} "
                f"đến {ke_hoach['ngay_ve']}"
            )
            st.write(
                f"**Ngân sách dự kiến:** "
                f"{ke_hoach['ngan_sach']:,} VNĐ"
            )

            dia_diem_hien_thi = ket_qua[:15]
            khoa_dia_diem = {
                f"{i}. {item.get('ten', 'Địa điểm')} — {item.get('dia_chi', '')}".strip(" —"): item
                for i, item in enumerate(dia_diem_hien_thi, start=1)
            }
            if ket_qua:
                st.caption("M có thể chọn riêng những địa điểm muốn lưu vào lịch trình. Bỏ chọn địa điểm nào thì địa điểm đó sẽ không được lưu.")
                lua_chon_dia_diem = st.multiselect(
                    "📌 Chọn địa điểm muốn lưu",
                    options=list(khoa_dia_diem.keys()),
                    default=st.session_state.get("dia_diem_duoc_chon", list(khoa_dia_diem.keys())),
                    key="chon_dia_diem_luu",
                )
                st.session_state["dia_diem_duoc_chon"] = lua_chon_dia_diem
                for i, dia_diem in enumerate(dia_diem_hien_thi, start=1):
                    with st.container(border=True):
                        st.markdown(f"**{i}. {dia_diem['ten']}**")
                        st.caption(dia_diem["danh_muc"])

                        if dia_diem.get("dia_chi"):
                            st.write(
                                f"Địa chỉ: {dia_diem['dia_chi']}"
                            )

                        if dia_diem.get("thanh_pho"):
                            st.write(
                                f"Khu vực: {dia_diem['thanh_pho']}"
                            )

                        if dia_diem.get("map"):
                            st.link_button(
    "🗺️ Chỉ đường",
    dia_diem["map"],
)

                if len(ket_qua) > 15:
                    st.caption(
                        f"Còn {len(ket_qua) - 15} địa điểm khác "
                        "chưa hiển thị."
                    )
            else:
                st.info(
                    "Chưa tìm thấy địa điểm phù hợp. "
                    "M thử nhập tên thành phố hoặc địa chỉ ngắn hơn nha."
                )

            st.divider()
            st.subheader("💾 3. Lưu chuyến đi")

            st.caption(
                "Chuyến đi sẽ được lưu vào cơ sở dữ liệu SQLite."
            )

            if st.button(
                "💾 Lưu lịch trình này",
                use_container_width=True,
                key="luu_chuyen_di",
            ):
                try:
                    cac_nhan_duoc_chon = st.session_state.get("dia_diem_duoc_chon", [])
                    cac_dia_diem_duoc_luu = [
                        item for nhan, item in khoa_dia_diem.items()
                        if nhan in cac_nhan_duoc_chon
                    ]
                    noi_dung = {
                        "thong_tin": ke_hoach,
                        "dia_diem": cac_dia_diem_duoc_luu,
                    }

                    luu_chuyen_di(
                        ten=ke_hoach["ten_chuyen_di"],
                        diem_den=ke_hoach["diem_den"],
                        ngay_di=ke_hoach["ngay_di"],
                        ngay_ve=ke_hoach["ngay_ve"],
                        noi_dung=noi_dung,
                    )

                    # Kiểm tra ngay sau khi ghi để tránh báo thành công giả.
                    cac_chuyen_di_sau_khi_luu = doc_chuyen_di()
                    st.success(
                        f"Đã lưu chuyến đi cùng {len(cac_dia_diem_duoc_luu)} địa điểm đã chọn! "
                        f"Hiện có {len(cac_chuyen_di_sau_khi_luu)} chuyến đi trong danh sách đã lưu."
                    )

                except Exception as e:
                    st.error(f"Chưa lưu được chuyến đi: {e}")

# 12. CHUYẾN ĐI ĐÃ LƯU
# =====================================================

elif trang == "🧳 Chuyến đi đã lưu":
    st.title("🧳 Chuyến đi đã lưu")
    st.write(
        "Những lịch trình đã lưu trong cơ sở dữ liệu SQLite."
    )
    if st.button("🔄 Làm mới danh sách chuyến đi", key="lam_moi_chuyen_di"):
        st.rerun()

    try:
        cac_chuyen_di = doc_chuyen_di()

        if not cac_chuyen_di:
            st.info(
                "Chưa có chuyến đi nào. "
                "M tạo lịch trình trong mục 'Lập kế hoạch chuyến đi' "
                "rồi lưu lại nha."
            )
        else:
            for chuyen_di in cac_chuyen_di:
                with st.container(border=True):
                    st.subheader(chuyen_di["ten"])
                    st.write(
                        f"**Điểm đến:** {chuyen_di['diem_den']}"
                    )
                    st.write(
                        f"**Thời gian:** {chuyen_di['ngay_di']} "
                        f"đến {chuyen_di['ngay_ve']}"
                    )

                    noi_dung = chuyen_di.get("noi_dung", {})
                    thong_tin = noi_dung.get("thong_tin", {})
                    dia_diem = noi_dung.get("dia_diem", [])

                    if thong_tin:
                        st.write(
                            f"**Thú cưng:** "
                            f"{thong_tin.get('ten_pet') or 'Chưa đặt tên'}"
                        )
                        st.write(
                            f"**Loài:** {thong_tin.get('loai_pet', '')}"
                        )
                        st.write(
                            f"**Kích thước:** "
                            f"{thong_tin.get('size_pet', '')}"
                        )
                        st.write(
                            f"**Ngân sách:** "
                            f"{thong_tin.get('ngan_sach', 0):,} VNĐ"
                        )
                      
                    if dia_diem:
                        with st.expander("📍 Xem địa điểm trong lịch trình"):
                            for item in dia_diem:
                                st.markdown(
                                    f"- **{item.get('ten', 'Địa điểm')}** "
                                    f"({item.get('danh_muc', '')})"
                                )

                                if item.get("dia_chi"):
                                    st.caption(item["dia_chi"])

                                if item.get("map"):
                                    # Bản Streamlit đang dùng không hỗ trợ key cho st.link_button.
                                    st.link_button(
                                        "Xem bản đồ",
                                        item["map"],
                                    )

                    if st.session_state.get("xac_nhan_xoa") == chuyen_di["id"]:
                        st.warning(
                            "Bạn có chắc muốn xóa lịch trình này không?"
                        )


                        x1, x2 = st.columns(2)

                        with x1:
                            if st.button(
                                "Có, xóa",
                                key=f"dong_y_xoa_{chuyen_di['id']}",
                            ):
                                xoa_chuyen_di(chuyen_di["id"])
                                st.session_state.pop(
                                    "xac_nhan_xoa",
                                    None,
                                )
                                st.rerun()

                        with x2:
                            if st.button(
                                "Không xóa",
                                key=f"huy_xoa_{chuyen_di['id']}",
                            ):
                                st.session_state.pop(
                                    "xac_nhan_xoa",
                                    None,
                                )
                                st.rerun()

    except Exception as e:
        st.error(f"Không thể tải danh sách chuyến đi: {e}")
        
# =====================================================
# 13. CHECKLIST CHUẨN BỊ CHO CHÓ VÀ MÈO
# =====================================================

elif trang == "✅ Checklist chuẩn bị":
    st.title("✅ Checklist chuẩn bị chuyến đi")
    st.write(
        "Đánh dấu từng món đã chuẩn bị để chuyến đi "
        "cùng thú cưng thuận tiện hơn."
    )

    danh_sach_checklist = {
        "Chó": [
            "Dây dắt chắc chắn và vòng cổ/yếm vừa vặn",
            "Thẻ tên có số điện thoại liên hệ",
            "Rọ mõm phù hợp (nếu cần hoặc nơi đến yêu cầu)",
            "Lồng/đai an toàn khi đi ô tô hoặc phương tiện phù hợp",
            "Thức ăn quen thuộc đủ cho cả chuyến đi và một ít dự phòng",
            "Bát ăn, bát uống gấp gọn",
            "Nước uống sạch và bình nước mang theo",
            "Bánh thưởng và đồ chơi quen thuộc",
            "Túi nhặt chất thải và khăn giấy",
            "Túi rác nhỏ và khăn lau chân",
            "Khăn tắm, bàn chải/lược và đồ vệ sinh",
            "Tấm lót hoặc chăn/đệm quen thuộc",
            "Sổ tiêm phòng và hồ sơ sức khỏe cần thiết",
            "Thuốc đang dùng theo hướng dẫn của bác sĩ thú y (nếu có)",
            "Thông tin liên hệ bác sĩ thú y và phòng khám gần điểm đến",
            "Xác nhận nơi lưu trú/địa điểm có nhận chó",
            "Kiểm tra thời tiết, tuyến đường và điểm nghỉ chân",
            "Đảm bảo chó đã vận động, đi vệ sinh trước khi khởi hành",
        ],
        "Mèo": [
            "Lồng vận chuyển chắc chắn, thông thoáng và khóa an toàn",
            "Thẻ tên/thông tin liên hệ gắn trên lồng hoặc vòng cổ",
            "Thức ăn quen thuộc đủ cho chuyến đi và một ít dự phòng",
            "Bát ăn, bát uống gấp gọn",
            "Nước uống sạch và bình/chai nước mang theo",
            "Khay vệ sinh di động",
            "Cát vệ sinh quen thuộc và xẻng xúc cát",
            "Túi đựng chất thải và túi rác nhỏ",
            "Tấm lót thấm hút cho lồng vận chuyển",
            "Khăn hoặc chăn có mùi quen thuộc",
            "Đồ chơi nhỏ giúp mèo bình tĩnh",
            "Bàn chải/lược và khăn lau",
            "Sổ tiêm phòng và hồ sơ sức khỏe cần thiết",
            "Thuốc đang dùng theo hướng dẫn của bác sĩ thú y (nếu có)",
            "Thông tin bác sĩ thú y và phòng khám gần điểm đến",
            "Xác nhận nơi lưu trú/địa điểm có nhận mèo",
            "Chuẩn bị không gian yên tĩnh, kín và an toàn khi đến nơi",
            "Kiểm tra cửa lồng, cửa xe và cửa phòng để tránh mèo chạy mất",
        ],
    }

    loai_checklist = st.radio(
        "Chọn thú cưng",
        ["Chó", "Mèo"],
        horizontal=True,
        key="loai_checklist",
    )

    danh_sach = danh_sach_checklist[loai_checklist]
    trang_thai = doc_checklist(loai_checklist)

    # Hiển thị tiến độ
    so_da_chon = sum(
        1 for ten_viec in danh_sach
        if trang_thai.get(ten_viec, False)
    )
    tong_so = len(danh_sach)
    phan_tram = so_da_chon / tong_so if tong_so else 0

    st.subheader(
        f"📦 Đã chuẩn bị {so_da_chon}/{tong_so} món"
    )
    st.progress(phan_tram)
    st.caption(f"Hoàn thành {phan_tram * 100:.0f}%")

    st.divider()

    so_moi_nhom = (len(danh_sach) + 3) // 4
    cac_nhom = [
        ("🎒 Di chuyển và nhận diện", danh_sach[:4]),
        ("🍽️ Ăn uống và sinh hoạt", danh_sach[4:8]),
        ("🧼 Vệ sinh và tiện nghi", danh_sach[8:12]),
        ("🩺 Sức khỏe và an toàn", danh_sach[12:16]),
        ("📍 Kiểm tra trước chuyến đi", danh_sach[16:]),
    ]
    for nhom, cac_viec in cac_nhom:
        if not cac_viec:
            continue
        st.markdown(f"### {nhom}")

        for ten_viec in cac_viec:
            da_tich = st.checkbox(
                ten_viec,
                value=trang_thai.get(ten_viec, False),
                key=f"check_{loai_checklist}_{ten_viec}",
            )

            # Lưu trạng thái ngay sau khi đánh dấu
            if da_tich != trang_thai.get(ten_viec, False):
                luu_checklist(
                    loai_checklist,
                    ten_viec,
                    da_tich,
                )
                st.rerun()

    # Cho phép người dùng tự thêm đồ cần mang; mục tự thêm cũng được lưu SQLite.
    st.divider()
    st.markdown("### ➕ Tự thêm đồ cần mang")
    with st.form(f"form_them_checklist_{loai_checklist}", clear_on_submit=True):
        do_tu_them = st.text_input(
            "Tên đồ dùng / việc cần chuẩn bị",
            placeholder="Ví dụ: Áo mưa cho bé, thuốc chống say xe...",
        )
        them_do = st.form_submit_button("➕ Thêm vào checklist")

    if them_do:
        ten_moi = do_tu_them.strip()
        if not ten_moi:
            st.warning("M nhập tên món đồ trước nha.")
        elif ten_moi in danh_sach:
            st.info("Món này đã có sẵn trong checklist rồi.")
        else:
            luu_checklist(loai_checklist, ten_moi, False)
            st.success(f"Đã thêm ‘{ten_moi}’ vào checklist {loai_checklist.lower()}.")
            st.rerun()

    # Hiển thị những mục tự thêm từ database (không nằm trong danh sách gợi ý mặc định).
    cac_muc_tu_them = [ten for ten in trang_thai if ten not in danh_sach]
    if cac_muc_tu_them:
        st.markdown("### 🧺 Đồ bạn tự thêm")
        for ten_viec in cac_muc_tu_them:
            da_tich = st.checkbox(
                ten_viec,
                value=trang_thai.get(ten_viec, False),
                key=f"check_custom_{loai_checklist}_{ten_viec}",
            )
            col_check, col_delete = st.columns([5, 1])
            with col_check:
                if da_tich != trang_thai.get(ten_viec, False):
                    luu_checklist(loai_checklist, ten_viec, da_tich)
                    st.rerun()
            with col_delete:
                if st.button("🗑️", key=f"xoa_custom_{loai_checklist}_{ten_viec}", help="Xóa mục tự thêm"):
                    with sqlite3.connect(DB_FILE) as conn:
                        conn.execute(
                            "DELETE FROM checklist WHERE loai_pet = ? AND ten_viec = ?",
                            (loai_checklist, ten_viec),
                        )
                        conn.commit()
                    st.rerun()

    if st.button(
        "🔄 Đặt lại checklist",
        use_container_width=True,
    ):
        dat_lai_checklist(loai_checklist)
        st.rerun()

    st.caption(
        "Các mục được lưu trong SQLite của ứng dụng. "
        "Danh sách là gợi ý chung; hãy điều chỉnh theo sức khỏe, "
        "độ tuổi, tính cách của bé và quy định của phương tiện/nơi đến."
    )


# =====================================================
# =====================================================
# 15. GIỚI THIỆU PETTRIP
# =====================================================
