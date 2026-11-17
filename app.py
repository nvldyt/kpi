import base64
import io
import json
import os
import sqlite3
import urllib.error
import urllib.request
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Quản trị KPI - Tổ Dược Lâm Sàng",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==============================================================================
# GIAO DIỆN CSS CHUẨN Y KHOA - HIỂN thị RÕ RÀNG, KHÔNG LỖI FONT ICON
# ==============================================================================
st.markdown(
    """
    <style>
    html, body, .stApp, h1, h2, h3, h4, h5, h6,
    p, label, button, input, textarea,
    [data-testid="stSidebar"], [data-testid="stMetricValue"],
    [data-testid="stMetricLabel"], [data-testid="stMetricDelta"],
    [data-baseweb="tab"], [data-testid="stMarkdownContainer"] {
        font-family: Arial, Helvetica, sans-serif !important;
    }

    [data-testid="stIconMaterial"], .material-icons, span[translate="no"] {
        font-family: "Material Symbols Rounded", "Material Icons" !important;
    }

    .stApp {
        background-color: #f8fafc;
    }
    header[data-testid="stHeader"] {
        background-color: transparent !important;
    }
    .block-container {
        padding-top: 3.2rem !important;
        padding-bottom: 2rem !important;
        max-width: 98% !important;
    }

    .dls-header {
        background: linear-gradient(135deg, #0f4c81 0%, #0284c7 100%);
        color: #ffffff;
        padding: 14px 20px;
        border-radius: 10px;
        margin-top: 4px;
        margin-bottom: 12px;
        box-shadow: 0 2px 8px rgba(15, 76, 129, 0.15);
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .dls-header-title {
        font-size: 19px !important;
        font-weight: 700 !important;
        margin: 0 !important;
        line-height: 1.35 !important;
        color: #ffffff !important;
    }
    .dls-header-sub {
        font-size: 12.5px !important;
        color: #e0f2fe !important;
        margin-top: 4px !important;
    }
    .dls-badge {
        background: rgba(255, 255, 255, 0.18);
        padding: 6px 12px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 600;
        color: #ffffff;
        white-space: nowrap;
    }

    [data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-left: 4px solid #0284c7;
        padding: 8px 12px !important;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
    }
    [data-testid="stMetricLabel"] p {
        font-size: 12px !important;
        font-weight: 700 !important;
        color: #475569 !important;
    }
    [data-testid="stMetricValue"] {
        font-size: 17.5px !important;
        font-weight: 700 !important;
        color: #0f172a !important;
        line-height: 1.3 !important;
    }
    [data-testid="stMetricDelta"] {
        font-size: 11px !important;
        font-weight: 600 !important;
    }

    [data-baseweb="tab-list"] {
        gap: 5px !important;
        background-color: #e2e8f0;
        padding: 5px 6px;
        border-radius: 8px;
        margin-bottom: 10px;
    }
    [data-baseweb="tab"] {
        height: 36px !important;
        padding: 0 12px !important;
        border-radius: 6px !important;
        font-size: 12.5px !important;
        font-weight: 700 !important;
        color: #334155 !important;
        background-color: transparent !important;
        border: none !important;
    }
    [aria-selected="true"][data-baseweb="tab"] {
        background-color: #ffffff !important;
        color: #0284c7 !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1) !important;
    }

    .mod-title {
        font-size: 16px !important;
        font-weight: 700 !important;
        color: #0f172a !important;
        margin: 4px 0 2px 0 !important;
    }
    .mod-desc {
        font-size: 12.5px !important;
        color: #64748b !important;
        margin-bottom: 10px !important;
    }

    .excel-summary-bar {
        background-color: #f0f9ff;
        border: 1px solid #bae6fd;
        border-left: 4px solid #0284c7;
        padding: 8px 14px;
        border-radius: 6px;
        font-size: 13px;
        color: #0f172a;
        margin-top: 6px;
        margin-bottom: 10px;
    }

    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        font-size: 14px !important;
        font-weight: 700 !important;
        color: #0f172a !important;
        margin-top: 8px !important;
    }
    .stButton > button, .stDownloadButton > button {
        font-size: 12.5px !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
        padding: 0.35rem 0.8rem !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

STATE_VERSION = "v6.0_persistent_split_rx_daily_counsel"
SQLITE_DB_FILE = "kpi_dls_persistent_v6.db"
AUTOSAVE_JSON_FILE = "kpi_dls_15tv_autosave_v6.json"
GITHUB_BACKUP_PATH = "kpi_dls_15tv_autosave_v6.json"
WEEKDAYS = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"]

CLINICAL_DEPARTMENTS = [
    "Ban Giám đốc & Phòng KHTH (Điều phối chung)",
    "Khoa Khám bệnh & Điều trị Ngoại trú",
    "Khoa Hồi sức tích cực - Chống độc (ICU)",
    "Khoa Nội Tim mạch - Lão học",
    "Khoa Nội tiết - Đái tháo đường",
    "Khoa Bệnh nhiệt đới & Truyền nhiễm",
    "Khoa Thận - Tiết niệu & Lọc máu",
    "Khoa Nội Hô hấp - Tiêu hóa",
    "Khoa Sản - Phụ khoa",
    "Khoa Nhi tổng hợp",
    "Khoa Ung bướu & Huyết học lâm sàng",
    "Khoa Phẫu thuật - Gây mê hồi sức",
    "Khoa Ngoại Tổng hợp & Chấn thương",
    "Khoa Thần kinh - Cơ xương khớp",
    "Đơn vị Quản lý Đau & Chăm sóc giảm nhẹ",
    "Trung tâm Thông tin thuốc & Cảnh giác Dược (DI & ADR)",
    "Phòng Bảo hiểm Y tế & Kinh tế Dược",
    "Trung tâm Đào tạo & Nghiên cứu Khoa học",
]

DEPARTMENT_DEFAULT_TARGETS = {
    "Khoa Khám bệnh & Điều trị Ngoại trú": {
        "hsba": 0, "rx_out": 300, "rx_in": 0, "counsel_out": 150, "consult": 10, "dose": 15
    },
    "Khoa Hồi sức tích cực - Chống độc (ICU)": {
        "hsba": 65, "rx_out": 0, "rx_in": 200, "counsel_out": 0, "consult": 25, "dose": 30
    },
    "Khoa Ngoại Tổng hợp & Chấn thương": {
        "hsba": 70, "rx_out": 0, "rx_in": 175, "counsel_out": 0, "consult": 15, "dose": 20
    },
    "Khoa Bệnh nhiệt đới & Truyền nhiễm": {
        "hsba": 80, "rx_out": 0, "rx_in": 200, "counsel_out": 0, "consult": 25, "dose": 25
    },
    "Khoa Nội tiết - Đái tháo đường": {
        "hsba": 65, "rx_out": 150, "rx_in": 125, "counsel_out": 80, "consult": 15, "dose": 20
    },
    "Khoa Nội Hô hấp - Tiêu hóa": {
        "hsba": 75, "rx_out": 50, "rx_in": 175, "counsel_out": 30, "consult": 15, "dose": 20
    },
    "Khoa Nội Tim mạch - Lão học": {
        "hsba": 70, "rx_out": 150, "rx_in": 150, "counsel_out": 75, "consult": 15, "dose": 20
    },
    "Khoa Thần kinh - Cơ xương khớp": {
        "hsba": 65, "rx_out": 50, "rx_in": 150, "counsel_out": 30, "consult": 15, "dose": 20
    },
}


# ==============================================================================
# DỮ LIỆU GỐC 15 DƯỢC SĨ LÂM SÀNG (KHỞI TẠO SỐ THỰC HIỆN = 0 ĐỂ NHẬP THỰC TẾ)
# ==============================================================================
def get_default_15_members_records():
    raw_members = [
        ("#01", "Ngô Chí Hiếu", "Khoa Khám bệnh & Điều trị Ngoại trú",
         "Rà soát đơn ngoại trú; trao đổi bác sĩ lựa chọn thuốc, chỉnh liều, xử lý tương tác; tư vấn người bệnh."),
        ("#02", "Nguyễn Thị Nhung", "Khoa Hồi sức tích cực - Chống độc (ICU)",
         "Hội chẩn sử dụng thuốc; rà soát bệnh án, hiệu chỉnh liều và can thiệp DLS tại khoa HSTC-CĐ, Sản, Nhi 01."),
        ("#03", "Nguyễn Tuấn Hùng", "Khoa Ngoại Tổng hợp & Chấn thương",
         "Hội chẩn sử dụng thuốc; rà soát bệnh án, hiệu chỉnh liều và can thiệp DLS tại khoa Ngoại tổng hợp - Chấn thương 01."),
        ("#04", "Nguyễn Hữu Hồng Đức", "Khoa Khám bệnh & Điều trị Ngoại trú",
         "Rà soát đơn ngoại trú; trao đổi bác sĩ lựa chọn thuốc, chỉnh liều, xử lý tương tác; tư vấn người bệnh."),
        ("#05", "Nguyễn Khắc Huy", "Khoa Bệnh nhiệt đới & Truyền nhiễm",
         "Hội chẩn sử dụng thuốc; rà soát bệnh án, hiệu chỉnh liều và can thiệp DLS tại khoa Bệnh nhiệt đới, Nội tim mạch."),
        ("#06", "Nguyễn Văn Huy", "Khoa Nội tiết - Đái tháo đường",
         "Hội chẩn sử dụng thuốc; rà soát bệnh án, hiệu chỉnh liều và can thiệp DLS tại khoa Nội tổng hợp 01."),
        ("#07", "Trịnh Thị Ngọc Lê", "Khoa Nội Hô hấp - Tiêu hóa",
         "Hội chẩn sử dụng thuốc; rà soát bệnh án, hiệu chỉnh liều và can thiệp DLS tại khoa TMH, RHM, Mắt, YHCT 01."),
        ("#08", "Lê Thị Thúy Hằng", "Khoa Khám bệnh & Điều trị Ngoại trú",
         "Kiểm soát an toàn đơn thuốc ngoại trú, dược lý cơ bản và tư vấn hướng dẫn sử dụng thuốc."),
        ("#09", "Lê Trần Đạt", "Khoa Hồi sức tích cực - Chống độc (ICU)",
         "Hội chẩn sử dụng thuốc; rà soát bệnh án, hiệu chỉnh liều và can thiệp DLS tại khoa HSTC-CĐ 02."),
        ("#10", "Ngũ Thị Hồng Nhung", "Khoa Nội Tim mạch - Lão học",
         "Rà soát đơn ngoại trú, tư vấn người bệnh; Hội chẩn, rà soát bệnh án và chỉnh liều tại khoa Nội tổng hợp 02."),
        ("#11", "Trịnh Thị Quỳnh Nga", "Khoa Ngoại Tổng hợp & Chấn thương",
         "Hội chẩn sử dụng thuốc; rà soát bệnh án, hiệu chỉnh liều và can thiệp DLS tại khoa Ngoại tổng hợp 02, TMH, Mắt, RHM."),
        ("#12", "Phạm Thị Thùy Ngân", "Khoa Nội Hô hấp - Tiêu hóa",
         "Hội chẩn sử dụng thuốc; rà soát bệnh án, hiệu chỉnh liều và can thiệp DLS tại Cơ sở 03."),
        ("#13", "Trần Văn Đông", "Khoa Thần kinh - Cơ xương khớp",
         "Hội chẩn sử dụng thuốc; rà soát bệnh án, hiệu chỉnh liều và can thiệp DLS tại khoa YHCT."),
        ("#14", "Đậu Đình Hùng Cường", "Khoa Khám bệnh & Điều trị Ngoại trú",
         "Rà soát đơn ngoại trú; trao đổi bác sĩ lựa chọn thuốc, chỉnh liều, xử lý tương tác; tư vấn người bệnh."),
        ("#15", "Lê Thị Thảo", "Khoa Khám bệnh & Điều trị Ngoại trú",
         "Hội chẩn sử dụng thuốc; rà soát bệnh án, duyệt đơn, hiệu chỉnh liều và can thiệp DLS tại Cơ sở 03."),
    ]

    records = []
    for mid, name, dept, task in raw_members:
        prof = DEPARTMENT_DEFAULT_TARGETS.get(
            dept,
            {"hsba": 60, "rx_out": 100, "rx_in": 100, "counsel_out": 50, "consult": 15, "dose": 15},
        )
        records.append(
            {
                "id": mid,
                "name": name,
                "department": dept,
                "assigned_task": task,
                # --- CHỈ TIÊU & CHỈ SỐ GIAO Ở MODULE 1 ---
                "target_hsba": prof["hsba"],
                "target_rx_out": prof["rx_out"],
                "target_rx_in": prof["rx_in"],
                "target_counsel_out": prof["counsel_out"],
                "target_consult": prof["consult"],
                "target_dose": prof["dose"],
                "doctor_consensus_pct": 95.0,
                "intervention_success_pct": 95.0,
                "task_quality_pct": 96.0,
                # --- M2: HỒ SƠ BỆNH ÁN (T2 - T6) ---
                "hsba_t2": 0, "hsba_t3": 0, "hsba_t4": 0, "hsba_t5": 0, "hsba_t6": 0,
                # --- M3: DUYỆT ĐƠN NGOẠI TRÚ (T2 - T6) ---
                "rx_out_t2": 0, "rx_out_t3": 0, "rx_out_t4": 0, "rx_out_t5": 0, "rx_out_t6": 0,
                # --- M4: DUYỆT ĐƠN NỘI TRÚ (T2 - T6) ---
                "rx_in_t2": 0, "rx_in_t3": 0, "rx_in_t4": 0, "rx_in_t5": 0, "rx_in_t6": 0,
                # --- M5: TƯ VẤN ĐƠN THUỐC NGOẠI TRÚ HẰNG NGÀY (T2 - T6) ---
                "coun_out_t2": 0, "coun_out_t3": 0, "coun_out_t4": 0, "coun_out_t5": 0, "coun_out_t6": 0,
                # --- M6: HỘI CHẨN LÂM SÀNG (T2 - T6) ---
                "consult_t2": 0, "consult_t3": 0, "consult_t4": 0, "consult_t5": 0, "consult_t6": 0,
                # --- M7: CAN THIỆP CHỈNH LIỀU (T2 - T6) ---
                "dose_t2": 0, "dose_t3": 0, "dose_t4": 0, "dose_t5": 0, "dose_t6": 0,
            }
        )
    return records


def get_default_config():
    return {
        "max_score_cap": 100.0,
        "ignore_zero_targets": True,
        "weights": {
            "w_m2_hsba": 20,
            "w_m3_rx_out": 15,
            "w_m4_rx_in": 15,
            "w_m5_counsel": 15,
            "w_m6_consult": 15,
            "w_m7_dose": 10,
            "w_m1_task": 10,
        },
    }


# ==============================================================================
# ĐỒNG BỘ ĐÁM MÂY GITHUB (CHỐNG RESET KHI STREAMLIT CLOUD NGỦ QUA ĐÊM)
# ==============================================================================
def get_github_credentials():
    token = os.environ.get("GITHUB_TOKEN", "")
    repo = os.environ.get("GITHUB_REPO", "")
    branch = os.environ.get("GITHUB_BRANCH", "main")
    try:
        if not token and "GITHUB_TOKEN" in st.secrets:
            token = st.secrets["GITHUB_TOKEN"]
        if not repo and "GITHUB_REPO" in st.secrets:
            repo = st.secrets["GITHUB_REPO"]
        if "GITHUB_BRANCH" in st.secrets:
            branch = st.secrets["GITHUB_BRANCH"]
    except Exception:
        pass
    return str(token).strip(), str(repo).strip(), str(branch).strip()


def push_payload_to_github(payload_dict):
    token, repo, branch = get_github_credentials()
    if not token or not repo:
        return False
    url = f"https://api.github.com/repos/{repo}/contents/{GITHUB_BACKUP_PATH}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "KPI-DLS-Streamlit-App",
    }
    sha = None
    try:
        req_get = urllib.request.Request(f"{url}?ref={branch}", headers=headers, method="GET")
        with urllib.request.urlopen(req_get, timeout=5) as resp:
            existing = json.loads(resp.read().decode("utf-8"))
            sha = existing.get("sha")
    except Exception:
        sha = None

    content_bytes = json.dumps(payload_dict, ensure_ascii=False, indent=2).encode("utf-8")
    body = {
        "message": f"Auto-save KPI DLS {payload_dict.get('updated_at', '')}",
        "content": base64.b64encode(content_bytes).decode("utf-8"),
        "branch": branch,
    }
    if sha:
        body["sha"] = sha

    try:
        req_put = urllib.request.Request(
            url,
            data=json.dumps(body).encode("utf-8"),
            headers={**headers, "Content-Type": "application/json"},
            method="PUT",
        )
        with urllib.request.urlopen(req_put, timeout=6) as resp:
            return resp.status in (200, 201)
    except Exception:
        return False


def pull_payload_from_github():
    token, repo, branch = get_github_credentials()
    if not token or not repo:
        return None
    url = f"https://api.github.com/repos/{repo}/contents/{GITHUB_BACKUP_PATH}?ref={branch}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "KPI-DLS-Streamlit-App",
    }
    try:
        req = urllib.request.Request(url, headers=headers, method="GET")
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            raw_json = base64.b64decode(data["content"]).decode("utf-8")
            parsed = json.loads(raw_json)
            if "records" in parsed and len(parsed["records"]) > 0:
                return parsed
    except Exception:
        return None
    return None


# ==============================================================================
# CƠ SỞ DỮ LIỆU DÙNG CHUNG TOÀN SERVER (SQLITE WAL + SHARED MEMORY)
# ==============================================================================
def init_sqlite_db():
    conn = sqlite3.connect(SQLITE_DB_FILE, timeout=15, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS kpi_store (
            store_key TEXT PRIMARY KEY,
            payload_json TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.commit()
    conn.close()


def read_payload_from_disk_or_cloud():
    init_sqlite_db()
    # 1. Đọc từ SQLite nội bộ trước
    try:
        conn = sqlite3.connect(SQLITE_DB_FILE, timeout=15, check_same_thread=False)
        cur = conn.cursor()
        cur.execute("SELECT payload_json FROM kpi_store WHERE store_key = 'main_state'")
        row = cur.fetchone()
        conn.close()
        if row and row[0]:
            data = json.loads(row[0])
            if "records" in data and len(data["records"]) > 0:
                return data
    except Exception:
        pass

    # 2. Đọc từ file JSON nội bộ
    if os.path.exists(AUTOSAVE_JSON_FILE):
        try:
            with open(AUTOSAVE_JSON_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if "records" in data and len(data["records"]) > 0:
                return data
        except Exception:
            pass

    # 3. Đọc từ GitHub Cloud nếu container vừa khởi động lại sau khi ngủ qua đêm
    cloud_data = pull_payload_from_github()
    if cloud_data is not None:
        write_payload_to_disk_and_cloud(cloud_data, sync_cloud=False)
        return cloud_data

    return None


def write_payload_to_disk_and_cloud(payload, sync_cloud=True):
    payload["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    json_str = json.dumps(payload, ensure_ascii=False, indent=2)

    # 1. Ghi vào SQLite (An toàn đa luồng)
    try:
        init_sqlite_db()
        conn = sqlite3.connect(SQLITE_DB_FILE, timeout=15, check_same_thread=False)
        conn.execute(
            """
            INSERT INTO kpi_store (store_key, payload_json, updated_at)
            VALUES ('main_state', ?, ?)
            ON CONFLICT(store_key) DO UPDATE SET
                payload_json = excluded.payload_json,
                updated_at = excluded.updated_at
            """,
            (json_str, payload["updated_at"]),
        )
        conn.commit()
        conn.close()
    except Exception:
        pass

    # 2. Ghi song song ra file JSON
    try:
        with open(AUTOSAVE_JSON_FILE, "w", encoding="utf-8") as f:
            f.write(json_str)
    except Exception:
        pass

    # 3. Đồng bộ lên GitHub nếu có cấu hình Token
    if sync_cloud:
        push_payload_to_github(payload)


@st.cache_resource
def get_shared_server_store():
    """Bộ nhớ chung cho tất cả các máy truy cập cùng lúc, không bao giờ bị tab cũ ghi đè tab mới."""
    existing = read_payload_from_disk_or_cloud()
    if existing is not None:
        default_rec_map = {r["id"]: r for r in get_default_15_members_records()}
        for r in existing["records"]:
            base_r = default_rec_map.get(r["id"], {})
            for k, v in base_r.items():
                if k not in r or pd.isna(r[k]):
                    r[k] = v
        if "config" not in existing:
            existing["config"] = get_default_config()
        return existing

    initial_payload = {
        "app_version": STATE_VERSION,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "config": get_default_config(),
        "records": get_default_15_members_records(),
    }
    write_payload_to_disk_and_cloud(initial_payload, sync_cloud=True)
    return initial_payload


SHARED_STORE = get_shared_server_store()

if "editor_ver" not in st.session_state:
    st.session_state.editor_ver = 0


# ==============================================================================
# CẬP NHẬT TỪNG Ô NGUYÊN TỬ (ATOMIC CELL UPDATE) - NHẬP SỐ LÀ TỰ LƯU NGAY
# ==============================================================================
def make_atomic_callback(editor_key_prefix, id_list_key, int_cols=None, float_cols=None):
    int_cols = set(int_cols or [])
    float_cols = set(float_cols or [])

    def _callback():
        full_key = f"{editor_key_prefix}_{st.session_state.editor_ver}"
        state_val = st.session_state.get(full_key, {})
        edited_rows = state_val.get("edited_rows", {})
        view_ids = st.session_state.get(id_list_key, [])
        if not edited_rows:
            return

        rec_map = {r["id"]: r for r in SHARED_STORE["records"]}
        changed_any = False

        for r_idx_str, col_changes in edited_rows.items():
            r_idx = int(r_idx_str)
            if r_idx < len(view_ids):
                mid = view_ids[r_idx]
                target_rec = rec_map.get(mid)
                if target_rec is not None:
                    for col_name, new_val in col_changes.items():
                        if col_name in int_cols:
                            clean_val = max(0, int(float(new_val))) if new_val is not None else 0
                        elif col_name in float_cols:
                            clean_val = max(0.0, float(new_val)) if new_val is not None else 0.0
                        else:
                            clean_val = str(new_val).strip() if new_val is not None else ""
                        target_rec[col_name] = clean_val
                        changed_any = True

        if changed_any:
            write_payload_to_disk_and_cloud(SHARED_STORE, sync_cloud=True)

    return _callback


# ==============================================================================
# HÀM TÍNH ĐIỂM CHI TIẾT CHO TỪNG MODULE (M1 -> M8)
# ==============================================================================
def compute_all_kpi_dataframes(records, config):
    df = pd.DataFrame(records).copy()
    max_cap = float(config.get("max_score_cap", 100.0))
    ignore_zero = bool(config.get("ignore_zero_targets", True))
    weights = config.get("weights", get_default_config()["weights"])

    def calc_simple_qty_score(done_series, target_series):
        pct_list, score_list = [], []
        for d, t in zip(done_series, target_series):
            if t <= 0:
                pct = 100.0 if d == 0 else min(max_cap, 100.0)
                sc = 100.0
            else:
                pct = round((d / t) * 100.0, 1)
                sc = round(min(max_cap, max(0.0, pct)), 1)
            pct_list.append(pct)
            score_list.append(sc)
        return pd.Series(pct_list), pd.Series(score_list)

    # --- M2: HỒ SƠ BỆNH ÁN ---
    hsba_days = ["hsba_t2", "hsba_t3", "hsba_t4", "hsba_t5", "hsba_t6"]
    df["done_hsba"] = df[hsba_days].sum(axis=1)
    df["diff_hsba"] = df["done_hsba"] - df["target_hsba"]
    df["pct_hsba"], df["score_m2_hsba"] = calc_simple_qty_score(df["done_hsba"], df["target_hsba"])

    # --- M3: DUYỆT ĐƠN NGOẠI TRÚ ---
    rx_out_days = ["rx_out_t2", "rx_out_t3", "rx_out_t4", "rx_out_t5", "rx_out_t6"]
    df["done_rx_out"] = df[rx_out_days].sum(axis=1)
    df["diff_rx_out"] = df["done_rx_out"] - df["target_rx_out"]
    df["pct_rx_out"], df["score_m3_rx_out"] = calc_simple_qty_score(df["done_rx_out"], df["target_rx_out"])

    # --- M4: DUYỆT ĐƠN NỘI TRÚ ---
    rx_in_days = ["rx_in_t2", "rx_in_t3", "rx_in_t4", "rx_in_t5", "rx_in_t6"]
    df["done_rx_in"] = df[rx_in_days].sum(axis=1)
    df["diff_rx_in"] = df["done_rx_in"] - df["target_rx_in"]
    df["pct_rx_in"], df["score_m4_rx_in"] = calc_simple_qty_score(df["done_rx_in"], df["target_rx_in"])

    # --- M5: TƯ VẤN ĐƠN THUỐC NGOẠI TRÚ HẰNG NGÀY ---
    coun_days = ["coun_out_t2", "coun_out_t3", "coun_out_t4", "coun_out_t5", "coun_out_t6"]
    df["done_counsel_out"] = df[coun_days].sum(axis=1)
    df["diff_counsel_out"] = df["done_counsel_out"] - df["target_counsel_out"]
    df["pct_counsel_out"], df["score_m5_counsel"] = calc_simple_qty_score(
        df["done_counsel_out"], df["target_counsel_out"]
    )
    # Tỷ lệ % đơn ngoại trú được tư vấn so với số đơn ngoại trú đã duyệt
    df["rate_counsel_vs_rx_out"] = df.apply(
        lambda r: round(min(100.0, (r["done_counsel_out"] / r["done_rx_out"]) * 100.0), 1)
        if r["done_rx_out"] > 0
        else (100.0 if r["done_counsel_out"] > 0 else 0.0),
        axis=1,
    )

    # --- M6: HỘI CHẨN LÂM SÀNG (65% Số lượng + 35% BS đồng thuận) ---
    con_days = ["consult_t2", "consult_t3", "consult_t4", "consult_t5", "consult_t6"]
    df["done_consult"] = df[con_days].sum(axis=1)
    df["diff_consult"] = df["done_consult"] - df["target_consult"]

    def calc_m6(row):
        t, d, cons = row["target_consult"], row["done_consult"], row["doctor_consensus_pct"]
        if t <= 0:
            return pd.Series([100.0, 100.0])
        if d <= 0:
            return pd.Series([0.0, 0.0])
        qty_pct = round((d / t) * 100.0, 1)
        sc = round(min(max_cap, 0.65 * min(max_cap, qty_pct) + 0.35 * cons), 1)
        return pd.Series([qty_pct, sc])

    df[["pct_consult", "score_m6_consult"]] = df.apply(calc_m6, axis=1)

    # --- M7: CAN THIỆP CHỈNH LIỀU (65% Số lượng + 35% Thành công) ---
    dose_days = ["dose_t2", "dose_t3", "dose_t4", "dose_t5", "dose_t6
