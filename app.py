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
    dose_days = ["dose_t2", "dose_t3", "dose_t4", "dose_t5", "dose_t6"]
    df["done_dose"] = df[dose_days].sum(axis=1)
    df["diff_dose"] = df["done_dose"] - df["target_dose"]

    def calc_m7(row):
        t, d, succ = row["target_dose"], row["done_dose"], row["intervention_success_pct"]
        if t <= 0:
            return pd.Series([100.0, 100.0])
        if d <= 0:
            return pd.Series([0.0, 0.0])
        qty_pct = round((d / t) * 100.0, 1)
        sc = round(min(max_cap, 0.65 * min(max_cap, qty_pct) + 0.35 * succ), 1)
        return pd.Series([qty_pct, sc])

    df[["pct_dose", "score_m7_dose"]] = df.apply(calc_m7, axis=1)

    # --- M8: TỔNG HỢP KPI TOÀN DIỆN ---
    def calc_total_kpi(row):
        components = [
            (row["score_m2_hsba"], weights["w_m2_hsba"], row["target_hsba"], row["done_hsba"]),
            (row["score_m3_rx_out"], weights["w_m3_rx_out"], row["target_rx_out"], row["done_rx_out"]),
            (row["score_m4_rx_in"], weights["w_m4_rx_in"], row["target_rx_in"], row["done_rx_in"]),
            (row["score_m5_counsel"], weights["w_m5_counsel"], row["target_counsel_out"], row["done_counsel_out"]),
            (row["score_m6_consult"], weights["w_m6_consult"], row["target_consult"], row["done_consult"]),
            (row["score_m7_dose"], weights["w_m7_dose"], row["target_dose"], row["done_dose"]),
            (row["task_quality_pct"], weights["w_m1_task"], 1, 1),
        ]
        weighted_sum = 0.0
        active_weight = 0.0
        for sc, w, target_val, done_val in components:
            if w <= 0:
                continue
            # Nếu bật chế độ công bằng: chỉ tiêu nào = 0 (không phân công) và thực hiện = 0 thì không tính vào mẫu số
            if ignore_zero and target_val <= 0 and done_val <= 0:
                continue
            weighted_sum += sc * w
            active_weight += w

        if active_weight <= 0:
            return 0.0
        return round(weighted_sum / active_weight, 1)

    df["kpi_total"] = df.apply(calc_total_kpi, axis=1)

    def rank_kpi(val):
        if val >= 95.0:
            return "Xuất sắc (A+)"
        elif val >= 90.0:
            return "Tốt (A)"
        elif val >= 80.0:
            return "Đạt chuẩn (B)"
        elif val > 0:
            return "Cần cải thiện (C)"
        return "Chưa nhập (—)"

    df["kpi_rank"] = df["kpi_total"].apply(rank_kpi)
    return df


# ==============================================================================
# XUẤT / NHẬP FILE EXCEL 8 SHEET ĐẦY ĐỦ
# ==============================================================================
def export_all_to_excel_bytes(df_calc):
    output = io.BytesIO()
    for eng in ["openpyxl", "xlsxwriter"]:
        try:
            with pd.ExcelWriter(output, engine=eng) as writer:
                df_calc[
                    [
                        "id", "name", "department",
                        "done_hsba", "target_hsba", "score_m2_hsba",
                        "done_rx_out", "target_rx_out", "score_m3_rx_out",
                        "done_rx_in", "target_rx_in", "score_m4_rx_in",
                        "done_counsel_out", "target_counsel_out", "score_m5_counsel",
                        "done_consult", "target_consult", "score_m6_consult",
                        "done_dose", "target_dose", "score_m7_dose",
                        "task_quality_pct", "kpi_total", "kpi_rank",
                    ]
                ].to_excel(writer, sheet_name="M8_TongHop_KPI", index=False)

                df_calc[
                    [
                        "id", "name", "department", "assigned_task",
                        "target_hsba", "target_rx_out", "target_rx_in",
                        "target_counsel_out", "target_consult", "target_dose",
                        "doctor_consensus_pct", "intervention_success_pct", "task_quality_pct",
                    ]
                ].to_excel(writer, sheet_name="M1_GiaoChiTieu", index=False)

                df_calc[
                    ["id", "name", "department", "target_hsba", "hsba_t2", "hsba_t3", "hsba_t4", "hsba_t5", "hsba_t6", "done_hsba", "score_m2_hsba"]
                ].to_excel(writer, sheet_name="M2_BenhAn", index=False)

                df_calc[
                    ["id", "name", "department", "target_rx_out", "rx_out_t2", "rx_out_t3", "rx_out_t4", "rx_out_t5", "rx_out_t6", "done_rx_out", "score_m3_rx_out"]
                ].to_excel(writer, sheet_name="M3_DuyetDon_NgoaiTru", index=False)

                df_calc[
                    ["id", "name", "department", "target_rx_in", "rx_in_t2", "rx_in_t3", "rx_in_t4", "rx_in_t5", "rx_in_t6", "done_rx_in", "score_m4_rx_in"]
                ].to_excel(writer, sheet_name="M4_DuyetDon_NoiTru", index=False)

                df_calc[
                    ["id", "name", "department", "target_counsel_out", "coun_out_t2", "coun_out_t3", "coun_out_t4", "coun_out_t5", "coun_out_t6", "done_counsel_out", "score_m5_counsel"]
                ].to_excel(writer, sheet_name="M5_TuVan_NgoaiTru", index=False)

                df_calc[
                    ["id", "name", "department", "target_consult", "consult_t2", "consult_t3", "consult_t4", "consult_t5", "consult_t6", "done_consult", "doctor_consensus_pct", "score_m6_consult"]
                ].to_excel(writer, sheet_name="M6_HoiChan", index=False)

                df_calc[
                    ["id", "name", "department", "target_dose", "dose_t2", "dose_t3", "dose_t4", "dose_t5", "dose_t6", "done_dose", "intervention_success_pct", "score_m7_dose"]
                ].to_excel(writer, sheet_name="M7_ChinhLieu", index=False)

                pd.DataFrame(SHARED_STORE["records"]).to_excel(writer, sheet_name="DuLieuGoc_Full", index=False)
            return output.getvalue()
        except ImportError:
            output = io.BytesIO()
            continue
    return None


def import_records_from_excel(uploaded_file):
    sheets = pd.read_excel(uploaded_file, sheet_name=None)
    if "DuLieuGoc_Full" in sheets:
        df_full = sheets["DuLieuGoc_Full"]
        SHARED_STORE["records"] = df_full.to_dict(orient="records")
        write_payload_to_disk_and_cloud(SHARED_STORE, sync_cloud=True)
        return True

    rec_map = {r["id"]: r for r in SHARED_STORE["records"]}
    sheet_mappings = [
        ("M1_GiaoChiTieu", ["name", "department", "assigned_task", "target_hsba", "target_rx_out", "target_rx_in", "target_counsel_out", "target_consult", "target_dose", "doctor_consensus_pct", "intervention_success_pct", "task_quality_pct"]),
        ("M2_BenhAn", ["target_hsba", "hsba_t2", "hsba_t3", "hsba_t4", "hsba_t5", "hsba_t6"]),
        ("M3_DuyetDon_NgoaiTru", ["target_rx_out", "rx_out_t2", "rx_out_t3", "rx_out_t4", "rx_out_t5", "rx_out_t6"]),
        ("M4_DuyetDon_NoiTru", ["target_rx_in", "rx_in_t2", "rx_in_t3", "rx_in_t4", "rx_in_t5", "rx_in_t6"]),
        ("M5_TuVan_NgoaiTru", ["target_counsel_out", "coun_out_t2", "coun_out_t3", "coun_out_t4", "coun_out_t5", "coun_out_t6"]),
        ("M6_HoiChan", ["target_consult", "consult_t2", "consult_t3", "consult_t4", "consult_t5", "consult_t6", "doctor_consensus_pct"]),
        ("M7_ChinhLieu", ["target_dose", "dose_t2", "dose_t3", "dose_t4", "dose_t5", "dose_t6", "intervention_success_pct"]),
    ]
    updated = False
    for s_name, cols in sheet_mappings:
        if s_name in sheets:
            df_s = sheets[s_name]
            for _, row in df_s.iterrows():
                mid = str(row.get("id", "")).strip()
                if mid in rec_map:
                    for c in cols:
                        if c in df_s.columns and pd.notna(row[c]):
                            rec_map[mid][c] = row[c]
                    updated = True
    if updated:
        write_payload_to_disk_and_cloud(SHARED_STORE, sync_cloud=True)
    return updated


# ==============================================================================
# THANH BÊN (SIDEBAR): CẤU HÌNH TRỌNG SỐ, LỌC KHOA & SAO LƯU AN TOÀN
# ==============================================================================
cfg = SHARED_STORE["config"]
gh_token, gh_repo, _ = get_github_credentials()

with st.sidebar:
    st.markdown("## 🏥 QUẢN TRỊ KPI DƯỢC LÂM SÀNG")
    st.success(f"🟢 Tự động lưu: **{SHARED_STORE.get('updated_at', 'Vừa xong')}**")
    if gh_token and gh_repo:
        st.caption(f"☁️ Đồng bộ Cloud GitHub: **Đang bật ({gh_repo})**")
    else:
        st.caption("💾 Chế độ lưu: **SQLite + JSON Máy chủ**")

    all_depts_in_use = sorted(list({r["department"] for r in SHARED_STORE["records"]}))
    selected_depts = st.multiselect(
        "🔍 Lọc theo Khoa phụ trách:",
        options=all_depts_in_use,
        default=all_depts_in_use,
    )

    st.markdown("### ⚖️ Cấu hình Chấm điểm & Trọng số")
    ignore_zero_targets = st.checkbox(
        "Khấu trừ chỉ tiêu = 0 (Chỉ tính điểm các phần có giao chỉ tiêu)",
        value=bool(cfg.get("ignore_zero_targets", True)),
        help="Giúp công bằng giữa Dược sĩ Nội trú (CT Ngoại trú = 0) và Dược sĩ Ngoại trú (CT Nội trú = 0).",
    )
    max_score_cap = st.number_input(
        "Trần điểm tối đa từng phần (100 - 120):",
        min_value=100.0,
        max_value=120.0,
        value=float(cfg.get("max_score_cap", 100.0)),
        step=5.0,
    )

    w_cur = cfg["weights"]
    w_m2_hsba = st.slider("M2: Hồ sơ Bệnh án (%)", 0, 50, int(w_cur["w_m2_hsba"]), 5)
    w_m3_rx_out = st.slider("M3: Duyệt đơn Ngoại trú (%)", 0, 50, int(w_cur["w_m3_rx_out"]), 5)
    w_m4_rx_in = st.slider("M4: Duyệt đơn Nội trú (%)", 0, 50, int(w_cur["w_m4_rx_in"]), 5)
    w_m5_counsel = st.slider("M5: Tư vấn đơn Ngoại trú (%)", 0, 50, int(w_cur["w_m5_counsel"]), 5)
    w_m6_consult = st.slider("M6: Hội chẩn Lâm sàng (%)", 0, 50, int(w_cur["w_m6_consult"]), 5)
    w_m7_dose = st.slider("M7: Can thiệp Chỉnh liều (%)", 0, 50, int(w_cur["w_m7_dose"]), 5)
    w_m1_task = st.slider("M1: Chất lượng Nhiệm vụ (%)", 0, 50, int(w_cur["w_m1_task"]), 5)

    new_weights = {
        "w_m2_hsba": w_m2_hsba,
        "w_m3_rx_out": w_m3_rx_out,
        "w_m4_rx_in": w_m4_rx_in,
        "w_m5_counsel": w_m5_counsel,
        "w_m6_consult": w_m6_consult,
        "w_m7_dose": w_m7_dose,
        "w_m1_task": w_m1_task,
    }
    if (
        new_weights != cfg["weights"]
        or ignore_zero_targets != cfg.get("ignore_zero_targets")
        or max_score_cap != cfg.get("max_score_cap")
    ):
        cfg["weights"] = new_weights
        cfg["ignore_zero_targets"] = ignore_zero_targets
        cfg["max_score_cap"] = max_score_cap
        write_payload_to_disk_and_cloud(SHARED_STORE, sync_cloud=True)

    st.caption(f"Tổng trọng số cấu hình: **{sum(new_weights.values())}%**")


# Tính toán dữ liệu tức thì
df_all = compute_all_kpi_dataframes(SHARED_STORE["records"], SHARED_STORE["config"])
df_filtered = df_all[df_all["department"].isin(selected_depts)].copy()
filtered_ids = set(df_filtered["id"].tolist())

with st.sidebar:
    st.markdown("### 📊 Sao lưu & Nạp Dữ liệu")
    excel_bytes = export_all_to_excel_bytes(df_all)
    if excel_bytes is not None:
        st.download_button(
            label="📥 Xuất File Excel Đầy Đủ (.xlsx)",
            data=excel_bytes,
            file_name=f"KPI_To_DLS_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
    st.download_button(
        label="⬇️ Tải File Sao lưu Dự phòng (.json)",
        data=json.dumps(SHARED_STORE, ensure_ascii=False, indent=2),
        file_name=f"KPI_To_DLS_Backup_{datetime.now().strftime('%Y%m%d')}.json",
        mime="application/json",
        use_container_width=True,
    )

    uploaded_file = st.file_uploader("⬆️ Khôi phục từ Excel (.xlsx) / (.json)", type=["xlsx", "json"])
    if uploaded_file is not None:
        if st.button("✅ Xác nhận Nạp Dữ liệu", type="primary", use_container_width=True):
            try:
                if uploaded_file.name.lower().endswith(".xlsx"):
                    import_records_from_excel(uploaded_file)
                else:
                    loaded = json.load(uploaded_file)
                    if "records" in loaded:
                        SHARED_STORE["records"] = loaded["records"]
                        if "config" in loaded:
                            SHARED_STORE["config"] = loaded["config"]
                        write_payload_to_disk_and_cloud(SHARED_STORE, sync_cloud=True)
                st.session_state.editor_ver += 1
                st.success("Đã nạp và lưu dữ liệu thành công!")
                st.rerun()
            except Exception as e:
                st.error(f"Lỗi nạp file: {e}")

    with st.expander("🔒 Vùng Quản trị (Sang tuần mới / Reset)"):
        st.warning("Chỉ Tổ trưởng sử dụng khi bắt đầu tuần làm việc mới.")
        confirm_unlock = st.checkbox("Tôi xác nhận muốn thao tác xóa số liệu tuần cũ")
        if st.button(
            "🧹 Đặt số Thứ 2–Thứ 6 về 0 (Giữ nguyên Chỉ tiêu M1)",
            disabled=not confirm_unlock,
            use_container_width=True,
        ):
            day_prefixes = ["hsba_t", "rx_out_t", "rx_in_t", "coun_out_t", "consult_t", "dose_t"]
            for r in SHARED_STORE["records"]:
                for p in day_prefixes:
                    for d_idx in ["2", "3", "4", "5", "6"]:
                        r[f"{p}{d_idx}"] = 0
            write_payload_to_disk_and_cloud(SHARED_STORE, sync_cloud=True)
            st.session_state.editor_ver += 1
            st.success("Đã đặt số thực hiện Thứ 2–Thứ 6 về 0 cho tuần mới!")
            st.rerun()


# ==============================================================================
# BANNER TIÊU ĐỀ & 7 THẺ CHỈ SỐ NHANH
# ==============================================================================
st.markdown(
    f"""
    <div class="dls-header">
        <div>
            <div class="dls-header-title">🏥 HỆ THỐNG QUẢN TRỊ & CHẤM ĐIỂM KPI — TỔ DƯỢC LÂM SÀNG</div>
            <div class="dls-header-sub">Nhập số trực tiếp trên bảng • Nhấn Enter hoặc chuyển ô là tự động lưu vĩnh viễn • Không tự xóa qua ngày</div>
        </div>
        <div class="dls-badge">👥 {len(SHARED_STORE["records"])} Dược sĩ Lâm sàng</div>
    </div>
    """,
    unsafe_allow_html=True,
)

c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
with c1:
    st.metric(
        "📁 M2: Bệnh án",
        f"{int(df_filtered['done_hsba'].sum()):,}/{int(df_filtered['target_hsba'].sum()):,}",
        f"TB: {df_filtered['score_m2_hsba'].mean():.1f} đ",
    )
with c2:
    st.metric(
        "💊 M3: Đơn Ngoại",
        f"{int(df_filtered['done_rx_out'].sum()):,}/{int(df_filtered['target_rx_out'].sum()):,}",
        f"TB: {df_filtered['score_m3_rx_out'].mean():.1f} đ",
    )
with c3:
    st.metric(
        "🏥 M4: Đơn Nội",
        f"{int(df_filtered['done_rx_in'].sum()):,}/{int(df_filtered['target_rx_in'].sum()):,}",
        f"TB: {df_filtered['score_m4_rx_in'].mean():.1f} đ",
    )
with c4:
    st.metric(
        "🗣️ M5: TV Ngoại",
        f"{int(df_filtered['done_counsel_out'].sum()):,}/{int(df_filtered['target_counsel_out'].sum()):,}",
        f"TB: {df_filtered['score_m5_counsel'].mean():.1f} đ",
    )
with c5:
    st.metric(
        "🩺 M6: Hội chẩn",
        f"{int(df_filtered['done_consult'].sum()):,}/{int(df_filtered['target_consult'].sum()):,}",
        f"TB: {df_filtered['score_m6_consult'].mean():.1f} đ",
    )
with c6:
    st.metric(
        "⚖️ M7: Chỉnh liều",
        f"{int(df_filtered['done_dose'].sum()):,}/{int(df_filtered['target_dose'].sum()):,}",
        f"TB: {df_filtered['score_m7_dose'].mean():.1f} đ",
    )
with c7:
    st.metric(
        "🏆 M8: KPI Tổ",
        f"{df_filtered['kpi_total'].mean():.1f} đ",
        f"{int((df_filtered['kpi_rank'] == 'Xuất sắc (A+)').sum())}/{len(df_filtered)} đạt A+",
    )


# ==============================================================================
# THANH 8 MODULE CHỨC NĂNG
# ==============================================================================
tab_m1, tab_m2, tab_m3, tab_m4, tab_m5, tab_m6, tab_m7, tab_m8 = st.tabs(
    [
        "🎯 M1: Giao Chỉ tiêu",
        "📁 M2: Bệnh án",
        "💊 M3: Duyệt Đơn Ngoại",
        "🏥 M4: Duyệt Đơn Nội",
        "🗣️ M5: Tư vấn Ngoại trú",
        "🩺 M6: Hội chẩn",
        "⚖️ M7: Chỉnh liều",
        "🏆 M8: Tổng hợp KPI",
    ]
)

# ------------------------------------------------------------------------------
# MODULE 1: GIAO CÁC CHỈ TIÊU VÀ CHỈ SỐ TỪNG PHẦN CHO TỪNG NGƯỜI
# ------------------------------------------------------------------------------
with tab_m1:
    st.markdown(
        '<div class="mod-title">🎯 M1: Giao Chỉ tiêu Tuần & Chỉ số Chất lượng Từng phần cho Từng Dược sĩ</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="mod-desc">Nhập hoặc điều chỉnh trực tiếp chỉ tiêu của từng người (nếu phần nào không phụ trách thì điền 0). Gõ số xong nhấn Enter là hệ thống tự lưu và cập nhật ngay sang các Module M2–M8.</div>',
        unsafe_allow_html=True,
    )

    m1_cols = [
        "id", "name", "department", "assigned_task",
        "target_hsba", "target_rx_out", "target_rx_in",
        "target_counsel_out", "target_consult", "target_dose",
        "doctor_consensus_pct", "intervention_success_pct", "task_quality_pct",
    ]
    df_m1_view = df_filtered[m1_cols].copy()
    st.session_state["_ids_m1"] = df_m1_view["id"].tolist()

    st.data_editor(
        df_m1_view,
        column_config={
            "id": st.column_config.TextColumn("Mã", disabled=True, width="small"),
            "name": st.column_config.TextColumn("Họ và tên Dược sĩ", width="medium"),
            "department": st.column_config.SelectboxColumn(
                "Khoa Lâm sàng Phụ trách", options=CLINICAL_DEPARTMENTS, width="medium"
            ),
            "assigned_task": st.column_config.TextColumn("Nhiệm vụ chuyên môn giao", width="large"),
            "target_hsba": st.column_config.NumberColumn("CT Bệnh án", min_value=0, max_value=2000, step=5),
            "target_rx_out": st.column_config.NumberColumn("CT Đơn Ngoại", min_value=0, max_value=5000, step=10),
            "target_rx_in": st.column_config.NumberColumn("CT Đơn Nội", min_value=0, max_value=5000, step=10),
            "target_counsel_out": st.column_config.NumberColumn("CT Tư vấn Ngoại", min_value=0, max_value=5000, step=5),
            "target_consult": st.column_config.NumberColumn("CT Hội chẩn", min_value=0, max_value=500, step=1),
            "target_dose": st.column_config.NumberColumn("CT Chỉnh liều", min_value=0, max_value=500, step=1),
            "doctor_consensus_pct": st.column_config.NumberColumn(
                "CS Đồng thuận BS (%)", min_value=0.0, max_value=100.0, step=1.0, format="%.1f%%"
            ),
            "intervention_success_pct": st.column_config.NumberColumn(
                "CS Chỉnh liều Đạt (%)", min_value=0.0, max_value=100.0, step=1.0, format="%.1f%%"
            ),
            "task_quality_pct": st.column_config.NumberColumn(
                "Điểm Nhiệm vụ (%)", min_value=0.0, max_value=120.0, step=1.0, format="%.1f%%"
            ),
        },
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key=f"ed_m1_{st.session_state.editor_ver}",
        on_change=make_atomic_callback(
            "ed_m1",
            "_ids_m1",
            int_cols=["target_hsba", "target_rx_out", "target_rx_in", "target_counsel_out", "target_consult", "target_dose"],
            float_cols=["doctor_consensus_pct", "intervention_success_pct", "task_quality_pct"],
        ),
    )

    st.markdown(
        f"""<div class="excel-summary-bar">
        📊 <b>Tổng Chỉ tiêu Tuần toàn Tổ:</b>
        Bệnh án: <b>{int(df_m1_view['target_hsba'].sum()):,}</b> &nbsp;|&nbsp;
        Đơn Ngoại trú: <b>{int(df_m1_view['target_rx_out'].sum()):,}</b> &nbsp;|&nbsp;
        Đơn Nội trú: <b>{int(df_m1_view['target_rx_in'].sum()):,}</b> &nbsp;|&nbsp;
        Tư vấn Ngoại trú: <b>{int(df_m1_view['target_counsel_out'].sum()):,}</b> &nbsp;|&nbsp;
        Hội chẩn: <b>{int(df_m1_view['target_consult'].sum()):,}</b> &nbsp;|&nbsp;
        Chỉnh liều: <b>{int(df_m1_view['target_dose'].sum()):,}</b>
        </div>""",
        unsafe_allow_html=True,
    )

# ------------------------------------------------------------------------------
# MODULE 2: KIỂM TRA HỒ SƠ BỆNH ÁN (THỨ 2 – THỨ 6)
# ------------------------------------------------------------------------------
with tab_m2:
    st.markdown(
        '<div class="mod-title">📁 M2: Kiểm tra Hồ sơ Bệnh án Hằng ngày (Thứ 2 – Thứ 6)</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="mod-desc">Nhập số lượng hồ sơ bệnh án rà soát từng ngày. Gõ số và nhấn Enter là tự động lưu.</div>',
        unsafe_allow_html=True,
    )

    m2_cols = [
        "id", "name", "department", "target_hsba",
        "hsba_t2", "hsba_t3", "hsba_t4", "hsba_t5", "hsba_t6",
        "done_hsba", "diff_hsba", "pct_hsba", "score_m2_hsba",
    ]
    df_m2_view = df_filtered[m2_cols].copy()
    st.session_state["_ids_m2"] = df_m2_view["id"].tolist()

    st.data_editor(
        df_m2_view,
        column_config={
            "id": st.column_config.TextColumn("Mã", disabled=True, width="small"),
            "name": st.column_config.TextColumn("Họ và tên Dược sĩ", disabled=True, width="medium"),
            "department": st.column_config.TextColumn("Khoa phụ trách", disabled=True, width="medium"),
            "target_hsba": st.column_config.NumberColumn("Chỉ tiêu Tuần", min_value=0, max_value=2000, step=5),
            "hsba_t2": st.column_config.NumberColumn("Thứ 2", min_value=0, max_value=500, step=1),
            "hsba_t3": st.column_config.NumberColumn("Thứ 3", min_value=0, max_value=500, step=1),
            "hsba_t4": st.column_config.NumberColumn("Thứ 4", min_value=0, max_value=500, step=1),
            "hsba_t5": st.column_config.NumberColumn("Thứ 5", min_value=0, max_value=500, step=1),
            "hsba_t6": st.column_config.NumberColumn("Thứ 6", min_value=0, max_value=500, step=1),
            "done_hsba": st.column_config.NumberColumn("Tổng Tuần", disabled=True),
            "diff_hsba": st.column_config.NumberColumn("Lệch (+/-)", disabled=True),
            "pct_hsba": st.column_config.NumberColumn("Đạt (%)", disabled=True, format="%.1f%%"),
            "score_m2_hsba": st.column_config.ProgressColumn(
                "Điểm M2", min_value=0, max_value=float(cfg["max_score_cap"]), format="%.1f đ"
            ),
        },
        use_container_width=True,
        hide_index=True,
        key=f"ed_m2_{st.session_state.editor_ver}",
        on_change=make_atomic_callback(
            "ed_m2",
            "_ids_m2",
            int_cols=["target_hsba", "hsba_t2", "hsba_t3", "hsba_t4", "hsba_t5", "hsba_t6"],
        ),
    )

    st.markdown(
        f"""<div class="excel-summary-bar">
        📊 <b>Tổng cộng M2 (Bệnh án):</b> Thực hiện: <b>{int(df_m2_view['done_hsba'].sum()):,}/{int(df_m2_view['target_hsba'].sum()):,}</b> &nbsp;|&nbsp;
        T2: <b>{int(df_m2_view['hsba_t2'].sum()):,}</b> • T3: <b>{int(df_m2_view['hsba_t3'].sum()):,}</b> •
        T4: <b>{int(df_m2_view['hsba_t4'].sum()):,}</b> • T5: <b>{int(df_m2_view['hsba_t5'].sum()):,}</b> •
        T6: <b>{int(df_m2_view['hsba_t6'].sum()):,}</b> &nbsp;|&nbsp;
        Điểm TB M2: <b>{df_m2_view['score_m2_hsba'].mean():.1f} đ</b>
        </div>""",
        unsafe_allow_html=True,
    )

# ------------------------------------------------------------------------------
# MODULE 3: DUYỆT ĐƠN THUỐC NGOẠI TRÚ RIÊNG (THỨ 2 – THỨ 6)
# ------------------------------------------------------------------------------
with tab_m3:
    st.markdown(
        '<div class="mod-title">💊 M3: Duyệt Đơn Thuốc Ngoại trú Hằng ngày (Thứ 2 – Thứ 6)</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="mod-desc">Module độc lập chuyên theo dõi số đơn thuốc ngoại trú được rà soát/duyệt từng ngày.</div>',
        unsafe_allow_html=True,
    )

    m3_cols = [
        "id", "name", "department", "target_rx_out",
        "rx_out_t2", "rx_out_t3", "rx_out_t4", "rx_out_t5", "rx_out_t6",
        "done_rx_out", "diff_rx_out", "pct_rx_out", "score_m3_rx_out",
    ]
    df_m3_view = df_filtered[m3_cols].copy()
    st.session_state["_ids_m3"] = df_m3_view["id"].tolist()

    st.data_editor(
        df_m3_view,
        column_config={
            "id": st.column_config.TextColumn("Mã", disabled=True, width="small"),
            "name": st.column_config.TextColumn("Họ và tên Dược sĩ", disabled=True, width="medium"),
            "department": st.column_config.TextColumn("Khoa phụ trách", disabled=True, width="medium"),
            "target_rx_out": st.column_config.NumberColumn("CT Đơn Ngoại", min_value=0, max_value=5000, step=10),
            "rx_out_t2": st.column_config.NumberColumn("Thứ 2", min_value=0, max_value=1500, step=1),
            "rx_out_t3": st.column_config.NumberColumn("Thứ 3", min_value=0, max_value=1500, step=1),
            "rx_out_t4": st.column_config.NumberColumn("Thứ 4", min_value=0, max_value=1500, step=1),
            "rx_out_t5": st.column_config.NumberColumn("Thứ 5", min_value=0, max_value=1500, step=1),
            "rx_out_t6": st.column_config.NumberColumn("Thứ 6", min_value=0, max_value=1500, step=1),
            "done_rx_out": st.column_config.NumberColumn("Tổng Đơn Ngoại", disabled=True),
            "diff_rx_out": st.column_config.NumberColumn("Lệch (+/-)", disabled=True),
            "pct_rx_out": st.column_config.NumberColumn("Đạt (%)", disabled=True, format="%.1f%%"),
            "score_m3_rx_out": st.column_config.ProgressColumn(
                "Điểm M3", min_value=0, max_value=float(cfg["max_score_cap"]), format="%.1f đ"
            ),
        },
        use_container_width=True,
        hide_index=True,
        key=f"ed_m3_{st.session_state.editor_ver}",
        on_change=make_atomic_callback(
            "ed_m3",
            "_ids_m3",
            int_cols=["target_rx_out", "rx_out_t2", "rx_out_t3", "rx_out_t4", "rx_out_t5", "rx_out_t6"],
        ),
    )

    st.markdown(
        f"""<div class="excel-summary-bar">
        📊 <b>Tổng cộng M3 (Duyệt Đơn Ngoại trú):</b> Thực hiện: <b>{int(df_m3_view['done_rx_out'].sum()):,}/{int(df_m3_view['target_rx_out'].sum()):,}</b> &nbsp;|&nbsp;
        T2: <b>{int(df_m3_view['rx_out_t2'].sum()):,}</b> • T3: <b>{int(df_m3_view['rx_out_t3'].sum()):,}</b> •
        T4: <b>{int(df_m3_view['rx_out_t4'].sum()):,}</b> • T5: <b>{int(df_m3_view['rx_out_t5'].sum()):,}</b> •
        T6: <b>{int(df_m3_view['rx_out_t6'].sum()):,}</b> &nbsp;|&nbsp;
        Điểm TB M3: <b>{df_m3_view['score_m3_rx_out'].mean():.1f} đ</b>
        </div>""",
        unsafe_allow_html=True,
    )

# ------------------------------------------------------------------------------
# MODULE 4: DUYỆT ĐƠN THUỐC NỘI TRÚ RIÊNG (THỨ 2 – THỨ 6)
# ------------------------------------------------------------------------------
with tab_m4:
    st.markdown(
        '<div class="mod-title">🏥 M4: Duyệt Đơn Thuốc / Phiếu Điều trị Nội trú Hằng ngày (Thứ 2 – Thứ 6)</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="mod-desc">Module độc lập chuyên theo dõi số đơn thuốc / y lệnh nội trú được rà soát hằng ngày.</div>',
        unsafe_allow_html=True,
    )

    m4_cols = [
        "id", "name", "department", "target_rx_in",
        "rx_in_t2", "rx_in_t3", "rx_in_t4", "rx_in_t5", "rx_in_t6",
        "done_rx_in", "diff_rx_in", "pct_rx_in", "score_m4_rx_in",
    ]
    df_m4_view = df_filtered[m4_cols].copy()
    st.session_state["_ids_m4"] = df_m4_view["id"].tolist()

    st.data_editor(
        df_m4_view,
        column_config={
            "id": st.column_config.TextColumn("Mã", disabled=True, width="small"),
            "name": st.column_config.TextColumn("Họ và tên Dược sĩ", disabled=True, width="medium"),
            "department": st.column_config.TextColumn("Khoa phụ trách", disabled=True, width="medium"),
            "target_rx_in": st.column_config.NumberColumn("CT Đơn Nội", min_value=0, max_value=5000, step=10),
            "rx_in_t2": st.column_config.NumberColumn("Thứ 2", min_value=0, max_value=1500, step=1),
            "rx_in_t3": st.column_config.NumberColumn("Thứ 3", min_value=0, max_value=1500, step=1),
            "rx_in_t4": st.column_config.NumberColumn("Thứ 4", min_value=0, max_value=1500, step=1),
            "rx_in_t5": st.column_config.NumberColumn("Thứ 5", min_value=0, max_value=1500, step=1),
            "rx_in_t6": st.column_config.NumberColumn("Thứ 6", min_value=0, max_value=1500, step=1),
            "done_rx_in": st.column_config.NumberColumn("Tổng Đơn Nội", disabled=True),
            "diff_rx_in": st.column_config.NumberColumn("Lệch (+/-)", disabled=True),
            "pct_rx_in": st.column_config.NumberColumn("Đạt (%)", disabled=True, format="%.1f%%"),
            "score_m4_rx_in": st.column_config.ProgressColumn(
                "Điểm M4", min_value=0, max_value=float(cfg["max_score_cap"]), format="%.1f đ"
            ),
        },
        use_container_width=True,
        hide_index=True,
        key=f"ed_m4_{st.session_state.editor_ver}",
        on_change=make_atomic_callback(
            "ed_m4",
            "_ids_m4",
            int_cols=["target_rx_in", "rx_in_t2", "rx_in_t3", "rx_in_t4", "rx_in_t5", "rx_in_t6"],
        ),
    )

    st.markdown(
        f"""<div class="excel-summary-bar">
        📊 <b>Tổng cộng M4 (Duyệt Đơn Nội trú):</b> Thực hiện: <b>{int(df_m4_view['done_rx_in'].sum()):,}/{int(df_m4_view['target_rx_in'].sum()):,}</b> &nbsp;|&nbsp;
        T2: <b>{int(df_m4_view['rx_in_t2'].sum()):,}</b> • T3: <b>{int(df_m4_view['rx_in_t3'].sum()):,}</b> •
        T4: <b>{int(df_m4_view['rx_in_t4'].sum()):,}</b> • T5: <b>{int(df_m4_view['rx_in_t5'].sum()):,}</b> •
        T6: <b>{int(df_m4_view['rx_in_t6'].sum()):,}</b> &nbsp;|&nbsp;
        Điểm TB M4: <b>{df_m4_view['score_m4_rx_in'].mean():.1f} đ</b>
        </div>""",
        unsafe_allow_html=True,
    )

# ------------------------------------------------------------------------------
# MODULE 5: TƯ VẤN ĐƠN THUỐC NGOẠI TRÚ THEO HẰNG NGÀY (THỨ 2 – THỨ 6)
# ------------------------------------------------------------------------------
with tab_m5:
    st.markdown(
        '<div class="mod-title">🗣️ M5: Tư vấn Đơn Thuốc Ngoại trú Theo Hằng ngày (Thứ 2 – Thứ 6)</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="mod-desc">Nhập trực tiếp số đơn thuốc ngoại trú đã tư vấn từng ngày (Thứ 2 đến Thứ 6); hệ thống tự tính Tổng đơn tư vấn tuần, Tỷ lệ đạt Chỉ tiêu và Tỷ lệ % so với Đơn ngoại trú đã duyệt.</div>',
        unsafe_allow_html=True,
    )

    m5_cols = [
        "id", "name", "department", "target_counsel_out",
        "coun_out_t2", "coun_out_t3", "coun_out_t4", "coun_out_t5", "coun_out_t6",
        "done_counsel_out", "done_rx_out", "rate_counsel_vs_rx_out", "pct_counsel_out", "score_m5_counsel",
    ]
    df_m5_view = df_filtered[m5_cols].copy()
    st.session_state["_ids_m5"] = df_m5_view["id"].tolist()

    st.data_editor(
        df_m5_view,
        column_config={
            "id": st.column_config.TextColumn("Mã", disabled=True, width="small"),
            "name": st.column_config.TextColumn("Họ và tên Dược sĩ", disabled=True, width="medium"),
            "department": st.column_config.TextColumn("Khoa phụ trách", disabled=True, width="medium"),
            "target_counsel_out": st.column_config.NumberColumn("CT Tư vấn Tuần", min_value=0, max_value=5000, step=5),
            "coun_out_t2": st.column_config.NumberColumn("TV Thứ 2", min_value=0, max_value=1500, step=1),
            "coun_out_t3": st.column_config.NumberColumn("TV Thứ 3", min_value=0, max_value=1500, step=1),
            "coun_out_t4": st.column_config.NumberColumn("TV Thứ 4", min_value=0, max_value=1500, step=1),
            "coun_out_t5": st.column_config.NumberColumn("TV Thứ 5", min_value=0, max_value=1500, step=1),
            "coun_out_t6": st.column_config.NumberColumn("TV Thứ 6", min_value=0, max_value=1500, step=1),
            "done_counsel_out": st.column_config.NumberColumn("Tổng TV Tuần", disabled=True),
            "done_rx_out": st.column_config.NumberColumn("Đơn Ngoại (M3)", disabled=True),
            "rate_counsel_vs_rx_out": st.column_config.NumberColumn("% TV/Đơn Ngoại", disabled=True, format="%.1f%%"),
            "pct_counsel_out": st.column_config.NumberColumn("Đạt CT (%)", disabled=True, format="%.1f%%"),
            "score_m5_counsel": st.column_config.ProgressColumn(
                "Điểm M5", min_value=0, max_value=float(cfg["max_score_cap"]), format="%.1f đ"
            ),
        },
        use_container_width=True,
        hide_index=True,
        key=f"ed_m5_{st.session_state.editor_ver}",
        on_change=make_atomic_callback(
            "ed_m5",
            "_ids_m5",
            int_cols=["target_counsel_out", "coun_out_t2", "coun_out_t3", "coun_out_t4", "coun_out_t5", "coun_out_t6"],
        ),
    )

    st.markdown(
        f"""<div class="excel-summary-bar">
        📊 <b>Tổng cộng M5 (Tư vấn Đơn Ngoại trú):</b> Thực hiện: <b>{int(df_m5_view['done_counsel_out'].sum()):,}/{int(df_m5_view['target_counsel_out'].sum()):,}</b> &nbsp;|&nbsp;
        T2: <b>{int(df_m5_view['coun_out_t2'].sum()):,}</b> • T3: <b>{int(df_m5_view['coun_out_t3'].sum()):,}</b> •
        T4: <b>{int(df_m5_view['coun_out_t4'].sum()):,}</b> • T5: <b>{int(df_m5_view['coun_out_t5'].sum()):,}</b> •
        T6: <b>{int(df_m5_view['coun_out_t6'].sum()):,}</b> &nbsp;|&nbsp;
        Điểm TB M5: <b>{df_m5_view['score_m5_counsel'].mean():.1f} đ</b>
        </div>""",
        unsafe_allow_html=True,
    )

# ------------------------------------------------------------------------------
# MODULE 6: HỘI CHẨN LÂM SÀNG (THỨ 2 – THỨ 6)
# ------------------------------------------------------------------------------
with tab_m6:
    st.markdown(
        '<div class="mod-title">🩺 M6: Hội chẩn Lâm sàng Hằng ngày (Thứ 2 – Thứ 6)</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="mod-desc">Điểm M6 = 65% × (% Hoàn thành số ca hội chẩn) + 35% × (% Bác sĩ đồng thuận).</div>',
        unsafe_allow_html=True,
    )

    m6_cols = [
        "id", "name", "department", "target_consult",
        "consult_t2", "consult_t3", "consult_t4", "consult_t5", "consult_t6",
        "done_consult", "pct_consult", "doctor_consensus_pct", "score_m6_consult",
    ]
    df_m6_view = df_filtered[m6_cols].copy()
    st.session_state["_ids_m6"] = df_m6_view["id"].tolist()

    st.data_editor(
        df_m6_view,
        column_config={
            "id": st.column_config.TextColumn("Mã", disabled=True, width="small"),
            "name": st.column_config.TextColumn("Họ và tên Dược sĩ", disabled=True, width="medium"),
            "department": st.column_config.TextColumn("Khoa phụ trách", disabled=True, width="medium"),
            "target_consult": st.column_config.NumberColumn("CT Hội chẩn", min_value=0, max_value=500, step=1),
            "consult_t2": st.column_config.NumberColumn("Thứ 2", min_value=0, max_value=100, step=1),
            "consult_t3": st.column_config.NumberColumn("Thứ 3", min_value=0, max_value=100, step=1),
            "consult_t4": st.column_config.NumberColumn("Thứ 4", min_value=0, max_value=100, step=1),
            "consult_t5": st.column_config.NumberColumn("Thứ 5", min_value=0, max_value=100, step=1),
            "consult_t6": st.column_config.NumberColumn("Thứ 6", min_value=0, max_value=100, step=1),
            "done_consult": st.column_config.NumberColumn("Tổng Ca", disabled=True),
            "pct_consult": st.column_config.NumberColumn("Đạt SL (%)", disabled=True, format="%.1f%%"),
            "doctor_consensus_pct": st.column_config.NumberColumn(
                "BS Đồng thuận (%)", min_value=0.0, max_value=100.0, step=0.5, format="%.1f%%"
            ),
            "score_m6_consult": st.column_config.ProgressColumn(
                "Điểm M6", min_value=0, max_value=float(cfg["max_score_cap"]), format="%.1f đ"
            ),
        },
        use_container_width=True,
        hide_index=True,
        key=f"ed_m6_{st.session_state.editor_ver}",
        on_change=make_atomic_callback(
            "ed_m6",
            "_ids_m6",
            int_cols=["target_consult", "consult_t2", "consult_t3", "consult_t4", "consult_t5", "consult_t6"],
            float_cols=["doctor_consensus_pct"],
        ),
    )

    st.markdown(
        f"""<div class="excel-summary-bar">
        📊 <b>Tổng cộng M6 (Hội chẩn):</b> Thực hiện: <b>{int(df_m6_view['done_consult'].sum()):,}/{int(df_m6_view['target_consult'].sum()):,}</b> &nbsp;|&nbsp;
        BS đồng thuận TB: <b>{df_m6_view['doctor_consensus_pct'].mean():.1f}%</b> &nbsp;|&nbsp;
        Điểm TB M6: <b>{df_m6_view['score_m6_consult'].mean():.1f} đ</b>
        </div>""",
        unsafe_allow_html=True,
    )

# ------------------------------------------------------------------------------
# MODULE 7: CAN THIỆP & HIỆU CHỈNH LIỀU (THỨ 2 – THỨ 6)
# ------------------------------------------------------------------------------
with tab_m7:
    st.markdown(
        '<div class="mod-title">⚖️ M7: Can thiệp & Hiệu chỉnh liều Hằng ngày (Thứ 2 – Thứ 6)</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="mod-desc">Điểm M7 = 65% × (% Hoàn thành số ca chỉnh liều) + 35% × (% Can thiệp thành công).</div>',
        unsafe_allow_html=True,
    )

    m7_cols = [
        "id", "name", "department", "target_dose",
        "dose_t2", "dose_t3", "dose_t4", "dose_t5", "dose_t6",
        "done_dose", "pct_dose", "intervention_success_pct", "score_m7_dose",
    ]
    df_m7_view = df_filtered[m7_cols].copy()
    st.session_state["_ids_m7"] = df_m7_view["id"].tolist()

    st.data_editor(
        df_m7_view,
        column_config={
            "id": st.column_config.TextColumn("Mã", disabled=True, width="small"),
            "name": st.column_config.TextColumn("Họ và tên Dược sĩ", disabled=True, width="medium"),
            "department": st.column_config.TextColumn("Khoa phụ trách", disabled=True, width="medium"),
            "target_dose": st.column_config.NumberColumn("CT Chỉnh liều", min_value=0, max_value=500, step=1),
            "dose_t2": st.column_config.NumberColumn("Thứ 2", min_value=0, max_value=100, step=1),
            "dose_t3": st.column_config.NumberColumn("Thứ 3", min_value=0, max_value=100, step=1),
            "dose_t4": st.column_config.NumberColumn("Thứ 4", min_value=0, max_value=100, step=1),
            "dose_t5": st.column_config.NumberColumn("Thứ 5", min_value=0, max_value=100, step=1),
            "dose_t6": st.column_config.NumberColumn("Thứ 6", min_value=0, max_value=100, step=1),
            "done_dose": st.column_config.NumberColumn("Tổng Ca", disabled=True),
            "pct_dose": st.column_config.NumberColumn("Đạt SL (%)", disabled=True, format="%.1f%%"),
            "intervention_success_pct": st.column_config.NumberColumn(
                "Thành công (%)", min_value=0.0, max_value=100.0, step=0.5, format="%.1f%%"
            ),
            "score_m7_dose": st.column_config.ProgressColumn(
                "Điểm M7", min_value=0, max_value=float(cfg["max_score_cap"]), format="%.1f đ"
            ),
        },
        use_container_width=True,
        hide_index=True,
        key=f"ed_m7_{st.session_state.editor_ver}",
        on_change=make_atomic_callback(
            "ed_m7",
            "_ids_m7",
            int_cols=["target_dose", "dose_t2", "dose_t3", "dose_t4", "dose_t5", "dose_t6"],
            float_cols=["intervention_success_pct"],
        ),
    )

    st.markdown(
        f"""<div class="excel-summary-bar">
        📊 <b>Tổng cộng M7 (Chỉnh liều):</b> Thực hiện: <b>{int(df_m7_view['done_dose'].sum()):,}/{int(df_m7_view['target_dose'].sum()):,}</b> &nbsp;|&nbsp;
        Tỷ lệ thành công TB: <b>{df_m7_view['intervention_success_pct'].mean():.1f}%</b> &nbsp;|&nbsp;
        Điểm TB M7: <b>{df_m7_view['score_m7_dose'].mean():.1f} đ</b>
        </div>""",
        unsafe_allow_html=True,
    )

# ------------------------------------------------------------------------------
# MODULE 8: TỔNG HỢP KPI TOÀN DIỆN & XẾP LOẠI
# ------------------------------------------------------------------------------
with tab_m8:
    st.markdown(
        '<div class="mod-title">🏆 M8: Bảng Tổng hợp Điểm Tất cả Chỉ tiêu & Xếp loại KPI Cá nhân</div>',
        unsafe_allow_html=True,
    )
    w = cfg["weights"]
    st.markdown(
        f'<div class="mod-desc">Trọng số: M2 (Bệnh án): <b>{w["w_m2_hsba"]}%</b> • M3 (Đơn Ngoại): <b>{w["w_m3_rx_out"]}%</b> • '
        f'M4 (Đơn Nội): <b>{w["w_m4_rx_in"]}%</b> • M5 (TV Ngoại): <b>{w["w_m5_counsel"]}%</b> • '
        f'M6 (Hội chẩn): <b>{w["w_m6_consult"]}%</b> • M7 (Chỉnh liều): <b>{w["w_m7_dose"]}%</b> • '
        f'M1 (Nhiệm vụ): <b>{w["w_m1_task"]}%</b></div>',
        unsafe_allow_html=True,
    )

    df_kpi_sorted = df_filtered.sort_values(by="kpi_total", ascending=False).copy()

    st.dataframe(
        df_kpi_sorted[
            [
                "id", "name", "department",
                "done_hsba", "score_m2_hsba",
                "done_rx_out", "score_m3_rx_out",
                "done_rx_in", "score_m4_rx_in",
                "done_counsel_out", "score_m5_counsel",
                "done_consult", "score_m6_consult",
                "done_dose", "score_m7_dose",
                "task_quality_pct", "kpi_total", "kpi_rank",
            ]
        ],
        column_config={
            "id": "Mã",
            "name": "Họ và tên Dược sĩ",
            "department": "Khoa phụ trách",
            "done_hsba": st.column_config.NumberColumn("Số BA"),
            "score_m2_hsba": st.column_config.NumberColumn("Điểm M2", format="%.1f"),
            "done_rx_out": st.column_config.NumberColumn("Đơn Ngoại"),
            "score_m3_rx_out": st.column_config.NumberColumn("Điểm M3", format="%.1f"),
            "done_rx_in": st.column_config.NumberColumn("Đơn Nội"),
            "score_m4_rx_in": st.column_config.NumberColumn("Điểm M4", format="%.1f"),
            "done_counsel_out": st.column_config.NumberColumn("TV Ngoại"),
            "score_m5_counsel": st.column_config.NumberColumn("Điểm M5", format="%.1f"),
            "done_consult": st.column_config.NumberColumn("Ca HC"),
            "score_m6_consult": st.column_config.NumberColumn("Điểm M6", format="%.1f"),
            "done_dose": st.column_config.NumberColumn("Ca CL"),
            "score_m7_dose": st.column_config.NumberColumn("Điểm M7", format="%.1f"),
            "task_quality_pct": st.column_config.NumberColumn("Điểm M1", format="%.1f"),
            "kpi_total": st.column_config.ProgressColumn(
                "TỔNG KPI", min_value=0, max_value=float(cfg["max_score_cap"]), format="%.1f đ"
            ),
            "kpi_rank": "Xếp loại",
        },
        use_container_width=True,
        hide_index=True,
    )

    ch1, ch2 = st.columns(2)
    with ch1:
        st.markdown('<div class="mod-title">📈 Biểu đồ Tổng Điểm KPI theo Dược sĩ</div>', unsafe_allow_html=True)
        chart_df = df_kpi_sorted.set_index("name")[
            ["kpi_total", "score_m2_hsba", "score_m3_rx_out", "score_m4_rx_in", "score_m5_counsel"]
        ]
        chart_df.columns = ["Tổng KPI", "M2: Bệnh án", "M3: Đơn Ngoại", "M4: Đơn Nội", "M5: Tư vấn Ngoại"]
        st.bar_chart(chart_df, height=380)

    with ch2:
        st.markdown('<div class="mod-title">📊 Điểm KPI Trung bình theo Khoa Lâm sàng</div>', unsafe_allow_html=True)
        dept_kpi = (
            df_kpi_sorted.groupby("department", as_index=True)
            .agg(Điểm_KPI_TB=("kpi_total", "mean"))
            .round(1)
        )
        st.bar_chart(dept_kpi, height=380)
