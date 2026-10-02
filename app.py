"""
祿野營地訂位管理系統 (Lu-Ye Campsite Booking Management System)
===========================================================
- 營區區域：雨棚區 (8帳)、草地區 (4帳)、通舖房間 (2間)
- 安全機制：4 位數 PIN 碼登入與管理員權限控制
- 核心功能：訂位管理、三大核心索引搜尋、Excel/CSV 匯入匯出、操作稽核日誌 (Audit Log)
- 雲端備份：Google Drive API 自動同步備份 (預設目標 DVCG 資料夾)，具備平滑降級本地機制
"""

import io
import json
import os
import sqlite3
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# ==============================================================================
# 1. 系統常數與營區規格設定
# ==============================================================================
APP_TITLE = "祿野營地訂位管理系統"
APP_SUBTITLE = "專業・溫馨・便利的露營營位管理後台"

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
DB_PATH = os.path.join(DATA_DIR, "luye_camping.db")
ADMIN_CONFIG_PATH = os.path.join(DATA_DIR, "admins.json")
SERVICE_ACCOUNT_PATH = os.path.join(DATA_DIR, "service_account.json")

# Google Drive 預設目標資料夾 (DVCG)
DEFAULT_GDRIVE_FOLDER_ID = "1JFzbSQuolGniXeAeeaX3okbUZPQ-Rzok"

# 營區三大區域規格與費用設定
CAMP_AREAS: Dict[str, Dict[str, Any]] = {
    "雨棚區": {
        "capacity": 8,
        "unit": "帳",
        "spec": "營位寬 4 米深 6 米（含 1 車）",
        "price_per_night": 1200,
        "night_rush_price": 600,  # 50%
        "color": "#2E7D32",
    },
    "草地區": {
        "capacity": 4,
        "unit": "帳",
        "spec": "開闊天然草地（含 1 車）",
        "price_per_night": 1200,
        "night_rush_price": 600,  # 50%
        "color": "#1565C0",
    },
    "通舖房間": {
        "capacity": 2,
        "unit": "間",
        "spec": "每間容納 4 人，含客廳與廁所，無衛浴，含戶外雨棚區（含 1 車）",
        "price_per_night": 2000,
        "night_rush_price": 1000,  # 50%
        "color": "#E65100",
    },
}

CUSTOMER_TAGS = ["一般", "好客人", "需要注意的客人"]

TAG_BADGE_STYLE = {
    "好客人": "background-color: #E8F5E9; color: #1B5E20; border: 1px solid #A5D6A7; padding: 2px 8px; border-radius: 12px; font-weight: 600; font-size: 0.85rem;",
    "需要注意的客人": "background-color: #FFEBEE; color: #B71C1C; border: 1px solid #EF9A9A; padding: 2px 8px; border-radius: 12px; font-weight: 600; font-size: 0.85rem;",
    "一般": "background-color: #E3F2FD; color: #0D47A1; border: 1px solid #90CAF9; padding: 2px 8px; border-radius: 12px; font-weight: 600; font-size: 0.85rem;",
}


# ==============================================================================
# 2. 資料庫與儲存層初始化
# ==============================================================================
def ensure_directories():
    """確保資料儲存目錄存在"""
    os.makedirs(DATA_DIR, exist_ok=True)


def init_admins():
    """初始化管理人 PIN 碼設定檔 (預設 0218: KC, 0311: Aileen)"""
    ensure_directories()
    if not os.path.exists(ADMIN_CONFIG_PATH):
        default_admins = {
            "0218": "KC",
            "0311": "Aileen",
        }
        with open(ADMIN_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(default_admins, f, ensure_ascii=False, indent=2)


def get_admins() -> Dict[str, str]:
    """取得所有管理員資訊 (PIN -> 名稱)"""
    init_admins()
    try:
        with open(ADMIN_CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        st.error(f"讀取管理人設定失敗: {e}")
        return {"0218": "KC", "0311": "Aileen"}


def save_admins(admins: Dict[str, str]) -> bool:
    """儲存管理員資訊"""
    ensure_directories()
    try:
        with open(ADMIN_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(admins, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        st.error(f"寫入管理人設定失敗: {e}")
        return False


def get_db_connection():
    """取得 SQLite 資料庫連線"""
    ensure_directories()
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """初始化資料庫表單"""
    conn = get_db_connection()
    cursor = conn.cursor()
    # 訂單資料表
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS bookings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            booking_code TEXT UNIQUE,
            check_in_date TEXT NOT NULL,
            check_out_date TEXT NOT NULL,
            name TEXT NOT NULL,
            phone TEXT NOT NULL,
            area TEXT NOT NULL,
            quantity INTEGER NOT NULL DEFAULT 1,
            night_rush INTEGER NOT NULL DEFAULT 0,
            visit_count INTEGER NOT NULL DEFAULT 1,
            tag TEXT NOT NULL DEFAULT '一般',
            notes TEXT DEFAULT '',
            total_amount INTEGER NOT NULL DEFAULT 0,
            created_by TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )

    # 稽核操作記錄表 (Audit Log)
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            operator TEXT NOT NULL,
            action TEXT NOT NULL,
            details TEXT NOT NULL
        )
        """
    )
    conn.commit()
    conn.close()


def add_audit_log(operator: str, action: str, details: str):
    """新增一筆操作紀錄"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute(
            "INSERT INTO audit_logs (timestamp, operator, action, details) VALUES (?, ?, ?, ?)",
            (now_str, operator, action, details),
        )
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[Audit Log Error] {e}")


def get_audit_logs(limit: int = 200) -> pd.DataFrame:
    """取得稽核操作紀錄清單"""
    try:
        conn = get_db_connection()
        query = "SELECT timestamp as '時間戳記', operator as '操作者', action as '操作動作', details as '具體細節' FROM audit_logs ORDER BY id DESC LIMIT ?"
        df = pd.read_sql_query(query, conn, params=(limit,))
        conn.close()
        return df
    except Exception as e:
        st.error(f"讀取操作日誌失敗: {e}")
        return pd.DataFrame()


# ==============================================================================
# 3. 費用計算與訂單邏輯輔助函式
# ==============================================================================
def calculate_booking_fee(
    area: str, quantity: int, check_in: date, check_out: date, night_rush: bool
) -> Tuple[int, int, int, int]:
    """
    計算營位訂位費用
    回傳: (總金額, 住宿晚數, 營位費用, 夜衝費用)
    """
    if area not in CAMP_AREAS:
        return (0, 1, 0, 0)

    area_info = CAMP_AREAS[area]
    price_per_night = area_info["price_per_night"]
    night_rush_price = area_info["night_rush_price"]

    nights = max(1, (check_out - check_in).days)
    base_fee = price_per_night * quantity * nights
    night_rush_fee = night_rush_price * quantity if night_rush else 0
    total_fee = base_fee + night_rush_fee

    return (total_fee, nights, base_fee, night_rush_fee)


def generate_booking_code(check_in: date) -> str:
    """產生唯一直覺的訂單編號 (如 LY-20261001-0812)"""
    rand_suffix = datetime.now().strftime("%H%M%S")[-4:]
    return f"LY-{check_in.strftime('%Y%m%d')}-{rand_suffix}"


def query_historical_stats(phone: str, name: str = "") -> Tuple[int, str]:
    """根據電話 (與姓名) 查詢歷史來訪次數與歷史標籤"""
    if not phone:
        return 1, "一般"
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*), tag FROM bookings WHERE phone = ? GROUP BY tag", (phone.strip(),))
        row = cursor.fetchone()
        cursor.execute("SELECT COUNT(*) FROM bookings WHERE phone = ?", (phone.strip(),))
        total_count = cursor.fetchone()[0]
        conn.close()

        tag = row[1] if row else "一般"
        return total_count, tag
    except Exception:
        return 1, "一般"


def format_phone_number(raw_val) -> str:
    """清理並格式化電話號碼：去除小數點 .0，轉為純字串並補齊開頭 0"""
    if pd.isna(raw_val) or raw_val is None:
        return ""
    val_str = str(raw_val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    # 清除破折號與空格
    val_str = val_str.replace("-", "").replace(" ", "")
    if val_str.lower() in ["nan", "none", "null"]:
        return ""
    # 若為純數字且長度為9碼(通常遺失開頭0)，自動補齊開頭 '0'
    if val_str.isdigit() and len(val_str) == 9 and val_str.startswith("9"):
        val_str = "0" + val_str
    elif val_str.isdigit() and 0 < len(val_str) <= 9 and not val_str.startswith("0"):
        val_str = "0" + val_str
    return val_str


# ==============================================================================
# 4. Google Drive 雲端同步模組 (含自動降級機制)
# ==============================================================================
class GoogleDriveSync:
    """Google Drive API 存取與備份同步器，具備完整的自動降級防呆機制"""

    def __init__(self, folder_id: str = DEFAULT_GDRIVE_FOLDER_ID):
        self.folder_id = folder_id
        self.service = None
        self.is_connected = False
        self.status_message = "尚未初始化"
        self._initialize_service()

    def _initialize_service(self):
        """嘗試透過 Service Account 金鑰或 Streamlit Secrets 進行認證"""
        try:
            from google.oauth2 import service_account
            from googleapiclient.discovery import build

            creds = None
            # 優先嘗試 1: 本地 service_account.json
            if os.path.exists(SERVICE_ACCOUNT_PATH):
                creds = service_account.Credentials.from_service_account_file(
                    SERVICE_ACCOUNT_PATH,
                    scopes=["https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"],
                )
                self.status_message = "已透過本地 service_account.json 認證"

            # 嘗試 2: Streamlit secrets
            elif hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
                creds_info = dict(st.secrets["gcp_service_account"])
                creds = service_account.Credentials.from_service_account_info(
                    creds_info,
                    scopes=["https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"],
                )
                self.status_message = "已透過 Streamlit Secrets 認證"

            if creds:
                self.service = build("drive", "v3", credentials=creds, cache_discovery=False)
                self.is_connected = True
            else:
                self.is_connected = False
                self.status_message = "未偵測到 Google 憑證金鑰，系統已平滑降級為「本地檔案模式 (SQLite/JSON)」"
        except Exception as e:
            self.is_connected = False
            self.status_message = f"雲端連線嘗試失敗 (已平滑降級為本地儲存): {str(e)}"

    def upload_file_to_drive(self, file_content_bytes: bytes, file_name: str, mime_type: str) -> Tuple[bool, str]:
        """上傳檔案至目標 Google Drive 資料夾"""
        if not self.is_connected or not self.service:
            return False, f"Google Drive 未連線：{self.status_message}"

        try:
            from googleapiclient.http import MediaIoBaseUpload

            file_metadata = {
                "name": file_name,
                "parents": [self.folder_id],
            }
            media = MediaIoBaseUpload(io.BytesIO(file_content_bytes), mimetype=mime_type, resumable=True)

            file = self.service.files().create(body=file_metadata, media_body=media, fields="id, name, webViewLink").execute()
            link = file.get("webViewLink", "")
            return True, f"檔案成功備份至雲端！檔案 ID: {file.get('id')}，檔名: {file_name}"
        except Exception as e:
            return False, f"上傳至 Google Drive 失敗: {str(e)}"

    def backup_database_now(self, operator: str = "系統") -> Tuple[bool, str]:
        """
        將目前所有訂單與 Log 打包匯出成 Excel 與 JSON，並嘗試同步到 Google Drive 與本地備份區
        """
        try:
            conn = get_db_connection()
            df_bookings = pd.read_sql_query("SELECT * FROM bookings ORDER BY check_in_date ASC", conn)
            df_logs = pd.read_sql_query("SELECT * FROM audit_logs ORDER BY id DESC", conn)
            conn.close()

            timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_filename = f"祿野營地資料庫完整備份_{timestamp_str}.xlsx"

            # 產生 Excel Bytes
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                df_bookings.to_excel(writer, sheet_name="露營訂單明細", index=False)
                df_logs.to_excel(writer, sheet_name="操作稽核日誌", index=False)
            excel_bytes = excel_buffer.getvalue()

            # 本地備份檔案存檔
            local_backup_path = os.path.join(DATA_DIR, excel_filename)
            with open(local_backup_path, "wb") as f:
                f.write(excel_bytes)

            # 嘗試 Google Drive 雲端同步
            if self.is_connected:
                success, msg = self.upload_file_to_drive(
                    excel_bytes,
                    excel_filename,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                if success:
                    add_audit_log(operator, "雲端同步備份", f"已成功同步至 Google Drive (DVCG 資料夾)：{excel_filename}")
                    return True, f"雲端同步備份成功！本地亦已保存副本。\n{msg}"
                else:
                    add_audit_log(operator, "雲端同步異常", f"雲端上傳失敗，已保存在本地備份：{local_backup_path}")
                    return False, f"本地備份成功，但雲端備份失敗：{msg}"
            else:
                add_audit_log(operator, "本地備份", f"雲端未啟用，已生成本地備份：{local_backup_path}")
                return True, f"本地資料備份成功！（已儲存至 {local_backup_path}）。因未配置 Google 憑證，雲端同步處於離線狀態。"
        except Exception as e:
            return False, f"備份過程發生錯誤: {str(e)}"


# 建立全域或快取 GDriveSync 實例
@st.cache_resource
def get_gdrive_sync():
    return GoogleDriveSync()


def trigger_auto_sync(operator: str):
    """在訂單、管理員或日誌變動時觸發自動備份"""
    try:
        sync_tool = get_gdrive_sync()
        sync_tool.backup_database_now(operator=operator)
    except Exception as e:
        print(f"[Auto Sync Warning] {e}")


# ==============================================================================
# 5. 自定義 CSS 樣式與頁面配置
# ==============================================================================
def apply_custom_styles():
    st.set_page_config(
        page_title="祿野營地訂位管理系統",
        page_icon="⛺",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.markdown(
        """
        <style>
        /* 主體背景與全域樣式 */
        .main {
            background-color: #F8F9FA;
        }
        
        /* 頂部標題區美化 */
        .app-header {
            background: linear-gradient(135deg, #1B5E20 0%, #2E7D32 50%, #43A047 100%);
            color: white;
            padding: 24px;
            border-radius: 12px;
            margin-bottom: 24px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        }
        .app-header h1 {
            color: #FFFFFF !important;
            margin: 0;
            font-size: 2.1rem;
            font-weight: 700;
        }
        .app-header p {
            color: #E8F5E9;
            margin-top: 6px;
            margin-bottom: 0;
            font-size: 1.05rem;
        }

        /* 資訊卡片樣式 */
        .luye-card {
            background: white;
            border-radius: 10px;
            padding: 18px 20px;
            border: 1px solid #E0E0E0;
            box-shadow: 0 2px 6px rgba(0,0,0,0.04);
            margin-bottom: 16px;
            transition: all 0.2s ease-in-out;
        }
        .luye-card:hover {
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
            border-color: #BDBDBD;
        }

        /* 營區容量徽章 */
        .capacity-pill {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 16px;
            font-weight: 600;
            font-size: 0.9rem;
            margin-right: 8px;
        }

        /* 表格與按鈕細節優化 */
        .stButton button {
            border-radius: 8px;
            font-weight: 600;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ==============================================================================
# 6. 安全登入與管理員狀態管理
# ==============================================================================
def check_authentication() -> bool:
    """檢查使用者是否已登入"""
    if "authenticated_user" not in st.session_state:
        st.session_state["authenticated_user"] = None

    return st.session_state["authenticated_user"] is not None


def render_login_view():
    """渲染 PIN 碼登入畫面"""
    col1, col2, col3 = st.columns([1, 1.8, 1])
    with col2:
        st.markdown(
            """
            <div style="text-align: center; padding: 30px 20px; background: white; border-radius: 16px; box-shadow: 0 8px 24px rgba(0,0,0,0.08); border-top: 6px solid #2E7D32;">
                <h1 style="color: #1B5E20; font-size: 2.2rem; margin-bottom: 6px;">⛺ 祿野營地</h1>
                <h3 style="color: #424242; font-weight: 500; margin-top: 0;">訂位管理系統</h3>
                <p style="color: #757575; font-size: 0.95rem;">請輸入 4 位數管理員 PIN 碼登入系統</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.write("")

        with st.form("login_form"):
            pin_input = st.text_input("🔑 4 位數 PIN 碼", type="password", max_chars=4, placeholder="請輸入 4 位數字...")
            submit = st.form_submit_button("進入系統", use_container_width=True, type="primary")

            if submit:
                admins = get_admins()
                if pin_input in admins:
                    user_name = admins[pin_input]
                    st.session_state["authenticated_user"] = user_name
                    st.session_state["authenticated_pin"] = pin_input
                    add_audit_log(user_name, "登入系統", "管理員成功通過 PIN 碼登入")
                    st.success(f"歡迎回來，{user_name}！登入成功。")
                    st.rerun()
                else:
                    st.error("PIN 碼錯誤！請重新確認或聯繫負責人。")

        st.markdown(
            """
            <div style="margin-top: 25px; padding: 12px; background: #F1F8E9; border-radius: 8px; font-size: 0.85rem; color: #33691E; text-align: center;">
                💡 <b>預設管理人 PIN 碼提醒</b>：<br>
                <code>0218</code> ➔ KC &nbsp;|&nbsp; <code>0311</code> ➔ Aileen
            </div>
            """,
            unsafe_allow_html=True,
        )


# ==============================================================================
# 7. 頁面功能：營位現況與儀表板 (Dashboard)
# ==============================================================================
def render_dashboard_view():
    st.markdown("### 📊 營區營位現況與快速總覽")

    today = date.today()
    selected_date = st.date_input("📅 選擇欲檢視的營位日期：", value=today, min_value=today - timedelta(days=365))

    # 查詢該日期各區預約狀態
    conn = get_db_connection()
    query = """
        SELECT area, SUM(quantity) as booked_units, COUNT(*) as booking_count
        FROM bookings
        WHERE ? >= check_in_date AND ? < check_out_date
        GROUP BY area
    """
    df_occupied = pd.read_sql_query(query, conn, params=(selected_date.strftime("%Y-%m-%d"), selected_date.strftime("%Y-%m-%d")))
    conn.close()

    occupied_map = {row["area"]: row["booked_units"] for _, row in df_occupied.iterrows()}

    # 顯示三大區域卡片
    cols = st.columns(3)
    for i, (area_name, info) in enumerate(CAMP_AREAS.items()):
        capacity = info["capacity"]
        unit = info["unit"]
        booked = int(occupied_map.get(area_name, 0))
        remaining = max(0, capacity - booked)
        occupancy_rate = min(100, int((booked / capacity) * 100)) if capacity > 0 else 0

        with cols[i]:
            card_html = f"""
            <div class="luye-card" style="border-left: 5px solid {info['color']};">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <h3 style="margin:0; color:{info['color']}; font-size: 1.3rem;">{area_name}</h3>
                    <span style="font-size:0.85rem; background:#EEE; padding:2px 8px; border-radius:12px;">上限 {capacity} {unit}</span>
                </div>
                <p style="margin: 6px 0; color: #666; font-size: 0.85rem;">規格：{info['spec']}</p>
                <div style="margin-top: 12px; display: flex; justify-content: space-between; align-items: baseline;">
                    <div>
                        <span style="font-size: 1.8rem; font-weight: bold; color: {'#C62828' if remaining==0 else '#2E7D32'};">{booked}</span>
                        <span style="color: #666; font-size: 0.9rem;"> / {capacity} {unit} 已訂</span>
                    </div>
                    <div>
                        <span style="font-size: 1.1rem; font-weight: bold; color: {'#C62828' if remaining==0 else '#1565C0'};">剩餘 {remaining}</span>
                    </div>
                </div>
                <div style="background:#ECEFF1; border-radius:6px; height:8px; width:100%; margin-top:8px; overflow:hidden;">
                    <div style="background:{info['color']}; width:{occupancy_rate}%; height:100%;"></div>
                </div>
                <div style="font-size:0.75rem; color:#888; text-align:right; margin-top:4px;">佔用率 {occupancy_rate}%</div>
                <div style="margin-top: 8px; font-size: 0.85rem; color: #424242; border-top: 1px dashed #DDD; padding-top: 8px;">
                    💰 原價：NT$ {info['price_per_night']:,}/{unit}/晚<br>
                    ⚡ 夜衝：NT$ {info['night_rush_price']:,}/{unit} (前晚 18:00~23:00)
                </div>
            </div>
            """
            st.markdown(card_html, unsafe_allow_html=True)

    st.write("---")

    # 今日/選定日期當天入住名單與退房名單
    col_in, col_out = st.columns(2)
    conn = get_db_connection()
    with col_in:
        st.markdown(f"#### 🚪 {selected_date.strftime('%Y-%m-%d')} 當日【辦理入住】名單")
        df_in = pd.read_sql_query(
            "SELECT booking_code as '訂單編號', name as '姓名', phone as '電話', area as '區域', quantity as '數量', night_rush as '夜衝', tag as '標籤' FROM bookings WHERE check_in_date = ?",
            conn,
            params=(selected_date.strftime("%Y-%m-%d"),),
        )
        if not df_in.empty:
            df_in["夜衝"] = df_in["夜衝"].apply(lambda x: "⚡ 是" if x == 1 else "否")
            st.dataframe(df_in, use_container_width=True, hide_index=True)
        else:
            st.info("當日無新入住訂單。")

    with col_out:
        st.markdown(f"#### 🚗 {selected_date.strftime('%Y-%m-%d')} 當日【預計退房】名單")
        df_out = pd.read_sql_query(
            "SELECT booking_code as '訂單編號', name as '姓名', phone as '電話', area as '區域', quantity as '數量', total_amount as '金額' FROM bookings WHERE check_out_date = ?",
            conn,
            params=(selected_date.strftime("%Y-%m-%d"),),
        )
        if not df_out.empty:
            st.dataframe(df_out, use_container_width=True, hide_index=True)
        else:
            st.info("當日無預計退房訂單。")
    conn.close()


# ==============================================================================
# 8. 頁面功能：訂位清單與三大核心索引搜尋 (Search & Booking Management)
# ==============================================================================
def render_bookings_view():
    st.markdown("### 📋 露營客訂單管理與快速檢索")

    # ==========================
    # 三大核心搜尋索引工具列
    # ==========================
    st.markdown(
        """
        <div style="background:#FAFAFA; border:1px solid #E0E0E0; border-radius:10px; padding:16px; margin-bottom:16px;">
            <div style="font-weight:600; color:#2E7D32; margin-bottom:8px;">🔍 三大核心搜尋索引（支援聯集或交集即時篩選）</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    search_col1, search_col2, search_col3, search_col4 = st.columns([1.5, 1.2, 1.2, 1])

    with search_col1:
        date_filter_mode = st.selectbox(
            "1. 依入住/住宿日期篩選",
            options=["所有日期", "今天", "未來 7 天內", "自訂特定日期", "自訂日期區間"],
            index=0,
        )

    with search_col2:
        search_name = st.text_input("2. 依露營人姓名搜尋", placeholder="輸入姓名或關鍵字...").strip()

    with search_col3:
        search_phone = st.text_input("3. 依露營人電話搜尋", placeholder="輸入手機號碼關鍵字...").strip()

    with search_col4:
        area_filter = st.selectbox("預約區域篩選", ["全部區域"] + list(CAMP_AREAS.keys()))

    # 自訂日期額外欄位
    custom_single_date = None
    custom_date_range = None
    if date_filter_mode == "自訂特定日期":
        custom_single_date = st.date_input("選擇日期：", value=date.today())
    elif date_filter_mode == "自訂日期區間":
        custom_date_range = st.date_input("選擇入住日期區間：", value=(date.today(), date.today() + timedelta(days=7)))

    # 建立動態 SQL 查詢
    query = "SELECT * FROM bookings WHERE 1=1"
    params: List[Any] = []

    # 1. 日期條件
    today_str = date.today().strftime("%Y-%m-%d")
    if date_filter_mode == "今天":
        query += " AND ? >= check_in_date AND ? < check_out_date"
        params.extend([today_str, today_str])
    elif date_filter_mode == "未來 7 天內":
        end_date_str = (date.today() + timedelta(days=7)).strftime("%Y-%m-%d")
        query += " AND check_in_date >= ? AND check_in_date <= ?"
        params.extend([today_str, end_date_str])
    elif date_filter_mode == "自訂特定日期" and custom_single_date:
        d_str = custom_single_date.strftime("%Y-%m-%d")
        query += " AND ? >= check_in_date AND ? < check_out_date"
        params.extend([d_str, d_str])
    elif date_filter_mode == "自訂日期區間" and isinstance(custom_date_range, tuple) and len(custom_date_range) == 2:
        s_str = custom_date_range[0].strftime("%Y-%m-%d")
        e_str = custom_date_range[1].strftime("%Y-%m-%d")
        query += " AND check_in_date <= ? AND check_out_date >= ?"
        params.extend([e_str, s_str])

    # 2. 姓名條件 (模糊搜尋)
    if search_name:
        query += " AND name LIKE ?"
        params.append(f"%{search_name}%")

    # 3. 電話條件 (模糊搜尋)
    if search_phone:
        query += " AND phone LIKE ?"
        params.append(f"%{search_phone}%")

    # 4. 區域條件
    if area_filter != "全部區域":
        query += " AND area = ?"
        params.append(area_filter)

    query += " ORDER BY check_in_date DESC, id DESC"

    conn = get_db_connection()
    df_results = pd.read_sql_query(query, conn, params=params)
    conn.close()

    st.markdown(f"**找到 `{len(df_results)}` 筆符合條件的訂單**")

    # 表格或卡片切換
    view_tab1, view_tab2 = st.tabs(["📑 詳細清單表格", "🗂️ 訂單卡片檢視"])

    with view_tab1:
        if not df_results.empty:
            display_df = df_results.copy()
            display_df["night_rush"] = display_df["night_rush"].apply(lambda x: "⚡ 是" if x == 1 else "否")
            display_df["total_amount"] = display_df["total_amount"].apply(lambda x: f"NT$ {x:,}")
            display_df = display_df.rename(
                columns={
                    "booking_code": "訂單編號",
                    "check_in_date": "入住日期",
                    "check_out_date": "退房日期",
                    "name": "露營人姓名",
                    "phone": "電話",
                    "area": "預約區域",
                    "quantity": "帳/間數",
                    "night_rush": "夜衝",
                    "visit_count": "來過次數",
                    "tag": "標籤",
                    "total_amount": "總計金額",
                    "notes": "備註",
                    "created_by": "經辦人",
                    "created_at": "建立時間",
                }
            )
            columns_to_show = [
                "訂單編號",
                "入住日期",
                "退房日期",
                "露營人姓名",
                "電話",
                "預約區域",
                "帳/間數",
                "夜衝",
                "來過次數",
                "標籤",
                "總計金額",
                "備註",
                "經辦人",
            ]
            st.dataframe(display_df[columns_to_show], use_container_width=True, hide_index=True)
        else:
            st.info("查無符合條件的訂單。")

    with view_tab2:
        if not df_results.empty:
            for _, row in df_results.iterrows():
                tag_style = TAG_BADGE_STYLE.get(row["tag"], TAG_BADGE_STYLE["一般"])
                area_unit = CAMP_AREAS.get(row["area"], {}).get("unit", "帳")

                card_box = st.container()
                with card_box:
                    st.markdown(
                        f"""
                        <div class="luye-card">
                            <div style="display:flex; justify-content:space-between; align-items:center;">
                                <div>
                                    <span style="font-size:1.2rem; font-weight:700; color:#1B5E20;">{row['name']}</span>
                                    <span style="font-size:0.95rem; color:#666; margin-left:12px;">📞 {row['phone']}</span>
                                    <span style="{tag_style}; margin-left:10px;">{row['tag']}</span>
                                    <span style="background:#E0F2F1; color:#00695C; padding:2px 8px; border-radius:12px; font-size:0.8rem; margin-left:6px;">第 {row['visit_count']} 次到訪</span>
                                </div>
                                <div>
                                    <span style="font-weight:700; font-size:1.25rem; color:#E65100;">NT$ {row['total_amount']:,}</span>
                                </div>
                            </div>
                            <div style="margin-top:10px; font-size:0.92rem; color:#333; display:flex; flex-wrap:wrap; gap:16px;">
                                <div>🏕️ <b>區域</b>：{row['area']} ({row['quantity']} {area_unit})</div>
                                <div>📅 <b>入住~退房</b>：{row['check_in_date']} 至 {row['check_out_date']}</div>
                                <div>⚡ <b>夜衝</b>：{'是 (18:00~23:00)' if row['night_rush'] == 1 else '否'}</div>
                                <div>🔖 <b>訂單號</b>：<code>{row['booking_code']}</code></div>
                                <div>👤 <b>經手人</b>：{row['created_by']}</div>
                            </div>
                            {f'<div style="margin-top:8px; padding:6px 10px; background:#FFF9C4; border-radius:6px; font-size:0.88rem; color:#5D4037;">📝 備註：{row["notes"]}</div>' if row["notes"] else ''}
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    btn_c1, btn_c2, btn_c3 = st.columns([1, 1, 6])
                    with btn_c1:
                        if st.button("✏️ 編輯", key=f"edit_btn_{row['id']}"):
                            st.session_state["editing_booking_id"] = row["id"]
                            st.session_state["nav_page"] = "➕ 新增/編輯預約"
                            st.rerun()
                    with btn_c2:
                        with st.popover("🗑️ 刪除"):
                            st.write(f"確定要刪除訂單 **{row['booking_code']}** ({row['name']}) 嗎？")
                            if st.button("確認刪除", key=f"del_confirm_{row['id']}", type="primary"):
                                current_user = st.session_state.get("authenticated_user", "管理員")
                                del_conn = get_db_connection()
                                del_cursor = del_conn.cursor()
                                del_cursor.execute("DELETE FROM bookings WHERE id = ?", (row["id"],))
                                del_conn.commit()
                                del_conn.close()

                                add_audit_log(
                                    current_user,
                                    "刪除訂單",
                                    f"刪除訂單 {row['booking_code']} (露營人: {row['name']}, 電話: {row['phone']}, 區域: {row['area']}, 金額: {row['total_amount']})",
                                )
                                trigger_auto_sync(current_user)
                                st.success(f"訂單 {row['booking_code']} 已順利刪除！")
                                st.rerun()
        else:
            st.info("尚無符合搜尋條件的訂單記錄。")


# ==============================================================================
# 9. 頁面功能：新增與編輯預約 (Booking Form)
# ==============================================================================
def render_booking_form_view():
    editing_id = st.session_state.get("editing_booking_id", None)
    is_editing = editing_id is not None

    booking_data = None
    if is_editing:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM bookings WHERE id = ?", (editing_id,))
        row = cursor.fetchone()
        conn.close()
        if row:
            booking_data = dict(row)
        else:
            st.warning("找不到該筆編輯訂單，已切換回新增模式。")
            st.session_state["editing_booking_id"] = None
            is_editing = False

    st.markdown(f"### {'✏️ 編輯露營訂位資料' if is_editing else '➕ 新增露營客預約訂單'}")

    if is_editing:
        st.info(f"正在編輯訂單：**{booking_data['booking_code']}** (露營客: {booking_data['name']})")
        if st.button("⬅️ 取消編輯，切換回新增預約模式"):
            st.session_state["editing_booking_id"] = None
            st.rerun()

    # 預設值設定
    default_name = booking_data["name"] if is_editing else ""
    default_phone = booking_data["phone"] if is_editing else ""
    default_area = booking_data["area"] if is_editing else "雨棚區"
    default_quantity = booking_data["quantity"] if is_editing else 1
    default_night_rush = bool(booking_data["night_rush"]) if is_editing else False
    default_tag = booking_data["tag"] if is_editing else "一般"
    default_notes = booking_data["notes"] if is_editing else ""
    default_visit_count = booking_data["visit_count"] if is_editing else 1

    if is_editing:
        default_check_in = datetime.strptime(booking_data["check_in_date"], "%Y-%m-%d").date()
        default_check_out = datetime.strptime(booking_data["check_out_date"], "%Y-%m-%d").date()
    else:
        default_check_in = date.today() + timedelta(days=1)
        default_check_out = default_check_in + timedelta(days=1)

    with st.form("booking_editor_form"):
        st.markdown("#### 1. 露營客基本資料")
        col_c1, col_c2, col_c3 = st.columns([1.2, 1.2, 1])

        with col_c1:
            name_val = st.text_input("露營人姓名 *", value=default_name, placeholder="例如：王小明")
        with col_c2:
            phone_val = st.text_input("露營人電話 *", value=default_phone, placeholder="例如：0912345678")
        with col_c3:
            tag_val = st.selectbox(
                "客戶標籤 *",
                options=CUSTOMER_TAGS,
                index=CUSTOMER_TAGS.index(default_tag) if default_tag in CUSTOMER_TAGS else 0,
            )

        # 自動統計來訪次數
        hist_count, _ = query_historical_stats(phone_val)
        calc_visit_count = default_visit_count if is_editing else max(1, hist_count + 1 if phone_val else 1)

        col_v1, col_v2 = st.columns([1, 2])
        with col_v1:
            visit_count_val = st.number_input(
                "來過次數 (系統自動統計或手動微調)",
                min_value=1,
                max_value=999,
                value=int(calc_visit_count),
                help="輸入電話時若已存在歷史訂單，系統會自動參考歷史到訪紀錄。",
            )

        st.markdown("---")
        st.markdown("#### 2. 營位預訂細節與日期")

        col_b1, col_b2, col_b3 = st.columns([1.2, 1, 1])
        with col_b1:
            area_val = st.selectbox(
                "預約區域 *",
                options=list(CAMP_AREAS.keys()),
                index=list(CAMP_AREAS.keys()).index(default_area),
            )
        with col_b2:
            max_cap = CAMP_AREAS[area_val]["capacity"]
            unit_str = CAMP_AREAS[area_val]["unit"]
            quantity_val = st.number_input(
                f"預約數量 ({unit_str}) *",
                min_value=1,
                max_value=max_cap,
                value=min(default_quantity, max_cap),
            )
        with col_b3:
            night_rush_val = st.checkbox("⛺ 前晚夜衝 (18:00–23:00 半價)", value=default_night_rush)

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            check_in_val = st.date_input("主要入住日期 *", value=default_check_in)
        with col_d2:
            check_out_val = st.date_input("預計退房日期 *", value=default_check_out)

        notes_val = st.text_area("空白文字備註 (特殊需求、車牌、同團成員等)", value=default_notes, placeholder="請在此自由填寫備註...")

        # 即時費用計算
        calc_total, nights_count, base_fee, rush_fee = calculate_booking_fee(
            area_val, quantity_val, check_in_val, check_out_val, night_rush_val
        )

        st.markdown("---")
        st.markdown("#### 3. 預估費用計算明細")
        f_col1, f_col2, f_col3, f_col4 = st.columns(4)
        f_col1.metric("預計住宿晚數", f"{nights_count} 晚")
        f_col2.metric("營位基礎費用", f"NT$ {base_fee:,}")
        f_col3.metric("夜衝費用", f"NT$ {rush_fee:,}")
        f_col4.metric("應收總金額", f"NT$ {calc_total:,}", delta="含車位規格")

        submit_btn = st.form_submit_button(
            "💾 儲存並建立訂單" if not is_editing else "💾 儲存修改內容",
            type="primary",
            use_container_width=True,
        )

        if submit_btn:
            # 檢核必要欄位
            if not name_val.strip():
                st.error("請填寫露營人姓名！")
            elif not phone_val.strip():
                st.error("請填寫露營人電話！")
            elif check_out_val <= check_in_val:
                st.error("退房日期必須在入住日期之後！")
            else:
                current_user = st.session_state.get("authenticated_user", "管理員")
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                conn = get_db_connection()
                cursor = conn.cursor()

                try:
                    if is_editing:
                        # 更新現有訂單
                        cursor.execute(
                            """
                            UPDATE bookings
                            SET check_in_date = ?, check_out_date = ?, name = ?, phone = ?,
                                area = ?, quantity = ?, night_rush = ?, visit_count = ?,
                                tag = ?, notes = ?, total_amount = ?, updated_at = ?
                            WHERE id = ?
                            """,
                            (
                                check_in_val.strftime("%Y-%m-%d"),
                                check_out_val.strftime("%Y-%m-%d"),
                                name_val.strip(),
                                phone_val.strip(),
                                area_val,
                                quantity_val,
                                1 if night_rush_val else 0,
                                visit_count_val,
                                tag_val,
                                notes_val.strip(),
                                calc_total,
                                now_str,
                                editing_id,
                            ),
                        )
                        conn.commit()
                        add_audit_log(
                            current_user,
                            "修改訂單",
                            f"修改訂單 {booking_data['booking_code']} (客戶: {name_val}, 區域: {area_val} {quantity_val}{unit_str}, 總額: {calc_total})",
                        )
                        st.session_state["editing_booking_id"] = None
                        st.success(f"訂單 {booking_data['booking_code']} 修改成功！")
                    else:
                        # 新增訂單
                        b_code = generate_booking_code(check_in_val)
                        cursor.execute(
                            """
                            INSERT INTO bookings (
                                booking_code, check_in_date, check_out_date, name, phone,
                                area, quantity, night_rush, visit_count, tag, notes,
                                total_amount, created_by, created_at, updated_at
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                b_code,
                                check_in_val.strftime("%Y-%m-%d"),
                                check_out_val.strftime("%Y-%m-%d"),
                                name_val.strip(),
                                phone_val.strip(),
                                area_val,
                                quantity_val,
                                1 if night_rush_val else 0,
                                visit_count_val,
                                tag_val,
                                notes_val.strip(),
                                calc_total,
                                current_user,
                                now_str,
                                now_str,
                            ),
                        )
                        conn.commit()
                        add_audit_log(
                            current_user,
                            "新增訂單",
                            f"新增訂單 {b_code} (客戶: {name_val}, 電話: {phone_val}, 區域: {area_val} {quantity_val}{unit_str}, 入住: {check_in_val}, 金額: {calc_total})",
                        )
                        st.success(f"🎉 訂單建立成功！訂單編號：{b_code}，總金額：NT$ {calc_total:,}")

                    conn.close()
                    # 觸發雲端與本地自動同步備份
                    trigger_auto_sync(current_user)
                    st.balloons()
                except Exception as e:
                    conn.close()
                    st.error(f"儲存訂單時發生錯誤: {e}")


# ==============================================================================
# 10. 頁面功能：Excel/CSV 匯入與匯出 (Import / Export)
# ==============================================================================
def render_import_export_view():
    st.markdown("### 📤 Excel / CSV 批次資料匯入與匯出")

    tab_export, tab_import = st.tabs(["📥 一鍵資料匯出 (Export)", "📤 批次上傳匯入 (Import)"])

    # 匯出資料
    with tab_export:
        st.markdown("#### 1. 匯出目前營地所有露營客訂單資料")
        st.write("點擊下方按鈕可將系統中完整的露營客訂位紀錄、金額與備註匯出為 Excel 或 CSV 試算表。")

        conn = get_db_connection()
        df_all = pd.read_sql_query("SELECT * FROM bookings ORDER BY check_in_date DESC", conn)
        conn.close()

        if not df_all.empty:
            df_export = df_all.copy()
            df_export["night_rush"] = df_export["night_rush"].apply(lambda x: "是" if x == 1 else "否")
            df_export = df_export.rename(
                columns={
                    "booking_code": "訂單編號",
                    "check_in_date": "入住日期",
                    "check_out_date": "退房日期",
                    "name": "露營人姓名",
                    "phone": "露營人電話",
                    "area": "預約區域",
                    "quantity": "帳數/間數",
                    "night_rush": "是否夜衝",
                    "visit_count": "來過次數",
                    "tag": "客戶標籤",
                    "total_amount": "總計金額",
                    "notes": "備註說明",
                    "created_by": "建立人員",
                    "created_at": "建立時間",
                    "updated_at": "最後更新時間",
                }
            )

            # Excel 下載
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                df_export.to_excel(writer, index=False, sheet_name="祿野營地訂位名單")
            excel_bytes = excel_buffer.getvalue()

            # CSV 下載 (包含 utf-8-sig 以免繁體中文亂碼)
            csv_bytes = df_export.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

            now_tag = datetime.now().strftime("%Y%m%d_%H%M%S")
            col_ex1, col_ex2 = st.columns(2)
            with col_ex1:
                st.download_button(
                    label="📊 匯出為 Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"祿野營地訂位名單_{now_tag}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary",
                )
            with col_ex2:
                st.download_button(
                    label="📄 匯出為 CSV 檔 (.csv)",
                    data=csv_bytes,
                    file_name=f"祿野營地訂位名單_{now_tag}.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

            st.write(f"目前資料庫中共有 **{len(df_all)}** 筆訂單資料。")
        else:
            st.info("資料庫目前尚無訂單資料可供匯出。")

    # 匯入資料
    with tab_import:
        st.markdown("#### 2. 批次匯入露營客訂單名單")
        st.write("上傳 Excel (.xlsx) 或 CSV (.csv) 格式檔案，系統將自動解析並寫入訂位管理資料庫中。")

        # 提供範本下載
        template_df = pd.DataFrame(
            [
                {
                    "主要入住日期": "2026-10-10",
                    "退房日期": "2026-10-12",
                    "露營人姓名": "張小露",
                    "露營人電話": "0988123456",
                    "預約區域": "雨棚區",
                    "帳數/間數": 2,
                    "是否夜衝": "是",
                    "來過次數": 1,
                    "客戶標籤": "好客人",
                    "空白文字備註": "需靠近水槽、帶 1 隻寵物",
                },
                {
                    "主要入住日期": "2026-10-15",
                    "退房日期": "2026-10-16",
                    "露營人姓名": "李大野",
                    "露營人電話": "0911654321",
                    "預約區域": "草地區",
                    "帳數/間數": 1,
                    "是否夜衝": "否",
                    "來過次數": 3,
                    "客戶標籤": "一般",
                    "空白文字備註": "第一次夜衝需說明",
                },
            ]
        )
        template_buffer = io.BytesIO()
        with pd.ExcelWriter(template_buffer, engine="openpyxl") as writer:
            template_df.to_excel(writer, index=False, sheet_name="匯入範本")
        st.download_button(
            label="📥 下載標準匯入範本 (.xlsx)",
            data=template_buffer.getvalue(),
            file_name="祿野營地訂單匯入標準範本.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        uploaded_file = st.file_uploader("請選擇欲上傳之 Excel 或 CSV 檔案", type=["xlsx", "csv"])
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith(".csv"):
                    df_upload = pd.read_csv(uploaded_file)
                else:
                    df_upload = pd.read_excel(uploaded_file)

                # 欄位靈活配對
                in_col = next((c for c in ["主要入住日期", "入住日期", "check_in_date", "check_in"] if c in df_upload.columns), None)
                out_col = next((c for c in ["退房日期", "check_out_date", "check_out"] if c in df_upload.columns), None)
                phone_col = next((c for c in ["露營人電話", "電話", "phone", "mobile"] if c in df_upload.columns), None)

                # 1. 增強日期欄位容錯解析機制 (pd.to_datetime(..., errors='coerce'))
                today_str = date.today().strftime("%Y-%m-%d")
                if in_col:
                    df_upload[in_col] = pd.to_datetime(df_upload[in_col], format="mixed", errors="coerce").dt.strftime("%Y-%m-%d").fillna(today_str)
                    # 避免空字串情況，統一填補當天日期
                    df_upload[in_col] = df_upload[in_col].apply(lambda x: today_str if not str(x).strip() or str(x).strip().lower() == "nan" else str(x).strip())

                if out_col:
                    df_upload[out_col] = pd.to_datetime(df_upload[out_col], format="mixed", errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
                    df_upload[out_col] = df_upload[out_col].apply(lambda x: "" if not str(x).strip() or str(x).strip().lower() == "nan" else str(x).strip())

                # 2. 增強電話號碼欄位容錯解析機制 (去除 .0 並補齊開頭 0)
                if phone_col:
                    df_upload[phone_col] = df_upload[phone_col].apply(format_phone_number)

                st.write("預覽上傳內容（前 5 筆）：")
                st.dataframe(df_upload.head(5), use_container_width=True)

                if st.button("🚀 開始執行批次匯入", type="primary"):
                    success_count = 0
                    skip_count = 0
                    error_logs = []
                    current_user = st.session_state.get("authenticated_user", "管理員")
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    conn = get_db_connection()
                    cursor = conn.cursor()

                    for idx, row in df_upload.iterrows():
                        try:
                            # 姓名與電話解析
                            name = str(row.get("露營人姓名", row.get("name", ""))).strip()
                            if name.lower() == "nan":
                                name = ""

                            raw_phone = row.get(phone_col) if phone_col else row.get("露營人電話", row.get("phone", ""))
                            phone = format_phone_number(raw_phone)

                            # 預約區域解析
                            area = str(row.get("預約區域", row.get("area", "雨棚區"))).strip()
                            if area not in CAMP_AREAS:
                                area = "雨棚區"

                            # 日期安全解析 (確保絕不丟出 ValueError)
                            in_raw = str(row.get(in_col, "") if in_col else row.get("主要入住日期", row.get("check_in_date", ""))).strip()
                            try:
                                if in_raw and in_raw.lower() != "nan":
                                    check_in = datetime.strptime(in_raw[:10], "%Y-%m-%d").date()
                                else:
                                    check_in = date.today()
                            except Exception:
                                check_in = date.today()

                            out_raw = str(row.get(out_col, "") if out_col else row.get("退房日期", row.get("check_out_date", ""))).strip()
                            try:
                                if out_raw and out_raw.lower() != "nan":
                                    check_out = datetime.strptime(out_raw[:10], "%Y-%m-%d").date()
                                else:
                                    check_out = check_in + timedelta(days=1)
                            except Exception:
                                check_out = check_in + timedelta(days=1)

                            # 若退房日期早於或等於入住日期，自動修正為入住隔天
                            if check_out <= check_in:
                                check_out = check_in + timedelta(days=1)

                            # 數量解析
                            raw_qty = row.get("帳數/間數", row.get("quantity", 1))
                            try:
                                qty = int(float(raw_qty)) if not pd.isna(raw_qty) else 1
                                qty = max(1, qty)
                            except Exception:
                                qty = 1

                            rush_val = str(row.get("是否夜衝", row.get("night_rush", "否"))).strip().lower()
                            night_rush = rush_val in ["是", "true", "1", "yes", "y"]

                            # 來訪次數
                            raw_v = row.get("來過次數", row.get("visit_count", 1))
                            try:
                                v_count = int(float(raw_v)) if not pd.isna(raw_v) else 1
                                v_count = max(1, v_count)
                            except Exception:
                                v_count = 1

                            tag = str(row.get("客戶標籤", row.get("tag", "一般"))).strip()
                            if tag not in CUSTOMER_TAGS:
                                tag = "一般"

                            notes = str(row.get("空白文字備註", row.get("notes", ""))).strip()
                            if notes.lower() in ["nan", "none", "null"]:
                                notes = ""

                            if not name or not phone:
                                skip_count += 1
                                error_logs.append(f"第 {idx+1} 列：缺少姓名或電話，略過。")
                                continue

                            # 自動計算費用與訂單編號
                            calc_total, _, _, _ = calculate_booking_fee(area, qty, check_in, check_out, night_rush)
                            b_code = generate_booking_code(check_in) + f"-{idx}"

                            cursor.execute(
                                """
                                INSERT INTO bookings (
                                    booking_code, check_in_date, check_out_date, name, phone,
                                    area, quantity, night_rush, visit_count, tag, notes,
                                    total_amount, created_by, created_at, updated_at
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                """,
                                (
                                    b_code,
                                    check_in.strftime("%Y-%m-%d"),
                                    check_out.strftime("%Y-%m-%d"),
                                    name,
                                    phone,
                                    area,
                                    qty,
                                    1 if night_rush else 0,
                                    v_count,
                                    tag,
                                    notes,
                                    calc_total,
                                    current_user,
                                    now_str,
                                    now_str,
                                ),
                            )
                            success_count += 1
                        except Exception as e_row:
                            skip_count += 1
                            error_logs.append(f"第 {idx+1} 列匯入失敗: {e_row}")

                    conn.commit()
                    conn.close()

                    add_audit_log(
                        current_user,
                        "批次匯入訂單",
                        f"自檔案 {uploaded_file.name} 匯入，成功 {success_count} 筆，略過/失敗 {skip_count} 筆",
                    )
                    trigger_auto_sync(current_user)

                    st.success(f"匯入作業完成！成功匯入 {success_count} 筆訂單。")
                    if error_logs:
                        with st.expander("檢視異常/略過清單"):
                            for err in error_logs:
                                st.write(f"- {err}")
            except Exception as e:
                st.error(f"檔案讀取或解析失敗: {e}")


# ==============================================================================
# 11. 頁面功能：操作日誌 Note Log (Audit Log)
# ==============================================================================
def render_audit_log_view():
    st.markdown("### 📜 系統操作日誌 (Audit Log)")
    st.write("記錄每一次「新增訂單」、「修改訂單」、「刪除訂單」、「管理人調整」等關鍵異動操作，確保營區營運歷程清晰可稽。")

    df_logs = get_audit_logs(limit=500)

    if not df_logs.empty:
        # 日誌篩選
        col_l1, col_l2 = st.columns([1, 2])
        with col_l1:
            operators = ["全部操作者"] + list(df_logs["操作者"].unique())
            selected_op = st.selectbox("依操作人員篩選", operators)
        with col_l2:
            search_keyword = st.text_input("搜尋動作或細節關鍵字", placeholder="例如：雨棚區、刪除、KC...")

        filtered_logs = df_logs.copy()
        if selected_op != "全部操作者":
            filtered_logs = filtered_logs[filtered_logs["操作者"] == selected_op]
        if search_keyword:
            filtered_logs = filtered_logs[
                filtered_logs["操作動作"].str.contains(search_keyword, case=False, na=False)
                | filtered_logs["具體細節"].str.contains(search_keyword, case=False, na=False)
            ]

        st.dataframe(filtered_logs, use_container_width=True, hide_index=True)

        # 匯出 Log
        csv_logs = filtered_logs.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            "📥 下載操作日誌紀錄 (CSV)",
            data=csv_logs,
            file_name=f"祿野營地操作日誌_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
        )
    else:
        st.info("目前尚無任何操作日誌。")


# ==============================================================================
# 12. 頁面功能：管理人帳號管理 (Admin Management)
# ==============================================================================
def render_admin_management_view():
    st.markdown("### 🔑 管理人帳號與 PIN 碼管理")
    st.write("維護可登入本系統之管理員 PIN 碼名冊。系統至少需保留一組有效的 4 位數 PIN 碼以防無法登入。")

    admins = get_admins()
    current_user = st.session_state.get("authenticated_user", "管理員")

    col_list, col_add = st.columns([1.2, 1])

    with col_list:
        st.markdown("#### 目前授權管理人清單")
        for pin, name in list(admins.items()):
            is_current = name == current_user
            with st.container():
                st.markdown(
                    f"""
                    <div style="background:white; border:1px solid #E0E0E0; border-radius:8px; padding:12px 16px; margin-bottom:10px; display:flex; justify-content:space-between; align-items:center;">
                        <div>
                            <span style="font-weight:700; font-size:1.1rem; color:#2E7D32;">{name}</span>
                            {' <span style="font-size:0.75rem; background:#C8E6C9; color:#1B5E20; padding:2px 6px; border-radius:8px;">目前在線</span>' if is_current else ''}
                            <div style="color:#757575; font-size:0.85rem; margin-top:2px;">PIN 碼：<code>****{pin[-2:]}</code></div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                # 刪除防呆
                with st.popover(f"刪除 {name} 的權限"):
                    st.write(f"確定要移除管理人 **{name}** 的 PIN 碼登入權限嗎？")
                    if len(admins) <= 1:
                        st.error("防呆保護：系統至少必須保留一組管理員 PIN 碼，無法刪除最後一人！")
                    else:
                        if st.button("確認刪除", key=f"del_pin_{pin}", type="primary"):
                            del admins[pin]
                            if save_admins(admins):
                                add_audit_log(current_user, "刪除管理人", f"移除了管理人 {name} (PIN: {pin}) 的登入權限")
                                trigger_auto_sync(current_user)
                                st.success(f"已成功刪除管理人 {name}！")
                                st.rerun()

    with col_add:
        st.markdown("#### ➕ 新增授權管理人 PIN 碼")
        with st.form("add_admin_form"):
            new_name = st.text_input("管理人稱呼/姓名", placeholder="例如：Alex")
            new_pin = st.text_input("4 位數字 PIN 碼", max_chars=4, placeholder="4 位數字，如 1234")
            add_btn = st.form_submit_button("確認新增管理人", type="primary", use_container_width=True)

            if add_btn:
                if not new_name.strip():
                    st.error("請填寫管理人名稱！")
                elif not (new_pin.isdigit() and len(new_pin) == 4):
                    st.error("PIN 碼必須為正好 4 位純數字！")
                elif new_pin in admins:
                    st.error(f"該 PIN 碼已被管理人【{admins[new_pin]}】使用，請使用其他 PIN 碼！")
                else:
                    admins[new_pin] = new_name.strip()
                    if save_admins(admins):
                        add_audit_log(current_user, "新增管理人", f"新增管理人 {new_name.strip()} (PIN: {new_pin})")
                        trigger_auto_sync(current_user)
                        st.success(f"🎉 成功新增管理人【{new_name.strip()}】！")
                        st.rerun()


# ==============================================================================
# 13. 頁面功能：Google Drive 雲端同步與備份狀態 (Cloud Sync)
# ==============================================================================
def render_cloud_sync_view():
    st.markdown("### ☁️ Google Drive 雲端同步與備份中心")
    st.write("設定與檢視 Google Drive 備份狀態。每次資料有新增、修改或刪除時，系統均會觸發自動備份機制。")

    sync_tool = get_gdrive_sync()
    current_user = st.session_state.get("authenticated_user", "管理員")

    st.markdown(
        f"""
        <div class="luye-card">
            <h4>📁 Google Drive 對接資訊</h4>
            <ul>
                <li><b>目標資料夾 ID</b>：<code>{DEFAULT_GDRIVE_FOLDER_ID}</code> (資料夾名稱：DVCG)</li>
                <li><b>雲端連線狀態</b>：{'<span style="color:#2E7D32; font-weight:700;">🟢 已連線正常</span>' if sync_tool.is_connected else '<span style="color:#E65100; font-weight:700;">🟡 本地儲存模式 (離線平滑降級)</span>'}</li>
                <li><b>詳細連線說明</b>：{sync_tool.status_message}</li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### ⚡ 立即執行備份")
        st.write("手動打包所有訂單與 Audit Log，產出最新 Excel 副本並同步至雲端與本地備份目錄。")
        if st.button("🚀 立即手動同步與備份", type="primary", use_container_width=True):
            with st.spinner("正在進行資料備份中..."):
                success, msg = sync_tool.backup_database_now(operator=current_user)
                if success:
                    st.success(msg)
                else:
                    st.warning(msg)

    with c2:
        st.markdown("#### 🔑 Google Service Account 金鑰設定")
        st.write("若欲啟用 Google Drive 雲端自動上傳，請上傳具備存取權限的 `service_account.json` 憑證檔：")
        sa_file = st.file_uploader("上傳 service_account.json", type=["json"])
        if sa_file is not None:
            if st.button("安裝憑證並重新連線"):
                try:
                    ensure_directories()
                    with open(SERVICE_ACCOUNT_PATH, "wb") as f:
                        f.write(sa_file.getvalue())
                    st.cache_resource.clear()
                    st.success("憑證檔案儲存成功！系統正重新載入雲端服務...")
                    add_audit_log(current_user, "更新雲端憑證", "上傳了新的 Google Service Account 金鑰")
                    st.rerun()
                except Exception as e:
                    st.error(f"儲存憑證失敗: {e}")

    # 本地備份檔案列表
    st.markdown("---")
    st.markdown("#### 💾 本地備份檔案列表 (備份歷程)")
    ensure_directories()
    backup_files = [f for f in os.listdir(DATA_DIR) if f.startswith("祿野營地資料庫完整備份_") and f.endswith(".xlsx")]
    backup_files.sort(reverse=True)

    if backup_files:
        st.write(f"目前本地已保存 **{len(backup_files)}** 份歷史備份檔：")
        for b_file in backup_files[:5]:
            f_path = os.path.join(DATA_DIR, b_file)
            f_size = os.path.getsize(f_path) / 1024
            st.markdown(f"- 📄 `{b_file}` ({f_size:.1f} KB)")
    else:
        st.info("尚無本地歷史備份檔案，請點擊上方按鈕立即產出第一份備份。")


# ==============================================================================
# 14. 主程式入口與導覽路由 (Main App Entry)
# ==============================================================================
def main():
    apply_custom_styles()
    init_db()
    init_admins()

    # 檢查登入狀態
    if not check_authentication():
        render_login_view()
        return

    # 頂部橫幅 Header
    st.markdown(
        f"""
        <div class="app-header">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <div>
                    <h1>⛺ {APP_TITLE}</h1>
                    <p>{APP_SUBTITLE}</p>
                </div>
                <div style="text-align:right;">
                    <span style="background:rgba(255,255,255,0.2); padding:6px 14px; border-radius:20px; font-size:0.95rem;">
                        👤 管理人員：<b>{st.session_state.get('authenticated_user', '管理員')}</b>
                    </span>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # 側邊欄導覽
    with st.sidebar:
        st.markdown("### 🌲 祿野營地導覽選單")
        st.write(f"目前登入者：**{st.session_state.get('authenticated_user', '管理員')}**")

        default_nav = st.session_state.get("nav_page", "⛺ 營位總覽看板")
        nav_options = [
            "⛺ 營位總覽看板",
            "📋 訂單管理與搜尋",
            "➕ 新增/編輯預約",
            "📤 資料匯入與匯出",
            "📜 操作日誌 (Log)",
            "🔑 管理人帳號管理",
            "☁️ 雲端備份與同步",
        ]

        if default_nav not in nav_options:
            default_nav = nav_options[0]

        page_selection = st.radio(
            "前往分頁：",
            nav_options,
            index=nav_options.index(default_nav),
        )
        st.session_state["nav_page"] = page_selection

        st.write("---")

        # 營地基本費率速查卡片
        st.markdown(
            """
            <div style="font-size:0.85rem; color:#424242; background:#F1F8E9; padding:10px; border-radius:8px;">
                <b>⛺ 營區規格速查</b><br>
                • <b>雨棚區 (8帳)</b>：1200元/晚<br>
                • <b>草地區 (4帳)</b>：1200元/晚<br>
                • <b>通舖房間 (2間)</b>：2000元/晚<br>
                • <b>前晚夜衝 (18~23時)</b>：半價 50%
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.write("")
        if st.button("🚪 登出系統", use_container_width=True):
            user_name = st.session_state.get("authenticated_user", "管理員")
            add_audit_log(user_name, "登出系統", "管理員登出成功")
            st.session_state["authenticated_user"] = None
            st.session_state["authenticated_pin"] = None
            st.session_state["editing_booking_id"] = None
            st.rerun()

    # 路由渲染對應頁面
    if page_selection == "⛺ 營位總覽看板":
        render_dashboard_view()
    elif page_selection == "📋 訂單管理與搜尋":
        render_bookings_view()
    elif page_selection == "➕ 新增/編輯預約":
        render_booking_form_view()
    elif page_selection == "📤 資料匯入與匯出":
        render_import_export_view()
    elif page_selection == "📜 操作日誌 (Log)":
        render_audit_log_view()
    elif page_selection == "🔑 管理人帳號管理":
        render_admin_management_view()
    elif page_selection == "☁️ 雲端備份與同步":
        render_cloud_sync_view()


if __name__ == "__main__":
    main()
