"""
祿野營地訂位管理系統 (Lu-Ye Campsite Booking Management System)
===========================================================
- 營區區域：雨棚區 (8帳)、草地區 (4帳)、通舖房間 (2間)
- 安全機制：4 位數 PIN 碼登入與管理員權限控制
- 核心功能：訂位管理、三大核心索引搜尋、Excel/CSV 匯入匯出、操作稽核日誌 (Audit Log)
- 雲端備份：Google Sheets 完整資料庫轉譯備份 (DB_Full_Backup 分頁)，具備平滑降級本地機制
"""

import io
import json
import os
import re
import sqlite3
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# ==============================================================================
# 1. 系統常數與營區規格設定
# ==============================================================================
APP_TITLE = "祿野營地訂位管理系統"
APP_SUBTITLE = "專業・溫馨・便利的露營營位管理後台"

# 應用程式根目錄：以 pathlib 絕對鎖定 (resolve 會一併解析符號連結)，
# 確保 Streamlit 睡眠喚醒/重開機後即使工作目錄改變，SQLite 資料庫仍
# 絕對指向 app.py 所在目錄，不會誤用相對路徑或建立空資料庫檔案
APP_ROOT = Path(__file__).resolve().parent
DATA_DIR = os.path.join(str(APP_ROOT), "data")
DB_PATH = os.path.join(DATA_DIR, "luye_camping.db")
ADMIN_CONFIG_PATH = os.path.join(DATA_DIR, "admins.json")
SERVICE_ACCOUNT_PATH = os.path.join(DATA_DIR, "service_account.json")

# Google Drive 目標資料夾 (DVCG 資料夾 ID)
DVCG_FOLDER_ID = "1JFzbSQuolGniXeAeeaX3okbUZPQ-Rzok"
DEFAULT_GDRIVE_FOLDER_ID = DVCG_FOLDER_ID


def get_gdrive_folder_id() -> str:
    """取得 Google Drive 目標資料夾 ID (依序檢查 st.secrets 與環境變數，預設回退至 DVCG_FOLDER_ID)

    (註：st.secrets 巢狀區段回傳 AttrDict，屬 Mapping 而非 dict 子類別，
     故直接以 in / [] 存取，不可使用 isinstance(..., dict) 判斷)
    """
    try:
        if hasattr(st, "secrets"):
            # 優先嘗試 1: st.secrets["google_drive"]["folder_id"]
            if "google_drive" in st.secrets:
                drive_section = st.secrets["google_drive"]
                if "folder_id" in drive_section:
                    return str(drive_section["folder_id"]).strip()
            # 嘗試 2: st.secrets["gcp_service_account"]["folder_id"]
            if "gcp_service_account" in st.secrets:
                sa_section = st.secrets["gcp_service_account"]
                if "folder_id" in sa_section:
                    return str(sa_section["folder_id"]).strip()
            # 嘗試 3: 其他常見單層 Secrets
            if "gdrive_folder_id" in st.secrets:
                return str(st.secrets["gdrive_folder_id"]).strip()
            if "GDRIVE_FOLDER_ID" in st.secrets:
                return str(st.secrets["GDRIVE_FOLDER_ID"]).strip()
            if "FOLDER_ID" in st.secrets:
                return str(st.secrets["FOLDER_ID"]).strip()
    except Exception:
        pass
    return os.environ.get("GDRIVE_FOLDER_ID", os.environ.get("FOLDER_ID", DVCG_FOLDER_ID)).strip()


def get_service_account_email() -> str:
    """取得 Service Account 的 client_email

    用途：Service Account 預設沒有 Drive 儲存配額，無法新建檔案；
    需將目標備份檔案预先分享給此信箱 (編輯者權限)，系統才能以覆寫方式同步。
    """
    try:
        if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
            sa_info = dict(st.secrets["gcp_service_account"])
            if isinstance(sa_info, dict) and "client_email" in sa_info:
                return str(sa_info["client_email"]).strip()
        if os.path.exists(SERVICE_ACCOUNT_PATH):
            with open(SERVICE_ACCOUNT_PATH, "r", encoding="utf-8") as f:
                sa_info = json.load(f)
            if isinstance(sa_info, dict) and "client_email" in sa_info:
                return str(sa_info["client_email"]).strip()
    except Exception as e:
        print(f"[Service Account] 讀取 client_email 失敗: {e}")
    return ""


# 全域 FOLDER_ID (供備份模組調用)
FOLDER_ID = get_gdrive_folder_id()

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


@st.cache_data(ttl=600)
def get_admins() -> Dict[str, str]:
    """取得所有管理員資訊 (PIN -> 名稱) - 快取 600 秒"""
    init_admins()
    try:
        with open(ADMIN_CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        return {"0218": "KC", "0311": "Aileen"}


def save_admins(admins: Dict[str, str]) -> bool:
    """儲存管理員資訊並清空快取"""
    ensure_directories()
    try:
        with open(ADMIN_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(admins, f, ensure_ascii=False, indent=2)
        st.cache_data.clear()
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
    """初始化資料庫表單與自動升級欄位"""
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
            night_rush_quantity INTEGER NOT NULL DEFAULT 0,
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

    # 平滑遷移新增 night_rush_quantity 欄位
    try:
        cursor.execute("ALTER TABLE bookings ADD COLUMN night_rush_quantity INTEGER DEFAULT 0")
    except Exception:
        pass
        
    try:
        cursor.execute("ALTER TABLE bookings ADD COLUMN night_rush_detail TEXT DEFAULT ''")
    except Exception:
        pass

    # 舊資料庫相容：若 night_rush=1 且 night_rush_quantity 為 0，自動對齊為 quantity
    try:
        cursor.execute("UPDATE bookings SET night_rush_quantity = quantity WHERE night_rush = 1 AND (night_rush_quantity IS NULL OR night_rush_quantity = 0)")
    except Exception:
        pass

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
    """新增一筆操作紀錄並清除快取"""
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
        st.cache_data.clear()
    except Exception as e:
        print(f"[Audit Log Error] {e}")


@st.cache_data(ttl=600)
def get_audit_logs(limit: int = 200) -> pd.DataFrame:
    """取得稽核操作紀錄清單 - 快取 600 秒"""
    try:
        conn = get_db_connection()
        query = "SELECT timestamp as '時間戳記', operator as '操作者', action as '操作動作', details as '具體細節' FROM audit_logs ORDER BY id DESC LIMIT ?"
        df = pd.read_sql_query(query, conn, params=(limit,))
        conn.close()
        return df
    except Exception as e:
        return pd.DataFrame()


# ==============================================================================
# 3. 費用計算、週末連動與訂單邏輯輔助函式
# ==============================================================================
def calculate_booking_fee(
    area: str, quantity: int, check_in: date, check_out: date, night_rush_qty_str: str = ""
) -> Tuple[int, int, int, int]:
    """
    計算營位訂位費用：
    - 支援多區域字串解析 (如 "草地區 (2 帳)、通鋪房間 (1 間)")
    - 正常營位費用 = 各區域預約數量 * 單價 * 住宿天數
    - 夜衝費用 = 各區域夜衝數量 * (單價 * 0.5)
    回傳: (總金額, 住宿晚數, 正常營位費用, 夜衝費用)
    """
    area_dict = parse_area_quantities(str(area), int(quantity))
    rush_dict = parse_area_quantities(str(night_rush_qty_str), 0) if isinstance(night_rush_qty_str, str) else parse_area_quantities("", int(night_rush_qty_str))
    
    nights = max(1, (check_out - check_in).days)
    
    base_fee = 0
    night_rush_fee = 0
    
    # 計算基礎費用與夜衝費用分配
    for area_name, qty in area_dict.items():
        if area_name in CAMP_AREAS and qty > 0:
            price = CAMP_AREAS[area_name]["price_per_night"]
            base_fee += price * qty * nights
            
            # 使用獨立的夜衝數量
            apply_rush = max(0, min(qty, rush_dict.get(area_name, 0)))
            if apply_rush > 0:
                night_rush_fee += int(price * 0.5) * apply_rush
                
    total_fee = base_fee + night_rush_fee
    return (total_fee, nights, base_fee, night_rush_fee)


def get_upcoming_saturday(base_date: Optional[date] = None) -> date:
    """計算本週六（若今天是週六則為今天，若週日則為當天或下週六）"""
    if base_date is None:
        base_date = date.today()
    weekday = base_date.weekday()  # 週一=0, 週六=5, 週日=6
    if weekday == 5:
        return base_date
    elif weekday == 6:
        return base_date + timedelta(days=6)
    else:
        return base_date + timedelta(days=(5 - weekday))


def get_next_weekend(base_date: Optional[date] = None) -> date:
    """計算下週末（下週六）"""
    return get_upcoming_saturday(base_date) + timedelta(days=7)


# 熱門連續假期快捷日期清單 (台灣常見連續假期)
POPULAR_HOLIDAYS: List[Tuple[str, date]] = [
    ("國慶連假", date(2026, 10, 10)),
    ("元旦跨年", date(2027, 1, 1)),
    ("農曆春節", date(2027, 2, 6)),
    ("和平紀念日(228)", date(2027, 2, 27)),
    ("清明兒童連假", date(2027, 4, 3)),
    ("端午連假", date(2027, 6, 19)),
    ("中秋連假", date(2027, 9, 25)),
]


def parse_area_quantities(area_str: str, total_qty: int = 1) -> Dict[str, int]:
    """解析複合區域字串，回傳各區域的帳數字典"""
    import re
    res = {}
    if "(" in area_str:
        for k in CAMP_AREAS.keys():
            m = re.search(re.escape(k) + r"\s*\(\s*(\d+)", area_str)
            if m:
                res[k] = int(m.group(1))
    else:
        if area_str in CAMP_AREAS:
            res[area_str] = int(total_qty)
    return res


def is_shelter_group_booking(area: str, quantity: int) -> bool:
    """判斷是否為雨棚區包區 (雨棚區且預約數量 >= 6)"""
    parsed = parse_area_quantities(str(area), int(quantity))
    return parsed.get("雨棚區", 0) >= 6


@st.cache_data(ttl=600)
def get_dashboard_campsite_slots(target_date_str: str) -> Dict[str, Any]:
    """
    查詢 target_date_str 當天涵蓋的所有訂單，並映射為各區域各編號營位的詳細狀態
    支援「雨棚區包區機制」：若雨棚區訂單帳數 >= 6，自動連動標示為雨棚包區
    """
    conn = get_db_connection()
    # 查詢入住日期 <= 選擇日期 < 退房日期的所有訂單
    query = """
        SELECT * FROM bookings
        WHERE check_in_date <= ? AND check_out_date > ?
        ORDER BY check_in_date ASC, id ASC
    """
    df = pd.read_sql_query(query, conn, params=(target_date_str, target_date_str))
    conn.close()

    slots_by_area: Dict[str, List[Dict[str, Any]]] = {}
    shelter_group_info: Optional[Dict[str, Any]] = None

    for area_name, area_info in CAMP_AREAS.items():
        capacity = area_info["capacity"]
        unit = area_info["unit"]
        
        area_bookings = []
        for _, row in df.iterrows():
            parsed = parse_area_quantities(str(row.get("area", "")), int(row.get("quantity", 1)))
            if parsed.get(area_name, 0) > 0:
                b_copy = dict(row)
                b_copy["qty_for_this_area"] = parsed[area_name]
                area_bookings.append(b_copy)

        # 檢測是否有雨棚包區訂單
        has_shelter_group = False
        group_booking_ref = None
        if area_name == "雨棚區":
            for b in area_bookings:
                if is_shelter_group_booking(str(b.get("area", "")), int(b.get("quantity", 0))):
                    has_shelter_group = True
                    group_booking_ref = b
                    shelter_group_info = b
                    break

        slots: List[Dict[str, Any]] = []
        slot_idx = 1
        for b in area_bookings:
            qty = b["qty_for_this_area"]
            # 若勾選夜衝，則此訂單全額享有夜衝標籤 (依 checkbox 判定)
            # 相容舊資料：若有 night_rush_quantity，則也納入判定
            rush_qty = int(b.get("night_rush_quantity", 1 if b.get("night_rush") else 0))
            is_group = is_shelter_group_booking(str(b.get("area", "")), int(b.get("quantity", 1)))
            for i in range(qty):
                if slot_idx <= capacity:
                    has_night_rush = (b.get("night_rush") == 1) or (rush_qty > 0)
                    slots.append({
                        "slot_no": slot_idx,
                        "status": "booked",
                        "booking": b,
                        "has_night_rush": has_night_rush,
                        "rush_qty": rush_qty,
                        "is_group_booking": is_group,
                    })
                    slot_idx += 1

        # 剩餘未被個別分配的營位
        while slot_idx <= capacity:
            # 若為雨棚包區狀態，剩餘未滿帳之帳位亦連動標示為包區空額
            if area_name == "雨棚區" and has_shelter_group and group_booking_ref:
                slots.append({
                    "slot_no": slot_idx,
                    "status": "group_reserved",
                    "booking": group_booking_ref,
                    "has_night_rush": False,
                    "rush_qty": 0,
                    "is_group_booking": True,
                })
            else:
                slots.append({
                    "slot_no": slot_idx,
                    "status": "available",
                    "booking": None,
                    "has_night_rush": False,
                    "rush_qty": 0,
                    "is_group_booking": False,
                })
            slot_idx += 1

        slots_by_area[area_name] = slots

    return {
        "slots_by_area": slots_by_area,
        "shelter_group_info": shelter_group_info,
    }


def generate_booking_code(check_in: date) -> str:
    """產生唯一直覺的訂單編號 (如 LY-20261001-0812)"""
    rand_suffix = datetime.now().strftime("%H%M%S")[-4:]
    return f"LY-{check_in.strftime('%Y%m%d')}-{rand_suffix}"


@st.cache_data(ttl=600)
def query_historical_stats(phone: str, name: str = "") -> Tuple[int, str]:
    """根據電話 (與姓名) 查詢歷史來訪次數與歷史標籤 - 快取 600 秒"""
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


@st.cache_data(ttl=600)
def get_dashboard_summary(selected_date_str: str) -> Tuple[Dict[str, int], pd.DataFrame, pd.DataFrame]:
    """查詢選定日期各區域佔用數量、當日入住名單、當日退房名單 - 快取 600 秒"""
    conn = get_db_connection()
    query_occ = """
        SELECT area, quantity
        FROM bookings
        WHERE ? >= check_in_date AND ? < check_out_date
    """
    df_occupied = pd.read_sql_query(query_occ, conn, params=(selected_date_str, selected_date_str))
    
    occupied_map = {"雨棚區": 0, "草地區": 0, "通舖房間": 0}
    for _, row in df_occupied.iterrows():
        parsed = parse_area_quantities(str(row["area"]), int(row["quantity"]))
        for k, v in parsed.items():
            if k in occupied_map:
                occupied_map[k] += v

    df_in = pd.read_sql_query(
        "SELECT booking_code as '訂單編號', name as '姓名', phone as '電話', area as '區域', quantity as '數量', night_rush as '夜衝', tag as '標籤' FROM bookings WHERE check_in_date = ?",
        conn,
        params=(selected_date_str,),
    )

    df_out = pd.read_sql_query(
        "SELECT booking_code as '訂單編號', name as '姓名', phone as '電話', area as '區域', quantity as '數量', total_amount as '金額' FROM bookings WHERE check_out_date = ?",
        conn,
        params=(selected_date_str,),
    )
    conn.close()
    return occupied_map, df_in, df_out


@st.cache_data(ttl=600)
def fetch_bookings_df(query: str, params: Tuple[Any, ...]) -> pd.DataFrame:
    """查詢過濾訂單列表 - 快取 600 秒"""
    conn = get_db_connection()
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


@st.cache_data(ttl=600)
def fetch_all_bookings_df() -> pd.DataFrame:
    """查詢全量訂單列表 - 快取 600 秒"""
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM bookings ORDER BY check_in_date DESC", conn)
    conn.close()
    return df


def load_data() -> pd.DataFrame:
    """載入全量訂單資料 (統一資料載入入口)

    於每次應用程式重跑/啟動時由 main() 呼叫，確保 Streamlit 睡眠喚醒後
    session_state 與各頁面 UI 元件皆能正確載入最新訂單資料。
    """
    return fetch_all_bookings_df()


def refresh_session_state_data() -> pd.DataFrame:
    """同步完成後重新查詢全量訂單並明確寫入 st.session_state

    於 Google Sheets 拉取/同步流程完成後呼叫，確保表格、指標與
    日曆等 UI 元件皆能取得非空的最新訂單資料。
    """
    df = load_data()
    st.session_state["bookings_df"] = df
    st.session_state["bookings_last_refreshed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return df


def deduplicate_bookings() -> int:
    """啟動時自動清理 bookings 表中的重複訂單

    依複合鍵 (入住日期 check_in_date, 露營人姓名 name,
    露營人電話 phone) 分組，每組僅保留 id 最大
    (最近建立/最新版本) 的記錄，刪除其餘重複項。

    回傳刪除的重複筆數 (0 表示無重複)；失敗時回傳 -1，
    不影響啟動流程。
    """
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(
            """
            DELETE FROM bookings
            WHERE id NOT IN (
                SELECT MAX(id)
                FROM bookings
                GROUP BY check_in_date, name, phone
            )
            """
        )
        deleted = max(0, cursor.rowcount)
        conn.commit()
        conn.close()

        if deleted > 0:
            print(f"[Dedup] 啟動自動去重：刪除重複訂單 {deleted} 筆 (保留最新版本)")
            add_audit_log(
                "系統",
                "自動去重清理",
                f"啟動時依 (入住日期, 姓名, 電話) 複合鍵自動清理重複訂單 {deleted} 筆，僅保留最新版本",
            )
        return deleted
    except Exception as e:
        # 清理失敗不影響啟動流程，僅背景記錄
        print(f"[Dedup] 啟動自動去重失敗 (不影響啟動): {e}")
        return -1


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

    def __init__(self, folder_id: Optional[str] = None):
        self.folder_id = folder_id or get_gdrive_folder_id()
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

    def upload_to_google_drive(self, file_content_bytes: bytes, filename: str, mime_type: str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") -> Tuple[bool, str]:
        """
        上傳檔案至目標 Google Drive 資料夾 (DVCG)。

        針對 Service Account 無 Drive 儲存配額的問題
        (HttpError 403: Service Accounts do not have storage quota)：
        1. file_metadata 明確指定 parents 為目標資料夾 (st.secrets['google_drive']['folder_id'])
        2. 優先「既有檔案覆蓋 (Update)」：覆寫舊檔不消耗新的儲存配額
        3. 資料夾內找不到時，全 Drive 搜尋同名舊檔 (含早期建立在根目錄的檔案)，
           覆寫內容並搬移 (addParents/removeParents) 至目標資料夾
        4. 皆找不到才新建檔；若新建遭遇配額錯誤，回傳明確處理指引
        """
        if not self.is_connected or not self.service:
            return False, f"Google Drive 未連線：{self.status_message}"

        # 目標資料夾 ID：優先取用 st.secrets['google_drive']['folder_id']，
        # 其餘來源 (gcp_service_account.folder_id / 環境變數 / DVCG 預設) 由 get_gdrive_folder_id() 支援
        FOLDER_ID = self.folder_id or get_gdrive_folder_id()
        drive_service = self.service

        try:
            from googleapiclient.http import MediaIoBaseUpload

            def _build_media():
                return MediaIoBaseUpload(io.BytesIO(file_content_bytes), mimetype=mime_type, resumable=True)

            media = _build_media()

            # 1. 上傳前先查詢 DVCG 資料夾內是否已存在同名檔案
            escaped_name = filename.replace("'", "\\'")
            query = f"'{FOLDER_ID}' in parents and name = '{escaped_name}' and trashed = false"
            response = drive_service.files().list(
                q=query,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
                fields="files(id, name, webViewLink)",
            ).execute()
            items = response.get("files", [])

            # 上傳元資料：務必明確指定 parents，確保檔案直接建立在共用資料夾內部
            file_metadata = {
                "name": filename,
                "parents": [FOLDER_ID],
            }

            if len(items) > 0:
                # 2. 資料夾內已存在：使用 update 覆蓋舊檔 (不消耗 Service Account 儲存配額)
                file = drive_service.files().update(
                    fileId=items[0]["id"],
                    media_body=media,
                    supportsAllDrives=True,
                ).execute()
                return True, f"檔案已成功覆蓋更新至雲端！檔案 ID: {file.get('id')}，檔名: {filename}"

            # 3. 資料夾內找不到：全 Drive 搜尋同名檔 (處理早期建立在根目錄的舊檔)
            global_query = f"name = '{escaped_name}' and trashed = false"
            global_response = drive_service.files().list(
                q=global_query,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
                fields="files(id, name, parents, webViewLink)",
            ).execute()
            global_items = global_response.get("files", [])

            if len(global_items) > 0:
                existing = global_items[0]
                existing_id = existing["id"]
                old_parents = existing.get("parents") or []
                move_params = {}
                if FOLDER_ID not in old_parents:
                    move_params["addParents"] = FOLDER_ID
                    if old_parents:
                        move_params["removeParents"] = ",".join(old_parents)
                try:
                    # 覆寫內容並搬移至 DVCG 資料夾 (搬移檔案不消耗配額)
                    file = drive_service.files().update(
                        fileId=existing_id,
                        media_body=media,
                        supportsAllDrives=True,
                        **move_params,
                    ).execute()
                    return True, f"檔案已覆寫並搬移至 DVCG 資料夾！檔案 ID: {file.get('id')}，檔名: {filename}"
                except Exception:
                    # 搬移失敗時退而只覆寫內容 (保留原位置，至少確保備份內容正確)
                    fallback_media = _build_media()
                    file = drive_service.files().update(
                        fileId=existing_id,
                        media_body=fallback_media,
                        supportsAllDrives=True,
                    ).execute()
                    return True, f"檔案已覆寫更新至雲端 (保留原資料夾位置)！檔案 ID: {file.get('id')}，檔名: {filename}"

            # 4. 完全不存在：建立新檔並強制掛載在 FOLDER_ID 下
            file = drive_service.files().create(
                body=file_metadata,
                media_body=media,
                supportsAllDrives=True,
                fields="id, name, webViewLink",
            ).execute()
            return True, f"檔案成功備份至雲端！檔案 ID: {file.get('id')}，檔名: {filename}"
        except Exception as e:
            error_text = str(e).lower()
            if "quota" in error_text:
                return False, (
                    f"上傳至 Google Drive 失敗：Service Account 沒有 Drive 儲存配額，無法新建檔案。"
                    f"請先手動在 DVCG 資料夾內建立「{filename}」，並將該檔案分享給服務帳號信箱"
                    "(編輯者權限) 後再試；或将備份資料夾設為 Shared Drive 以使用組織儲存空間。"
                )
            return False, f"上傳至 Google Drive 失敗: {str(e)}"

    # 方法別名以相容既有呼叫
    upload_file_to_drive = upload_to_google_drive
    sync_to_drive = upload_to_google_drive

    def backup_database_now(self, operator: str = "系統") -> Tuple[bool, str]:
        """
        資料庫備份：
        - 本地備份：Excel 檔案存至 data/ 備份目錄
        - 雲端備份：採「Google Sheets 完整資料庫轉譯備份」(DB_Full_Backup 分頁)
          (個人版 Google Drive 的 Service Account 預設無儲存配額，無法作為檔案
           擁有者上傳實體檔案，故不再呼叫 Drive API，改以 gspread 寫入試算表)
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

            # 雲端備份：Google Sheets 完整資料庫轉譯 (取代 Drive API 實體檔案上傳)
            success, msg = backup_database_to_google_sheets_full()
            if success:
                add_audit_log(operator, "雲端同步備份", f"已成功同步至 Google Sheets「{DB_FULL_BACKUP_WORKSHEET_TITLE}」分頁")
                st.toast("✅ 本地資料庫與雲端 Google Sheets 完整備份成功！")
                return True, f"本地與雲端備份皆成功！\n雲端：{msg}\n本地：{local_backup_path}"
            else:
                add_audit_log(operator, "雲端同步異常", f"雲端備份失敗，已保存在本地備份：{local_backup_path}")
                return False, f"本地備份成功，但雲端 Google Sheets 備份失敗：{msg}"
        except Exception as e:
            return False, f"備份過程發生錯誤: {str(e)}"


# 建立全域或快取 GDriveSync 實例
@st.cache_resource
def get_gdrive_sync(folder_id: Optional[str] = None):
    return GoogleDriveSync(folder_id=folder_id)


def upload_to_google_drive(file_content_bytes: bytes, filename: str, mime_type: str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", folder_id: Optional[str] = None) -> Tuple[bool, str]:
    """模組級函式：上傳或覆蓋備份檔案至 Google Drive DVCG 資料夾"""
    sync_tool = get_gdrive_sync(folder_id=folder_id)
    return sync_tool.upload_to_google_drive(file_content_bytes, filename, mime_type)

sync_to_drive = upload_to_google_drive


def trigger_auto_sync(operator: str):
    """在訂單、管理員或日誌變動時觸發自動備份"""
    try:
        sync_tool = get_gdrive_sync()
        sync_tool.backup_database_now(operator=operator)
    except Exception as e:
        print(f"[Auto Sync Warning] {e}")


# ==============================================================================
# 4-1. Google Sheets 即時同步模組 (gspread)
# ==============================================================================
GSHEET_WORKSHEET_TITLE = "訂單紀錄"
DB_FULL_BACKUP_WORKSHEET_TITLE = "DB_Full_Backup"

# SQLite 欄位 → Google Sheets 中文標頭對照 (與 Excel 匯出欄位一致)
GSHEET_COLUMN_MAP = {
    "booking_code": "訂單編號",
    "check_in_date": "入住日期",
    "check_out_date": "退房日期",
    "name": "露營人姓名",
    "phone": "露營人電話",
    "area": "預約區域",
    "quantity": "帳數/間數",
    "night_rush": "是否夜衝",
    "night_rush_quantity": "夜衝數量",
    "night_rush_detail": "夜衝明細",
    "visit_count": "來過次數",
    "tag": "客戶標籤",
    "notes": "備註說明",
    "total_amount": "總計金額",
    "created_by": "建立人員",
    "created_at": "建立時間",
    "updated_at": "最後更新時間",
}


def get_spreadsheet_id() -> str:
    """取得 Google Sheets 目標試算表 ID

    防呆讀取：依序檢查 st.secrets 的 [google_drive] 或 [google_sheets] 區段，
    並支援 SPREADSHEET_ID / spreadsheet_id 大寫或小寫鍵名。
    (註：st.secrets 巢狀區段回傳 AttrDict，屬 Mapping 而非 dict 的子類別，
     因此不可用 isinstance(..., dict) 判斷，否則在真實 Streamlit 環境永遠讀取失敗)
    """
    try:
        if hasattr(st, "secrets"):
            # 嘗試從 [google_drive] 或 [google_sheets] 讀取大寫或小寫的 ID
            drive_secrets = st.secrets.get("google_drive", st.secrets.get("google_sheets", {}))
            spreadsheet_id = drive_secrets.get("SPREADSHEET_ID") or drive_secrets.get("spreadsheet_id")

            if not spreadsheet_id:
                print("[Google Sheets] 未設定試算表 ID，請檢查 .streamlit/secrets.toml")
                return ""
            return str(spreadsheet_id).strip()
    except Exception as e:
        print(f"[Google Sheets] 讀取 SPREADSHEET_ID 失敗: {e}")
    return ""


@st.cache_resource(show_spinner=False)
def get_gspread_client():
    """建立 Google Sheets API 連線客戶端 (gspread) 並快取，避免每次重新認證

    認證來源：優先 st.secrets['gcp_service_account']，其次本機 data/service_account.json；
    連線後會透過 open_by_key 驗證目標試算表可存取。
    """
    from google.oauth2 import service_account
    import gspread

    creds = None
    # 認證來源 1: Streamlit Secrets
    if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
        try:
            creds_info = dict(st.secrets["gcp_service_account"])
            creds = service_account.Credentials.from_service_account_info(
                creds_info,
                scopes=["https://www.googleapis.com/auth/spreadsheets"],
            )
        except Exception as e:
            print(f"[Google Sheets] Secrets 認證失敗: {e}")

    # 認證來源 2: 本機 service_account.json (本地開發環境)
    if creds is None and os.path.exists(SERVICE_ACCOUNT_PATH):
        try:
            creds = service_account.Credentials.from_service_account_file(
                SERVICE_ACCOUNT_PATH,
                scopes=["https://www.googleapis.com/auth/spreadsheets"],
            )
        except Exception as e:
            print(f"[Google Sheets] 本機憑證認證失敗: {e}")

    if creds is None:
        raise RuntimeError(
            "找不到 Google Service Account 憑證，請設定 st.secrets['gcp_service_account'] 或放置 data/service_account.json"
        )

    spreadsheet_id = get_spreadsheet_id()
    if not spreadsheet_id:
        raise RuntimeError(
            "找不到目標試算表 ID，請設定 st.secrets['google_drive']['SPREADSHEET_ID'] 或 st.secrets['google_sheets']['spreadsheet_id']"
        )

    client = gspread.authorize(creds)
    # 連線至指定試算表並驗證可存取性
    client.open_by_key(spreadsheet_id)
    return client


def sync_db_to_google_sheet() -> bool:
    """將 SQLite bookings 資料表全量同步至 Google Sheets 指定工作表

    流程：讀取最新完整訂單 DataFrame → 全欄位轉字串 → 清除工作表舊資料 → 一次性覆寫標頭與資料列。
    同步失敗僅顯示警告訊息，不影響 SQLite 原本的寫入流程。
    """
    try:
        client = get_gspread_client()
        spreadsheet_id = get_spreadsheet_id()
        spreadsheet = client.open_by_key(spreadsheet_id)

        # 讀取 SQLite 最新完整訂單資料
        conn = get_db_connection()
        df = pd.read_sql_query(
            """
            SELECT booking_code, check_in_date, check_out_date, name, phone, area,
                   quantity, night_rush, night_rush_quantity, night_rush_detail,
                   visit_count, tag, notes, total_amount, created_by, created_at, updated_at
            FROM bookings
            ORDER BY check_in_date DESC
            """,
            conn,
        )
        conn.close()

        # 空值以空字串表示，並將所有欄位 (含日期/時間/數字) 全數轉換為字串格式
        df = df.fillna("").astype(str)

        # 取得目標工作表；若不存在則自動建立
        worksheet = None
        for ws in spreadsheet.worksheets():
            if ws.title == GSHEET_WORKSHEET_TITLE:
                worksheet = ws
                break
        if worksheet is None:
            worksheet = spreadsheet.add_worksheet(
                title=GSHEET_WORKSHEET_TITLE,
                rows=1000,
                cols=len(GSHEET_COLUMN_MAP),
            )

        # 組合中文標頭與資料列
        df = df.rename(columns=GSHEET_COLUMN_MAP)
        header = [str(column) for column in df.columns]
        data = [header] + df.values.tolist()

        # 清除工作表舊資料並一次性覆寫，確保 Google Sheets 與 SQLite 資料完全一致
        worksheet.clear()
        worksheet.update(range_name="A1", values=data)

        st.toast(f"☁️ Google Sheets 同步完成，共 {len(df)} 筆訂單", icon="✅")
        return True
    except Exception as e:
        st.warning(f"⚠️ Google Sheets 即時同步失敗（SQLite 資料不受影響）：{e}")
        return False


def _normalize_text(value: Any) -> str:
    """字串標準化：去除首尾空白"""
    return "" if value is None else str(value).strip()


def _normalize_phone(value: Any) -> str:
    """電話格式標準化：去除空格、破折號與括號，確保跨系統比對一致"""
    if value is None:
        return ""
    s = str(value).strip()
    for ch in (" ", "　", "-", "－", "(", ")", "（", "）"):
        s = s.replace(ch, "")
    return s


def _normalize_date_str(value: Any) -> str:
    """日期字串標準化為 YYYY-MM-DD

    支援 Google Sheets 常見格式：2026-10-25、2026/10/25、
    2026年10月25日、2026-10-25 00:00:00、純數字 20261025。
    """
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    m = re.search(r"(\d{4})\s*[-/年]\s*(\d{1,2})\s*[-/月]\s*(\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    digits = re.sub(r"\D", "", s)
    if len(digits) == 8:
        return f"{digits[:4]}-{digits[4:6]}-{digits[6:8]}"
    return s[:10]


def _normalize_int(value: Any, default: int = 0) -> int:
    """數值標準化：轉換失敗時回傳預設值"""
    try:
        return int(float(str(value).strip()))
    except (TypeError, ValueError):
        return default


def pull_google_sheet_to_db() -> Tuple[int, int]:
    """從 Google Sheets「訂單紀錄」分頁拉取訂單，以最新資料覆寫本機 SQLite

    覆寫規則 (複合鍵)：入住日期 (check_in_date) + 露營人姓名 (name) +
    露營人電話 (phone) 三者完全一致即視為同一筆訂單；
    雲端列代表最新更新 (區域/備註/客戶資料等)，故先刪除
    SQLite 既有舊記錄再插入最新版本。

    比對前會先標準化欄位格式 (去除空白、電話去符號、日期統一為
    YYYY-MM-DD)，以容忍 Sheets 與 SQLite 之間的格式差異。

    回傳 (寫入筆數, 跳過無效列數)；試算表不可達或憑證缺失等任何失敗
    均僅背景記錄日誌並回傳 (0, 0)，不會中斷 UI 流程。
    """
    try:
        spreadsheet_id = get_spreadsheet_id()
        if not spreadsheet_id:
            print("[Google Sheets Pull] 未設定試算表 ID，跳過啟動拉取")
            return 0, 0

        client = get_gspread_client()
        spreadsheet = client.open_by_key(spreadsheet_id)

        # 取得目標工作表 (若不存在視為尚無雲端資料可拉取)
        worksheet = None
        for ws in spreadsheet.worksheets():
            if ws.title == GSHEET_WORKSHEET_TITLE:
                worksheet = ws
                break
        if worksheet is None:
            print(f"[Google Sheets Pull] 工作表「{GSHEET_WORKSHEET_TITLE}」不存在，跳過")
            return 0, 0

        rows = worksheet.get_all_values()
        if not rows or len(rows) < 2:
            return 0, 0

        # 建立中文標頭 → SQLite 欄位反向對照
        header_to_column = {chinese: col for col, chinese in GSHEET_COLUMN_MAP.items()}
        header = [str(h).strip() for h in rows[0]]
        column_index = {}
        for idx, h in enumerate(header):
            col = header_to_column.get(h)
            if col:
                column_index[col] = idx

        # 必要欄位檢查 (去重複合鍵 + 入住/退房日)
        required = ("check_in_date", "name", "phone", "check_out_date")
        missing = [c for c in required if c not in column_index]
        if missing:
            print(f"[Google Sheets Pull] 工作表缺少必要欄位 {missing}，跳過")
            return 0, 0

        def _cell(row: List[Any], col: str, default: str = "") -> str:
            idx = column_index.get(col)
            if idx is None or idx >= len(row):
                return default
            return str(row[idx]).strip()

        # 讀取 SQLite 既有訂單的複合鍵 → id 集合 (標準化後比對)
        conn = get_db_connection()
        existing_key_ids = {}
        for r in conn.execute("SELECT id, check_in_date, name, phone FROM bookings").fetchall():
            _key = (
                _normalize_date_str(r[1]),
                _normalize_text(r[2]),
                _normalize_phone(r[3]),
            )
            existing_key_ids.setdefault(_key, []).append(r[0])

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        skipped = 0
        # 以複合鍵建立最新版本映射：批次內更晚出現的列自動覆寫先前列
        latest_records = {}
        for raw in rows[1:]:
            row = list(raw) + [""] * max(0, len(header) - len(raw))

            check_in_str = _normalize_date_str(_cell(row, "check_in_date"))
            name_val = _normalize_text(_cell(row, "name"))
            phone_val = _normalize_phone(_cell(row, "phone"))
            check_out_str = _normalize_date_str(_cell(row, "check_out_date"))

            # 空白列或關鍵欄位缺失則跳過
            if not check_in_str or not name_val or not phone_val or not check_out_str:
                skipped += 1
                continue

            composite_key = (check_in_str, name_val, phone_val)
            booking_code = _cell(row, "booking_code") or generate_booking_code(date.today())
            night_rush_raw = _cell(row, "night_rush").lower()
            latest_records[composite_key] = (
                booking_code,
                check_in_str,
                check_out_str,
                name_val,
                phone_val,
                _cell(row, "area"),
                _normalize_int(_cell(row, "quantity"), 1) or 1,
                1 if night_rush_raw in ("1", "true", "是", "y", "yes") else 0,
                _normalize_int(_cell(row, "night_rush_quantity"), 0),
                _cell(row, "night_rush_detail"),
                _normalize_int(_cell(row, "visit_count"), 1) or 1,
                _cell(row, "tag") or "一般",
                _cell(row, "notes"),
                _normalize_int(_cell(row, "total_amount"), 0),
                _cell(row, "created_by") or "系統匯入",
                _cell(row, "created_at") or now_str,
                _cell(row, "updated_at") or now_str,
            )

        # 最新覆寫：依複合鍵收集既有舊記錄 id 並刪除
        # (雲端列代表最新更新，取代 SQLite 既有版本)
        delete_ids = []
        for key in latest_records:
            delete_ids.extend(existing_key_ids.get(key, []))
        if delete_ids:
            conn.executemany(
                "DELETE FROM bookings WHERE id = ?",
                [(i,) for i in delete_ids],
            )

        # 批次插入最新版本並清除資料快取
        if latest_records:
            conn.executemany(
                """
                INSERT INTO bookings (
                    booking_code, check_in_date, check_out_date, name, phone,
                    area, quantity, night_rush, night_rush_quantity, night_rush_detail,
                    visit_count, tag, notes, total_amount, created_by, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                list(latest_records.values()),
            )
            conn.commit()
            st.cache_data.clear()
        conn.close()

        inserted = len(latest_records)
        overwritten = len(delete_ids)

        # 更新 session_state 啟動拉取狀態
        st.session_state["gsheet_startup_pull_done"] = True
        st.session_state["gsheet_last_pull_at"] = now_str
        st.session_state["gsheet_last_pull_new"] = inserted
        st.session_state["gsheet_last_pull_overwritten"] = overwritten
        st.session_state["gsheet_last_pull_skipped"] = skipped

        print(
            f"[Google Sheets Pull] 啟動拉取完成：寫入 {inserted} 筆"
            f" (覆寫既有 {overwritten} 筆)，跳過無效 {skipped} 筆"
        )
        return inserted, skipped
    except Exception as e:
        # 試算表不可達/憑證異常等：背景記錄並於畫面顯示警告，但不中斷 UI
        print(f"[Google Sheets Pull] 啟動拉取失敗 (UI 不受影響)：{e}")
        st.warning(f"⚠️ Google Sheets 啟動拉取失敗，將以本機 SQLite 資料顯示：{e}")
        st.session_state["gsheet_startup_pull_done"] = True
        return 0, 0


def backup_database_to_google_sheets_full() -> Tuple[bool, str]:
    """將 SQLite 完整資料庫轉譯並覆寫寫入 Google Sheets 的 DB_Full_Backup 分頁

    讀取所有資料表 (bookings、audit_logs 等)，以「表名標題 → 欄位名稱 → 全部資料列」
    的完整表格格式依序寫入，取代實體檔案上傳
    (個人版 Google Drive 的 Service Account 預設無儲存配額，無法作為檔案擁有者上傳)。
    """
    try:
        spreadsheet_id = get_spreadsheet_id()
        if not spreadsheet_id:
            st.error("⚠️ 未設定 Google Sheets 試算表 ID，請檢查 .streamlit/secrets.toml！")
            return False, "未設定 Google Sheets 試算表 ID，請檢查 .streamlit/secrets.toml 設定"

        client = get_gspread_client()
        spreadsheet = client.open_by_key(spreadsheet_id)

        # 讀取 SQLite 所有資料表 (排除 SQLite 內部表)
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
        )
        table_names = [row[0] for row in cursor.fetchall()]

        # 組合完整表格資料：每個資料表依序寫入 [表名標題列 / 欄位名稱列 / 資料列 / 空行分隔]
        data_rows: List[List[str]] = []
        for table_name in table_names:
            cursor.execute(f'SELECT * FROM "{table_name}"')
            columns = [str(desc[0]) for desc in cursor.description]
            rows = cursor.fetchall()

            data_rows.append([f"📦 資料表：{table_name} (共 {len(rows)} 列)"])
            data_rows.append(columns)
            for row in rows:
                data_rows.append(["" if value is None else str(value) for value in row])
            data_rows.append([])  # 空行分隔
        conn.close()

        if not data_rows:
            data_rows = [["(資料庫目前無任何資料表)"]]

        # 取得或建立 DB_Full_Backup 工作表
        worksheet = None
        for ws in spreadsheet.worksheets():
            if ws.title == DB_FULL_BACKUP_WORKSHEET_TITLE:
                worksheet = ws
                break
        if worksheet is None:
            worksheet = spreadsheet.add_worksheet(
                title=DB_FULL_BACKUP_WORKSHEET_TITLE,
                rows=max(1000, len(data_rows) + 100),
                cols=40,
            )
        elif worksheet.row_count < len(data_rows) + 10:
            # 資料量成長超過既有工作表容量時自動擴容，避免寫入失敗
            worksheet.resize(rows=len(data_rows) + 100, cols=max(40, worksheet.col_count))

        # 清除舊資料並一次性覆寫，確保 Google Sheets 與 SQLite 資料庫完全一致
        worksheet.clear()
        worksheet.update(range_name="A1", values=data_rows)

        return True, f"已將 {len(table_names)} 個資料表完整轉譯至 Google Sheets「{DB_FULL_BACKUP_WORKSHEET_TITLE}」分頁"
    except Exception as e:
        return False, str(e)


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


def inject_auto_logout_js(timeout_seconds: int = 60):
    """在 Streamlit 應用程式中嵌入輕量 JavaScript 監聽使用者動作，60 秒無動作自動登出"""
    js_code = f"""
    <script>
    (function() {{
        const TIMEOUT_MS = {timeout_seconds} * 1000;
        let timer = null;

        function triggerLogout() {{
            try {{
                const targetWin = window.parent || window;
                const currentUrl = new URL(targetWin.location.href);
                currentUrl.searchParams.set("auto_logout", "1");
                targetWin.location.href = currentUrl.toString();
            }} catch (e) {{
                window.location.search = "?auto_logout=1";
            }}
        }}

        function resetTimer() {{
            if (timer) clearTimeout(timer);
            timer = setTimeout(triggerLogout, TIMEOUT_MS);
        }}

        const events = ['mousemove', 'keydown', 'click', 'scroll', 'touchstart'];
        try {{
            const targetDoc = window.parent.document || document;
            events.forEach(function(evt) {{
                targetDoc.addEventListener(evt, resetTimer, true);
            }});
        }} catch (e) {{
            events.forEach(function(evt) {{
                document.addEventListener(evt, resetTimer, true);
            }});
        }}
        resetTimer();
    }})();
    </script>
    """
    components.html(js_code, height=0, width=0)


def render_login_view():
    """渲染極簡安全 PIN 碼登入畫面"""
    col1, col2, col3 = st.columns([1, 1.6, 1])
    with col2:
        if st.session_state.get("timeout_logged_out"):
            st.warning("因超過 60 秒無操作，系統已自動登出以維護資安。")
            st.session_state["timeout_logged_out"] = False

        st.markdown(
            """
            <div style="text-align: center; padding: 32px 24px; background: white; border-radius: 16px; box-shadow: 0 8px 24px rgba(0,0,0,0.06); border-top: 6px solid #2E7D32;">
                <h1 style="color: #1B5E20; font-size: 2.2rem; margin-bottom: 8px;">⛺ 祿野營地</h1>
                <h3 style="color: #424242; font-weight: 500; margin-top: 0; margin-bottom: 6px;">訂位管理系統</h3>
                <p style="color: #757575; font-size: 0.95rem; margin-bottom: 0;">請輸入管理員 PIN 碼</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.write("")

        with st.form("login_form"):
            pin_input = st.text_input("🔑 管理員 PIN 碼", type="password", max_chars=8, placeholder="請輸入 PIN 碼...")
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
                    st.error("PIN 碼錯誤，請重新輸入")


# ==============================================================================
# 6-1. 刪除後端與卡片內嵌編輯/刪除控件 (主鍵 id 綁定)
# ==============================================================================
def execute_delete_booking(
    booking_id: int,
    booking_info: Optional[Dict[str, Any]] = None,
) -> None:
    """依主鍵 id 刪除訂單並立即重整 UI 與雲端同步

    刪除後端共用邏後端共用邏輯：參數化 DELETE ... WHERE id = ? →
    清除資料快取 → 重查並重整 session_state 資料集 →
    即時同步 Google Sheets → 寫入稽核日誌 → 觸發自動備份。
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM bookings WHERE id = ?", (booking_id,))
    conn.commit()
    conn.close()

    st.cache_data.clear()
    # 立即重查並重整 session_state 資料集
    refresh_session_state_data()
    # SQLite 刪除成功後，即時同步至 Google Sheets
    sync_db_to_google_sheet()

    current_user = st.session_state.get("authenticated_user", "管理員")
    info = booking_info or {}
    add_audit_log(
        current_user,
        "刪除訂單",
        f"刪除訂單 {info.get('booking_code', '')} (露營人: {info.get('name', '')}, 電話: {info.get('phone', '')}, 區域: {info.get('area', '')}, 金額: {info.get('total_amount', '')})",
    )
    trigger_auto_sync(current_user)


def _parse_date_value(raw_value: Any, fallback: date) -> date:
    """安全解析訂單日期字串為 date 物件

    歷史資料可能含占位日期 (如「歷史客戶/無預約日期」)，
    解析失敗時回退至 fallback，避免 date_input 拋錯。
    """
    try:
        return datetime.strptime(str(raw_value)[:10], "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return fallback


# ==============================================================================
# 6-2. 編輯與刪除卡片回調 (Session State 驅動)
# ==============================================================================
def trigger_edit_booking(record_id: int, row_data: Dict[str, Any]):
    """編輯卡片回調：載入資料並跳轉至編輯表單"""
    st.session_state["editing_booking_id"] = record_id
    st.session_state["editing_booking_data"] = row_data
    st.session_state["jump_to_page"] = "➕ 新增露營預約"
    reset_booking_form_date_keys()


def trigger_delete_dialog(record_id: int, row_data: Dict[str, Any]):
    """刪除卡片回調：觸發刪除對話框"""
    st.session_state["deleting_booking_id"] = record_id
    st.session_state["deleting_booking_data"] = row_data


# ==============================================================================
# 6-3. 刪除確認對話框 (原生 Modal)
# ==============================================================================
@st.dialog("⚠️ 確認刪除預約")
def render_delete_confirm_modal():
    """刪除確認視窗：執行 DELETE 並重整資料"""
    booking_id = st.session_state["deleting_booking_id"]
    row = st.session_state["deleting_booking_data"]
    
    st.write(f"確定要刪除這筆預約嗎？")
    st.info(f"姓名: {row['name']}\n入住日期: {row['check_in_date']}")
    
    col1, col2 = st.columns(2)
    if col1.button("確定刪除", type="primary", use_container_width=True):
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM bookings WHERE id = ?", (booking_id,))
        conn.commit()
        conn.close()
        
        st.cache_data.clear()
        refresh_session_state_data()
        sync_db_to_google_sheet()
        
        # 稽核日誌
        current_user = st.session_state.get("authenticated_user", "管理員")
        add_audit_log(current_user, "刪除訂單", f"刪除訂單 {row.get('booking_code')} (ID {booking_id})")
        trigger_auto_sync(current_user)
        
        st.session_state["deleting_booking_id"] = None
        st.session_state["deleting_booking_data"] = None
        st.rerun()
        
    if col2.button("取消", use_container_width=True):
        st.session_state["deleting_booking_id"] = None
        st.session_state["deleting_booking_data"] = None
        st.rerun()


def render_card_edit_delete_controls(row: pd.Series) -> None:
    """訂單卡片內嵌的編輯/刪除按鈕 (on_click 綁定)"""
    btn_c1, btn_c2 = st.columns([1, 1])
    with btn_c1:
        st.button(
            "✏️ 編輯", 
            key=f"card_edit_{row['id']}", 
            use_container_width=True,
            on_click=trigger_edit_booking,
            args=(row["id"], dict(row))
        )
    with btn_c2:
        st.button(
            "🗑️ 刪除", 
            key=f"card_del_{row['id']}", 
            use_container_width=True,
            on_click=trigger_delete_dialog,
            args=(row["id"], dict(row))
        )


# ==============================================================================
# 7. 頁面功能：營位總覽看板 (日曆式看板與點擊拉出資料卡抽屜)
# ==============================================================================
def render_dashboard_view():
    st.markdown("### 📊 營區營位總覽看板 (日曆視角)")

    # 1. 統一管理狀態
    if "dash_view_mode" not in st.session_state:
        st.session_state["dash_view_mode"] = "週視角"
    if "dashboard_selected_date" not in st.session_state:
        st.session_state["dashboard_selected_date"] = date.today()
    if "dash_calendar_base" not in st.session_state:
        st.session_state["dash_calendar_base"] = date.today()

    # 當月月份醒目顯示 (Prominent Month Display)
    base_dt = st.session_state["dash_calendar_base"]
    st.markdown(
        f"""
        <div style="background: linear-gradient(135deg, #1B5E20, #2E7D32); color: white; padding: 16px 24px; border-radius: 12px; margin-bottom: 20px; box-shadow: 0 4px 12px rgba(46, 125, 50, 0.2); text-align: center;">
            <span style="font-size: 1.8rem; font-weight: 800; letter-spacing: 2px;">📅 {base_dt.year} 年 {base_dt.month} 月</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # 視角切換器與快捷跳轉按鈕工具列 (全面同步 session_state)
    b_col1, b_col2, b_col3 = st.columns([1.5, 3, 1.5])
    with b_col1:
        view_mode = st.radio("切換日曆視角", ["週視角", "月視角"], index=0 if st.session_state["dash_view_mode"] == "週視角" else 1, horizontal=True, label_visibility="collapsed")
        if view_mode != st.session_state["dash_view_mode"]:
            st.session_state["dash_view_mode"] = view_mode
            st.rerun()

    with b_col2:
        c1, c2, c3 = st.columns(3)
        if view_mode == "週視角":
            if c1.button("⬅️ 上週", use_container_width=True): 
                st.session_state["dash_calendar_base"] -= timedelta(days=7)
                st.rerun()
            if c2.button("📅 本週", use_container_width=True): 
                st.session_state["dash_calendar_base"] = date.today()
                st.session_state["dashboard_selected_date"] = date.today()
                st.rerun()
            if c3.button("下週 ➡️", use_container_width=True): 
                st.session_state["dash_calendar_base"] += timedelta(days=7)
                st.rerun()
        else:
            if c1.button("⬅️ 上個月", use_container_width=True):
                base = st.session_state["dash_calendar_base"]
                st.session_state["dash_calendar_base"] = (base.replace(day=1) - timedelta(days=1)).replace(day=1)
                st.rerun()
            if c2.button("📅 本月", use_container_width=True): 
                st.session_state["dash_calendar_base"] = date.today()
                st.session_state["dashboard_selected_date"] = date.today()
                st.rerun()
            if c3.button("下個月 ➡️", use_container_width=True):
                base = st.session_state["dash_calendar_base"]
                st.session_state["dash_calendar_base"] = (base.replace(day=28) + timedelta(days=4)).replace(day=1)
                st.rerun()

    with b_col3:
        def on_dashboard_date_change():
            st.session_state["dashboard_selected_date"] = st.session_state["dash_date_picker"]
            st.session_state["dash_calendar_base"] = st.session_state["dash_date_picker"]

        if "dash_date_picker" not in st.session_state or st.session_state["dash_date_picker"] != st.session_state["dashboard_selected_date"]:
            st.session_state["dash_date_picker"] = st.session_state["dashboard_selected_date"]

        st.date_input(
            "📅 指定日期：",
            key="dash_date_picker",
            min_value=date.today() - timedelta(days=365),
            on_change=on_dashboard_date_change,
        )

    # 2. 取出整批資料作快取計算與日曆網格渲染
    df_all = fetch_all_bookings_df()

    import calendar
    cal = calendar.Calendar(firstweekday=0)
    base_dt = st.session_state["dash_calendar_base"]

    if view_mode == "週視角":
        start_dt = base_dt - timedelta(days=base_dt.weekday())
        grid = [[start_dt + timedelta(days=i) for i in range(7)]]
    else:
        grid = cal.monthdatescalendar(base_dt.year, base_dt.month)

    st.markdown("---")
    
    day_names = ["週一", "週二", "週三", "週四", "週五", "週六", "週日"]
    cols = st.columns(7)
    for i, name in enumerate(day_names):
        cols[i].markdown(f"<div style='text-align:center; font-weight:bold; color:#424242; padding-bottom:8px;'>{name}</div>", unsafe_allow_html=True)
        
    for week in grid:
        w_cols = st.columns(7)
        for i, dt in enumerate(week):
            with w_cols[i]:
                is_selected = (dt == st.session_state["dashboard_selected_date"])
                is_today = (dt == date.today())
                btn_type = "primary" if is_selected else "secondary"
                day_label = f"{dt.month}/{dt.day}" if dt.day == 1 else f"{dt.day}"
                weekday_label = day_names[dt.weekday()]
                
                # 計算當日各區域預約狀況
                mask = (pd.to_datetime(df_all["check_in_date"]).dt.date <= dt) & (pd.to_datetime(df_all["check_out_date"]).dt.date > dt)
                daily_df = df_all[mask]
                
                summary = {"雨棚區": 0, "草地區": 0, "通舖房間": 0}
                has_group = False
                has_night_rush = False
                
                for _, row in daily_df.iterrows():
                    area = row["area"]
                    qty = int(row["quantity"])
                    # 解析跨區複合區域字串 (如 "雨棚區 (2 帳)、草地區 (1 帳)")
                    # 確保各區域狀態點點 (雨棚/草地/通鋪) 都能正確累加顯示
                    parsed_areas = parse_area_quantities(str(area), qty)
                    for area_name, area_qty in parsed_areas.items():
                        if area_name in summary:
                            summary[area_name] += area_qty
                    if is_shelter_group_booking(area, qty):
                        has_group = True
                    # 檢測當日是否為夜衝日 (即入住當日，且有夜衝帳數)
                    if str(row["check_in_date"]) == dt.strftime("%Y-%m-%d"):
                        rush_qty = int(row.get("night_rush_quantity", 1 if row.get("night_rush") else 0))
                        if rush_qty > 0:
                            has_night_rush = True
                
                # 格子樣式：根據選取/今日/一般狀態決定邊框與背景色
                if is_selected:
                    border_color, bg_color = "#1565C0", "#E3F2FD"
                elif is_today:
                    border_color, bg_color = "#2E7D32", "#E8F5E9"
                else:
                    border_color, bg_color = "#E0E0E0", "#FFFFFF"

                # 區域摘要徽章 (compact font, font-weight 600)
                badges_html = ""
                if summary["雨棚區"] > 0:
                    badges_html += f'<div style="font-size:0.85rem; color:#1B5E20; font-weight:600;"><span style="display:inline-block; width:10px; height:10px; border-radius:50%; background:#2E7D32; margin-right:4px;"></span>雨棚 x{summary["雨棚區"]}</div>'
                if summary["草地區"] > 0:
                    badges_html += f'<div style="font-size:0.85rem; color:#E65100; font-weight:600;"><span style="display:inline-block; width:10px; height:10px; border-radius:50%; background:#EF6C00; margin-right:4px;"></span>草地 x{summary["草地區"]}</div>'
                if summary["通舖房間"] > 0:
                    badges_html += f'<div style="font-size:0.85rem; color:#6A1B9A; font-weight:600;"><span style="display:inline-block; width:10px; height:10px; border-radius:50%; background:#8E24AA; margin-right:4px;"></span>通鋪 x{summary["通舖房間"]}</div>'

                # 特徵膠囊標籤 (soft padding, inline pills)
                tags_html = ""
                if has_group:
                    tags_html += '<span style="background:#F57F17; color:#FFFFFF; padding:2px 6px; border-radius:10px; font-size:0.75rem; font-weight:700;">🏷️ 包區</span> '
                if has_night_rush:
                    tags_html += '<span style="background:#E65100; color:#FFFFFF; padding:2px 6px; border-radius:10px; font-size:0.75rem; font-weight:700;">⚡ 夜衝</span>'

                # 單一完整 HTML 卡片：Header + Badges + Tags 全部封裝在單一 div 內
                card_html = f"""
                <div style="border: 2px solid {border_color}; background-color: {bg_color}; border-radius: 10px; padding: 10px; min-height: 125px; box-shadow: 0 2px 4px rgba(0,0,0,0.03);">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                        <span style="font-size: 1.1rem; font-weight: 800; color: #212121;">{day_label}</span>
                        <span style="font-size: 0.8rem; color: #757575;">{weekday_label}</span>
                    </div>
                    <div style="display: flex; flex-direction: column; gap: 4px; margin-bottom: 6px;">
                        {badges_html}
                    </div>
                    <div style="display: flex; gap: 4px; flex-wrap: wrap;">
                        {tags_html}
                    </div>
                </div>
                """
                st.markdown(card_html, unsafe_allow_html=True)

                # 獨立選擇按鈕（置於卡片下方，避免 Streamlit 關閉 div 導致徽章溢出）
                if st.button("選擇", key=f"cal_btn_{dt.isoformat()}", use_container_width=True, type="primary" if is_selected else "secondary"):
                    st.session_state["dashboard_selected_date"] = dt
                    st.rerun()

    # 3. 日期點擊與側邊/下方「訂單資料抽屜 (Drawer)」
    selected_dt = st.session_state["dashboard_selected_date"]
    st.markdown("---")
    
    drawer_col1, drawer_col2 = st.columns([1, 1])
    with drawer_col1:
        st.markdown(f"#### 📝 {selected_dt.strftime('%Y-%m-%d')} 訂位明細抽屜")
    with drawer_col2:
        if st.button(f"➕ 快速新增 {selected_dt.strftime('%m/%d')} 露營預約", use_container_width=True, type="primary"):
            # 無縫跳轉至「新增露營預約」分頁，並自動帶入點選的入住日期
            st.session_state["selected_booking_date"] = selected_dt
            st.session_state["jump_to_page"] = "➕ 新增露營預約"
            st.session_state["editing_booking_id"] = None
            st.session_state["editing_booking_data"] = None
            # 重置表單日期鍵值，確保表單載入時正確帶入點選日期與隔天退房
            reset_booking_form_date_keys()
            st.rerun()
            
    mask_selected = (pd.to_datetime(df_all["check_in_date"]).dt.date <= selected_dt) & (pd.to_datetime(df_all["check_out_date"]).dt.date > selected_dt)
    df_selected = df_all[mask_selected]
    
    if df_selected.empty:
        st.info("當日無任何營位預約記錄。")
    else:
        for _, row in df_selected.iterrows():
            tag_style = TAG_BADGE_STYLE.get(row["tag"], TAG_BADGE_STYLE["一般"])
            area_unit = CAMP_AREAS.get(row["area"], {}).get("unit", "帳")
            rush_num = int(row.get("night_rush_quantity", 1 if row.get("night_rush") else 0))
            # 僅當天入住的訂單才會顯示夜衝，否則只是一般停留
            has_rush_flag = (row.get("night_rush") == 1 or rush_num > 0) and (str(row["check_in_date"]) == selected_dt.strftime("%Y-%m-%d"))
            is_group_shelter = is_shelter_group_booking(row.get("area", ""), row.get("quantity", 0))
            
            shelter_group_badge = (
                ' <span style="background:#FFF8E1; color:#F57F17; border:1px solid #FFE082; padding:2px 8px; border-radius:12px; font-weight:700; font-size:0.82rem;">🏷️ 雨棚包區</span>'
                if is_group_shelter else ""
            )
            # 各區域狀態點點標籤 (🟢 雨棚 / 🟠 草地 / 🟣 通鋪)，與包區標籤同時並存顯示
            card_area_qtys = parse_area_quantities(str(row.get("area", "")), int(row.get("quantity", 0)))
            canopy_qty = card_area_qtys.get("雨棚區", 0)
            grass_qty = card_area_qtys.get("草地區", 0)
            room_qty = card_area_qtys.get("通舖房間", 0)
            area_dot_badges = ""
            if canopy_qty > 0:
                area_dot_badges += f' <span style="background:#E8F5E9; color:#1B5E20; border:1px solid #A5D6A7; padding:2px 8px; border-radius:12px; font-weight:700; font-size:0.82rem;">🟢 雨棚 x{canopy_qty}</span>'
            if grass_qty > 0:
                area_dot_badges += f' <span style="background:#FFF3E0; color:#E65100; border:1px solid #FFCC80; padding:2px 8px; border-radius:12px; font-weight:700; font-size:0.82rem;">🟠 草地 x{grass_qty}</span>'
            if room_qty > 0:
                area_dot_badges += f' <span style="background:#F3E5F5; color:#6A1B9A; border:1px solid #CE93D8; padding:2px 8px; border-radius:12px; font-weight:700; font-size:0.82rem;">🟣 通鋪 x{room_qty}</span>'
            
            card_html = (
                f"""
                <div class="luye-card" style="{'border-left: 5px solid #FBC02D;' if is_group_shelter else ''} margin-bottom:12px;">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <div>
                            <span style="font-size:1.2rem; font-weight:700; color:#1B5E20;">{row['name']}</span>
                            <span style="font-size:0.95rem; color:#666; margin-left:12px;">📞 {row['phone']}</span>
                            <span style="{tag_style}; margin-left:10px;">{row['tag']}</span>
                            {area_dot_badges}
                            {shelter_group_badge}
                        </div>
                        <div>
                            <span style="font-weight:700; font-size:1.25rem; color:#E65100;">NT$ {int(row['total_amount']):,}</span>
                        </div>
                    </div>
                    <div style="margin-top:10px; font-size:0.92rem; color:#333; display:flex; flex-wrap:wrap; gap:16px;">
                        <div>🏕️ <b>區域</b>：{row['area']} {'(滿 6 帳包區)' if is_group_shelter else ''}</div>
                        <div>📅 <b>入住~退房</b>：{row['check_in_date']} 至 {row['check_out_date']}</div>
                        <div>⚡ <b>夜衝</b>：{f'是 ({rush_num} {area_unit})' if has_rush_flag else '否'}</div>
                        <div>🔖 <b>訂單號</b>：<code>{row['booking_code']}</code></div>
                    </div>
                    {f'<div style="margin-top:8px; padding:6px 10px; background:#FFF9C4; border-radius:6px; font-size:0.88rem; color:#5D4037;">📝 備註：{row["notes"]}</div>' if row["notes"] else ''}
                </div>
                """
            )
            
            if hasattr(st, "html"):
                st.html(card_html)
            else:
                st.markdown(card_html, unsafe_allow_html=True)

            # 卡片內嵌編輯/刪除控件 (session_state 驅動，
            # 表單與確認直接渲染於卡片下方)
            render_card_edit_delete_controls(row)


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

    # 使用快取函式 (600秒快取)
    df_results = fetch_bookings_df(query, tuple(params))

    st.markdown(f"**找到 `{len(df_results)}` 筆符合條件的訂單**")

    # 表格或卡片切換
    view_tab1, view_tab2 = st.tabs(["📑 詳細清單表格", "🗂️ 訂單卡片檢視"])

    with view_tab1:
        if not df_results.empty:
            display_df = df_results.copy()
            display_df["area"] = display_df.apply(
                lambda r: f"{r['area']} (🏷️包區)" if is_shelter_group_booking(r.get("area", ""), r.get("quantity", 0)) else r["area"],
                axis=1,
            )
            display_df["night_rush"] = display_df.apply(
                lambda r: f"⚡ {int(r.get('night_rush_quantity', 1))} {CAMP_AREAS.get(r.get('area'), {}).get('unit', '帳')}"
                if (r.get("night_rush") == 1 or (r.get("night_rush_quantity") is not None and r.get("night_rush_quantity") > 0))
                else "否",
                axis=1,
            )
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
                    "night_rush": "夜衝數量",
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
                "夜衝數量",
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
                rush_num = int(row.get("night_rush_quantity", 1 if row.get("night_rush") else 0))
                has_rush_flag = (row.get("night_rush") == 1 or rush_num > 0)
                is_group_shelter = is_shelter_group_booking(row.get("area", ""), row.get("quantity", 0))
                shelter_group_badge = (
                    ' <span style="background:#FFF8E1; color:#F57F17; border:1px solid #FFE082; padding:2px 8px; border-radius:12px; font-weight:700; font-size:0.82rem;">🏷️ 雨棚包區</span>'
                    if is_group_shelter
                    else ""
                )

                card_box = st.container()
                with card_box:
                    card_html = (
                        f"""
                        <div class="luye-card" style="{'border-left: 5px solid #FBC02D;' if is_group_shelter else ''}">
                            <div style="display:flex; justify-content:space-between; align-items:center;">
                                <div>
                                    <span style="font-size:1.2rem; font-weight:700; color:#1B5E20;">{row['name']}</span>
                                    <span style="font-size:0.95rem; color:#666; margin-left:12px;">📞 {row['phone']}</span>
                                    <span style="{tag_style}; margin-left:10px;">{row['tag']}</span>
                                    {shelter_group_badge}
                                    <span style="background:#E0F2F1; color:#00695C; padding:2px 8px; border-radius:12px; font-size:0.8rem; margin-left:6px;">第 {row['visit_count']} 次到訪</span>
                                </div>
                                <div>
                                    <span style="font-weight:700; font-size:1.25rem; color:#E65100;">NT$ {row['total_amount']:,}</span>
                                </div>
                            </div>
                            <div style="margin-top:10px; font-size:0.92rem; color:#333; display:flex; flex-wrap:wrap; gap:16px;">
                                <div>🏕️ <b>區域</b>：{row['area']} {'(滿 6 帳包區)' if is_group_shelter else ''}</div>
                                <div>📅 <b>入住~退房</b>：{row['check_in_date']} 至 {row['check_out_date']}</div>
                                <div>⚡ <b>夜衝</b>：{f'是 ({rush_num} {area_unit})' if has_rush_flag else '否'}</div>
                                <div>🔖 <b>訂單號</b>：<code>{row['booking_code']}</code></div>
                                <div>👤 <b>經手人</b>：{row['created_by']}</div>
                            </div>
                            {f'<div style="margin-top:8px; padding:6px 10px; background:#FFF9C4; border-radius:6px; font-size:0.88rem; color:#5D4037;">📝 備註：{row["notes"]}</div>' if row["notes"] else ''}
                        </div>
                        """
                    )
                    if hasattr(st, "html"):
                        st.html(card_html)
                    else:
                        st.markdown(card_html, unsafe_allow_html=True)

                    # 卡片內嵌編輯/刪除控件 (session_state 驅動，
                    # 表單與確認直接渲染於卡片下方)
                    render_card_edit_delete_controls(row)

        else:
            st.info("尚無符合搜尋條件的訂單記錄。")


# ==============================================================================
# 9. 頁面功能：新增與編輯預約 (Booking Form)
# ==============================================================================
# 表單入住/退房日期連動所使用的 session_state 鍵值
FORM_CHECK_IN_KEY = "booking_form_check_in_date"
FORM_CHECK_OUT_KEY = "booking_form_check_out_date"
FORM_PREV_CHECK_IN_KEY = "booking_form_prev_check_in"


def reset_booking_form_date_keys():
    """重置表單日期鍵值，讓表單下次渲染時正確帶入預設的入住/退房日期"""
    for _dk in (FORM_CHECK_IN_KEY, FORM_CHECK_OUT_KEY, FORM_PREV_CHECK_IN_KEY):
        st.session_state.pop(_dk, None)


def compute_linked_checkout(
    check_in_val: date,
    default_check_out: date,
    prev_check_in: Optional[date],
    current_check_out: Optional[date],
) -> Tuple[date, date]:
    """純函式：計算預約表單入住/退房日期自動連動

    入住日期選取時自動計算 default_check_out = check_in + timedelta(days=1)，
    退房日期選擇器的 value 與 min_value 皆動態對應該值，
    確保最低住宿一晚，並防止選取同日或早於入住日的日期。

    參數：
    - check_in_val: 目前入住日期
    - default_check_out: 表單預設退房日 (編輯模式為原訂單退房日；新增模式為入住日 +1 天)
    - prev_check_in: 前次渲染的入住日期 (用於偵測變更)
    - current_check_out: 目前 session_state 中的退房日期

    回傳 (min_checkout, linked_check_out)：
    - min_checkout: 退房日期最小合規值 (= check_in + 1 天)
    - linked_check_out: 應寫入 session_state 的退房日期
    """
    # 退房日期最小合規值 (入住日 +1 天，兩天一夜預設)
    min_checkout = check_in_val + timedelta(days=1)

    # 入住日期變更同步機制：入住日期改變時，強制更新退房日期為新入住日 +1 天
    # (註：Streamlit 規範禁止 st.form 內的 widget 使用 on_change 回調，
    #  否則會拋出 StreamlitInvalidFormCallbackError，故以此同步機制實現相同效果)
    if prev_check_in is not None and check_in_val != prev_check_in:
        return min_checkout, min_checkout

    # 退房日期 Session State 初始化：首次渲染時以表單預設值建立鍵值
    # (編輯模式為原訂單退房日，新增/快速新增模式為入住日 +1 天)
    if current_check_out is None:
        current_check_out = default_check_out

    # 防衛邏輯：若退房日期小於 min_checkout (或小於等於入住日期)，強制更新為 min_checkout
    # 確保 st.date_input 的 value 絕對不小於 min_value，徹底杜絕 StreamlitValueBelowMinError
    if current_check_out < min_checkout or current_check_out <= check_in_val:
        current_check_out = min_checkout

    return min_checkout, current_check_out


def render_booking_form_view(is_quick_mode: bool = False):
    editing_id = st.session_state.get("editing_booking_id", None)
    is_editing = editing_id is not None

    booking_data = None
    if is_editing:
        if "editing_booking_data" in st.session_state and st.session_state["editing_booking_data"]:
            booking_data = st.session_state["editing_booking_data"]
        else:
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
                st.session_state["editing_booking_data"] = None
                reset_booking_form_date_keys()
                is_editing = False

    if not is_quick_mode:
        st.markdown(f"### {'✏️ 編輯露營訂位資料' if is_editing else '➕ 新增露營客預約訂單'}")

    if is_editing and booking_data:
        st.info(f"正在編輯訂單：**{booking_data['booking_code']}** (露營客: {booking_data['name']})")
        if st.button("⬅️ 取消編輯，切換回新增預約模式"):
            st.session_state["editing_booking_id"] = None
            st.session_state["editing_booking_data"] = None
            reset_booking_form_date_keys()
            st.rerun()

    # 預設值設定
    default_name = booking_data["name"] if is_editing else ""
    default_phone = booking_data["phone"] if is_editing else ""
    default_area = booking_data["area"] if is_editing else ""
    default_tag = booking_data["tag"] if is_editing else "一般"
    default_notes = booking_data["notes"] if is_editing else ""
    default_visit_count = booking_data["visit_count"] if is_editing else 1

    if is_editing:
        # 安全解析入住/退房日期：批次匯入的歷史記錄可能為
        # 「歷史客戶/無預約日期」等占位字串，解析失敗時回退至今天，
        # 避免編輯表單因 ValueError 無法開啟
        try:
            default_check_in = datetime.strptime(str(booking_data["check_in_date"]), "%Y-%m-%d").date()
        except (TypeError, ValueError):
            default_check_in = date.today()
        try:
            default_check_out = datetime.strptime(str(booking_data["check_out_date"]), "%Y-%m-%d").date()
        except (TypeError, ValueError):
            default_check_out = date.today() + timedelta(days=1)
    else:
        # 快速預約：自動帶入點選日期 (優先 selected_booking_date，相容 quick_booking_date / selected_date)
        quick_date = (
            st.session_state.get("selected_booking_date")
            or st.session_state.get("quick_booking_date")
            or st.session_state.get("selected_date")
        )
        if quick_date:
            default_check_in = quick_date
            st.session_state["selected_booking_date"] = None
            st.session_state["quick_booking_date"] = None
            if "selected_date" in st.session_state:
                st.session_state["selected_date"] = None
        else:
            default_check_in = date.today() + timedelta(days=1)
        default_check_out = default_check_in + timedelta(days=1)

    # 表單 widget 唯一鍵值（綁定 session_state，確保編輯切換時正確帶入預填值）
    FORM_NAME_KEY = "booking_form_name"
    FORM_PHONE_KEY = "booking_form_phone"
    FORM_TAG_KEY = "booking_form_tag"
    FORM_VISIT_KEY = "booking_form_visit"
    FORM_NOTES_KEY = "booking_form_notes"
    FORM_QTY_SHELTER_KEY = "booking_form_qty_shelter"
    FORM_RUSH_SHELTER_KEY = "booking_form_rush_shelter"
    FORM_QTY_GRASS_KEY = "booking_form_qty_grass"
    FORM_RUSH_GRASS_KEY = "booking_form_rush_grass"
    FORM_QTY_ROOM_KEY = "booking_form_qty_room"
    FORM_RUSH_ROOM_KEY = "booking_form_rush_room"

    # 跨區數量與各區獨立夜衝數量初始化（從 form 內移出，供種子化使用）
    if is_quick_mode:
        # 快速預約模式一律歸零
        canopy_qty = 0
        grass_qty = 0
        room_qty = 0
        canopy_night_qty = 0
        grass_night_qty = 0
        room_night_qty = 0
    elif is_editing and booking_data:
        # 編輯模式：精準還原原訂單各區域數量 (禁止直接歸零)
        # 舊格式無括號字串時，以訂單總數量作為該區域數量
        fallback_total = int(booking_data.get("quantity", 0) or 0)
        parsed_defaults = parse_area_quantities(default_area, fallback_total)
        canopy_qty = parsed_defaults.get("雨棚區", 0)
        grass_qty = parsed_defaults.get("草地區", 0)
        room_qty = parsed_defaults.get("通舖房間", 0)
        # 各區獨立夜衝數量：優先解析 night_rush_detail 明細
        rush_detail_str = str(booking_data.get("night_rush_detail") or "")
        if rush_detail_str:
            parsed_rush = parse_area_quantities(rush_detail_str, 0)
            canopy_night_qty = parsed_rush.get("雨棚區", 0)
            grass_night_qty = parsed_rush.get("草地區", 0)
            room_night_qty = parsed_rush.get("通舖房間", 0)
        else:
            # 舊資料相容：無明細時將夜衝總數配給首個有預約的區域
            canopy_night_qty = 0
            grass_night_qty = 0
            room_night_qty = 0
            old_rush_qty = int(booking_data.get("night_rush_quantity", 0) or 0)
            if old_rush_qty > 0:
                if canopy_qty > 0:
                    canopy_night_qty = min(canopy_qty, old_rush_qty)
                elif grass_qty > 0:
                    grass_night_qty = min(grass_qty, old_rush_qty)
                elif room_qty > 0:
                    room_night_qty = min(room_qty, old_rush_qty)
    else:
        # 一般新增模式：全部歸零
        canopy_qty = 0
        grass_qty = 0
        room_qty = 0
        canopy_night_qty = 0
        grass_night_qty = 0
        room_night_qty = 0

    # 偵測編輯內容切換：切換編輯對象（或進入/退出編輯模式）時，
    # 以 booking_data 種子化 widget 狀態，確保預填值生效且使用者修改可被捕捉
    _form_ctx = f"{'quick' if is_quick_mode else 'normal'}_{editing_id or 'new'}"
    if st.session_state.get("booking_form_ctx") != _form_ctx:
        st.session_state["booking_form_ctx"] = _form_ctx
        st.session_state[FORM_NAME_KEY] = default_name
        st.session_state[FORM_PHONE_KEY] = default_phone
        if default_tag in CUSTOMER_TAGS:
            st.session_state[FORM_TAG_KEY] = default_tag
        st.session_state[FORM_VISIT_KEY] = int(default_visit_count)
        st.session_state[FORM_NOTES_KEY] = default_notes
        st.session_state[FORM_QTY_SHELTER_KEY] = int(min(8, canopy_qty))
        st.session_state[FORM_RUSH_SHELTER_KEY] = int(min(min(8, canopy_qty), canopy_night_qty))
        st.session_state[FORM_QTY_GRASS_KEY] = int(min(4, grass_qty))
        st.session_state[FORM_RUSH_GRASS_KEY] = int(min(min(4, grass_qty), grass_night_qty))
        st.session_state[FORM_QTY_ROOM_KEY] = int(min(2, room_qty))
        st.session_state[FORM_RUSH_ROOM_KEY] = int(min(min(2, room_qty), room_night_qty))
        st.session_state[FORM_CHECK_IN_KEY] = default_check_in
        st.session_state[FORM_CHECK_OUT_KEY] = default_check_out
        st.session_state[FORM_PREV_CHECK_IN_KEY] = default_check_in

    with st.form("booking_editor_form"):
        st.markdown("#### 1. 露營客基本資料")
        col_c1, col_c2, col_c3 = st.columns([1.2, 1.2, 1])

        with col_c1:
            name_val = st.text_input("露營人姓名 *", key=FORM_NAME_KEY, placeholder="例如：王小明")
        with col_c2:
            phone_val = st.text_input("露營人電話 *", key=FORM_PHONE_KEY, placeholder="例如：0912345678")
        with col_c3:
            tag_val = st.selectbox(
                "客戶標籤 *",
                options=CUSTOMER_TAGS,
                key=FORM_TAG_KEY,
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
                key=FORM_VISIT_KEY,
                help="輸入電話時若已存在歷史訂單，系統會自動參考歷史到訪紀錄。",
            )

        st.markdown("---")
        st.markdown("#### 2. 營位預訂細節與日期 (支援跨區複選)")

        col_b1, col_b2, col_b3 = st.columns(3)

        with col_b1:
            qty_shelter = st.number_input("雨棚區 (帳數)", min_value=0, max_value=8, key=FORM_QTY_SHELTER_KEY)
            rush_shelter = st.number_input("⚡ 雨棚區夜衝 (帳數)", min_value=0, max_value=8, key=FORM_RUSH_SHELTER_KEY)
            if rush_shelter > qty_shelter:
                st.warning("雨棚區夜衝數量不可大於預約數量！")
        with col_b2:
            qty_grass = st.number_input("草地區 (帳數)", min_value=0, max_value=4, key=FORM_QTY_GRASS_KEY)
            rush_grass = st.number_input("⚡ 草地區夜衝 (帳數)", min_value=0, max_value=4, key=FORM_RUSH_GRASS_KEY)
            if rush_grass > qty_grass:
                st.warning("草地區夜衝數量不可大於預約數量！")
        with col_b3:
            qty_room = st.number_input("通舖房間 (間數)", min_value=0, max_value=2, key=FORM_QTY_ROOM_KEY)
            rush_room = st.number_input("⚡ 通舖夜衝 (間數)", min_value=0, max_value=2, key=FORM_RUSH_ROOM_KEY)
            if rush_room > qty_room:
                st.warning("通舖房間夜衝數量不可大於預約數量！")

        col_d1, col_d2 = st.columns(2)

        with col_d1:
            check_in_val = st.date_input(
                "主要入住日期 *",
                value=default_check_in,
                key=FORM_CHECK_IN_KEY,
            )

        # 退房日期自動連動：純函式計算，防止 StreamlitValueBelowMinError
        # 動態計算 min_checkout = check_in + 1 天
        # 若目前退房日期 <= 入住日期，自動強制設為 min_checkout
        min_checkout = check_in_val + timedelta(days=1)
        current_checkout = st.session_state.get(FORM_CHECK_OUT_KEY, default_check_out)
        if current_checkout <= check_in_val:
            current_checkout = min_checkout
        st.session_state[FORM_CHECK_OUT_KEY] = current_checkout

        with col_d2:
            check_out_val = st.date_input(
                "預計退房日期 *",
                value=current_checkout,
                min_value=min_checkout,
                key=FORM_CHECK_OUT_KEY,
            )

        total_qty = qty_shelter + qty_grass + qty_room
        area_parts = []
        if qty_shelter > 0: area_parts.append(f"雨棚區 ({qty_shelter} 帳)")
        if qty_grass > 0: area_parts.append(f"草地區 ({qty_grass} 帳)")
        if qty_room > 0: area_parts.append(f"通舖房間 ({qty_room} 間)")
        area_val = "、".join(area_parts) if area_parts else "未選擇"

        # 各區獨立夜衝數量 (不超過該區域預約數量) 與夜衝明細字串
        eff_rush_shelter = min(qty_shelter, rush_shelter)
        eff_rush_grass = min(qty_grass, rush_grass)
        eff_rush_room = min(qty_room, rush_room)
        total_rush_qty = eff_rush_shelter + eff_rush_grass + eff_rush_room
        rush_parts = []
        if eff_rush_shelter > 0: rush_parts.append(f"雨棚區 ({eff_rush_shelter} 帳)")
        if eff_rush_grass > 0: rush_parts.append(f"草地區 ({eff_rush_grass} 帳)")
        if eff_rush_room > 0: rush_parts.append(f"通舖房間 ({eff_rush_room} 間)")
        rush_val = "、".join(rush_parts) if rush_parts else ""
        has_night_rush = total_rush_qty > 0
        
        if qty_shelter >= 6:
            st.info("🏷️ **雨棚區包區提示**：雨棚預約達 6 帳以上，系統自動認定為【雨棚區包區】，營位總覽看板將連動保留雨棚區 1~8 號全區！")

        notes_val = st.text_area("空白文字備註 (特殊需求、車牌、同團成員等)", key=FORM_NOTES_KEY, placeholder="請在此自由填寫備註...")

        # 即時費用計算 (傳入各區獨立夜衝明細)
        calc_total, nights_count, base_fee, rush_fee = calculate_booking_fee(
            area_val, total_qty, check_in_val, check_out_val, rush_val
        )

        st.markdown("---")
        st.markdown("#### 3. 預估費用計算明細")
        f_col1, f_col2, f_col3, f_col4 = st.columns(4)
        f_col1.metric("預計住宿晚數", f"{nights_count} 晚")
        f_col2.metric("營位基礎費用", f"NT$ {base_fee:,}", f"共 {total_qty} 單位")
        f_col3.metric("夜衝費用 (50%)", f"NT$ {rush_fee:,}", "包含夜衝" if has_night_rush else "無夜衝")
        f_col4.metric("應收總金額", f"NT$ {calc_total:,}", delta="含跨區組合" if len(area_parts) > 1 else "")

        submit_btn = st.form_submit_button(
            "💾 儲存並建立訂單" if not is_editing else "💾 儲存修改內容",
            type="primary",
            use_container_width=True,
        )

        if submit_btn:
            # 1. Input Validations
            if not name_val.strip():
                st.error("❌ 請填寫露營人姓名！")
            elif not phone_val.strip():
                st.error("❌ 請填寫露營人電話！")
            elif total_qty == 0:
                st.error("❌ 請至少選擇一個區域並輸入數量！")
            elif check_out_val <= check_in_val:
                st.error("❌ 退房日期必須在入住日期之後！")
            else:
                current_user = st.session_state.get("authenticated_user", "管理員")
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                has_night_rush_flag = 1 if has_night_rush else 0

                try:
                    conn = get_db_connection()
                    cursor = conn.cursor()

                    if is_editing:
                        # 2A. Clear & Explicit Update
                        cursor.execute(
                            """
                            UPDATE bookings
                            SET check_in_date = ?, check_out_date = ?, name = ?, phone = ?,
                                area = ?, quantity = ?, night_rush = ?, night_rush_quantity = ?, night_rush_detail = ?,
                                visit_count = ?, tag = ?, notes = ?, total_amount = ?, updated_at = ?
                            WHERE id = ?
                            """,
                            (
                                check_in_val.strftime("%Y-%m-%d"),
                                check_out_val.strftime("%Y-%m-%d"),
                                name_val.strip(),
                                phone_val.strip(),
                                area_val,
                                total_qty,
                                has_night_rush_flag,
                                int(total_rush_qty),
                                rush_val,
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
                            f"修改訂單 ID {editing_id} (客戶: {name_val}, 區域: {area_val}, 總金額: {calc_total})",
                        )
                    else:
                        # 2B. Insert New Order
                        b_code = generate_booking_code(check_in_val)
                        cursor.execute(
                            """
                            INSERT INTO bookings (
                                booking_code, check_in_date, check_out_date, name, phone,
                                area, quantity, night_rush, night_rush_quantity, night_rush_detail, visit_count, tag, notes,
                                total_amount, created_by, created_at, updated_at
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                b_code,
                                check_in_val.strftime("%Y-%m-%d"),
                                check_out_val.strftime("%Y-%m-%d"),
                                name_val.strip(),
                                phone_val.strip(),
                                area_val,
                                total_qty,
                                has_night_rush_flag,
                                int(total_rush_qty),
                                rush_val,
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
                            f"新增訂單 {b_code} (客戶: {name_val}, 電話: {phone_val}, 金額: {calc_total})",
                        )

                    conn.close()

                    # 3. Cache & Sync Refinement
                    st.cache_data.clear()
                    refresh_session_state_data()
                    sync_db_to_google_sheet()
                    trigger_auto_sync(current_user)

                    # 4. Clean State & UI Feedback
                    st.session_state["editing_booking_id"] = None
                    st.session_state["editing_booking_data"] = None
                    reset_booking_form_date_keys()

                    st.toast("✅ 訂單已成功儲存並同步至 Google Sheets！", icon="🎉")
                    st.rerun()

                except Exception as e:
                    st.error(f"⚠️ 儲存資料庫時發生嚴重錯誤: {str(e)}")
                    print(f"[Form Submit Error] {str(e)}")


# ==============================================================================
# 10. 頁面功能：Excel/CSV 匯入與匯出 (Import / Export)
# ==============================================================================
def match_import_notes_column(columns) -> Optional[str]:
    """批次匯入檔案的備註欄位靈活配對

    依優先順序配對：備註 > 備註說明 > 空白文字備註 >
    notes > remarks > 說明，確保不同來源的 XLSX/CSV
    備註欄位皆能正確讀入。
    """
    return next(
        (c for c in ["備註", "備註說明", "空白文字備註", "notes", "remarks", "說明"] if c in columns),
        None,
    )


def parse_import_notes(row: Any, notes_col: Optional[str]) -> str:
    """解析批次匯入列的備註內容

    依配對到的備註欄位讀取，NaN/None 或空白儲存格安全轉為
    空字串；並合併「地區/注意/車宿/寵物/其他」等補充欄位，
    以「 | 」連接。
    """
    # 備註主欄位解析 (NaN/空白安全轉為空字串)
    base_notes = ""
    if notes_col:
        raw_notes = row.get(notes_col)
        if raw_notes is not None and not pd.isna(raw_notes):
            parsed_notes = str(raw_notes).strip()
            if parsed_notes and parsed_notes.lower() not in ["nan", "none", "null"]:
                base_notes = parsed_notes

    # 備註整合邏輯 (合併 "地區", "注意", "車宿", "寵物", "其他")
    notes_parts = []
    if base_notes:
        notes_parts.append(base_notes)
    for col_name in ["地區", "注意", "車宿", "寵物", "其他"]:
        val = str(row.get(col_name, "")).strip()
        if val and val.lower() not in ["nan", "none", "null"]:
            notes_parts.append(f"{col_name}: {val}")

    return " | ".join(notes_parts)


def render_import_export_view():
    st.markdown("### 📤 Excel / CSV 批次資料匯入與匯出")

    tab_export, tab_import = st.tabs(["📥 一鍵資料匯出 (Export)", "📤 批次上傳匯入 (Import)"])

    # 匯出資料
    with tab_export:
        st.markdown("#### 1. 匯出目前營地所有露營客訂單資料")
        st.write("點擊下方按鈕可將系統中完整的露營客訂位紀錄、金額與備註匯出為 Excel 或 CSV 試算表。")

        # 使用快取全量查詢 (600秒快取)
        df_all = fetch_all_bookings_df()

        if not df_all.empty:
            df_export = df_all.copy()
            df_export["night_rush"] = df_export.apply(lambda r: "是" if r.get("night_rush") or r.get("night_rush_quantity", 0) > 0 else "否", axis=1)
            df_export["night_rush_quantity"] = df_export["night_rush_quantity"].fillna(0).astype(int)
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
                    "night_rush_quantity": "夜衝數量",
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
                    "夜衝數量": 1,
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
                    "夜衝數量": 0,
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
                    overwrite_count = 0
                    error_logs = []
                    current_user = st.session_state.get("authenticated_user", "管理員")
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    # 備註欄位靈活配對 (詳見 match_import_notes_column)
                    notes_col = match_import_notes_column(df_upload.columns)

                    conn = get_db_connection()
                    cursor = conn.cursor()

                    # 最新覆寫前置檢查：載入 SQLite 既有訂單的複合鍵
                    # (入住日期, 姓名, 電話) 對應 id 集合 (標準化後比對)
                    existing_key_ids = {}
                    for _r in cursor.execute(
                        "SELECT id, check_in_date, name, phone FROM bookings"
                    ).fetchall():
                        _key = (
                            _normalize_date_str(_r[1]),
                            _normalize_text(_r[2]),
                            _normalize_phone(_r[3]),
                        )
                        existing_key_ids.setdefault(_key, []).append(_r[0])

                    for idx, row in df_upload.iterrows():
                        try:
                            # 姓名與電話解析
                            name = str(row.get("露營人姓名", row.get("name", row.get("姓名", "")))).strip()
                            if name.lower() in ["nan", "none", "null"]:
                                name = ""

                            raw_phone = row.get(phone_col) if phone_col else row.get("露營人電話", row.get("phone", row.get("電話", "")))
                            phone = format_phone_number(raw_phone)

                            # 預約區域解析
                            area = str(row.get("預約區域", row.get("area", "雨棚區"))).strip()
                            if area not in CAMP_AREAS:
                                area = "雨棚區"

                            # 日期安全解析 (確保絕不丟出 ValueError，允許無日期)
                            in_raw = str(row.get(in_col, "") if in_col else row.get("主要入住日期", row.get("check_in_date", ""))).strip()
                            check_in_str = "歷史客戶/無預約日期"
                            check_out_str = "歷史客戶/無預約日期"
                            is_valid_date = False
                            
                            if in_raw and in_raw.lower() not in ["nan", "none", "null", "nat"]:
                                try:
                                    check_in_dt = datetime.strptime(in_raw[:10], "%Y-%m-%d").date()
                                    check_in_str = check_in_dt.strftime("%Y-%m-%d")
                                    is_valid_date = True
                                except Exception:
                                    pass

                            out_raw = str(row.get(out_col, "") if out_col else row.get("退房日期", row.get("check_out_date", ""))).strip()
                            if is_valid_date:
                                if out_raw and out_raw.lower() not in ["nan", "none", "null", "nat"]:
                                    try:
                                        check_out_dt = datetime.strptime(out_raw[:10], "%Y-%m-%d").date()
                                        if check_out_dt <= check_in_dt:
                                            check_out_dt = check_in_dt + timedelta(days=1)
                                        check_out_str = check_out_dt.strftime("%Y-%m-%d")
                                    except Exception:
                                        check_out_dt = check_in_dt + timedelta(days=1)
                                        check_out_str = check_out_dt.strftime("%Y-%m-%d")
                                else:
                                    check_out_dt = check_in_dt + timedelta(days=1)
                                    check_out_str = check_out_dt.strftime("%Y-%m-%d")

                            # 數量解析
                            raw_qty = row.get("帳數/間數", row.get("quantity", 1))
                            try:
                                qty = int(float(raw_qty)) if not pd.isna(raw_qty) else 1
                                qty = max(1, qty)
                            except Exception:
                                qty = 1

                            # 夜衝數量解析 (彈性支援獨立夜衝帳數或原是否夜衝欄位)
                            raw_rush_qty = row.get("夜衝數量", row.get("夜衝帳數/間數", row.get("night_rush_quantity", None)))
                            if raw_rush_qty is not None and not pd.isna(raw_rush_qty):
                                try:
                                    rush_qty = int(float(raw_rush_qty))
                                except Exception:
                                    rush_qty = 0
                            else:
                                rush_val = str(row.get("是否夜衝", row.get("night_rush", "否"))).strip().lower()
                                has_rush_bool = rush_val in ["是", "true", "1", "yes", "y"]
                                rush_qty = qty if has_rush_bool else 0

                            rush_qty = max(0, min(qty, rush_qty))
                            night_rush = 1 if rush_qty > 0 else 0

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

                            # 備註解析 (NaN/空白安全處理，詳見 parse_import_notes)
                            notes = parse_import_notes(row, notes_col)

                            if not name or not phone:
                                skip_count += 1
                                error_logs.append(f"第 {idx+1} 列：缺少姓名或電話，略過。")
                                continue

                            # 最新覆寫機制：本行代表最新更新 (區域/備註/
                            # 客戶資料等)，複合鍵相同時先刪除既有舊記錄
                            # (依 id 刪除，容忍電話/姓名格式差異) 再插入
                            import_key = (
                                _normalize_date_str(check_in_str),
                                _normalize_text(name),
                                _normalize_phone(phone),
                            )
                            if import_key in existing_key_ids:
                                for _old_id in existing_key_ids[import_key]:
                                    cursor.execute(
                                        "DELETE FROM bookings WHERE id = ?",
                                        (_old_id,),
                                    )
                                overwrite_count += 1
                            existing_key_ids[import_key] = []

                            # 自動計算費用與訂單編號 (夜衝帳數獨立計費)
                            if is_valid_date:
                                calc_total, _, _, _ = calculate_booking_fee(area, qty, check_in_dt, check_out_dt, rush_qty)
                                b_code = generate_booking_code(check_in_dt) + f"-{idx}"
                            else:
                                calc_total, _, _, _ = calculate_booking_fee(area, qty, date.today(), date.today() + timedelta(days=1), rush_qty)
                                b_code = f"LY-HISTORY-{datetime.now().strftime('%Y%m%d%H%M%S')}-{idx}"

                            cursor.execute(
                                """
                                INSERT INTO bookings (
                                    booking_code, check_in_date, check_out_date, name, phone,
                                    area, quantity, night_rush, night_rush_quantity, visit_count, tag, notes,
                                    total_amount, created_by, created_at, updated_at
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                """,
                                (
                                    b_code,
                                    check_in_str,
                                    check_out_str,
                                    name,
                                    phone,
                                    area,
                                    qty,
                                    night_rush,
                                    rush_qty,
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
                            # 記錄新插入 id，供後續批次內覆寫比對
                            existing_key_ids[import_key].append(cursor.lastrowid)
                        except Exception as e_row:
                            skip_count += 1
                            error_logs.append(f"第 {idx+1} 列匯入失敗: {e_row}")

                    conn.commit()
                    conn.close()

                    st.cache_data.clear()
                    # 立即重整 session_state 資料集，顯示最新匯入/覆寫結果
                    refresh_session_state_data()
                    # SQLite 寫入成功後，即時同步至 Google Sheets
                    sync_db_to_google_sheet()
                    add_audit_log(
                        current_user,
                        "批次匯入訂單",
                        f"自檔案 {uploaded_file.name} 匯入，成功 {success_count} 筆"
                        + (f"，覆寫既有 {overwrite_count} 筆" if overwrite_count else "")
                        + f"，略過/失敗 {skip_count} 筆",
                    )
                    trigger_auto_sync(current_user)

                    success_msg = f"匯入作業完成！成功匯入 {success_count} 筆訂單"
                    if overwrite_count:
                        success_msg += f"（其中 {overwrite_count} 筆為覆寫既有訂單的最新版本）"
                    st.success(success_msg)
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
            is_current = (name == current_user)
            with st.container():
                online_badge = (
                    '<span style="font-size:0.75rem; background:#C8E6C9; color:#1B5E20; padding:2px 8px; border-radius:10px; margin-left:8px; font-weight:600;">目前在線</span>'
                    if is_current else ""
                )
                admin_card_html = (
                    f'<div style="background:#FFFFFF; border:1px solid #E0E0E0; border-radius:10px; padding:14px 16px; margin-bottom:8px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">'
                    f'<div style="display:flex; justify-content:space-between; align-items:center;">'
                    f'<div>'
                    f'<div style="display:flex; align-items:center;">'
                    f'<span style="font-weight:700; font-size:1.1rem; color:#2E7D32;">{name}</span>'
                    f'{online_badge}'
                    f'</div>'
                    f'<div style="color:#757575; font-size:0.85rem; margin-top:4px;">PIN 碼：<code style="background:#F5F5F5; padding:2px 6px; border-radius:4px; font-weight:600;">****{pin[-2:]}</code></div>'
                    f'</div>'
                    f'</div>'
                    f'</div>'
                )
                if hasattr(st, "html"):
                    st.html(admin_card_html)
                else:
                    st.markdown(admin_card_html, unsafe_allow_html=True)

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
    st.markdown("### ☁️ 雲端同步與完整備份中心 (Google Sheets)")
    st.write("設定與檢視雲端備份狀態。每次資料有新增、修改或刪除時，系統均會觸發自動備份機制 (完整資料庫轉譯至 Google Sheets DB_Full_Backup 分頁)。")

    sync_tool = get_gdrive_sync()
    current_user = st.session_state.get("authenticated_user", "管理員")

    sa_email = get_service_account_email()
    sa_email_display = f"<code>{sa_email}</code>" if sa_email else "<code>未偵測到 (可於右側上傳 service_account.json)</code>"

    # 偵測 Google Sheets 雲端備份連線狀態
    try:
        get_gspread_client()
        sheets_status_html = '<span style="color:#2E7D32; font-weight:700;">🟢 已連線正常</span>'
        sheets_detail = f"目標試算表 ID：<code>{get_spreadsheet_id()}</code>"
    except Exception as e:
        sheets_status_html = '<span style="color:#E65100; font-weight:700;">🟡 未連線 (僅本地備份模式)</span>'
        sheets_detail = f"連線說明：{e}"

    st.markdown(
        f"""
        <div class="luye-card">
            <h4>☁️ 雲端備份對接資訊 (Google Sheets)</h4>
            <ul>
                <li><b>備份目標</b>：Google Sheets 試算表的 <code>{DB_FULL_BACKUP_WORKSHEET_TITLE}</code> 分頁 (完整資料庫轉譯備份)</li>
                <li><b>雲端連線狀態</b>：{sheets_status_html}</li>
                <li><b>服務帳號信箱</b>：{sa_email_display} <span style="color:#5F6368;">(請確認目標試算表已分享編輯權限給此信箱)</span></li>
                <li><b>詳細連線說明</b>：{sheets_detail}</li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### ⚡ 立即執行備份")
        st.write("手動打包所有訂單與 Audit Log 存至本地，並將完整資料庫轉譯備份至 Google Sheets「DB_Full_Backup」分頁。")
        if st.button("🚀 立即手動同步與備份", type="primary", use_container_width=True):
            with st.spinner("正在進行資料備份中..."):
                success, msg = sync_tool.backup_database_now(operator=current_user)
                if success:
                    st.success(msg)
                else:
                    st.warning(msg)

    with c2:
        st.markdown("#### 🔑 Google Service Account 金鑰設定")
        st.write("若欲啟用 Google Sheets 雲端備份，請上傳具備存取權限的 `service_account.json` 憑證檔 (並請將目標試算表分享編輯權限給服務帳號信箱)：")
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
# 14. 系統管理整合中心 (System Management Tabs)
# ==============================================================================
def render_system_management_view():
    st.markdown("### ⚙️ 系統管理與資料維護")
    tab_admin, tab_data, tab_logs, tab_cloud = st.tabs([
        "👤 管理人帳號管理",
        "📁 資料匯入與匯出",
        "📜 操作日誌 (Note Log)",
        "☁️ 雲端備份與設定",
    ])
    with tab_admin:
        render_admin_management_view()
    with tab_data:
        render_import_export_view()
    with tab_logs:
        render_audit_log_view()
    with tab_cloud:
        render_cloud_sync_view()


# ==============================================================================
# 15. 主程式入口與導覽路由 (Main App Entry)
# ==============================================================================
def main():
    apply_custom_styles()
    # 每次應用程式重跑/啟動皆執行：建表 (CREATE TABLE IF NOT EXISTS)
    init_db()

    # 啟動自動去重清理：依 (入住日期, 姓名, 電話) 複合鍵刪除重複訂單，
    # 僅保留 id 最大 (最近建立/最新版本) 的記錄；清理發生時清除資料快取，
    # 確保後續查詢與 session_state 立即反映無重複的結果
    if deduplicate_bookings() > 0:
        st.cache_data.clear()

    # 啟動拉取：僅於應用程式首次啟動時執行一次，
    # 避免每次重跑皆從 Google Sheets 拉取舊資料覆寫剛寫入 SQLite 的更新
    if not st.session_state.get("gsheet_startup_pull_done", False):
        pull_google_sheet_to_db()
        st.session_state["gsheet_startup_pull_done"] = True

    # 同步完成後重新查詢全量訂單並寫入 session_state，
    # 確保表格/指標/日曆等 UI 元件取得非空資料
    refresh_session_state_data()

    init_admins()

    # 檢查是否因 60 秒無動作觸發資安自動登出
    if st.query_params.get("auto_logout") == "1":
        st.query_params.clear()
        user_name = st.session_state.get("authenticated_user", "管理員")
        if user_name:
            add_audit_log(user_name, "資安自動登出", "超過 60 秒無操作，系統自動登出以維護資安")
        st.session_state["authenticated_user"] = None
        st.session_state["authenticated_pin"] = None
        st.session_state["editing_booking_id"] = None
        st.session_state["editing_booking_data"] = None
        reset_booking_form_date_keys()
        st.session_state["timeout_logged_out"] = True
        st.rerun()

    # 檢查登入狀態
    if not check_authentication():
        render_login_view()
        return

    # 已登入：嵌入 60 秒無動作自動登出監聽機制 (監聽 MouseMove, KeyDown, Click, Scroll, Touch)
    inject_auto_logout_js(60)

    # 導覽選單清單
    nav_options = [
        "📊 營位總覽看板",
        "➕ 新增露營預約",
        "🔍 訂位管理與搜尋",
        "⚙️ 系統管理",
    ]

    # 若有外部跳轉請求 (例如點擊卡片上的【編輯】按鈕)，在選單 widget 實例化前優先套用，杜絕 StreamlitWidgetAlreadyInstantiatedError
    if "jump_to_page" in st.session_state and st.session_state["jump_to_page"] in nav_options:
        target_page = st.session_state.pop("jump_to_page")
        st.session_state["main_navigation"] = target_page

    if "main_navigation" not in st.session_state or st.session_state["main_navigation"] not in nav_options:
        st.session_state["main_navigation"] = nav_options[0]

    current_nav = st.session_state["main_navigation"]
    nav_index = nav_options.index(current_nav)

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

    # 側邊欄導覽 (使用固定 key 管理 State，防止切換延遲與多次點擊問題)
    with st.sidebar:
        st.markdown("### 🌲 祿野營地導覽選單")
        st.write(f"目前登入者：**{st.session_state.get('authenticated_user', '管理員')}**")

        page_selection = st.radio(
            "前往模組：",
            nav_options,
            index=nav_index,
            key="main_navigation",
        )

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
            st.session_state["editing_booking_data"] = None
            reset_booking_form_date_keys()
            st.rerun()

    # 路由渲染對應頁面
    if page_selection == "📊 營位總覽看板":
        render_dashboard_view()
    elif page_selection == "➕ 新增露營預約":
        render_booking_form_view()
    elif page_selection == "🔍 訂位管理與搜尋":
        render_bookings_view()
    elif page_selection == "⚙️ 系統管理":
        render_system_management_view()

    # 頁面路由後，確認是否有刪除待辦
    if st.session_state.get("deleting_booking_id"):
        render_delete_confirm_modal()



if __name__ == "__main__":
    main()
