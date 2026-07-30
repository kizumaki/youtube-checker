import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import isodate
import pycountry
import requests
import zipfile
import os
import re
import datetime
import io
import json
from collections import Counter
from PIL import Image as PILImage
from supabase import create_client, Client
from concurrent.futures import ThreadPoolExecutor, as_completed

# Page Config
st.set_page_config(
    page_title="YT CHECKER PRO", 
    page_icon="🚀", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- MODERN ARTISTIC LIGHT THEME (PALETA COLOR PRO SYSTEM) ---
# Palette: #D95F26 (Tangerine), #3D2F29 (Espresso), #F4F2F1 (Warm White), #47A5D1 (Ocean), #FFFFFF (White), #6B7280 (Ash)
st.markdown("""
<style>
/* Base Theme */
.stApp {
    background-color: #F4F2F1 !important; 
    color: #3D2F29 !important; 
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
}

/* Sidebar Styling */
section[data-testid="stSidebar"] {
    background-color: #FFFFFF !important;
    border-right: 1px solid #E5E7EB !important;
    box-shadow: 2px 0 10px rgba(61, 47, 41, 0.03) !important;
}

section[data-testid="stSidebar"] .stMarkdown {
    color: #3D2F29 !important;
}

/* Header & Tab Area */
header[data-testid="stHeader"] {
    background-color: transparent !important;
}

/* HIGH-END ARTISTIC TABS DESIGN */
.stTabs [data-baseweb="tab-list"] {
    gap: 12px;
    background-color: #FFFFFF;
    padding: 10px 14px;
    border-radius: 16px;
    border: 1px solid #E5E7EB;
    box-shadow: 0 4px 20px rgba(61, 47, 41, 0.05);
}

.stTabs [data-baseweb="tab"] {
    background-color: #F4F2F1;
    border: 1px solid transparent !important;
    border-radius: 10px;
    color: #3D2F29 !important;
    font-weight: 700;
    font-size: 0.92rem;
    padding: 10px 22px;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    cursor: pointer !important;
}

/* Tab Hover Effect (Artistic Lift & Glow) */
.stTabs [data-baseweb="tab"]:hover {
    background-color: #FFFFFF !important;
    color: #D95F26 !important;
    border: 1px solid #D95F26 !important;
    transform: translateY(-3px) scale(1.02);
    box-shadow: 0 6px 16px rgba(217, 95, 38, 0.18) !important;
}

/* Active Tab (Gradient Tangerine) */
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #D95F26 0%, #B84A1A 100%) !important;
    color: #FFFFFF !important;
    border: none !important;
    transform: translateY(-1px) scale(1.03);
    box-shadow: 0 6px 18px rgba(217, 95, 38, 0.35) !important;
}

/* Standard Card Container Styling */
div[data-testid="stVerticalBlockBorderWrapper"] {
    background-color: #FFFFFF !important;
    border: 1px solid #E5E7EB !important;
    border-radius: 14px !important;
    padding: 10px !important;
    box-shadow: 0 2px 8px rgba(61, 47, 41, 0.04) !important;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}

div[data-testid="stVerticalBlockBorderWrapper"]:hover {
    box-shadow: 0 6px 18px rgba(61, 47, 41, 0.08) !important;
}

/* Active Inspected Card Highlight (Vivid Ocean Accent) */
div[data-testid="stVerticalBlockBorderWrapper"]:has(div.active-card-marker) {
    background-color: #EBF8FF !important; 
    border: 2.5px solid #47A5D1 !important; 
    box-shadow: 0 6px 20px rgba(71, 165, 209, 0.25) !important;
}

div[data-testid="stVerticalBlockBorderWrapper"]:has(div.active-card-marker) div[data-testid="stVerticalBlock"],
div[data-testid="stVerticalBlockBorderWrapper"]:has(div.active-card-marker) div[data-testid="stColumn"],
div[data-testid="stVerticalBlockBorderWrapper"]:has(div.active-card-marker) div[data-testid="element-container"],
div[data-testid="stVerticalBlockBorderWrapper"]:has(div.active-card-marker) div[data-testid="stMarkdownContainer"] {
    background-color: transparent !important;
}

/* In-Cart Card Highlight (Emerald Soft Green) */
div[data-testid="stVerticalBlockBorderWrapper"]:has(div.in-cart-marker) {
    background-color: #F0FDF4 !important;
    border: 2.5px solid #10B981 !important;
    box-shadow: 0 6px 20px rgba(16, 185, 129, 0.2) !important;
}

div[data-testid="stVerticalBlockBorderWrapper"]:has(div.in-cart-marker) div[data-testid="stVerticalBlock"],
div[data-testid="stVerticalBlockBorderWrapper"]:has(div.in-cart-marker) div[data-testid="stColumn"],
div[data-testid="stVerticalBlockBorderWrapper"]:has(div.in-cart-marker) div[data-testid="element-container"],
div[data-testid="stVerticalBlockBorderWrapper"]:has(div.in-cart-marker) div[data-testid="stMarkdownContainer"] {
    background-color: transparent !important;
}

/* Inputs & Selectboxes */
.stTextInput input, .stTextArea textarea, .stSelectbox select {
    background-color: #FFFFFF !important;
    color: #3D2F29 !important;
    border: 1px solid #D1D5DB !important;
    border-radius: 8px !important;
}

/* Buttons */
.stButton button {
    border-radius: 8px !important;
    font-weight: 700 !important;
    border: 1px solid #D1D5DB !important;
    background-color: #FFFFFF !important;
    color: #3D2F29 !important;
    transition: all 0.2s ease !important;
}

.stButton button:hover {
    border-color: #D95F26 !important;
    color: #D95F26 !important;
    background-color: #FFFBF9 !important;
}

/* Primary Button (Tangerine CTA) */
.stButton button[kind="primary"] {
    background-color: #D95F26 !important;
    color: #FFFFFF !important;
    border: none !important;
}

.stButton button[kind="primary"]:hover {
    background-color: #B84A1A !important;
    color: #FFFFFF !important;
    box-shadow: 0 4px 14px rgba(217, 95, 38, 0.4) !important;
}

/* Badges */
.badge-pro {
    display: inline-block;
    padding: 4px 12px;
    border-radius: 9999px;
    font-size: 0.75rem;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}

.badge-emerald { background-color: #D1FAE5; color: #065F46; border: 1px solid #10B981; }
.badge-tangerine { background-color: #FFEDD5; color: #9A3412; border: 1px solid #D95F26; }
.badge-ocean { background-color: #E0F2FE; color: #0369A1; border: 1px solid #47A5D1; }
</style>
""", unsafe_allow_html=True)

# Global Default API Key
DEFAULT_API_KEY = st.secrets.get("YOUTUBE_API_KEY", "AIzaSyDDBEJscqkGGpG1xtuL4wYPuFkS4BIL854")

# Connect to Supabase
@st.cache_resource
def init_supabase():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = init_supabase()

# --- DATABASE PERSISTENCE HELPERS ---
def load_api_keys_from_db():
    try:
        res = supabase.table("app_config").select("value").eq("key", "api_keys").execute()
        if res.data and len(res.data) > 0:
            return res.data[0]["value"]
    except Exception: pass
    return None

def save_api_keys_to_db(key_string):
    try:
        supabase.table("app_config").upsert({"key": "api_keys", "value": key_string}, on_conflict="key").execute()
        return True
    except Exception:
        return False

def load_cart_from_db():
    cart_dict = {}
    try:
        res = supabase.table("cart_items").select("*").execute()
        if res.data:
            for row in res.data:
                h = row["handle"]
                c_data = row.get("channel_data")
                if isinstance(c_data, str):
                    c_data = json.loads(c_data)
                cart_dict[h] = c_data
    except Exception: pass
    return cart_dict

def add_to_cart_db(pure_handle, channel_data):
    try:
        data_clean = dict(channel_data)
        if "recent_videos" in data_clean:
            del data_clean["recent_videos"]
        supabase.table("cart_items").upsert({"handle": pure_handle, "channel_data": data_clean}, on_conflict="handle").execute()
    except Exception: pass

def remove_from_cart_db(pure_handle):
    try:
        supabase.table("cart_items").delete().eq("handle", pure_handle).execute()
    except Exception: pass

def clear_cart_db():
    try:
        supabase.table("cart_items").delete().neq("handle", "___NONE___").execute()
    except Exception: pass

# --- INITIALIZE PERSISTENT STATE ---
if 'global_api_keys' not in st.session_state:
    db_keys = load_api_keys_from_db()
    st.session_state['global_api_keys'] = db_keys if db_keys else DEFAULT_API_KEY

if 'cart' not in st.session_state or 'cart_loaded' not in st.session_state:
    st.session_state['cart'] = load_cart_from_db()
    st.session_state['cart_loaded'] = True

if 'pending_seed_input' in st.session_state:
    st.session_state['seed_input_tab3'] = st.session_state['pending_seed_input']
    del st.session_state['pending_seed_input']

if 'seed_input_tab3' not in st.session_state:
    st.session_state['seed_input_tab3'] = "@NickDiGiovanni"

if 'pending_keywords' in st.session_state:
    st.session_state['custom_kw_tab3'] = st.session_state['pending_keywords']
    del st.session_state['pending_keywords']

if 'custom_kw_tab3' not in st.session_state:
    st.session_state['custom_kw_tab3'] = ""

if 'video_preview_cache' not in st.session_state:
    st.session_state['video_preview_cache'] = {}

if 'active_inspected_handle' not in st.session_state:
    st.session_state['active_inspected_handle'] = None

# --- CALLBACKS & HELPERS ---
def set_active_inspected_channel(pure_handle):
    st.session_state['active_inspected_handle'] = pure_handle

def set_api_keys(key_string):
    keys = [k.strip() for k in re.split(r'[\n,]+', key_string) if k.strip()]
    st.session_state['api_keys'] = keys if keys else [DEFAULT_API_KEY]

set_api_keys(st.session_state['global_api_keys'])

# --- SIDEBAR BRANDING & CONFIG ---
with st.sidebar:
    col1, col2, col3 = st.columns([1, 6, 1])
    with col2:
        try:
            st.image("logo.png", use_container_width=True)
        except Exception:
            st.markdown("<h2 style='text-align: center; color: #D95F26;'>🚀 LOGO</h2>", unsafe_allow_html=True)
            
    st.markdown("""
        <div style="text-align: center; padding-bottom: 12px;">
            <h2 style="margin: 6px 0 4px 0; font-weight: 900; font-size: 1.3rem; color: #3D2F29; letter-spacing: -0.02em;">YT CHECKER PRO</h2>
            <span class="badge-pro badge-emerald">Supabase Live</span>
        </div>
    """, unsafe_allow_html=True)
    st.divider()

    st.subheader("⚙️ Cấu Hình API Keys")
    active_key_count = len(st.session_state.get('api_keys', []))
    st.markdown(f"API Keys đang hoạt động: <span class='badge-pro badge-tangerine'>{active_key_count} Keys Active</span>", unsafe_allow_html=True)
    
    keys_input = st.text_area("Danh sách API Keys (Mỗi dòng 1 key):", value=st.session_state['global_api_keys'], height=180, key="api_keys_text_area")
    
    if st.button("💾 Lưu API Keys Vĩnh Viễn", type="primary", use_container_width=True):
        st.session_state['global_api_keys'] = keys_input
        set_api_keys(keys_input)
        save_api_keys_to_db(keys_input)
        st.toast("🎉 Đã lưu vĩnh viễn danh sách API Keys vào Supabase!")
        st.rerun()

    st.caption("💡 Mẹo: Keys và Giỏ Hàng lưu đám mây vĩnh viễn, không bao giờ bị mất khi F5 hay tắt máy.")
    st.divider()
    
    if st.button("🔄 Làm Mới Màn Hình Display", use_container_width=True, help="Dọn dẹp danh sách hiển thị để màn hình gọn gàng"):
        keys_to_clear = ['passed_channels', 'rejected_channels', 'last_inspected_data', 'last_inspected_handle', 'audit_success_msg', 'batch_check_new', 'batch_check_existing', 'active_inspected_handle']
        for key in keys_to_clear:
            if key in st.session_state:
                del st.session_state[key]
        for key in list(st.session_state.keys()):
            if key.startswith('audit_file_'):
                del st.session_state[key]
        st.rerun()

# --- HELPER TO DELETE CHANNEL COMPLETELY FROM SYSTEM ---
def delete_channel_from_system(pure_handle):
    if not pure_handle: return
    try:
        supabase.table("channels").delete().eq("handle", pure_handle).execute()
    except Exception: pass

    remove_from_cart_db(pure_handle)
    if pure_handle in st.session_state.get('cart', {}):
        del st.session_state['cart'][pure_handle]

    if st.session_state.get('active_inspected_handle') == pure_handle:
        st.session_state['active_inspected_handle'] = None

    if 'passed_channels' in st.session_state:
        st.session_state['passed_channels'] = [ch for ch in st.session_state['passed_channels'] if to_pure_id(ch.get('Handle')) != pure_handle]

    if 'rejected_channels' in st.session_state:
        st.session_state['rejected_channels'] = [ch for ch in st.session_state['rejected_channels'] if to_pure_id(ch.get('Handle')) != pure_handle]

    if 'batch_check_new' in st.session_state:
        st.session_state['batch_check_new'] = [ch for ch in st.session_state['batch_check_new'] if to_pure_id(ch.get('Handle')) != pure_handle]

    if 'batch_check_existing' in st.session_state:
        st.session_state['batch_check_existing'] = [ch for ch in st.session_state['batch_check_existing'] if to_pure_id(ch.get('Handle')) != pure_handle]

    audit_key = f"audit_file_{pure_handle}"
    if audit_key in st.session_state: del st.session_state[audit_key]
    if pure_handle in st.session_state.get('video_preview_cache', {}): del st.session_state['video_preview_cache'][pure_handle]

# --- API QUOTA ROTATION MANAGER ---
def yt_execute(request_func):
    keys = st.session_state.get('api_keys', [DEFAULT_API_KEY])
    if not keys: keys = [DEFAULT_API_KEY]
    idx = st.session_state.get('api_key_idx', 0)
    
    for attempt in range(len(keys)):
        key = keys[idx]
        try:
            yt = build("youtube", "v3", developerKey=key)
            req = request_func(yt)
            return req.execute()
        except HttpError as e:
            if e.resp.status in [403, 400]:
                idx = (idx + 1) % len(keys)
                st.session_state['api_key_idx'] = idx
            else:
                raise e
    raise Exception("❌ Toàn bộ API Keys bạn nhập đã bị chết hoặc cạn sạch Quota!")

# --- HELPER FUNCTIONS ---
def to_pure_id(raw_val):
    if not raw_val or pd.isna(raw_val) or str(raw_val).strip().upper() in ["N/A", "NAN", "NONE", ""]: return None
    s = str(raw_val).strip()
    m_url = re.search(r'youtube\.com/(?:@|c/|user/|channel/)?([^\s?#/]+)', s, re.IGNORECASE)
    if m_url: s = m_url.group(1)
    s = re.sub(r'[\s]+', '', s)
    s = re.sub(r'^@+', '', s).strip().lower()
    return s if s else None

def is_long_form_video(v, min_seconds=180):
    title = v.get('Title', '').lower()
    if '#shorts' in title or '#short' in title: return False
    if v.get('Seconds', 0) <= min_seconds: return False
    return True

def extract_handles_from_text(text_block):
    if not text_block: return []
    lines = re.split(r'[\n,\t\r]+', str(text_block))
    handles = []
    for line in lines:
        p = to_pure_id(line)
        if p and p not in handles: handles.append(p)
    return handles

def extract_handles_from_file(uploaded_file):
    handles = []
    fname = uploaded_file.name.lower()
    try:
        if fname.endswith('.txt'):
            content = uploaded_file.read().decode("utf-8", errors="ignore")
            handles = extract_handles_from_text(content)
        elif fname.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
            for col in df.columns:
                for val in df[col].dropna():
                    p = to_pure_id(val)
                    if p and p not in handles: handles.append(p)
        elif fname.endswith('.xlsx') or fname.endswith('.xls'):
            df = pd.read_excel(uploaded_file)
            for col in df.columns:
                for val in df[col].dropna():
                    p = to_pure_id(val)
                    if p and p not in handles: handles.append(p)
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
    return handles

def extract_handle_from_filename(filename):
    base = os.path.basename(filename)
    base_no_ext = os.path.splitext(base)[0]
    pattern = r'_(?:backlog|\d{4}|\d{2,4}[-_/.]\d{1,2}[-_/.]\d{1,2}|\d{1,2}[-_/.]\d{2,4}|\d{6,8})(?:_.*)?$'
    cleaned = re.sub(pattern, '', base_no_ext, flags=re.IGNORECASE)
    cleaned = re.sub(r'[\s]+', '', cleaned)
    pure_id = re.sub(r'^@+', '', cleaned).strip().lower()
    return pure_id if pure_id else None

def is_within_last_90_days(date_str):
    if not date_str or date_str == "N/A": return False
    s = str(date_str).strip().lower()
    today = datetime.date.today()
    m_iso = re.match(r'^(\d{4})-(\d{2})-(\d{2})', s)
    if m_iso:
        try:
            dt = datetime.date(int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3)))
            return 0 <= (today - dt).days <= 90
        except Exception: return False
    return False

# --- STRICT FILTERS CONFIGURATION ---
EXCLUDED_COUNTRIES = {'CN', 'TW', 'HK', 'TH', 'IN', 'VN'}
EXCLUDED_TEXT_REGEX = re.compile(r'[\u0E00-\u0E7F]|[\u4E00-\u9FFF]|[\u0900-\u097F]|[àáảãạâầấẩẫậăằắẳẵặèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]', re.IGNORECASE)
EXCLUDED_KEYWORDS = ['official mv', 'music video', 'official audio', 'album', 'song', 'records', 'lyrics', 'remix', 'vocal', 'cover', 'news', 'politics', 'lgbt', 'lgbtq', 'gay', 'lesbian', 'transgender', 'war', 'military', 'ukraine', 'russia', 'tin tức', 'chính trị', 'thời sự', 'chiến tranh', 'đảng', 'quân sự']
STOP_WORDS = {'the', 'a', 'an', 'and', 'or', 'in', 'on', 'at', 'to', 'for', 'of', 'with', 'by', 'is', 'it', 'this', 'that', 'ep', 'episode', 'part', 'video', 'shorts', 'full', 'hd', '2024', '2025', '2026', 'official', 'channel', 'vs', 'dude', 'perfect', 'nick', 'digiovanni', 'mrbeast', 'pewdiepie'}

def passes_layer1_metadata_filter(title, desc, country_code):
    if country_code in EXCLUDED_COUNTRIES: return False, f"Quốc gia bị loại ({country_code})"
    combined_text = f"{title} {desc}".lower()
    if EXCLUDED_TEXT_REGEX.search(combined_text): return False, "Ngôn ngữ không phù hợp (Trung, Thái, Hindi, Việt)"
    for kw in EXCLUDED_KEYWORDS:
        if kw in combined_text: return False, f"Chứa từ khóa bị cấm: '{kw}'"
    return True, "OK"

def clean_and_extract_keywords(text, seed_handle=""):
    seed_clean = seed_handle.replace('@', '').lower()
    words = re.findall(r'\b[a-zA-Z0-9]{3,}\b', text.lower())
    filtered = [w for w in words if w not in STOP_WORDS and w not in seed_clean]
    return filtered

# --- YOUTUBE API OPERATIONS (CACHE OPTIMIZED) ---
@st.cache_data(ttl=86400, show_spinner=False)
def get_channel_id_by_handle(handle):
    clean = handle.replace('@', '').split('/')[-1].strip()
    try:
        res = yt_execute(lambda yt: yt.channels().list(part="id", forHandle=clean))
        if 'items' in res and len(res['items']) > 0: return res['items'][0]['id']
    except Exception: pass
    try:
        res = yt_execute(lambda yt: yt.search().list(part="snippet", q=clean, type="channel", maxResults=1))
        if 'items' in res and len(res['items']) > 0: return res['items'][0]['snippet']['channelId']
    except Exception: pass
    return None

@st.cache_data(ttl=86400, show_spinner=False)
def extract_channel_master_keywords(channel_id):
    keywords_pool, channel_keywords, top_tags, categories = [], [], [], []
    try:
        ch_res = yt_execute(lambda yt: yt.channels().list(part="brandingSettings,contentDetails,snippet,topicDetails", id=channel_id))
        if 'items' in ch_res and len(ch_res['items']) > 0:
            item = ch_res['items'][0]
            raw_kw = item.get('brandingSettings', {}).get('channel', {}).get('keywords', '')
            found_kw = re.findall(r'"([^"]+)"|\b([a-zA-Z0-9]{3,})\b', raw_kw)
            for k1, k2 in found_kw:
                kw = k1 or k2
                if kw and len(kw) > 2 and kw.lower() not in STOP_WORDS:
                    keywords_pool.append(kw.lower())
                    channel_keywords.append(kw.lower())
                    
            topics = item.get('topicDetails', {}).get('topicCategories', [])
            for t in topics: categories.append(t.split('/')[-1].replace('_', ' '))
                
            uploads_playlist = item.get('contentDetails', {}).get('relatedPlaylists', {}).get('uploads', '')
            if uploads_playlist:
                v_res = yt_execute(lambda yt: yt.playlistItems().list(part="snippet", playlistId=uploads_playlist, maxResults=15))
                v_ids = [v['snippet']['resourceId']['videoId'] for v in v_res.get('items', [])]
                if v_ids:
                    v_detail_res = yt_execute(lambda yt: yt.videos().list(part="snippet", id=','.join(v_ids)))
                    for v_item in v_detail_res.get('items', []):
                        for tag in v_item.get('snippet', {}).get('tags', []):
                            if len(tag) > 2 and tag.lower() not in STOP_WORDS:
                                keywords_pool.append(tag.lower())
                                top_tags.append(tag.lower())
    except Exception: pass

    most_common_kws = [word for word, count in Counter(keywords_pool).most_common(15)]
    top_tag_counts = [word for word, count in Counter(top_tags).most_common(10)]
    return {"master_keywords": most_common_kws, "channel_keywords": list(set(channel_keywords))[:10], "top_tags": top_tag_counts, "categories": categories}

def get_channel_details(channel_id):
    res = yt_execute(lambda yt: yt.channels().list(part="snippet,contentDetails,statistics", id=channel_id))
    if 'items' in res and len(res['items']) > 0:
        item = res['items'][0]
        playlist_id = item.get('contentDetails', {}).get('relatedPlaylists', {}).get('uploads', '')
        sub_count = int(item['statistics'].get('subscriberCount', 0))
        description = item['snippet'].get('description', 'No description available.')
        joined_date_raw = item['snippet'].get('publishedAt', '')
        joined_date = ""
        if joined_date_raw:
            try: joined_date = pd.to_datetime(joined_date_raw).strftime("%b %d, %Y")
            except Exception: joined_date = joined_date_raw[:10]
        country_code = item['snippet'].get('country', 'N/A')
        country_name = pycountry.countries.get(alpha_2=country_code).name if country_code != 'N/A' and pycountry.countries.get(alpha_2=country_code) else country_code
        avatar_url = item['snippet'].get('thumbnails', {}).get('high', {}).get('url', '')
        return playlist_id, sub_count, description, joined_date, country_name, country_code, avatar_url
    return None, 0, "", "", "", "", ""

def get_video_details(video_ids):
    video_data = []
    total = len(video_ids)
    if total == 0: return video_data
    for i in range(0, total, 50):
        try:
            chunk = video_ids[i:i+50]
            res = yt_execute(lambda yt: yt.videos().list(part="snippet,contentDetails,statistics", id=','.join(chunk)))
            for item in res.get('items', []):
                duration_seconds = int(isodate.parse_duration(item['contentDetails']['duration']).total_seconds())
                h, rem = divmod(duration_seconds, 3600)
                m, s = divmod(rem, 60)
                dur_str = f"{h}:{m:02d}:{s:02d}" if h > 0 else f"{m}:{s:02d}"

                pub_date = item['snippet']['publishedAt']
                try: formatted_date = pd.to_datetime(pub_date).strftime("%d-%m-%Y")
                except Exception: formatted_date = pub_date[:10]
                
                video_data.append({
                    'Title': item['snippet']['title'], 
                    'Link': f"https://youtube.com/watch?v={item['id']}",
                    'Length (Exact)': dur_str, 
                    'Seconds': duration_seconds,
                    'Views': int(item['statistics'].get('viewCount', 0)), 
                    'Published Date': formatted_date,
                    'Video ID': item['id']
                })
        except Exception: pass
    return video_data

@st.cache_data(ttl=43200, show_spinner=False)
def get_6_recent_videos(pure_handle):
    long_vids = []
    try:
        cid = get_channel_id_by_handle(pure_handle)
        if cid:
            playlist_id, _, _, _, _, _, _ = get_channel_details(cid)
            if playlist_id:
                next_token = None
                for _ in range(2):
                    v_res = yt_execute(lambda yt: yt.playlistItems().list(part="snippet", playlistId=playlist_id, maxResults=50, pageToken=next_token))
                    v_ids = [v_item['snippet']['resourceId']['videoId'] for v_item in v_res.get('items', [])]
                    if v_ids:
                        v_details = get_video_details(v_ids)
                        for v in v_details:
                            if is_long_form_video(v, min_seconds=180) and v not in long_vids:
                                long_vids.append(v)
                            if len(long_vids) >= 6: break
                    next_token = v_res.get('nextPageToken')
                    if not next_token or len(long_vids) >= 6: break
                return long_vids[:6]
    except Exception: pass
    return long_vids[:6]

# --- STREAMLIT MODAL DIALOG ---
@st.dialog("🎬 6 Video Dài (Long-form) Mới Nhất", width="large")
def show_video_dialog(pure_handle, pre_fetched_videos=None):
    st.markdown(f"### 📺 Kênh đang xem: `@{pure_handle}`")
    st.markdown(f"🔗 **[Mở thẳng Tab Videos trên YouTube](https://youtube.com/@{pure_handle}/videos)**")
    
    vids = []
    if pre_fetched_videos:
        vids = [v for v in pre_fetched_videos if is_long_form_video(v, min_seconds=180)]
        
    if len(vids) < 6:
        vids = get_6_recent_videos(pure_handle)
        
    vids = vids[:6]
    
    if vids:
        st.divider()
        for row_idx in range(0, len(vids), 2):
            col1, col2 = st.columns(2)
            
            v1 = vids[row_idx]
            with col1:
                vid_id1 = v1.get('Video ID') or (v1.get('Link', '').split('v=')[-1] if 'v=' in v1.get('Link', '') else '')
                if vid_id1:
                    st.image(f"https://img.youtube.com/vi/{vid_id1}/hqdefault.jpg", use_container_width=True)
                st.markdown(f"**[{v1['Title'][:45]}...]({v1['Link']})**")
                st.caption(f"👀 {v1.get('Views', 0):,} views | ⏳ {v1.get('Length (Exact)', 'N/A')} | 📅 {v1.get('Published Date', '')}")
            
            if row_idx + 1 < len(vids):
                v2 = vids[row_idx + 1]
                with col2:
                    vid_id2 = v2.get('Video ID') or (v2.get('Link', '').split('v=')[-1] if 'v=' in v2.get('Link', '') else '')
                    if vid_id2:
                        st.image(f"https://img.youtube.com/vi/{vid_id2}/hqdefault.jpg", use_container_width=True)
                    st.markdown(f"**[{v2['Title'][:45]}...]({v2['Link']})**")
                    st.caption(f"👀 {v2.get('Views', 0):,} views | ⏳ {v2.get('Length (Exact)', 'N/A')} | 📅 {v2.get('Published Date', '')}")
            
            st.divider()
    else:
        st.caption("Không tìm thấy video dài (Kênh này chỉ đăng Shorts hoặc chưa có video dài trên 3 phút).")

    if st.button("❌ Đóng Cửa Sổ Preview", type="primary", use_container_width=True):
        st.rerun()

def generate_v414_excel_report(clean_handle, sub_count, channel_desc, channel_joined, channel_country, avatar_url, video_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = clean_handle[:31]

    date_str = datetime.datetime.now().strftime("%d-%m-%Y")
    total_videos = len(video_data)
    total_views = sum(v['Views'] for v in video_data)
    total_minutes = round(sum(v['Seconds'] for v in video_data) / 60)

    # TAB 1: MAIN DATA SHEET
    ws.merge_cells('A1:E1')
    ws['A1'] = f"{clean_handle.upper()} YOUTUBE CHANNEL SUMMARY REPORT - up to {date_str}"
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws['A2'] = f"Total Videos: {total_videos:,}"
    ws['A3'] = f"Total Duration: {total_minutes:,} minutes"
    ws['A4'] = f"Total Views: {total_views:,}"
    ws['A5'] = f"Total Subscribers: {sub_count:,}"
    ws['A6'] = f"Country Location: {channel_country}"
    ws['A7'] = f"Channel Joined Date: {channel_joined}"

    ws['A9'] = "Channel Description Text:"
    ws['A9'].font = Font(bold=True, italic=True)
    ws['A10'] = channel_desc
    ws['A10'].alignment = Alignment(vertical="top", wrap_text=True)

    for row in range(2, 8): ws[f'A{row}'].font = Font(bold=True)

    if avatar_url:
        try:
            res = requests.get(avatar_url, timeout=5)
            if res.status_code == 200:
                img = PILImage.open(io.BytesIO(res.content))
                img = img.resize((140, 140))
                temp_buf = io.BytesIO()
                img.save(temp_buf, format="PNG")
                temp_buf.seek(0)
                ws.add_image(ExcelImage(temp_buf), 'C10')
        except Exception: pass

    headers = ["Video Title", "Link", "Length", "Views", "Published Date"]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=12, column=col_num, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[12].height = 24

    for idx, v in enumerate(video_data):
        r = idx + 13
        cA = ws.cell(row=r, column=1, value=v['Title']); cA.font = Font(name="Calibri", size=11)
        cB = ws.cell(row=r, column=2, value=v['Link'])
        if v.get('Link'): cB.hyperlink = v['Link']; cB.font = Font(name="Calibri", size=11, color="0563C1", underline="single")
        ws.cell(row=r, column=3, value=v['Length (Exact)']).alignment = Alignment(horizontal="center", vertical="center")
        cD = ws.cell(row=r, column=4, value=v['Views']); cD.number_format = '#,##0'
        ws.cell(row=r, column=5, value=v['Published Date']).alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions['A'].width = 55; ws.column_dimensions['B'].width = 45; ws.column_dimensions['C'].width = 22; ws.column_dimensions['D'].width = 15; ws.column_dimensions['E'].width = 15

    # TAB 2: DASHBOARD
    ws_charts = wb.create_sheet(title="Top 10 Video Title")
    top_10_videos = sorted(video_data, key=lambda x: x['Views'], reverse=True)[:10]

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid"); header_font = Font(bold=True, color="FFFFFF")
    ws_charts['A1'] = "Top 10 Most Viewed Videos (Click to Watch)"; ws_charts['A1'].font, ws_charts['A1'].fill = header_font, header_fill
    ws_charts['B1'] = "Views"; ws_charts['B1'].font, ws_charts['B1'].fill = header_font, header_fill

    PALETTE = [{"fill": "2F5597", "font": "FFFFFF"}, {"fill": "C00000", "font": "FFFFFF"}, {"fill": "70AD47", "font": "FFFFFF"}, {"fill": "7030A0", "font": "FFFFFF"}, {"fill": "00C0C0", "font": "FFFFFF"}, {"fill": "E37222", "font": "FFFFFF"}, {"fill": "41536B", "font": "FFFFFF"}, {"fill": "A04000", "font": "FFFFFF"}, {"fill": "385723", "font": "FFFFFF"}, {"fill": "626262", "font": "FFFFFF"}]
    ws_charts['D1'] = f"📊 Top 10 Most Viewed Videos - {clean_handle}"
    ws_charts['D1'].font = Font(bold=True, size=14, color="1F4E78")

    for row_idx, video in enumerate(top_10_videos, start=2):
        color_idx = (row_idx - 2) % len(PALETTE); style = PALETTE[color_idx]
        title_cell = ws_charts.cell(row=row_idx, column=1, value=video['Title'][:45] + "...")
        title_cell.hyperlink = video['Link']; title_cell.font = Font(bold=True, color=style["font"], underline="single"); title_cell.fill = PatternFill("solid", fgColor=style["fill"])
        view_cell = ws_charts.cell(row=row_idx, column=2, value=video['Views']); view_cell.number_format = '#,##0'
        view_cell.font, view_cell.fill = Font(bold=True, color=style["font"]), PatternFill("solid", fgColor=style["fill"])

    ws_charts.column_dimensions['A'].width = 50; ws_charts.column_dimensions['B'].width = 15

    if len(top_10_videos) > 0:
        chart = BarChart(); chart.type = "col"; chart.y_axis.title = 'Total Views'; chart.x_axis.title = 'Videos'
        chart.add_data(Reference(ws_charts, min_col=2, min_row=1, max_col=2, max_row=len(top_10_videos)+1), titles_from_data=True)
        chart.set_categories(Reference(ws_charts, min_col=1, min_row=2, max_col=1, max_row=len(top_10_videos)+1))
        chart.legend = None
        for idx in range(len(top_10_videos)): chart.series[0].dPt.append(DataPoint(idx=idx, graphicalProperties={"solidFill": PALETTE[idx]["fill"]}))
        chart.width = 22; chart.height = 14
        ws_charts.add_chart(chart, "D3")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# --- REUSABLE COMPONENT: RENDER SHARED CART ---
def render_shared_cart_ui(key_suffix=""):
    st.divider()
    cart_items = st.session_state['cart']
    st.subheader(f"🛒 Giỏ Hàng Dùng Chung ({len(cart_items)} Kênh)")
    if cart_items:
        df_cart = pd.DataFrame(list(cart_items.values()))
        if 'Handle' in df_cart.columns:
            df_cart['Tab Videos'] = df_cart['Handle'].apply(lambda h: f"https://youtube.com/{to_pure_id(h)}/videos")
            df_cart['Link Kênh'] = df_cart['Handle'].apply(lambda h: f"https://youtube.com/{to_pure_id(h)}")
            
        if 'recent_videos' in df_cart.columns:
            df_cart = df_cart.drop(columns=['recent_videos'])

        st.dataframe(df_cart, use_container_width=True, column_config={
            "Link Kênh": st.column_config.LinkColumn("Trang Chủ", display_text="🏠 Kênh"),
            "Tab Videos": st.column_config.LinkColumn("Tab Videos", display_text="🎬 Videos")
        })
        
        c1, c2, c3, c4 = st.columns(4)
        c1.download_button("📄 Tải TXT", data="\n".join([i["Handle"] for i in cart_items.values()]), file_name="gio_hang_dung_chung.txt", use_container_width=True, key=f"dl_txt_cart_{key_suffix}")
        buf_xl = io.BytesIO(); df_cart.to_excel(buf_xl, index=False)
        c2.download_button("📊 Tải Excel", data=buf_xl.getvalue(), file_name="gio_hang_dung_chung.xlsx", use_container_width=True, key=f"dl_xl_cart_{key_suffix}")
        if c3.button("⚡ Nạp Toàn Bộ Vào DB", type="primary", use_container_width=True, key=f"push_db_cart_{key_suffix}"):
            data_db = [{"handle": to_pure_id(i["Handle"]), "youtuber_name": i.get("Tên Kênh", ""), "source": "Cart Import"} for i in cart_items.values()]
            supabase.table("channels").upsert(data_db, on_conflict="handle").execute()
            st.success(f"🎉 Đã nạp {len(data_db)} kênh vào Database!")
        if c4.button("🧹 Xóa Sạch Giỏ Hàng", use_container_width=True, key=f"clear_cart_{key_suffix}"): 
            clear_cart_db()
            st.session_state['cart'] = {}
            st.success("🎉 Đã xóa sạch Giỏ Hàng!")
            st.rerun()
    else:
        st.info("Giỏ hàng đang trống. Bấm '🛒 Thêm' ở Tab 1 hoặc Tab 3 để nhặt kênh vào giỏ!")

# --- APP HEADER ---
st.markdown("""
    <div style="padding: 5px 0 15px 0;">
        <h1 style="font-weight: 900; color: #3D2F29; margin-bottom: 0; letter-spacing: -0.03em;">YT CHECKER PRO</h1>
        <p style="color: #6B7280; font-size: 1.05rem;">Hệ thống phân tích, tìm kiếm kênh đồng ngách Đa Luồng Siêu Tốc.</p>
    </div>
""", unsafe_allow_html=True)

# Tabs
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🔍 Tra cứu Handle Hàng Loạt", 
    "⚡ Cào Live & Tạo Báo Cáo Audit", 
    "🎯 Săn Kênh Tương Tự (Multi-Threaded)",
    "📤 Upload Cập nhật Data", 
    "📊 Xem Database",
    "✨ Soi Từ Khóa Kênh (SEO Inspector)"
])

# --- TAB 1: BATCH SEARCH ---
with tab1:
    st.subheader("🔍 Kiểm tra Trùng Lặp Danh Sách Handle Hàng Loạt")
    col_s1, col_s2 = st.columns([2, 1])
    with col_s1: text_input_area = st.text_area("Dán danh sách Handle/Link kênh vào đây (mỗi kênh 1 dòng):", height=180)
    with col_s2: file_input_check = st.file_uploader("Hoặc Upload file danh sách (.txt, .csv, .xlsx):")
        
    if st.button("🔎 Bắt Đầu Kiểm Tra Hàng Loạt", type="primary"):
        all_target_handles = set()
        if text_input_area:
            for h in extract_handles_from_text(text_input_area): all_target_handles.add(h)
        if file_input_check:
            for h in extract_handles_from_file(file_input_check): all_target_handles.add(h)
                
        target_list = list(all_target_handles)
        if not target_list: st.warning("⚠️ Vui lòng dán danh sách Handle hoặc chọn file để kiểm tra!")
        else:
            with st.spinner(f"Đang đối chiếu {len(target_list)} Handle với Database Supabase..."):
                response = supabase.table("channels").select("handle, youtuber_name").in_("handle", target_list).execute()
                db_matches = {item["handle"].lower(): item for item in response.data} if response.data else {}
                
                new_handles, existing_handles = [], []
                for h in target_list:
                    p_id = to_pure_id(h)
                    if p_id in db_matches: existing_handles.append({"Handle": f"@{p_id}", "Tên Kênh": db_matches[p_id].get("youtuber_name", "N/A"), "Trạng thái": "❌ Đã có trong DB"})
                    else: new_handles.append({"Handle": f"@{p_id}", "Tên Kênh": p_id.upper(), "Link Kênh": f"https://www.youtube.com/@{p_id}", "Trạng thái": "✅ Kênh Mới (Chưa làm)"})

                st.session_state['batch_check_new'] = new_handles
                st.session_state['batch_check_existing'] = existing_handles

    if 'batch_check_new' in st.session_state or 'batch_check_existing' in st.session_state:
        new_handles = st.session_state.get('batch_check_new', [])
        existing_handles = st.session_state.get('batch_check_existing', [])

        st.divider()
        m1, m2, m3 = st.columns(3)
        m1.metric("Tổng số kênh kiểm tra", f"{len(new_handles) + len(existing_handles)} kênh")
        m2.metric("❌ Kênh Đã Tồn Tại", f"{len(existing_handles)} kênh")
        m3.metric("✅ Kênh Mới Có Thể Làm", f"{len(new_handles)} kênh")
        
        res_tab1, res_tab2 = st.tabs([f"✅ Kênh Mới Chưa Làm ({len(new_handles)})", f"❌ Kênh Đã Tồn Tại ({len(existing_handles)})"])
        with res_tab1:
            if new_handles:
                if st.button("🛒 Thêm TẤT CẢ Kênh Mới vào Giỏ Hàng", type="primary", key="btn_add_all_t1"):
                    for item in new_handles:
                        p_id = to_pure_id(item["Handle"])
                        item_data = {
                            "Handle": item["Handle"],
                            "Tên Kênh": item.get("Tên Kênh", p_id.upper()),
                            "Link Kênh": f"https://www.youtube.com/@{p_id}",
                            "Trạng Thái DB": "✅ KÊNH MỚI"
                        }
                        st.session_state['cart'][p_id] = item_data
                        add_to_cart_db(p_id, item_data)
                    st.success(f"🎉 Đã thêm {len(new_handles)} kênh mới vào Giỏ hàng chung!")
                    st.rerun()

                st.divider()
                for idx, item in enumerate(new_handles):
                    p_id = to_pure_id(item["Handle"])
                    is_active = (p_id == st.session_state.get('active_inspected_handle'))
                    is_in_cart = p_id in st.session_state['cart']
                    
                    with st.container(border=True):
                        if is_active:
                            st.markdown('<div class="active-card-marker"></div>', unsafe_allow_html=True)
                            st.markdown('<div style="background-color: #47A5D1; color: white; padding: 6px 12px; border-radius: 6px; font-weight: bold; margin-bottom: 10px;">🔍 ĐANG XEM 6 VIDEO MỚI CỦA KÊNH NÀY</div>', unsafe_allow_html=True)
                        elif is_in_cart:
                            st.markdown('<div class="in-cart-marker"></div>', unsafe_allow_html=True)
                            st.markdown('<div style="background-color: #10B981; color: white; padding: 6px 12px; border-radius: 6px; font-weight: bold; margin-bottom: 10px;">🛒 ĐÃ CÓ TRONG GIỎ HÀNG</div>', unsafe_allow_html=True)

                        c1, c2, c3 = st.columns([3.5, 3.5, 3.0])
                        with c1:
                            st.markdown(f"### [{item['Handle']}]({item['Link Kênh']})")
                            st.write(f"**{item.get('Tên Kênh', p_id.upper())}**")
                            if st.button("👁️ Xem 6 Video Mới", key=f"btn_prev_t1_{idx}_{p_id}", on_click=set_active_inspected_channel, args=(p_id,)):
                                show_video_dialog(p_id)
                        with c2:
                            st.markdown(f"**Trạng thái:** {item['Trạng thái']}")
                        with c3:
                            st.write("**Thao tác:**")
                            bc1, bc2 = st.columns(2)
                            if is_in_cart:
                                if bc1.button("❌ Bỏ Giỏ", key=f"rm_t1_{idx}_{p_id}", use_container_width=True):
                                    remove_from_cart_db(p_id)
                                    del st.session_state['cart'][p_id]
                                    st.rerun()
                            else:
                                if bc1.button("🛒 Thêm Giỏ", key=f"add_t1_{idx}_{p_id}", use_container_width=True):
                                    item_data = {
                                        "Handle": item["Handle"],
                                        "Tên Kênh": item.get("Tên Kênh", p_id.upper()),
                                        "Link Kênh": f"https://www.youtube.com/@{p_id}",
                                        "Trạng Thái DB": "✅ KÊNH MỚI"
                                    }
                                    st.session_state['cart'][p_id] = item_data
                                    add_to_cart_db(p_id, item_data)
                                    st.rerun()

                            if bc2.button("🗑️ Xóa", key=f"del_t1_new_{idx}_{p_id}", use_container_width=True, help="Loại bỏ kênh này khỏi danh sách"):
                                delete_channel_from_system(p_id)
                                st.toast(f"🗑️ Đã xóa kênh @{p_id}!")
                                st.rerun()
            else:
                st.info("Tất cả kênh đều đã tồn tại trong Database!")

        with res_tab2:
            if existing_handles:
                for idx, item in enumerate(existing_handles):
                    p_id = to_pure_id(item["Handle"])
                    is_active = (p_id == st.session_state.get('active_inspected_handle'))
                    
                    with st.container(border=True):
                        if is_active:
                            st.markdown('<div class="active-card-marker"></div>', unsafe_allow_html=True)
                            st.markdown('<div style="background-color: #47A5D1; color: white; padding: 6px 12px; border-radius: 6px; font-weight: bold; margin-bottom: 10px;">🔍 ĐANG XEM 6 VIDEO MỚI CỦA KÊNH NÀY</div>', unsafe_allow_html=True)

                        c1, c2, c3 = st.columns([4.0, 4.0, 2.0])
                        with c1:
                            st.markdown(f"### [{item['Handle']}](https://youtube.com/@{p_id})")
                            st.write(f"**{item.get('Tên Kênh', 'N/A')}**")
                            if st.button("👁️ Xem 6 Video Mới", key=f"btn_prev_t1_ext_{idx}_{p_id}", on_click=set_active_inspected_channel, args=(p_id,)):
                                show_video_dialog(p_id)
                        with c2:
                            st.markdown(f"**Trạng thái:** {item['Trạng thái']}")
                        with c3:
                            if st.button("🗑️ Xóa DB", key=f"del_t1_ext_{idx}_{p_id}", use_container_width=True):
                                delete_channel_from_system(p_id)
                                st.toast(f"🗑️ Đã xóa kênh @{p_id} khỏi DB!")
                                st.rerun()

    render_shared_cart_ui(key_suffix="tab1")

# --- TAB 2: LIVE API SCRAPER ---
with tab2:
    st.subheader("⚡ Cào dữ liệu Live & Xuất Báo Cáo Audit chuẩn V4.14")
    channel_url_input = st.text_input("Dán Link kênh hoặc Handle vào đây:", value="@4wd247")

    if channel_url_input and st.button("🚀 Xử lý Kênh & Tạo Báo Cáo V4.14", type="primary"):
        pure_h = to_pure_id(channel_url_input)
        if pure_h:
            try:
                b_data, f_name = run_single_channel_audit(pure_h)
                if b_data:
                    supabase.table("channels").upsert([{"handle": pure_h, "youtuber_name": pure_h.upper(), "source": "YouTube API V4.14"}], on_conflict="handle").execute()
                    st.success(f"🎉 Đã dựng xong báo cáo Audit!")
                    st.download_button("📥 Tải về File Audit V4.14", data=b_data, file_name=f_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except Exception as e: st.error(f"Lỗi: {e}")

# --- TAB 3: MULTI-THREADED SMART RELATED FINDER ---
def process_single_candidate(item, min_subs_choice, min_duration_choice, db_existing_set):
    c_handle = to_pure_id(item['snippet'].get('customUrl', '')) or item['id'].lower()
    c_title = item['snippet']['title']
    c_desc = item['snippet'].get('description', '')
    c_country = item['snippet'].get('country', 'N/A')
    c_subs = int(item['statistics'].get('subscriberCount', 0))
    c_video_count = int(item['statistics'].get('videoCount', 0))
    c_url = f"https://www.youtube.com/@{c_handle}"
    db_status = "❌ Đã có trong DB" if c_handle in db_existing_set else "✅ KÊNH MỚI"
    
    c_playlist = item.get('contentDetails', {}).get('relatedPlaylists', {}).get('uploads', '')
    latest_date, has_qualifying_video = "N/A", False
    recent_vids = []
    
    if c_playlist and c_video_count > 0:
        try:
            v_res = yt_execute(lambda yt: yt.playlistItems().list(part="snippet", playlistId=c_playlist, maxResults=50))
            v_ids = [v_item['snippet']['resourceId']['videoId'] for v_item in v_res.get('items', [])]
            if v_ids:
                v_details = get_video_details(v_ids)
                long_vids = [v for v in v_details if is_long_form_video(v, min_seconds=180)]
                if long_vids:
                    latest_date = long_vids[0]['Published Date']
                    has_qualifying_video = any(v['Seconds'] >= min_duration_choice for v in long_vids)
                    recent_vids = long_vids[:6]
        except Exception: pass

    base_data = {"Handle": f"@{c_handle}", "Link Kênh": c_url, "Tên Kênh": c_title, "Subscribers": f"{c_subs:,}", "Quốc gia": c_country, "Video Gần Nhất": latest_date, "Tổng Số Video": f"{c_video_count:,}", "Trạng Thái DB": db_status, "recent_videos": recent_vids}

    if c_subs < min_subs_choice: 
        base_data["Lý do loại"] = f"Dưới {min_subs_choice:,} Subs"
        return False, base_data
    passes_l1, l1_reason = passes_layer1_metadata_filter(c_title, c_desc, c_country)
    if not passes_l1: 
        base_data["Lý do loại"] = l1_reason
        return False, base_data
    if c_video_count == 0 or not c_playlist: 
        base_data["Lý do loại"] = "Kênh trống"
        return False, base_data
    if not is_within_last_90_days(latest_date): 
        base_data["Lý do loại"] = f"Bỏ trống (Mới nhất: {latest_date})"
        return False, base_data
    if not has_qualifying_video: 
        base_data["Lý do loại"] = "Shorts-only"
        return False, base_data
        
    return True, base_data

with tab3:
    st.subheader("🎯 Săn Kênh Tương Tự & Giỏ Hàng (Multi-threaded Speed)")
    
    if 'audit_success_msg' in st.session_state:
        st.success(st.session_state['audit_success_msg'])
        del st.session_state['audit_success_msg']

    trigger_auto_start_search = False
    if st.session_state.get('trigger_deep_search_now', False):
        st.session_state['trigger_deep_search_now'] = False
        trigger_auto_start_search = True

    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        seed_channel_input = st.text_input("Nhập Handle Kênh Mồi (ví dụ: @NickDiGiovanni):", key="seed_input_tab3")
        
        if st.button("✨ Tự Động Phân Tích từ Kênh Mồi"):
            pure_s_auto = to_pure_id(seed_channel_input)
            if pure_s_auto:
                try:
                    cid_auto = get_channel_id_by_handle(pure_s_auto)
                    if cid_auto:
                        ext = extract_channel_master_keywords(cid_auto)
                        st.session_state['pending_keywords'] = ", ".join(ext['master_keywords'][:6])
                        st.rerun()
                except Exception as e: st.error(f"Lỗi: {e}")
                    
        custom_keywords_input = st.text_input("Từ khóa chủ đề (Tự động liên kết):", key="custom_kw_tab3")
        
    with col_f2:
        min_subs_choice = st.selectbox("Mốc Subscribers Tối Thiểu:", options=[100000, 250000, 500000, 1000000], index=3, format_func=lambda x: f"{x:,} Subs")
        min_duration_choice = st.selectbox("Lọc Loại Bỏ Kênh Shorts:", options=[60, 180, 300, 600], index=0, format_func=lambda x: f"Loại Shorts < {x//60} phút" if x < 600 else "Có Video > 10 phút")

    start_btn = st.button("🚀 Bắt Đầu Săn Kênh Đồng Ngách", type="primary")

    if (start_btn or trigger_auto_start_search) and seed_channel_input:
        pure_seed = to_pure_id(seed_channel_input)
        try:
            st.info(f"🔍 Đang kết nối API và phân tích `{pure_seed}`...")
            seed_id = get_channel_id_by_handle(pure_seed)
            if not seed_id: st.error("Không tìm thấy kênh mồi này trên YouTube!")
            else:
                playlist_id, _, seed_desc, _, _, _, _ = get_channel_details(seed_id)
                
                if custom_keywords_input: top_kw_list = clean_and_extract_keywords(custom_keywords_input, seed_handle=pure_seed)
                else:
                    ext_info = extract_channel_master_keywords(seed_id)
                    top_kw_list = ext_info['master_keywords'][:4] if ext_info['master_keywords'] else [pure_seed.replace('_', ' ')]
                    
                st.write(f"🏷️ **Từ khóa quét:** `{', '.join(top_kw_list)}`")
                st.info("🌐 Đang quét đa luồng tự động hàng trăm ứng viên...")
                
                candidate_channel_ids = set()
                q_chan = " ".join(top_kw_list[:2])
                c_search_res = yt_execute(lambda yt: yt.search().list(part="snippet", q=q_chan, type="channel", maxResults=50))
                for c_item in c_search_res.get('items', []):
                    if c_item['snippet']['channelId'] != seed_id: candidate_channel_ids.add(c_item['snippet']['channelId'])
                    
                search_queries = [" ".join(top_kw_list[:2]), " ".join(top_kw_list[2:4])] if len(top_kw_list) >= 4 else [" ".join(top_kw_list)]
                for q in search_queries:
                    if not q.strip(): continue
                    v_search_res = yt_execute(lambda yt: yt.search().list(part="snippet", q=q, type="video", maxResults=50))
                    for v_item in v_search_res.get('items', []):
                        if v_item['snippet']['channelId'] != seed_id: candidate_channel_ids.add(v_item['snippet']['channelId'])
                        
                candidate_ids_list = list(candidate_channel_ids)
                
                if not candidate_ids_list: st.warning("Không quét được ứng viên nào!")
                else:
                    st.info(f"📊 Đang phân tích siêu tốc {len(candidate_ids_list)} kênh qua luồng xử lý song song...")
                    passed_channels, rejected_channels = [], []
                    channel_items, candidate_handles = [], []
                    
                    for i in range(0, len(candidate_ids_list), 50):
                        chan_res = yt_execute(lambda yt: yt.channels().list(part="snippet,contentDetails,statistics", id=','.join(candidate_ids_list[i:i+50])))
                        for item in chan_res.get('items', []):
                            c_h = to_pure_id(item['snippet'].get('customUrl', '')) or item['id'].lower()
                            candidate_handles.append(c_h)
                            channel_items.append(item)

                    db_res = supabase.table("channels").select("handle").in_("handle", candidate_handles).execute()
                    db_existing_set = {r["handle"].lower() for r in db_res.data} if db_res.data else set()
                    
                    with ThreadPoolExecutor(max_workers=8) as executor:
                        futures = [executor.submit(process_single_candidate, item, min_subs_choice, min_duration_choice, db_existing_set) for item in channel_items]
                        for future in as_completed(futures):
                            is_pass, res_data = future.result()
                            if is_pass:
                                passed_channels.append(res_data)
                            else:
                                rejected_channels.append(res_data)

                    st.session_state['passed_channels'] = passed_channels
                    st.session_state['rejected_channels'] = rejected_channels

                    st.divider()
                    c_m1, c_m2, c_m3 = st.columns(3)
                    c_m1.metric("Tổng ứng viên", len(candidate_ids_list))
                    c_m2.metric(f"✅ Đạt Chuẩn (>{min_subs_choice:,} Subs)", len(passed_channels))
                    c_m3.metric("❌ Bị Loại", len(rejected_channels))

        except Exception as e:
            st.error(f"Lỗi: {e}")

    if 'passed_channels' in st.session_state or 'rejected_channels' in st.session_state:
        passed_list = st.session_state.get('passed_channels', [])
        rejected_list = st.session_state.get('rejected_channels', [])
        
        tab_pass, tab_rej = st.tabs([f"✅ Kênh Đạt Chuẩn ({len(passed_list)})", f"❌ Kênh Bị Loại ({len(rejected_list)})"])
        
        # --- TAB PASSED ---
        with tab_pass:
            if passed_list:
                if st.button("🛒 Thêm TẤT CẢ Kênh Mới vào Giỏ", type="primary"):
                    for row in passed_list:
                        if "✅" in row["Trạng Thái DB"]: 
                            p_id = to_pure_id(row["Handle"])
                            st.session_state['cart'][p_id] = dict(row)
                            add_to_cart_db(p_id, dict(row))
                    st.rerun()
                st.divider()

                for idx, row in enumerate(passed_list):
                    p_id = to_pure_id(row['Handle'])
                    is_active = (p_id == st.session_state.get('active_inspected_handle'))
                    is_in_cart = p_id in st.session_state['cart']

                    with st.container(border=True):
                        if is_active:
                            st.markdown('<div class="active-card-marker"></div>', unsafe_allow_html=True)
                            st.markdown('<div style="background-color: #47A5D1; color: white; padding: 6px 12px; border-radius: 6px; font-weight: bold; margin-bottom: 10px;">🔍 ĐANG XEM 6 VIDEO MỚI CỦA KÊNH NÀY</div>', unsafe_allow_html=True)
                        elif is_in_cart:
                            st.markdown('<div class="in-cart-marker"></div>', unsafe_allow_html=True)
                            st.markdown('<div style="background-color: #10B981; color: white; padding: 6px 12px; border-radius: 6px; font-weight: bold; margin-bottom: 10px;">🛒 ĐÃ CÓ TRONG GIỎ HÀNG</div>', unsafe_allow_html=True)

                        c1, c2, c3, c4 = st.columns([2.2, 3.0, 1.8, 3.0])
                        with c1:
                            st.markdown(f"### [{row['Handle']}]({row['Link Kênh']})")
                            st.write(f"**{row['Tên Kênh']}**")
                            if st.button("👁️ Xem 6 Video Mới", key=f"btn_prev_pass_{p_id}", on_click=set_active_inspected_channel, args=(p_id,)):
                                show_video_dialog(p_id, pre_fetched_videos=row.get('recent_videos'))
                        with c2:
                            st.write(f"👥 **Subs:** `{row['Subscribers']}` | 🌍 **Q.Gia:** `{row['Quốc gia']}`")
                            st.write(f"🎬 **Tổng Video:** `{row['Tổng Số Video']}` | 📅 **Mới nhất:** `{row['Video Gần Nhất']}`")
                        with c3:
                            st.write(f"**Database:**\n{row['Trạng Thái DB']}")
                        with c4:
                            bc1, bc2 = st.columns(2)
                            if is_in_cart:
                                if bc1.button("❌ Bỏ Giỏ", key=f"rm_p_{p_id}", use_container_width=True):
                                    remove_from_cart_db(p_id)
                                    del st.session_state['cart'][p_id]
                                    st.rerun()
                            else:
                                if bc1.button("🛒 Thêm Giỏ", key=f"add_p_{p_id}", use_container_width=True):
                                    st.session_state['cart'][p_id] = dict(row)
                                    add_to_cart_db(p_id, dict(row))
                                    st.rerun()
                                    
                            audit_key = f"audit_file_{p_id}"
                            if audit_key in st.session_state:
                                bc2.download_button("📥 Tải Audit", data=st.session_state[audit_key]["bytes"], file_name=st.session_state[audit_key]["filename"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_p_{p_id}", use_container_width=True)
                            else:
                                if bc2.button("📄 Tạo Audit", key=f"btn_p_{p_id}", use_container_width=True):
                                    with st.spinner("Đang dựng Audit..."):
                                        b_data, f_name = run_single_channel_audit(p_id)
                                        if b_data:
                                            supabase.table("channels").upsert([{"handle": p_id, "youtuber_name": row['Tên Kênh'], "source": "Smart Finder Audit"}], on_conflict="handle").execute()
                                            st.session_state['audit_success_msg'] = f"🎉 Đã lưu **@{p_id}** vào Database!"
                                            st.session_state[audit_key] = {"bytes": b_data, "filename": f_name}
                                            st.rerun()

                            bc3, bc4 = st.columns(2)
                            if bc3.button("🎯 Đào Sâu", key=f"deep_p_{p_id}", type="secondary", use_container_width=True):
                                cid_deep = get_channel_id_by_handle(p_id)
                                if cid_deep:
                                    ext_deep = extract_channel_master_keywords(cid_deep)
                                    st.session_state['pending_keywords'] = ", ".join(ext_deep['master_keywords'][:6])
                                    st.session_state['pending_seed_input'] = f"@{p_id}"
                                    st.session_state['trigger_deep_search_now'] = True
                                    st.rerun()

                            if bc4.button("🗑️ Xóa", key=f"del_p_{p_id}", use_container_width=True, help="Loại bỏ kênh này khỏi hệ thống & danh sách"):
                                delete_channel_from_system(p_id)
                                st.toast(f"🗑️ Đã xóa kênh @{p_id}!")
                                st.rerun()
            else:
                st.info("Không có kênh nào đạt chuẩn.")
                
        # --- TAB REJECTED ---
        with tab_rej:
            if rejected_list:
                for idx, row in enumerate(rejected_list):
                    p_id = to_pure_id(row['Handle'])
                    is_active = (p_id == st.session_state.get('active_inspected_handle'))
                    is_in_cart = p_id in st.session_state['cart']
                    
                    with st.container(border=True):
                        if is_active:
                            st.markdown('<div class="active-card-marker"></div>', unsafe_allow_html=True)
                            st.markdown('<div style="background-color: #47A5D1; color: white; padding: 6px 12px; border-radius: 6px; font-weight: bold; margin-bottom: 10px;">🔍 ĐANG XEM 6 VIDEO MỚI CỦA KÊNH NÀY</div>', unsafe_allow_html=True)
                        elif is_in_cart:
                            st.markdown('<div class="in-cart-marker"></div>', unsafe_allow_html=True)
                            st.markdown('<div style="background-color: #10B981; color: white; padding: 6px 12px; border-radius: 6px; font-weight: bold; margin-bottom: 10px;">🛒 ĐÃ CÓ TRONG GIỎ HÀNG</div>', unsafe_allow_html=True)

                        c1, c2, c3, c4 = st.columns([2.2, 3.0, 1.8, 3.0])
                        with c1:
                            st.markdown(f"### [{row['Handle']}]({row['Link Kênh']})")
                            st.write(f"**{row['Tên Kênh']}**")
                            if st.button("👁️ Xem 6 Video Mới", key=f"btn_prev_rej_{p_id}", on_click=set_active_inspected_channel, args=(p_id,)):
                                show_video_dialog(p_id, pre_fetched_videos=row.get('recent_videos'))
                        with c2:
                            st.write(f"👥 **Subs:** `{row['Subscribers']}` | 🌍 **Q.Gia:** `{row.get('Quốc gia', '')}`")
                            st.write(f"🎬 **Tổng Video:** `{row.get('Tổng Số Video', '')}` | 📅 **Mới nhất:** `{row.get('Video Gần Nhất', '')}`")
                        with c3:
                            st.write(f"**Database:** {row.get('Trạng Thái DB', '')}")
                            st.write(f"❌ **Lý do:** `{row['Lý do loại']}`")
                        with c4:
                            bc1, bc2 = st.columns(2)
                            if is_in_cart:
                                if bc1.button("❌ Bỏ Giỏ", key=f"rm_r_{p_id}", use_container_width=True):
                                    remove_from_cart_db(p_id)
                                    del st.session_state['cart'][p_id]
                                    st.rerun()
                            else:
                                if bc1.button("🛒 Thêm Giỏ", key=f"add_r_{p_id}", use_container_width=True):
                                    st.session_state['cart'][p_id] = dict(row)
                                    add_to_cart_db(p_id, dict(row))
                                    st.rerun()

                            audit_key = f"audit_file_{p_id}"
                            if audit_key in st.session_state:
                                bc2.download_button("📥 Tải Audit", data=st.session_state[audit_key]["bytes"], file_name=st.session_state[audit_key]["filename"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_r_{p_id}", use_container_width=True)
                            else:
                                if bc2.button("📄 Tạo Audit", key=f"btn_r_{p_id}", use_container_width=True):
                                    with st.spinner("Đang dựng Audit..."):
                                        b_data, f_name = run_single_channel_audit(p_id)
                                        if b_data:
                                            supabase.table("channels").upsert([{"handle": p_id, "youtuber_name": row['Tên Kênh'], "source": "Smart Finder Audit"}], on_conflict="handle").execute()
                                            st.session_state['audit_success_msg'] = f"🎉 Đã lưu **@{p_id}** vào Database!"
                                            st.session_state[audit_key] = {"bytes": b_data, "filename": f_name}
                                            st.rerun()

                            bc3, bc4 = st.columns(2)
                            if bc3.button("🎯 Đào Sâu", key=f"deep_r_{p_id}", type="secondary", use_container_width=True):
                                cid_deep = get_channel_id_by_handle(p_id)
                                if cid_deep:
                                    ext_deep = extract_channel_master_keywords(cid_deep)
                                    st.session_state['pending_keywords'] = ", ".join(ext_deep['master_keywords'][:6])
                                    st.session_state['pending_seed_input'] = f"@{p_id}"
                                    st.session_state['trigger_deep_search_now'] = True
                                    st.rerun()

                            if bc4.button("🗑️ Xóa", key=f"del_r_{p_id}", use_container_width=True, help="Loại bỏ kênh này khỏi hệ thống & danh sách"):
                                delete_channel_from_system(p_id)
                                st.toast(f"🗑️ Đã xóa kênh @{p_id}!")
                                st.rerun()

    render_shared_cart_ui(key_suffix="tab3")

# --- TAB 4, TAB 5, TAB 6 ---
with tab4:
    st.subheader("Upload file .ZIP hoặc .TXT để cập nhật Database")
    uploaded_files = st.file_uploader("Kéo thả file `.zip` (chứa các báo cáo Excel) hoặc file `.txt` vào đây:", type=["zip", "txt", "xlsx"], accept_multiple_files=True)
    if uploaded_files and st.button("🚀 Bắt đầu xử lý & Nạp vào Database", type="primary"):
        new_handles_to_insert = []
        for file in uploaded_files:
            file_name = file.name
            if file_name.endswith('.zip'):
                with zipfile.ZipFile(file, 'r') as zip_ref:
                    extract_path = "temp_zip_extract"
                    zip_ref.extractall(extract_path)
                    for root, _, filenames in os.walk(extract_path):
                        for fn in filenames:
                            if fn.endswith('.xlsx') or fn.endswith('.xls'):
                                h = extract_handle_from_filename(fn)
                                if h: new_handles_to_insert.append({"handle": h, "youtuber_name": h, "source": file_name})
            elif file_name.endswith('.txt'):
                content = file.read().decode("utf-8", errors="ignore")
                for line in content.splitlines():
                    h = to_pure_id(line)
                    if h: new_handles_to_insert.append({"handle": h, "youtuber_name": h, "source": file_name})
        if new_handles_to_insert:
            df_insert = pd.DataFrame(new_handles_to_insert).drop_duplicates(subset=["handle"])
            supabase.table("channels").upsert(df_insert.to_dict(orient="records"), on_conflict="handle").execute()
            st.success(f"🎉 Đã xử lý & đồng bộ thành công {len(df_insert)} Handle vào Database đám mây!")

with tab5:
    st.subheader("📊 Danh sách toàn bộ Channel trong Database")
    res = supabase.table("channels").select("*").execute()
    if res.data:
        df_all = pd.DataFrame(res.data)
        st.write(f"Tổng số kênh hiện có: **{len(df_all)}**")
        
        search_db = st.text_input("🔍 Tìm kiếm kênh trong Database (Handle hoặc Tên):", "")
        if search_db:
            df_filtered = df_all[
                df_all['handle'].str.contains(search_db, case=False, na=False) | 
                df_all['youtuber_name'].str.contains(search_db, case=False, na=False)
            ]
        else:
            df_filtered = df_all

        items_per_page = 20
        total_pages = max(1, (len(df_filtered) + items_per_page - 1) // items_per_page)
        page = st.number_input("Trang:", min_value=1, max_value=total_pages, value=1, step=1)
        
        start_idx = (int(page) - 1) * items_per_page
        end_idx = start_idx + items_per_page
        page_data = df_filtered.iloc[start_idx:end_idx]

        st.divider()
        for idx, row in page_data.iterrows():
            p_id = to_pure_id(row['handle'])
            is_active = (p_id == st.session_state.get('active_inspected_handle'))
            
            with st.container(border=True):
                if is_active:
                    st.markdown('<div class="active-card-marker"></div>', unsafe_allow_html=True)
                    st.markdown('<div style="background-color: #47A5D1; color: white; padding: 6px 12px; border-radius: 6px; font-weight: bold; margin-bottom: 10px;">🔍 ĐANG XEM 6 VIDEO MỚI CỦA KÊNH NÀY</div>', unsafe_allow_html=True)

                c1, c2, c3 = st.columns([4.0, 4.0, 2.0])
                with c1:
                    st.markdown(f"### [@{p_id}](https://youtube.com/@{p_id})")
                    st.write(f"**Tên YouTuber:** {row.get('youtuber_name', 'N/A')}")
                    if st.button("👁️ Xem 6 Video Mới", key=f"btn_prev_db_{idx}_{p_id}", on_click=set_active_inspected_channel, args=(p_id,)):
                        show_video_dialog(p_id)
                with c2:
                    st.write(f"**Nguồn dữ liệu:** {row.get('source', 'N/A')}")
                with c3:
                    if st.button("🗑️ Xóa DB", key=f"del_db_{idx}_{p_id}", use_container_width=True):
                        delete_channel_from_system(p_id)
                        st.toast(f"🗑️ Đã xóa kênh @{p_id} khỏi Database!")
                        st.rerun()

        st.divider()
        csv = df_all.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Tải về toàn bộ Database (CSV)", data=csv, file_name="master_youtube_database.csv", mime="text/csv")

with tab6:
    st.subheader("✨ Soi Từ Khóa Kênh (Channel & Video Tags SEO Inspector)")
    inspect_handle_input = st.text_input("Nhập Handle Kênh cần soi:", value="@NickDiGiovanni")
    if inspect_handle_input and st.button("🔍 Soi Từ Khóa Ngay", type="primary"):
        pure_inspect = to_pure_id(inspect_handle_input)
        if pure_inspect:
            try:
                cid_insp = get_channel_id_by_handle(pure_inspect)
                if not cid_insp:
                    st.error("Không tìm thấy Channel ID cho kênh này!")
                else:
                    ext_data = extract_channel_master_keywords(cid_insp)
                    st.session_state['pending_keywords'] = ", ".join(ext_data['master_keywords'])
                    st.session_state['last_inspected_data'] = ext_data
                    st.session_state['last_inspected_handle'] = pure_inspect
                    st.rerun()
            except Exception as e:
                st.error(f"Lỗi khi soi từ khóa: {e}")

    if 'last_inspected_data' in st.session_state:
        ext_data, pure_inspect = st.session_state['last_inspected_data'], st.session_state.get('last_inspected_handle', '')
        st.divider(); st.success(f"✨ Đã liên kết tự động bộ từ khóa này sang Tab 3 ('Săn Kênh Tương Tự')!")
        st.markdown(f"### 🏷️ Dữ Liệu Từ Khóa Của Kênh `@{pure_inspect}`")
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            st.markdown("#### 🔑 Thẻ Từ Khóa Ẩn (Channel Keywords):")
            for kw in ext_data['channel_keywords']: st.write(f"• `{kw}`")
        with col_t2:
            st.markdown("#### 📌 Top Video Tags Xuất Hiện Nhiều Nhất:")
            for tag in ext_data['top_tags']: st.write(f"• `{tag}`")
        st.code(", ".join(ext_data['master_keywords']), language="text")
