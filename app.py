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
import xml.etree.ElementTree as ET
from collections import Counter
from PIL import Image as PILImage
from supabase import create_client, Client
from concurrent.futures import ThreadPoolExecutor, as_completed

# Page Config
st.set_page_config(
    page_title="YT CHECKER PRO", 
    page_icon="🎙️", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Connect to Supabase
@st.cache_resource
def init_supabase():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = init_supabase()

# --- INITIALIZE PERSISTENT STATE ---
if 'app_theme' not in st.session_state:
    st.session_state['app_theme'] = 'Studio Peach (Sáng)'

if 'selected_channels' not in st.session_state:
    st.session_state['selected_channels'] = set()

# Callback for Selection Sync
def toggle_select_channel(pure_handle):
    if pure_handle in st.session_state['selected_channels']:
        st.session_state['selected_channels'].remove(pure_handle)
    else:
        st.session_state['selected_channels'].add(pure_handle)

def clear_selected_channels():
    st.session_state['selected_channels'].clear()

# Theme CSS Dynamic Injection
is_dark = st.session_state['app_theme'] == 'Studio Espresso (Tối)'
bg_color = "#1E1816" if is_dark else "#F4F2F1"
card_bg = "#2A221F" if is_dark else "#FFFFFF"
text_color = "#F4F2F1" if is_dark else "#3D2F29"
border_color = "#3D2F29" if is_dark else "#E5E7EB"
sidebar_bg = "#241D1A" if is_dark else "#FFFFFF"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap');

/* Base Theme */
.stApp {{
    background-color: {bg_color} !important;
    color: {text_color} !important;
    font-family: 'Montserrat', sans-serif !important;
}}

/* Sidebar Styling */
section[data-testid="stSidebar"] {{
    background-color: {sidebar_bg} !important;
    border-right: 1px solid {border_color} !important;
    box-shadow: 4px 0 15px rgba(0, 0, 0, 0.05) !important;
}}

/* Header */
header[data-testid="stHeader"] {{
    background-color: transparent !important;
}}

/* HIGH-END ARTISTIC TABS */
.stTabs [data-baseweb="tab-list"] {{
    gap: 32px;
    background-color: transparent;
    padding: 0 0 4px 0;
    border-bottom: 2px solid #D1D5DB;
}}

.stTabs [data-baseweb="tab"] {{
    background-color: transparent !important;
    border: none !important;
    border-bottom: 3px solid transparent !important;
    border-radius: 0 !important;
    color: #6B7280 !important;
    font-weight: 700;
    font-size: 0.9rem;
    padding: 10px 4px;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    transition: all 0.3s ease !important;
    cursor: pointer !important;
}}

.stTabs [data-baseweb="tab"]:hover {{
    color: #D95F26 !important;
    transform: translateY(-1px);
}}

.stTabs [aria-selected="true"] {{
    color: #D95F26 !important;
    border-bottom: 3px solid #D95F26 !important;
    transform: translateY(0);
}}

/* Standard Card Container Styling */
div[data-testid="stVerticalBlockBorderWrapper"] {{
    background-color: {card_bg} !important;
    border: 1px solid {border_color} !important;
    border-radius: 12px !important;
    padding: 12px !important;
    box-shadow: 0 2px 10px rgba(0, 0, 0, 0.03) !important;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}}

div[data-testid="stVerticalBlockBorderWrapper"]:hover {{
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08) !important;
}}

/* ACTIVE INSPECTED CARD (TANGERINE) */
div[data-testid="stVerticalBlockBorderWrapper"]:has(div.active-card-marker) {{
    border: 2px solid #D95F26 !important; 
    box-shadow: 0 8px 24px rgba(217, 95, 38, 0.2) !important;
}}

/* IN-CART CARD (OCEAN BLUE) */
div[data-testid="stVerticalBlockBorderWrapper"]:has(div.in-cart-marker) {{
    border: 2px solid #47A5D1 !important;
    box-shadow: 0 8px 24px rgba(71, 165, 209, 0.2) !important;
}}

/* Active Banner Tag Styling (Strict 100% White Text) */
.active-banner-tag {{
    background-color: #D95F26 !important;
    color: #FFFFFF !important;
    padding: 6px 14px !important;
    border-radius: 8px !important;
    font-weight: 800 !important;
    margin-bottom: 12px !important;
    font-size: 0.85rem !important;
    letter-spacing: 0.05em !important;
    display: inline-block !important;
    box-shadow: 0 3px 10px rgba(217, 95, 38, 0.25) !important;
}}
.active-banner-tag * {{ color: #FFFFFF !important; }}

.in-cart-banner-tag {{
    background-color: #47A5D1 !important;
    color: #FFFFFF !important;
    padding: 6px 14px !important;
    border-radius: 8px !important;
    font-weight: 800 !important;
    margin-bottom: 12px !important;
    font-size: 0.85rem !important;
    letter-spacing: 0.05em !important;
    display: inline-block !important;
    box-shadow: 0 3px 10px rgba(71, 165, 209, 0.25) !important;
}}
.in-cart-banner-tag * {{ color: #FFFFFF !important; }}

/* Inputs & Selectboxes */
.stTextInput input, .stTextArea textarea, .stSelectbox select {{
    background-color: {card_bg} !important;
    color: {text_color} !important;
    border: 1px solid #D1D5DB !important;
    border-radius: 8px !important;
    font-family: 'Montserrat', sans-serif !important;
}}

.stTextInput input:focus, .stTextArea textarea:focus, .stSelectbox select:focus {{
    border-color: #D95F26 !important;
    box-shadow: 0 0 0 1px #D95F26 !important;
}}

/* Default Buttons */
.stButton button {{
    border-radius: 8px !important;
    font-weight: 700 !important;
    font-family: 'Montserrat', sans-serif !important;
    border: 1px solid #D1D5DB !important;
    background-color: {card_bg} !important;
    color: {text_color} !important;
    text-transform: uppercase;
    font-size: 0.8rem !important;
    letter-spacing: 0.05em;
    transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
}}

.stButton button:hover {{
    border-color: #D95F26 !important;
    color: #D95F26 !important;
    transform: translateY(-1px) !important;
}}

/* ARTISTIC PRIMARY BUTTONS (#D95F26 Gradient & Soft 1px Lift) */
.stButton button[kind="primary"],
.stButton button[kind="primary"] *,
.stButton button[kind="primary"] p,
.stButton button[kind="primary"] span,
.stButton button[kind="primary"] div {{
    background: linear-gradient(135deg, #D95F26 0%, #E66A32 100%) !important;
    color: #FFFFFF !important;
    border: none !important;
    box-shadow: 0 3px 10px rgba(217, 95, 38, 0.22) !important;
}}

.stButton button[kind="primary"]:hover,
.stButton button[kind="primary"]:hover *,
.stButton button[kind="primary"]:hover p,
.stButton button[kind="primary"]:hover span,
.stButton button[kind="primary"]:hover div {{
    background: linear-gradient(135deg, #C24E18 0%, #D95F26 100%) !important;
    color: #FFFFFF !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 16px rgba(217, 95, 38, 0.32) !important;
}}

/* Badges */
.badge-pro {{
    display: inline-block;
    padding: 6px 12px;
    border-radius: 9999px;
    font-size: 0.7rem;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 0.08em;
}}

.badge-tangerine, .badge-tangerine * {{ background-color: #D95F26 !important; color: #FFFFFF !important; border: none !important; }}
.badge-ocean, .badge-ocean * {{ background-color: #47A5D1 !important; color: #FFFFFF !important; border: none !important; }}
</style>
""", unsafe_allow_html=True)

# Global Default API Key
DEFAULT_API_KEY = st.secrets.get("YOUTUBE_API_KEY", "AIzaSyDDBEJscqkGGpG1xtuL4wYPuFkS4BIL854")

# --- DATABASE PERSISTENCE HELPERS ---
def load_api_keys_from_db():
    try:
        res = supabase.table("app_config").select("value").eq("key", "api_keys").execute()
        if res.data and len(res.data) > 0: return res.data[0]["value"]
    except Exception: pass
    return None

def save_api_keys_to_db(key_string):
    try:
        supabase.table("app_config").upsert({"key": "api_keys", "value": key_string}, on_conflict="key").execute()
        return True
    except Exception: return False

def load_cart_from_db():
    cart_dict = {}
    try:
        res = supabase.table("cart_items").select("*").execute()
        if res.data:
            for row in res.data:
                h = row["handle"]
                c_data = row.get("channel_data")
                if isinstance(c_data, str): c_data = json.loads(c_data)
                cart_dict[h] = c_data
    except Exception: pass
    return cart_dict

def add_to_cart_db(pure_handle, channel_data):
    try:
        data_clean = dict(channel_data)
        if "recent_videos" in data_clean: del data_clean["recent_videos"]
        supabase.table("cart_items").upsert({"handle": pure_handle, "channel_data": data_clean}, on_conflict="handle").execute()
    except Exception: pass

def remove_from_cart_db(pure_handle):
    try: supabase.table("cart_items").delete().eq("handle", pure_handle).execute()
    except Exception: pass

def clear_cart_db():
    try: supabase.table("cart_items").delete().neq("handle", "___NONE___").execute()
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

if 'seed_input_tab3' not in st.session_state: st.session_state['seed_input_tab3'] = "@NickDiGiovanni"
if 'pending_keywords' in st.session_state:
    st.session_state['custom_kw_tab3'] = st.session_state['pending_keywords']
    del st.session_state['pending_keywords']

if 'custom_kw_tab3' not in st.session_state: st.session_state['custom_kw_tab3'] = ""
if 'video_preview_cache' not in st.session_state: st.session_state['video_preview_cache'] = {}
if 'active_inspected_handle' not in st.session_state: st.session_state['active_inspected_handle'] = None

def set_active_inspected_channel(pure_handle):
    st.session_state['active_inspected_handle'] = pure_handle

def set_api_keys(key_string):
    keys = [k.strip() for k in re.split(r'[\n,]+', key_string) if k.strip()]
    st.session_state['api_keys'] = keys if keys else [DEFAULT_API_KEY]

set_api_keys(st.session_state['global_api_keys'])

# --- CLEAN & ROBUST CUSTOM KPI RENDERER ---
def render_kpi_cards(kpi_data):
    cards_html = ""
    for title, val, color in kpi_data:
        cards_html += f'<div style="flex: 1; min-width: 200px; background-color: {card_bg}; border: 1px solid {border_color}; border-radius: 12px; padding: 18px; text-align: center; box-shadow: 0 4px 10px rgba(0,0,0,0.04);"><div style="font-size: 0.8rem; color: #6B7280; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 6px;">{title}</div><div style="font-size: 2.2rem; font-weight: 900; color: {color};">{val}</div></div>'
    full_html = f'<div style="display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap;">{cards_html}</div>'
    st.markdown(full_html, unsafe_allow_html=True)

# --- SIDEBAR BRANDING & CONFIG ---
with st.sidebar:
    col1, col2, col3 = st.columns([0.5, 8, 0.5])
    with col2:
        logo_paths = ["logo.png", "logo_2.png", "logo.jpg"]
        found_logo = None
        for path in logo_paths:
            if os.path.exists(path):
                found_logo = path
                break
        if found_logo: st.image(found_logo, use_container_width=True)
        else:
            st.markdown("""
            <div style="text-align: center; padding: 10px 0;">
                <svg width="65" height="65" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <rect x="10" y="45" width="8" height="10" rx="2" fill="#D95F26"/>
                    <rect x="22" y="30" width="8" height="40" rx="2" fill="#D95F26"/>
                    <rect x="34" y="15" width="8" height="70" rx="2" fill="#D95F26"/>
                    <rect x="46" y="5" width="8" height="90" rx="2" fill="#D95F26"/>
                    <rect x="58" y="20" width="8" height="60" rx="2" fill="#D95F26"/>
                    <rect x="70" y="35" width="8" height="30" rx="2" fill="#D95F26"/>
                    <rect x="82" y="42" width="8" height="16" rx="2" fill="#D95F26"/>
                </svg>
            </div>
            """, unsafe_allow_html=True)
            
    st.markdown("""
        <div style="text-align: center; padding-bottom: 8px;">
            <h2 style="margin: 0 0 4px 0; font-weight: 800; font-size: 1.15rem; letter-spacing: -0.02em;">YT CHECKER PRO</h2>
            <span class="badge-pro badge-ocean">Supabase Live</span>
        </div>
    """, unsafe_allow_html=True)
    
    st.selectbox("🎨 Giao diện App:", options=["Studio Peach (Sáng)", "Studio Espresso (Tối)"], key="app_theme")
    st.divider()

    st.markdown("<h4 style='font-weight: 700; font-size: 0.95rem;'>⚙️ Cấu Hình API Keys</h4>", unsafe_allow_html=True)
    active_key_count = len(st.session_state.get('api_keys', []))
    st.markdown(f"Đang chạy: <span class='badge-pro badge-tangerine'>{active_key_count} Keys</span>", unsafe_allow_html=True)
    
    keys_input = st.text_area("Danh sách API (1 key/dòng):", value=st.session_state['global_api_keys'], height=140, key="api_keys_text_area")
    
    if st.button("💾 Lưu Cấu Hình Vĩnh Viễn", type="primary", use_container_width=True):
        st.session_state['global_api_keys'] = keys_input
        set_api_keys(keys_input)
        save_api_keys_to_db(keys_input)
        st.toast("🎉 Đã lưu vĩnh viễn danh sách API Keys!")
        st.rerun()

    st.divider()
    if st.button("🔄 Làm Mới Màn Hình", use_container_width=True):
        keys_to_clear = ['passed_channels', 'rejected_channels', 'last_inspected_data', 'last_inspected_handle', 'audit_success_msg', 'batch_check_new', 'batch_check_existing', 'active_inspected_handle', 'selected_channels']
        for key in keys_to_clear:
            if key in st.session_state: del st.session_state[key]
        for key in list(st.session_state.keys()):
            if key.startswith('audit_file_'): del st.session_state[key]
        st.rerun()

# --- HELPER TO DELETE CHANNEL COMPLETELY FROM SYSTEM ---
def delete_channel_from_system(pure_handle):
    if not pure_handle: return
    try: supabase.table("channels").delete().eq("handle", pure_handle).execute()
    except Exception: pass

    remove_from_cart_db(pure_handle)
    if pure_handle in st.session_state.get('cart', {}): del st.session_state['cart'][pure_handle]
    if st.session_state.get('active_inspected_handle') == pure_handle: st.session_state['active_inspected_handle'] = None
    if pure_handle in st.session_state['selected_channels']: st.session_state['selected_channels'].remove(pure_handle)

    for key_list in ['passed_channels', 'rejected_channels', 'batch_check_new', 'batch_check_existing']:
        if key_list in st.session_state:
            st.session_state[key_list] = [ch for ch in st.session_state[key_list] if to_pure_id(ch.get('Handle')) != pure_handle]

    audit_key = f"audit_file_{pure_handle}"
    if audit_key in st.session_state: del st.session_state[audit_key]

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
            else: raise e
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
    except Exception as e: st.error(f"Lỗi đọc file: {e}")
    return handles

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

# --- YOUTUBE API OPERATIONS ---
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

# --- ULTRA-FAST 0-QUOTA RSS PREVIEW FETCH ---
@st.cache_data(ttl=43200, show_spinner=False)
def get_6_recent_videos(pure_handle):
    long_vids = []
    try:
        cid = get_channel_id_by_handle(pure_handle)
        if cid:
            rss_url = f"https://www.youtube.com/feeds/videos.xml?channel_id={cid}"
            resp = requests.get(rss_url, timeout=4)
            if resp.status_code == 200:
                root = ET.fromstring(resp.content)
                ns = {'atom': 'http://www.w3.org/2005/Atom', 'yt': 'http://www.youtube.com/xml/schemas/2015'}
                rss_v_ids = []
                for entry in root.findall('atom:entry', ns):
                    v_id_el = entry.find('yt:videoId', ns)
                    if v_id_el is not None and v_id_el.text: rss_v_ids.append(v_id_el.text)
                    if len(rss_v_ids) >= 12: break
                if rss_v_ids:
                    v_details = get_video_details(rss_v_ids)
                    for v in v_details:
                        if is_long_form_video(v, min_seconds=180): long_vids.append(v)
                        if len(long_vids) >= 6: break

            if len(long_vids) < 6:
                playlist_id, _, _, _, _, _, _ = get_channel_details(cid)
                if playlist_id:
                    v_res = yt_execute(lambda yt: yt.playlistItems().list(part="snippet", playlistId=playlist_id, maxResults=30))
                    v_ids = [v_item['snippet']['resourceId']['videoId'] for v_item in v_res.get('items', [])]
                    if v_ids:
                        v_details = get_video_details(v_ids)
                        for v in v_details:
                            if is_long_form_video(v, min_seconds=180) and v not in long_vids: long_vids.append(v)
                            if len(long_vids) >= 6: break
    except Exception: pass
    return long_vids[:6]

# --- STREAMLIT MODAL DIALOGS ---
@st.dialog("🎬 6 Video Dài (Long-form) Mới Nhất", width="large")
def show_video_dialog(pure_handle, pre_fetched_videos=None):
    st.markdown(f"<h3 style='font-weight: 800;'>📺 Kênh đang xem: <span style='color: #D95F26;'>@{pure_handle}</span></h3>", unsafe_allow_html=True)
    st.markdown(f"🔗 **[Mở thẳng Tab Videos trên YouTube](https://youtube.com/@{pure_handle}/videos)**")
    
    vids = []
    if pre_fetched_videos: vids = [v for v in pre_fetched_videos if is_long_form_video(v, min_seconds=180)]
    if len(vids) < 6: vids = get_6_recent_videos(pure_handle)
    vids = vids[:6]
    
    if vids:
        st.divider()
        for row_idx in range(0, len(vids), 2):
            col1, col2 = st.columns(2)
            v1 = vids[row_idx]
            with col1:
                vid_id1 = v1.get('Video ID') or (v1.get('Link', '').split('v=')[-1] if 'v=' in v1.get('Link', '') else '')
                if vid_id1: st.image(f"https://img.youtube.com/vi/{vid_id1}/hqdefault.jpg", use_container_width=True)
                st.markdown(f"**[{v1['Title'][:45]}...]({v1['Link']})**")
                st.caption(f"👀 {v1.get('Views', 0):,} views | ⏳ {v1.get('Length (Exact)', 'N/A')} | 📅 {v1.get('Published Date', '')}")
            
            if row_idx + 1 < len(vids):
                v2 = vids[row_idx + 1]
                with col2:
                    vid_id2 = v2.get('Video ID') or (v2.get('Link', '').split('v=')[-1] if 'v=' in v2.get('Link', '') else '')
                    if vid_id2: st.image(f"https://img.youtube.com/vi/{vid_id2}/hqdefault.jpg", use_container_width=True)
                    st.markdown(f"**[{v2['Title'][:45]}...]({v2['Link']})**")
                    st.caption(f"👀 {v2.get('Views', 0):,} views | ⏳ {v2.get('Length (Exact)', 'N/A')} | 📅 {v2.get('Published Date', '')}")
            st.divider()
    else: st.caption("Không tìm thấy video dài (Kênh này chỉ đăng Shorts hoặc chưa có video dài trên 3 phút).")
    if st.button("❌ Đóng Cửa Sổ Preview", type="primary", use_container_width=True): st.rerun()

def get_selected_channel_data():
    selected = st.session_state['selected_channels']
    data = []
    lists = st.session_state.get('batch_check_new', []) + st.session_state.get('batch_check_existing', []) + st.session_state.get('passed_channels', []) + st.session_state.get('rejected_channels', [])
    seen = set()
    for item in lists:
        p = to_pure_id(item.get('Handle'))
        if p in selected and p not in seen:
            data.append(item)
            seen.add(p)
    return data

@st.dialog("⚖️ Bảng So Sánh Kênh Trực Diện", width="large")
def compare_channels_dialog(channel_data_list):
    if not channel_data_list:
        st.warning("Không có dữ liệu kênh để so sánh.")
        return
    st.markdown("<h3 style='text-align: center; color: #D95F26; font-weight: 800; margin-bottom: 20px;'>📊 SO SÁNH CHỈ SỐ KÊNH</h3>", unsafe_allow_html=True)
    
    # ENRICH DATA LIVE IF METRICS ARE N/A (E.G. FROM TAB 1)
    enriched_list = []
    with st.spinner("Đang kết nối API cào dữ liệu chi tiết để so sánh..."):
        for ch in channel_data_list:
            c_dict = dict(ch)
            pure_h = to_pure_id(c_dict.get('Handle'))
            if pure_h and (c_dict.get('Subscribers') is None or c_dict.get('Subscribers') == 'N/A'):
                cid = get_channel_id_by_handle(pure_h)
                if cid:
                    playlist_id, sub_count, channel_desc, channel_joined, country_name, country_code, avatar_url = get_channel_details(cid)
                    recent_vids = get_6_recent_videos(pure_h)
                    latest_date = recent_vids[0]['Published Date'] if recent_vids else 'N/A'
                    
                    try:
                        c_res = yt_execute(lambda yt: yt.channels().list(part="statistics", id=cid))
                        video_count = int(c_res['items'][0]['statistics'].get('videoCount', 0)) if (c_res.get('items') and len(c_res['items']) > 0) else 0
                    except Exception: video_count = 0

                    c_dict['Subscribers'] = f"{sub_count:,}"
                    c_dict['Tổng Số Video'] = f"{video_count:,}"
                    c_dict['Quốc gia'] = country_name if country_name else 'N/A'
                    c_dict['Video Gần Nhất'] = latest_date
            enriched_list.append(c_dict)

    cols = st.columns(len(enriched_list))
    for idx, ch in enumerate(enriched_list):
        with cols[idx]:
            st.markdown(f"<div style='background-color: {card_bg}; padding: 15px; border-radius: 12px; border: 1px solid {border_color}; text-align: center;'>", unsafe_allow_html=True)
            st.markdown(f"<h4 style='color: #47A5D1; font-weight: 800; margin-bottom: 5px;'><a href='{ch.get('Link Kênh', f'https://youtube.com/@{to_pure_id(ch.get(\"Handle\"))}')}' style='text-decoration: none; color: inherit;'>{ch.get('Handle')}</a></h4>", unsafe_allow_html=True)
            st.markdown(f"<p style='font-size: 0.9rem; color: #6B7280; font-weight: 600;'>{ch.get('Tên Kênh', 'N/A')}</p>", unsafe_allow_html=True)
            st.divider()
            st.markdown(f"<p style='font-size: 0.8rem; color: #6B7280; margin-bottom: 0;'>👥 SUBSCRIBERS</p><p style='font-size: 1.5rem; font-weight: 800; color: #D95F26; margin-top: 0;'>{ch.get('Subscribers', 'N/A')}</p>", unsafe_allow_html=True)
            st.markdown(f"<p style='font-size: 0.8rem; color: #6B7280; margin-bottom: 0;'>🎬 TỔNG VIDEO</p><p style='font-size: 1.3rem; font-weight: 700; color: #3D2F29; margin-top: 0;'>{ch.get('Tổng Số Video', 'N/A')}</p>", unsafe_allow_html=True)
            st.markdown(f"<p style='font-size: 0.8rem; color: #6B7280; margin-bottom: 0;'>🌍 QUỐC GIA</p><p style='font-size: 1.1rem; font-weight: 600; color: #3D2F29; margin-top: 0;'>{ch.get('Quốc gia', 'N/A')}</p>", unsafe_allow_html=True)
            st.markdown(f"<p style='font-size: 0.8rem; color: #6B7280; margin-bottom: 0;'>📅 GẦN NHẤT</p><p style='font-size: 1.1rem; font-weight: 600; color: #47A5D1; margin-top: 0;'>{ch.get('Video Gần Nhất', 'N/A')}</p>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
    st.write("")
    if st.button("❌ Đóng Cửa Sổ So Sánh", type="primary", use_container_width=True): st.rerun()

def run_single_channel_audit(pure_handle):
    cid = get_channel_id_by_handle(pure_handle)
    if not cid: return None, None
    playlist_id, sub_count, channel_desc, channel_joined, channel_country, c_code, avatar_url = get_channel_details(cid)
    v_ids = []
    next_token = None
    while True:
        res = yt_execute(lambda yt: yt.playlistItems().list(part="snippet", playlistId=playlist_id, maxResults=50, pageToken=next_token))
        for v_item in res.get('items', []): v_ids.append(v_item['snippet']['resourceId']['videoId'])
        next_token = res.get('nextPageToken')
        if not next_token: break
    v_data = get_video_details(v_ids)
    excel_bytes = generate_v414_excel_report(pure_handle, sub_count, channel_desc, channel_joined, channel_country, avatar_url, v_data)
    out_fname = f"{pure_handle}_{datetime.datetime.now().strftime('%d-%m-%Y')}.xlsx"
    return excel_bytes, out_fname

def generate_v414_excel_report(clean_handle, sub_count, channel_desc, channel_joined, channel_country, avatar_url, video_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = clean_handle[:31]
    date_str = datetime.datetime.now().strftime("%d-%m-%Y")
    total_videos = len(video_data)
    total_views = sum(v['Views'] for v in video_data)
    total_minutes = round(sum(v['Seconds'] for v in video_data) / 60)
    ws.merge_cells('A1:E1')
    ws['A1'] = f"{clean_handle.upper()} YOUTUBE CHANNEL SUMMARY REPORT - up to {date_str}"
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="D95F26", end_color="D95F26", fill_type="solid")
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
        cell.fill = PatternFill(start_color="D1D5DB", end_color="D1D5DB", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[12].height = 24
    for idx, v in enumerate(video_data):
        r = idx + 13
        cA = ws.cell(row=r, column=1, value=v['Title']); cA.font = Font(name="Calibri", size=11)
        cB = ws.cell(row=r, column=2, value=v['Link'])
        if v.get('Link'): cB.hyperlink = v['Link']; cB.font = Font(name="Calibri", size=11, color="47A5D1", underline="single")
        ws.cell(row=r, column=3, value=v['Length (Exact)']).alignment = Alignment(horizontal="center", vertical="center")
        cD = ws.cell(row=r, column=4, value=v['Views']); cD.number_format = '#,##0'
        ws.cell(row=r, column=5, value=v['Published Date']).alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions['A'].width = 55; ws.column_dimensions['B'].width = 45; ws.column_dimensions['C'].width = 22; ws.column_dimensions['D'].width = 15; ws.column_dimensions['E'].width = 15
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# --- REUSABLE COMPONENT: RENDER SHARED CART ---
def render_shared_cart_ui(key_suffix=""):
    st.divider()
    cart_items = st.session_state['cart']
    st.markdown(f"<h3 style='font-weight: 800;'>🛒 Giỏ Hàng Dùng Chung ({len(cart_items)} Kênh)</h3>", unsafe_allow_html=True)
    if cart_items:
        df_cart = pd.DataFrame(list(cart_items.values()))
        if 'Handle' in df_cart.columns:
            df_cart['Tab Videos'] = df_cart['Handle'].apply(lambda h: f"https://youtube.com/{to_pure_id(h)}/videos")
            df_cart['Link Kênh'] = df_cart['Handle'].apply(lambda h: f"https://youtube.com/{to_pure_id(h)}")
            
        if 'recent_videos' in df_cart.columns: df_cart = df_cart.drop(columns=['recent_videos'])

        st.dataframe(df_cart, use_container_width=True, column_config={
            "Link Kênh": st.column_config.LinkColumn("Trang Chủ", display_text="🏠 Kênh"),
            "Tab Videos": st.column_config.LinkColumn("Tab Videos", display_text="🎬 Videos"),
            "Tag": st.column_config.TextColumn("🏷️ Nhãn Trạng Thái")
        })
        
        c1, c2, c3, c4 = st.columns(4)
        c1.download_button("📄 Tải TXT", data="\n".join([i["Handle"] for i in cart_items.values()]), file_name="gio_hang_dung_chung.txt", use_container_width=True, key=f"dl_txt_cart_{key_suffix}")
        buf_xl = io.BytesIO(); df_cart.to_excel(buf_xl, index=False)
        c2.download_button("📊 Tải Excel", data=buf_xl.getvalue(), file_name="gio_hang_dung_chung.xlsx", use_container_width=True, key=f"dl_xl_cart_{key_suffix}")
        if c3.button("⚡ Nạp Toàn Bộ Vào DB", type="primary", use_container_width=True, key=f"push_db_cart_{key_suffix}"):
            data_db = [{"handle": to_pure_id(i["Handle"]), "youtuber_name": i.get("Tên Kênh", ""), "source": f"Cart Import [{i.get('Tag', '')}]"} for i in cart_items.values()]
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
        <h1 style="font-weight: 900; margin-bottom: 5px; font-size: 2.4rem; letter-spacing: -0.03em;">YT CHECKER <span style="color: #D95F26;">PRO</span></h1>
        <p style="font-size: 1.05rem; font-weight: 500; opacity: 0.8;">Hệ thống phân tích, tìm kiếm kênh đồng ngách Đa Luồng Siêu Tốc & Tiết kiệm Quota.</p>
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

# --- HELPER FOR IN-MEMORY SORTING & FILTERING ---
def sort_and_filter_channels(channel_list, search_query, sort_by):
    filtered = list(channel_list)
    if search_query:
        q = search_query.strip().lower()
        filtered = [c for c in filtered if q in c.get('Handle', '').lower() or q in c.get('Tên Kênh', '').lower()]
    if sort_by == "Subscribers (Cao -> Thấp)":
        def parse_subs(val):
            try: return int(str(val.get('Subscribers', '0')).replace(',', ''))
            except Exception: return 0
        filtered.sort(key=parse_subs, reverse=True)
    elif sort_by == "Tên Kênh (A -> Z)": filtered.sort(key=lambda x: x.get('Tên Kênh', '').lower())
    elif sort_by == "Mới Đăng Video": filtered.sort(key=lambda x: x.get('Video Gần Nhất', ''), reverse=True)
    return filtered

# --- TAB 1: BATCH SEARCH ---
with tab1:
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>🔍 Kiểm tra Trùng Lặp Danh Sách Handle Hàng Loạt</h3>", unsafe_allow_html=True)
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
            with st.spinner(f"Đang đối chiếu đa luồng {len(target_list)} Handle với Database Supabase..."):
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
        render_kpi_cards([
            ("TỔNG SỐ KIỂM TRA", f"{len(new_handles) + len(existing_handles)}", "#47A5D1"),
            ("✅ KÊNH MỚI CÓ THỂ LÀM", f"{len(new_handles)}", "#10B981"),
            ("❌ KÊNH ĐÃ TỒN TẠI", f"{len(existing_handles)}", "#EF4444")
        ])
        
        res_tab1, res_tab2 = st.tabs([f"✅ Kênh Mới Chưa Làm ({len(new_handles)})", f"❌ Kênh Đã Tồn Tại ({len(existing_handles)})"])
        with res_tab1:
            if new_handles:
                cart_keys = set(st.session_state['cart'].keys())
                selected_set = st.session_state['selected_channels']
                
                # SEPARATED COUNTS FOR ACCURATE LOGIC
                selected_not_in_cart = [p for p in selected_set if p not in cart_keys]
                cnt_for_cart = len(selected_not_in_cart)
                cnt_total_sel = len(selected_set)

                tb1, tb2, tb3, tb4, tb5 = st.columns([2.5, 2.5, 2.0, 2.0, 1.0])
                with tb1:
                    if st.button("🛒 Thêm TẤT CẢ Kênh Mới vào Giỏ", type="primary", key="btn_add_all_t1", use_container_width=True):
                        for item in new_handles:
                            p_id = to_pure_id(item["Handle"])
                            item_data = {"Handle": item["Handle"], "Tên Kênh": item.get("Tên Kênh", p_id.upper()), "Link Kênh": f"https://www.youtube.com/@{p_id}", "Trạng Thái DB": "✅ KÊNH MỚI", "Tag": "📌 Chưa phân loại"}
                            st.session_state['cart'][p_id] = item_data
                            add_to_cart_db(p_id, item_data)
                        st.success(f"🎉 Đã thêm {len(new_handles)} kênh mới vào Giỏ hàng chung!")
                        st.rerun()
                with tb2:
                    if st.button(f"🛒 Thêm ({cnt_for_cart}) Kênh Mới Vào Giỏ", key="btn_add_sel_t1", use_container_width=True):
                        if cnt_for_cart > 0:
                            for item in new_handles:
                                p_id = to_pure_id(item["Handle"])
                                if p_id in selected_not_in_cart:
                                    item_data = {"Handle": item["Handle"], "Tên Kênh": item.get("Tên Kênh", p_id.upper()), "Link Kênh": f"https://www.youtube.com/@{p_id}", "Trạng Thái DB": "✅ KÊNH MỚI", "Tag": "📌 Chưa phân loại"}
                                    st.session_state['cart'][p_id] = item_data
                                    add_to_cart_db(p_id, item_data)
                            st.success(f"🎉 Đã thêm {cnt_for_cart} kênh mới vào giỏ!")
                            st.rerun()
                        else: st.warning("Không có kênh mới nào chưa được thêm vào giỏ trong các kênh bạn chọn!")
                with tb3:
                    if st.button(f"⚖️ So Sánh ({cnt_total_sel}) Kênh", key="btn_cmp_sel_t1", use_container_width=True):
                        if 1 < cnt_total_sel <= 5: compare_channels_dialog(get_selected_channel_data())
                        else: st.warning("Vui lòng chọn từ 2 đến 5 kênh để so sánh!")
                with tb4:
                    if st.button(f"🗑️ Xóa ({cnt_total_sel}) Đã Chọn", key="btn_del_sel_t1", use_container_width=True):
                        if cnt_total_sel > 0:
                            for p_id in list(selected_set): delete_channel_from_system(p_id)
                            st.session_state['selected_channels'].clear()
                            st.success(f"🗑️ Đã xóa {cnt_total_sel} kênh!")
                            st.rerun()
                        else: st.warning("Vui lòng tick chọn ít nhất 1 kênh!")
                with tb5:
                    if st.button("❌ Bỏ Chọn", key="btn_clear_sel_t1", use_container_width=True):
                        clear_selected_channels()
                        st.rerun()

                st.divider()
                
                # Pagination
                items_per_page = 20
                total_pages = max(1, (len(new_handles) + items_per_page - 1) // items_per_page)
                page_new = st.number_input("Trang (Kênh Mới):", min_value=1, max_value=total_pages, value=1, step=1, key="page_new_t1")
                start_idx = (page_new - 1) * items_per_page
                paged_new = new_handles[start_idx:start_idx + items_per_page]

                for idx, item in enumerate(paged_new):
                    p_id = to_pure_id(item["Handle"])
                    is_active = (p_id == st.session_state.get('active_inspected_handle'))
                    is_in_cart = p_id in st.session_state['cart']
                    
                    with st.container(border=True):
                        if is_active:
                            st.markdown('<div class="active-card-marker"></div>', unsafe_allow_html=True)
                            st.markdown('<div class="active-banner-tag">🔍 ĐANG XEM 6 VIDEO MỚI CỦA KÊNH NÀY</div>', unsafe_allow_html=True)
                        elif is_in_cart:
                            st.markdown('<div class="in-cart-marker"></div>', unsafe_allow_html=True)
                            st.markdown('<div class="in-cart-banner-tag">🛒 ĐÃ CÓ TRONG GIỎ HÀNG</div>', unsafe_allow_html=True)

                        c0, c1, c2, c3 = st.columns([0.4, 3.1, 3.5, 3.0])
                        with c0:
                            st.checkbox("", key=f"chk_t1_{p_id}", value=(p_id in st.session_state['selected_channels']), on_change=toggle_select_channel, args=(p_id,))
                        with c1:
                            st.markdown(f"<h3 style='margin:0; font-weight:800; font-size:1.3rem;'><a href='{item['Link Kênh']}' style='color:#D95F26; text-decoration:none;'>{item['Handle']}</a></h3>", unsafe_allow_html=True)
                            st.write(f"**{item.get('Tên Kênh', p_id.upper())}**")
                            if st.button("👁️ Xem 6 Video Mới", key=f"btn_prev_t1_{p_id}", on_click=set_active_inspected_channel, args=(p_id,)):
                                show_video_dialog(p_id)
                        with c2:
                            st.markdown(f"**Trạng thái:** <span style='color:#47A5D1;'>{item['Trạng thái']}</span>", unsafe_allow_html=True)
                        with c3:
                            st.write("**Thao tác:**")
                            bc1, bc2 = st.columns(2)
                            if is_in_cart:
                                current_tag = st.session_state['cart'][p_id].get("Tag", "📌 Chưa phân loại")
                                new_tag = st.selectbox("Gắn Nhãn:", ["📌 Chưa phân loại", "🔥 Ưu tiên làm", "📩 Đã liên hệ", "⏳ Đang chờ duyệt", "✅ Đã chốt", "❌ Bỏ qua"], index=["📌 Chưa phân loại", "🔥 Ưu tiên làm", "📩 Đã liên hệ", "⏳ Đang chờ duyệt", "✅ Đã chốt", "❌ Bỏ qua"].index(current_tag), key=f"tag_t1_{p_id}")
                                if new_tag != current_tag:
                                    st.session_state['cart'][p_id]["Tag"] = new_tag
                                    add_to_cart_db(p_id, st.session_state['cart'][p_id])
                                if bc1.button("❌ Bỏ Giỏ", key=f"rm_t1_{p_id}", use_container_width=True):
                                    remove_from_cart_db(p_id)
                                    del st.session_state['cart'][p_id]
                                    st.rerun()
                            else:
                                if bc1.button("🛒 Thêm Giỏ", key=f"add_t1_{p_id}", use_container_width=True):
                                    item_data = {"Handle": item["Handle"], "Tên Kênh": item.get("Tên Kênh", p_id.upper()), "Link Kênh": f"https://www.youtube.com/@{p_id}", "Trạng Thái DB": "✅ KÊNH MỚI", "Tag": "📌 Chưa phân loại"}
                                    st.session_state['cart'][p_id] = item_data
                                    add_to_cart_db(p_id, item_data)
                                    st.rerun()

                            if bc2.button("🗑️ Xóa Kênh", key=f"del_t1_{p_id}", use_container_width=True):
                                delete_channel_from_system(p_id)
                                st.toast(f"🗑️ Đã xóa kênh @{p_id}!")
                                st.rerun()
            else:
                st.info("Tất cả kênh đều đã tồn tại trong Database!")

        with res_tab2:
            if existing_handles:
                items_per_page_ex = 20
                total_pages_ex = max(1, (len(existing_handles) + items_per_page_ex - 1) // items_per_page_ex)
                page_ex = st.number_input("Trang (Kênh Tồn Tại):", min_value=1, max_value=total_pages_ex, value=1, step=1, key="page_ex_t1")
                start_idx_ex = (page_ex - 1) * items_per_page_ex
                paged_ex = existing_handles[start_idx_ex:start_idx_ex + items_per_page_ex]

                for idx, item in enumerate(paged_ex):
                    p_id = to_pure_id(item["Handle"])
                    is_active = (p_id == st.session_state.get('active_inspected_handle'))
                    
                    with st.container(border=True):
                        if is_active:
                            st.markdown('<div class="active-card-marker"></div>', unsafe_allow_html=True)
                            st.markdown('<div class="active-banner-tag">🔍 ĐANG XEM 6 VIDEO MỚI CỦA KÊNH NÀY</div>', unsafe_allow_html=True)

                        c1, c2, c3 = st.columns([4.0, 4.0, 2.0])
                        with c1:
                            st.markdown(f"<h3 style='margin:0; font-weight:800; font-size:1.3rem;'><a href='https://youtube.com/@{p_id}' style='text-decoration:none;'>{item['Handle']}</a></h3>", unsafe_allow_html=True)
                            st.write(f"**{item.get('Tên Kênh', 'N/A')}**")
                            if st.button("👁️ Xem 6 Video Mới", key=f"btn_prev_t1_ext_{p_id}", on_click=set_active_inspected_channel, args=(p_id,)):
                                show_video_dialog(p_id)
                        with c2:
                            st.markdown(f"**Trạng thái:** <span style='color:#D95F26;'>{item['Trạng thái']}</span>", unsafe_allow_html=True)
                        with c3:
                            if st.button("🗑️ Xóa DB", key=f"del_t1_ext_{p_id}", use_container_width=True):
                                delete_channel_from_system(p_id)
                                st.toast(f"🗑️ Đã xóa kênh @{p_id} khỏi DB!")
                                st.rerun()

    render_shared_cart_ui(key_suffix="tab1")

# --- TAB 2: LIVE API SCRAPER ---
with tab2:
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>⚡ Cào dữ liệu Live & Xuất Báo Cáo Audit chuẩn V4.14</h3>", unsafe_allow_html=True)
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
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>🎯 Săn Kênh Tương Tự & Giỏ Hàng (Multi-threaded Speed)</h3>", unsafe_allow_html=True)
    
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
                            if is_pass: passed_channels.append(res_data)
                            else: rejected_channels.append(res_data)

                    st.session_state['passed_channels'] = passed_channels
                    st.session_state['rejected_channels'] = rejected_channels

        except Exception as e: st.error(f"Lỗi: {e}")

    if 'passed_channels' in st.session_state or 'rejected_channels' in st.session_state:
        passed_list = st.session_state.get('passed_channels', [])
        rejected_list = st.session_state.get('rejected_channels', [])
        
        st.divider()
        render_kpi_cards([
            ("TỔNG ỨNG VIÊN", f"{len(passed_list) + len(rejected_list)}", "#47A5D1"),
            (f"✅ ĐẠT CHUẨN (>{min_subs_choice:,} SUBS)", f"{len(passed_list)}", "#10B981"),
            ("❌ BỊ LOẠI", f"{len(rejected_list)}", "#EF4444")
        ])

        tab_pass, tab_rej = st.tabs([f"✅ Kênh Đạt Chuẩn ({len(passed_list)})", f"❌ Kênh Bị Loại ({len(rejected_list)})"])
        
        # --- TAB PASSED ---
        with tab_pass:
            if passed_list:
                sf_col1, sf_col2, sf_col3 = st.columns([3, 2, 2])
                with sf_col1: filter_q = st.text_input("⚡ Lọc nhanh tên/handle:", key="filter_pass_q", placeholder="Gõ tên kênh để lọc...")
                with sf_col2: sort_by = st.selectbox("Sắp xếp danh sách:", options=["Subscribers (Cao -> Thấp)", "Tên Kênh (A -> Z)", "Mới Đăng Video"], key="sort_pass_by")
                with sf_col3:
                    st.write(""); st.write("")
                    if st.button("🛒 Thêm TẤT CẢ Kênh Mới Vào Giỏ", type="primary", use_container_width=True):
                        for row in passed_list:
                            if "✅" in row["Trạng Thái DB"]: 
                                p_id = to_pure_id(row["Handle"])
                                item_data = dict(row); item_data["Tag"] = "📌 Chưa phân loại"
                                st.session_state['cart'][p_id] = item_data
                                add_to_cart_db(p_id, item_data)
                        st.success("🎉 Đã thêm tất cả vào giỏ hàng!")
                        st.rerun()

                display_passed = sort_and_filter_channels(passed_list, filter_q, sort_by)

                cart_keys = set(st.session_state['cart'].keys())
                selected_set = st.session_state['selected_channels']
                
                # SEPARATED COUNTS FOR ACCURATE LOGIC
                selected_not_in_cart = [p for p in selected_set if p not in cart_keys]
                cnt_for_cart = len(selected_not_in_cart)
                cnt_total_sel = len(selected_set)

                ba1, ba2, ba3, ba4, ba5 = st.columns([2.5, 2.5, 2.0, 2.0, 1.0])
                with ba1:
                    if st.button(f"🛒 Thêm ({cnt_for_cart}) Kênh Mới Vào Giỏ", key="btn_add_sel_pass", use_container_width=True):
                        if cnt_for_cart > 0:
                            for row in display_passed:
                                p_id = to_pure_id(row['Handle'])
                                if p_id in selected_not_in_cart:
                                    item_data = dict(row); item_data["Tag"] = "📌 Chưa phân loại"
                                    st.session_state['cart'][p_id] = item_data
                                    add_to_cart_db(p_id, item_data)
                            st.success(f"🎉 Đã thêm {cnt_for_cart} kênh mới đã chọn!")
                            st.rerun()
                        else: st.warning("Không có kênh mới nào chưa được thêm vào giỏ trong các kênh bạn chọn!")
                with ba2:
                    if st.button(f"⚖️ So Sánh ({cnt_total_sel}) Kênh", key="btn_cmp_sel_pass", use_container_width=True):
                        if 1 < cnt_total_sel <= 5: compare_channels_dialog(get_selected_channel_data())
                        else: st.warning("Vui lòng chọn từ 2 đến 5 kênh để so sánh!")
                with ba3:
                    if st.button(f"🗑️ Xóa ({cnt_total_sel}) Đã Chọn", key="btn_del_sel_pass", use_container_width=True):
                        if cnt_total_sel > 0:
                            for p_id in list(selected_set): delete_channel_from_system(p_id)
                            st.session_state['selected_channels'].clear()
                            st.success(f"🗑️ Đã xóa {cnt_total_sel} kênh!")
                            st.rerun()
                        else: st.warning("Vui lòng tick chọn ít nhất 1 kênh!")
                with ba4:
                    if st.button("❌ Bỏ Chọn", key="btn_clear_sel_pass", use_container_width=True):
                        clear_selected_channels()
                        st.rerun()

                st.divider()

                # Pagination setup
                items_per_page = 20
                total_pages = max(1, (len(display_passed) + items_per_page - 1) // items_per_page)
                page_pass = st.number_input("Trang (Kênh Đạt Chuẩn):", min_value=1, max_value=total_pages, value=1, step=1, key="page_pass_t3")
                start_idx = (page_pass - 1) * items_per_page
                paged_passed = display_passed[start_idx:start_idx + items_per_page]

                for idx, row in enumerate(paged_passed):
                    p_id = to_pure_id(row['Handle'])
                    is_active = (p_id == st.session_state.get('active_inspected_handle'))
                    is_in_cart = p_id in st.session_state['cart']

                    with st.container(border=True):
                        if is_active:
                            st.markdown('<div class="active-card-marker"></div>', unsafe_allow_html=True)
                            st.markdown('<div class="active-banner-tag">🔍 ĐANG XEM 6 VIDEO MỚI CỦA KÊNH NÀY</div>', unsafe_allow_html=True)
                        elif is_in_cart:
                            st.markdown('<div class="in-cart-marker"></div>', unsafe_allow_html=True)
                            st.markdown('<div class="in-cart-banner-tag">🛒 ĐÃ CÓ TRONG GIỎ HÀNG</div>', unsafe_allow_html=True)

                        c0, c1, c2, c3, c4 = st.columns([0.4, 2.2, 3.0, 1.8, 3.0])
                        with c0:
                            st.checkbox("", key=f"chk_p_{p_id}", value=(p_id in st.session_state['selected_channels']), on_change=toggle_select_channel, args=(p_id,))
                        with c1:
                            st.markdown(f"<h3 style='margin:0; font-weight:800; font-size:1.3rem;'><a href='{row['Link Kênh']}' style='color:#D95F26; text-decoration:none;'>{row['Handle']}</a></h3>", unsafe_allow_html=True)
                            st.write(f"**{row['Tên Kênh']}**")
                            if st.button("👁️ Xem 6 Video Mới", key=f"btn_prev_pass_{p_id}", on_click=set_active_inspected_channel, args=(p_id,)):
                                show_video_dialog(p_id, pre_fetched_videos=row.get('recent_videos'))
                        with c2:
                            st.write(f"👥 **Subs:** `{row['Subscribers']}` | 🌍 **Q.Gia:** `{row['Quốc gia']}`")
                            st.write(f"🎬 **Tổng Video:** `{row['Tổng Số Video']}` | 📅 **Mới nhất:** `{row['Video Gần Nhất']}`")
                        with c3:
                            st.write(f"**Database:**\n<span style='color:#47A5D1; font-weight:700;'>{row['Trạng Thái DB']}</span>", unsafe_allow_html=True)
                        with c4:
                            bc1, bc2 = st.columns(2)
                            if is_in_cart:
                                current_tag = st.session_state['cart'][p_id].get("Tag", "📌 Chưa phân loại")
                                new_tag = st.selectbox("Gắn Nhãn:", ["📌 Chưa phân loại", "🔥 Ưu tiên làm", "📩 Đã liên hệ", "⏳ Đang chờ duyệt", "✅ Đã chốt", "❌ Bỏ qua"], index=["📌 Chưa phân loại", "🔥 Ưu tiên làm", "📩 Đã liên hệ", "⏳ Đang chờ duyệt", "✅ Đã chốt", "❌ Bỏ qua"].index(current_tag), key=f"tag_p_{p_id}")
                                if new_tag != current_tag:
                                    st.session_state['cart'][p_id]["Tag"] = new_tag
                                    add_to_cart_db(p_id, st.session_state['cart'][p_id])
                                if bc1.button("❌ Bỏ Giỏ", key=f"rm_p_{p_id}", use_container_width=True):
                                    remove_from_cart_db(p_id)
                                    del st.session_state['cart'][p_id]
                                    st.rerun()
                            else:
                                if bc1.button("🛒 Thêm Giỏ", key=f"add_p_{p_id}", use_container_width=True):
                                    item_data = dict(row); item_data["Tag"] = "📌 Chưa phân loại"
                                    st.session_state['cart'][p_id] = item_data
                                    add_to_cart_db(p_id, item_data)
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

                            if bc4.button("🗑️ Xóa Kênh", key=f"del_p_{p_id}", use_container_width=True):
                                delete_channel_from_system(p_id)
                                st.toast(f"🗑️ Đã xóa kênh @{p_id}!")
                                st.rerun()
            else:
                st.info("Không có kênh nào đạt chuẩn.")
                
        # --- TAB REJECTED ---
        with tab_rej:
            if rejected_list:
                rf_col1, rf_col2 = st.columns([3, 3])
                with rf_col1: filter_rej_q = st.text_input("⚡ Lọc nhanh kênh bị loại:", key="filter_rej_q", placeholder="Gõ tên/handle...")
                with rf_col2: sort_rej_by = st.selectbox("Sắp xếp:", options=["Subscribers (Cao -> Thấp)", "Tên Kênh (A -> Z)"], key="sort_rej_by")

                display_rejected = sort_and_filter_channels(rejected_list, filter_rej_q, sort_rej_by)
                
                items_per_page_rej = 20
                total_pages_rej = max(1, (len(display_rejected) + items_per_page_rej - 1) // items_per_page_rej)
                page_rej = st.number_input("Trang (Kênh Bị Loại):", min_value=1, max_value=total_pages_rej, value=1, step=1, key="page_rej_t3")
                start_idx_rej = (page_rej - 1) * items_per_page_rej
                paged_rejected = display_rejected[start_idx_rej:start_idx_rej + items_per_page_rej]

                for idx, row in enumerate(paged_rejected):
                    p_id = to_pure_id(row['Handle'])
                    is_active = (p_id == st.session_state.get('active_inspected_handle'))
                    is_in_cart = p_id in st.session_state['cart']
                    
                    with st.container(border=True):
                        if is_active:
                            st.markdown('<div class="active-card-marker"></div>', unsafe_allow_html=True)
                            st.markdown('<div class="active-banner-tag">🔍 ĐANG XEM 6 VIDEO MỚI CỦA KÊNH NÀY</div>', unsafe_allow_html=True)
                        elif is_in_cart:
                            st.markdown('<div class="in-cart-marker"></div>', unsafe_allow_html=True)
                            st.markdown('<div class="in-cart-banner-tag">🛒 ĐÃ CÓ TRONG GIỎ HÀNG</div>', unsafe_allow_html=True)

                        c0, c1, c2, c3, c4 = st.columns([0.4, 2.2, 3.0, 1.8, 3.0])
                        with c0:
                            st.checkbox("", key=f"chk_r_{p_id}", value=(p_id in st.session_state['selected_channels']), on_change=toggle_select_channel, args=(p_id,))
                        with c1:
                            st.markdown(f"<h3 style='margin:0; font-weight:800; font-size:1.3rem;'><a href='{row['Link Kênh']}' style='text-decoration:none;'>{row['Handle']}</a></h3>", unsafe_allow_html=True)
                            st.write(f"**{row['Tên Kênh']}**")
                            if st.button("👁️ Xem 6 Video Mới", key=f"btn_prev_rej_{p_id}", on_click=set_active_inspected_channel, args=(p_id,)):
                                show_video_dialog(p_id, pre_fetched_videos=row.get('recent_videos'))
                        with c2:
                            st.write(f"👥 **Subs:** `{row['Subscribers']}` | 🌍 **Q.Gia:** `{row.get('Quốc gia', '')}`")
                            st.write(f"🎬 **Tổng Video:** `{row.get('Tổng Số Video', '')}` | 📅 **Mới nhất:** `{row.get('Video Gần Nhất', '')}`")
                        with c3:
                            st.write(f"**Database:** {row.get('Trạng Thái DB', '')}")
                            st.markdown(f"❌ **Lý do:** <span style='color:#D95F26; font-weight:700;'>{row['Lý do loại']}</span>", unsafe_allow_html=True)
                        with c4:
                            bc1, bc2 = st.columns(2)
                            if is_in_cart:
                                current_tag = st.session_state['cart'][p_id].get("Tag", "📌 Chưa phân loại")
                                new_tag = st.selectbox("Gắn Nhãn:", ["📌 Chưa phân loại", "🔥 Ưu tiên làm", "📩 Đã liên hệ", "⏳ Đang chờ duyệt", "✅ Đã chốt", "❌ Bỏ qua"], index=["📌 Chưa phân loại", "🔥 Ưu tiên làm", "📩 Đã liên hệ", "⏳ Đang chờ duyệt", "✅ Đã chốt", "❌ Bỏ qua"].index(current_tag), key=f"tag_r_{p_id}")
                                if new_tag != current_tag:
                                    st.session_state['cart'][p_id]["Tag"] = new_tag
                                    add_to_cart_db(p_id, st.session_state['cart'][p_id])
                                if bc1.button("❌ Bỏ Giỏ", key=f"rm_r_{p_id}", use_container_width=True):
                                    remove_from_cart_db(p_id)
                                    del st.session_state['cart'][p_id]
                                    st.rerun()
                            else:
                                if bc1.button("🛒 Thêm Giỏ", key=f"add_r_{p_id}", use_container_width=True):
                                    item_data = dict(row); item_data["Tag"] = "📌 Chưa phân loại"
                                    st.session_state['cart'][p_id] = item_data
                                    add_to_cart_db(p_id, item_data)
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

                            if bc4.button("🗑️ Xóa", key=f"del_r_{p_id}", use_container_width=True):
                                delete_channel_from_system(p_id)
                                st.toast(f"🗑️ Đã xóa kênh @{p_id}!")
                                st.rerun()

    render_shared_cart_ui(key_suffix="tab3")

# --- TAB 4, TAB 5, TAB 6 ---
with tab4:
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>📤 Upload file .ZIP hoặc .TXT để cập nhật Database</h3>", unsafe_allow_html=True)
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
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>📊 Danh sách toàn bộ Channel trong Database</h3>", unsafe_allow_html=True)
    res = supabase.table("channels").select("*").execute()
    if res.data:
        df_all = pd.DataFrame(res.data)
        st.markdown(f"Tổng số kênh hiện có: <span style='font-weight:800; color:#D95F26;'>{len(df_all)}</span>", unsafe_allow_html=True)
        
        search_db = st.text_input("🔍 Tìm kiếm kênh trong Database (Handle hoặc Tên):", "")
        if search_db:
            df_filtered = df_all[
                df_all['handle'].str.contains(search_db, case=False, na=False) | 
                df_all['youtuber_name'].str.contains(search_db, case=False, na=False)
            ]
        else: df_filtered = df_all

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
                    st.markdown('<div class="active-banner-tag">🔍 ĐANG XEM 6 VIDEO MỚI CỦA KÊNH NÀY</div>', unsafe_allow_html=True)

                c1, c2, c3 = st.columns([4.0, 4.0, 2.0])
                with c1:
                    st.markdown(f"<h3 style='margin:0; font-weight:800; font-size:1.3rem;'><a href='https://youtube.com/@{p_id}' style='text-decoration:none;'>@{p_id}</a></h3>", unsafe_allow_html=True)
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
        st.download_button("📥 Tải về toàn bộ Database (CSV)", data=csv, file_name="master_youtube_database.csv", mime="text/csv", type="primary")

with tab6:
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>✨ Soi Từ Khóa Kênh (Channel & Video Tags SEO Inspector)</h3>", unsafe_allow_html=True)
    inspect_handle_input = st.text_input("Nhập Handle Kênh cần soi:", value="@NickDiGiovanni")
    if inspect_handle_input and st.button("🔍 Soi Từ Khóa Ngay", type="primary"):
        pure_inspect = to_pure_id(inspect_handle_input)
        if pure_inspect:
            try:
                cid_insp = get_channel_id_by_handle(pure_inspect)
                if not cid_insp: st.error("Không tìm thấy Channel ID cho kênh này!")
                else:
                    ext_data = extract_channel_master_keywords(cid_insp)
                    st.session_state['pending_keywords'] = ", ".join(ext_data['master_keywords'])
                    st.session_state['last_inspected_data'] = ext_data
                    st.session_state['last_inspected_handle'] = pure_inspect
                    st.rerun()
            except Exception as e: st.error(f"Lỗi khi soi từ khóa: {e}")

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
