import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import requests
from bs4 import BeautifulSoup
import xml.etree.ElementTree as ET
import zipfile
import os
import re
import datetime
import io
from supabase import create_client, Client

# Page Config
st.set_page_config(page_title="YouTube Channel Master DB", page_icon="📺", layout="wide")

# Persistent HTTP Session
session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9"
})

# Connect to Supabase
@st.cache_resource
def init_supabase():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = init_supabase()

# --- HELPER FUNCTIONS ---
def to_pure_id(raw_val):
    if not raw_val or pd.isna(raw_val) or str(raw_val).strip().upper() in ["N/A", "NAN", "NONE", ""]:
        return None
    s = str(raw_val).strip()
    m_url = re.search(r'youtube\.com/(?:@|c/|user/|channel/)?([^\s?#/]+)', s, re.IGNORECASE)
    if m_url:
        s = m_url.group(1)
    s = re.sub(r'[\s]+', '', s)
    s = re.sub(r'^@+', '', s).strip().lower()
    return s if s else None

def format_display_handle(raw_val):
    pure = to_pure_id(raw_val)
    return f"@{pure}" if pure else "N/A"

def extract_handle_from_filename(filename):
    base = os.path.basename(filename)
    base_no_ext = os.path.splitext(base)[0]
    pattern = r'_(?:backlog|\d{4}|\d{2,4}[-_/.]\d{1,2}[-_/.]\d{1,2}|\d{1,2}[-_/.]\d{1,2}[-_/.]\d{2,4}|\d{6,8})(?:_.*)?$'
    cleaned = re.sub(pattern, '', base_no_ext, flags=re.IGNORECASE)
    cleaned = re.sub(r'[\s]+', '', cleaned)
    pure_id = re.sub(r'^@+', '', cleaned).strip().lower()
    return pure_id if pure_id else None

def is_within_last_90_days(date_str):
    if not date_str or date_str == "N/A":
        return False
    s = str(date_str).strip().lower()
    today = datetime.date.today()
    
    m_iso = re.match(r'^(\d{4})-(\d{2})-(\d{2})', s)
    if m_iso:
        try:
            dt = datetime.date(int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3)))
            return 0 <= (today - dt).days <= 90
        except Exception:
            return False
            
    m_rel = re.search(r'(\d+)\s*(second|minute|hour|day|week|month|year)', s)
    if m_rel:
        num = int(m_rel.group(1))
        unit = m_rel.group(2)
        if 'day' in unit: return num <= 90
        elif 'week' in unit: return (num * 7) <= 90
        elif 'month' in unit: return num <= 3
        elif 'year' in unit: return False
        elif 'hour' in unit or 'minute' in unit or 'second' in unit: return True
            
    return False

def get_channel_info_live(input_str):
    pure_h = to_pure_id(input_str)
    if not pure_h:
        return None
        
    url = f"https://www.youtube.com/@{pure_h}"
    channel_title = pure_h
    channel_id = None
    latest_date = "N/A"
    
    try:
        r = session.get(url, timeout=7)
        if r.status_code == 200:
            html = r.text
            
            # Extract Channel Title
            m_title = re.search(r'<meta\s+property=["\']og:title["\']\s+content=["\']([^"\']+)["\']', html)
            if m_title:
                channel_title = m_title.group(1)
                
            # Extract Channel ID
            patterns = [
                r'itemprop=["\']channelId["\']\s+content=["\'](UC[A-Za-z0-9_.-]+)["\']',
                r'["\']channelId["\']:\s*["\'](UC[A-Za-z0-9_.-]+)["\']',
                r'["\']externalId["\']:\s*["\'](UC[A-Za-z0-9_.-]+)["\']',
                r'youtube\.com/channel/(UC[A-Za-z0-9_.-]+)'
            ]
            for p in patterns:
                m = re.search(p, html, re.IGNORECASE)
                if m:
                    channel_id = m.group(1)
                    break
    except Exception:
        pass

    # Method A: RSS Feed
    if channel_id:
        try:
            rss_url = f"https://www.youtube.com/feeds/videos.xml?channel_id={channel_id}"
            res = session.get(rss_url, timeout=6)
            if res.status_code == 200:
                tree = ET.fromstring(res.content)
                ns = {'atom': 'http://www.w3.org/2005/Atom', 'yt': 'http://www.youtube.com/xml/schemas/2015'}
                entries = tree.findall('atom:entry', ns)
                for entry in entries[:5]:
                    v_id_el = entry.find('yt:videoId', ns)
                    pub_el = entry.find('atom:published', ns)
                    if v_id_el is not None and pub_el is not None:
                        video_id = v_id_el.text
                        pub_text = pub_el.text
                        short_check_url = f"https://www.youtube.com/shorts/{video_id}"
                        try:
                            head_res = session.head(short_check_url, allow_redirects=False, timeout=3)
                            if head_res.status_code != 200:
                                latest_date = pub_text[:10]
                                break
                        except Exception:
                            latest_date = pub_text[:10]
                            break
        except Exception:
            pass

    # Method B: Scrape /videos tab
    if latest_date == "N/A":
        try:
            v_url = f"https://www.youtube.com/@{pure_h}/videos"
            res = session.get(v_url, timeout=6)
            if res.status_code == 200:
                m_date = re.search(r'"publishedTimeText":\s*\{\s*"simpleText":\s*"([^"]+)"', res.text)
                if not m_date:
                    m_date = re.search(r'"publishedTimeText":\s*\{\s*"runs":\s*\[\s*\{\s*"text":\s*"([^"]+)"', res.text)
                if m_date:
                    latest_date = m_date.group(1)
        except Exception:
            pass

    return {
        "title": channel_title,
        "handle": f"@{pure_h}",
        "url": url,
        "latest_date": latest_date,
        "is_active": is_within_last_90_days(latest_date)
    }

def generate_channel_report_bytes(youtuber_name, handle, yt_url, latest_video_date):
    """Generates a professional Excel report matching the standard audit format."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Channel Report"
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11)
    link_font = Font(name="Calibri", size=11, color="0563C1", underline="single")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    
    # Title Block
    ws['A1'] = "YOUTUBE CHANNEL AUDIT REPORT"
    ws['A1'].font = Font(name="Calibri", size=16, bold=True, color="1F4E78")
    ws['A2'] = f"Generated Date: {datetime.date.today().strftime('%Y-%m-%d')}"
    ws['A2'].font = Font(name="Calibri", size=10, italic=True, color="595959")
    
    # Table Headers
    headers = ["Youtuber Name", "Handle", "YouTube Link", "Latest Long Video Date", "Activity Status"]
    for col_num, h_text in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_num, value=h_text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Data Row
    act_status = "Active (<90 days)" if is_within_last_90_days(latest_video_date) else "Inactive / Shorts-Only"
    row_data = [youtuber_name, handle, yt_url, latest_video_date, act_status]
    for col_num, val in enumerate(row_data, 1):
        cell = ws.cell(row=5, column=col_num, value=val)
        cell.font = data_font
        cell.border = thin_border
        cell.alignment = Alignment(vertical="center")
        
        if col_num == 1 or col_num == 3:  # Hyperlinks
            cell.hyperlink = yt_url
            cell.font = link_font
        elif col_num in [2, 4, 5]:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions['A'].width = 32
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 22
    
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- APP UI HEADER ---
st.title("📺 YouTube Channel Master Database")
st.caption("Hệ thống quản lý, tra cứu, cào kênh live & xuất báo cáo 24/7")

# Tabs
tab1, tab2, tab3, tab4 = st.tabs(["🔍 Tra cứu Handle", "⚡ Tạo Báo Cáo Kênh", "📤 Upload Cập nhật Data", "📊 Xem Database"])

# --- TAB 1: SEARCH ---
with tab1:
    st.subheader("Kiểm tra Kênh đã có trong Database chưa")
    search_input = st.text_input("Nhập Handle hoặc Tên kênh cần tra cứu (ví dụ: @MrBeast, PewDiePie):")
    
    if search_input:
        pure_search = to_pure_id(search_input)
        if pure_search:
            response = supabase.table("channels").select("*").ilike("handle", pure_search).execute()
            
            if response.data:
                st.error(f"❌ **KÊNH ĐÃ TỒN TẠI!** Handle `@ {pure_search}` đã có trong Database.")
                st.json(response.data)
            else:
                st.success(f"✅ **KÊNH HOÀN TOÀN MỚI!** Handle `@ {pure_search}` CHƯA CÓ trong Database. Bạn có thể cào dữ liệu!")

# --- TAB 2: LIVE FETCH & GENERATE REPORT ---
with tab2:
    st.subheader("⚡ Cào dữ liệu Live & Xuất Báo Cáo Excel cho 1 Kênh")
    st.markdown("Dán Link YouTube hoặc Handle bất kỳ để kiểm tra, đồng bộ vào Database và tải file báo cáo Excel chuẩn.")
    
    channel_url_input = st.text_input("Dán Link kênh hoặc Handle vào đây:", placeholder="https://www.youtube.com/@aCookieGod hoặc @aCookieGod")
    
    if channel_url_input and st.button("🚀 Xử lý Kênh & Tạo Báo Cáo"):
        with st.spinner("Đang kết nối YouTube, kiểm tra ngày video dài và khởi tạo báo cáo..."):
            info = get_channel_info_live(channel_url_input)
            
            if info:
                pure_h = to_pure_id(info["handle"])
                
                # Check DB status
                db_res = supabase.table("channels").select("*").ilike("handle", pure_h).execute()
                exists_in_db = len(db_res.data) > 0
                
                # Auto Upsert into Supabase
                supabase.table("channels").upsert([{
                    "handle": pure_h,
                    "youtuber_name": info["title"],
                    "source": "Live Web Scraper"
                }], on_conflict="handle").execute()
                
                # Display Results
                st.divider()
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    st.markdown(f"### 🎯 **{info['title']}**")
                    st.write(f"• **Handle:** `{info['handle']}`")
                    st.write(f"• **Link Kênh:** [{info['url']}]({info['url']})")
                    st.write(f"• **Ngày đăng Video Dài gần nhất:** `{info['latest_date']}`")
                    
                    if info['is_active']:
                        st.success("🟢 **Trạng thái:** Đang hoạt động tích cực (Có video dài < 90 ngày)")
                    else:
                        st.warning("⚠️ **Trạng thái:** Không hoạt động hoặc Chỉ đăng Short (> 90 ngày)")
                        
                with col2:
                    if exists_in_db:
                        st.info("ℹ️ Kênh này **đã có sẵn** trong Database (Đã cập nhật lại thông tin mới).")
                    else:
                        st.success("✨ Kênh mới! **Đã tự động lưu** vào Database Cloud.")
                        
                    # Generate Excel bytes
                    excel_bytes = generate_channel_report_bytes(
                        youtuber_name=info["title"],
                        handle=info["handle"],
                        yt_url=info["url"],
                        latest_video_date=info["latest_date"]
                    )
                    
                    today_str = datetime.date.today().strftime('%Y-%m-%d')
                    report_filename = f"{pure_h}_{today_str}.xlsx"
                    
                    st.download_button(
                        label="📥 Tải về File Báo Cáo (.xlsx)",
                        data=excel_bytes,
                        file_name=report_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            else:
                st.error("Không tìm thấy thông tin kênh YouTube này. Vui lòng kiểm tra lại đường link hoặc Handle!")

# --- TAB 3: UPLOAD & UPDATE ---
with tab3:
    st.subheader("Upload file .ZIP hoặc .TXT để cập nhật Database")
    uploaded_files = st.file_uploader("Kéo thả file `.zip` (chứa các báo cáo Excel) hoặc file `.txt` vào đây:", type=["zip", "txt", "xlsx"], accept_multiple_files=True)
    
    if uploaded_files and st.button("🚀 Bắt đầu xử lý & Nạp vào Database"):
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
                                if h:
                                    new_handles_to_insert.append({"handle": h, "youtuber_name": h, "source": file_name})
                                    
            elif file_name.endswith('.txt'):
                content = file.read().decode("utf-8", errors="ignore")
                for line in content.splitlines():
                    h = to_pure_id(line)
                    if h:
                        new_handles_to_insert.append({"handle": h, "youtuber_name": h, "source": file_name})

        if new_handles_to_insert:
            df_insert = pd.DataFrame(new_handles_to_insert).drop_duplicates(subset=["handle"])
            data_dict = df_insert.to_dict(orient="records")
            
            try:
                supabase.table("channels").upsert(data_dict, on_conflict="handle").execute()
                st.success(f"🎉 Đã xử lý & đồng bộ thành công {len(data_dict)} Handle vào Database đám mây!")
            except Exception as e:
                st.error(f"Lỗi khi lưu dữ liệu: {e}")

# --- TAB 4: VIEW & DOWNLOAD ---
with tab4:
    st.subheader("Danh sách toàn bộ Channel trong Database")
    res = supabase.table("channels").select("*").execute()
    if res.data:
        df_all = pd.DataFrame(res.data)
        st.write(f"Tổng số kênh hiện có: **{len(df_all)}**")
        st.dataframe(df_all, use_container_width=True)
        
        csv = df_all.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Tải về toàn bộ Database (CSV)", data=csv, file_name="master_youtube_database.csv", mime="text/csv")
