import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from googleapiclient.discovery import build
import isodate
import pycountry
import requests
import zipfile
import os
import re
import datetime
import io
from collections import Counter
from PIL import Image as PILImage
from supabase import create_client, Client

# Page Config
st.set_page_config(page_title="YouTube Master DB & Related Finder", page_icon="📺", layout="wide")

# Safe State Lifecycle Management - Update pending state before widget instantiation
if 'pending_keywords' in st.session_state:
    st.session_state['custom_kw_tab3'] = st.session_state['pending_keywords']
    del st.session_state['pending_keywords']

if 'custom_kw_tab3' not in st.session_state:
    st.session_state['custom_kw_tab3'] = ""

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

def extract_handles_from_text(text_block):
    if not text_block:
        return []
    lines = re.split(r'[\n,\t\r]+', str(text_block))
    handles = []
    for line in lines:
        p = to_pure_id(line)
        if p and p not in handles:
            handles.append(p)
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
                    if p and p not in handles:
                        handles.append(p)
        elif fname.endswith('.xlsx') or fname.endswith('.xls'):
            df = pd.read_excel(uploaded_file)
            for col in df.columns:
                for val in df[col].dropna():
                    p = to_pure_id(val)
                    if p and p not in handles:
                        handles.append(p)
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
    return handles

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
    return False

# --- STRICT FILTERS CONFIGURATION ---
EXCLUDED_COUNTRIES = {'CN', 'TW', 'HK', 'TH', 'IN', 'VN'}

EXCLUDED_TEXT_REGEX = re.compile(
    r'[\u0E00-\u0E7F]|'  # Thai
    r'[\u4E00-\u9FFF]|'  # Chinese
    r'[\u0900-\u097F]|'  # Hindi
    r'[àáảãạâầấẩẫậăằắẳẵặèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]', # Vietnamese
    re.IGNORECASE
)

EXCLUDED_KEYWORDS = [
    'official mv', 'music video', 'official audio', 'album', 'song', 'records', 'lyrics', 'remix', 'vocal', 'cover',
    'news', 'politics', 'lgbt', 'lgbtq', 'gay', 'lesbian', 'transgender', 'war', 'military', 'ukraine', 'russia', 
    'tin tức', 'chính trị', 'thời sự', 'chiến tranh', 'đảng', 'quân sự'
]

STOP_WORDS = {
    'the', 'a', 'an', 'and', 'or', 'in', 'on', 'at', 'to', 'for', 'of', 'with', 'by', 'is', 'it', 'this', 'that',
    'ep', 'episode', 'part', 'video', 'shorts', 'full', 'hd', '2024', '2025', '2026', 'official', 'channel', 'vs',
    'dude', 'perfect', 'nick', 'digiovanni', 'mrbeast', 'pewdiepie'
}

def passes_layer1_metadata_filter(title, desc, country_code):
    if country_code in EXCLUDED_COUNTRIES:
        return False, f"Quốc gia bị loại ({country_code})"
        
    combined_text = f"{title} {desc}".lower()
    
    if EXCLUDED_TEXT_REGEX.search(combined_text):
        return False, "Ngôn ngữ không phù hợp (Trung, Thái, Hindi, Việt)"
        
    for kw in EXCLUDED_KEYWORDS:
        if kw in combined_text:
            return False, f"Chứa từ khóa bị cấm: '{kw}'"
            
    return True, "OK"

def clean_and_extract_keywords(text, seed_handle=""):
    seed_clean = seed_handle.replace('@', '').lower()
    words = re.findall(r'\b[a-zA-Z0-9]{3,}\b', text.lower())
    filtered = [w for w in words if w not in STOP_WORDS and w not in seed_clean]
    return filtered

# --- YOUTUBE DATA API ENGINE ---
def get_channel_id_by_handle(youtube, handle):
    clean = handle.replace('@', '').split('/')[-1].strip()
    try:
        request = youtube.channels().list(part="id", forHandle=clean)
        response = request.execute()
        if 'items' in response and len(response['items']) > 0:
            return response['items'][0]['id']
    except Exception:
        pass
    request = youtube.search().list(part="snippet", q=clean, type="channel", maxResults=1)
    response = request.execute()
    if 'items' in response and len(response['items']) > 0:
        return response['items'][0]['snippet']['channelId']
    return None

def extract_channel_master_keywords(youtube, channel_id):
    keywords_pool = []
    channel_keywords = []
    top_tags = []
    categories = []
    
    try:
        ch_req = youtube.channels().list(part="brandingSettings,contentDetails,snippet,topicDetails", id=channel_id)
        ch_res = ch_req.execute()
        
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
            for t in topics:
                cat_name = t.split('/')[-1].replace('_', ' ')
                categories.append(cat_name)
                
            uploads_playlist = item.get('contentDetails', {}).get('relatedPlaylists', {}).get('uploads', '')
            if uploads_playlist:
                v_req = youtube.playlistItems().list(part="snippet", playlistId=uploads_playlist, maxResults=15)
                v_res = v_req.execute()
                v_ids = [v['snippet']['resourceId']['videoId'] for v in v_res.get('items', [])]
                
                if v_ids:
                    v_detail_req = youtube.videos().list(part="snippet", id=','.join(v_ids))
                    v_detail_res = v_detail_req.execute()
                    
                    for v_item in v_detail_res.get('items', []):
                        tags = v_item.get('snippet', {}).get('tags', [])
                        for tag in tags:
                            if len(tag) > 2 and tag.lower() not in STOP_WORDS:
                                keywords_pool.append(tag.lower())
                                top_tags.append(tag.lower())
    except Exception:
        pass

    most_common_kws = [word for word, count in Counter(keywords_pool).most_common(15)]
    top_tag_counts = [word for word, count in Counter(top_tags).most_common(10)]
    
    return {
        "master_keywords": most_common_kws,
        "channel_keywords": list(set(channel_keywords))[:10],
        "top_tags": top_tag_counts,
        "categories": categories
    }

def get_channel_details(youtube, channel_id):
    request = youtube.channels().list(part="snippet,contentDetails,statistics", id=channel_id)
    response = request.execute()
    if 'items' in response and len(response['items']) > 0:
        item = response['items'][0]
        playlist_id = item.get('contentDetails', {}).get('relatedPlaylists', {}).get('uploads', '')
        sub_count = int(item['statistics'].get('subscriberCount', 0))
        description = item['snippet'].get('description', 'No description available.')

        joined_date_raw = item['snippet'].get('publishedAt', '')
        joined_date = ""
        if joined_date_raw:
            try:
                joined_date = pd.to_datetime(joined_date_raw).strftime("%b %d, %Y")
            except Exception:
                joined_date = joined_date_raw[:10]

        country_code = item['snippet'].get('country', 'N/A')
        country_name = "N/A"
        if country_code != 'N/A':
            country_obj = pycountry.countries.get(alpha_2=country_code)
            country_name = country_obj.name if country_obj else country_code

        thumbnails = item['snippet'].get('thumbnails', {})
        avatar_url = thumbnails.get('high', {}).get('url', thumbnails.get('medium', {}).get('url', ''))

        return playlist_id, sub_count, description, joined_date, country_name, country_code, avatar_url
    return None, 0, "", "", "", "", ""

def get_video_details(youtube, video_ids, progress_bar=None):
    video_data = []
    total = len(video_ids)
    if total == 0:
        return video_data
        
    for i in range(0, total, 50):
        try:
            request = youtube.videos().list(part="snippet,contentDetails,statistics", id=','.join(video_ids[i:i+50]))
            response = request.execute()
            for item in response.get('items', []):
                duration_seconds = int(isodate.parse_duration(item['contentDetails']['duration']).total_seconds())
                h, rem = divmod(duration_seconds, 3600)
                m, s = divmod(rem, 60)

                pub_date = item['snippet']['publishedAt']
                try:
                    formatted_date = pd.to_datetime(pub_date).strftime("%Y-%m-%d")
                except Exception:
                    formatted_date = pub_date[:10]

                video_data.append({
                    'Title': item['snippet']['title'],
                    'Link': f"https://youtube.com/watch?v={item['id']}",
                    'Length (Exact)': f"{h:02d}:{m:02d}:{s:02d}",
                    'Seconds': duration_seconds,
                    'Views': int(item['statistics'].get('viewCount', 0)),
                    'Published Date': formatted_date
                })
        except Exception:
            pass
        if progress_bar and total > 0:
            progress_bar.progress(min(1.0, (i + 50) / total))
            
    return video_data

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

    for row in range(2, 8):
        ws[f'A{row}'].font = Font(bold=True)

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
        except Exception:
            pass

    headers = ["Video Title", "Link", "Length", "Views", "Published Date"]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=12, column=col_num, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[12].height = 24

    for idx, v in enumerate(video_data):
        r = idx + 13
        cA = ws.cell(row=r, column=1, value=v['Title'])
        cA.font = Font(name="Calibri", size=11)
        cA.alignment = Alignment(horizontal="left", vertical="center")

        cB = ws.cell(row=r, column=2, value=v['Link'])
        if v.get('Link'):
            cB.hyperlink = v['Link']
            cB.font = Font(name="Calibri", size=11, color="0563C1", underline="single")
        cB.alignment = Alignment(horizontal="left", vertical="center")

        ws.cell(row=r, column=3, value=v['Length (Exact)']).alignment = Alignment(horizontal="center", vertical="center")
        
        cD = ws.cell(row=r, column=4, value=v['Views'])
        cD.font = Font(name="Calibri", size=11)
        cD.number_format = '#,##0'
        cD.alignment = Alignment(horizontal="right", vertical="center")

        ws.cell(row=r, column=5, value=v['Published Date']).alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions['A'].width = 55
    ws.column_dimensions['B'].width = 45
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15

    # TAB 2: DASHBOARD
    ws_charts = wb.create_sheet(title="Top 10 Video Title")
    top_10_videos = sorted(video_data, key=lambda x: x['Views'], reverse=True)[:10]

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    ws_charts['A1'] = "Top 10 Most Viewed Videos (Click to Watch)"
    ws_charts['B1'] = "Views"

    c1 = ws_charts['A1']
    c2 = ws_charts['B1']
    c1.font = header_font
    c2.font = header_font
    c1.fill = header_fill
    c2.fill = header_fill
    c1.alignment = Alignment(horizontal="center", vertical="center")
    c2.alignment = Alignment(horizontal="center", vertical="center")

    PALETTE = [
        {"fill": "2F5597", "font": "FFFFFF"}, {"fill": "C00000", "font": "FFFFFF"},
        {"fill": "70AD47", "font": "FFFFFF"}, {"fill": "7030A0", "font": "FFFFFF"},
        {"fill": "00C0C0", "font": "FFFFFF"}, {"fill": "E37222", "font": "FFFFFF"},
        {"fill": "41536B", "font": "FFFFFF"}, {"fill": "A04000", "font": "FFFFFF"},
        {"fill": "385723", "font": "FFFFFF"}, {"fill": "626262", "font": "FFFFFF"}
    ]

    ws_charts['D1'] = f"📊 Top 10 Most Viewed Videos - {clean_handle}"
    ws_charts['D1'].font = Font(bold=True, size=14, color="1F4E78")
    ws_charts['D1'].alignment = Alignment(vertical="center")

    for row_idx, video in enumerate(top_10_videos, start=2):
        color_idx = (row_idx - 2) % len(PALETTE)
        current_style = PALETTE[color_idx]

        short_title = video['Title'][:45] + "..." if len(video['Title']) > 45 else video['Title']

        title_cell = ws_charts.cell(row=row_idx, column=1, value=short_title)
        title_cell.hyperlink = video['Link']
        title_cell.font = Font(bold=True, color=current_style["font"], underline="single")
        title_cell.fill = PatternFill(start_color=current_style["fill"], end_color=current_style["fill"], fill_type="solid")

        view_cell = ws_charts.cell(row=row_idx, column=2, value=video['Views'])
        view_cell.number_format = '#,##0'
        view_cell.font = Font(bold=True, color=current_style["font"])
        view_cell.fill = PatternFill(start_color=current_style["fill"], end_color=current_style["fill"], fill_type="solid")
        view_cell.alignment = Alignment(horizontal="right")

    ws_charts.column_dimensions['A'].width = 50
    ws_charts.column_dimensions['B'].width = 15

    if len(top_10_videos) > 0:
        chart = BarChart()
        chart.type = "col"
        chart.title = None
        chart.y_axis.title = 'Total Views'
        chart.x_axis.title = 'Videos'

        data_ref = Reference(ws_charts, min_col=2, min_row=1, max_col=2, max_row=len(top_10_videos)+1)
        cats_ref = Reference(ws_charts, min_col=1, min_row=2, max_col=1, max_row=len(top_10_videos)+1)

        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.legend = None

        series = chart.series[0]
        for idx in range(len(top_10_videos)):
            dp = DataPoint(idx=idx)
            dp.graphicalProperties.solidFill = PALETTE[idx]["fill"]
            series.dPt.append(dp)

        chart.width = 22
        chart.height = 14
        ws_charts.add_chart(chart, "D3")

    output_buf = io.BytesIO()
    wb.save(output_buf)
    return output_buf.getvalue()

# --- APP UI HEADER ---
st.title("📺 YouTube Channel Master Database")
st.caption("Hệ thống tra cứu, cào live, Săn Kênh Đồng Ngách & Soi Từ Khóa Kênh 24/7")

# Tabs
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🔍 Tra cứu Handle Hàng Loạt", 
    "⚡ Cào Live & Tạo Báo Cáo Audit", 
    "🎯 Săn Kênh Tương Tự (Content-Based)",
    "📤 Upload Cập nhật Data", 
    "📊 Xem Database",
    "✨ Soi Từ Khóa Kênh (SEO Inspector)"
])

# --- TAB 1: BATCH SEARCH ---
with tab1:
    st.subheader("🔍 Kiểm tra Trùng Lặp Danh Sách Handle Hàng Loạt")
    col_s1, col_s2 = st.columns([2, 1])
    with col_s1:
        text_input_area = st.text_area(
            "Cách 1: Dán danh sách Handle/Link kênh vào đây (mỗi kênh 1 dòng):",
            placeholder="@MrBeast\n@PewDiePie\nhttps://www.youtube.com/@aCookieGod\n@123GO_",
            height=180
        )
    with col_s2:
        file_input_check = st.file_uploader("Cách 2: Hoặc Upload file danh sách (.txt, .csv, .xlsx):", type=["txt", "csv", "xlsx", "xls"])
        
    if st.button("🔎 Bắt Đầu Kiểm Tra Hàng Loạt", type="primary"):
        all_target_handles = set()
        if text_input_area:
            for h in extract_handles_from_text(text_input_area): all_target_handles.add(h)
        if file_input_check:
            for h in extract_handles_from_file(file_input_check): all_target_handles.add(h)
                
        target_list = list(all_target_handles)
        if not target_list:
            st.warning("⚠️ Vui lòng dán danh sách Handle hoặc chọn file để kiểm tra!")
        else:
            with st.spinner(f"Đang đối chiếu {len(target_list)} Handle với Database Supabase..."):
                response = supabase.table("channels").select("handle, youtuber_name, source").in_("handle", target_list).execute()
                db_matches = {item["handle"].lower(): item for item in response.data} if response.data else {}
                
                new_handles, existing_handles, report_data = [], [], []
                for h in target_list:
                    if h in db_matches:
                        matched_item = db_matches[h]
                        existing_handles.append({
                            "Handle": f"@{h}",
                            "Tên / Ghi chú": matched_item.get("youtuber_name", "N/A"),
                            "Nguồn Dữ Liệu": matched_item.get("source", "N/A"),
                            "Trạng thái": "❌ Đã có trong DB"
                        })
                        report_data.append({"Handle": f"@{h}", "Trạng Thái": "❌ Đã có", "Tên/Nguồn": matched_item.get("youtuber_name", "N/A")})
                    else:
                        new_handles.append({"Handle": f"@{h}", "Trạng thái": "✅ Kênh Mới (Chưa làm)"})
                        report_data.append({"Handle": f"@{h}", "Trạng Thái": "✅ Mới", "Tên/Nguồn": "Sẵn sàng cào dữ liệu"})

                st.divider()
                m1, m2, m3 = st.columns(3)
                m1.metric("Tổng số kênh kiểm tra", f"{len(target_list)} kênh")
                m2.metric("❌ Kênh Đã Tồn Tại", f"{len(existing_handles)} kênh")
                m3.metric("✅ Kênh Mới Có Thể Làm", f"{len(new_handles)} kênh")
                
                res_tab1, res_tab2 = st.tabs([f"✅ Kênh Mới Chưa Làm ({len(new_handles)})", f"❌ Kênh Đã Tồn Tại ({len(existing_handles)})"])
                with res_tab1:
                    if new_handles:
                        df_new = pd.DataFrame(new_handles)
                        st.dataframe(df_new, use_container_width=True)
                        col_dl1, col_dl2 = st.columns(2)
                        with col_dl1:
                            st.download_button("📥 Tải Danh Sách Kênh Mới (.txt)", data="\n".join([item["Handle"] for item in new_handles]), file_name=f"danh_sach_kenh_moi_{datetime.date.today().strftime('%d-%m-%Y')}.txt", mime="text/plain")
                        with col_dl2:
                            buf_new = io.BytesIO()
                            df_new.to_excel(buf_new, index=False)
                            st.download_button("📥 Tải Danh Sách Kênh Mới (.xlsx)", data=buf_new.getvalue(), file_name=f"danh_sach_kenh_moi_{datetime.date.today().strftime('%d-%m-%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    else:
                        st.info("Tất cả các kênh trong danh sách của bạn đều đã tồn tại trong Database!")

                with res_tab2:
                    if existing_handles:
                        st.dataframe(pd.DataFrame(existing_handles), use_container_width=True)
                    else:
                        st.success("🎉 Tuyệt vời! Không có kênh nào trùng lặp!")

# --- TAB 2: LIVE API SCRAPER ---
with tab2:
    st.subheader("⚡ Cào dữ liệu Live & Xuất Báo Cáo Audit chuẩn V4.14")
    col_input1, col_input2 = st.columns([2, 1])
    with col_input1:
        channel_url_input = st.text_input("Dán Link kênh hoặc Handle vào đây:", value="@4wd247", placeholder="https://www.youtube.com/@4wd247 hoặc @4wd247")
    with col_input2:
        default_api_key = st.secrets.get("YOUTUBE_API_KEY", "AIzaSyBrTtmMp-txQ7ID15wrJZUpN-i53SRVzgk")
        api_key_input = st.text_input("YouTube Data API Key:", value=default_api_key, type="password")

    if channel_url_input and st.button("🚀 Xử lý Kênh & Tạo Báo Cáo V4.14"):
        pure_h = to_pure_id(channel_url_input)
        if not pure_h:
            st.error("Handle không hợp lệ!")
        else:
            try:
                youtube = build("youtube", "v3", developerKey=api_key_input)
                channel_id = get_channel_id_by_handle(youtube, pure_h)
                if not channel_id:
                    st.error("Không tìm thấy Channel ID!")
                else:
                    playlist_id, sub_count, channel_desc, channel_joined, channel_country, c_code, avatar_url = get_channel_details(youtube, channel_id)
                    video_ids = []
                    next_page_token = None
                    try:
                        while True:
                            req = youtube.playlistItems().list(part="snippet", playlistId=playlist_id, maxResults=50, pageToken=next_page_token)
                            res = req.execute()
                            for item in res['items']:
                                video_ids.append(item['snippet']['resourceId']['videoId'])
                            next_page_token = res.get('nextPageToken')
                            if not next_page_token: break
                    except Exception:
                        pass
                        
                    prog_bar = st.progress(0.0)
                    video_data = get_video_details(youtube, video_ids, progress_bar=prog_bar)
                    
                    supabase.table("channels").upsert([{"handle": pure_h, "youtuber_name": pure_h.upper(), "source": "YouTube API V4.14"}], on_conflict="handle").execute()
                    
                    st.divider()
                    col_res1, col_res2 = st.columns([1, 2])
                    latest_date = video_data[0]['Published Date'] if video_data else "N/A"
                    with col_res1:
                        if avatar_url: st.image(avatar_url, width=150)
                    with col_res2:
                        st.markdown(f"### 🎯 **@{pure_h}**")
                        st.write(f"• **Số lượng Video:** `{len(video_data):,}` | **Subs:** `{sub_count:,}`")
                        st.write(f"• **Quốc gia:** `{channel_country}` | **Gần nhất:** `{latest_date}`")

                    excel_v414_bytes = generate_v414_excel_report(pure_h, sub_count, channel_desc, channel_joined, channel_country, avatar_url, video_data)
                    date_now_str = datetime.datetime.now().strftime("%d-%m-%Y")
                    output_file_name = f"{pure_h}_{date_now_str}.xlsx"
                    st.download_button(f"📥 Tải về File Audit V4.14 ({output_file_name})", data=excel_v414_bytes, file_name=output_file_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except Exception as e:
                st.error(f"Lỗi: {e}")

# --- TAB 3: CONTENT-BASED SMART RELATED FINDER WITH AUTOMATIC DYNAMIC LINKAGE ---
with tab3:
    st.subheader("🎯 Săn Kênh Tương Tự Theo Nội Dung & Xuất Báo Cáo Audit 1-Click")
    st.markdown("Hệ thống tự phân tích **Nội dung Video & Tags** $\rightarrow$ Quét rộng các Creator cùng chủ đề $\rightarrow$ Lọc tiêu chuẩn $\rightarrow$ Xuất Báo Cáo Audit V4.14 cho **bất kỳ kênh nào**.")
    
    col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
    with col_f1:
        seed_channel_input = st.text_input("Nhập Handle Kênh Mồi (ví dụ: @dudeperfect, @NickDiGiovanni, @4wd247):", value="@NickDiGiovanni", key="seed_input_tab3")
        
        # Auto-Extract Keywords Button in Tab 3
        if st.button("✨ Tự Động Phân Tích từ Kênh Mồi", help="Bấm để YouTube API tự bóc tách thẻ từ khóa chuẩn nhất từ kênh mồi và điền vào bên dưới"):
            pure_s_auto = to_pure_id(seed_channel_input)
            if pure_s_auto:
                try:
                    yt_auto = build("youtube", "v3", developerKey=st.secrets.get("YOUTUBE_API_KEY", "AIzaSyDDBEJscqkGGpG1xtuL4wYPuFkS4BIL854"))
                    cid_auto = get_channel_id_by_handle(yt_auto, pure_s_auto)
                    if cid_auto:
                        extracted = extract_channel_master_keywords(yt_auto, cid_auto)
                        st.session_state['pending_keywords'] = ", ".join(extracted['master_keywords'][:6])
                        st.rerun()
                except Exception as e:
                    st.error(f"Lỗi bóc từ khóa: {e}")
                    
        custom_keywords_input = st.text_input(
            "Từ khóa chủ đề (Tự động liên kết từ Tab 6 hoặc bấm nút phân tích ở trên):", 
            key="custom_kw_tab3",
            placeholder="Ví dụ: Cooking, Food, Recipe, Chef"
        )
        
    with col_f2:
        min_subs_choice = st.selectbox(
            "Mốc Subscribers Tối Thiểu:",
            options=[100000, 250000, 500000, 1000000],
            index=3, # Default 1,000,000 Subs
            format_func=lambda x: f"{x:,} Subs ({'1 Triệu' if x==1000000 else f'{x//1000}k'})"
        )
        min_duration_choice = st.selectbox(
            "Lọc Loại Bỏ Kênh Shorts:",
            options=[60, 180, 300, 600],
            index=0, # Default >= 60 seconds
            format_func=lambda x: f"Loại Shorts < {x//60} phút" if x < 600 else "Bắt buộc có Video > 10 phút"
        )
    with col_f3:
        default_api_key_tab3 = st.secrets.get("YOUTUBE_API_KEY", "AIzaSyDDBEJscqkGGpG1xtuL4wYPuFkS4BIL854")
        api_key_tab3 = st.text_input("YouTube Data API Key:", value=default_api_key_tab3, type="password", key="api_key_tab3")

    if seed_channel_input and st.button("🚀 Bắt Đầu Săn Kênh Đồng Ngách"):
        pure_seed = to_pure_id(seed_channel_input)
        if not pure_seed:
            st.error("Handle kênh mồi không hợp lệ!")
        else:
            try:
                youtube = build("youtube", "v3", developerKey=api_key_tab3)
                
                st.info("🔍 Đang tìm kiếm Channel ID của kênh mồi...")
                seed_id = get_channel_id_by_handle(youtube, pure_seed)
                
                if not seed_id:
                    st.error("Không tìm thấy kênh mồi này trên YouTube!")
                else:
                    playlist_id, _, seed_desc, _, _, _, _ = get_channel_details(youtube, seed_id)
                    
                    st.info("💡 Đang bóc tách Từ Khóa Chủ Đề từ nội dung & thẻ Tags...")
                    
                    if custom_keywords_input:
                        top_kw_list = clean_and_extract_keywords(custom_keywords_input, seed_handle=pure_seed)
                    else:
                        ext_info = extract_channel_master_keywords(youtube, seed_id)
                        top_kw_list = ext_info['master_keywords'][:4] if ext_info['master_keywords'] else [pure_seed.replace('_', ' ')]
                        
                    st.write(f"🏷️ **Từ khóa ngách chính dùng để quét:** `{', '.join(top_kw_list)}`")
                    
                    st.info("🌐 Đang quét tìm rộng hàng trăm Kênh & Videos thuộc chủ đề này...")
                    
                    candidate_channel_ids = set()
                    
                    # 1. Search Channels directly by query
                    q_chan = " ".join(top_kw_list[:2])
                    c_search_req = youtube.search().list(part="snippet", q=q_chan, type="channel", maxResults=50)
                    c_search_res = c_search_req.execute()
                    for c_item in c_search_res.get('items', []):
                        found_cid = c_item['snippet']['channelId']
                        if found_cid != seed_id: candidate_channel_ids.add(found_cid)
                        
                    # 2. Search Videos in the exact niche to find producing Channels
                    search_queries = [" ".join(top_kw_list[:2]), " ".join(top_kw_list[2:4])] if len(top_kw_list) >= 4 else [" ".join(top_kw_list)]
                    for q in search_queries:
                        if not q.strip(): continue
                        v_search_req = youtube.search().list(part="snippet", q=q, type="video", maxResults=50)
                        v_search_res = v_search_req.execute()
                        for v_item in v_search_res.get('items', []):
                            found_cid = v_item['snippet']['channelId']
                            if found_cid != seed_id: candidate_channel_ids.add(found_cid)
                            
                    candidate_ids_list = list(candidate_channel_ids)
                    
                    if not candidate_ids_list:
                        st.warning("Không quét thêm được kênh ứng viên nào cùng chủ đề!")
                    else:
                        st.info(f"📊 Tìm thấy {len(candidate_ids_list)} kênh ứng viên. Đang áp dụng phễu lọc tiêu chuẩn (Subs $\ge$ {min_subs_choice:,})...")
                        
                        passed_channels = []
                        rejected_channels = []
                        
                        channel_item_map = {}
                        candidate_handles = []
                        
                        for i in range(0, len(candidate_ids_list), 50):
                            chan_req = youtube.channels().list(part="snippet,contentDetails,statistics", id=','.join(candidate_ids_list[i:i+50]))
                            chan_res = chan_req.execute()
                            for item in chan_res.get('items', []):
                                c_handle = to_pure_id(item['snippet'].get('customUrl', '')) or item['id'].lower()
                                candidate_handles.append(c_handle)
                                channel_item_map[c_handle] = item

                        # Query DB in batch
                        db_res = supabase.table("channels").select("handle").in_("handle", candidate_handles).execute()
                        db_existing_set = {r["handle"].lower() for r in db_res.data} if db_res.data else set()
                        
                        for c_handle, item in channel_item_map.items():
                            c_title = item['snippet']['title']
                            c_desc = item['snippet'].get('description', '')
                            c_country = item['snippet'].get('country', 'N/A')
                            c_subs = int(item['statistics'].get('subscriberCount', 0))
                            c_url = f"https://www.youtube.com/@{c_handle}"
                            
                            # --- SUBSCRIBER FILTER ---
                            if c_subs < min_subs_choice:
                                rejected_channels.append({"Handle": f"@{c_handle}", "Link Kênh": c_url, "Tên Kênh": c_title, "Subscribers": f"{c_subs:,}", "Lý do loại": f"Dưới mốc chọn (<{min_subs_choice:,} Subs)"})
                                continue

                            # --- LAYER 1 FILTER ---
                            passes_l1, l1_reason = passes_layer1_metadata_filter(c_title, c_desc, c_country)
                            if not passes_l1:
                                rejected_channels.append({"Handle": f"@{c_handle}", "Link Kênh": c_url, "Tên Kênh": c_title, "Subscribers": f"{c_subs:,}", "Lý do loại": l1_reason})
                                continue
                                
                            # --- LAYER 2 FILTER ---
                            c_playlist = item.get('contentDetails', {}).get('relatedPlaylists', {}).get('uploads', '')
                            if not c_playlist:
                                rejected_channels.append({"Handle": f"@{c_handle}", "Link Kênh": c_url, "Tên Kênh": c_title, "Subscribers": f"{c_subs:,}", "Lý do loại": "Không có Playlist Uploads"})
                                continue

                            try:
                                v_req = youtube.playlistItems().list(part="snippet", playlistId=c_playlist, maxResults=10)
                                v_res = v_req.execute()
                                v_ids = [v_item['snippet']['resourceId']['videoId'] for v_item in v_res.get('items', [])]
                            except Exception:
                                rejected_channels.append({"Handle": f"@{c_handle}", "Link Kênh": c_url, "Tên Kênh": c_title, "Subscribers": f"{c_subs:,}", "Lý do loại": "Playlist ẩn hoặc lỗi 404"})
                                continue
                                
                            if not v_ids:
                                rejected_channels.append({"Handle": f"@{c_handle}", "Link Kênh": c_url, "Tên Kênh": c_title, "Subscribers": f"{c_subs:,}", "Lý do loại": "Kênh không có video nào"})
                                continue
                                
                            v_details = get_video_details(youtube, v_ids)
                            latest_date = v_details[0]['Published Date'] if v_details else "N/A"
                            
                            if not is_within_last_90_days(latest_date):
                                rejected_channels.append({"Handle": f"@{c_handle}", "Link Kênh": c_url, "Tên Kênh": c_title, "Subscribers": f"{c_subs:,}", "Lý do loại": f"Bỏ trống > 90 ngày (Mới nhất: {latest_date})"})
                                continue
                                
                            has_qualifying_video = any(v['Seconds'] >= min_duration_choice for v in v_details)
                            if not has_qualifying_video:
                                rejected_channels.append({"Handle": f"@{c_handle}", "Link Kênh": c_url, "Tên Kênh": c_title, "Subscribers": f"{c_subs:,}", "Lý do loại": f"Video ngắn dưới {min_duration_choice//60} phút (Shorts-only)"})
                                continue
                                
                            in_db = c_handle in db_existing_set
                            passed_channels.append({
                                "Handle": f"@{c_handle}",
                                "Link Kênh": c_url,
                                "Tên Kênh": c_title,
                                "Subscribers": f"{c_subs:,}",
                                "Quốc gia": c_country,
                                "Video Gần Nhất": latest_date,
                                "Trạng Thái DB": "❌ Đã có trong DB" if in_db else "✅ KÊNH MỚI TIỀM NĂNG"
                            })

                        st.session_state['passed_channels'] = passed_channels
                        st.session_state['rejected_channels'] = rejected_channels

                        st.divider()
                        st.markdown(f"### 🎉 Kết Quả Săn Kênh Đồng Ngách Từ `{pure_seed}`")
                        
                        col_m1, col_f2, col_f3 = st.columns(3)
                        col_m1.metric("Tổng ứng viên đã quét", len(candidate_ids_list))
                        col_f2.metric(f"✅ Đạt Chuẩn (>{min_subs_choice:,} Subs)", len(passed_channels))
                        col_f3.metric("❌ Bị Loại Bởi Bộ Lọc", len(rejected_channels))

            except Exception as e:
                st.error(f"Lỗi khi tìm kênh tương tự: {e}")

    # Display Tables & 1-Click Audit Report Generator
    if 'passed_channels' in st.session_state or 'rejected_channels' in st.session_state:
        passed_list = st.session_state.get('passed_channels', [])
        rejected_list = st.session_state.get('rejected_channels', [])
        
        tab_pass, tab_rej = st.tabs([f"✅ Kênh Đạt Chuẩn ({len(passed_list)})", f"❌ Kênh Bị Loại ({len(rejected_list)})"])
        
        with tab_pass:
            if passed_list:
                df_pass = pd.DataFrame(passed_list)
                st.dataframe(
                    df_pass,
                    use_container_width=True,
                    column_config={"Link Kênh": st.column_config.LinkColumn("Link Kênh", display_text="Xem Kênh 🔗")}
                )
                new_only_handles = [row["Handle"] for row in passed_list if "✅" in row["Trạng Thái DB"]]
                if new_only_handles:
                    st.download_button("📥 Tải Danh Sách Handle Kênh Mới (.txt)", data="\n".join(new_only_handles), file_name="kenh_moi_da_loc.txt", mime="text/plain")
            else:
                st.info("Không có kênh nào đạt chuẩn.")
                
        with tab_rej:
            if rejected_list:
                df_rej = pd.DataFrame(rejected_list)
                st.dataframe(
                    df_rej,
                    use_container_width=True,
                    column_config={"Link Kênh": st.column_config.LinkColumn("Link Kênh", display_text="Xem Kênh 🔗")}
                )
            else:
                st.info("Không có kênh nào bị loại.")

        st.divider()
        st.subheader("📄 Tạo & Tải File Báo Cáo Audit V4.14 Cho Kênh Bất Kỳ")
        st.caption("Chọn bất kỳ kênh nào trong danh sách kết quả (kể cả kênh đạt chuẩn hay **kênh bị loại**) để xuất ngay 1 file Excel Audit 2 Sheet tiêu chuẩn.")
        
        channel_options_map = {}
        for item in passed_list:
            label = f"{item['Handle']} | {item['Tên Kênh']} | {item['Subscribers']} Subs (✅ Đạt Chuẩn)"
            channel_options_map[label] = item['Handle']
            
        for item in rejected_list:
            label = f"{item['Handle']} | {item['Tên Kênh']} | {item['Subscribers']} Subs (❌ Loại: {item['Lý do loại']})"
            channel_options_map[label] = item['Handle']
            
        if channel_options_map:
            selected_label = st.selectbox("Chọn kênh bạn muốn xuất file báo cáo Audit V4.14:", options=list(channel_options_map.keys()))
            selected_handle = channel_options_map[selected_label]
            
            if st.button("🚀 Cào Toàn Bộ Video & Dựng File Audit V4.14"):
                pure_audit_h = to_pure_id(selected_handle)
                try:
                    yt_audit = build("youtube", "v3", developerKey=api_key_tab3)
                    cid_audit = get_channel_id_by_handle(yt_audit, pure_audit_h)
                    
                    if not cid_audit:
                        st.error("Không tìm thấy Channel ID!")
                    else:
                        with st.spinner(f"Đang trích xuất toàn bộ video & dựng báo cáo Audit 2 Sheet cho {pure_audit_h}..."):
                            playlist_id, sub_count, channel_desc, channel_joined, channel_country, c_code, avatar_url = get_channel_details(yt_audit, cid_audit)
                            
                            v_ids = []
                            next_token = None
                            while True:
                                req = yt_audit.playlistItems().list(part="snippet", playlistId=playlist_id, maxResults=50, pageToken=next_token)
                                res = req.execute()
                                for v_item in res.get('items', []):
                                    v_ids.append(v_item['snippet']['resourceId']['videoId'])
                                next_token = res.get('nextPageToken')
                                if not next_token: break
                                
                            prog_bar_audit = st.progress(0.0)
                            v_data_audit = get_video_details(yt_audit, v_ids, progress_bar=prog_bar_audit)
                            
                            excel_audit_bytes = generate_v414_excel_report(
                                clean_handle=pure_audit_h,
                                sub_count=sub_count,
                                channel_desc=channel_desc,
                                channel_joined=channel_joined,
                                channel_country=channel_country,
                                avatar_url=avatar_url,
                                video_data=v_data_audit
                            )
                            
                            out_fname = f"{pure_audit_h}_{datetime.datetime.now().strftime('%d-%m-%Y')}.xlsx"
                            
                            st.success(f"🎉 Đã dựng xong báo cáo Audit V4.14 cho @{pure_audit_h} ({len(v_data_audit)} video)!")
                            st.download_button(
                                label=f"📥 Tải Về Báo Cáo Audit Excel ({out_fname})",
                                data=excel_audit_bytes,
                                file_name=out_fname,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                except Exception as e:
                    st.error(f"Lỗi khi tạo báo cáo: {e}")

# --- TAB 4: UPLOAD & UPDATE ---
with tab4:
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
                                if h: new_handles_to_insert.append({"handle": h, "youtuber_name": h, "source": file_name})
            elif file_name.endswith('.txt'):
                content = file.read().decode("utf-8", errors="ignore")
                for line in content.splitlines():
                    h = to_pure_id(line)
                    if h: new_handles_to_insert.append({"handle": h, "youtuber_name": h, "source": file_name})

        if new_handles_to_insert:
            df_insert = pd.DataFrame(new_handles_to_insert).drop_duplicates(subset=["handle"])
            data_dict = df_insert.to_dict(orient="records")
            try:
                supabase.table("channels").upsert(data_dict, on_conflict="handle").execute()
                st.success(f"🎉 Đã xử lý & đồng bộ thành công {len(data_dict)} Handle vào Database đám mây!")
            except Exception as e:
                st.error(f"Lỗi khi lưu dữ liệu: {e}")

# --- TAB 5: VIEW & DOWNLOAD ---
with tab5:
    st.subheader("Danh sách toàn bộ Channel trong Database")
    res = supabase.table("channels").select("*").execute()
    if res.data:
        df_all = pd.DataFrame(res.data)
        st.write(f"Tổng số kênh hiện có: **{len(df_all)}**")
        st.dataframe(df_all, use_container_width=True)
        csv = df_all.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Tải về toàn bộ Database (CSV)", data=csv, file_name="master_youtube_database.csv", mime="text/csv")

# --- TAB 6: DEDICATED CHANNEL KEYWORD INSPECTOR ---
with tab6:
    st.subheader("✨ Soi Từ Khóa Kênh (Channel & Video Tags SEO Inspector)")
    st.markdown("Nhập bất kỳ Handle nào để bóc tách **Thẻ từ khóa ẩn của Kênh (Channel Keywords)**, **Top Video Tags** và **Phân loại AI của YouTube**.")
    
    col_k1, col_f2 = st.columns([2, 1])
    with col_k1:
        inspect_handle_input = st.text_input("Nhập Handle Kênh cần soi (ví dụ: @NickDiGiovanni, @dudeperfect):", value="@NickDiGiovanni", key="inspect_input")
    with col_f2:
        api_key_tab6 = st.text_input("YouTube Data API Key:", value=default_api_key_tab3, type="password", key="api_key_tab6")

    if inspect_handle_input and st.button("🔍 Soi Từ Khóa Ngay"):
        pure_inspect = to_pure_id(inspect_handle_input)
        if not pure_inspect:
            st.error("Handle không hợp lệ!")
        else:
            try:
                yt_insp = build("youtube", "v3", developerKey=api_key_tab6)
                cid_insp = get_channel_id_by_handle(yt_insp, pure_inspect)
                
                if not cid_insp:
                    st.error("Không tìm thấy Channel ID cho kênh này!")
                else:
                    with st.spinner("Đang bóc tách dữ liệu từ YouTube Studio & Tags..."):
                        ext_data = extract_channel_master_keywords(yt_insp, cid_inspect)
                        master_str = ", ".join(ext_data['master_keywords'])
                        
                        # Store in pending_keywords for Streamlit lifecycle safety
                        st.session_state['pending_keywords'] = master_str
                        st.session_state['last_inspected_data'] = ext_data
                        st.session_state['last_inspected_handle'] = pure_inspect
                        
            except Exception as e:
                st.error(f"Lỗi khi soi từ khóa: {e}")

    # Display inspection results if present in session_state
    if 'last_inspected_data' in st.session_state:
        ext_data = st.session_state['last_inspected_data']
        pure_inspect = st.session_state.get('last_inspected_handle', '')
        master_str = ", ".join(ext_data['master_keywords'])
        
        st.divider()
        st.success(f"✨ Đã liên kết tự động bộ từ khóa này sang Tab 3 ('Săn Kênh Tương Tự')!")
        st.markdown(f"### 🏷️ Dữ Liệu Từ Khóa Của Kênh `@{pure_inspect}`")
        
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            st.markdown("#### 🔑 Thẻ Từ Khóa Ẩn Của Kênh (Channel Keywords):")
            if ext_data['channel_keywords']:
                for kw in ext_data['channel_keywords']:
                    st.write(f"• `{kw}`")
            else:
                st.info("Kênh này không cài đặt Thẻ từ khóa ẩn.")
                
            st.markdown("#### 📂 Phân Loại Chủ Đề Của YouTube (Topics):")
            if ext_data['categories']:
                for cat in ext_data['categories']:
                    st.write(f"• **{cat}**")
            else:
                st.info("Chưa có thông tin Topic Category.")

        with col_t2:
            st.markdown("#### 📌 Top Video Tags Xuất Hiện Nhiều Nhất:")
            if ext_data['top_tags']:
                for tag in ext_data['top_tags']:
                    st.write(f"• `{tag}`")
            else:
                st.info("Không tìm thấy Video Tags.")

        st.divider()
        st.markdown("#### 🎯 Bộ Từ Khóa Gợi Ý (Đã tự động gửi sang Tab 3):")
        st.code(master_str, language="text")
