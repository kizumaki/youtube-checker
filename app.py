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
from PIL import Image as PILImage
from supabase import create_client, Client

# Page Config
st.set_page_config(page_title="YouTube Channel Master DB", page_icon="📺", layout="wide")

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
    """Extracts pure handle IDs from multi-line or comma-separated text."""
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
    """Extracts handles from uploaded TXT, CSV, or XLSX files."""
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

# --- YOUTUBE DATA API V4.14 SCRAPER ENGINE ---
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

def get_channel_details(youtube, channel_id):
    request = youtube.channels().list(part="snippet,contentDetails,statistics", id=channel_id)
    response = request.execute()
    if 'items' in response and len(response['items']) > 0:
        item = response['items'][0]
        playlist_id = item['contentDetails']['relatedPlaylists']['uploads']
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

        return playlist_id, sub_count, description, joined_date, country_name, avatar_url
    return None, 0, "", "", "", ""

def get_video_details(youtube, video_ids, progress_bar=None):
    video_data = []
    total = len(video_ids)
    
    for i in range(0, total, 50):
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

    # TAB 2: "Top 10 Video Title" DASHBOARD
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
        {"fill": "2F5597", "font": "FFFFFF"},
        {"fill": "C00000", "font": "FFFFFF"},
        {"fill": "70AD47", "font": "FFFFFF"},
        {"fill": "7030A0", "font": "FFFFFF"},
        {"fill": "00C0C0", "font": "FFFFFF"},
        {"fill": "E37222", "font": "FFFFFF"},
        {"fill": "41536B", "font": "FFFFFF"},
        {"fill": "A04000", "font": "FFFFFF"},
        {"fill": "385723", "font": "FFFFFF"},
        {"fill": "626262", "font": "FFFFFF"}
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
st.caption("Hệ thống quản lý, tra cứu hàng loạt, cào kênh live & xuất báo cáo Audit chuẩn 24/7")

# Tabs
tab1, tab2, tab3, tab4 = st.tabs(["🔍 Tra cứu Handle Hàng Loạt", "⚡ Cào Live & Tạo Báo Cáo Audit", "📤 Upload Cập nhật Data", "📊 Xem Database"])

# --- TAB 1: BATCH SEARCH ---
with tab1:
    st.subheader("🔍 Kiểm tra Trùng Lặp Danh Sách Handle Hàng Loạt")
    st.markdown("Dán danh sách Handle/Link hoặc upload file để đối chiếu nhanh với Database Cloud.")
    
    col_s1, col_s2 = st.columns([2, 1])
    
    with col_s1:
        text_input_area = st.text_area(
            "Cách 1: Dán danh sách Handle/Link kênh vào đây (mỗi kênh 1 dòng):",
            placeholder="@MrBeast\n@PewDiePie\nhttps://www.youtube.com/@aCookieGod\n@123GO_",
            height=180
        )
        
    with col_s2:
        file_input_check = st.file_uploader(
            "Cách 2: Hoặc Upload file danh sách (.txt, .csv, .xlsx):",
            type=["txt", "csv", "xlsx", "xls"]
        )
        
    if st.button("🔎 Bắt Đầu Kiểm Tra Hàng Loạt", type="primary"):
        # Combine handles from both inputs
        all_target_handles = set()
        
        if text_input_area:
            for h in extract_handles_from_text(text_input_area):
                all_target_handles.add(h)
                
        if file_input_check:
            for h in extract_handles_from_file(file_input_check):
                all_target_handles.add(h)
                
        target_list = list(all_target_handles)
        
        if not target_list:
            st.warning("⚠️ Vui lòng dán danh sách Handle hoặc chọn file để kiểm tra!")
        else:
            with st.spinner(f"Đang đối chiếu {len(target_list)} Handle với Database Supabase..."):
                # Supabase BATCH query using in_ operator
                response = supabase.table("channels").select("handle, youtuber_name, source").in_("handle", target_list).execute()
                
                db_matches = {item["handle"].lower(): item for item in response.data} if response.data else {}
                
                # Categorize
                new_handles = []
                existing_handles = []
                report_data = []
                
                for h in target_list:
                    if h in db_matches:
                        matched_item = db_matches[h]
                        existing_handles.append({
                            "Handle": f"@{h}",
                            "Tên / Ghi chú": matched_item.get("youtuber_name", "N/A"),
                            "Nguồn Dữ Liệu": matched_item.get("source", "N/A"),
                            "Trạng thái": "❌ Đã có trong DB"
                        })
                        report_data.append({
                            "Handle": f"@{h}",
                            "Trạng Thái": "❌ Đã có",
                            "Tên/Nguồn": matched_item.get("youtuber_name", "N/A")
                        })
                    else:
                        new_handles.append({
                            "Handle": f"@{h}",
                            "Trạng thái": "✅ Kênh Mới (Chưa làm)"
                        })
                        report_data.append({
                            "Handle": f"@{h}",
                            "Trạng Thái": "✅ Mới",
                            "Tên/Nguồn": "Sẵn sàng cào dữ liệu"
                        })

                # Display Metrics
                st.divider()
                m1, m2, m3 = st.columns(3)
                m1.metric("Tổng số kênh kiểm tra", f"{len(target_list)} kênh")
                m2.metric("❌ Kênh Đã Tồn Tại", f"{len(existing_handles)} kênh")
                m3.metric("✅ Kênh Mới Có Thể Làm", f"{len(new_handles)} kênh")
                
                # Results Tabs
                res_tab1, res_tab2 = st.tabs([f"✅ Kênh Mới Chưa Làm ({len(new_handles)})", f"❌ Kênh Đã Tồn Tại ({len(existing_handles)})"])
                
                with res_tab1:
                    if new_handles:
                        df_new = pd.DataFrame(new_handles)
                        st.dataframe(df_new, use_container_width=True)
                        
                        # Prepare plain text list for download
                        txt_content = "\n".join([item["Handle"] for item in new_handles])
                        
                        col_dl1, col_dl2 = st.columns(2)
                        with col_dl1:
                            st.download_button(
                                label="📥 Tải Danh Sách Kênh Mới (.txt)",
                                data=txt_content,
                                file_name=f"danh_sach_kenh_moi_{datetime.date.today().strftime('%d-%m-%Y')}.txt",
                                mime="text/plain"
                            )
                        with col_dl2:
                            # Excel format download
                            buf_new = io.BytesIO()
                            df_new.to_excel(buf_new, index=False)
                            st.download_button(
                                label="📥 Tải Danh Sách Kênh Mới (.xlsx)",
                                data=buf_new.getvalue(),
                                file_name=f"danh_sach_kenh_moi_{datetime.date.today().strftime('%d-%m-%Y')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.info("Tất cả các kênh trong danh sách của bạn đều **đã tồn tại** trong Database!")

                with res_tab2:
                    if existing_handles:
                        df_exist = pd.DataFrame(existing_handles)
                        st.dataframe(df_exist, use_container_width=True)
                    else:
                        st.success("🎉 Tuyệt vời! Không có kênh nào trùng lặp trong danh sách này!")
                        
                # Overall report download
                st.divider()
                df_report = pd.DataFrame(report_data)
                buf_rep = io.BytesIO()
                df_report.to_excel(buf_rep, index=False)
                
                st.download_button(
                    label="📊 Tải Báo Cáo Kiểm Tra Tổng Hợp (.xlsx)",
                    data=buf_rep.getvalue(),
                    file_name=f"bao_cao_kiem_tra_trung_lap_{datetime.date.today().strftime('%d-%m-%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

# --- TAB 2: LIVE API SCRAPER & V4.14 AUDIT REPORT ---
with tab2:
    st.subheader("⚡ Cào dữ liệu Live & Xuất Báo Cáo Audit chuẩn V4.14")
    st.markdown("Cào chính thức qua **YouTube Data API v3** (lấy đủ 1000+ video, Avatar, Subscriptions, Views, Lengths, v.v.).")
    
    col_input1, col_input2 = st.columns([2, 1])
    with col_input1:
        channel_url_input = st.text_input("Dán Link kênh hoặc Handle vào đây:", value="@4wd247", placeholder="https://www.youtube.com/@4wd247 hoặc @4wd247")
    with col_input2:
        default_api_key = st.secrets.get("YOUTUBE_API_KEY", "AIzaSyBrTtmMp-txQ7ID15wrJZUpN-i53SRVzgk")
        api_key_input = st.text_input("YouTube Data API Key (Tùy chọn):", value=default_api_key, type="password")

    if channel_url_input and st.button("🚀 Xử lý Kênh & Tạo Báo Cáo V4.14"):
        pure_h = to_pure_id(channel_url_input)
        if not pure_h:
            st.error("Handle hoặc đường link không hợp lệ!")
        else:
            try:
                st.info("🔄 Đang kết nối YouTube Data API v3...")
                youtube = build("youtube", "v3", developerKey=api_key_input)
                
                channel_id = get_channel_id_by_handle(youtube, pure_h)
                if not channel_id:
                    st.error("❌ Không tìm thấy Channel ID cho kênh này trên YouTube!")
                else:
                    playlist_id, sub_count, channel_desc, channel_joined, channel_country, avatar_url = get_channel_details(youtube, channel_id)
                    
                    st.info("📥 Đang trích xuất toàn bộ danh sách Video trong Playlist...")
                    video_ids = []
                    next_page_token = None
                    
                    while True:
                        req = youtube.playlistItems().list(part="snippet", playlistId=playlist_id, maxResults=50, pageToken=next_page_token)
                        res = req.execute()
                        for item in res['items']:
                            video_ids.append(item['snippet']['resourceId']['videoId'])
                        next_page_token = res.get('nextPageToken')
                        if not next_page_token:
                            break
                            
                    st.write(f"📊 Tìm thấy **{len(video_ids)}** video. Đang lấy chi tiết thời lượng & số lượt xem...")
                    prog_bar = st.progress(0.0)
                    video_data = get_video_details(youtube, video_ids, progress_bar=prog_bar)
                    
                    # Auto Upsert into Supabase
                    supabase.table("channels").upsert([{
                        "handle": pure_h,
                        "youtuber_name": pure_h.upper(),
                        "source": "YouTube API V4.14"
                    }], on_conflict="handle").execute()
                    
                    # Show Overview UI
                    st.divider()
                    col_res1, col_res2 = st.columns([1, 2])
                    
                    latest_date = video_data[0]['Published Date'] if video_data else "N/A"
                    is_act = is_within_last_90_days(latest_date)
                    
                    with col_res1:
                        if avatar_url:
                            st.image(avatar_url, width=150)
                    with col_res2:
                        st.markdown(f"### 🎯 **@{pure_h}**")
                        st.write(f"• **Số lượng Video:** `{len(video_data):,}`")
                        st.write(f"• **Lượt đăng ký (Subs):** `{sub_count:,}`")
                        st.write(f"• **Quốc gia:** `{channel_country}` | **Ngày tham gia:** `{channel_joined}`")
                        st.write(f"• **Video gần nhất:** `{latest_date}`")
                        
                        if is_act:
                            st.success("🟢 **Trạng thái:** Hoạt động tích cực (Đăng video dài < 90 ngày)")
                        else:
                            st.warning("⚠️ **Trạng thái:** Không hoạt động quá 90 ngày")

                    # Generate V4.14 Excel
                    st.info("📈 Đang khởi tạo file Excel Audit V4.14...")
                    excel_v414_bytes = generate_v414_excel_report(
                        clean_handle=pure_h,
                        sub_count=sub_count,
                        channel_desc=channel_desc,
                        channel_joined=channel_joined,
                        channel_country=channel_country,
                        avatar_url=avatar_url,
                        video_data=video_data
                    )
                    
                    date_now_str = datetime.datetime.now().strftime("%d-%m-%Y")
                    output_file_name = f"{pure_h}_{date_now_str}.xlsx"
                    
                    st.download_button(
                        label=f"📥 Tải về File Báo Cáo Audit Chuẩn V4.14 ({output_file_name})",
                        data=excel_v414_bytes,
                        file_name=output_file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f"❌ Có lỗi xảy ra trong quá trình gọi YouTube API: {e}")

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
