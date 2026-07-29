import streamlit as st
import pandas as pd
import zipfile
import os
import re
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

# Helper Functions
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

def extract_handle_from_filename(filename):
    base = os.path.basename(filename)
    base_no_ext = os.path.splitext(base)[0]
    pattern = r'_(?:backlog|\d{4}|\d{2,4}[-_/.]\d{1,2}[-_/.]\d{1,2}|\d{1,2}[-_/.]\d{1,2}[-_/.]\d{2,4}|\d{6,8})(?:_.*)?$'
    cleaned = re.sub(pattern, '', base_no_ext, flags=re.IGNORECASE)
    cleaned = re.sub(r'[\s]+', '', cleaned)
    pure_id = re.sub(r'^@+', '', cleaned).strip().lower()
    return pure_id if pure_id else None

# App UI Header
st.title("📺 YouTube Channel Master Database")
st.caption("Hệ thống quản lý, tra cứu & gọt trùng Handle YouTube đám mây 24/7")

# Tabs
tab1, tab2, tab3 = st.tabs(["🔍 Tra cứu Handle", "📤 Upload Cập nhật Data", "📊 Xem Database"])

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

# --- TAB 2: UPLOAD & UPDATE ---
with tab2:
    st.subheader("Upload file .ZIP hoặc .TXT để cập nhật Database")
    uploaded_files = st.file_uploader("Kéo thả file `.zip` (chứa các báo cáo Excel) hoặc file `.txt` vào đây:", type=["zip", "txt", "xlsx"], accept_multiple_files=True)
    
    if uploaded_files and st.button("🚀 Bắt đầu xử lý & Nạp vào Database"):
        new_handles_to_insert = []
        
        for file in uploaded_files:
            file_name = file.name
            
            # Handle ZIP
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
                                    
            # Handle TXT
            elif file_name.endswith('.txt'):
                content = file.read().decode("utf-8", errors="ignore")
                for line in content.splitlines():
                    h = to_pure_id(line)
                    if h:
                        new_handles_to_insert.append({"handle": h, "youtuber_name": h, "source": file_name})

        if new_handles_to_insert:
            # Batch upsert to Supabase
            df_insert = pd.DataFrame(new_handles_to_insert).drop_duplicates(subset=["handle"])
            data_dict = df_insert.to_dict(orient="records")
            
            try:
                # Upsert ignore duplicates
                response = supabase.table("channels").upsert(data_dict, on_conflict="handle").execute()
                st.success(f"🎉 Đã xử lý & đồng bộ thành công {len(data_dict)} Handle vào Database đám mây!")
            except Exception as e:
                st.error(f"Lỗi khi lưu dữ liệu: {e}")

# --- TAB 3: VIEW & DOWNLOAD ---
with tab3:
    st.subheader("Danh sách toàn bộ Channel trong Database")
    res = supabase.table("channels").select("*").execute()
    if res.data:
        df_all = pd.DataFrame(res.data)
        st.write(f"Tổng số kênh hiện có: **{len(df_all)}**")
        st.dataframe(df_all, use_container_width=True)
        
        # Download Excel
        csv = df_all.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Tải về toàn bộ Database (CSV)", data=csv, file_name="master_youtube_database.csv", mime="text/csv")
