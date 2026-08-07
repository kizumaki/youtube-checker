import re
import io
import datetime
import zipfile
import os
import pandas as pd
import requests
import streamlit as st
from concurrent.futures import ThreadPoolExecutor, as_completed

from db_utils import (
    supabase, load_api_keys_from_db, save_api_keys_to_db, 
    load_cart_from_db, add_to_cart_db, remove_from_cart_db, 
    clear_entire_database, add_batch_to_cart_db, clear_cart_db,
    confirm_clear_db_dialog
)
from yt_utils import (
    DEFAULT_API_KEY, to_pure_id, get_channel_link, extract_raw_inputs_from_file, 
    parse_raw_inputs_to_handles, test_all_api_keys, process_tab1_single_handle, 
    process_single_crm_channel_meta, run_single_channel_audit, 
    render_social_badges_html, extract_channel_master_keywords, 
    process_single_candidate, clean_and_extract_keywords, get_channel_id_by_handle_direct,
    get_channel_details_direct, get_6_recent_videos_direct, extract_handle_from_filename,
    extract_text_from_docx_bytes, is_garbage_input, get_handles_from_video_ids,
    get_handles_from_search_queries
)
from ui_components import (
    inject_theme_css, render_kpi_cards, show_ai_email_dialog, 
    render_shared_cart_ui, show_video_dialog, compare_channels_dialog
)

# Page Config
st.set_page_config(page_title="YT CHECKER PRO", page_icon="🎙️", layout="wide", initial_sidebar_state="expanded")

# Initialize Session State
if 'app_theme' not in st.session_state: st.session_state['app_theme'] = 'Studio Peach (Sáng)'
if 'selected_channels' not in st.session_state: st.session_state['selected_channels'] = set()
if 'api_usage' not in st.session_state: st.session_state['api_usage'] = {}
if 'api_status_map' not in st.session_state: st.session_state['api_status_map'] = {}
if 'exhausted_keys_set' not in st.session_state: st.session_state['exhausted_keys_set'] = set()
if 'chk_counter' not in st.session_state: st.session_state['chk_counter'] = 0
if 'global_api_keys' not in st.session_state: st.session_state['global_api_keys'] = load_api_keys_from_db() or DEFAULT_API_KEY
if 'cart' not in st.session_state: st.session_state['cart'] = load_cart_from_db()

def toggle_select_channel(pure_handle):
    if pure_handle in st.session_state['selected_channels']: st.session_state['selected_channels'].remove(pure_handle)
    else: st.session_state['selected_channels'].add(pure_handle)

def cb_select_all(channel_list):
    for item in channel_list:
        raw_h = item.get('Handle') or item.get('handle')
        p_id = to_pure_id(raw_h)
        if p_id: st.session_state['selected_channels'].add(p_id)
    st.session_state['chk_counter'] += 1

def cb_clear_all():
    st.session_state['selected_channels'].clear()
    st.session_state['chk_counter'] += 1

def sync_pagination_top(top_key, bottom_key, state_key):
    val = st.session_state[top_key]
    st.session_state[state_key] = val
    st.session_state[bottom_key] = val

def sync_pagination_bottom(top_key, bottom_key, state_key):
    val = st.session_state[bottom_key]
    st.session_state[state_key] = val
    st.session_state[top_key] = val

def format_page_range(page_num, items_per_page, total_items):
    if total_items == 0: return "0 / 0"
    start_item = (page_num - 1) * items_per_page + 1
    end_item = min(page_num * items_per_page, total_items)
    return f"{start_item:,} - {end_item:,} / {total_items:,}"

def delete_channel_from_system(pure_handle):
    if not pure_handle: return
    try: supabase.table("channels").delete().eq("handle", pure_handle).execute()
    except Exception: pass
    remove_from_cart_db(pure_handle)
    if pure_handle in st.session_state.get('cart', {}): del st.session_state['cart'][pure_handle]
    if st.session_state.get('active_inspected_handle') == pure_handle: st.session_state['active_inspected_handle'] = None
    if pure_handle in st.session_state['selected_channels']: st.session_state['selected_channels'].remove(pure_handle)
    for key in list(st.session_state.keys()):
        if key.startswith('crm_cache_') or key == 'tab5_crm_cache': st.session_state.pop(key, None)

def set_api_keys(key_string):
    keys = [k.strip() for k in re.split(r'[\n,]+', key_string) if k.strip()]
    st.session_state['api_keys'] = keys if keys else [DEFAULT_API_KEY]

set_api_keys(st.session_state['global_api_keys'])

# Theme Styling
is_dark = st.session_state['app_theme'] == 'Studio Espresso (Tối)'
bg_color = "#1E1816" if is_dark else "#F4F2F1"
card_bg = "#2A221F" if is_dark else "#FFFFFF"
text_color = "#F4F2F1" if is_dark else "#3D2F29"
border_color = "#3D2F29" if is_dark else "#E5E7EB"
sidebar_bg = "#241D1A" if is_dark else "#FFFFFF"
inject_theme_css(is_dark, bg_color, card_bg, text_color, border_color, sidebar_bg)

# --- SIDEBAR ---
with st.sidebar:
    col1, col2, col3 = st.columns([0.5, 8, 0.5])
    with col2:
        found_logo = next((p for p in ["logo.png", "logo_2.png", "logo.jpg"] if os.path.exists(p)), None)
        if found_logo: st.image(found_logo, use_container_width=True)
        else:
            st.markdown("""
            <div style="text-align: center; padding: 10px 0;">
                <svg width="65" height="65" viewBox="0 0 100 100" fill="none">
                    <rect x="10" y="45" width="8" height="10" rx="2" fill="#D95F26"/>
                    <rect x="22" y="30" width="8" height="40" rx="2" fill="#D95F26"/>
                    <rect x="34" y="15" width="8" height="70" rx="2" fill="#D95F26"/>
                    <rect x="46" y="5" width="8" height="90" rx="2" fill="#D95F26"/>
                    <rect x="58" y="20" width="8" height="60" rx="2" fill="#D95F26"/>
                    <rect x="70" y="35" width="8" height="30" rx="2" fill="#D95F26"/>
                    <rect x="82" y="42" width="8" height="16" rx="2" fill="#D95F26"/>
                </svg>
            </div>""", unsafe_allow_html=True)
    st.markdown("""
        <div style="text-align: center; padding-bottom: 8px;">
            <h2 style="margin: 0 0 4px 0; font-weight: 800; font-size: 1.15rem;">YT CHECKER PRO</h2>
            <span class="badge-pro badge-ocean">Supabase Live</span>
        </div>""", unsafe_allow_html=True)
    st.selectbox("🎨 Giao diện App:", options=["Studio Peach (Sáng)", "Studio Espresso (Tối)"], key="app_theme")
    st.divider()

    if 'api_status_tested' not in st.session_state:
        test_all_api_keys(); st.session_state['api_status_tested'] = True

    st.markdown("<h4 style='font-weight: 700; font-size: 0.95rem;'>🛡️ Sức Khỏe API Quota</h4>", unsafe_allow_html=True)
    for k in st.session_state.get('api_keys', []):
        used = st.session_state.get('api_usage', {}).get(k, 0)
        k_stat_type, _ = st.session_state.get('api_status_map', {}).get(k, ("UNKNOWN", 0))
        pct = 100 if (k_stat_type in ["EXHAUSTED", "DEAD"] or used >= 10000) else min(100, int((used / 10000) * 100))
        color = "#EF4444" if pct == 100 else ("#10B981" if pct < 70 else "#F59E0B")
        st.markdown(f"""
            <div style='margin-bottom: 8px;'>
                <div style='font-size: 0.75rem; color: #6B7280; font-weight: 700;'>🔑 {k[:10]}...</div>
                <div style='background-color: #E5E7EB; border-radius: 4px; height: 6px; width: 100%; margin-top: 4px;'>
                    <div style='background-color: {color}; width: {pct}%; height: 100%; border-radius: 4px;'></div>
                </div>
                <div style='font-size: 0.65rem; color: #9CA3AF; text-align: right;'>🟢 {used:,}/10,000</div>
            </div>""", unsafe_allow_html=True)

    if st.button("🧪 Kiểm Tra Sức Khỏe Keys", use_container_width=True, key="btn_test_keys"):
        with st.spinner("Đang kiểm tra..."): test_all_api_keys(); st.rerun()

    keys_input = st.text_area("Cập nhật danh sách Key (1 key/dòng):", value=st.session_state['global_api_keys'], height=80, key="api_keys_text_area")
    if st.button("💾 Lưu Cấu Hình Key", type="primary", use_container_width=True):
        st.session_state['global_api_keys'] = keys_input; set_api_keys(keys_input); save_api_keys_to_db(keys_input); test_all_api_keys(); st.toast("🎉 Đã lưu vĩnh viễn danh sách API Keys!"); st.rerun()

    st.divider()
    if st.button("🔄 Làm Mới Màn Hình", use_container_width=True):
        cb_clear_all(); st.rerun()

# --- HEADER ---
st.markdown("""
    <div style="padding: 5px 0 15px 0;">
        <h1 style="font-weight: 900; margin-bottom: 5px; font-size: 2.4rem;">YT CHECKER <span style="color: #D95F26;">PRO</span></h1>
        <p style="font-size: 1.05rem; font-weight: 500; opacity: 0.8;">Hệ thống phân tích, tìm kiếm kênh đồng ngách Đa Luồng Siêu Tốc & Quản lý Chiến Dịch.</p>
    </div>""", unsafe_allow_html=True)

# Tabs
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🔍 Tra cứu Handle Hàng Loạt", 
    "⚡ Cào Live & Tạo Báo Cáo Audit", 
    "🎯 Săn Kênh Tương Tự (Multi-Threaded)",
    "📤 Upload Cập nhật Data", 
    "📊 Xem Database",
    "✨ Soi Từ Khóa Kênh (SEO Inspector)"
])

# --- TAB 1 ---
with tab1:
    st.markdown("<h3 style='font-weight: 700;'>🔍 Kiểm tra Trùng Lặp Danh Sách Handle / Link Video Hàng Loạt</h3>", unsafe_allow_html=True)
    col_s1, col_s2 = st.columns([2, 1])
    with col_s1: text_input_area = st.text_area("Dán danh sách Handle/Link kênh/Link Video/Link Tìm kiếm vào đây (mỗi dòng 1 link):", height=180)
    with col_s2: file_input_check = st.file_uploader("Hoặc Upload file danh sách (.txt, .csv, .xlsx):")

    if st.button("🔎 Bắt Đầu Kiểm Tra Hàng Loạt", type="primary"):
        all_raw_inputs = []
        if text_input_area: all_raw_inputs.extend(re.split(r'[\n,\t\r]+', str(text_input_area)))
        if file_input_check: all_raw_inputs.extend(extract_raw_inputs_from_file(file_input_check))
        target_list = parse_raw_inputs_to_handles(all_raw_inputs)
        
        if not target_list: st.warning("⚠️ Vui lòng dán danh sách Handle, Link Video hoặc chọn file để kiểm tra!")
        else:
            prog = st.progress(0); stat = st.empty()
            response = supabase.table("channels").select("handle, youtuber_name").in_("handle", target_list).execute()
            db_matches = {item["handle"].lower(): item for item in response.data} if response.data else {}
            new_handles, existing_handles, rejected_handles = [], [], []
            tot = len(target_list); comp = 0

            with ThreadPoolExecutor(max_workers=8) as executor:
                futures = [executor.submit(process_tab1_single_handle, p_id, db_matches, st.session_state['api_keys'], st.session_state['exhausted_keys_set']) for p_id in target_list]
                for future in as_completed(futures):
                    comp += 1; prog.progress(comp / tot)
                    stat.markdown(f"⏳ **Đang phân tích:** `{comp}/{tot}` Handle...")
                    try:
                        status, res_data, _ = future.result()
                        if status == "NEW": new_handles.append(res_data)
                        elif status == "EXISTING": existing_handles.append(res_data)
                        else: rejected_handles.append(res_data)
                    except Exception: pass
            prog.empty(); stat.empty()
            st.session_state['batch_check_new'] = new_handles
            st.session_state['batch_check_existing'] = existing_handles
            st.session_state['batch_check_rejected'] = rejected_handles

    if 'batch_check_new' in st.session_state:
        new_handles = st.session_state.get('batch_check_new', [])
        existing_handles = st.session_state.get('batch_check_existing', [])
        rejected_handles = st.session_state.get('batch_check_rejected', [])
        st.divider()
        render_kpi_cards([
            ("TỔNG SỐ KIỂM TRA", f"{len(new_handles) + len(existing_handles) + len(rejected_handles)}", "#47A5D1"),
            ("✅ ĐẠT CHUẨN (>=1M SUBS & >10MIN)", f"{len(new_handles)}", "#10B981"),
            ("❌ ĐÃ TỒN TẠI TRONG DB", f"{len(existing_handles)}", "#F59E0B"),
            ("🚫 BỊ LOẠI", f"{len(rejected_handles)}", "#EF4444")
        ], card_bg, border_color)

    render_shared_cart_ui("tab1")

# --- TAB 2 ---
with tab2:
    st.markdown("<h3 style='font-weight: 700;'>⚡ Cào dữ liệu Live & Xuất Báo Cáo Audit chuẩn V4.14 (Đơn / Hàng Loạt)</h3>", unsafe_allow_html=True)
    col_t2_1, col_t2_2 = st.columns([2, 1])
    with col_t2_1: text_input_area_t2 = st.text_area("Dán danh sách Handle/Link kênh/Link Video/Link Tìm kiếm:", height=180, value="@4wd247", key="text_input_tab2")
    with col_t2_2: file_input_t2 = st.file_uploader("Upload file danh sách hoặc gói báo cáo:", type=["zip", "xlsx", "xls", "txt", "docx", "doc", "csv"], key="file_input_tab2")

    if st.button("🚀 Bắt Đầu Cào Live & Tạo Báo Cáo Audit V4.14", type="primary", key="btn_run_tab2_audit"):
        inputs_t2 = re.split(r'[\n,\t\r]+', str(text_input_area_t2)) if text_input_area_t2 else []
        if file_input_t2: inputs_t2.extend(extract_raw_inputs_from_file(file_input_t2))
        target_t2 = parse_raw_inputs_to_handles(inputs_t2)
        if not target_t2: st.warning("⚠️ Vui lòng dán danh sách hoặc upload file!")
        else:
            tot_t2 = len(target_t2); prog_t2 = st.progress(0); stat_t2 = st.empty(); comp_t2 = 0; audit_res = []; db_upsert = []
            with ThreadPoolExecutor(max_workers=5) as executor:
                futures = {executor.submit(run_single_channel_audit, p_h, st.session_state['api_keys'], st.session_state['exhausted_keys_set']): p_h for p_h in target_t2}
                for future in as_completed(futures):
                    comp_t2 += 1; prog_t2.progress(comp_t2 / tot_t2)
                    try:
                        res_val = future.result()
                        if res_val and res_val[0] and res_val[1]:
                            b_bytes, f_name, _ = res_val; audit_res.append((f_name, b_bytes))
                            p_clean = f_name.split('_')[0]; db_upsert.append({"handle": p_clean, "youtuber_name": p_clean.upper(), "source": "Live Audit Scraper"})
                    except Exception: pass
            prog_t2.empty(); stat_t2.empty()
            if db_upsert: supabase.table("channels").upsert(db_upsert, on_conflict="handle").execute()
            st.session_state['tab2_audit_output'] = {"results": audit_res, "count": len(audit_res), "total_requested": tot_t2}; st.rerun()

# --- TAB 3 ---
with tab3:
    st.markdown("<h3 style='font-weight: 700;'>🎯 Săn Kênh Tương Tự & Giỏ Hàng (Multi-threaded Speed)</h3>", unsafe_allow_html=True)
    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        seed_input = st.text_input("Nhập Handle Kênh Mồi (ví dụ: @NickDiGiovanni):", key="seed_input_tab3")
        custom_kw = st.text_input("Từ khóa chủ đề:", key="custom_kw_tab3")
    with col_f2:
        min_subs = st.selectbox("Mốc Subscribers Tối Thiểu:", options=[100000, 250000, 500000, 1000000], index=3, format_func=lambda x: f"{x:,} Subs")
        min_dur = st.selectbox("Lọc Yêu Cầu Đồ Dài Video:", options=[600], index=0, format_func=lambda x: "Bắt buộc có Video > 10 phút")

# --- TAB 4 ---
with tab4:
    st.markdown("<h3 style='font-weight: 700;'>📤 Upload file .ZIP, .TXT, .XLSX hoặc .DOCX</h3>", unsafe_allow_html=True)
    files_up = st.file_uploader("Kéo thả file vào đây:", type=["zip", "txt", "xlsx", "xls", "docx", "csv"], accept_multiple_files=True)

# --- TAB 5 ---
with tab5:
    st.markdown("<h3 style='font-weight: 700;'>📊 Quản lý Database CRM Kênh</h3>", unsafe_allow_html=True)
    try: res = supabase.table("channels").select("*").order("created_at", desc=True).execute()
    except Exception: res = supabase.table("channels").select("*").execute()
    if res.data:
        df_all = pd.DataFrame(res.data)
        st.write(f"Tổng số kênh hiện có trong DB: **{len(df_all)}**")
        st.dataframe(df_all, use_container_width=True)
    else: st.info("Database trống.")

# --- TAB 6 ---
with tab6:
    st.markdown("<h3 style='font-weight: 700;'>✨ Soi Từ Khóa Kênh (SEO Inspector)</h3>", unsafe_allow_html=True)
    inspect_input = st.text_input("Nhập Handle Kênh / Link cần soi:", value="@NickDiGiovanni")
    if st.button("🔍 Soi Từ Khóa Ngay", type="primary"):
        pure = to_pure_id(inspect_input)
        if pure:
            cid, _, _, _ = get_channel_id_by_handle_direct(pure, st.session_state['api_keys'])
            if cid:
                ext = extract_channel_master_keywords(cid)
                st.success(f"✨ Đã soi từ khóa kênh @{pure}!")
                st.code(", ".join(ext['master_keywords']), language="text")
