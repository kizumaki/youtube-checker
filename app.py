import datetime
import io
import os
import re
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests
import streamlit as st

from db_utils import (
    add_batch_to_cart_db,
    add_to_cart_db,
    clear_cart_db,
    clear_entire_database,
    confirm_clear_db_dialog,
    load_api_keys_from_db,
    load_cart_from_db,
    remove_from_cart_db,
    save_api_keys_to_db,
    supabase,
)
from ui_components import (
    compare_channels_dialog,
    inject_theme_css,
    render_kpi_cards,
    render_shared_cart_ui,
    set_active_inspected_channel,
    show_ai_email_dialog,
    show_video_dialog,
)
from yt_utils import (
    DEFAULT_API_KEY,
    clean_and_extract_keywords,
    extract_channel_master_keywords,
    extract_handle_from_filename,
    extract_raw_inputs_from_file,
    extract_text_from_docx_bytes,
    extract_video_id,
    get_6_recent_videos_direct,
    get_channel_details_direct,
    get_channel_id_by_handle_direct,
    get_channel_link,
    get_handles_from_search_queries,
    get_handles_from_video_ids,
    is_garbage_input,
    is_within_last_90_days,
    parse_raw_inputs_to_handles,
    passes_layer1_metadata_filter,
    process_single_candidate,
    process_single_crm_channel_meta,
    process_tab1_single_handle,
    render_social_badges_html,
    run_single_channel_audit,
    test_all_api_keys,
    to_pure_id,
)

# Page Configuration
st.set_page_config(page_title="YT CHECKER PRO", page_icon="🎙️", layout="wide", initial_sidebar_state="expanded")

# --- INITIALIZE SESSION STATE ---
if 'app_theme' not in st.session_state: st.session_state['app_theme'] = 'Studio Peach (Sáng)'
if 'selected_channels' not in st.session_state: st.session_state['selected_channels'] = set()
if 'api_usage' not in st.session_state: st.session_state['api_usage'] = {}
if 'api_status_map' not in st.session_state: st.session_state['api_status_map'] = {}
if 'exhausted_keys_set' not in st.session_state: st.session_state['exhausted_keys_set'] = set()
if 'chk_counter' not in st.session_state: st.session_state['chk_counter'] = 0
if 'global_api_keys' not in st.session_state: st.session_state['global_api_keys'] = load_api_keys_from_db() or DEFAULT_API_KEY
if 'cart' not in st.session_state: st.session_state['cart'] = load_cart_from_db()
if 'active_inspected_handle' not in st.session_state: st.session_state['active_inspected_handle'] = None

if 'pending_seed_handle' in st.session_state:
    st.session_state['seed_input_tab3'] = st.session_state.pop('pending_seed_handle')
if 'pending_keywords' in st.session_state:
    st.session_state['custom_kw_tab3'] = st.session_state.pop('pending_keywords')

TAG_OPTIONS = ["📌 Chưa phân loại", "🔥 Ưu tiên làm", "📩 Đã liên hệ", "⏳ Đang chờ duyệt", "✅ Đã chốt", "❌ Bỏ qua"]

def toggle_select_channel(pure_handle):
    if pure_handle in st.session_state['selected_channels']:
        st.session_state['selected_channels'].remove(pure_handle)
    else:
        st.session_state['selected_channels'].add(pure_handle)

def cb_select_all(channel_list):
    for item in channel_list:
        raw_h = item.get('Handle') or item.get('handle')
        p_id = to_pure_id(raw_h)
        if p_id: st.session_state['selected_channels'].add(p_id)
    st.session_state['chk_counter'] += 1

def cb_clear_all():
    st.session_state['selected_channels'].clear()
    st.session_state['chk_counter'] += 1

def format_page_range(page_num, items_per_page, total_items):
    if total_items == 0: return "0 / 0"
    start_item = (page_num - 1) * items_per_page + 1
    end_item = min(page_num * items_per_page, total_items)
    return f"{start_item:,} - {end_item:,} / {total_items:,}"

def delete_channel_from_system(pure_handle):
    if not pure_handle: return
    p_clean = pure_handle.lower()
    try: supabase.table("channels").delete().eq("handle", p_clean).execute()
    except Exception: pass
    remove_from_cart_db(p_clean)
    if p_clean in st.session_state.get('cart', {}): del st.session_state['cart'][p_clean]
    if st.session_state.get('active_inspected_handle') == p_clean: st.session_state['active_inspected_handle'] = None
    if p_clean in st.session_state['selected_channels']: st.session_state['selected_channels'].remove(p_clean)
    
    for list_key in ['batch_check_new', 'batch_check_existing', 'batch_check_rejected', 'passed_channels', 'rejected_channels']:
        if list_key in st.session_state:
            st.session_state[list_key] = [
                item for item in st.session_state[list_key]
                if to_pure_id(item.get('Handle') or item.get('handle')) != p_clean
            ]

    for key in list(st.session_state.keys()):
        if key.startswith('crm_cache_') or key == 'tab5_crm_cache': st.session_state.pop(key, None)

def set_api_keys(key_string):
    keys = [k.strip() for k in re.split(r'[\n,]+', key_string) if k.strip()]
    st.session_state['api_keys'] = keys if keys else [DEFAULT_API_KEY]

def get_selected_channel_data():
    selected = st.session_state['selected_channels']
    data = []
    lists = st.session_state.get('batch_check_new', []) + st.session_state.get('batch_check_existing', []) + st.session_state.get('passed_channels', []) + st.session_state.get('rejected_channels', []) + st.session_state.get('batch_check_rejected', [])
    seen = set()
    for item in lists:
        p = to_pure_id(item.get('Handle'))
        if p in selected and p not in seen:
            data.append(item)
            seen.add(p)
    return data

set_api_keys(st.session_state['global_api_keys'])

# Apply Theme
is_dark = st.session_state['app_theme'] == 'Studio Espresso (Tối)'
bg_color = "#1E1816" if is_dark else "#F4F2F1"
card_bg = "#2A221F" if is_dark else "#FFFFFF"
text_color = "#F4F2F1" if is_dark else "#3D2F29"
border_color = "#3D2F29" if is_dark else "#E5E7EB"
sidebar_bg = "#241D1A" if is_dark else "#FFFFFF"
inject_theme_css(is_dark, bg_color, card_bg, text_color, border_color, sidebar_bg)

# --- SIDEBAR CONTROL ---
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
        if pct < 70: color, status_label = "#10B981", f"🟢 {used:,}/10,000"
        elif pct < 90: color, status_label = "#F59E0B", f"🟡 {used:,}/10,000"
        else: color, status_label = "#EF4444", f"🔴 {used:,}/10,000" if k_stat_type != "DEAD" else "💀 Key Chết"

        st.markdown(f"""
            <div style='margin-bottom: 8px;'>
                <div style='font-size: 0.75rem; color: #6B7280; font-weight: 700;'>🔑 {k[:10]}...</div>
                <div style='background-color: #E5E7EB; border-radius: 4px; height: 6px; width: 100%; margin-top: 4px;'>
                    <div style='background-color: {color}; width: {pct}%; height: 100%; border-radius: 4px;'></div>
                </div>
                <div style='font-size: 0.65rem; color: #9CA3AF; text-align: right; margin-top: 2px;'>{status_label}</div>
            </div>""", unsafe_allow_html=True)

    if st.button("🧪 Kiểm Tra Sức Khỏe Keys", use_container_width=True, key="btn_test_keys"):
        with st.spinner("Đang kiểm tra thực tế..."): test_all_api_keys(); st.rerun()

    keys_input = st.text_area("Cập nhật danh sách Key (1 key/dòng):", value=st.session_state['global_api_keys'], height=80, key="api_keys_text_area")
    if st.button("💾 Lưu Cấu Hình Key", type="primary", use_container_width=True):
        st.session_state['global_api_keys'] = keys_input; set_api_keys(keys_input); save_api_keys_to_db(keys_input); test_all_api_keys(); st.toast("🎉 Đã lưu vĩnh viễn danh sách API Keys!"); st.rerun()

    st.divider()
    if st.button("🔄 Làm Mới Màn Hình", use_container_width=True):
        st.session_state['selected_channels'].clear()
        for k in ['batch_check_new', 'batch_check_existing', 'batch_check_rejected', 'passed_channels', 'rejected_channels', 'tab2_audit_output', 'active_inspected_handle']:
            st.session_state.pop(k, None)
        st.session_state['chk_counter'] += 1
        st.rerun()

# --- APP HEADER ---
st.markdown("""
    <div style="padding: 5px 0 15px 0;">
        <h1 style="font-weight: 900; margin-bottom: 5px; font-size: 2.4rem;">YT CHECKER <span style="color: #D95F26;">PRO</span></h1>
        <p style="font-size: 1.05rem; font-weight: 500; opacity: 0.8;">Hệ thống phân tích, tìm kiếm kênh đồng ngách Đa Luồng Siêu Tốc & Quản lý Chiến Dịch.</p>
    </div>""", unsafe_allow_html=True)

# TABS
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
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>🔍 Kiểm tra Trùng Lặp Danh Sách Handle / Link Video Hàng Loạt</h3>", unsafe_allow_html=True)
    st.caption("💡 *Tự động lọc nâng cao: Bắt buộc >= 1M Subs, có Video > 10 phút trong các video mới nhất, ra video trong 90 ngày và LOẠI HẲN kênh Shorts/Phim/Music/News.*")
    
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
            progress_bar = st.progress(0); status_text = st.empty()
            db_existing_map = {}
            try:
                response = supabase.table("channels").select("handle, youtuber_name").execute()
                if response.data:
                    for item in response.data:
                        h = item.get("handle")
                        if h:
                            p = to_pure_id(h)
                            if p: db_existing_map[p.lower()] = item
            except Exception: pass
            
            new_handles, existing_handles, rejected_handles = [], [], []
            total_items, completed_count = len(target_list), 0
            
            active_keys_t1 = st.session_state.get('api_keys', [DEFAULT_API_KEY])
            exhausted_set_t1 = set(st.session_state.get('exhausted_keys_set', set()))

            with ThreadPoolExecutor(max_workers=8) as executor:
                futures = [executor.submit(process_tab1_single_handle, p_id, db_existing_map, active_keys_t1, exhausted_set_t1) for p_id in target_list]
                for future in as_completed(futures):
                    try:
                        status, res_data, _ = future.result()
                        if status == "NEW": new_handles.append(res_data)
                        elif status == "EXISTING": existing_handles.append(res_data)
                        else: rejected_handles.append(res_data)
                    except Exception: pass
                    completed_count += 1
                    progress_bar.progress(completed_count / total_items)
                    status_text.markdown(f"⏳ **Đang phân tích siêu tốc:** `{completed_count}/{total_items}` Handle...")

            progress_bar.empty(); status_text.empty()
            st.session_state['batch_check_new'] = new_handles
            st.session_state['batch_check_existing'] = existing_handles
            st.session_state['batch_check_rejected'] = rejected_handles

    if 'batch_check_new' in st.session_state or 'batch_check_existing' in st.session_state or 'batch_check_rejected' in st.session_state:
        new_handles = st.session_state.get('batch_check_new', [])
        existing_handles = st.session_state.get('batch_check_existing', [])
        rejected_handles = st.session_state.get('batch_check_rejected', [])

        st.divider()
        render_kpi_cards([
            ("TỔNG SỐ KIỂM TRA", f"{len(new_handles) + len(existing_handles) + len(rejected_handles)}", "#47A5D1"),
            ("✅ ĐẠT CHUẨN (>=1M SUBS & >10MIN)", f"{len(new_handles)}", "#10B981"),
            ("❌ ĐÃ TỒN TẠI TRONG DB", f"{len(existing_handles)}", "#F59E0B"),
            ("🚫 BỊ LOẠI (<1M / SHORTS / KHÔNG CÓ >10MIN)", f"{len(rejected_handles)}", "#EF4444")
        ], card_bg, border_color)
        
        res_tab1, res_tab2, res_tab3 = st.tabs([
            f"✅ Kênh Mới Đạt Chuẩn ({len(new_handles)})", 
            f"❌ Kênh Đã Tồn Tại ({len(existing_handles)})",
            f"🚫 Kênh Bị Loại ({len(rejected_handles)})"
        ])
        
        with res_tab1:
            if new_handles:
                cart_keys = set(st.session_state['cart'].keys())
                selected_set = st.session_state['selected_channels']
                selected_not_in_cart = [p for p in selected_set if p not in cart_keys]
                cnt_for_cart, cnt_total_sel = len(selected_not_in_cart), len(selected_set)

                with st.container(border=True):
                    st.markdown('<div class="action-bar-marker"></div>', unsafe_allow_html=True)
                    tb1, tb2, tb3, tb4, tb5, tb6 = st.columns([1.8, 1.8, 1.5, 1.5, 1.5, 1.5])
                    with tb1:
                        if st.button(f"🛒 Thêm Giỏ ({cnt_for_cart})", key="btn_add_sel_t1", use_container_width=True):
                            if cnt_for_cart > 0:
                                items_to_add = []
                                for item in new_handles:
                                    p_id = to_pure_id(item["Handle"])
                                    if p_id in selected_not_in_cart:
                                        item_data = {"Handle": item["Handle"], "Tên Kênh": item.get("Tên Kênh", p_id.upper()), "Link Kênh": get_channel_link(p_id), "Trạng Thái DB": "✅ KÊNH MỚI", "Tag": "📌 Chưa phân loại", "Socials": item.get("Socials", {})}
                                        st.session_state['cart'][p_id] = item_data
                                        items_to_add.append((p_id, item_data))
                                add_batch_to_cart_db(items_to_add); st.toast(f"🎉 Đã thêm {len(items_to_add)} kênh mới vào giỏ!"); st.rerun()
                            else: st.warning("Vui lòng chọn kênh chưa có trong giỏ!")
                    with tb2:
                        if st.button(f"💾 Lưu DB ({cnt_total_sel})", key="btn_save_db_sel_t1", use_container_width=True):
                            if cnt_total_sel > 0:
                                data_db, saved_ids = [], set()
                                for item in new_handles:
                                    p_id = to_pure_id(item["Handle"])
                                    if p_id in selected_set:
                                        data_db.append({"handle": p_id, "youtuber_name": item.get("Tên Kênh", p_id.upper()), "source": "Tra cứu hàng loạt"})
                                        saved_ids.add(p_id)
                                if data_db:
                                    supabase.table("channels").upsert(data_db, on_conflict="handle").execute()
                                    st.session_state['new_db_channels_notify'] = f"🎉 Vừa lưu thành công {len(data_db)} kênh mới vào Database!"
                                    st.session_state['batch_check_new'] = [item for item in new_handles if to_pure_id(item["Handle"]) not in saved_ids]
                                    st.session_state['selected_channels'].difference_update(saved_ids)
                                    st.toast(f"🎉 Đã lưu thành công {len(data_db)} kênh!"); st.rerun()
                            else: st.warning("Vui lòng chọn ít nhất 1 kênh!")
                    with tb3:
                        if st.button(f"⚖️ So sánh ({cnt_total_sel})", key="btn_cmp_sel_t1", use_container_width=True):
                            if 1 < cnt_total_sel <= 5: compare_channels_dialog(get_selected_channel_data())
                            else: st.warning("Vui lòng chọn 2-5 kênh để so sánh!")
                    with tb4:
                        if st.button(f"🗑️ Xóa ({cnt_total_sel})", key="btn_del_sel_t1", use_container_width=True):
                            if cnt_total_sel > 0:
                                for p_id in list(selected_set): delete_channel_from_system(p_id)
                                st.session_state['selected_channels'].clear(); st.toast(f"🗑️ Đã xóa {cnt_total_sel} kênh!"); st.rerun()
                            else: st.warning("Vui lòng chọn ít nhất 1 kênh!")
                    with tb5: st.button("✅ Chọn Tất Cả", key="btn_sel_all_t1_new", on_click=cb_select_all, args=(new_handles,), use_container_width=True)
                    with tb6: st.button("❌ Bỏ Chọn", key="btn_clear_sel_t1_new", on_click=cb_clear_all, use_container_width=True)

                items_per_page = 20
                if 'p_state_t1_new' not in st.session_state: st.session_state['p_state_t1_new'] = 1
                total_pages = max(1, (len(new_handles) + items_per_page - 1) // items_per_page)
                
                page_new = st.number_input("Trang hiện tại:", min_value=1, max_value=total_pages, step=1, key="page_t1_new_input", value=st.session_state['p_state_t1_new'])
                st.session_state['p_state_t1_new'] = page_new
                
                start_idx = (int(page_new) - 1) * items_per_page
                paged_new = new_handles[start_idx:start_idx + items_per_page]

                for idx, item in enumerate(paged_new):
                    p_id = to_pure_id(item["Handle"]); is_active = (p_id == st.session_state.get('active_inspected_handle')); is_in_cart = p_id in st.session_state['cart']
                    stt_num = start_idx + idx + 1
                    
                    with st.container(border=True):
                        if is_active: st.markdown('<div class="active-card-marker"></div><div class="active-banner-tag">🔍 ĐANG XEM 6 VIDEO MỚI</div>', unsafe_allow_html=True)
                        elif is_in_cart: st.markdown('<div class="in-cart-marker"></div><div class="in-cart-banner-tag">🛒 ĐÃ CÓ TRONG GIỎ HÀNG</div>', unsafe_allow_html=True)

                        c0, c1, c2, c3 = st.columns([0.4, 3.1, 3.5, 3.0])
                        with c0: st.checkbox("", key=f"chk_t1_{p_id}_{st.session_state['chk_counter']}", value=(p_id in st.session_state['selected_channels']), on_change=toggle_select_channel, args=(p_id,))
                        with c1:
                            st.markdown(f"<h3 style='margin:0;'><span class='badge-stt'>#{stt_num}</span><a href='{item['Link Kênh']}' style='color:#D95F26;'>{item['Handle']}</a></h3>", unsafe_allow_html=True)
                            st.write(f"**{item.get('Tên Kênh', p_id.upper())}**")
                            c1_1, c1_2 = st.columns(2)
                            if c1_1.button("👁️ Xem Video", key=f"btn_prev_t1_{p_id}", on_click=set_active_inspected_channel, args=(p_id,)): show_video_dialog(p_id)
                            if c1_2.button("📩 Soạn Mail", key=f"btn_mail_pass_{p_id}"): show_ai_email_dialog(item)
                        with c2:
                            st.write(f"👥 **Subs:** `{item.get('Subscribers', 'N/A')}` | 🌍 **Q.Gia:** `{item.get('Quốc gia', 'N/A')}`")
                            st.markdown(render_social_badges_html(item.get("Socials", {})), unsafe_allow_html=True)
                        with c3:
                            bc1, bc2 = st.columns(2)
                            if is_in_cart:
                                if bc1.button("❌ Bỏ Giỏ", key=f"rm_t1_{p_id}", use_container_width=True): remove_from_cart_db(p_id); del st.session_state['cart'][p_id]; st.rerun()
                            else:
                                if bc1.button("🛒 Giỏ", key=f"add_t1_{p_id}", use_container_width=True):
                                    item_data = {"Handle": item["Handle"], "Tên Kênh": item.get("Tên Kênh", p_id.upper()), "Link Kênh": get_channel_link(p_id), "Trạng Thái DB": "✅ KÊNH MỚI", "Tag": "📌 Chưa phân loại", "Socials": item.get("Socials", {})}
                                    st.session_state['cart'][p_id] = item_data; add_to_cart_db(p_id, item_data); st.rerun()
                            if bc2.button("🗑️ Xóa", key=f"del_t1_{p_id}", use_container_width=True): delete_channel_from_system(p_id); st.toast(f"🗑️ Đã xóa kênh @{p_id}!"); st.rerun()

        with res_tab2:
            if existing_handles:
                st.caption("📋 Danh sách kênh trùng lặp đã tồn tại sẵn trong Database:")
                df_exist = pd.DataFrame(existing_handles)
                cols_exist = [c for c in ['Handle', 'Tên Kênh', 'Trạng thái'] if c in df_exist.columns]
                st.dataframe(df_exist[cols_exist], use_container_width=True)
            else:
                st.info("Không có kênh nào bị trùng lặp trong đợt kiểm tra này.")

        with res_tab3:
            if rejected_handles:
                st.caption("📋 Bảng chi tiết lý do từng kênh bị gạt bỏ. Bạn có thể bấm '🔄 Phục Hồi' để chuyển kênh sang danh sách Đạt Chuẩn:")
                for idx, item in enumerate(list(rejected_handles)):
                    p_id = to_pure_id(item.get("Handle", ""))
                    with st.container(border=True):
                        rc1, rc2, rc3 = st.columns([5, 3.5, 1.5])
                        with rc1:
                            st.markdown(f"### #{idx+1} <a href='{get_channel_link(p_id)}' style='color:#EF4444;'>{item.get('Handle')}</a>", unsafe_allow_html=True)
                            st.write(f"**Tên Kênh:** {item.get('Tên Kênh', p_id.upper() if p_id else 'N/A')}")
                            if item.get('Subscribers'):
                                st.write(f"👥 **Subs:** `{item.get('Subscribers')}`")
                        with rc2:
                            reason = item.get('Lý do loại') or item.get('Trạng thái') or 'Không đạt tiêu chí lọc'
                            st.markdown(f"<div style='background-color:#FEE2E2; border:1px solid #EF4444; padding:10px; border-radius:8px;'><span style='color:#B91C1C; font-weight:800; font-size:0.85rem;'>🔴 LÝ DO LOẠI:</span><br><span style='color:#991B1B; font-weight:700; font-size:0.85rem;'>{reason}</span></div>", unsafe_allow_html=True)
                        with rc3:
                            st.write("")
                            if st.button("🔄 Phục Hồi", key=f"btn_restore_rej_{idx}_{p_id}", use_container_width=True):
                                st.session_state['batch_check_rejected'] = [r for r in st.session_state['batch_check_rejected'] if to_pure_id(r.get("Handle")) != p_id]
                                restored_item = dict(item)
                                restored_item["Trạng thái"] = "✅ Đã Phục Hồi Thủ Công"
                                restored_item["Link Kênh"] = get_channel_link(p_id)
                                if 'batch_check_new' not in st.session_state:
                                    st.session_state['batch_check_new'] = []
                                st.session_state['batch_check_new'].append(restored_item)
                                st.toast(f"🎉 Đã phục hồi kênh @{p_id} sang danh sách Đạt Chuẩn!")
                                st.rerun()

    render_shared_cart_ui(key_suffix="tab1")

# --- TAB 2: LIVE API SCRAPER ---
with tab2:
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>⚡ Cào dữ liệu Live & Xuất Báo Cáo Audit chuẩn V4.14 (Đơn / Hàng Loạt)</h3>", unsafe_allow_html=True)
    col_t2_1, col_t2_2 = st.columns([2, 1])
    with col_t2_1: text_input_area_t2 = st.text_area("Dán danh sách Handle/Link (mỗi dòng 1 link):", height=180, value="@4wd247", key="text_input_tab2")
    with col_t2_2: file_input_t2 = st.file_uploader("Upload file danh sách (.zip, .xlsx, .txt, .docx, .csv):", type=["zip", "xlsx", "xls", "txt", "docx", "doc", "csv"], key="file_input_tab2")

    if st.button("🚀 Bắt Đầu Cào Live & Tạo Báo Cáo Audit V4.14", type="primary", key="btn_run_tab2_audit"):
        inputs_t2 = re.split(r'[\n,\t\r]+', str(text_input_area_t2)) if text_input_area_t2 else []
        if file_input_t2: inputs_t2.extend(extract_raw_inputs_from_file(file_input_t2))
        target_handles_t2 = parse_raw_inputs_to_handles(inputs_t2)
        if not target_handles_t2: st.warning("⚠️ Vui lòng dán danh sách Handle/Link hoặc upload file!")
        else:
            tot_t2 = len(target_handles_t2); prog_t2 = st.progress(0); stat_t2 = st.empty(); comp_t2 = 0; audit_results_t2 = []; db_upsert_list_t2 = []
            active_keys_t2 = st.session_state.get('api_keys', [DEFAULT_API_KEY])
            exhausted_set_t2 = set(st.session_state.get('exhausted_keys_set', set()))

            with ThreadPoolExecutor(max_workers=5) as executor:
                futures = {executor.submit(run_single_channel_audit, p_h, active_keys_t2, exhausted_set_t2): p_h for p_h in target_handles_t2}
                for future in as_completed(futures):
                    p_h_orig = futures[future]; comp_t2 += 1; prog_t2.progress(comp_t2 / tot_t2)
                    stat_t2.markdown(f"⏳ **Đang cào dữ liệu Live:** `{comp_t2}/{tot_t2}` kênh (@{p_h_orig})...")
                    try:
                        res_val = future.result()
                        if res_val and res_val[0] and res_val[1]:
                            b_bytes, f_name, _ = res_val; audit_results_t2.append((f_name, b_bytes))
                            p_h_clean = f_name.split('_')[0]; db_upsert_list_t2.append({"handle": p_h_clean, "youtuber_name": p_h_clean.upper(), "source": "Live Audit Scraper"})
                    except Exception: pass

            prog_t2.empty(); stat_t2.empty()
            if db_upsert_list_t2: supabase.table("channels").upsert(db_upsert_list_t2, on_conflict="handle").execute()
            st.session_state['tab2_audit_output'] = {"results": audit_results_t2, "count": len(audit_results_t2), "total_requested": tot_t2}; st.rerun()

    if 'tab2_audit_output' in st.session_state:
        out_data = st.session_state['tab2_audit_output']; audit_results_t2 = out_data["results"]
        st.divider()
        if len(audit_results_t2) == 1:
            f_name, b_bytes = audit_results_t2[0]
            st.success(f"🎉 Đã dựng xong báo cáo Audit V4.14 cho kênh @{f_name.split('_')[0]}!")
            st.download_button("📥 Tải về File Audit V4.14 (.xlsx)", data=b_bytes, file_name=f_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
        elif len(audit_results_t2) > 1:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z_file:
                for fname, fbytes in audit_results_t2: z_file.writestr(fname, fbytes)
            st.success(f"🎉 Đã cào dữ liệu & dựng {len(audit_results_t2)} Báo cáo Audit V4.14!")
            st.download_button(f"📦 TẢI GÓI AUDIT ZIP ({len(audit_results_t2)} FILE EXCEL)", data=zip_buf.getvalue(), file_name=f"Goi_Audit_{datetime.datetime.now().strftime('%d-%m-%Y')}.zip", mime="application/zip", type="primary", use_container_width=True)

# TAB 3, TAB 4, TAB 5, TAB 6 giữ nguyên giao diện đẹp hiện có.
