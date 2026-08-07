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

# --- PROCESS PENDING DATA FROM TAB 6 TO TAB 3 (BEFORE WIDGET INSTANTIATION) ---
if 'pending_seed_handle' in st.session_state:
    st.session_state['seed_input_tab3'] = st.session_state.pop('pending_seed_handle')
if 'pending_keywords' in st.session_state:
    st.session_state['custom_kw_tab3'] = st.session_state.pop('pending_keywords')

TAG_OPTIONS = ["📌 Chưa phân loại", "🔥 Ưu tiên làm", "📩 Đã liên hệ", "⏳ Đang chờ duyệt", "✅ Đã chốt", "❌ Bỏ qua"]

def get_tag_index(tag):
    return TAG_OPTIONS.index(tag) if tag in TAG_OPTIONS else 0

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

# Apply Theme Styling
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
        if pct < 70:
            color = "#10B981"
            status_label = f"🟢 {used:,}/10,000"
        elif pct < 90:
            color = "#F59E0B"
            status_label = f"🟡 {used:,}/10,000"
        else:
            color = "#EF4444"
            status_label = f"🔴 {used:,}/10,000" if k_stat_type != "DEAD" else "💀 Key Chết"

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
        keys_to_clear = [
            'batch_check_new', 'batch_check_existing', 'batch_check_rejected',
            'passed_channels', 'rejected_channels', 'tab2_audit_output',
            'active_inspected_handle', 'last_inspected_data',
            'last_inspected_handle', 'seed_input_tab3', 'custom_kw_tab3',
            'pending_seed_handle', 'pending_keywords'
        ]
        for k in keys_to_clear:
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
            response = supabase.table("channels").select("handle, youtuber_name").in_("handle", target_list).execute()
            db_matches = {item["handle"].lower(): item for item in response.data} if response.data else {}
            
            new_handles, existing_handles, rejected_handles = [], [], []
            total_items, completed_count = len(target_list), 0
            
            active_keys_t1 = st.session_state.get('api_keys', [DEFAULT_API_KEY])
            exhausted_set_t1 = set(st.session_state.get('exhausted_keys_set', set()))

            with ThreadPoolExecutor(max_workers=8) as executor:
                futures = [executor.submit(process_tab1_single_handle, p_id, db_matches, active_keys_t1, exhausted_set_t1) for p_id in target_list]
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
                page_new = st.session_state['p_state_t1_new']
                start_idx = (page_new - 1) * items_per_page
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
                            if bc2.button("🗑️ Xóa", key=f"del_t1_{p_id}", use_container_width=True): delete_channel_from_system(p_id); st.rerun()

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
                            
                            if st.button("🛒 Phục Hồi & Vào Giỏ", key=f"btn_restore_cart_{idx}_{p_id}", use_container_width=True):
                                st.session_state['batch_check_rejected'] = [r for r in st.session_state['batch_check_rejected'] if to_pure_id(r.get("Handle")) != p_id]
                                restored_item = dict(item)
                                restored_item["Trạng thái"] = "✅ Đã Phục Hồi Thủ Công"
                                restored_item["Link Kênh"] = get_channel_link(p_id)
                                if 'batch_check_new' not in st.session_state:
                                    st.session_state['batch_check_new'] = []
                                st.session_state['batch_check_new'].append(restored_item)
                                
                                item_data = {
                                    "Handle": item.get("Handle", f"@{p_id}"),
                                    "Tên Kênh": item.get("Tên Kênh", p_id.upper()),
                                    "Link Kênh": get_channel_link(p_id),
                                    "Trạng Thái DB": "✅ ĐÃ PHỤC HỒI",
                                    "Tag": "📌 Chưa phân loại",
                                    "Socials": item.get("Socials", {})
                                }
                                st.session_state['cart'][p_id] = item_data
                                add_to_cart_db(p_id, item_data)
                                st.toast(f"🎉 Đã phục hồi & thêm kênh @{p_id} vào Giỏ hàng vĩnh viễn!")
                                st.rerun()
            else:
                st.info("Không có kênh nào bị loại.")

    render_shared_cart_ui(key_suffix="tab1")

# --- TAB 2: LIVE API SCRAPER ---
with tab2:
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>⚡ Cào dữ liệu Live & Xuất Báo Cáo Audit chuẩn V4.14 (Đơn / Hàng Loạt)</h3>", unsafe_allow_html=True)
    st.caption("💡 *Hỗ trợ dán danh sách nhiều Handle/Link YouTube/Link Tìm kiếm, hoặc Upload file `.zip`, `.xlsx`, `.txt`, `.docx` (Word), `.csv`.*")

    col_t2_1, col_t2_2 = st.columns([2, 1])
    with col_t2_1: text_input_area_t2 = st.text_area("Dán danh sách Handle/Link kênh/Link Video/Link Tìm kiếm (mỗi dòng 1 link):", height=180, value="@4wd247", key="text_input_tab2")
    with col_t2_2: file_input_t2 = st.file_uploader("Upload file danh sách hoặc gói báo cáo (.zip, .xlsx, .txt, .docx, .csv):", type=["zip", "xlsx", "xls", "txt", "docx", "doc", "csv"], key="file_input_tab2")

    if st.button("🚀 Bắt Đầu Cào Live & Tạo Báo Cáo Audit V4.14", type="primary", key="btn_run_tab2_audit"):
        inputs_t2 = re.split(r'[\n,\t\r]+', str(text_input_area_t2)) if text_input_area_t2 else []
        if file_input_t2: inputs_t2.extend(extract_raw_inputs_from_file(file_input_t2))
        target_handles_t2 = parse_raw_inputs_to_handles(inputs_t2)
        if not target_handles_t2: st.warning("⚠️ Vui lòng dán danh sách Handle/Link hoặc upload file để cào báo cáo!")
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
        out_data = st.session_state['tab2_audit_output']; audit_results_t2 = out_data["results"]; tot_t2 = out_data["total_requested"]
        st.divider()
        if len(audit_results_t2) == 1:
            f_name, b_bytes = audit_results_t2[0]
            st.success(f"🎉 Đã dựng xong báo cáo Audit V4.14 cho kênh @{f_name.split('_')[0]}!")
            st.download_button("📥 Tải về File Audit V4.14 (.xlsx)", data=b_bytes, file_name=f_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
        elif len(audit_results_t2) > 1:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z_file:
                for fname, fbytes in audit_results_t2: z_file.writestr(fname, fbytes)
            st.success(f"🎉 Đã cào dữ liệu & dựng {len(audit_results_t2)} / {tot_t2} Báo cáo Audit V4.14!")
            st.download_button(f"📦 TẢI GÓI AUDIT ZIP ({len(audit_results_t2)} FILE EXCEL)", data=zip_buf.getvalue(), file_name=f"Goi_Audit_{datetime.datetime.now().strftime('%d-%m-%Y')}.zip", mime="application/zip", type="primary", use_container_width=True)

# --- TAB 3: SMART RELATED FINDER ---
with tab3:
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>🎯 Săn Kênh Tương Tự & Giỏ Hàng (Multi-threaded Speed)</h3>", unsafe_allow_html=True)
    if 'audit_success_msg' in st.session_state: st.success(st.session_state['audit_success_msg']); del st.session_state['audit_success_msg']

    trigger_auto_start_search = False
    if st.session_state.get('trigger_deep_search_now', False):
        st.session_state['trigger_deep_search_now'] = False; trigger_auto_start_search = True

    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        seed_channel_input = st.text_input("Nhập Handle Kênh Mồi (ví dụ: @NickDiGiovanni):", key="seed_input_tab3")
        if st.button("✨ Tự Động Phân Tích từ Kênh Mồi"):
            pure_s_auto = to_pure_id(seed_channel_input)
            if pure_s_auto:
                try:
                    active_keys = st.session_state.get('api_keys', [DEFAULT_API_KEY])
                    cid_auto, _, _, _ = get_channel_id_by_handle_direct(pure_s_auto, active_keys)
                    if cid_auto:
                        ext = extract_channel_master_keywords(cid_auto)
                        st.session_state['custom_kw_tab3'] = ", ".join(ext['master_keywords'][:6]); st.rerun()
                except Exception as e: st.error(f"Lỗi: {e}")
        
        custom_keywords_input = st.text_input("Từ khóa chủ đề (Tự động liên kết):", key="custom_kw_tab3")
    with col_f2:
        min_subs_choice = st.selectbox("Mốc Subscribers Tối Thiểu:", options=[100000, 250000, 500000, 1000000], index=3, format_func=lambda x: f"{x:,} Subs")
        min_duration_choice = st.selectbox("Lọc Yêu Cầu Đồ Dài Video:", options=[600], index=0, format_func=lambda x: "Bắt buộc có Video > 10 phút")

    start_btn = st.button("🚀 Bắt Đầu Săn Kênh Đồng Ngách", type="primary")

    if (start_btn or trigger_auto_start_search) and seed_channel_input:
        pure_seed = to_pure_id(seed_channel_input)
        try:
            st.info(f"🔍 Đang kết nối API và phân tích `{pure_seed}`...")
            active_keys_t3 = st.session_state.get('api_keys', [DEFAULT_API_KEY])
            exhausted_set_t3 = set(st.session_state.get('exhausted_keys_set', set()))
            seed_id, _, _, _ = get_channel_id_by_handle_direct(pure_seed, active_keys_t3, exhausted_set_t3)

            if not seed_id: st.error("Không tìm thấy kênh mồi này!")
            else:
                top_kw_list = clean_and_extract_keywords(custom_keywords_input, seed_handle=pure_seed) if custom_keywords_input else extract_channel_master_keywords(seed_id)['master_keywords'][:4]
                st.write(f"🏷️ **Từ khóa quét:** `{', '.join(top_kw_list)}`")
                candidate_channel_ids = set()
                q_chan = " ".join(top_kw_list[:2])
                c_search_res, _, _, _ = yt_execute_safe(lambda yt: yt.search().list(part="snippet", q=q_chan, type="channel", maxResults=50), active_keys_t3, exhausted_set_t3, cost=100)
                for c_item in c_search_res.get('items', []):
                    if c_item['snippet']['channelId'] != seed_id: candidate_channel_ids.add(c_item['snippet']['channelId'])

                candidate_ids_list = list(candidate_channel_ids)
                if not candidate_ids_list: st.warning("Không quét được ứng viên nào!")
                else:
                    passed_channels, rejected_channels, channel_items, candidate_handles = [], [], [], []
                    for i in range(0, len(candidate_ids_list), 50):
                        chan_res, _, _, _ = yt_execute_safe(lambda yt: yt.channels().list(part="snippet,contentDetails,statistics", id=','.join(candidate_ids_list[i:i+50])), active_keys_t3, exhausted_set_t3, cost=1)
                        for item in chan_res.get('items', []):
                            c_h = to_pure_id(item['snippet'].get('customUrl', '')) or item['id']
                            candidate_handles.append(c_h.lower()); channel_items.append(item)

                    db_res = supabase.table("channels").select("handle").in_("handle", candidate_handles).execute()
                    db_existing_set = {r["handle"].lower() for r in db_res.data} if db_res.data else set()

                    prog_t3 = st.progress(0); stat_t3 = st.empty(); tot_cand = len(channel_items); comp_cand = 0
                    with ThreadPoolExecutor(max_workers=8) as executor:
                        futures = [executor.submit(process_single_candidate, item, min_subs_choice, min_duration_choice, db_existing_set, active_keys_t3, exhausted_set_t3) for item in channel_items]
                        for future in as_completed(futures):
                            try:
                                is_pass, res_data = future.result()
                                if is_pass: passed_channels.append(res_data)
                                else: rejected_channels.append(res_data)
                            except Exception: pass
                            comp_cand += 1; prog_t3.progress(comp_cand / tot_cand)
                            stat_t3.markdown(f"📊 **Đang phân tích siêu tốc:** `{comp_cand}/{tot_cand}` ứng viên...")

                    prog_t3.empty(); stat_t3.empty()
                    st.session_state['passed_channels'] = passed_channels
                    st.session_state['rejected_channels'] = rejected_channels
        except Exception as e: st.error(f"Lỗi: {e}")

    if 'passed_channels' in st.session_state or 'rejected_channels' in st.session_state:
        passed_list = st.session_state.get('passed_channels', [])
        rejected_list = st.session_state.get('rejected_channels', [])
        st.divider()
        render_kpi_cards([
            ("TỔNG ỨNG VIÊN", f"{len(passed_list) + len(rejected_list)}", "#47A5D1"),
            (f"✅ ĐẠT CHUẨN (>{min_subs_choice:,} SUBS & >10MIN)", f"{len(passed_list)}", "#10B981"),
            ("❌ BỊ LOẠI", f"{len(rejected_list)}", "#EF4444")
        ], card_bg, border_color)

    render_shared_cart_ui(key_suffix="tab3")

# --- TAB 4: UPLOAD DATA ---
with tab4:
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>📤 Upload file .ZIP, .TXT, .XLSX hoặc .DOCX để cập nhật Database</h3>", unsafe_allow_html=True)
    st.caption("💡 *Hỗ trợ tải lên trực tiếp các file Excel báo cáo lẻ (.xlsx), file Word (.docx), file nén .ZIP hoặc file danh sách .TXT.*")
    
    uploaded_files = st.file_uploader("Kéo thả file vào đây:", type=["zip", "txt", "xlsx", "xls", "docx", "doc", "csv"], accept_multiple_files=True)
    
    if uploaded_files and st.button("🚀 Bắt đầu xử lý & Nạp vào Database", type="primary"):
        new_handles_to_insert, skipped_details = [], []
        for file in uploaded_files:
            file_name = file.name.lower()
            if file_name.endswith('.zip'):
                with zipfile.ZipFile(file, 'r') as zip_ref:
                    extract_path = "temp_zip_extract"; zip_ref.extractall(extract_path)
                    for root, _, filenames in os.walk(extract_path):
                        for fn in filenames:
                            if fn.startswith('~$') or fn.startswith('._'): continue
                            if fn.endswith('.xlsx') or fn.endswith('.xls'):
                                h = extract_handle_from_filename(fn)
                                if h: new_handles_to_insert.append({"handle": h, "youtuber_name": h.upper(), "source": file.name, "filename": fn})
            elif file_name.endswith('.txt'):
                content = file.read().decode("utf-8", errors="ignore")
                for line in content.splitlines():
                    h = to_pure_id(line)
                    if h: new_handles_to_insert.append({"handle": h, "youtuber_name": h.upper(), "source": file.name, "filename": file.name})
            elif file_name.endswith('.xlsx') or file_name.endswith('.xls'):
                h_from_fn = extract_handle_from_filename(file.name)
                file.seek(0)
                try:
                    df_excel = pd.read_excel(file)
                    target_cols = [col for col in df_excel.columns if any(k in str(col).lower() for k in ['search', 'youtuber', 'handle', 'link', 'kênh'])]
                    cols_to_use = target_cols if target_cols else df_excel.columns
                    for col in cols_to_use:
                        if 'stats' in str(col).lower() or 'kz.youtubers' in str(col).lower(): continue
                        for val in df_excel[col].dropna():
                            if is_garbage_input(val): continue
                            p = to_pure_id(val)
                            if p: new_handles_to_insert.append({"handle": p, "youtuber_name": p.upper(), "source": file.name, "filename": file.name})
                except Exception: pass

        if new_handles_to_insert:
            df_raw = pd.DataFrame(new_handles_to_insert)
            df_insert = df_raw.drop_duplicates(subset=["handle"])
            supabase.table("channels").upsert(df_insert[["handle", "youtuber_name", "source"]].to_dict(orient="records"), on_conflict="handle").execute()
            for k in list(st.session_state.keys()):
                if k.startswith('crm_cache_') or k == 'tab5_crm_cache': st.session_state.pop(k, None)
            st.session_state['new_db_channels_notify'] = f"🎉 Vừa tải lên & nạp thành công {len(df_insert)} kênh mới vào Database!"
            st.success(f"🎉 Đã đồng bộ thành công {len(df_insert)} Handle vào Supabase!")
        else: st.warning("⚠️ Không tìm thấy Handle hợp lệ nào!")

# --- TAB 5: CRM DATABASE VIEWER (FULL DETAILS & 2 MODES) ---
with tab5:
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>📊 Quản lý Database CRM Kênh</h3>", unsafe_allow_html=True)
    if 'new_db_channels_notify' in st.session_state: st.success(st.session_state['new_db_channels_notify'])

    try: res = supabase.table("channels").select("*").order("created_at", desc=True).execute()
    except Exception:
        try: res = supabase.table("channels").select("*").order("id", desc=True).execute()
        except Exception: res = supabase.table("channels").select("*").execute()

    if res.data:
        df_all = pd.DataFrame(res.data)
        if 'created_at' in df_all.columns:
            df_all['created_at_dt'] = pd.to_datetime(df_all['created_at'], errors='coerce')
            df_all = df_all.sort_values(by='created_at_dt', ascending=False)

        c_top1, c_top2 = st.columns([7, 3])
        with c_top1: st.markdown(f"Tổng số kênh hiện có trong DB: <span style='font-weight:800; color:#D95F26;'>{len(df_all)}</span>", unsafe_allow_html=True)
        with c_top2:
            if st.button("💣 Xóa Vĩnh Viễn Toàn Bộ DB", use_container_width=True, key="btn_trigger_wipe_db"):
                confirm_clear_db_dialog(cb_clear_all)

        st.divider()
        st.markdown("#### 🔍 Bộ Lọc Database Chuyên Sâu:")
        fc1, fc2, fc3, fc4 = st.columns([2, 2, 2, 2])
        with fc1: search_db = st.text_input("Tìm kiếm thủ công (Handle, Tên, Link Kênh, Link Video):", placeholder="Dán @handle, link Youtube hoặc link Video...")
        with fc2: sel_sub_range = st.selectbox("Lọc theo Mốc Subscribers:", options=["-- Tất Cả Mốc Subs --", "< 100K Subs", "100K - 500K Subs", "500K - 1M Subs", "> 1M Subs"])
        with fc3: sel_source = st.selectbox("Lọc theo Nguồn Nạp (File ZIP / Source):", options=["-- Tất Cả Nguồn --"] + (list(df_all['source'].dropna().unique()) if 'source' in df_all.columns else []))
        with fc4: view_mode = st.radio("Chế độ hiển thị:", ["🎨 Card View (Thẻ chi tiết)", "📊 Table Grid View (Bảng nén gọn)"], horizontal=True)

        df_pre = df_all.copy()
        if search_db.strip():
            s_input = search_db.strip(); v_id = extract_video_id(s_input); resolved_handles = []
            if v_id:
                with st.spinner("🔍 Đang giải mã Link Video sang Handle..."): resolved_handles = get_handles_from_video_ids([v_id])
            pure_h = to_pure_id(s_input); clean_kw = re.sub(r'^@+', '', s_input).strip().lower()
            cond_handle = df_pre['handle'].str.contains(clean_kw, case=False, na=False)
            cond_name = df_pre['youtuber_name'].str.contains(clean_kw, case=False, na=False)
            if pure_h: cond_handle = cond_handle | df_pre['handle'].str.contains(pure_h, case=False, na=False)
            if resolved_handles: df_pre = df_pre[cond_handle | cond_name | df_pre['handle'].apply(to_pure_id).isin([h.lower() for h in resolved_handles])]
            else: df_pre = df_pre[cond_handle | cond_name]

        if sel_source != "-- Tất Cả Nguồn --": df_pre = df_pre[df_pre['source'] == sel_source]

        crm_meta_map = {}
        if sel_sub_range != "-- Tất Cả Mốc Subs --":
            handles_to_check = [to_pure_id(h) for h in df_pre['handle'].tolist() if to_pure_id(h)]
            matched_handles = []
            if handles_to_check:
                prog_db = st.progress(0); stat_db = st.empty(); comp_h = 0
                active_keys_t5_filter = st.session_state.get('api_keys', [DEFAULT_API_KEY])
                exhausted_set_t5_filter = set(st.session_state.get('exhausted_keys_set', set()))

                with ThreadPoolExecutor(max_workers=10) as executor:
                    futures = [executor.submit(process_single_crm_channel_meta, p_h, active_keys_t5_filter, exhausted_set_t5_filter) for p_h in handles_to_check]
                    for future in as_completed(futures):
                        try:
                            p_h, meta, logs = future.result(); crm_meta_map[p_h] = meta; s_num = meta['sub_count']
                            is_match = False
                            if sel_sub_range == "< 100K Subs" and (0 < s_num < 100000): is_match = True
                            elif sel_sub_range == "100K - 500K Subs" and (100000 <= s_num < 500000): is_match = True
                            elif sel_sub_range == "500K - 1M Subs" and (500000 <= s_num < 1000000): is_match = True
                            elif sel_sub_range == "> 1M Subs" and s_num >= 1000000: is_match = True
                            if is_match: matched_handles.append(p_h)
                        except Exception: pass
                        comp_h += 1; prog_db.progress(comp_h / len(handles_to_check))
                prog_db.empty(); stat_db.empty()
                df_filtered = df_pre[df_pre['handle'].apply(to_pure_id).isin(matched_handles)]
            else: df_filtered = df_pre
        else: df_filtered = df_pre

        st.caption(f"🎯 Kết quả khớp: **{len(df_filtered)}** / {len(df_all)} kênh")

        if view_mode == "📊 Table Grid View (Bảng nén gọn)":
            df_grid = df_filtered.copy()
            df_grid['Link Kênh'] = df_grid['handle'].apply(lambda h: get_channel_link(to_pure_id(h)))
            df_grid['Tab Videos'] = df_grid['handle'].apply(lambda h: f"{get_channel_link(to_pure_id(h))}/videos" if to_pure_id(h) else "")
            if 'created_at' in df_grid.columns:
                df_grid['Ngày Cập Nhật'] = pd.to_datetime(df_grid['created_at'], errors='coerce').dt.strftime('%d-%m-%Y %H:%M')
            cols_show = [c for c in ['handle', 'youtuber_name', 'source', 'Ngày Cập Nhật', 'Link Kênh', 'Tab Videos'] if c in df_grid.columns]
            df_grid.index = range(1, len(df_grid) + 1)
            st.dataframe(df_grid[cols_show], use_container_width=True, column_config={
                "Link Kênh": st.column_config.LinkColumn("Trang Chủ", display_text="🏠 Kênh"),
                "Tab Videos": st.column_config.LinkColumn("Tab Videos", display_text="🎬 Videos")
            })
        else:
            items_per_page = 20
            if 'p_state_t5_db' not in st.session_state: st.session_state['p_state_t5_db'] = 1
            total_pages = max(1, (len(df_filtered) + items_per_page - 1) // items_per_page)
            page = st.session_state['p_state_t5_db']
            start_idx = (int(page) - 1) * items_per_page; end_idx = start_idx + items_per_page
            page_data = df_filtered.iloc[start_idx:end_idx]

            col_db_p1, col_db_p2 = st.columns([2, 8])
            with col_db_p1:
                st.number_input("Trang:", min_value=1, max_value=total_pages, step=1, key="page_db_viewer_top", on_change=sync_pagination_top, args=("page_db_viewer_top", "page_db_viewer_bottom", "p_state_t5_db"))
            with col_db_p2:
                st.write(""); st.markdown(f"📄 **Trang {int(page)} / {total_pages}** *(Hiển thị {format_page_range(int(page), items_per_page, len(df_filtered))} kênh)*")

            paged_handles = [to_pure_id(r['handle']) for _, r in page_data.iterrows() if to_pure_id(r['handle'])]
            missing_paged = [p_h for p_h in paged_handles if p_h not in crm_meta_map]
            active_keys_t5 = st.session_state.get('api_keys', [DEFAULT_API_KEY])
            exhausted_set_t5 = set(st.session_state.get('exhausted_keys_set', set()))

            if missing_paged:
                with ThreadPoolExecutor(max_workers=10) as executor:
                    futures = [executor.submit(process_single_crm_channel_meta, p_h, active_keys_t5, exhausted_set_t5) for p_h in missing_paged]
                    for future in as_completed(futures):
                        try:
                            p_h, meta, logs = future.result(); crm_meta_map[p_h] = meta
                        except Exception: pass

            with st.container(border=True):
                st.markdown('<div class="action-bar-marker"></div>', unsafe_allow_html=True)
                cnt_total_sel_db = len(st.session_state['selected_channels'])
                db_act1, db_act2, db_act3, db_act4 = st.columns([3, 2, 2, 1.5])
                with db_act1:
                    if st.button(f"🗑️ Xóa ({cnt_total_sel_db}) Kênh Đã Chọn Khỏi DB", key="btn_del_sel_db", use_container_width=True):
                        if cnt_total_sel_db > 0:
                            for p_id in list(st.session_state['selected_channels']): delete_channel_from_system(p_id)
                            cb_clear_all(); st.toast(f"🗑️ Đã xóa {cnt_total_sel_db} kênh khỏi Database!"); st.rerun()
                        else: st.warning("Vui lòng tick chọn ít nhất 1 kênh trong DB!")
                with db_act2: st.button("✅ Chọn Tất Cả (Trang Này)", key="btn_sel_page_db", on_click=cb_select_all, args=(page_data.to_dict('records'),), use_container_width=True)
                with db_act3: st.button("✅ Chọn Tất Cả (Toàn DB)", key="btn_sel_all_db", on_click=cb_select_all, args=(df_filtered.to_dict('records'),), use_container_width=True)
                with db_act4: st.button("❌ Bỏ Chọn", key="btn_clear_sel_db", on_click=cb_clear_all, use_container_width=True)

            st.divider()
            for idx, row in page_data.iterrows():
                p_id = to_pure_id(row['handle']); is_active = (p_id == st.session_state.get('active_inspected_handle')); stt_num_db = start_idx + idx + 1
                crm_meta = crm_meta_map.get(p_id) or {"sub_count": -1, "sub_str": "N/A", "country": "N/A", "socials": {}}
                created_str = pd.to_datetime(row['created_at']).strftime("%d-%m-%Y %H:%M") if row.get('created_at') else ""

                with st.container(border=True):
                    if is_active: st.markdown('<div class="active-card-marker"></div><div class="active-banner-tag">🔍 ĐANG XEM 6 VIDEO MỚI</div>', unsafe_allow_html=True)
                    c0, c1, c2, c3 = st.columns([0.4, 3.8, 3.8, 2.0])
                    with c0: st.checkbox("", key=f"chk_db_{p_id}_{st.session_state['chk_counter']}", value=(p_id in st.session_state['selected_channels']), on_change=toggle_select_channel, args=(p_id,))
                    with c1:
                        st.markdown(f"### <span class='badge-stt'>#{stt_num_db}</span><a href='{get_channel_link(p_id)}'>@{p_id}</a>", unsafe_allow_html=True)
                        st.write(f"**Tên YouTuber:** {row.get('youtuber_name', 'N/A')}")
                        if st.button("👁️ Xem 6 Video Mới", key=f"btn_prev_db_{idx}_{p_id}", on_click=set_active_inspected_channel, args=(p_id,)): show_video_dialog(p_id)
                    with c2:
                        st.write(f"👥 **Subs:** `{crm_meta['sub_str']}` | 🌍 **Q.Gia:** `{crm_meta['country']}`")
                        st.write(f"📁 **Nguồn:** {row.get('source', 'N/A')}" + (f" | 📅 **Cập nhật:** `{created_str}`" if created_str else ""))
                        st.markdown(render_social_badges_html(crm_meta.get("socials", {})), unsafe_allow_html=True)
                    with c3:
                        if st.button("🗑️ Xóa DB", key=f"del_db_{idx}_{p_id}", use_container_width=True): delete_channel_from_system(p_id); st.toast(f"🗑️ Đã xóa kênh @{p_id}!"); st.rerun()

        st.divider()
        st.download_button("📥 Tải về toàn bộ Database (CSV)", data=df_all.to_csv(index=False).encode('utf-8'), file_name="master_youtube_database.csv", mime="text/csv", type="primary")
    else: st.info("Database hiện đang trống!")

# --- TAB 6: CHANNEL & VIDEO TAGS SEO INSPECTOR ---
with tab6:
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>✨ Soi Từ Khóa Kênh (Channel & Video Tags SEO Inspector)</h3>", unsafe_allow_html=True)
    st.caption("💡 *Nhập Handle thủ công hoặc chọn nhanh các kênh đã có sẵn trong Database CRM.*")

    col_i1, col_i2 = st.columns([1, 1])
    with col_i1: inspect_handle_input = st.text_input("Nhập Handle Kênh / Link cần soi:", value="@NickDiGiovanni", key="input_inspect_manual")
    with col_i2:
        db_channels_options = ["-- Chọn Kênh từ Database --"]
        try:
            res_db = supabase.table("channels").select("handle, youtuber_name").order("handle").execute()
            if res_db.data:
                for row in res_db.data:
                    h = row.get("handle"); name = row.get("youtuber_name") or h
                    if h: db_channels_options.append(f"@{h} - {name}")
        except Exception: pass
        selected_db_channel = st.selectbox("Hoặc chọn kênh từ Database CRM:", options=db_channels_options, key="sel_inspect_db")

    if st.button("🔍 Soi Từ Khóa Ngay", type="primary", key="btn_run_inspect_tab6"):
        target_inspect = inspect_handle_input
        if selected_db_channel != "-- Chọn Kênh từ Database --":
            target_inspect = selected_db_channel.split(" - ")[0].strip()

        pure_inspect = to_pure_id(target_inspect)
        if pure_inspect:
            try:
                active_keys = st.session_state.get('api_keys', [DEFAULT_API_KEY])
                cid_insp, _, _, _ = get_channel_id_by_handle_direct(pure_inspect, active_keys)
                if not cid_insp: st.error(f"Không tìm thấy Channel ID cho kênh @{pure_inspect}!")
                else:
                    ext_data = extract_channel_master_keywords(cid_insp)
                    kw_str = ", ".join(ext_data['master_keywords'])
                    
                    st.session_state['pending_seed_handle'] = f"@{pure_inspect}"
                    st.session_state['pending_keywords'] = kw_str
                    st.session_state['last_inspected_data'] = ext_data
                    st.session_state['last_inspected_handle'] = pure_inspect
                    st.rerun()
            except Exception as e: st.error(f"Lỗi khi soi từ khóa: {e}")
        else: st.warning("⚠️ Vui lòng nhập hoặc chọn một kênh hợp lệ!")

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
            if ext_data['top_tags']:
                for tag in ext_data['top_tags']: st.write(f"• `{tag}`")
            else:
                st.caption("Chưa có hoặc Kênh không khai báo Video Tags.")
        
        st.markdown("#### 🚀 Dãy Từ Khóa Chủ Đạo (Master Keywords String - Dùng để Copy hoặc Nạp tự động):")
        st.code(", ".join(ext_data['master_keywords']), language="text")
