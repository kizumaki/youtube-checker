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

    st.write("")
    keys_input = st.text_area("Cập nhật danh sách Key (1 key/dòng):", value=st.session_state['global_api_keys'], height=80, key="api_keys_text_area")
    if st.button("💾 Lưu Cấu Hình Key", type="primary", use_container_width=True):
        st.session_state['global_api_keys'] = keys_input; set_api_keys(keys_input); save_api_keys_to_db(keys_input); test_all_api_keys(); st.toast("🎉 Đã lưu vĩnh viễn danh sách API Keys!"); st.rerun()

    st.divider()
    if st.button("🔄 Làm Mới Màn Hình", use_container_width=True):
        cb_clear_all(); st.rerun()

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
                    with st.container(border=True):
                        if is_active: st.markdown('<div class="active-banner-tag">🔍 ĐANG XEM 6 VIDEO MỚI</div>', unsafe_allow_html=True)
                        elif is_in_cart: st.markdown('<div class="in-cart-banner-tag">🛒 ĐÃ CÓ TRONG GIỎ HÀNG</div>', unsafe_allow_html=True)

                        c0, c1, c2, c3 = st.columns([0.4, 3.1, 3.5, 3.0])
                        with c0: st.checkbox("", key=f"chk_t1_{p_id}_{st.session_state['chk_counter']}", value=(p_id in st.session_state['selected_channels']), on_change=toggle_select_channel, args=(p_id,))
                        with c1:
                            st.markdown(f"### <a href='{item['Link Kênh']}' style='color:#D95F26;'>{item['Handle']}</a>", unsafe_allow_html=True)
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

    render_shared_cart_ui(key_suffix="tab1")

# --- TAB 2: LIVE AUDIT ---
with tab2:
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>⚡ Cào dữ liệu Live & Xuất Báo Cáo Audit chuẩn V4.14</h3>", unsafe_allow_html=True)
    col_t2_1, col_t2_2 = st.columns([2, 1])
    with col_t2_1: text_input_area_t2 = st.text_area("Dán danh sách Handle/Link:", height=180, value="@4wd247", key="text_input_tab2")
    with col_t2_2: file_input_t2 = st.file_uploader("Upload file danh sách (.zip, .xlsx, .txt, .docx, .csv):", key="file_input_tab2")

    if st.button("🚀 Bắt Đầu Cào Live & Tạo Báo Cáo Audit V4.14", type="primary", key="btn_run_tab2_audit"):
        inputs_t2 = re.split(r'[\n,\t\r]+', str(text_input_area_t2)) if text_input_area_t2 else []
        if file_input_t2: inputs_t2.extend(extract_raw_inputs_from_file(file_input_t2))
        target_handles_t2 = parse_raw_inputs_to_handles(inputs_t2)
        if target_handles_t2:
            tot_t2 = len(target_handles_t2); prog_t2 = st.progress(0); stat_t2 = st.empty(); comp_t2 = 0; audit_results_t2 = []; db_upsert_list_t2 = []
            active_keys_t2 = st.session_state.get('api_keys', [DEFAULT_API_KEY])
            exhausted_set_t2 = set(st.session_state.get('exhausted_keys_set', set()))

            with ThreadPoolExecutor(max_workers=5) as executor:
                futures = {executor.submit(run_single_channel_audit, p_h, active_keys_t2, exhausted_set_t2): p_h for p_h in target_handles_t2}
                for future in as_completed(futures):
                    comp_t2 += 1; prog_t2.progress(comp_t2 / tot_t2)
                    try:
                        res_val = future.result()
                        if res_val and res_val[0] and res_val[1]:
                            b_bytes, f_name, _ = res_val; audit_results_t2.append((f_name, b_bytes))
                            p_h_clean = f_name.split('_')[0]; db_upsert_list_t2.append({"handle": p_h_clean, "youtuber_name": p_h_clean.upper(), "source": "Live Audit Scraper"})
                    except Exception: pass

            prog_t2.empty(); stat_t2.empty()
            if db_upsert_list_t2: supabase.table("channels").upsert(db_upsert_list_t2, on_conflict="handle").execute()
            st.session_state['tab2_audit_output'] = {"results": audit_results_t2, "count": len(audit_results_t2), "total_requested": tot_t2}; st.rerun()

# --- TAB 3: SMART RELATED FINDER ---
with tab3:
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>🎯 Săn Kênh Tương Tự & Giỏ Hàng</h3>", unsafe_allow_html=True)
    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        seed_channel_input = st.text_input("Nhập Handle Kênh Mồi (ví dụ: @NickDiGiovanni):", key="seed_input_tab3")
        custom_keywords_input = st.text_input("Từ khóa chủ đề:", key="custom_kw_tab3")
    with col_f2:
        min_subs_choice = st.selectbox("Mốc Subscribers Tối Thiểu:", options=[100000, 250000, 500000, 1000000], index=3, format_func=lambda x: f"{x:,} Subs")
        min_duration_choice = st.selectbox("Lọc Yêu Cầu Đồ Dài Video:", options=[600], index=0, format_func=lambda x: "Bắt buộc có Video > 10 phút")

# --- TAB 4: UPLOAD DATA ---
with tab4:
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>📤 Upload file .ZIP, .TXT, .XLSX hoặc .DOCX để cập nhật Database</h3>", unsafe_allow_html=True)
    uploaded_files = st.file_uploader("Kéo thả file vào đây:", type=["zip", "txt", "xlsx", "xls", "docx", "doc", "csv"], accept_multiple_files=True)

# --- TAB 5: DATABASE CRM VIEWER ---
with tab5:
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>📊 Quản lý Database CRM Kênh</h3>", unsafe_allow_html=True)
    if 'new_db_channels_notify' in st.session_state: st.success(st.session_state['new_db_channels_notify'])

    try: res = supabase.table("channels").select("*").order("created_at", desc=True).execute()
    except Exception: res = supabase.table("channels").select("*").execute()

    if res.data:
        df_all = pd.DataFrame(res.data)
        st.markdown(f"Tổng số kênh hiện có trong DB: **{len(df_all)}**")
        st.divider()
        fc1, fc2, fc3, fc4 = st.columns([2, 2, 2, 2])
        with fc1: search_db = st.text_input("Tìm kiếm thủ công (Handle, Tên):", placeholder="Dán @handle...")
        with fc2: sel_sub_range = st.selectbox("Lọc theo Subs:", options=["-- Tất Cả Mốc Subs --", "< 100K Subs", "100K - 500K Subs", "500K - 1M Subs", "> 1M Subs"])
        with fc3: sel_source = st.selectbox("Lọc theo Nguồn:", options=["-- Tất Cả Nguồn --"] + list(df_all['source'].dropna().unique()) if 'source' in df_all.columns else ["-- Tất Cả Nguồn --"])
        with fc4: view_mode = st.radio("Chế độ hiển thị:", ["🎨 Card View (Thẻ chi tiết)", "📊 Table Grid View (Bảng nén gọn)"], horizontal=True)

        if view_mode == "📊 Table Grid View (Bảng nén gọn)":
            df_grid = df_all.copy()
            df_grid['Link Kênh'] = df_grid['handle'].apply(lambda h: get_channel_link(to_pure_id(h)))
            st.dataframe(df_grid, use_container_width=True)
        else:
            for idx, row in df_all.head(20).iterrows():
                p_id = to_pure_id(row['handle'])
                with st.container(border=True):
                    c1, c2, c3 = st.columns([3, 4, 2])
                    with c1: st.write(f"**@{p_id}** ({row.get('youtuber_name', 'N/A')})")
                    with c2: st.write(f"📁 Nguồn: {row.get('source', 'N/A')}")
                    with c3:
                        if st.button("🗑️ Xóa DB", key=f"del_db_main_{idx}_{p_id}", use_container_width=True):
                            delete_channel_from_system(p_id); st.rerun()

# --- TAB 6: SEO INSPECTOR ---
with tab6:
    st.markdown("<h3 style='font-weight: 700; margin-top: 15px;'>✨ Soi Từ Khóa Kênh (SEO Inspector)</h3>", unsafe_allow_html=True)
    inspect_handle_input = st.text_input("Nhập Handle Kênh cần soi:", value="@NickDiGiovanni")
    if st.button("🔍 Soi Từ Khóa Ngay", type="primary"):
        pure_inspect = to_pure_id(inspect_handle_input)
        if pure_inspect:
            active_keys = st.session_state.get('api_keys', [DEFAULT_API_KEY])
            cid_insp, _, _, _ = get_channel_id_by_handle_direct(pure_inspect, active_keys)
            if cid_insp:
                ext_data = extract_channel_master_keywords(cid_insp)
                st.success(f"✨ Đã soi từ khóa kênh @{pure_inspect}!")
                st.code(", ".join(ext_data['master_keywords']), language="text")
