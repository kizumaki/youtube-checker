import io
import zipfile
import pandas as pd
import requests
import streamlit as st
from db_utils import add_to_cart_db, remove_from_cart_db, clear_cart_db, load_campaigns, save_campaigns, add_batch_to_cart_db, supabase
from yt_utils import render_social_badges_html, to_pure_id, get_channel_link, run_single_channel_audit

def inject_theme_css(is_dark, bg_color, card_bg, text_color, border_color, sidebar_bg):
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap');
    .stApp {{ background-color: {bg_color} !important; color: {text_color} !important; font-family: 'Montserrat', sans-serif !important; }}
    section[data-testid="stSidebar"] {{ background-color: {sidebar_bg} !important; border-right: 1px solid {border_color} !important; box-shadow: 4px 0 15px rgba(0, 0, 0, 0.05) !important; }}
    header[data-testid="stHeader"] {{ background-color: transparent !important; }}
    .stTabs [data-baseweb="tab-list"] {{ gap: 32px; background-color: transparent; padding: 0 0 4px 0; border-bottom: 2px solid #D1D5DB; }}
    .stTabs [data-baseweb="tab"] {{ background-color: transparent !important; border: none !important; border-bottom: 3px solid transparent !important; border-radius: 0 !important; color: #6B7280 !important; font-weight: 700; font-size: 0.9rem; padding: 10px 4px; text-transform: uppercase; letter-spacing: 0.05em; transition: all 0.3s ease !important; cursor: pointer !important; }}
    .stTabs [data-baseweb="tab"]:hover {{ color: #D95F26 !important; transform: translateY(-1px); }}
    .stTabs [aria-selected="true"] {{ color: #D95F26 !important; border-bottom: 3px solid #D95F26 !important; }}
    div[data-testid="stVerticalBlockBorderWrapper"] {{ background-color: {card_bg} !important; border: 1px solid {border_color} !important; border-radius: 12px !important; padding: 12px !important; box-shadow: 0 2px 10px rgba(0, 0, 0, 0.03) !important; transition: transform 0.2s ease, box-shadow 0.2s ease; margin-bottom: 15px !important; }}
    div[data-testid="stVerticalBlockBorderWrapper"]:hover {{ box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08) !important; }}
    div[data-testid="stVerticalBlockBorderWrapper"]:has(div.active-card-marker) {{ border: 2px solid #D95F26 !important; box-shadow: 0 8px 24px rgba(217, 95, 38, 0.2) !important; }}
    div[data-testid="stVerticalBlockBorderWrapper"]:has(div.in-cart-marker) {{ border: 2px solid #47A5D1 !important; box-shadow: 0 8px 24px rgba(71, 165, 209, 0.2) !important; }}
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.action-bar-marker) {{ border: 2px solid #D95F26 !important; border-top: 5px solid #D95F26 !important; padding: 10px 14px !important; margin-bottom: 25px !important; box-shadow: 0 6px 15px rgba(0,0,0,0.08) !important; }}
    .active-banner-tag {{ background-color: #D95F26 !important; color: #FFFFFF !important; padding: 6px 14px !important; border-radius: 8px !important; font-weight: 800 !important; margin-bottom: 12px !important; font-size: 0.85rem !important; letter-spacing: 0.05em !important; display: inline-block !important; box-shadow: 0 3px 10px rgba(217, 95, 38, 0.25) !important; }}
    .active-banner-tag * {{ color: #FFFFFF !important; }}
    .in-cart-banner-tag {{ background-color: #47A5D1 !important; color: #FFFFFF !important; padding: 6px 14px !important; border-radius: 8px !important; font-weight: 800 !important; margin-bottom: 12px !important; font-size: 0.85rem !important; display: inline-block !important; box-shadow: 0 3px 10px rgba(71, 165, 209, 0.25) !important; }}
    .in-cart-banner-tag * {{ color: #FFFFFF !important; }}
    .stTextInput input, .stTextArea textarea, .stSelectbox select {{ background-color: {card_bg} !important; color: {text_color} !important; border: 1px solid #D1D5DB !important; border-radius: 8px !important; font-family: 'Montserrat', sans-serif !important; }}
    .stTextInput input:focus, .stTextArea textarea:focus, .stSelectbox select:focus {{ border-color: #D95F26 !important; box-shadow: 0 0 0 1px #D95F26 !important; }}
    .stButton button {{ border-radius: 8px !important; font-weight: 700 !important; font-family: 'Montserrat', sans-serif !important; border: 1px solid #D1D5DB !important; background-color: {card_bg} !important; color: {text_color} !important; text-transform: uppercase; font-size: 0.8rem !important; letter-spacing: 0.05em; transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important; }}
    .stButton button:hover {{ border-color: #D95F26 !important; color: #D95F26 !important; transform: translateY(-1px) !important; }}
    .stButton button[kind="primary"], .stButton button[kind="primary"] *, .stButton button[kind="primary"] p, .stButton button[kind="primary"] span, .stButton button[kind="primary"] div {{ background: linear-gradient(135deg, #D95F26 0%, #E66A32 100%) !important; color: #FFFFFF !important; border: none !important; box-shadow: 0 3px 10px rgba(217, 95, 38, 0.22) !important; }}
    .stButton button[kind="primary"]:hover, .stButton button[kind="primary"]:hover *, .stButton button[kind="primary"]:hover p, .stButton button[kind="primary"]:hover span, .stButton button[kind="primary"]:hover div {{ background: linear-gradient(135deg, #C24E18 0%, #D95F26 100%) !important; color: #FFFFFF !important; transform: translateY(-1px) !important; box-shadow: 0 6px 16px rgba(217, 95, 38, 0.32) !important; }}
    .social-badge {{ display: inline-block; padding: 4px 10px; border-radius: 6px; font-size: 0.75rem; font-weight: 700; margin-right: 6px; margin-bottom: 6px; text-decoration: none !important; color: #FFFFFF !important; box-shadow: 0 2px 5px rgba(0,0,0,0.15); }}
    .social-email {{ background-color: #EA4335 !important; }}
    .social-ig {{ background-color: #E1306C !important; }}
    .social-tt {{ background-color: #000000 !important; border: 1px solid #555; }}
    .social-x {{ background-color: #1DA1F2 !important; }}
    .social-discord {{ background-color: #5865F2 !important; }}
    .social-fb {{ background-color: #1877F2 !important; }}
    .social-web {{ background-color: #10B981 !important; }}
    .badge-pro {{ display: inline-block; padding: 6px 12px; border-radius: 9999px; font-size: 0.7rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.08em; }}
    .badge-ocean, .badge-ocean * {{ background-color: #47A5D1 !important; color: #FFFFFF !important; border: none !important; }}
    .badge-score {{ padding: 4px 8px; border-radius: 6px; font-weight: 800; font-size: 0.8rem; background-color: #FFF2EB; color: #D95F26; border: 1px solid #D95F26; display: inline-block; }}
    .badge-stt {{ font-weight: 800; font-size: 0.85rem; color: #6B7280; background-color: #E5E7EB; padding: 2px 8px; border-radius: 6px; margin-right: 8px; inline-block; }}
    </style>
    """, unsafe_allow_html=True)

def render_kpi_cards(kpi_data, card_bg, border_color):
    cards_html = ""
    for title, val, color in kpi_data:
        cards_html += f'<div style="flex: 1; min-width: 200px; background-color: {card_bg}; border: 1px solid {border_color}; border-radius: 12px; padding: 18px; text-align: center; box-shadow: 0 4px 10px rgba(0,0,0,0.04);"><div style="font-size: 0.8rem; color: #6B7280; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 6px;">{title}</div><div style="font-size: 2.2rem; font-weight: 900; color: {color};">{val}</div></div>'
    st.markdown(f'<div style="display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap;">{cards_html}</div>', unsafe_allow_html=True)

@st.dialog("🤖 AI Soạn Mail Hợp Tác", width="large")
def show_ai_email_dialog(channel_data):
    st.markdown(f"<h3 style='color: #D95F26; font-weight: 800;'>Thư ngỏ gửi <span style='color: #47A5D1;'>{channel_data.get('Tên Kênh', 'Creator')}</span></h3>", unsafe_allow_html=True)
    st.caption("✨ AI đã tự động phân tích chỉ số và thiết kế thư cá nhân hóa. Bạn có thể chỉnh sửa trước khi Copy.")
    er_text = channel_data.get('ER', 'rất ấn tượng')
    if er_text != 'N/A' and er_text != 'rất ấn tượng': er_text = f"lên tới {er_text}"
        
    template = f"""Subject: Collaboration Opportunity with {channel_data.get('Tên Kênh', 'your channel')} 🚀

Hi {channel_data.get('Tên Kênh', 'there')},

I’ve been closely following your content on YouTube, especially your latest videos, and I am absolutely blown away by the quality! It’s no surprise your channel is growing so fast, with an incredible engagement rate {er_text}. 

I am reaching out from Backstreet Voice Studio. We specialize in high-end voiceover, localization, and audio optimization, helping top-tier creators like you expand their reach into new global markets effortlessly. 

Given your strong audience retention and niche focus, translating and dubbing your content could easily double your current viewership without requiring any extra production effort on your end.

Would you be open to a quick 10-minute chat this week to see how we could partner up to maximize your channel's revenue? 

Keep up the amazing work!

Best regards,
[Your Name/Title]
Backstreet Voice Studio"""
    st.text_area("Bản Thảo Email (Sẵn sàng Copy):", value=template, height=350)
    if st.button("❌ Đóng Cửa Sổ", type="primary", use_container_width=True): st.rerun()

def render_shared_cart_ui(key_suffix="cart_ui"):
    st.divider()
    cart_items = st.session_state['cart']
    col_t1, col_t2 = st.columns([7, 3])
    with col_t1: st.markdown(f"<h3 style='font-weight: 800;'>🛒 Giỏ Hàng Dùng Chung ({len(cart_items)} Kênh)</h3>", unsafe_allow_html=True)
    with col_t2:
        with st.expander("📁 Quản Lý Chiến Dịch"):
            camps = load_campaigns()
            new_camp_name = st.text_input("Tên chiến dịch mới:", key=f"new_camp_name_{key_suffix}")
            if st.button("💾 Lưu Giỏ Hàng Thành Chiến Dịch", use_container_width=True, key=f"save_camp_btn_{key_suffix}"):
                if new_camp_name: camps[new_camp_name] = cart_items; save_campaigns(camps); st.success(f"Đã lưu chiến dịch '{new_camp_name}'!")
                else: st.warning("Vui lòng nhập tên chiến dịch.")
            sel_camp = st.selectbox("Tải lại chiến dịch cũ:", options=["-- Chọn --"] + list(camps.keys()), key=f"sel_camp_{key_suffix}")
            if st.button("📂 Tải Dữ Liệu", use_container_width=True, key=f"load_camp_btn_{key_suffix}"):
                if sel_camp != "-- Chọn --": st.session_state['cart'] = camps[sel_camp]; clear_cart_db(); add_batch_to_cart_db([(k, v) for k, v in st.session_state['cart'].items()]); st.success(f"Đã tải thành công chiến dịch {sel_camp}!"); st.rerun()

    if cart_items:
        df_cart = pd.DataFrame(list(cart_items.values()))
        if 'Handle' in df_cart.columns:
            df_cart['Tab Videos'] = df_cart['Handle'].apply(lambda h: f"{get_channel_link(to_pure_id(h))}/videos" if to_pure_id(h) else "")
            df_cart['Link Kênh'] = df_cart['Handle'].apply(lambda h: get_channel_link(to_pure_id(h)) if to_pure_id(h) else "")
        if 'recent_videos' in df_cart.columns: df_cart = df_cart.drop(columns=['recent_videos'])
        df_cart.index = range(1, len(df_cart) + 1)

        st.dataframe(df_cart, use_container_width=True, column_config={
            "Link Kênh": st.column_config.LinkColumn("Trang Chủ", display_text="🏠 Kênh"),
            "Tab Videos": st.column_config.LinkColumn("Tab Videos", display_text="🎬 Videos"),
            "Tag": st.column_config.TextColumn("🏷️ Nhãn Trạng Thái")
        })
        
        with st.expander("🚀 Đẩy Dữ Liệu & Xuất File (Google Sheets / Excel / Audit ZIP)", expanded=True):
            c1, c2, c3, c4 = st.columns([1.5, 1.5, 3.2, 1.8])
            c1.download_button("📄 Tải TXT", data="\n".join([i["Handle"] for i in cart_items.values()]), file_name="gio_hang_dung_chung.txt", use_container_width=True, key=f"dl_txt_cart_{key_suffix}")
            buf_xl = io.BytesIO(); df_cart.to_excel(buf_xl, index=False)
            c2.download_button("📊 Tải Excel", data=buf_xl.getvalue(), file_name="gio_hang_dung_chung.xlsx", use_container_width=True, key=f"dl_xl_cart_{key_suffix}")
            
            if c3.button("⚡ NẠP DB & TẠO BÁO CÁO AUDIT", type="primary", use_container_width=True, key=f"push_db_cart_{key_suffix}"):
                data_db = [{"handle": to_pure_id(i["Handle"]), "youtuber_name": i.get("Tên Kênh", ""), "source": f"Cart Import [{i.get('Tag', '')}]"} for i in cart_items.values()]
                supabase.table("channels").upsert(data_db, on_conflict="handle").execute()
                for k in list(st.session_state.keys()):
                    if k.startswith('crm_cache_') or k == 'tab5_crm_cache': st.session_state.pop(k, None)
                st.session_state['new_db_channels_notify'] = f"🎉 Vừa nạp thành công {len(data_db)} kênh mới vào Database! Tất cả kênh mới được ưu tiên hiển thị ở đầu danh sách."

                handles_to_audit = [to_pure_id(i["Handle"]) for i in cart_items.values() if to_pure_id(i["Handle"])]
                tot_cart_audit = len(handles_to_audit)
                
                if tot_cart_audit > 0:
                    prog_audit = st.progress(0); stat_audit = st.empty(); comp_audit = 0; audit_results = []
                    active_keys_cart = st.session_state.get('api_keys', [])
                    exhausted_set_cart = set(st.session_state.get('exhausted_keys_set', set()))

                    with ThreadPoolExecutor(max_workers=5) as executor:
                        futures = {executor.submit(run_single_channel_audit, p_h, active_keys_cart, exhausted_set_cart): p_h for p_h in handles_to_audit}
                        for future in as_completed(futures):
                            comp_audit += 1; prog_audit.progress(comp_audit / tot_cart_audit)
                            stat_audit.markdown(f"⏳ **Đang cào dữ liệu & Dựng Audit V4.14:** `{comp_audit}/{tot_cart_audit}` kênh...")
                            try:
                                res_val = future.result()
                                if res_val and res_val[0] and res_val[1]:
                                    b_bytes, f_name, logs = res_val
                                    audit_results.append((f_name, b_bytes))
                            except Exception: pass

                    prog_audit.empty(); stat_audit.empty()
                    if audit_results:
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for fname, fbytes in audit_results: zip_file.writestr(fname, fbytes)
                        st.session_state[f"cart_audit_zip_{key_suffix}"] = {
                            "bytes": zip_buffer.getvalue(),
                            "filename": f"Goi_Bao_Cao_Audit_V414_{datetime.datetime.now().strftime('%d-%m-%Y')}.zip",
                            "count": len(audit_results)
                        }
                        st.success(f"🎉 Đã nạp thành công {len(data_db)} kênh vào DB & Dựng xong {len(audit_results)} file Báo cáo Audit V4.14!"); st.rerun()

            if c4.button("🧹 Xóa Giỏ Hàng", use_container_width=True, key=f"clear_cart_{key_suffix}"): 
                clear_cart_db(); st.session_state['cart'] = {}; st.session_state.pop(f"cart_audit_zip_{key_suffix}", None); st.success("🎉 Đã xóa sạch Giỏ Hàng!"); st.rerun()

            zip_key = f"cart_audit_zip_{key_suffix}"
            if zip_key in st.session_state:
                st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
                zip_info = st.session_state[zip_key]
                st.download_button(
                    label=f"📦 TẢI GÓI AUDIT ZIP ({zip_info['count']} FILE EXCEL)",
                    data=zip_info["bytes"], file_name=zip_info["filename"], mime="application/zip", type="primary", use_container_width=True, key=f"dl_zip_btn_{key_suffix}"
                )

            st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
            wh_col1, wh_col2 = st.columns([3, 1])
            with wh_col1: webhook_url = st.text_input("🔗 Dán Google Apps Script Webhook URL (Để đồng bộ lên Sheets):", key=f"webhook_url_{key_suffix}")
            with wh_col2: 
                st.write("")
                if st.button("🚀 Push to Google Sheets", type="primary", use_container_width=True, key=f"push_gsheets_btn_{key_suffix}"):
                    if webhook_url:
                        try:
                            requests.post(webhook_url, json={"data": df_cart.to_dict(orient="records")})
                            st.success("🎉 Đã đẩy dữ liệu thành công!")
                        except Exception as e: st.error(f"Lỗi: {e}")
                    else: st.warning("Vui lòng dán Webhook URL!")
    else: st.info("Giỏ hàng đang trống. Bấm '🛒 Thêm' ở Tab 1 hoặc Tab 3 để nhặt kênh vào giỏ!")
