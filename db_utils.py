import json
import streamlit as st
from supabase import create_client, Client

@st.cache_resource
def init_supabase():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = init_supabase()

def load_api_keys_from_db():
    try:
        res = supabase.table("app_config").select("value").eq("key", "api_keys").execute()
        if res.data and len(res.data) > 0:
            return res.data[0]["value"]
    except Exception as e:
        print(f"[Supabase Log] Lỗi đọc API keys từ DB: {e}")
    return None

def save_api_keys_to_db(key_string):
    try:
        supabase.table("app_config").upsert({"key": "api_keys", "value": key_string}, on_conflict="key").execute()
        return True
    except Exception as e:
        print(f"[Supabase Log] Lỗi lưu API keys vào DB: {e}")
        return False

def load_cart_from_db():
    cart_dict = {}
    try:
        res = supabase.table("cart_items").select("*").execute()
        if res.data:
            for row in res.data:
                h = row["handle"]
                c_data = row.get("channel_data")
                if isinstance(c_data, str):
                    c_data = json.loads(c_data)
                cart_dict[h] = c_data
    except Exception as e:
        print(f"[Supabase Log] Lỗi tải Giỏ hàng: {e}")
    return cart_dict

def add_to_cart_db(pure_handle, channel_data):
    try:
        data_clean = dict(channel_data)
        if "recent_videos" in data_clean:
            del data_clean["recent_videos"]
        supabase.table("cart_items").upsert({"handle": pure_handle, "channel_data": data_clean}, on_conflict="handle").execute()
    except Exception as e:
        print(f"[Supabase Log] Lỗi thêm kênh vào Giỏ hàng DB: {e}")

def add_batch_to_cart_db(channels_list):
    try:
        rows = []
        for pure_handle, channel_data in channels_list:
            if not pure_handle:
                continue
            data_clean = dict(channel_data)
            if "recent_videos" in data_clean:
                del data_clean["recent_videos"]
            rows.append({"handle": pure_handle, "channel_data": data_clean})
        if rows:
            supabase.table("cart_items").upsert(rows, on_conflict="handle").execute()
    except Exception as e:
        print(f"[Supabase Log] Lỗi thêm lô vào Giỏ hàng DB: {e}")

def remove_from_cart_db(pure_handle):
    try:
        p_raw = str(pure_handle).strip()
        supabase.table("cart_items").delete().ilike("handle", p_raw).execute()
        supabase.table("cart_items").delete().eq("handle", p_raw).execute()
        supabase.table("cart_items").delete().eq("handle", p_raw.lower()).execute()
    except Exception as e:
        print(f"[Supabase Log] Lỗi xóa kênh khỏi Giỏ hàng DB: {e}")

def clear_cart_db():
    try:
        supabase.table("cart_items").delete().neq("handle", "___NONE___").execute()
    except Exception as e:
        print(f"[Supabase Log] Lỗi làm sạch Giỏ hàng DB: {e}")

def clear_entire_database(cb_clear_all=None):
    try:
        supabase.table("channels").delete().neq("handle", "___NONE___").execute()
        if cb_clear_all:
            cb_clear_all()
        return True
    except Exception as e:
        print(f"[Supabase Log] Lỗi xóa toàn bộ Database: {e}")
        return False

def load_campaigns():
    try:
        res = supabase.table("app_config").select("value").eq("key", "campaigns").execute()
        if res.data and len(res.data) > 0:
            return json.loads(res.data[0]["value"])
    except Exception as e:
        print(f"[Supabase Log] Lỗi đọc chiến dịch: {e}")
    return {}

def save_campaigns(camps_dict):
    try:
        supabase.table("app_config").upsert({"key": "campaigns", "value": json.dumps(camps_dict)}, on_conflict="key").execute()
    except Exception as e:
        print(f"[Supabase Log] Lỗi lưu chiến dịch: {e}")

@st.dialog("⚠️ CẢNH BÁO: XÓA SẠCH DATABASE", width="small")
def confirm_clear_db_dialog(cb_clear_all=None):
    st.error("🚨 Hành động này sẽ XÓA VĨNH VIỄN toàn bộ danh sách kênh trong Supabase và KHÔNG THỂ HỒI PHỤC!")
    st.write("Vui lòng gõ **`XOA DATABASE`** vào ô bên dưới để xác nhận:")
    confirm_txt = st.text_input("Xác nhận:", key="input_confirm_db_wipe")
    if st.button("💣 XÁC NHẬN XÓA SẠCH DATABASE", type="primary", use_container_width=True):
        if confirm_txt.strip().upper() == "XOA DATABASE":
            if clear_entire_database(cb_clear_all):
                st.success("🎉 Đã xóa sạch vĩnh viễn toàn bộ Database!")
                st.rerun()
            else:
                st.error("❌ Đã xảy ra lỗi khi kết nối Supabase!")
        else:
            st.warning("⚠️ Mã xác nhận không đúng! Vui lòng gõ 'XOA DATABASE'.")
