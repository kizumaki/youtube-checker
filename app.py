# Custom Scrollbar & Advanced UI Injection
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap');

/* Base Theme */
.stApp {{ background-color: {bg_color} !important; color: {text_color} !important; font-family: 'Montserrat', sans-serif !important; }}
section[data-testid="stSidebar"] {{ background-color: {sidebar_bg} !important; border-right: 1px solid {border_color} !important; box-shadow: 4px 0 15px rgba(0, 0, 0, 0.05) !important; }}

/* Custom Scrollbar */
::-webkit-scrollbar {{ width: 8px; height: 8px; }}
::-webkit-scrollbar-track {{ background: {bg_color}; }}
::-webkit-scrollbar-thumb {{ background: #D95F26; border-radius: 4px; }}
::-webkit-scrollbar-thumb:hover {{ background: #C24E18; }}

/* HIGH-END ARTISTIC TABS */
.stTabs [data-baseweb="tab-list"] {{ gap: 24px; background-color: transparent; border-bottom: 2px solid #E5E7EB; }}
.stTabs [data-baseweb="tab"] {{ background-color: transparent !important; border: none !important; border-bottom: 3px solid transparent !important; color: #6B7280 !important; font-weight: 700; font-size: 0.88rem; padding: 12px 6px; text-transform: uppercase; letter-spacing: 0.04em; transition: all 0.25s ease !important; }}
.stTabs [data-baseweb="tab"]:hover {{ color: #D95F26 !important; }}
.stTabs [aria-selected="true"] {{ color: #D95F26 !important; border-bottom: 3px solid #D95F26 !important; }}

/* Standard Card Container Styling with Hover Scale */
div[data-testid="stVerticalBlockBorderWrapper"] {{ background-color: {card_bg} !important; border: 1px solid {border_color} !important; border-radius: 14px !important; padding: 14px !important; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03) !important; transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1); margin-bottom: 16px !important; }}
div[data-testid="stVerticalBlockBorderWrapper"]:hover {{ transform: translateY(-2px); box-shadow: 0 8px 24px rgba(217, 95, 38, 0.12) !important; border-color: #D95F26 !important; }}

/* Pill Badges */
.pill-badge {{ display: inline-flex; align-items: center; padding: 3px 10px; border-radius: 9999px; font-size: 0.75rem; font-weight: 700; margin-right: 6px; }}
.pill-green {{ background-color: #DEF7EC; color: #03543F; }}
.pill-orange {{ background-color: #FEF3C7; color: #92400E; }}
.pill-blue {{ background-color: #E0F2FE; color: #1E40AF; }}

/* Modern Score Bar */
.score-container {{ display: inline-block; background: #FFF2EB; border: 1px solid #D95F26; border-radius: 8px; padding: 4px 10px; font-weight: 800; font-size: 0.8rem; color: #D95F26; }}
</style>
""", unsafe_allow_html=True)
