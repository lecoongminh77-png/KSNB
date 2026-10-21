import streamlit as st

st.set_page_config(
    page_title="Lê Công Minh Workspace",
    page_icon="💼",
    layout="wide"
)

# CSS giao diện Apple Dashboard cao cấp
st.markdown("""
<style>
    [data-testid="stSidebarNav"] { display: none !important; }
    
    .stApp {
        background-color: #f5f5f7;
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", sans-serif;
        color: #1d1d1f;
    }
    
    /* Header */
    .header-box {
        text-align: center;
        padding: 40px 0 25px 0;
    }
    .main-title {
        font-size: 42px;
        font-weight: 700;
        letter-spacing: -0.025em;
        margin-bottom: 6px;
    }
    .sub-title {
        font-size: 17px;
        color: #86868b;
        font-weight: 400;
    }

    /* Thẻ ứng dụng (Card) */
    .app-card {
        background: #ffffff;
        border-radius: 20px;
        padding: 28px 24px 20px 24px;
        border: 1px solid rgba(0, 0, 0, 0.06);
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.04);
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        min-height: 220px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .app-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 28px rgba(0, 0, 0, 0.08);
        border-color: rgba(0, 113, 227, 0.2);
    }
    .badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
        margin-bottom: 12px;
    }
    .badge-green { background: #e3f8e5; color: #1e7e34; }
    .badge-blue { background: #e1f0ff; color: #0071e3; }
    
    .card-title {
        font-size: 22px;
        font-weight: 700;
        color: #1d1d1f;
        margin-bottom: 8px;
    }
    .card-desc {
        font-size: 14px;
        color: #6e6e73;
        line-height: 1.5;
        margin-bottom: 16px;
    }
    
    /* Nút bấm Apple Pill */
    div.stButton > button {
        background: #0071e3 !important;
        color: #ffffff !important;
        border-radius: 980px !important;
        border: none !important;
        padding:
