import streamlit as st

# 1. Cấu hình trang
st.set_page_config(
    page_title="Lê Công Minh Workspace",
    page_icon="💼",
    layout="wide"
)

# 2. Giao diện phong cách Apple tối giản
st.markdown("""
<style>
    /* Ẩn sidebar mặc định */
    [data-testid="stSidebarNav"] { display: none !important; }
    
    .stApp {
        background-color: #f5f5f7;
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", sans-serif;
        color: #1d1d1f;
    }
    
    .header-box {
        text-align: center;
        padding: 50px 0 30px 0;
    }
    
    .main-title {
        font-size: 40px;
        font-weight: 700;
        letter-spacing: -0.02em;
        margin-bottom: 8px;
    }
    
    .sub-title {
        font-size: 17px;
        color: #86868b;
        font-weight: 400;
    }
    
    /* Tùy chỉnh nút bấm dạng viên thuốc */
    div.stButton > button {
        background: #0071e3 !important;
        color: #ffffff !important;
        border-radius: 980px !important;
        border: none !important;
        padding: 14px 28px !important;
        font-weight: 600 !important;
        font-size: 16px !important;
        box-shadow: 0 4px 14px rgba(0, 113, 227, 0.25) !important;
        transition: all 0.2s ease !important;
        height: auto !important;
    }
    
    div.stButton > button:hover {
        background: #0077ed !important;
        transform: translateY(-2px);
        box-shadow: 0 6px 18px rgba(0, 113, 227, 0.35) !important;
    }
</style>

<div class="header-box">
    <div class="main-title">💼 Lê Công Minh Workspace</div>
    <div class="sub-title">Chọn công cụ làm việc bạn muốn truy cập</div>
</div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# 3. Khu vực chứa các nút chức năng (Căn giữa màn hình)
_, center_col, _ = st.columns([1, 2, 1])

with center_col:
    # Nút bấm chuyển hướng sang trang Sforum
    if st.button("✍️ Truy cập Sforum Content Studio", use_container_width=True):
        st.switch_page("pages/Sforum.py")

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

    # Nút bấm dự phòng hoặc tiện ích khác (nếu có trang nghiệm thu trong pages/)
    # if st.button("📄 Đối Chiếu Nghiệm Thu HRC", use_container_width=True):
    #     st.switch_page("pages/Doi_Chieu_HRC.py")
