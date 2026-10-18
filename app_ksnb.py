# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN CHUẨN APPLE & THANH TIỆN ÍCH
# -------------------------------------------------------------
st.set_page_config(page_title="Tiện Ích - KSNB", page_icon="🛡️", layout="wide")

# CSS ẩn menu sidebar và tạo phong cách Apple
st.markdown("""
<style>
    [data-testid="stSidebarNav"] { display: none !important; }
    .stApp {
        background-color: #f5f5f7;
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", sans-serif;
        color: #1d1d1f;
    }
    .apple-header-box {
        display: flex;
        justify-content: space-between;
        align-items: flex-end;
        padding: 30px 0 15px 0;
        margin-bottom: 10px;
    }
    .apple-main-title {
        font-size: 46px;
        font-weight: 700;
        color: #1d1d1f;
        letter-spacing: -0.025em;
        margin: 0;
    }
    .apple-sub-text {
        font-size: 19px;
        font-weight: 600;
        color: #1d1d1f;
        text-align: right;
        margin: 0;
        line-height: 1.35;
    }
    .apple-sub-text span {
        color: #0071e3;
        font-size: 15px;
        display: block;
        margin-top: 4px;
        font-weight: 500;
    }
    div.stButton > button:first-child {
        background: #0071e3 !important;
        color: #ffffff !important;
        border-radius: 980px !important;
        border: none !important;
        padding: 12px 24px !important;
        font-weight: 600 !important;
        font-size: 15px !important;
        box-shadow: 0 4px 12px rgba(0, 113, 227, 0.2) !important;
    }
</style>
""", unsafe_allow_html=True)

# Header Tiện Ích
st.markdown("""
<div class="apple-header-box">
    <div>
        <h1 class="apple-main-title">🛠️ Tiện Ích</h1>
    </div>
    <div>
        <div class="apple-sub-text">
            Nền tảng công cụ hỗ trợ<br>công việc thông minh.
            <span>Lê Công Minh Workspace ›</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# Thanh nút bấm tiện ích
col_nav1, col_nav2, col_nav3, _ = st.columns([1.5, 1.5, 1.8, 5])
with col_nav1:
    # Bấm Sforum là nhảy sang trang viết bài Sforum ngay
    if st.button("✍️ Sforum", use_container_width=True):
        st.switch_page("pages/Sforum.py")
with col_nav2:
    st.button("🛡️ KSNB", use_container_width=True)
with col_nav3:
    st.button("➕ Thêm tiện ích", disabled=True, use_container_width=True)

st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 15px 0 25px 0;'>", unsafe_allow_html=True)
