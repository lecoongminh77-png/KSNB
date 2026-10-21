import streamlit as st

st.set_page_config(
    page_title="Lê Công Minh Workspace",
    page_icon="💼",
    layout="wide"
)

# Giao diện phong cách Apple Dashboard
st.markdown("""
<style>
    [data-testid="stSidebarNav"] { display: none !important; }
    
    .stApp {
        background-color: #f5f5f7;
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", sans-serif;
        color: #1d1d1f;
    }
    
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
    
    div.stButton > button {
        background: #0071e3 !important;
        color: #ffffff !important;
        border-radius: 980px !important;
        border: none !important;
        padding: 10px 20px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        box-shadow: 0 4px 12px rgba(0, 113, 227, 0.2) !important;
        transition: all 0.2s ease !important;
    }
    div.stButton > button:hover {
        background: #0077ed !important;
        transform: scale(1.02) !important;
    }
</style>

<div class="header-box">
    <div class="main-title">💼 Lê Công Minh Workspace</div>
    <div class="sub-title">Cổng điều phối các công cụ tự động hóa & kiểm soát dữ liệu</div>
</div>
""", unsafe_allow_html=True)

# Lưới hiển thị 2 thẻ công cụ
_, col1, col2, _ = st.columns([1, 2.5, 2.5, 1], gap="medium")

with col1:
    st.markdown("""
    <div class="app-card">
        <div>
            <span class="badge badge-blue">AI Automation</span>
            <div class="card-title">✍️ Sforum Studio</div>
            <div class="card-desc">Hệ thống AI tự động phân tích bài viết gốc, áp dụng quy chuẩn Google Docs và biên tập bài viết công nghệ chuyên sâu chuẩn SEO.</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("Truy cập Sforum Studio ➔", key="btn_sforum", use_container_width=True):
        st.switch_page("pages/Sforum.py")

with col2:
    st.markdown("""
    <div class="app-card">
        <div>
            <span class="badge badge-green">Kiểm Soát Nội Bộ</span>
            <div class="card-title">📄 Đối Chiếu HRC</div>
            <div class="card-desc">Bộ công cụ tự động đối soát dữ liệu 4 bên (BBNT, Packing List, Trạm cân, ERP, Bảo vệ) và tự động tô đỏ cảnh báo sai phạm.</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("Mở Đối Chiếu Nghiệm Thu ➔", key="btn_ksnb", use_container_width=True):
        st.info("💡 Nếu bạn chuyển code Nghiệm thu sang file pages/Doi_Chieu.py thì cấu hình st.switch_page tại đây.")
