import streamlit as st
import requests
from streamlit_lottie import st_lottie
import time  # <-- THÊM THƯ VIỆN NÀY ĐỂ ĐẾM THỜI GIAN

st.set_page_config(
    page_title="Lê Công Minh",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="collapsed"
)
st.markdown("""
<style>
    section[data-testid="stSidebar"], 
    [data-testid="stSidebar"], 
    [data-testid="collapsedControl"] {
        display: none !important;
    }
</style>
""", unsafe_allow_html=True)

# Hàm tải Lottie
def load_lottieurl(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()

# 1. KHỞI TẠO BỘ NHỚ KIỂM TRA ĐÃ QUA MÀN HÌNH INTRO CHƯA
if 'da_xem_intro' not in st.session_state:
    st.session_state.da_xem_intro = False

# 2. HIỆN MÀN HÌNH WELCOME TỰ ĐỘNG CHUYỂN TRANG
if not st.session_state.da_xem_intro:
    # Tạo một khung chứa tạm thời để hiển thị animation
    splash_placeholder = st.empty()
    
    with splash_placeholder.container():
        st.markdown("<br><br>", unsafe_allow_html=True)
        
        # Link Lottie của bạn
        lottie_welcome = load_lottieurl("https://lottie.host/4a9438d1-b8fb-4a3b-aed8-58a48b9413e2/BnRtyhhxBR.json") 
        
        # Hiển thị hình động
        st_lottie(lottie_welcome, height=350, key="welcome_animation")
        
    # Cho hệ thống dừng 3 giây để người dùng kịp nhìn animation
    time.sleep(5)
    
    # Đánh dấu là đã xem intro
    st.session_state.da_xem_intro = True
    
    # Xóa cái khung tạm thời chứa animation đi
    splash_placeholder.empty()
    
    # Tải lại trang (lúc này da_xem_intro = True nên sẽ bỏ qua màn hình Welcome và load trang chính)
    st.rerun()

# ==========================================
# GIAO DIỆN TRANG CHÍNH BẮT ĐẦU TỪ ĐÂY
# ==========================================

# Giao diện phong cách Apple Dashboard
st.markdown("""
<style>
[data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }

    
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

# Hoạt ảnh quả bóng xoay chính giữa Dashboard
_, col_center, _ = st.columns([3.8, 2.4, 3.8])
with col_center:
    logo_header = load_lottieurl("https://lottie.host/aa244008-30a5-46ac-a553-63b9d65b315a/kWzva6TSwg.json")
    if logo_header:
        st_lottie(logo_header, height=170, key="header_logo")

""", unsafe_allow_html=True)

# Lưới hiển thị các thẻ công cụ (Chia làm 3 cột chính)
_, col1, col2, col3, _ = st.columns([0.2, 3, 3, 3, 0.2], gap="medium")

with col1:
    st.markdown("""
    <div class="app-card" style="margin-bottom: 20px;">
        <div>
            <span class="badge badge-blue">AI Automation</span>
            <div class="card-title">✍️ Sforum Studio</div>
            <div class="card-desc">Hệ thống AI tự động phân tích bài viết gốc, áp dụng quy chuẩn Google Docs và biên tập bài viết.</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("Truy cập Sforum Studio ➔", key="btn_sforum", use_container_width=True):
        st.switch_page("pages/Sforum.py")

    st.markdown("""
    <div class="app-card">
        <div>
            <span class="badge badge-green">Trạm Cân & ERP</span>
            <div class="card-title">🚛 Hàng Ra Cổng</div>
            <div class="card-desc">Đối soát 4 bên (ERP 021, 105 và Trạm Cân NM, SMC), tính dung sai tự động.</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("Mở Đối Soát Cổng ➔", key="btn_cong", use_container_width=True):
        st.switch_page("pages/Kiem_Tra_Ra_Cong.py")

with col2:
    st.markdown("""
    <div class="app-card" style="margin-bottom: 20px;">
        <div>
            <span class="badge badge-green">Kiểm Soát Nội Bộ</span>
            <div class="card-title">📄 Đối Chiếu HRC</div>
            <div class="card-desc">Bộ công cụ tự động đối soát dữ liệu 4 bên và tự động tô đỏ cảnh báo sai phạm.</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("Mở Đối Chiếu Nghiệm Thu ➔", key="btn_ksnb", use_container_width=True):
        st.switch_page("pages/Doi_Chieu_HRC.py")

    # ĐÂY LÀ THẺ MỚI BỔ SUNG CHO TỔNG HỢP NXT
    st.markdown("""
    <div class="app-card">
        <div>
            <span class="badge badge-blue">Quản Lý Tồn Kho</span>
            <div class="card-title">📊 Tổng Hợp NXT</div>
            <div class="card-desc">Gộp tự động file INV 110 và 105 để xuất báo cáo Nhập Xuất Tồn và tình hình sử dụng NVL.</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("Mở Dashboard NXT ➔", key="btn_nxt", use_container_width=True):
        st.switch_page("pages/NXT_Tong_Hop.py")

with col3:
    st.markdown("""
    <div class="app-card">
        <div>
            <span class="badge badge-blue">Quản Lý Tồn Kho</span>
            <div class="card-title">📦 Kiểm Tra Mẫu</div>
            <div class="card-desc">Tự động điền công thức kiểm tra tồn kho cuộn: Hạng, Số ký chuẩn, Chênh lệch và Trùng mã.</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("Mở Kiểm Tra Mẫu ➔", key="btn_2356", use_container_width=True):
        st.switch_page("pages/Kiem_Tra_Mau.py")

