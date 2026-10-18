import streamlit as st
import re
import requests
from bs4 import BeautifulSoup
from google import genai
from google.genai import types

# -------------------------------------------------------------
# 1. CẤU HÌNH GIAO DIỆN CHUẨN APPLE STORE
# -------------------------------------------------------------
st.set_page_config(
    page_title="Tiện Ích - Sforum Studio",
    page_icon="🍎",
    layout="wide"
)

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
        padding: 25px 0 15px 0;
        margin-bottom: 5px;
    }
    .apple-main-title {
        font-size: 44px;
        font-weight: 700;
        color: #1d1d1f;
        letter-spacing: -0.025em;
        margin: 0;
    }
    .apple-sub-text {
        font-size: 18px;
        font-weight: 600;
        color: #1d1d1f;
        text-align: right;
        margin: 0;
        line-height: 1.35;
    }
    .apple-sub-text span {
        color: #0071e3;
        font-size: 14px;
        display: block;
        margin-top: 4px;
        font-weight: 500;
    }
    div.stButton > button:first-child {
        background: #0071e3 !important;
        color: #ffffff !important;
        border-radius: 980px !important;
        border: none !important;
        padding: 10px 24px !important;
        font-weight: 600 !important;
        font-size: 15px !important;
        box-shadow: 0 4px 12px rgba(0, 113, 227, 0.2) !important;
        transition: all 0.2s ease !important;
    }
    div.stButton > button:first-child:hover {
        background: #0077ed !important;
        transform: scale(1.02) !important;
    }
    .stTextInput > div > div > input, 
    .stTextArea > div > div > textarea, 
    .stSelectbox > div > div {
        background-color: #ffffff !important;
        border: 1px solid #d2d2d7 !important;
        border-radius: 12px !important;
        color: #1d1d1f !important;
        font-size: 15px !important;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; background-color: transparent; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 980px;
        font-weight: 600;
        padding: 8px 18px;
        background: #e8e8ed;
        color: #6e6e73;
        border: none;
    }
    .stTabs [aria-selected="true"] { background: #1d1d1f !important; color: #ffffff !important; }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# 2. HEADER TIỆN ÍCH CHUẨN APPLE
# -------------------------------------------------------------
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

# Dãy nút chuyển đổi tiện ích
col_nav1, col_nav2, col_nav3, _ = st.columns([1.5, 1.5, 1.8, 5])
with col_nav1:
    st.button("✍️ Sforum", use_container_width=True)
with col_nav2:
    if st.button("🛡️ KSNB", use_container_width=True):
        st.switch_page("app_ksnb.py")
with col_nav3:
    st.button("➕ Thêm tiện ích", disabled=True, use_container_width=True)

st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 15px 0 20px 0;'>", unsafe_allow_html=True)

# -------------------------------------------------------------
# 3. KẾT NỐI VÀ ĐỌC MẪU TRỰC TIẾP TỪ GOOGLE DOCS
# -------------------------------------------------------------
st.markdown("""
<div style="margin-bottom: 12px;">
    <h2 style="font-size: 26px; font-weight: 700; color: #1d1d1f; margin: 0;">Sforum Content Studio. <span style="color: #6e6e73; font-weight: 500;">Tự động học mẫu từ Google Docs.</span></h2>
</div>
""", unsafe_allow_html=True)

# Hàm tải văn bản từ Google Doc
def fetch_google_doc_text(doc_url: str):
    try:
        match = re.search(r"/document/d/([a-zA-Z0-9-_]+)", doc_url)
        doc_id = match.group(1) if match else doc_url.strip()
        export_url = f"https://docs.google.com/document/d/{doc_id}/export?format=txt"
        resp = requests.get(export_url, timeout=12)
        if resp.status_code == 200 and "accounts.google.com" not in resp.url:
            return resp.text.strip()
        return None
    except Exception:
        return None

# Khung cài đặt liên kết Google Doc quy chuẩn
with st.expander("📖 Cấu hình file Google Docs Quy chuẩn (Bấm để xem/sửa link)", expanded=False):
    doc_link = st.text_input(
        "Đường dẫn file Google Doc chứa mẫu quy chuẩn của bạn:",
        value="https://docs.google.com/document/d/1_bl4RrRBs4mv521i45FqVJzUe9Yh4ZOlJtk31QR0tuQ/edit"
    )
    st.caption("💡 Lưu ý: File Google Doc cần được bật quyền chia sẻ: **'Bất kỳ ai có đường liên kết đều có thể xem'** để hệ thống có thể đọc tự động.")

# Đọc thử nội dung từ Doc
loaded_rules = fetch_google_doc_text(doc_link)
if loaded_rules:
    st.success(f"🟢 Đã kết nối thành công với Google Docs ({len(loaded_rules.split())} từ quy chuẩn). Khi bạn sửa file Google Doc, web sẽ tự cập nhật theo!")
else:
    st.warning("⚠️ Chưa đọc được file Google Doc (vui lòng kiểm tra quyền chia sẻ 'Bất kỳ ai có liên kết đều xem được'). Hệ thống đang dùng bộ quy chuẩn dự phòng.")

# Lấy API Key
api_key = None
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    api_key = st.sidebar.text_input("🔑 Gemini API Key:", type="password")

# Hàm cào dữ liệu từ URL bài gốc
def extract_content_from_url(url: str):
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9,vi;q=0.8",
        }
        resp = requests.get(url, headers=headers, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        
        for s in soup(["script", "style", "nav", "footer", "header", "noscript", "aside"]):
            s.decompose()
            
        title = soup.find("h1")
        title_text = title.get_text(strip=True) if title else ""
        
        article_body = soup.find("article") or soup.find(id=re.compile(r"article-body|content|main", re.I)) or soup.find("main")
        target = article_body if article_body else soup
        
        paragraphs = []
        for elem in target.find_all(["h2", "h3", "p", "li"]):
            txt = elem.get_text(strip=True)
            if len(txt) > 25 and not any(skip in txt.lower() for skip in ["subscribe", "sign up", "newsletter", "cookie", "privacy"]):
                paragraphs.append(txt)
                
        content = "\n\n".join(paragraphs)
        
        images = []
        for img in target.find_all("img"):
            src = img.get("src") or img.get("data-src") or img.get("data-original")
            if src and any(ext in src.lower() for ext in [".jpg", ".jpeg", ".png", ".webp"]):
                if not src.startswith("http"):
                    src = requests.compat.urljoin(url, src)
                if src not in images and not any(k in src.lower() for k in ["avatar", "icon", "logo", "banner"]):
                    images.append(src)
                    
        return title_text, content, images[:12]
    except Exception as e:
        return "", "", []

# -------------------------------------------------------------
# 4. KHUNG NHẬP LIỆU GIAO DIỆN CHUẨN APPLE
# -------------------------------------------------------------
with st.container(border=True):
    c1, c2 = st.columns([2, 1], gap="large")

    with c1:
        url_input = st.text_input("🔗 Nhập URL bài viết gốc:", placeholder="https://www.windowscentral.com/...")
        manual_content = st.text_area(
            "📝 Hoặc dán trực tiếp toàn bộ nội dung bài gốc (Khuyên dùng):", 
            height=200,
            placeholder="Mẹo: Mở bài gốc bấm Ctrl + A > Ctrl + C rồi dán vào đây để đủ 100% ý..."
        )

    with c2:
        product_type = st.selectbox(
            "📦 Phân loại sản phẩm:",
            ["Phụ kiện / Thiết bị khác (Chỉ dùng H2, không dùng H3)", "Laptop (Chuẩn H2 + H3)"]
        )
        source_name = st.text_input("🌐 Tên trang web nguồn:", placeholder="Ví dụ: Windows Central, PCMag, NotebookCheck...")
        exchange_rate = st.number_input("💱 Tỷ giá quy đổi (USD/VND):", value=25850, step=50)

    st.markdown("<br>", unsafe_allow_html=True)
    btn_generate = st.button("🚀 Bắt đầu tạo bài viết tự động", use_container_width=True)

# -------------------------------------------------------------
# 5. XỬ LÝ VIẾT BÀI VỚI QUY CHUẨN TỪ GOOGLE DOCS
# -------------------------------------------------------------
if btn_generate:
    if not api_key:
        st.error("⚠️ Vui lòng cung cấp Gemini API Key!")
    else:
        with st.spinner("Đang đọc quy chuẩn từ Google Doc và biên tập bài viết..."):
            extracted_title = ""
            extracted_text = manual_content.strip()
            images_found = []
            
            if url_input:
                extracted_title, web_content, images_found = extract_content_from_url(url_input)
                if not extracted_text:
                    extracted_text = web_content
                    if extracted_text:
                        st.info(f"📊 Đã lấy được khoảng **{len(extracted_text.split())} từ** từ bài gốc.")
                    else:
                        st.warning("⚠️ Không thể tự động cào bài viết do trang web chặn. Vui lòng copy text bài gốc dán vào ô bên trên nhé!")
                        st.stop()
            
            if not extracted_text:
                st.error("Chưa có nội dung bài gốc! Vui lòng nhập URL hoặc dán nội dung chữ.")
                st.stop()
                
            detected_source = source_name if source_name else ("bài gốc" if not url_input else url_input.split("//")[-1].split("/")[0])
            
            # Sử dụng quy chuẩn đọc từ Google Docs (hoặc dự phòng)
            rules_prompt = loaded_rules if loaded_rules else """
QUY TẮC MẪU TRÊN TAY:
1. Title: "Trên tay + [Tên sản phẩm]:... + [giá quy đổi tiền Việt]"
2. Slug: Viết thường, không dấu, cách nhau dấu gạch ngang.
3. Page title: Bắt buộc từ 50 đến 60 ký tự.
4. Page description: Bắt buộc từ 150 đến 160 ký tự.
5. Meta Keywords: Bắt buộc đúng 3 từ khóa, chỉ gồm "trên tay + tên sản phẩm", viết thường toàn bộ, cách nhau bằng dấu phẩy.
6. Sapo: Bắt đầu bằng "Trên tay + [Tên sản phẩm]", đúng 1 câu dài 2 dòng.
7. Đoạn mở bài: 1 đoạn ngắn (3-4 câu) sau Sapo trước H2 đầu tiên.
8. Thân bài:
- Laptop: H2 Thiết kế (H3 Thiết kế 1 ảnh, H3 Bàn phím 2 ảnh); H2 Màn hình; H2 Cấu hình; H2 Tạm kết (2 đoạn có quy đổi giá).
- Phụ kiện: 3 - 5 mục H2 (không dùng H3), mỗi H2 có câu tóm tắt, khung ảnh 1200x675 (Alt text bắt đầu bằng Trên tay + tên sản phẩm), câu dẫn sản phẩm CellphoneS.
9. Dòng cuối cùng: "Nguồn: [Tên nguồn]"
"""

            SYSTEM_PROMPT = f"""
Bạn là Tổng biên tập kiêm Biên tập viên công nghệ cao cấp của Sforum (CellphoneS).
Nhiệm vụ của bạn là áp dụng BỘ QUY CHUẨN ĐƯỢC CUNG CẤP DƯỚI ĐÂY để chuyển thể bài đánh giá nước ngoài thành bài viết tiếng Việt CHUYÊN SÂU (1000 - 1500 từ), ĐẦY ĐỦ TẤT CẢ CÁC Ý, TUYỆT ĐỐI KHÔNG TÓM TẮT SƠ SÀI.

Tỷ giá quy đổi hiện tại: 1 USD = {exchange_rate:,} VNĐ.

BỘ QUY CHUẨN THỰC THI (ĐƯỢC ĐỌC TỪ GOOGLE DOCS):
---
{rules_prompt}
---

LƯU Ý THÊM:
- Phân loại sản phẩm hiện tại: {product_type}
- Nguồn bài viết: {detected_source}
"""

            user_prompt = f"""
TIÊU ĐỀ BÀI GỐC: {extracted_title}
NGUỒN BÀI VIẾT: {detected_source}
LINK BÀI GỐC: {url_input}
DANH SÁCH ẢNH: {images_found}

NỘI DUNG BÀI GỐC:
---
{extracted_text[:25000]}
---
Hãy viết hoàn chỉnh bài viết theo đúng 100% các tiêu chí trong Bộ quy chuẩn Google Docs ở trên.
"""
            try:
                client = genai.Client(api_key=api_key)
                response = client.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=user_prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_PROMPT,
                        temperature=0.3,
                    )
                )
                
                result_text = response.text
                st.balloons()
                st.success("🎉 Bài viết đã được tạo hoàn chỉnh theo đúng quy chuẩn Google Docs!")

                with st.container(border=True):
                    tab_article, tab_seo, tab_raw = st.tabs(["📄 Bài viết hoàn chỉnh", "🎯 Trường SEO CMS", "📋 Mã Markdown"])
                    
                    with tab_article:
                        st.markdown(result_text)
                        st.download_button(
                            "📥 Tải bài viết (.txt)", 
                            data=result_text, 
                            file_name="bai-viet-sforum.txt", 
                            mime="text/plain"
                        )
                        
                    with tab_seo:
                        st.subheader("📋 Sao chép nhanh vào CMS:")
                        title_match = re.search(r"Title:(.*)", result_text, re.IGNORECASE)
                        slug_match = re.search(r"Slug:(.*)", result_text, re.IGNORECASE)
                        page_title_match = re.search(r"Page title:(.*)", result_text, re.IGNORECASE)
                        page_desc_match = re.search(r"Page description:(.*)", result_text, re.IGNORECASE)
                        keywords_match = re.search(r"Meta Keywords:(.*)", result_text, re.IGNORECASE)
                        tags_match = re.search(r"Tags:(.*)", result_text, re.IGNORECASE)
                        
                        st.text_input("Tiêu đề (Title):", value=title_match.group(1).strip() if title_match else "")
                        st.text_input("Đường dẫn (Slug):", value=slug_match.group(1).strip() if slug_match else "")
                        st.text_input("Page Title (< 60 ký tự):", value=page_title_match.group(1).strip() if page_title_match else "")
                        st.text_area("Page Description (< 160 ký tự):", value=page_desc_match.group(1).strip() if page_desc_match else "", height=80)
                        st.text_input("Meta Keywords (3 từ cách nhau dấu phẩy):", value=keywords_match.group(1).strip() if keywords_match else "")
                        st.text_input("Tags:", value=tags_match.group(1).strip() if tags_match else "")
                        
                    with tab_raw:
                        st.code(result_text, language="markdown")
                        
            except Exception as e:
                st.error(f"Lỗi: {e}")
