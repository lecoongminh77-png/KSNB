import streamlit as st
import re
import requests
from bs4 import BeautifulSoup
from google import genai
from google.genai import types

# -------------------------------------------------------------
# 1. CẤU HÌNH TRANG GIAO DIỆN HIỆN ĐẠI & TINH TẾ
# -------------------------------------------------------------
st.set_page_config(
    page_title="Sforum Content Studio",
    page_icon="📝",
    layout="wide"
)

# Tùy biến nhẹ nhàng chuẩn màu CellphoneS (không lạm dụng CSS)
st.markdown("""
<style>
    /* Nút bấm đỏ CellphoneS thanh lịch */
    div.stButton > button:first-child {
        background-color: #d70018 !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 10px 24px !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
    }
    div.stButton > button:first-child:hover {
        background-color: #b80015 !important;
    }
    /* Chỉnh tiêu đề chính */
    .title-red {
        color: #d70018;
        font-weight: 800;
        margin-bottom: 0px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="title-red">📝 Sforum Content Studio</h1>', unsafe_allow_html=True)
st.caption("Công cụ tự động chuyển thể bài viết review quốc tế thành bài Trên tay / Đánh giá chuẩn SEO Sforum.")

# -------------------------------------------------------------
# 2. LẤY API KEY TỪ SECRETS HOẶC SIDEBAR
# -------------------------------------------------------------
api_key = None
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    api_key = st.sidebar.text_input("🔑 Gemini API Key:", type="password")

# -------------------------------------------------------------
# 3. HÀM CÀO DỮ LIỆU BÀI GỐC
# -------------------------------------------------------------
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
# 4. GIAO DIỆN NHẬP DỮ LIỆU (KHUNG VIỀN HIỆN ĐẠI)
# -------------------------------------------------------------
with st.container(border=True):
    col1, col2 = st.columns([2, 1], gap="medium")

    with col1:
        url_input = st.text_input("🔗 Nhập URL bài viết gốc:", placeholder="https://www.windowscentral.com/...")
        manual_content = st.text_area(
            "📝 Hoặc dán trực tiếp toàn bộ chữ bài gốc (Khuyên dùng để đủ 100% ý):", 
            height=200,
            placeholder="Nếu trang web chặn cào tự động, mở bài gốc bấm Ctrl + A > Ctrl + C rồi dán vào đây..."
        )

    with col2:
        product_type = st.selectbox(
            "📦 Phân loại sản phẩm:",
            ["Phụ kiện / Thiết bị khác (Chỉ dùng H2, không dùng H3)", "Laptop (Chuẩn H2 + H3)"]
        )
        source_name = st.text_input("🌐 Tên trang web nguồn:", placeholder="Ví dụ: Windows Central, PCMag, NotebookCheck...")
        exchange_rate = st.number_input("💱 Tỷ giá quy đổi (USD/VND):", value=25850, step=50)

    btn_generate = st.button("🚀 Bắt đầu tạo bài viết tự động", use_container_width=True)

# -------------------------------------------------------------
# 5. BỘ XỬ LÝ NỘI DUNG
# -------------------------------------------------------------
if btn_generate:
    if not api_key:
        st.error("⚠️ Vui lòng cung cấp Gemini API Key!")
    else:
        with st.spinner("Đang đọc dữ liệu bài gốc và biên tập bài viết..."):
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
            
            SYSTEM_PROMPT = f"""
Bạn là Tổng biên tập kiêm Biên tập viên công nghệ cao cấp của Sforum (CellphoneS). 
Nhiệm vụ của bạn là chuyển thể toàn bộ bài viết đánh giá nước ngoài thành một bài viết "Trên tay" tiếng Việt CHUYÊN SÂU, ĐẦY ĐỦ TẤT CẢ CÁC Ý, KHÔNG ĐƯỢC TÓM TẮT RÚT GỌN.

YÊU CẦU QUAN TRỌNG VỀ ĐỘ DÀI VÀ NỘI DUNG:
- Độ dài bài viết: Phải đạt từ 1000 - 1500 từ.
- TUYỆT ĐỐI KHÔNG BỎ SÓT Ý: Phải chuyển tải đầy đủ tất cả các phân tích, thông số, chi tiết thiết kế, trải nghiệm thực tế, ưu điểm, nhược điểm và nhận xét của tác giả bài gốc.
- Mỗi mục H2 phải viết từ 2 đến 3 đoạn văn dài, có chiều sâu, văn phong gần gũi công nghệ.
- Tỷ giá quy đổi: 1 USD = {exchange_rate:,} VNĐ.

QUY TẮC CẤU TRÚC BẮT BUỘC:
1. Title: "Trên tay + [Tên sản phẩm]:... + [giá quy đổi tiền Việt]"
2. Slug: Viết thường, không dấu, cách nhau dấu gạch ngang.
3. Page title: Bắt buộc từ 50 đến 60 ký tự.
4. Page description: Bắt buộc từ 150 đến 160 ký tự.
5. Meta Keywords: Bắt buộc đúng 3 từ khóa, chỉ gồm "trên tay + tên sản phẩm" (hoặc tên viết tắt), VIẾT THƯỜNG TOÀN BỘ, cách nhau bằng dấu phẩy.
6. Tags: Các từ khóa liên quan đến hãng, dòng sản phẩm, loại phụ kiện.
7. Sapo: Bắt buộc bắt đầu bằng cụm từ "Trên tay + [Tên sản phẩm]", đúng 1 câu duy nhất dài 2 dòng hiển thị.
8. Đoạn mở bài: 1 đoạn văn ngắn (3-4 câu) đặt ngay sau Sapo và trước tiêu đề H2 đầu tiên.

QUY TẮC THÂN BÀI:
A. NẾU LÀ LAPTOP:
- <H2> Thiết kế và ngoại hình + [Tên sản phẩm] (1 câu tóm tắt sau H2)
  - <H3> Thiết kế (đặt lại tên hay hơn, viết 2 đoạn, chèn 1 khung ảnh ở giữa kèm Alt text và Caption)
  - <H3> Bàn phím, touchpad, cổng kết nối (đặt lại tên hay hơn, viết 2 đoạn, chèn 2 khung ảnh ở giữa kèm Alt text và Caption)
- <H2> Màn hình hiển thị + [Tên sản phẩm] (1 câu tóm tắt sau H2)
  - <H3> Màn hình
  - <H3> Webcam, loa
- <H2> Cấu hình và hiệu năng + [Tên sản phẩm] (1 câu tóm tắt sau H2)
  - <H3> Cấu hình kèm khả năng nâng cấp
  - <H3> Hiệu năng
- <H2> Tạm kết (Xem quy tắc Tạm kết)

B. NẾU LÀ PHỤ KIỆN / THIẾT BỊ KHÁC:
- Dựa vào toàn bộ nội dung bài gốc, chia thành 3 - 5 tiêu đề H2 hợp lý để phản ánh ĐẦY ĐỦ từng khía cạnh: Thiết kế & công thái học, Switch/phím bấm/công nghệ lõi, Tính năng phụ/kết nối/phần mềm, Trợ năng/trải nghiệm thực tế.
- TUYỆT ĐỐI KHÔNG DÙNG THẺ H3 CHO PHỤ KIỆN.
- Sau mỗi H2 bắt buộc có 1 câu tóm tắt ngắn trước khi vào bài chi tiết.
- Mỗi mục H2 đều chèn khung [Ảnh minh họa - 1200 x 675 px], thẻ Alt text (bắt buộc bắt đầu bằng "Trên tay + [Tên sản phẩm]...") và Caption rõ nghĩa.
- Sau phần trải nghiệm/switch, chèn 1 câu dẫn kêu gọi hành động dưới 50 từ để giới thiệu sản phẩm tương tự đang bán tại CellphoneS kèm vị trí [Chèn box sản phẩm CellphoneS].

QUY TẮC PHẦN TẠM KẾT:
- <H2> Tạm kết
- Viết đúng 2 đoạn văn:
  + Đoạn 1: Tóm tắt cảm nhận cầm trên tay và đánh giá nhanh ưu/nhược điểm.
  + Đoạn 2: Ghi rõ giá bán ngoại tệ gốc VÀ quy đổi ra tiền Việt Nam theo tỷ giá hiện tại (nếu bài gốc nhắc tới bản tiền nhiệm thì quy đổi giá của bản tiền nhiệm luôn).
- Dòng cuối cùng: "Nguồn: {detected_source}"
"""

            user_prompt = f"""
LOẠI BÀI VIẾT: {product_type}
TIÊU ĐỀ BÀI GỐC: {extracted_title}
NGUỒN BÀI VIẾT: {detected_source}
LINK BÀI GỐC: {url_input}
DANH SÁCH ẢNH TRÍCH XUẤT ĐƯỢC: {images_found}

NỘI DUNG BÀI GỐC:
---
{extracted_text[:25000]}
---

YÊU CẦU: Hãy đọc kỹ toàn bộ nội dung bài gốc trên, viết một bài viết hoàn chỉnh, sâu sắc, dài trên 1000 từ, bảo đảm đầy đủ TẤT CẢ các ý của bài gốc theo đúng 100% các tiêu chí trong System Prompt.
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
                st.success("🎉 Bài viết đã được tạo hoàn chỉnh!")
                
                tab_article, tab_seo, tab_raw = st.tabs(["📄 Bài viết hoàn chỉnh", "🎯 Trường SEO CMS", "📋 Mã Markdown"])
                
                with tab_article:
                    with st.container(border=True):
                        st.markdown(result_text)
                    st.download_button(
                        "📥 Tải bài viết (.txt)", 
                        data=result_text, 
                        file_name="bai-viet-sforum.txt", 
                        mime="text/plain"
                    )
                    
                with tab_seo:
                    with st.container(border=True):
                        st.subheader("📋 Copy nhanh vào CMS:")
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
