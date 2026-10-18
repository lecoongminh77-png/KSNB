import streamlit as st
import re
import requests
from bs4 import BeautifulSoup
from google import genai
from google.genai import types

# -------------------------------------------------------------
# CẤU HÌNH TRANG STREAMLIT
# -------------------------------------------------------------
st.set_page_config(
    page_title="Sforum Auto Content Creator",
    page_icon="✍️",
    layout="wide"
)

st.title("✍️ Sforum Content Generator - Chuẩn Form Trên Tay")
st.caption("Tự động chuyển đổi bài review nước ngoài thành bài viết Trên tay chuẩn SEO CellphoneS / Sforum.")

# -------------------------------------------------------------
# LẤY API KEY TỪ STREAMLIT SECRETS HOẶC INPUT
# -------------------------------------------------------------
api_key = None
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    api_key = st.sidebar.text_input("🔑 Nhập Gemini API Key:", type="password")

if not api_key:
    st.info("💡 Mẹo: Để không phải nhập API Key mỗi lần vào web, hãy thêm `GEMINI_API_KEY = 'khoá_của_bạn'` vào mục **Settings > Secrets** trên Streamlit Cloud.")

# -------------------------------------------------------------
# HÀM BÓC TÁCH NỘI DUNG TỪ ĐƯỜNG DẪN URL
# -------------------------------------------------------------
def extract_content_from_url(url: str):
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        
        # Bỏ các thẻ thừa
        for s in soup(["script", "style", "nav", "footer", "header", "noscript"]):
            s.decompose()
            
        title = soup.find("h1")
        title_text = title.get_text(strip=True) if title else ""
        
        # Lấy các đoạn văn chính
        paragraphs = [p.get_text(strip=True) for p in soup.find_all("p") if len(p.get_text(strip=True)) > 40]
        content = "\n\n".join(paragraphs)
        
        # Lấy danh sách link ảnh gốc
        images = []
        for img in soup.find_all("img"):
            src = img.get("src") or img.get("data-src")
            if src and any(ext in src.lower() for ext in [".jpg", ".jpeg", ".png", ".webp"]):
                if not src.startswith("http"):
                    src = requests.compat.urljoin(url, src)
                if src not in images and "avatar" not in src.lower() and "logo" not in src.lower():
                    images.append(src)
                    
        return title_text, content, images[:10]
    except Exception as e:
        st.warning(f"Không thể tự động tải URL: {e}. Vui lòng dán trực tiếp văn bản bên dưới.")
        return "", "", []

# -------------------------------------------------------------
# GIAO DIỆN NHẬP LIỆU
# -------------------------------------------------------------
col1, col2 = st.columns([2, 1])

with col1:
    url_input = st.text_input("🔗 Nhập URL bài viết gốc (Windows Central, PCMag, NotebookCheck...):")
    manual_content = st.text_area("📝 Hoặc dán trực tiếp nội dung bài gốc (nếu có):", height=200)

with col2:
    product_type = st.selectbox(
        "📦 Loại sản phẩm:",
        ["Phụ kiện / Thiết bị khác (Chỉ dùng H2)", "Laptop (Chuẩn H2 + H3)"]
    )
    source_name = st.text_input("🌐 Tên trang web nguồn:", placeholder="Ví dụ: Windows Central, PCMag...")
    exchange_rate = st.number_input("💱 Tỷ giá USD/VND:", value=25850, step=50)

# -------------------------------------------------------------
# BỘ PROMPT CHUẨN ĐƯỢC THIẾT KẾ RIÊNG
# -------------------------------------------------------------
SYSTEM_PROMPT = f"""
Bạn là Biên tập viên công nghệ hàng đầu của trang tin Sforum (CellphoneS). 
Nhiệm vụ của bạn là chuyển thể nội dung bài viết đánh giá từ nước ngoài thành bài viết "Trên tay" tiếng Việt chuẩn 100% quy chuẩn sau:

TỶ GIÁ QUY ĐỔI HIỆN TẠI: 1 USD = {exchange_rate:,} VNĐ.

QUY TẮC BẮT BUỘC THEO FORM:
1. Title: Bắt buộc theo cấu trúc: "Trên tay + [Tên sản phẩm]:... + [giá tiền Việt ví dụ: giá 1.55 triệu đồng]"
2. Slug: Ngắn gọn, viết thường, không dấu, phân cách bằng dấu gạch ngang (VD: tren-tay-razer-tartarus-v2-pro).
3. Page title: Bắt buộc từ 50 đến 60 ký tự (Đếm chính xác ký tự).
4. Page description: Bắt buộc từ 150 đến 160 ký tự (Đếm chính xác ký tự).
5. Meta Keywords: Bắt buộc đúng 3 từ khóa, chỉ gồm "trên tay + tên sản phẩm" (hoặc tên viết tắt), VIẾT THƯỜNG TOÀN BỘ, cách nhau bằng dấu phẩy. TUYỆT ĐỐI không thêm từ ngữ nào khác.
6. Tags: Các từ khóa thương hiệu, dòng sản phẩm, loại phụ kiện liên quan.
7. Sapo: Bắt buộc bắt đầu bằng cụm từ "Trên tay + [Tên sản phẩm]", đúng 1 câu duy nhất và độ dài tương đương 2 dòng hiển thị.
8. Đoạn mở bài: Luôn có 1 đoạn văn ngắn (2-3 câu) đặt ngay sau Sapo và trước tiêu đề H2 đầu tiên.

CẤU TRÚC THÂN BÀI:

A. NẾU LÀ LAPTOP:
- <H2> Thiết kế và ngoại hình + [Tên sản phẩm]
  (1 câu tóm tắt ngay sau H2)
  - <H3> Thiết kế (đặt lại tên hay hơn, lấy thông tin bài gốc, viết 2 đoạn. Chèn khung [Ảnh minh họa - 1200x675] ở giữa kèm Alt text và Caption).
  - <H3> Bàn phím, touchpad, cổng kết nối (đặt lại tên hay hơn, 2 đoạn, chèn 2 khung [Ảnh minh họa - 1200x675] kèm Alt text và Caption).
- <H2> Màn hình hiển thị + [Tên sản phẩm]
  (1 câu tóm tắt ngay sau H2)
  - <H3> Màn hình (đặt lại tên hay hơn)
  - <H3> Webcam, loa (đặt lại tên hay hơn)
- <H2> Cấu hình và hiệu năng + [Tên sản phẩm]
  (1 câu tóm tắt ngay sau H2)
  - <H3> Cấu hình kèm khả năng nâng cấp
  - <H3> Hiệu năng
- <H2> Tạm kết (Xem quy tắc Tạm kết bên dưới).

B. NẾU LÀ PHỤ KIỆN / THIẾT BỊ KHÁC (KHÔNG PHẢI LAPTOP):
- Tiêu đề H2 phân bổ hợp lý theo bài gốc (ngắn gọn, dưới 60 ký tự). TUYỆT ĐỐI KHÔNG CÓ THẺ H3.
- Sau mỗi H2 bắt buộc có 1 câu tóm tắt ngắn.
- Các vị trí chèn ảnh phải ghi rõ: [Ảnh minh họa: Tên ảnh - Kích thước: 1200 x 675 px], thẻ Alt text (BẮT BUỘC mở đầu bằng "Trên tay + [Tên sản phẩm]...") và Caption rõ ràng.
- Chèn 1 câu dẫn kêu gọi hành động (CTA) tự nhiên dưới 50 từ để giới thiệu sản phẩm tương tự đang bán tại CellphoneS kèm vị trí [Chèn box sản phẩm CellphoneS].

QUY TẮC PHẦN TẠM KẾT:
- <H2> Tạm kết
- Viết đúng 2 đoạn văn:
  + Đoạn 1: Tóm tắt lại cảm nhận khi cầm trên tay thiết bị và các đánh giá nhanh về ưu/nhược điểm.
  + Đoạn 2: Ghi rõ giá bán ngoại tệ gốc của bài viết VÀ quy đổi ra tiền Việt Nam Đồng theo tỷ giá hiện tại (nếu có bản tiền nhiệm thì quy đổi giá của bản tiền nhiệm luôn).
- Dòng cuối cùng: "Nguồn: [Tên trang web]"
"""

# -------------------------------------------------------------
# NÚT XỬ LÝ
# -------------------------------------------------------------
if st.button("🚀 Bắt đầu tạo bài viết tự động", type="primary", use_container_width=True):
    if not api_key:
        st.error("Vui lòng cung cấp Gemini API Key để tiếp tục!")
    else:
        with st.spinner("Đang cào dữ liệu và phân tích bài viết..."):
            extracted_title = ""
            extracted_text = manual_content
            images_found = []
            
            if url_input:
                extracted_title, web_content, images_found = extract_content_from_url(url_input)
                if not extracted_text:
                    extracted_text = web_content
                    
            if not extracted_text:
                st.error("Không tìm thấy nội dung bài viết! Vui lòng nhập link hoặc dán nội dung bài gốc.")
                st.stop()
                
            detected_source = source_name if source_name else ("bài gốc" if not url_input else url_input.split("//")[-1].split("/")[0])
            
            user_prompt = f"""
LOẠI BÀI VIẾT: {product_type}
NGUỒN BÀI VIẾT: {detected_source}
LINK BÀI GỐC: {url_input}
DANH SÁCH ẢNH TÌM ĐƯỢC: {images_found}

Dưới đây là nội dung bài gốc cần chuyển đổi:
---
Tiêu đề gốc: {extracted_title}
Nội dung:
{extracted_text[:12000]}
---
Hãy viết hoàn chỉnh bài viết theo đúng 100% các tiêu chí trong System Prompt.
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
                
                st.success("🎉 Bài viết đã được tạo thành công!")
                
                # Chia tab hiển thị
                tab_article, tab_seo, tab_raw = st.tabs(["📄 Bài viết hoàn chỉnh", "🎯 Trường SEO CMS", "📋 Mã Markdown"])
                
                with tab_article:
                    st.markdown(result_text)
                    st.download_button(
                        label="📥 Tải bài viết (.txt)",
                        data=result_text,
                        file_name="bai-viet-sforum.txt",
                        mime="text/plain"
                    )
                    
                with tab_seo:
                    st.subheader("Copy nhanh vào hệ thống CMS:")
                    # Tìm trích xuất các trường SEO bằng regex
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
                st.error(f"Lỗi khi tạo bài viết: {e}")
