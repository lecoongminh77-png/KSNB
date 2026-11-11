import sys
import os
import io
import time
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN (CHUẨN APPLE STYLE)
# -------------------------------------------------------------
st.set_page_config(page_title='Tạo Tem Tự Động', page_icon='🏷️', layout='wide')

st.markdown("""
<style>
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    .stApp { background-color: #f5f5f7; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; color: #1d1d1f; }
    .apple-header-box { display: flex; justify-content: space-between; align-items: flex-end; padding: 20px 0 10px 0; margin-bottom: 5px; }
    .apple-main-title { font-size: 38px; font-weight: 700; color: #1d1d1f; letter-spacing: -0.025em; margin: 0; }
    div.stButton > button:first-child { background: #0071e3 !important; color: #ffffff !important; border-radius: 980px !important; border: none !important; padding: 10px 24px !important; font-weight: 600 !important; }
    div.stButton > button:first-child:hover { background: #0077ed !important; }
    div.stDownloadButton > button:first-child { background: #28a745 !important; color: #ffffff !important; border-radius: 980px !important; border: none !important; padding: 10px 24px !important; font-weight: 600 !important; }
    div.stDownloadButton > button:first-child:hover { background: #218838 !important; }
    [data-testid="stVerticalBlockBorderWrapper"] { background: #ffffff; border-radius: 16px; padding: 20px; border: 1px solid rgba(0, 0, 0, 0.08); box-shadow: 0 2px 12px rgba(0, 0, 0, 0.03); }
    div[data-testid="stMetric"] { background-color: #ffffff; border-radius: 12px; padding: 15px; border: 1px solid rgba(0,0,0,0.05); box-shadow: 0 2px 8px rgba(0,0,0,0.02); text-align: center; }
    .upload-title { font-size: 18px; font-weight: 700; color: #1d1d1f; margin-bottom: 12px; }
    .guide-box { background: #fbfbfd; border-radius: 12px; padding: 16px; margin-bottom: 16px; border: 1px solid rgba(0,0,0,0.04); }
    .guide-header { font-weight: 700; color: #0071e3; font-size: 14px; margin-bottom: 6px; }
</style>
""", unsafe_allow_html=True)

# Nút quay về trang chủ
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button('🏠 Trang chủ', use_container_width=True):
        st.switch_page('app_ksnb.py')

st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>", unsafe_allow_html=True)
st.markdown("<div class='apple-header-box'><div><h1 class='apple-main-title'>🏷️ Tạo Tem Tự Động Từ Excel</h1></div></div>", unsafe_allow_html=True)
st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 15px 0 25px 0;'>", unsafe_allow_html=True)

# -------------------------------------------------------------
# HÀM TẠO FILE WORD TEM NHÃN
# -------------------------------------------------------------
def tao_file_word_tem(danh_sach_tem, tieu_de_tem="TEM THÔNG TIN", so_cot=2):
    doc = Document()
    
    # Căn lề khổ A4: 1.27 cm (~0.5 inch)
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    num_rows = (len(danh_sach_tem) + so_cot - 1) // so_cot
    table = doc.add_table(rows=num_rows, cols=so_cot)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, content in enumerate(danh_sach_tem):
        r = i // so_cot
        c = i % so_cot
        cell = table.cell(r, c)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # Viền tem nét đứt màu xám nhạt (chuẩn đường cắt)
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="dashed" w:sz="6" w:space="0" w:color="B0B0B0"/>\n'
            f'  <w:left w:val="dashed" w:sz="6" w:space="0" w:color="B0B0B0"/>\n'
            f'  <w:bottom w:val="dashed" w:sz="6" w:space="0" w:color="B0B0B0"/>\n'
            f'  <w:right w:val="dashed" w:sz="6" w:space="0" w:color="B0B0B0"/>\n'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)

        # Viết nội dung tem
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Dòng tiêu đề nhỏ bên trên
        if tieu_de_tem.strip():
            p_hdr = p.add_run(f"{tieu_de_tem.strip().upper()}\n")
            p_hdr.font.name = 'Arial'
            p_hdr.font.size = Pt(8.5)
            p_hdr.font.color.rgb = RGBColor(120, 120, 120)

        # Dòng thông tin chính
        p_main = p.add_run(f"{str(content).strip()}\n")
        p_main.font.name = 'Arial'
        p_main.font.size = Pt(13)
        p_main.font.bold = True
        p_main.font.color.rgb = RGBColor(29, 29, 31)

    doc_stream = io.BytesIO()
    doc.save(doc_stream)
    doc_stream.seek(0)
    return doc_stream

# -------------------------------------------------------------
# LOGIC CHÍNH
# -------------------------------------------------------------
def trang_tao_tem():
    st.caption('Công cụ hỗ trợ • Chuyển đổi dữ liệu bảng tính thành nhãn in tức thì')
    st.markdown("### 📥 Nạp Dữ Liệu Nguồn")

    col_upload, col_config = st.columns([6, 4], gap="large")

    with col_upload:
        with st.container(border=True):
            st.markdown("<div class='upload-title'>1️⃣ Tải Lên Bảng Tính Excel</div><div class='guide-box'><div class='guide-header'>📌 Định dạng file hỗ trợ</div>Hệ thống hỗ trợ file <b>.xlsx</b> hoặc <b>.xls</b> có chứa cột thông tin cần in tem.</div>", unsafe_allow_html=True)
            f_excel = st.file_uploader('Chọn file Excel', type=['xlsx', 'xls'], label_visibility="collapsed")

    df = None
    if f_excel is not None:
        try:
            xl = pd.ExcelFile(f_excel)
            sheet_selected = xl.sheet_names[0]
            if len(xl.sheet_names) > 1:
                sheet_selected = st.selectbox("Chọn Sheet cần lấy dữ liệu:", xl.sheet_names)
            
            df = pd.read_excel(f_excel, sheet_name=sheet_selected)
        except Exception as e:
            st.error(f"Lỗi khi đọc file: {e}")

    with col_config:
        with st.container(border=True):
            st.markdown("<div class='upload-title'>2️⃣ Cấu Hình Khổ Tem</div>", unsafe_allow_html=True)
            if df is not None:
                cot_du_lieu = st.selectbox("Chọn cột lấy nội dung in tem:", list(df.columns))
                tieu_de_tem = st.text_input("Tiêu đề nhỏ trên tem (tùy chọn):", value="TEM ĐỊNH DANH")
                so_cot_in = st.radio("Bố cục số tem mỗi dòng:", [2, 3, 4], index=0, horizontal=True)
            else:
                st.info("💡 Vui lòng nạp file Excel trước để thiết lập các thông số.")
                cot_du_lieu = None
                tieu_de_tem = ""
                so_cot_in = 2

    # Hiển thị dữ liệu xem trước
    if df is not None:
        st.markdown("<br>", unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown("### 👁️ Xem Trước Dữ Liệu Nguồn (10 dòng đầu)")
            st.dataframe(df.head(10), use_container_width=True)

        if st.button('🚀 Bắt đầu tạo tem tự động', use_container_width=True):
            progress_bar = st.progress(0, text="Đang xử lý dữ liệu...")
            
            danh_sach = df[cot_du_lieu].dropna().astype(str).tolist()
            danh_sach = [x.strip() for x in danh_sach if x.strip() != ""]

            progress_bar.progress(50, text="Đang dựng khung in và bố cục văn bản...")
            time.sleep(0.3)

            if len(danh_sach) == 0:
                progress_bar.empty()
                st.warning("⚠️ Cột đã chọn không có dữ liệu hợp lệ để tạo tem!")
                return

            file_word = tao_file_word_tem(danh_sach, tieu_de_tem, so_cot=so_cot_in)
            
            progress_bar.progress(100, text="✅ Hoàn tất!")
            time.sleep(0.5)
            progress_bar.empty()

            st.markdown("### 📊 DASHBOARD: KẾT QUẢ")
            m1, m2, m3 = st.columns(3)
            m1.metric("📦 Tổng Dòng Dữ Liệu", f"{len(df)} dòng")
            m2.metric("🏷️ Số Tem Đã Tạo", f"{len(danh_sach)} tem")
            m3.metric("📑 Bố Cục Trang In", f"{so_cot_in} tem / hàng")

            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="📥 Tải File Word In Tem (.docx)",
                data=file_word,
                file_name="Danh_Sach_Tem_In.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

trang_tao_tem()
