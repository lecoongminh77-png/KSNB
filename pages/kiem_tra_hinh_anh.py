import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io

st.set_page_config(page_title="Tạo Tem Tự Động", page_icon="🏷️", layout="centered")

st.title("🏷️ Tool In Tem Danh Sách Tự Động")
st.write("Tải file Excel lên để hệ thống tự động tạo danh sách tem in ấn.")

uploaded_file = st.file_uploader("Chọn file Excel (.xlsx hoặc .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("Tải file thành công! Dưới đây là dữ liệu xem trước:")
        st.dataframe(df.head(10))

        # Cho phép chọn cột chứa tên
        column_names = list(df.columns)
        selected_col = st.selectbox("Chọn cột chứa họ và tên (hoặc nội dung tem):", column_names)

        # Cấu hình tem
        cols_per_row = st.selectbox("Số lượng tem trên một hàng:", [2, 3], index=0)

        if st.button("🚀 Bắt đầu tạo tem"):
            names = df[selected_col].dropna().astype(str).tolist()

            if not names:
                st.warning("Cột được chọn không có dữ liệu!")
            else:
                # Khởi tạo file Word
                doc = Document()

                # Căn lề trang 1.5cm (~0.6 inch) để tối ưu diện tích in
                sections = doc.sections
                for section in sections:
                    section.top_margin = Inches(0.5)
                    section.bottom_margin = Inches(0.5)
                    section.left_margin = Inches(0.5)
                    section.right_margin = Inches(0.5)

                # Tính số hàng cần thiết
                num_rows = (len(names) + cols_per_row - 1) // cols_per_row
                table = doc.add_table(rows=num_rows, cols=cols_per_row)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER

                # Đổ dữ liệu vào từng ô tem
                for i, name in enumerate(names):
                    r = i // cols_per_row
                    c = i % cols_per_row
                    cell = table.cell(r, c)
                    
                    # Căn giữa theo chiều dọc
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                    # Viền khung tem (tạo viền nét đứt hoặc viền mỏng dễ cắt)
                    tcPr = cell._tc.get_or_add_tcPr()
                    borders = parse_xml(
                        f'<w:tcBorders {nsdecls("w")}>\n'
                        f'  <w:top w:val="dashed" w:sz="6" w:space="0" w:color="A0A0A0"/>\n'
                        f'  <w:left w:val="dashed" w:sz="6" w:space="0" w:color="A0A0A0"/>\n'
                        f'  <w:bottom w:val="dashed" w:sz="6" w:space="0" w:color="A0A0A0"/>\n'
                        f'  <w:right w:val="dashed" w:sz="6" w:space="0" w:color="A0A0A0"/>\n'
                        f'</w:tcBorders>'
                    )
                    tcPr.append(borders)

                    # Nội dung tem
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # Dòng phụ (tiêu đề tem nếu cần)
                    p_sub = p.add_run("THẺ NHÂN VIÊN / TEM TÊN\n")
                    p_sub.font.size = Pt(9)
                    p_sub.font.color.rgb = RGBColor(120, 120, 120)

                    # Dòng tên chính
                    p_name = p.add_run(name.strip())
                    p_name.font.size = Pt(14)
                    p_name.font.bold = True

                # Lưu vào bộ nhớ đệm (buffer) để tải về
                doc_buffer = io.BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)

                st.success(f"Đã tạo thành công {len(names)} tem!")
                st.download_button(
                    label="📥 Tải file tem Word (.docx) về máy",
                    data=doc_buffer,
                    file_name="danh_sach_tem.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

    except Exception as e:
        st.error(f"Lỗi khi đọc file: {e}")
