import streamlit as st
import pandas as pd
from PIL import Image
import io
import time
import gc
import base64
from openai import OpenAI
import fitz  # PyMuPDF để xử lý file PDF scan
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN
# -------------------------------------------------------------
st.set_page_config(page_title='Tổng Hợp & Đối Chiếu Biên Bản Kho', page_icon='📦', layout='wide')

st.markdown("""
<style>
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    .stApp { background-color: #f5f5f7; font-family: -apple-system, sans-serif; color: #1d1d1f; }
    .apple-main-title { font-size: 34px; font-weight: 700; color: #1d1d1f; margin-bottom: 20px; text-align: center; }
    [data-testid="stVerticalBlockBorderWrapper"] { background: #ffffff; border-radius: 16px; padding: 15px; border: 1px solid rgba(0,0,0,0.08); box-shadow: 0 2px 12px rgba(0,0,0,0.03); }
    div[data-testid="stMetric"] { background-color: #ffffff; border-radius: 12px; padding: 15px; border: 1px solid rgba(0,0,0,0.05); box-shadow: 0 2px 8px rgba(0,0,0,0.02); text-align: center; }
    .upload-title { font-size: 18px; font-weight: 700; margin-bottom: 10px; color: #8b5cf6;}
    .list-box { background: #fff; padding: 15px; border-radius: 10px; border: 1px solid #e2e8f0; height: 100%; }
    button[kind="primary"] { background-color: #8b5cf6 !important; color: white !important; border-color: #8b5cf6 !important; }
</style>
""", unsafe_allow_html=True)

# Nút về trang chủ
col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button('🏠 Trang chủ', type='primary', use_container_width=True):
        try: st.switch_page('app_ksnb.py')
        except: pass

st.markdown("<h1 class='apple-main-title'>📦 Tổng Hợp & Đối Chiếu Biên Bản Kho Tôn Hoa Sen</h1>", unsafe_allow_html=True)
st.markdown("<hr style='border-top: 1px solid #e2e8f0; margin-bottom: 20px;'>", unsafe_allow_html=True)

# Hàm chuyển ảnh sang Base64 gửi qua OpenRouter Vision
# Hàm chuyển ảnh sang Base64 gửi qua OpenRouter Vision
def encode_image_to_base64(image_bytes):
    img = Image.open(io.BytesIO(image_bytes))
    img.thumbnail((1600, 1600))
    buffered = io.BytesIO()
    img.save(buffered, format="JPEG", quality=90)
    return base64.b64encode(buffered.getvalue()).decode('utf-8')

# -------------------------------------------------------------
# CẤU HÌNH OPENROUTER API
# -------------------------------------------------------------
st.markdown("### 🔑 Cấu hình OpenRouter API (AI Vision)")
api_key_input = st.text_input("Dán OpenRouter API Key vào đây:", type="password")

selected_model = st.selectbox(
    "🤖 Chọn Model AI Quét Biên Bản:",
    [
        "google/gemini-2.0-flash-exp:free",
        "meta-llama/llama-3.2-90b-vision-instruct:free",
        "qwen/qwen-2-vl-7b-instruct:free"
    ]
)

client = None
if api_key_input:
    try:
        client = OpenAI(
            base_url="https://openrouter.ai/api/v1",
            api_key=api_key_input.strip()
        )
        st.success("✅ Đã kết nối thành công với OpenRouter!")
    except Exception as e:
        st.error(f"❌ Lỗi khởi tạo Key: {e}")

st.markdown("<hr style='border-top: 1px dashed #ccc; margin: 20px 0;'>", unsafe_allow_html=True)

st.markdown("### 📥 Tải Lên Các Biên Bản Đối Chiếu (PDF hoặc Hình Ảnh)")
uploaded_files = st.file_uploader(
    "Tải lên các file biên bản (Hỗ trợ file .PDF, .JPG, .PNG)", 
    type=['pdf', 'jpg', 'jpeg', 'png'], 
    accept_multiple_files=True
)

# -------------------------------------------------------------
# XỬ LÝ CHÍNH
# -------------------------------------------------------------
if not uploaded_files:
    st.info("👆 Vui lòng tải lên các file biên bản đối chiếu hàng gửi kho để hệ thống tổng hợp.")
elif not api_key_input or not client:
    st.warning("⚠️ Vui lòng nhập OpenRouter API Key ở trên để AI đọc dữ liệu.")
else:
    if st.button("🚀 BẮT ĐẦU TỔNG HỢP & ĐỐI CHIẾU", use_container_width=True):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            extracted_docs = []
            total_files = len(uploaded_files)
            
            for idx, file in enumerate(uploaded_files):
                pct = int(((idx) / total_files) * 70)
                status_text.markdown(f"**⏳ Đang xử lý file {idx+1}/{total_files}: `{file.name}` ({pct}%)**")
                progress_bar.progress(pct)
                
                file_bytes = file.getvalue()
                images_to_process = []
                
                if file.type == "application/pdf" or file.name.lower().endswith('.pdf'):
                    # Chuyển PDF thành ảnh bằng PyMuPDF
                    doc = fitz.open(stream=file_bytes, filetype="pdf")
                    for page_num in range(len(doc)):
                        page = doc.load_page(page_num)
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2)) # Zoom x2 nét chữ
                        images_to_process.append(pix.tobytes("png"))
                else:
                    images_to_process.append(file_bytes)
                
                # Gọi AI Vision đọc thông tin khách hàng và số tổng
                for img_idx, img_bytes in enumerate(images_to_process):
                    base64_img = encode_image_to_base64(img_bytes)
                    
                    prompt = (
                        "Hãy đọc trang biên bản đối chiếu kho này và trả về JSON chính xác các thông tin sau:\n"
                        "1. 'ten_khach_hang': Tên đơn vị gửi hàng (Bên A) hoặc tên khách hàng đầy đủ.\n"
                        "2. 'tong_so_cuon': Tổng số lượng cuộn (hoặc dòng chi tiết) ở dòng Tổng cộng.\n"
                        "3. 'tong_khoi_luong': Tổng khối lượng còn lại (Kg) ở dòng Tổng cộng (chỉ lấy số nguyên).\n"
                        "4. 'trang_thai_ky': Đã ký đóng dấu đủ hay chưa ('Đã ký đủ 2 bên', 'Chưa ký', hoặc 'Ký sai vị trí').\n"
                        "Chỉ trả về định dạng JSON thuần túy."
                    )
                    
                    try:
                        response = client.chat.completions.create(
                            model=selected_model,
                            messages=[
                                {
                                    "role": "user",
                                    "content": [
                                        {"type": "text", "text": prompt},
                                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}}
                                    ]
                                }
                            ],
                            temperature=0.0,
                            max_tokens=500
                        ]
                        content = response.choices[0].message.content
                        # Clean JSON string if wrapped in markdown
                        if "```json" in content:
                            content = content.split("```json")[1].split("```")[0].strip()
                        elif "```" in content:
                            content = content.split("```")[1].split("```")[0].strip()
                            
                        import json
                        data_json = json.loads(content)
                        extracted_docs.append({
                            'file_name': file.name,
                            'data': data_json
                        })
                    except Exception as ex:
                        st.warning(f"⚠️ Không đọc được chi tiết file {file.name} trang {img_idx+1}: {ex}")

            progress_bar.progress(85)
            status_text.markdown("**🔄 Đang tổng hợp đối chiếu với danh sách chi nhánh...**")
            time.sleep(1)
            
            # Danh sách chuẩn các nhóm khách hàng & chi nhánh
            danh_sach_chuan = [
                {"stt": 1, "nhom": "ĐTHS", "ten": "Đầu Tư Hoa Sen", "dvt1": "Cuộn", "soluong1": 565, "dvt2": "Kg", "soluong2": 3118360},
                {"stt": 2, "nhom": "HTPP MB", "ten": "Công Ty Cổ Phần Tập Đoàn Hoa Sen - Chi Nhánh Thành Phố Hà Nội", "dvt1": "Cuộn", "soluong1": 21, "dvt2": "Kg", "soluong2": 85370},
                {"stt": 3, "nhom": "HTPP MN", "ten": "Công Ty Cổ Phần Hoa Sen Home - Chi Nhánh Tỉnh Đồng Nai", "dvt1": "Cuộn", "soluong1": 43, "dvt2": "Kg", "soluong2": 173190},
                {"stt": 4, "nhom": "HTPP MN", "ten": "Công Ty Cổ Phần Hoa Sen Home - Chi Nhánh Thành Phố Cần Thơ", "dvt1": "Cuộn", "soluong1": 13, "dvt2": "Kg", "soluong2": 44190},
                {"stt": 5, "nhom": "HTPP MN", "ten": "Công Ty Cổ Phần Hoa Sen Home - Chi Nhánh Thành Phố Hồ Chí Minh", "dvt1": "Cuộn", "soluong1": 114, "dvt2": "Kg", "soluong2": 540170},
                {"stt": 6, "nhom": "HTPP MN", "ten": "Công Ty Cổ Phần Hoa Sen Home - Chi Nhánh Tỉnh An Giang", "dvt1": "Cuộn", "soluong1": 4, "dvt2": "Kg", "soluong2": 14190},
                {"stt": 7, "nhom": "HTPP MN", "ten": "Công Ty Cổ Phần Hoa Sen Home - Chi Nhánh Tỉnh Vĩnh Long", "dvt1": "Cuộn", "soluong1": 51, "dvt2": "Kg", "soluong2": 183280},
                {"stt": 8, "nhom": "HTPP MN", "ten": "Công Ty Cổ Phần Hoa Sen Home - Chi Nhánh Tỉnh Lâm Đồng", "dvt1": "Cuộn", "soluong1": 23, "dvt2": "Kg", "soluong2": 80010},
                {"stt": 9, "nhom": "HTPP MN", "ten": "Công Ty Cổ Phần Hoa Sen Home - Chi Nhánh Tỉnh Tây Ninh", "dvt1": "Cuộn", "soluong1": 30, "dvt2": "Kg", "soluong2": 132652},
                {"stt": 10, "nhom": "KDTT MN", "ten": "Công Ty TNHH Tuyết Hảo Khánh Hòa", "dvt1": "Cuộn", "soluong1": 46, "dvt2": "Kg", "soluong2": 231460},
                {"stt": 11, "nhom": "KDTT MN", "ten": "Công Ty TNHH Sản Xuất Thương Mại Phan Long", "dvt1": "Cuộn", "soluong1": 27, "dvt2": "Kg", "soluong2": 134760},
                {"stt": 12, "nhom": "KDTT MN", "ten": "Công Ty TNHH Một Thành Viên Thương Mại Dịch Vụ Sản Xuất Cơ Khí Hải Phương Anh", "dvt1": "Cuộn", "soluong1": 66, "dvt2": "Kg", "soluong2": 283190},
                {"stt": 13, "nhom": "KDTT MN", "ten": "Công Ty TNHH Một Thành Viên Sản Xuất Thương Mại Dịch Vụ Thanh Huyền", "dvt1": "Cuộn", "soluong1": 125, "dvt2": "Kg", "soluong2": 714550},
                {"stt": 14, "nhom": "KDTT MN", "ten": "Công Ty Trách Nhiệm Hữu Hạn Sắt Thép Ngọc Biển", "dvt1": "Cuộn", "soluong1": 215, "dvt2": "Kg", "soluong2": 939030},
                {"stt": 15, "nhom": "KDTT MN", "ten": "Công Ty TNHH Tôn Thép Nguyễn Đức Anh Khoa", "dvt1": "Cuộn", "soluong1": 28, "dvt2": "Kg", "soluong2": 118720},
                {"stt": 16, "nhom": "KDTT MN", "ten": "Công Ty TNHH TM DV Thép Trường Sa", "dvt1": "Cuộn", "soluong1": 7, "dvt2": "Kg", "soluong2": 46720},
                {"stt": 17, "nhom": "KDTT MN", "ten": "Công Ty Trách Nhiệm Hữu Hạn Hi Văn", "dvt1": "Cuộn", "soluong1": 4, "dvt2": "Kg", "soluong2": 9680},
                {"stt": 18, "nhom": "KDTT MN", "ten": "Công Ty Trách Nhiệm Hữu Hạn Tân Mỹ Hạnh", "dvt1": "Cuộn", "soluong1": 39, "dvt2": "Kg", "soluong2": 172280},
                {"stt": 19, "nhom": "KDTT MN", "ten": "Công Ty Cổ Phần Vinatech Hồ Chí Minh", "dvt1": "Cuộn", "soluong1": 5, "dvt2": "Kg", "soluong2": 14240},
                {"stt": 20, "nhom": "KDTT MN", "ten": "Công Ty TNHH Tôn Thép Kokoro", "dvt1": "Cuộn", "soluong1": 1, "dvt2": "Kg", "soluong2": 4420},
                {"stt": 21, "nhom": "KDTT MN", "ten": "Công Ty TNHH Sản Xuất Thương Mại Trường Phúc Bình Thuận", "dvt1": "Cuộn", "soluong1": 2, "dvt2": "Kg", "soluong2": 7400},
                {"stt": 22, "nhom": "KDTT MN", "ten": "Công Ty TNHH Tôn Thép Vạn Niên", "dvt1": "Cuộn", "soluong1": 1, "dvt2": "Kg", "soluong2": 2640},
                {"stt": 23, "nhom": "KDTT MN", "ten": "Công Ty Cổ Phần Hoa Sen Home", "dvt1": "Cuộn", "soluong1": 742, "dvt2": "Kg", "soluong2": 3269830},
                {"stt": 24, "nhom": "KN", "ten": "Khách hàng Gia Huy", "dvt1": "Cuộn", "soluong1": 1, "dvt2": "Kg", "soluong2": 530}
            ]

            # Xử lý gán trạng thái dựa trên file quét được và danh sách chuẩn
            processed_rows = []
            for item in danh_sach_chuan:
                # Tìm xem đơn vị này có trong file được upload lên không
                matched_doc = None
                for doc in extracted_docs:
                    if any(w in doc['data'].get('ten_khach_hang', '').lower() for w in item['ten'].lower().split()[:3]):
                        matched_doc = doc
                        break
                
                row = item.copy()
                if matched_doc:
                    # Đã có biên bản
                    status_ky = matched_doc['data'].get('trang_thai_ky', '').lower()
                    if 'chưa' in status_ky:
                        row['chua_ky'] = 'x'
                        row['ky_2ben'] = ''
                        row['ky_3ben'] = ''
                        row['hoan_thanh'] = ''
                        row['nguyen_nhan'] = 'Chưa ký biên bản'
                    elif 'sai' in status_ky:
                        row['chua_ky'] = ''
                        row['ky_2ben'] = 'x'
                        row['ky_3ben'] = ''
                        row['hoan_thanh'] = ''
                        row['nguyen_nhan'] = 'Đã ký nhưng ký sai vị trí nên gửi yêu cầu ký lại'
                    else:
                        row['chua_ky'] = ''
                        row['ky_2ben'] = 'x'
                        row['ky_3ben'] = ''
                        row['hoan_thanh'] = 'x'
                        row['nguyen_nhan'] = ''
                    row['ket_qua'] = 'Chưa ghi nhận'
                    row['ghi_chu'] = 'Đã gửi kèm file biên bản'
                    row['is_error'] = False
                else:
                    # Chưa nhận được biên bản
                    row['chua_ky'] = 'x'
                    row['ky_2ben'] = ''
                    row['ky_3ben'] = ''
                    row['hoan_thanh'] = ''
                    row['nguyen_nhan'] = ''
                    row['ket_qua'] = ''
                    row['ghi_chu'] = 'Chưa nhận được biên bản'
                    row['is_error'] = False
                
                processed_rows.append(row)

            progress_bar.progress(100)
            status_text.success("✅ Hoàn tất tổng hợp đối chiếu!")
            time.sleep(1)
            progress_bar.empty()
            status_text.empty()

            # Hiển thị bảng tổng kết trên web
            st.markdown("### 📊 KẾT QUẢ ĐỐI CHIẾU & TỔNG HỢP")
            df_display = pd.DataFrame(processed_rows)
            st.dataframe(df_display[['stt', 'nhom', 'ten', 'soluong1', 'soluong2', 'ghi_chu']], use_container_width=True)

            # Tạo file Excel thành phẩm chuẩn form mẫu
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Tong_Hop"

            # Tiêu đề
            ws.merge_cells("A1:O1")
            ws["A1"] = "TỔNG HỢP ĐỐI CHIẾU HÀNG HÓA GỬI KHO TẠI NHÀ MÁY TÔN HOA SEN PHÚ MỸ"
            ws["A1"].font = Font(size=13, bold=True)
            ws["A1"].alignment = Alignment(horizontal="center")

            ws["A2"] = "Đến Hết Ngày 30/09/2026"
            ws["A2"].font = Font(italic=True)

            headers = [
                "STT", "Nhóm khách hàng", "Tên khách hàng", "ĐVT 1", "Số lượng theo ĐVT 1", 
                "ĐVT 2", "Số lượng theo ĐVT 2", "Chưa kí Biên bản", "Đã kí Biên bản 2 bên", 
                "Đã kí Biên bản 3 bên", "Đã hoàn thành, lưu trữ biên bản", "Nguyên nhân", 
                "Thời gian dự kiến hoàn thành ký Biên bản 3 bên", "Kết quả Kiểm tra Biên bản", "Ghi chú"
            ]
            ws.append([]) # Dòng 3 trống
            ws.append(headers) # Dòng 4 tiêu đề bảng

            # Định dạng header
            header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=4, column=col_num)
                cell.font = Font(bold=True)
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            # Đổ dữ liệu vào excel
            red_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid") # Màu đỏ nhạt báo lỗi

            for r_idx, r_data in enumerate(processed_rows, start=5):
                ws.append([
                    r_data['stt'], r_data['nhom'], r_data['ten'], r_data['dvt1'], r_data['soluong1'],
                    r_data['dvt2'], r_data['soluong2'], r_data['chua_ky'], r_data['ky_2ben'],
                    r_data['ky_3ben'], r_data['hoan_thanh'], r_data['nguyen_nhan'], "", r_data['ket_qua'], r_data['ghi_chu']
                ])
                
                # Nếu có lỗi hoặc nghi ngờ -> tô đỏ dòng
                if r_data.get('is_error', False):
                    for c_idx in range(1, len(headers) + 1):
                        ws.cell(row=r_idx, column=c_idx).fill = red_fill

            wb.save(output)

            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="📥 Tải Xuống File Excel Thành Phẩm",
                data=output.getvalue(),
                file_name="TONG_HOP_DOI_CHIEU_HANG_GUI_KHO_THANH_PHAM.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Đã xảy ra lỗi trong quá trình xử lý: {e}")
