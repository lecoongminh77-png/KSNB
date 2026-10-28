import sys
import os
import io
import time
import pandas as pd
import streamlit as st
from openpyxl.styles import PatternFill, Font, Border, Side

# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN
# -------------------------------------------------------------
st.set_page_config(page_title='Kiểm Soát Báo Cáo Cắt Tấm', page_icon='🛡️', layout='wide')

st.markdown("""
<style>
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    .stApp { background-color: #f5f5f7; font-family: -apple-system, sans-serif; color: #1d1d1f; }
    .apple-header-box { display: flex; justify-content: space-between; align-items: flex-end; padding: 20px 0 10px 0; margin-bottom: 5px; }
    .apple-main-title { font-size: 40px; font-weight: 700; color: #1d1d1f; letter-spacing: -0.025em; margin: 0; }
    div.stButton > button:first-child { background: #0071e3 !important; color: #ffffff !important; border-radius: 980px !important; border: none !important; padding: 10px 20px !important; font-weight: 600 !important; }
    div.stButton > button:first-child:hover { background: #0077ed !important; }
    div.stDownloadButton > button:first-child { background: #28a745 !important; color: #ffffff !important; border-radius: 980px !important; border: none !important; padding: 10px 20px !important; font-weight: 600 !important; }
    div.stDownloadButton > button:first-child:hover { background: #218838 !important; }
    [data-testid="stVerticalBlockBorderWrapper"] { background: #ffffff; border-radius: 16px; padding: 15px; border: 1px solid rgba(0, 0, 0, 0.08); box-shadow: 0 2px 12px rgba(0, 0, 0, 0.03); }
    div[data-testid="stMetric"] { background-color: #ffffff; border-radius: 12px; padding: 15px; border: 1px solid rgba(0,0,0,0.05); box-shadow: 0 2px 8px rgba(0,0,0,0.02); text-align: center; }
    .upload-title { font-size: 18px; font-weight: 700; color: #1d1d1f; margin-bottom: 16px; }
    .guide-box { background: #fbfbfd; border-radius: 12px; padding: 16px; margin-bottom: 16px; border: 1px solid rgba(0,0,0,0.04); }
    .guide-header { font-weight: 700; color: #0071e3; font-size: 14px; margin-bottom: 8px; }
</style>
""", unsafe_allow_html=True)

col_home, _ = st.columns([1.5, 8.5])
with col_home:
    if st.button('🏠 Trang chủ', use_container_width=True):
        st.switch_page('app_ksnb.py')

st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 10px 0 20px 0;'>", unsafe_allow_html=True)
st.markdown("<div class='apple-header-box'><div><h1 class='apple-main-title'>🏭 Kiểm Soát Báo Cáo Bao Bì & Cắt Tấm</h1></div></div>", unsafe_allow_html=True)
st.markdown("<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 15px 0 25px 0;'>", unsafe_allow_html=True)

# -------------------------------------------------------------
# HÀM HỖ TRỢ
# -------------------------------------------------------------
def is_negative(val):
    if pd.isna(val): return False
    if isinstance(val, (int, float)): return val < 0
    if isinstance(val, str):
        val = val.strip()
        if (val.startswith('(') and val.endswith(')')) or val.startswith('-'): return True
    return False

def clean_ma_lo(val):
    if pd.isna(val): return ""
    return str(val).strip().upper()

def parse_and_check_date(val, t_month, t_year):
    """Kiểm tra Ngày xuất kho có khớp với Tháng/Năm đang chọn hay không"""
    if pd.isna(val) or str(val).strip() == "": 
        return False, ""
    try:
        dt = pd.to_datetime(str(val), errors='coerce')
        if pd.notna(dt):
            if dt.month == t_month and dt.year == t_year:
                return True, dt.strftime("%d/%m/%Y")
    except:
        pass
    return False, ""

# -------------------------------------------------------------
# LOGIC CHÍNH
# -------------------------------------------------------------
def kiem_soat_san_xuat():
    st.caption('Kiểm soát nội bộ • Đối chiếu theo Ngày xuất kho giữa Báo cáo sản xuất và ERP 105')
    st.markdown("### 📥 Nạp Dữ Liệu Nguồn")
    
    col_t1, col_t2, _ = st.columns([2, 2, 6])
    with col_t1: thang_chon = st.selectbox('Chọn Tháng', list(range(1, 13)), index=9)
    with col_t2: nam_chon = st.selectbox('Chọn Năm', [2024, 2025, 2026, 2027], index=2)
    chuoi_yymm = f"{str(nam_chon)[-2:]}{thang_chon:02d}"
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_upload1, col_upload2 = st.columns(2, gap="large")
    with col_upload1:
        with st.container(border=True):
            st.markdown("<div class='upload-title'>1️⃣ Dữ liệu Báo cáo Cắt tấm</div><div class='guide-box'><div class='guide-header'>📌 Báo cáo cắt tấm (BCSX)</div>Dữ liệu lấy từ các sheet QLSX và NXT NGUYEN LIEU.</div>", unsafe_allow_html=True)
            f_bc = st.file_uploader('Tải lên File BÁO CÁO CẮT TẤM', type=['xlsx', 'xls'])

    with col_upload2:
        with st.container(border=True):
            st.markdown("<div class='upload-title'>2️⃣ Dữ liệu Giao dịch ERP 105</div><div class='guide-box'><div class='guide-header'>📌 Báo cáo OPM 105</div>Tích chọn Kho: <b>031</b> | Giao dịch: <b>Move order</b></div>", unsafe_allow_html=True)
            f_105 = st.file_uploader('Tải lên File OPM 105', type=['xlsx', 'xls'])

    if st.button('🚀 Bắt đầu đối chiếu tự động', use_container_width=True):
        if not f_bc and not f_105:
            st.warning("⚠️ Vui lòng tải lên ít nhất một file để kiểm tra!")
            return
        progress_bar = st.progress(0, text="Đang khởi động...")
        
        bc_sheets = []
        if f_bc:
            xl_bc = pd.ExcelFile(f_bc)
            bc_sheets = xl_bc.sheet_names

        # ==========================================
        # PHẦN 1: KIỂM TRA MÃ KIỆN BỊ ÂM & THIẾU
        # ==========================================
        if f_bc is not None:
            if 'NXT_TAM_LOT' not in bc_sheets:
                st.error("🚨 Không tìm thấy sheet 'NXT_TAM_LOT' trong file Báo Cáo!")
            else:
                progress_bar.progress(20, text="Đang quét dữ liệu báo cáo cắt tấm...")
                st.markdown("### 📊 KẾT QUẢ: KIỂM SOÁT CẮT TẤM")
                try:
                    df_nxt_tam_lot = pd.read_excel(f_bc, sheet_name='NXT_TAM_LOT', header=None, dtype=str)
                    ma_kien_am = [(str(row[1]).strip(), row[11]) for idx, row in df_nxt_tam_lot.iterrows() 
                                  if pd.notna(row[1]) and str(row[1]).strip().lower() != "nan" and idx > 1 and is_negative(row[11])]
                    
                    if ma_kien_am:
                        st.error(f"🚨 PHÁT HIỆN SỐ TẤM BỊ ÂM TRONG SHEET 'NXT_TAM_LOT':")
                        for mk, st_am in ma_kien_am: st.write(f"- Mã kiện: **{mk}** | Số tấm: **{st_am}**")
                    else:
                        st.success("✅ Sheet 'NXT_TAM_LOT': Không có trường hợp nào số tấm bị âm.")

                    ds_sheets = {'QLSX NIENG': 'SXNI', 'QLSX NHUNG': 'SXNH', 'QLSX_CAT_TAM MAT TRANG': 'CTMT', 'QLSX_CAT_TAM LOT NGOAI': 'CTLN', 'QLSX_CAT_TAM LOT TRONG': 'CTLT', 'QLSX TP MAT TRANG': 'SXMT'}
                    nxt_ma_kiens = set(df_nxt_tam_lot[1].dropna().astype(str).str.strip().tolist())

                    for sheet_name, prefix in ds_sheets.items():
                        sheet_real_name = next((s for s in bc_sheets if sheet_name.lower() in s.lower()), None)
                        if sheet_real_name:
                            try:
                                df_qlsx = pd.read_excel(f_bc, sheet_name=sheet_real_name, header=None, dtype=str)
                                max_seq = 0
                                for val in df_qlsx[12].dropna():
                                    val_str = str(val).strip()
                                    if val_str.startswith(f"{prefix}{chuoi_yymm}"):
                                        try:
                                            seq = int(val_str[-3:])
                                            if seq > max_seq: max_seq = seq
                                        except ValueError: pass
                                
                                with st.expander(f"Kết quả: {sheet_real_name} ({prefix}{chuoi_yymm})"):
                                    if max_seq == 0:
                                        st.info(f"Không có mã kiện thuộc tháng {thang_chon}/{nam_chon}.")
                                    else:
                                        st.write(f"Số lượng mã lớn nhất trong tháng: **{max_seq}** kiện.")
                                        thieu = [f"{prefix}{chuoi_yymm}{i:03d}" for i in range(1, max_seq + 1) if f"{prefix}{chuoi_yymm}{i:03d}" not in nxt_ma_kiens]
                                        if thieu:
                                            st.error(f"❌ Phát hiện thiếu **{len(thieu)}** mã kiện bên NXT_TAM_LOT:")
                                            st.write(", ".join(thieu))
                                        else:
                                            st.success("✅ Đủ số lượng, không rớt mã lô nào.")
                            except Exception as e: st.error(f"Lỗi đọc {sheet_real_name}: {e}")
                except Exception as e: st.error(f"Lỗi hệ thống: {e}")

        # ==========================================
        # PHẦN 2: ĐỐI CHIẾU ERP 105 <-> NXT (CÓ LỌC NGÀY XUẤT)
        # ==========================================
        if f_105 is not None and f_bc is not None:
            progress_bar.progress(60, text="Đang đồng bộ và đối chiếu chéo OPM 105 với sổ sách...")
            st.markdown("### 📊 DASHBOARD: ĐỐI CHIẾU CHÉO ERP 105 & NXT NGUYÊN LIỆU")
            
            try:
                # 1. Xử lý ERP 105 (Giữ nguyên)
                df_105_raw = pd.read_excel(f_105, header=None, dtype=str)
                df_105_raw['Mã Lô 105'] = df_105_raw[5].apply(clean_ma_lo)
                df_105_raw['Khối Lượng 105'] = pd.to_numeric(df_105_raw[9], errors='coerce').fillna(0)
                df_105_filtered = df_105_raw[(df_105_raw['Mã Lô 105'] != "") & (df_105_raw['Mã Lô 105'].str.len() > 3)]
                df_105_grouped = df_105_filtered.groupby('Mã Lô 105', as_index=False)['Khối Lượng 105'].sum()
                dict_105 = dict(zip(df_105_grouped['Mã Lô 105'], df_105_grouped['Khối Lượng 105']))

                # 2. Xử lý NXT NGUYÊN LIỆU (Có lọc ngày cột P - Index 15)
                nl_sheet_name = next((s for s in bc_sheets if "NXT NGUYEN LIEU" in s.upper()), None)
                if not nl_sheet_name:
                    st.error("🚨 Không tìm thấy sheet có tên 'NXT NGUYEN LIEU' trong Báo cáo.")
                else:
                    df_nl = pd.read_excel(f_bc, sheet_name=nl_sheet_name, header=None, dtype=str)
                    dict_nl = {}
                    dict_ngay_xuat = {} # Lưu lại ngày xuất kho
                    
                    for idx, row in df_nl.iterrows():
                        if idx < 2: continue # Bỏ qua header
                        
                        ngay_xuat_raw = row[15] # Cột P - Ngày xuất kho
                        is_match, date_str = parse_and_check_date(ngay_xuat_raw, thang_chon, nam_chon)
                        
                        # CHỈ LẤY CÁC DÒNG XUẤT KHO TRONG THÁNG ĐANG CHỌN
                        if is_match:
                            ma_lo_nxt = clean_ma_lo(row[1])
                            kl_nxt = pd.to_numeric(row[7], errors='coerce') if pd.notna(row[7]) else 0
                            if ma_lo_nxt != "" and len(ma_lo_nxt) > 3:
                                dict_nl[ma_lo_nxt] = dict_nl.get(ma_lo_nxt, 0) + kl_nxt
                                dict_ngay_xuat[ma_lo_nxt] = date_str

                    # 3. ĐỐI CHIẾU 2 CHIỀU (Full Outer Join)
                    all_ma_lo = set(dict_105.keys()).union(set(dict_nl.keys()))
                    
                    ket_qua = []
                    so_dat = so_lech = so_thieu_nxt = so_thieu_erp = 0

                    for ma in all_ma_lo:
                        kl_105 = dict_105.get(ma, 0) # Âm
                        kl_nxt = dict_nl.get(ma, 0) # Dương
                        chenh_lech = kl_105 + kl_nxt 
                        ngay_xuat = dict_ngay_xuat.get(ma, "")
                        
                        if ma in dict_105 and ma in dict_nl:
                            if round(chenh_lech, 2) == 0:
                                trang_thai = "Đã có thông tin trên ERP (Khớp 100%)"
                                so_dat += 1
                            else:
                                trang_thai = "Đã có thông tin trên ERP (Bị lệch Khối lượng)"
                                so_lech += 1
                        elif ma in dict_nl and ma not in dict_105:
                            trang_thai = "Chưa có thông tin trên ERP"
                            so_thieu_erp += 1
                        else:
                            # Lô này có trên ERP nhưng Kế toán chưa ghi vào NXT hoặc ghi lệch tháng
                            trang_thai = f"Có giao dịch ERP nhưng KHÔNG GHI SỔ xuất kho tháng {thang_chon}/{nam_chon}"
                            so_thieu_nxt += 1
                            
                        ket_qua.append({
                            'Ngày xuất kho': ngay_xuat,
                            'Mã số cuộn': ma,
                            'Khối lượng Sổ NXT (Dương)': kl_nxt,
                            'Khối lượng ERP 105 (Âm)': kl_105,
                            'Chênh lệch (kg)': round(chenh_lech, 2),
                            'Ghi chú': trang_thai
                        })

                    # Sắp xếp lại danh sách: Ưu tiên hiển thị lô có ngày xuất kho lên trước
                    ket_qua = sorted(ket_qua, key=lambda x: (x['Ngày xuất kho'] == "", x['Ngày xuất kho']))

                    # 4. Hiển thị Dashboard (4 Chỉ số)
                    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                    col_m1.metric("📦 Tổng Mã Lô", f"{len(all_ma_lo)} Lô")
                    col_m2.metric("✅ Đã Có ERP (Khớp)", f"{so_dat} Lô")
                    col_m3.metric("⚠️ Đã Có ERP (Lệch KL)", f"{so_lech} Lô", delta=f"{so_lech} lỗi", delta_color="inverse")
                    
                    loi_lech_ma = so_thieu_nxt + so_thieu_erp
                    col_m4.metric("🚨 Lỗi Thiếu / Khống", f"{loi_lech_ma} Lô", delta=f"{loi_lech_ma} lỗi", delta_color="inverse")

                    # 5. Hiển thị Bảng & Xuất Excel
                    if len(all_ma_lo) > 0:
                        df_kq = pd.DataFrame(ket_qua)
                        st.dataframe(df_kq, use_container_width=True)

                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_kq.to_excel(writer, index=False, sheet_name='Doi_Chieu_Xuat_Kho')
                            ws = writer.sheets['Doi_Chieu_Xuat_Kho']
                            
                            f_green = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
                            t_green = Font(color='006100')
                            f_yel = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
                            t_yel = Font(color='9C5700')
                            f_red = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
                            t_red = Font(color='9C0006')
                            b_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

                            for r in range(2, len(df_kq) + 2):
                                # Kiểm tra giá trị ở cột Ghi chú (Cột số 6)
                                status = str(ws.cell(row=r, column=6).value)
                                for c in range(1, 7):
                                    cell = ws.cell(row=r, column=c)
                                    cell.border = b_thin
                                    if "Khớp" in status: 
                                        cell.fill, cell.font = f_green, t_green
                                    elif "Lệch" in status: 
                                        cell.fill, cell.font = f_yel, t_yel
                                    elif "Chưa có" in status or "KHÔNG GHI SỔ" in status: 
                                        cell.fill, cell.font = f_red, t_red
                                        
                            for col in ws.columns:
                                max_len = max((len(str(c.value)) for c in col if c.value), default=0)
                                ws.column_dimensions[col[0].column_letter].width = max_len + 2

                        st.download_button("📥 Tải Báo Cáo Đối Chiếu", data=output.getvalue(), file_name=f"Doi_Chieu_Xuat_Kho_{chuoi_yymm}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    else: st.info("Không tìm thấy giao dịch nào xuất kho trong tháng này!")

            except Exception as e: st.error(f"Lỗi hệ thống khi đối chiếu: {e}")

        progress_bar.progress(100, text="✅ Hoàn tất!")
        time.sleep(1)
        progress_bar.empty()

kiem_soat_san_xuat()
