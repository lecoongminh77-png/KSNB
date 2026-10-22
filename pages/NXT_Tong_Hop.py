import streamlit as st
import pandas as pd
import math

# Cấu hình trang
st.set_page_config(page_title="Báo Cáo NXT NVL Chính", layout="wide")

st.title("📊 BÁO CÁO TỔNG HỢP NXT NVL CHÍNH")
st.markdown("Hệ thống tự động gộp các file dữ liệu kho và xuất ra dashboard tổng hợp.")

# --- HÀM XỬ LÝ DỮ LIỆU ---
@st.cache_data
def process_inv110(files):
    df_list = []
    for file in files:
        # Đọc dữ liệu, dòng tiêu đề thực tế nằm ở dòng thứ 11 (index 10)
        df = pd.read_excel(file, skiprows=10)
        
        # Đặt lại tên các cột quan trọng theo đúng index để tránh lỗi Unnamed
        # Index 0: Phân loại, Index 3: Tên hàng, Index 22: Khối lượng, Index 27: Số lượng
        if len(df.columns) >= 28:
            df = df.rename(columns={
                df.columns[0]: 'Phân loại',
                df.columns[3]: 'Tên hàng',
                df.columns[22]: 'Khối lượng (KG)',
                df.columns[27]: 'Số lượng'
            })
            df_list.append(df[['Phân loại', 'Tên hàng', 'Khối lượng (KG)', 'Số lượng']])
            
    if not df_list:
        return pd.DataFrame()
        
    # Gộp tất cả các file INV 110 lại
    df_merged = pd.concat(df_list, ignore_index=True)
    df_merged = df_merged.dropna(subset=['Phân loại', 'Khối lượng (KG)'], how='all')
    
    # Ép kiểu dữ liệu sang số để tính tổng
    df_merged['Khối lượng (KG)'] = pd.to_numeric(df_merged['Khối lượng (KG)'], errors='coerce').fillna(0)
    df_merged['Số lượng'] = pd.to_numeric(df_merged['Số lượng'], errors='coerce').fillna(0)
    
    # Quy đổi KG sang Tấn
    df_merged['Khối lượng (Tấn)'] = df_merged['Khối lượng (KG)'] / 1000
    
    # Group by để ra số tổng theo từng phân loại (HRC, PO, CR...)
    summary = df_merged.groupby('Phân loại').agg({'Số lượng': 'sum', 'Khối lượng (Tấn)': 'sum'}).reset_index()
    return summary

@st.cache_data
def process_105(files):
    df_list = []
    for file in files:
        # Dòng tiêu đề thực tế nằm ở dòng thứ 9 (index 8)
        df = pd.read_excel(file, skiprows=8)
        
        # Index 2: Loại Giao Dịch, Index 6: Tên Mặt Hàng, Index 11: Khối Lượng, Index 13: Số Lượng
        if len(df.columns) >= 14:
            df = df.rename(columns={
                df.columns[2]: 'Loại Giao Dịch',
                df.columns[6]: 'Tên Mặt Hàng',
                df.columns[11]: 'Khối lượng (KG)',
                df.columns[13]: 'Số lượng'
            })
            df_list.append(df[['Loại Giao Dịch', 'Tên Mặt Hàng', 'Khối lượng (KG)', 'Số lượng']])
            
    if not df_list:
        return pd.DataFrame()
        
    # Gộp tất cả các file 105 lại
    df_merged = pd.concat(df_list, ignore_index=True)
    df_merged = df_merged.dropna(subset=['Loại Giao Dịch', 'Khối lượng (KG)'], how='all')
    
    df_merged['Khối lượng (KG)'] = pd.to_numeric(df_merged['Khối lượng (KG)'], errors='coerce').fillna(0)
    # Quy đổi sang Tấn và lấy trị tuyệt đối (vì lệnh Issue thường mang dấu âm)
    df_merged['Khối lượng (Tấn)'] = (df_merged['Khối lượng (KG)'] / 1000).abs()
    
    # Group by theo Loại giao dịch
    summary = df_merged.groupby('Loại Giao Dịch').agg({'Khối lượng (Tấn)': 'sum'}).reset_index()
    return summary


# --- GIAO DIỆN UPLOAD FILE ---
st.markdown("### 📥 Nạp Dữ Liệu Nguồn")

# Khung ghi chú hướng dẫn lấy file chuẩn theo SOP
with st.expander("📖 Bấm vào đây để xem Hướng dẫn lấy file từ hệ thống"):
    # Chia làm 2 cột trái phải
    col_huongdan1, col_huongdan2 = st.columns(2)
    
    with col_huongdan1:
        st.markdown("""
        **INV 110**
        INV_110_GD2
        
        **HSG_KS_INV_Cung Ứng**
        - 031,035,03G,03H,041,046
        - Chọn mốc thời gian ngày thứ 5 tuần báo cáo

        **HSG_KS_INV_Vật tư**
        - 03B
        - Chọn mốc thời gian ngày thứ 5 tuần báo cáo

        **NMBD_KS_PO_ADMIN**
        - 047
        - Chọn mốc thời gian ngày thứ 5 tuần báo cáo
        - Xóa dữ liệu (Hợp kim kẽm nhôm magie (dạng xỉ))
        """)

    with col_huongdan2:
        st.markdown("""
        **OMP 105**
        HSG_OPM_105_BCKiemSoatThongTinGDTheoLo_GD2
        
        **HSG_KS_INV_Cung Ứng**
        - 031,03G,03H,046
        - Loại nghiệp vụ: Inventory, Job or Schedule, Move order
        - Chọn mốc thời gian từ thứ 6 tuần trước đến thứ 5 tuần này.

        **HSG_KS_INV_Vật tư**
        - 03B
        - Loại nghiệp vụ: Inventory, Job or Schedule, Move order
        - Chọn mốc thời gian từ thứ 6 tuần trước đến thứ 5 tuần này.

        **NMBD_KS_PO_ADMIN**
        - 047
        - Loại nghiệp vụ: Inventory, Job or Schedule, Move order
        - Chọn mốc thời gian từ thứ 6 tuần trước đến thứ 5 tuần này.
        """)

col_upload1, col_upload2 = st.columns(2)

with col_upload1:
    inv110_files = st.file_uploader(
        "1️⃣ Tải lên các file INV 110", 
        type=["xlsx"], 
        accept_multiple_files=True,
        help="Nạp cùng lúc các file INV 110 lấy từ Cung ứng, Vật tư, Admin."
    )
    if len(inv110_files) > 3:
        st.warning("Bạn đã chọn nhiều hơn 3 file. Hệ thống sẽ gộp tất cả nhưng khuyến cáo nên kiểm tra lại!")

with col_upload2:
    file105_files = st.file_uploader(
        "2️⃣ Tải lên các file 105 (Giao dịch sản xuất)", 
        type=["xlsx"], 
        accept_multiple_files=True,
        help="Nạp các file 105 tương ứng với kỳ báo cáo."
    )
    if len(file105_files) > 3:
        st.warning("Bạn đã chọn nhiều hơn 3 file. Hệ thống sẽ gộp tất cả nhưng khuyến cáo nên kiểm tra lại!")

# --- XỬ LÝ VÀ HIỂN THỊ DASHBOARD ---
if inv110_files and file105_files:
    st.divider()
    st.markdown("### 📈 BẢNG TỔNG HỢP (Dashboard NXT)")
    
    # Gọi hàm gộp data
    summary_inv = process_inv110(inv110_files)
    summary_105 = process_105(file105_files)
    
    col_dash1, col_dash2 = st.columns(2)
    
    # Cột bên trái: Dữ liệu Tồn kho (Mô phỏng cột trái file Excel)
    with col_dash1:
        st.success("TỒN KHO NVL CHÍNH (Từ INV 110)")
        # Định dạng lại bảng cho đẹp
        st.dataframe(
            summary_inv.style.format({"Số lượng": "{:,.0f}", "Khối lượng (Tấn)": "{:,.3f}"}),
            use_container_width=True,
            hide_index=True
        )
        
        tong_ton = summary_inv['Khối lượng (Tấn)'].sum()
        st.metric("Tổng Tồn Kho (Tấn)", f"{tong_ton:,.3f}")

    # Cột bên phải: Tình hình sử dụng (Mô phỏng cột phải file Excel)
    with col_dash2:
        st.info("TÌNH HÌNH SỬ DỤNG NVL / GIAO DỊCH (Từ file 105)")
        
        # Chỉ lọc những loại giao dịch quan trọng thường có trong báo cáo
        giaodich_quantrong = ['WIP Issue', 'WIP Return', 'Move Order Issue', 'Miscellaneous receipt', 'WIP Byproduct Completion', 'WIP Completion']
        filtered_105 = summary_105[summary_105['Loại Giao Dịch'].isin(giaodich_quantrong)]
        
        st.dataframe(
            filtered_105.style.format({"Khối lượng (Tấn)": "{:,.3f}"}),
            use_container_width=True,
            hide_index=True
        )

else:
    st.info("💡 Mẹo: Bạn có thể giữ phím Ctrl (hoặc Cmd) để chọn nhiều file Excel cùng lúc nạp vào mỗi ô ở trên.")

