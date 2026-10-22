import streamlit as st
import pandas as pd
import numpy as np

# Cấu hình hiển thị trang
st.set_page_config(
    page_title="Báo Cáo NXT NVL Chính",
    page_icon="📊",
    layout="wide"
)

st.title("📊 BÁO CÁO TỔNG HỢP NXT NVL CHÍNH")
st.markdown("Hệ thống tự động đọc, gộp dữ liệu kho ERP và xuất bảng báo cáo đối chiếu.")

# --- 1. HÀM XỬ LÝ DỮ LIỆU TỒN KHO (INV 110) ---
@st.cache_data
def process_inv110(files):
    df_list = []
    for file in files:
        try:
            # File INV 110 tiêu đề thực tế nằm ở dòng thứ 11 (skiprows=10)
            df = pd.read_excel(file, skiprows=10)
            if len(df.columns) >= 28:
                df = df.rename(columns={
                    df.columns[0]: 'Phân loại',
                    df.columns[3]: 'Tên hàng',
                    df.columns[22]: 'Khối lượng (KG)'
                })
                df_list.append(df[['Phân loại', 'Tên hàng', 'Khối lượng (KG)']])
        except Exception as e:
            st.error(f"Lỗi khi đọc file INV 110: {e}")

    if not df_list:
        return None, None

    # Gộp tất cả các file nạp vào thành 1 bảng lớn
    df = pd.concat(df_list, ignore_index=True)
    df = df.dropna(subset=['Phân loại', 'Khối lượng (KG)'], how='all')
    df['Khối lượng (KG)'] = pd.to_numeric(df['Khối lượng (KG)'], errors='coerce').fillna(0)
    df['Khối lượng (Tấn)'] = df['Khối lượng (KG)'] / 1000

    # Bảng 1: Tồn kho HRC-PO-CR-CM
    hrc = df[df['Phân loại'] == 'HRC']
    po = df[df['Phân loại'] == 'PO']
    cm_cr = df[df['Phân loại'].isin(['CM', 'CR'])]

    tb1_data = {
        'Tên hàng': [
            'Tồn kho HRC',
            'Tồn kho HRC ( hàng ngoài cảng)',
            'Tồn kho PO',
            'Tồn kho CM + CR'
        ],
        'Số lượng (Cuộn)': [
            len(hrc),
            0,
            len(po),
            len(cm_cr)
        ],
        'Khối lượng (Tấn)': [
            hrc['Khối lượng (Tấn)'].sum(),
            0.0,
            po['Khối lượng (Tấn)'].sum(),
            cm_cr['Khối lượng (Tấn)'].sum()
        ],
        'Ghi chú': ['', '', '', '']
    }
    df_tb1 = pd.DataFrame(tb1_data)
    
    # Thêm dòng Tổng kết
    df_tb1.loc[len(df_tb1)] = [
        'Tổng',
        df_tb1['Số lượng (Cuộn)'].sum(),
        df_tb1['Khối lượng (Tấn)'].sum(),
        ''
    ]

    # Bảng 2: Kim loại màu tại nhà máy NH
    cgg = df[df['Phân loại'].astype(str).str.contains('CGG', na=False, case=False)]
    nhom_kem = df[df['Phân loại'].astype(str).str.contains('nhôm kẽm', na=False, case=False)]
    kem_magie = df[df['Phân loại'].astype(str).str.contains('kẽm Magie', na=False, case=False)]
    pha_che = df[df['Phân loại'].astype(str).str.contains('pha chế', na=False, case=False)]

    tb2_data = {
        'Tên hàng': [
            'Tồn kho CGG',
            'Tồn kho Hợp kim nhôm kẽm',
            'Tồn kho Hợp kim kẽm Magie',
            'Tồn kho nguyên liệu pha chế'
        ],
        'Khối lượng (Tấn)': [
            cgg['Khối lượng (Tấn)'].sum(),
            nhom_kem['Khối lượng (Tấn)'].sum(),
            kem_magie['Khối lượng (Tấn)'].sum(),
            pha_che['Khối lượng (Tấn)'].sum()
        ],
        'Ghi chú': ['', '', '', '']
    }
    df_tb2 = pd.DataFrame(tb2_data)

    return df_tb1, df_tb2


# --- 2. HÀM XỬ LÝ DỮ LIỆU GIAO DỊCH (OMP 105) ---
@st.cache_data
def process_105(files):
    df_list = []
    for file in files:
        try:
            # File 105 tiêu đề thực tế nằm ở dòng thứ 9 (skiprows=8)
            df = pd.read_excel(file, skiprows=8)
            if len(df.columns) >= 14:
                df = df.rename(columns={
                    df.columns[2]: 'Loại Giao Dịch',
                    df.columns[5]: 'Mã Mặt Hàng',
                    df.columns[6]: 'Tên Mặt Hàng',
                    df.columns[11]: 'Khối lượng (KG)'
                })
                df_list.append(df[['Loại Giao Dịch', 'Mã Mặt Hàng', 'Tên Mặt Hàng', 'Khối lượng (KG)']])
        except Exception as e:
            st.error(f"Lỗi khi đọc file 105: {e}")

    if not df_list:
        return None, None

    df = pd.concat(df_list, ignore_index=True)
    df = df.dropna(subset=['Loại Giao Dịch', 'Khối lượng (KG)'], how='all')
    df['Khối lượng (KG)'] = pd.to_numeric(df['Khối lượng (KG)'], errors='coerce').fillna(0)

    # Lọc giao dịch sử dụng nguyên liệu sản xuất
    giao_dich_su_dung = ['WIP Issue', 'Move Order Issue']
    df_used = df[df['Loại Giao Dịch'].isin(giao_dich_su_dung)].copy()
    df_used['Khối lượng (Tấn)'] = df_used['Khối lượng (KG)'].abs() / 1000

    # Phân loại nguyên vật liệu chính
    hrc_used = df_used[df_used['Tên Mặt Hàng'].astype(str).str.contains('HRC|Cán nóng', na=False, case=False)]
    po_used = df_used[df_used['Tên Mặt Hàng'].astype(str).str.contains('PO|Tẩy rỉ', na=False, case=False)]
    cm_used = df_used[df_used['Tên Mặt Hàng'].astype(str).str.contains('CM|Cán nguội', na=False, case=False)]

    # Bảng 3: Tình hình sử dụng NVL chính
    tb3_data = {
        'Tên hàng': ['HRC', 'PO', 'CM', 'Hao hụt+ phế liệu'],
        'Số lượng (Cuộn)': [
            len(hrc_used),
            len(po_used),
            len(cm_used),
            0
        ],
        'Khối lượng (Tấn)': [
            hrc_used['Khối lượng (Tấn)'].sum(),
            po_used['Khối lượng (Tấn)'].sum(),
            cm_used['Khối lượng (Tấn)'].sum(),
            0.0
        ],
        'Ghi chú': ['', '', '', '']
    }
    df_tb3 = pd.DataFrame(tb3_data)
    
    # Thêm dòng Tổng kết
    df_tb3.loc[len(df_tb3)] = [
        'Tổng',
        df_tb3['Số lượng (Cuộn)'].sum(),
        df_tb3['Khối lượng (Tấn)'].sum(),
        ''
    ]

    # Bảng 4: Tình hình sử dụng NPL
    nhom_kem_used = df_used[df_used['Tên Mặt Hàng'].astype(str).str.contains('nhôm kẽm', na=False, case=False)]
    kem_used = df_used[
        df_used['Tên Mặt Hàng'].astype(str).str.contains('Kẽm', na=False, case=False) & 
        ~df_used['Tên Mặt Hàng'].astype(str).str.contains('nhôm|magie', na=False, case=False)
    ]
    magie_used = df_used[df_used['Tên Mặt Hàng'].astype(str).str.contains('Magie', na=False, case=False)]
    pha_che_used = df_used[df_used['Tên Mặt Hàng'].astype(str).str.contains('pha chế', na=False, case=False)]

    tb4_data = {
        'Tên hàng': [
            'Hợp kim nhôm kẽm',
            'Kẽm không hợp kim',
            'Hợp kim kẽm Magie',
            'Nguyên liệu pha chế'
        ],
        'Khối lượng (Tấn)': [
            nhom_kem_used['Khối lượng (Tấn)'].sum(),
            kem_used['Khối lượng (Tấn)'].sum(),
            magie_used['Khối lượng (Tấn)'].sum(),
            pha_che_used['Khối lượng (Tấn)'].sum()
        ],
        'Ghi chú': ['', '', '', '']
    }
    df_tb4 = pd.DataFrame(tb4_data)

    return df_tb3, df_tb4


# --- 3. GIAO DIỆN HƯỚNG DẪN VÀ NẠP FILE ---
st.markdown("### 📥 Nạp Dữ Liệu Nguồn")

with st.expander("📖 Bấm vào đây để xem Hướng dẫn lấy file từ hệ thống"):
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

with col_upload2:
    file105_files = st.file_uploader(
        "2️⃣ Tải lên các file 105 (Giao dịch sản xuất)", 
        type=["xlsx"], 
        accept_multiple_files=True,
        help="Nạp cùng lúc các file 105 lấy từ Cung ứng, Vật tư, Admin."
    )


# --- 4. XỬ LÝ VÀ HIỂN THỊ KẾT QUẢ THEO FORM MẪU ---
if inv110_files and file105_files:
    st.divider()

    # Hàm tạo style bảng: Tô vàng dòng Tổng, format số thập phân
    def format_table(df, is_khil_only=False):
        if is_khil_only:
            return df.style.format({
                "Khối lượng (Tấn)": "{:,.2f}"
            }).apply(
                lambda row: ['background-color: #ffff99; font-weight: bold' if row['Tên hàng'] == 'Tổng' else '' for _ in row],
                axis=1
            )
        else:
            return df.style.format({
                "Số lượng (Cuộn)": "{:,.0f}",
                "Khối lượng (Tấn)": "{:,.2f}"
            }).apply(
                lambda row: ['background-color: #ffff99; font-weight: bold' if row['Tên hàng'] == 'Tổng' else '' for _ in row],
                axis=1
            )

    df_tb1, df_tb2 = process_inv110(inv110_files)
    df_tb3, df_tb4 = process_105(file105_files)

    if df_tb1 is not None and df_tb3 is not None:
        col_dash1, col_dash2 = st.columns(2)

        # Cột trái: TỒN KHO NLC, NPL
        with col_dash1:
            st.markdown("<h4 style='text-align: center; color: #1e7e34;'>BÁO CÁO KHỐI LƯỢNG NLC, NPL TẠI CTC/NM</h4>", unsafe_allow_html=True)
            
            st.markdown("**1. Tồn kho HRC-PO-CR-CM**")
            st.dataframe(format_table(df_tb1), use_container_width=True, hide_index=True)

            st.markdown("**2. Kim loại màu tại nhà máy NH**")
            st.dataframe(format_table(df_tb2, is_khil_only=True), use_container_width=True, hide_index=True)

        # Cột phải: KHỐI LƯỢNG SỬ DỤNG
        with col_dash2:
            st.markdown("<h4 style='text-align: center; color: #1e7e34;'>BÁO CÁO KHỐI LƯỢNG SỬ DỤNG NLC, NPL TẠI CTC/NM</h4>", unsafe_allow_html=True)
            
            st.markdown("**3. Tình hình sử dụng NVL chính :**")
            st.dataframe(format_table(df_tb3), use_container_width=True, hide_index=True)

            st.markdown("**4. Tình hình sử dụng NPL :**")
            st.dataframe(format_table(df_tb4, is_khil_only=True), use_container_width=True, hide_index=True)
else:
    st.info("💡 Bạn hãy kéo/thả đầy đủ file INV 110 và file 105 vào 2 ô bên trên để hệ thống bắt đầu xử lý số liệu.")
