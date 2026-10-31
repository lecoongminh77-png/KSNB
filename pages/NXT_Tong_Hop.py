# --- 4. GIAO DIỆN HƯỚNG DẪN & UPLOAD ---

# Thêm CSS riêng cho các thẻ Hướng dẫn tải file
st.markdown("""
<style>
    .upload-card {
        background: #ffffff;
        border-radius: 16px;
        padding: 24px;
        border: 1px solid rgba(0, 0, 0, 0.08);
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.03);
        margin-bottom: 20px;
        min-height: 250px;
    }
    .upload-title {
        font-size: 18px;
        font-weight: 700;
        color: #1d1d1f;
        margin-bottom: 16px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .guide-box {
        background: #fbfbfd;
        border-radius: 12px;
        padding: 16px;
        margin-bottom: 16px;
        border: 1px solid rgba(0,0,0,0.04);
    }
    .guide-header {
        font-weight: 700;
        color: #0071e3;
        font-size: 14px;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .guide-text {
        font-size: 13px;
        color: #515154;
        line-height: 1.6;
        margin: 0;
    }
    .guide-text ul {
        margin-top: 4px;
        margin-bottom: 0;
        padding-left: 20px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("### 📥 Nạp Dữ Liệu Nguồn")

# Bọc phần Upload và Hướng dẫn vào 2 cột
col_upload1, col_upload2 = st.columns(2, gap="large")

with col_upload1:
    st.markdown("""
    <div class="upload-card">
        <div class="upload-title">1️⃣ Hướng dẫn tải file INV 110</div>
        
        <div class="guide-box">
            <div class="guide-header">📌 HSG_KS_INV_Cung Ứng</div>
            <div class="guide-text">
                Kho: 031, 035, 03G, 03H, 041, 046<br>
                Thời gian: Chọn ngày thứ 5 của tuần báo cáo
            </div>
        </div>
        
        <div class="guide-box">
            <div class="guide-header">📌 HSG_KS_INV_Vật tư</div>
            <div class="guide-text">
                Kho: 03B<br>
                Thời gian: Chọn ngày thứ 5 của tuần báo cáo
            </div>
        </div>
        
        <div class="guide-box">
            <div class="guide-header">📌 NMBD_KS_PO_ADMIN</div>
            <div class="guide-text">
                Kho: 047<br>
                Thời gian: Chọn ngày thứ 5 của tuần báo cáo<br>
                <i>*Lưu ý: Xóa dữ liệu (Hợp kim kẽm nhôm magie dạng xỉ)</i>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Nút upload đặt ngay dưới thẻ hướng dẫn tương ứng
    inv110_files = st.file_uploader(
        "Nạp các file INV 110 tại đây:",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Nạp cùng lúc các file INV 110 lấy từ Cung ứng, Vật tư, Admin."
    )

with col_upload2:
    st.markdown("""
    <div class="upload-card">
        <div class="upload-title">2️⃣ Hướng dẫn tải file OPM 105 (Giao dịch SX)</div>
        
        <div class="guide-box">
            <div class="guide-header">📌 HSG_KS_INV_Cung Ứng</div>
            <div class="guide-text">
                Kho: 031, 03G, 03H, 046<br>
                Nghiệp vụ: Inventory, Job or Schedule, Move order<br>
                Thời gian: Thứ 6 tuần trước đến Thứ 5 tuần này
            </div>
        </div>
        
        <div class="guide-box">
            <div class="guide-header">📌 HSG_KS_INV_Vật tư</div>
            <div class="guide-text">
                Kho: 03B<br>
                Nghiệp vụ: Inventory, Job or Schedule, Move order<br>
                Thời gian: Thứ 6 tuần trước đến Thứ 5 tuần này
            </div>
        </div>
        
        <div class="guide-box">
            <div class="guide-header">📌 NMBD_KS_PO_ADMIN</div>
            <div class="guide-text">
                Kho: 047<br>
                Nghiệp vụ: Inventory, Job or Schedule, Move order<br>
                Thời gian: Thứ 6 tuần trước đến Thứ 5 tuần này
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Nút upload đặt ngay dưới thẻ hướng dẫn tương ứng
    file105_files = st.file_uploader(
        "Nạp các file OPM 105 tại đây:",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Nạp cùng lúc các file 105 lấy từ Cung ứng, Vật tư, Admin."
    )


# --- 5. HIỂN THỊ KẾT QUẢ VÀ NÚT TẢI EXCEL ---
if inv110_files and file105_files:
    st.divider()

    def style_table(df, is_weight_only=False):
        if is_weight_only:
            return df.style.format({"Khối lượng (Tấn)": "{:,.2f}"}).apply(
                lambda r: ['background-color: #ffff00; font-weight: bold; color: red' if r['Tên hàng'] == 'Tổng' else '' for _ in r], axis=1
            )
        else:
            return df.style.format({
                "Số lượng (Cuộn)": lambda x: f"{x:,.0f}" if pd.notnull(x) else "",
                "Khối lượng (Tấn)": "{:,.2f}"
            }).apply(
                lambda r: ['background-color: #ffff00; font-weight: bold; color: red' if r['Tên hàng'] == 'Tổng' else '' for _ in r], axis=1
            )

    # GỌI HÀM XỬ LÝ (Hãy đảm bảo hàm process_inv110 vẫn đang tồn tại ở bên trên code của bạn nhé)
    df_tb1, df_tb2, report_date_str = process_inv110(inv110_files)
    df_tb3, df_tb4 = process_105(file105_files)

    if df_tb1 is not None and df_tb3 is not None:
        col_left, col_right = st.columns(2)

        # TỐI ƯU: Dùng trực tiếp Base64 nội bộ
        logo_html = f'<img src="data:image/png;base64,{HOASEN_LOGO_B64}" width="45" style="vertical-align: middle; margin-bottom: 6px;">'

        with col_left:
            st.markdown(f"""
            <div style="text-align: center;">
                {logo_html}
                <h4 style="color: #b30000; margin: 0;">BÁO CÁO KHỐI LƯỢNG NLC, NPL TẠI CTC/NM</h4>
                <p style="font-style: italic; font-size: 13px; color: #cc0000; margin-top: 3px;">Đến 23 giờ 59 phút ngày {report_date_str}</p>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("<b style='color: red;'>1. Tồn kho HRC-PO-CR-CM</b>", unsafe_allow_html=True)
            st.dataframe(style_table(df_tb1), use_container_width=True, hide_index=True)

            st.markdown("<b style='color: red;'>2. Kim loại màu tại nhà máy NH</b>", unsafe_allow_html=True)
            st.dataframe(style_table(df_tb2, is_weight_only=True), use_container_width=True, hide_index=True)

        with col_right:
            st.markdown(f"""
            <div style="text-align: center;">
                {logo_html}
                <h4 style="color: #b30000; margin: 0;">BÁO CÁO KHỐI LƯỢNG SỬ DỤNG NLC, NPL TẠI CTC/NM</h4>
                <p style="font-style: italic; font-size: 13px; color: #cc0000; margin-top: 3px;">Đến 23 giờ 59 phút ngày {report_date_str}</p>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("<b style='color: red;'>3. Tình hình sử dụng NVL chính :</b>", unsafe_allow_html=True)
            st.dataframe(style_table(df_tb3), use_container_width=True, hide_index=True)

            st.markdown("<b style='color: red;'>4. Tình hình sử dụng NPL :</b>", unsafe_allow_html=True)
            st.dataframe(style_table(df_tb4, is_weight_only=True), use_container_width=True, hide_index=True)

        st.write("")
        excel_data = export_to_excel(df_tb1, df_tb2, df_tb3, df_tb4, report_date_str)
        st.download_button(
            label="📥 Tải Báo Cáo Ra File Excel (.xlsx)",
            data=excel_data,
            file_name=f"BAO_CAO_NXT_{report_date_str.replace('/', '-')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
