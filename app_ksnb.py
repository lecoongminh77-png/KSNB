st.markdown("""
<style>
/* Đổi màu nền web sáng sủa hơn */
.stApp {
    background-color: #f4f7f6;
}
/* Ép các nút bấm (button) thành dạng thẻ nổi 3D hiện đại */
div.stButton > button:first-child {
    background-color: #ffffff;
    color: #0d6efd;
    border: 1px solid #e0e0e0;
    border-radius: 12px;
    padding: 20px;
    font-size: 16px;
    font-weight: bold;
    box-shadow: 0 4px 10px rgba(0,0,0,0.05);
    transition: all 0.3s ease;
    width: 100%;
}
/* Hiệu ứng bật lên và đổi màu khi di chuột vào nút */
div.stButton > button:first-child:hover {
    transform: translateY(-5px);
    box-shadow: 0 10px 20px rgba(13, 110, 253, 0.15);
    background-color: #0d6efd;
    color: #ffffff;
    border-color: #0d6efd;
}
/* Làm mượt góc các hộp màu (st.info, st.success) */
div[data-testid="stAlert"] {
    border-radius: 10px;
    border: none;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
}
</style>
""", unsafe_allow_html=True)
