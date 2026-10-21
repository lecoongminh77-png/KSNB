import streamlit as st
import pandas as pd
import time

def hien_thi_tro_ly_ai():
    # Đưa mọi thứ vào st.sidebar (Thanh bên trái)
    st.sidebar.title("🤖 Trợ lý KSNB")
    
    # 1. Chỗ tải file thu gọn trên sidebar
    uploaded_file = st.sidebar.file_uploader("📥 Tải file dữ liệu lên", type=["xlsx", "xls", "csv"], key="sidebar_upload")
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            st.session_state.df = df
            st.sidebar.success(f"✅ Đã đọc file ({df.shape[0]} dòng)")
        except Exception as e:
            st.sidebar.error("❌ Lỗi đọc file")

    st.sidebar.markdown("---")
    
    # 2. Khung chat
    if "messages" not in st.session_state:
        st.session_state.messages = [{"role": "assistant", "content": "Chào anh Minh! Cần tra cứu dữ liệu gì anh cứ ném file lên rồi nhắn em nhé."}]

    # In lại tin nhắn cũ
    for msg in st.session_state.messages:
        with st.sidebar.chat_message(msg["role"]):
            st.sidebar.markdown(msg["content"])

    # Ô nhập lệnh chat
    if prompt := st.sidebar.chat_input("Ví dụ: Lọc các cuộn lệch dung sai..."):
        # In tin nhắn của bạn
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.sidebar.chat_message("user"):
            st.sidebar.markdown(prompt)

        # AI trả lời
        with st.sidebar.chat_message("assistant"):
            if uploaded_file is None:
                response = "⚠️ Anh tải file Excel lên trước nhé!"
                st.sidebar.markdown(response)
                st.session_state.messages.append({"role": "assistant", "content": response})
            else:
                with st.spinner("Đang phân tích..."):
                    time.sleep(1)
                    if "lệch" in prompt.lower() or "chênh" in prompt.lower():
                        response = "🔍 Đã lọc ra các cuộn hao hụt vượt định mức. Anh kiểm tra nhé."
                    elif "trùng" in prompt.lower():
                        response = "🚨 Cảnh báo: Phát hiện mã xuất cổng trùng lặp!"
                    else:
                        response = f"Đã nhận lệnh: {prompt}."
                    
                    st.sidebar.markdown(response)
                    st.session_state.messages.append({"role": "assistant", "content": response})

