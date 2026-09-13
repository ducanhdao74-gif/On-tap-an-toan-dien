import streamlit as st
import pandas as pd
import os

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="Ôn Thi An Toàn Điện", page_icon="⚡", layout="wide")

# --- ĐỌC DỮ LIỆU EXCEL ---
@st.cache_data
def load_data():
    file_name = "PL1. Tong hop ngan hang cau hoi an toan nam 2025 fn (1).xlsx"
    if not os.path.exists(file_name):
        return None, f"Không tìm thấy file `{file_name}` trong thư mục!"
    
    try:
        xls = pd.ExcelFile(file_name)
        sheet_names = xls.sheet_names
        all_questions = []
        
        for sheet in sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            all_questions.append(df)
            
        combined_df = pd.concat(all_questions, ignore_index=True)
        return combined_df, None
    except Exception as e:
        return None, f"Lỗi khi đọc file Excel: {e}"

df_questions, error_message = load_data()

# --- GIAO DIỆN CHÍNH ---
st.title("⚡ Ôn Thi Sát Hạch An Toàn Điện")

if error_message:
    st.error(error_message)
    st.info("Vui lòng đảm bảo file Excel ngân hàng câu hỏi đã được đặt cùng thư mục với `app.py` trên GitHub.")
else:
    st.success(f"Đã tải thành công ngân hàng câu hỏi! Tổng số dòng dữ liệu: {len(df_questions)}")
    
    mode = st.sidebar.selectbox("Chọn chế độ ôn tập:", ["Học theo chủ đề", "Làm bài thi thử (Mock Test)"])
    
    if mode == "Học theo chủ đề":
        st.subheader("📖 Chế độ học tập")
        st.write("Chọn phần hoặc chủ đề để bắt đầu ôn luyện các câu hỏi.")
    else:
        st.subheader("📝 Chế độ thi thử")
        st.write("Đang phát triển bài thi mô phỏng 50 câu ngẫu nhiên.")