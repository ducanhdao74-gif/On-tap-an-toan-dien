import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Ôn Thi An Toàn Điện", page_icon="⚡", layout="wide")

@st.cache_data
def load_data():
    file_name = "PL1. Tong hop ngan hang cau hoi an toan nam 2025 fn (1).xlsx"
    if not os.path.exists(file_name):
        return None, f"Không tìm thấy file `{file_name}` trong thư mục!"
    try:
        xls = pd.ExcelFile(file_name)
        all_questions = []
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            df['Chủ đề/Phần'] = sheet # Thêm cột phân loại theo sheet
            all_questions.append(df)
        combined_df = pd.concat(all_questions, ignore_index=True)
        return combined_df, None
    except Exception as e:
        return None, f"Lỗi khi đọc file Excel: {e}"

df_questions, error_message = load_data()

st.title("⚡ Ôn Thi Sát Hạch An Toàn Điện")

if error_message:
    st.error(error_message)
else:
    st.success(f"Đã tải thành công ngân hàng câu hỏi! Tổng số dòng dữ liệu: {len(df_questions)}")
    
    mode = st.sidebar.selectbox("Chọn chế độ ôn tập:", ["Xem danh sách câu hỏi", "Làm bài thi thử"])
    
    if mode == "Xem danh sách câu hỏi":
        st.subheader("📖 Danh sách ngân hàng câu hỏi")
        # Tìm cột chứa nội dung câu hỏi (thường là cột 1 hoặc 2 tùy file)
        st.dataframe(df_questions, use_container_width=True)
    else:
        st.subheader("📝 Chế độ thi thử")
        st.write("Hệ thống mô phỏng bài thi sẽ hiển thị tại đây.")