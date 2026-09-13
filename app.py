import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện", page_icon="⚡", layout="wide")

# --- ĐỌC DỮ LIỆU EXCEL ---
@st.cache_data
def load_data():
    file_name = "PL1. Tong hop ngan hang cau hoi an toan nam 2025 fn (1).xlsx"
    if not os.path.exists(file_name):
        return None, f"Không tìm thấy file `{file_name}` trong thư mục!"
    try:
        xls = pd.ExcelFile(file_name)
        sheets_data = {}
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            sheets_data[sheet] = df
        return sheets_data, None
    except Exception as e:
        return None, f"Lỗi khi đọc file Excel: {e}"

sheets_data, error_message = load_data()

# --- SIDEBAR ---
st.sidebar.title("⚡ Menu Ôn Tập")
st.sidebar.markdown("### Chọn chế độ:")
mode = st.sidebar.radio("", ["📖 Ôn tập theo chuyên đề", "📝 Thi thử (Mock Test)"])

if error_message:
    st.error(error_message)
else:
    if mode == "📖 Ôn tập theo chuyên đề":
        sheet_list = list(sheets_data.keys())
        selected_sheet = st.sidebar.selectbox("📂 Chọn Chuyên Đề", sheet_list)
        
        df_current = sheets_data[selected_sheet]
        
        st.title("⚡ Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện")
        
        # Lọc ra các dòng có chứa dữ liệu câu hỏi (loại bỏ giá trị NaN)
        # Giả sử cột đầu tiên chứa nội dung câu hỏi
        col_question_name = df_current.columns[0]
        valid_rows = df_current.dropna(subset=[col_question_name]).reset_index(drop=True)
        
        total_qs = len(valid_rows)
        st.markdown(f"### Chuyên đề: {selected_sheet} (Tổng số: {total_qs} dòng/câu)")
        st.markdown("---")
        
        # Khởi tạo vị trí câu hỏi hiện tại trong session_state nếu chưa có
        if 'q_index' not in st.session_state:
            st.session_state.q_index = 0
            
        # Đảm bảo index không vượt quá giới hạn
        if st.session_state.q_index >= total_qs:
            st.session_state.q_index = 0

        row = valid_rows.iloc[st.session_state.q_index]
        question_text = row[col_question_name]
        
        st.markdown(f"**Câu {st.session_state.q_index + 1}:** {question_text}")
        
        # Hiển thị các cột tiếp theo làm các đáp án lựa chọn giả định
        options = [str(row[c]) for c in valid_rows.columns[1:5] if pd.notna(row[c])]
        
        if options:
            choice = st.radio("Chọn đáp án của bạn:", options, key=f"q_{st.session_state.q_index}")
        
        st.markdown("---")
        col1, col2, col3 = st.columns([1, 1, 4])
        with col1:
            if st.button("⬅️ Câu trước") and st.session_state.q_index > 0:
                st.session_state.q_index -= 1
                st.rerun()
        with col2:
            if st.button("Câu tiếp theo ➡️") and st.session_state.q_index < total_qs - 1:
                st.session_state.q_index += 1
                st.rerun()
                
    else:
        st.title("📝 Chế độ thi thử (Mock Test)")
        st.write("Hệ thống đề thi mô phỏng đang được cập nhật.")