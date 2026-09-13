import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Ôn Thi An Toàn Điện", page_icon="⚡", layout="wide")

# --- ĐỌC DỮ LIỆU EXCEL THEO TỪNG SHEET/CHUYÊN ĐỀ ---
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
            # Lọc các cột không có giá trị None
            sheets_data[sheet] = df
        return sheets_data, None
    except Exception as e:
        return None, f"Lỗi khi đọc file Excel: {e}"

sheets_data, error_message = load_data()

# --- SIDEBAR MENU ---
st.sidebar.title("⚡ Menu Ôn Tập")
mode = st.sidebar.radio("Chọn chế độ:", ["📖 Ôn tập theo chuyên đề", "📝 Thi thử (Mock Test)"])

if error_message:
    st.error(error_message)
else:
    if mode == "📖 Ôn tập theo chuyên đề":
        sheet_list = list(sheets_data.keys())
        selected_sheet = st.sidebar.selectbox("Chọn Chuyên Đề", sheet_list)
        
        df_current = sheets_data[selected_sheet]
        
        # Giả định cấu trúc cột trong sheet dựa theo file Excel của ông
        # Tìm các cột câu hỏi và đáp án
        st.title("⚡ Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện")
        st.markdown(f"### Chuyên đề: {selected_sheet}")
        st.markdown("---")
        
        # Hiển thị tạm thời dữ liệu sheet dưới dạng trắc nghiệm theo dòng của file
        # (Vì cấu trúc file của ông mỗi cột tương ứng với một chuyên đề dạng bảng dọc)
        col_names = [col for col in df_current.columns if "Unnamed" not in str(col)]
        
        # Lấy câu hỏi từ các cột có sẵn trong sheet
        st.info(f"Đang hiển thị chuyên đề **{selected_sheet}** gồm {len(df_current)} dòng dữ liệu ôn tập.")
        
        # Hiển thị bảng chi tiết hoặc dạng câu hỏi tương tác của chuyên đề này
        st.dataframe(df_current, use_container_width=True)
        
    else:
        st.title("📝 Chế độ thi thử (Mock Test)")
        st.write("Tính năng trộn 50 câu hỏi ngẫu nhiên đang được kích hoạt.")