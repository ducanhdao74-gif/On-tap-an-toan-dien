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

# --- SIDEBAR MENU ---
st.sidebar.title("⚡ Menu Ôn Tập")
st.sidebar.markdown("**Chọn chế độ:**")
mode = st.sidebar.radio(
    "",
    [
        "📖 Ôn tập theo chuyên đề",
        "🔄 Ôn lại câu trả lời sai",
        "📂 Ôn gộp tất cả (50 câu/phần)",
        "📝 Thi thử (Mock Test)"
    ],
    label_visibility="collapsed"
)

if error_message:
    st.error(error_message)
else:
    if mode == "📖 Ôn tập theo chuyên đề":
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 📂 Chọn Chuyên Đề")
        sheet_list = list(sheets_data.keys())
        selected_sheet = st.sidebar.selectbox("", sheet_list, label_visibility="collapsed")
        
        df_current = sheets_data[selected_sheet]
        col_q = df_current.columns[0]
        valid_df = df_current.dropna(subset=[col_q]).reset_index(drop=True)
        total_q = len(valid_df)
        
        # Quản lý state cho từng chuyên đề riêng biệt
        if f"q_idx_{selected_sheet}" not in st.session_state:
            st.session_state[f"q_idx_{selected_sheet}"] = 0
        if f"done_{selected_sheet}" not in st.session_state:
            st.session_state[f"done_{selected_sheet}"] = 0

        st.sidebar.markdown("---")
        st.sidebar.markdown(f"**Tổng số câu**\n### {total_q}")
        st.sidebar.markdown(f"**Đã làm**\n### {st.session_state[f'done_{selected_sheet}']}")
        
        if st.sidebar.button("🔄 Đặt lại tiến độ chuyên đề này"):
            st.session_state[f"q_idx_{selected_sheet}"] = 0
            st.session_state[f"done_{selected_sheet}"] = 0
            st.rerun()

        # --- GIAO DIỆN CHÍNH ---
        st.title("⚡ Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện")
        
        idx = st.session_state[f"q_idx_{selected_sheet}"]
        if idx >= total_q:
            idx = 0
            st.session_state[f"q_idx_{selected_sheet}"] = 0

        st.subheader(f"Chuyên đề: {selected_sheet} (Câu {idx + 1}/{total_q})")
        st.markdown("---")

        row = valid_df.iloc[idx]
        q_text = row[col_q]

        st.markdown(f"#### {q_text}")
        st.write("Chọn đáp án của bạn:")

        # Lấy các đáp án từ các cột tiếp theo trong hàng Excel
        options = []
        for c in valid_df.columns[1:5]:
            val = row[c]
            if pd.notna(val):
                options.append(str(val))

        if not options:
            options = ["A. Đang cập nhật đáp án", "B. Đang cập nhật đáp án", "C. Đang cập nhật đáp án", "D. Đang cập nhật đáp án"]

        choice = st.radio("Đáp án", options, key=f"radio_{selected_sheet}_{idx}", label_visibility="collapsed")

        st.markdown("---")
        if st.button("Bỏ qua / Sang câu tiếp theo ➡️", type="primary"):
            st.session_state[f"done_{selected_sheet}"] = min(total_q, st.session_state[f"done_{selected_sheet}"] + 1)
            if st.session_state[f"q_idx_{selected_sheet}"] < total_q - 1:
                st.session_state[f"q_idx_{selected_sheet}"] += 1
            else:
                st.session_state[f"q_idx_{selected_sheet}"] = 0
            st.rerun()
            
    else:
        st.title("📝 Chế độ thi thử (Mock Test)")
        st.write("Hệ thống đề thi mô phỏng 50 câu ngẫu nhiên đang được chuẩn bị.")