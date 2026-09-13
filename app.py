import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện", page_icon="⚡", layout="wide")

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
        
        # Lọc các dòng có câu hỏi (giả sử cột chứa câu hỏi là cột có chứa text hoặc lấy các cột từ 1 đến 5 làm câu hỏi và đáp án)
        valid_df = df_current.dropna(how="all").reset_index(drop=True)
        total_q = len(valid_df)
        
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

        st.title("⚡ Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện")
        
        idx = st.session_state[f"q_idx_{selected_sheet}"]
        if idx >= total_q:
            idx = 0
            st.session_state[f"q_idx_{selected_sheet}"] = 0

        row = valid_df.iloc[idx]
        
        # Lấy nội dung câu hỏi (thường ở cột 0 hoặc cột 1) và các đáp án ở các cột kế tiếp
        cols = list(valid_df.columns)
        q_text = row[cols[0]] if pd.notna(row[cols[0]]) else f"Câu hỏi {idx + 1}"
        
        st.subheader(f"Chuyên đề: {selected_sheet} (Câu {idx + 1}/{total_q})")
        st.markdown("---")
        st.markdown(f"#### {q_text}")
        st.write("Chọn đáp án của bạn:")

        # Lấy tất cả các ô có dữ liệu trong dòng này làm phương án lựa chọn
        options = []
        for c in cols[1:]:
            val = row[c]
            if pd.notna(val) and str(val).strip() != "":
                options.append(str(val))
                
        if not options:
            options = ["A. Không có dữ liệu đáp án", "B. Kiểm tra lại file Excel", "C. ---", "D. ---"]

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