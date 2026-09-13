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
            col = df.columns[0]
            
            questions = []
            current_q = None
            current_opts = []
            
            for val in df[col].dropna():
                val_str = str(val).strip()
                if val_str.lower().startswith("câu"):
                    if current_q:
                        questions.append({"question": current_q, "options": current_opts})
                    current_q = val_str
                    current_opts = []
                elif val_str.lower().startswith(("a.", "b.", "c.", "d.")):
                    current_opts.append(val_str)
                else:
                    if current_q and not current_opts:
                        current_q += " " + val_str
                    elif current_opts:
                        current_opts[-1] += " " + val_str
            if current_q:
                questions.append({"question": current_q, "options": current_opts})
                
            # Fallback nếu file cấu trúc khác
            if not questions:
                for idx, row in df.iterrows():
                    row_vals = [str(x) for x in row.values if pd.notna(x)]
                    if row_vals:
                        questions.append({"question": row_vals[0], "options": row_vals[1:] if len(row_vals)>1 else ["A. Đang cập nhật", "B. ---", "C. ---", "D. ---"]})
                        
            sheets_data[sheet] = questions
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
        
        q_list = sheets_data[selected_sheet]
        total_q = len(q_list)
        
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
        if idx >= total_q and total_q > 0:
            idx = 0
            st.session_state[f"q_idx_{selected_sheet}"] = 0

        if total_q == 0:
            st.warning("Chuyên đề này hiện chưa có câu hỏi nào được trích xuất.")
        else:
            q_item = q_list[idx]
            
            st.subheader(f"Chuyên đề: {selected_sheet} (Câu {idx + 1}/{total_q})")
            st.markdown("---")
            st.markdown(f"#### {q_item['question']}")
            st.write("Chọn đáp án của bạn:")

            options = q_item['options']
            if not options:
                options = ["A. Đang cập nhật đáp án", "B. ---", "C. ---", "D. ---"]

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