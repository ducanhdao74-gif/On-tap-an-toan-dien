import os
import time
from datetime import datetime
import pandas as pd
import streamlit as st
from apscheduler.schedulers.background import BackgroundScheduler

# --- CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Ôn Tập An Toàn Điện - Điện lực Điện Biên",
    page_icon="⚡",
    layout="centered",
)

# --- KHỞI TẠO SESSION STATE ---
if "start_time" not in st.session_state:
  st.session_state.start_time = time.time()

if "completed_questions" not in st.session_state:
  st.session_state.completed_questions = set()

if "total_questions_count" not in st.session_state:
  st.session_state.total_questions_count = 0


# --- HÀM TÍNH THỜI GIAN ÔN TẬP ---
def get_total_study_time():
  current_session_duration = time.time() - st.session_state.start_time
  total_seconds = (
      st.session_state.get("accumulated_time", 0) + current_session_duration
  )
  gio = int(total_seconds // 3600)
  phut = int((total_seconds % 3600) // 60)
  if gio > 0:
    return f"{gio} giờ {phut} phút"
  else:
    return f"{phut} phút"


# --- HÀM TẠO NỘI DUNG THÔNG BÁO TIẾN ĐỘ ---
def get_progress_email_content():
  so_cau_hoan_thanh = len(st.session_state.completed_questions)
  tong_so_cau = st.session_state.get("total_questions_count", 0)
  str_thoi_gian = get_total_study_time()

  noi_dung = f"""
Chào Đức Anh,

Hôm nay là một ngày mới rồi! Hãy dành ra chút thời gian để vào ôn tập ngân hàng câu hỏi An Toàn Điện nhé:

📊 Tiến độ hiện tại của ông:
- Số câu đã hoàn thành: {so_cau_hoan_thanh}/{tong_so_cau} câu
- Tổng thời gian đã ôn tập: {str_thoi_gian}

Chúc ông ôn thi thật tốt và đạt kết quả cao!
"""
  return noi_dung


# --- GIAO DIỆN CHÍNH ---
st.title("⚡ Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện")
st.write(
    "Chào Đức Anh! Hệ thống ôn tập phục vụ ôn thi Công ty Điện lực Điện Biên."
)

# Tìm file excel bắt đầu bằng "PL1" trong thư mục hiện tại cho chính xác
excel_file = None
for f in os.listdir("."):
  if f.endswith(".xlsx") and f.startswith("PL1"):
    excel_file = f
    break

# Nếu không tìm thấy bằng chữ bắt đầu, thử quét các file excel khác
if not excel_file:
  for f in os.listdir("."):
    if f.endswith(".xlsx"):
      excel_file = f
      break

if excel_file and os.path.exists(excel_file):
  try:
    df = pd.read_excel(excel_file)
    st.session_state.total_questions_count = len(df)

    st.success(
        f"Đã tải thành công ngân hàng câu hỏi từ file `{excel_file}`! Tổng số"
        f" câu: {len(df)} câu."
    )

    # Hiển thị thống kê nhanh
    col1, col2 = st.columns(2)
    with col1:
      st.metric(
          "Tiến độ câu hỏi",
          f"{len(st.session_state.completed_questions)} /"
          f" {st.session_state.total_questions_count}",
      )
    with col2:
      st.metric("Tổng thời gian ôn", get_total_study_time())

    st.divider()

    # --- KHU VỰC LÀM BÀI ÔN TẬP ---
    st.subheader("📝 Bắt đầu ôn tập / Trắc nghiệm")

    option = st.selectbox(
        "Chọn chế độ ôn tập:", ["Làm toàn bộ câu hỏi", "Ôn theo phần"]
    )

    if option == "Làm toàn bộ câu hỏi":
      st.write("Giao diện danh sách câu hỏi hoặc làm từng câu sẽ hiển thị ở đây.")
      # Hiển thị trước một vài dòng dữ liệu của file excel để xác nhận đúng phần ôn tập
      with st.expander("Xem trước danh sách câu hỏi trong file"):
        st.dataframe(df.head())

      if st.button("Đánh dấu hoàn thành toàn bộ câu hỏi (Test)"):
        st.session_state.completed_questions = set(range(len(df)))
        st.rerun()

    # Xem thử nội dung thông báo
    with st.expander("📩 Xem trước nội dung thông báo tiến độ gửi cho ông"):
      st.code(get_progress_email_content(), language="text")

  except Exception as e:
    st.error(f"Lỗi khi đọc file Excel câu hỏi: {e}")
else:
  st.warning(
      "⚠️ Không tìm thấy file Excel nào trong thư mục GitHub. Vui lòng kiểm tra"
      " lại tên file."
  )
