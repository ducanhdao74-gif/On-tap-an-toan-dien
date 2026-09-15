import time
from datetime import datetime
import streamlit as st

# --- CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Ôn Tập An Toàn Điện",
    page_icon="⚡",
    layout="centered",
)

# --- KHỞI TẠO SESSION STATE CHO TIẾN ĐỘ VÀ THỜI GIAN ---
if "start_time" not in st.session_state:
  st.session_state.start_time = time.time()

if "completed_questions" not in st.session_state:
  # Tập hợp chứa các câu hỏi đã hoàn thành (ví dụ lưu index hoặc ID câu hỏi)
  st.session_state.completed_questions = set()

if "total_questions_count" not in st.session_state:
  # Giá trị mặc định hoặc sẽ được cập nhật khi đọc file Excel
  st.session_state.total_questions_count = 100

# Tính tổng thời gian đã vào app ôn tập từ trước đến nay (tính bằng giây)
current_session_duration = time.time() - st.session_state.start_time
total_study_time_seconds = st.session_state.get(
    "accumulated_time", 0
) + current_session_duration


# --- HÀM TẠO NỘI DUNG THÔNG BÁO TIẾN ĐỘ ---
def get_progress_email_content():
  so_cau_hoan_thanh = len(st.session_state.completed_questions)
  tong_so_cau = st.session_state.total_questions_count

  # Quy đổi tổng thời gian ra giờ và phút
  gio = int(total_study_time_seconds // 3600)
  phut = int((total_study_time_seconds % 3600) // 60)

  if gio > 0:
    str_thoi_gian = f"{gio} giờ {phut} phút"
  else:
    str_thoi_gian = f"{phut} phút"

  noi_dung = f"""
Chào Đức Anh,

Hôm nay là một ngày mới rồi! Hãy dành ra chút thời gian để vào ôn tập ngân hàng câu hỏi An Toàn Điện nhé:

📊 Tiến độ hiện tại của ông:
- Số câu đã hoàn thành: {so_cau_hoan_thanh}/{tong_so_cau} câu
- Tổng thời gian đã ôn tập: {str_thoi_gian}

Chúc ông ôn thi thật tốt và đạt kết quả cao!
"""
  return noi_dung


# --- GIAO DIỆN CHÍNH CỦA ỨNG DỤNG ---
st.title("⚡ Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện")
st.write(
    "Chào Đức Anh! Hệ thống ôn tập đã sẵn sàng. Hãy chọn các tính năng bên dưới"
    " để bắt đầu."
)

# Hiển thị số liệu trực quan trên giao diện app luôn để tiện theo dõi
col1, col2 = st.columns(2)
with col1:
  st.metric(
      "Tiến độ câu hỏi",
      f"{len(st.session_state.completed_questions)} /"
      f" {st.session_state.total_questions_count}",
  )
with col2:
  gio_hien_tai = int(total_study_time_seconds // 3600)
  phut_hien_tai = int((total_study_time_seconds % 3600) // 60)
  st.metric(
      "Tổng thời gian ôn", f"{gio_hien_tai}h {phut_hien_tai}m"
  )

st.divider()

# Khu vực xem thử nội dung thông báo gửi cho bạn
with st.expander("📩 Xem trước nội dung thông báo tiến độ"):
  st.code(get_progress_email_content(), language="text")

# (Phần logic đọc Excel, làm bài quiz, gửi mail lịch trình apscheduler của bạn giữ nguyên bên dưới đây)
