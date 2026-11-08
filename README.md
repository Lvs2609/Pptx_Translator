# 🌟 PPTX Translator AI (Gemini Engine)
> **Bản địa hóa bài thuyết trình PowerPoint đa ngôn ngữ chuyên nghiệp: Bảo toàn 100% Định dạng, Màu sắc từng từ, Bảng biểu, Hình ảnh & Bố cục Slide.**

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=for-the-badge&logo=python)
![Google Gemini](https://img.shields.io/badge/Google%20Gemini-API%20v1-orange?style=for-the-badge&logo=google)
![Streamlit](https://img.shields.io/badge/Streamlit-1.30%2B-red?style=for-the-badge&logo=streamlit)
![SQLite](https://img.shields.io/badge/SQLite-WAL%20Mode-003B57?style=for-the-badge&logo=sqlite)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

---

## 👨‍💻 Tác giả & Thông tin phát triển
* **Người phát triển:** **Lương Văn Sơn (Sinh năm 2002)**
* **Mục tiêu dự án:** Cung cấp một giải pháp tự động hóa hoàn toàn việc dịch thuật bài giảng, báo cáo và tài liệu PowerPoint (`.pptx`) từ các ngôn ngữ quốc tế (Anh, Trung, Nhật, Hàn, Pháp, Đức, Tây Ban Nha) sang Tiếng Việt và ngược lại, giải quyết triệt để vấn đề vỡ khung, tràn chữ và mất định dạng màu sắc thường gặp.

---

## 🚀 Các Công Nghệ Sử Dụng (Tech Stack)

| Công nghệ | Vai trò trong hệ thống |
| :--- | :--- |
| **Python 3.10+** | Ngôn ngữ phát triển cốt lõi, xử lý đa luồng (`concurrent.futures`). |
| **Google GenAI SDK (`google-genai`)** | Kết nối trực tiếp mô hình Gemini AI thế hệ mới (`gemini-3.5-flash`, `gemini-3.5-flash-lite`, `gemini-3.1-flash-lite`) với Prompt 5 cấp độ chuyên gia. |
| **`python-pptx` & `lxml`** | Bóc tách tầng sâu cấu trúc OpenXML PowerPoint (`<a:p>`, `<a:r>`, `<a:rPr>`), chuẩn hóa c14n và phục hồi định dạng thẻ XML. |
| **Streamlit Dashboard** | Giao diện Web hiện đại, vừa khít 1 màn hình (`100vh`), hỗ trợ thích ứng cả Chế độ Sáng (**Light**) & Tối (**Dark**). |
| **SQLite 3 Database (WAL)** | Cơ sở dữ liệu nhúng cục bộ (`data/pptx_translator.db`) quản lý tập trung: Lịch sử tác vụ (Tasks), Bộ nhớ song ngữ (Translation Memory) và Hạn mức API Quota. Lưu trữ an toàn 100% trên máy người dùng, không tạo file rác. |
| **Translation Memory (TM)** | Bộ đếm song ngữ thông minh trong SQLite, tra cứu tức thì $O(1)$ sub-millisecond, tái sử dụng câu dịch cũ giúp tiết kiệm tối đa Quota API. |

---

## ✨ Điểm Nổi Bật Vượt Trội

1. **Bảo Toàn 100% Định Dạng Chi Tiết (Word-level Run Formatting):**
   - Khác với các công cụ dịch thô làm mất màu hoặc mất in đậm, hệ thống tự động đánh dấu các đoạn chữ có màu sắc, in đậm/nghiêng riêng biệt bằng thẻ `<r1>`, `<r2>`. Sau khi AI dịch xong, các thẻ này được ghép nối chính xác về đúng cụm từ tương ứng.
2. **AI Smart Auto-Layout (Tự động co chữ chống tràn viền):**
   - Tự động nhận diện độ dãn nở văn bản giữa các hệ chữ (Latinh $\leftrightarrow$ CJK Đông Á). Khi câu dịch dài hơn khung chứa ban đầu, thuật toán tự động giảm nhẹ cỡ chữ (`Auto-Shrink`) để vừa khít hộp Textbox và ô Bảng biểu.
3. **Theo Dõi Quota & Đếm Ngược Reset Thời Gian Thực:**
   - Đo độ trễ kết nối (ping ms), hiển thị số lượt đã dùng / còn lại (`Đã dùng: X / Y lượt`) kèm thanh % trực quan.
   - Khi gặp lỗi chạm hạn mức `429`, hệ thống tự động trích xuất chính xác thời gian Google mở lại và tự động chuyển sang model dự phòng (`gemini-3.5-flash-lite`).
4. **Bố Cục Giao Diện Tiện Lợi (Dashboard 2 Cột):**
   - **Cột Trái:** Chọn file, chọn ngôn ngữ và theo dõi **Live Terminal Log (220px)** chạy ngay dưới nút bấm dịch.
   - **Cột Phải:** Quản lý **Danh sách các tệp đã dịch trong SQLite DB**, kèm **nút Tải file (`📥 Tải file`)** riêng biệt cho từng task.

---

## 📁 Cấu Trúc Thư Mục Chuẩn Dev (MVC Architecture)

```text
pptx_translator/
├── .streamlit/             # Cấu hình giao diện và vô hiệu hóa telemetry
├── data/                   # Thư mục dữ liệu cục bộ (Local-Only, không đẩy lên Git)
│   ├── exports/            # Thư mục lưu trữ các file PowerPoint đã dịch (Tự động giữ 20 file gần nhất)
│   └── pptx_translator.db  # Cơ sở dữ liệu SQLite duy nhất (Tasks, TM Cache, API Quota Tracker)
├── samples/                # File PowerPoint mẫu để kiểm thử nhanh
│   └── sample_presentation.pptx
├── src/                    # Toàn bộ mã nguồn ứng dụng
│   ├── config.py           # Cấu hình tập trung, danh sách ngôn ngữ, font chữ CJK
│   ├── core/               # Lõi xử lý PPTX & Thuật toán
│   │   ├── autofit.py      # Thuật toán tự co cỡ chữ chống tràn khung
│   │   ├── parser.py       # Phân tích cú pháp XML, bảo toàn thẻ định dạng
│   │   └── translator.py   # Pipeline chia gói dịch đa luồng
│   ├── services/           # Kết nối dịch vụ bên ngoài
│   │   ├── db_service.py   # Quản lý SQLite DB và các file xuất
│   │   ├── gemini_service.py # Kết nối Gemini API, Fallback, Quota Tracker
│   │   └── memory_service.py # Quản lý bộ nhớ dịch thuật (TM)
│   └── ui/
│       └── styles.py       # CSS thích ứng theme Sáng/Tối, console log
├── app.py                  # Điểm khởi chạy Web Dashboard (Streamlit)
├── cli.py                  # Công cụ dòng lệnh (CLI)
├── run_app.bat             # File khởi chạy 1-click tiện lợi trên Windows
├── requirements.txt        # Danh sách thư viện Python cần cài đặt
├── .env.example            # Tệp mẫu hướng dẫn khai báo API Key
├── .gitignore              # Bộ lọc bảo vệ API Key và dữ liệu cá nhân khi đẩy Git
└── README.md               # Tài liệu hướng dẫn sử dụng này
```

---

## 🛠️ Hướng Dẫn Cài Đặt & Sử Dụng Chi Tiết

### Bước 1: Tải mã nguồn về máy
```bash
git clone https://github.com/<your-username>/pptx_translator.git
cd pptx_translator
```

### Bước 2: Khởi tạo và kích hoạt môi trường ảo Python
* **Trên Windows:**
  ```powershell
  python -m venv venv
  .\venv\Scripts\activate
  ```
* **Trên Linux / macOS:**
  ```bash
  python3 -m venv venv
  source venv/bin/activate
  ```

### Bước 3: Cài đặt các thư viện phụ thuộc
```bash
pip install -r requirements.txt
```

---

## 🔑 Hướng Dẫn Thêm Gemini API Key

1. **Lấy API Key hoàn toàn miễn phí:**
   - Truy cập vào: **[Google AI Studio](https://aistudio.google.com/app/apikey)**
   - Đăng nhập tài khoản Google và bấm **Create API key**.
   - Sao chép chuỗi khóa API (dạng `AQ.Ab8RN6...` hoặc `AIzaSy...`).

2. **Cách 1: Khai báo vào file `.env` (Khuyên dùng - Lưu vĩnh viễn):**
   - Đổi tên tệp `.env.example` thành `.env` (hoặc tạo file mới tên `.env` ở thư mục gốc dự án).
   - Dán khóa API của bạn vào:
     ```ini
     GEMINI_API_KEY=dán_khóa_api_của_bạn_vào_đây
     ```
   - Ứng dụng sẽ tự động nạp khóa mỗi khi mở lên mà không cần nhập lại.

3. **Cách 2: Nhập trực tiếp trên Giao diện Web:**
   - Khi mở giao diện Web, bạn có thể dán trực tiếp API Key vào ô **"🔑 Gemini API Key"** ở thanh bên trái.

---

## 🖥️ Cách Khởi Chạy Ứng Dụng

### Cách 1: Khởi chạy 1-Click trên Windows (Nhanh nhất)
* Chỉ cần **nhấp đúp chuột vào file `run_app.bat`**. 
* Ứng dụng sẽ tự động kích hoạt môi trường và mở trình duyệt web tại `http://localhost:8501`.

### Cách 2: Khởi chạy qua lệnh Web (Streamlit)
```powershell
streamlit run app.py
```

### Cách 3: Sử dụng qua công cụ dòng lệnh (CLI - Dành cho tự động hóa / Server)
```powershell
# Dịch file sang Tiếng Việt (vi)
python cli.py "duong_dan_file.pptx" -t vi

# Dịch sang Tiếng Anh với model chỉ định
python cli.py "duong_dan_file.pptx" -t en --model gemini-3.5-flash-lite

# Xem toàn bộ các tùy chọn hỗ trợ
python cli.py -h
```

---

## 🌐 Danh Sách Ngôn Ngữ Hỗ Trợ

| Mã (Code) | Tên hiển thị | Tự động gán Font chữ chuẩn |
| :---: | :--- | :--- |
| `vi` | Tiếng Việt | Mặc định hệ thống |
| `en` | Tiếng Anh (English) | Mặc định hệ thống |
| `zh` | Tiếng Trung Giản thể (Simplified) | `Microsoft YaHei` |
| `zh-tw` | Tiếng Trung Phồn thể (Traditional) | `Microsoft JhengHei` |
| `ja` | Tiếng Nhật (Japanese) | `Meiryo` |
| `ko` | Tiếng Hàn (Korean) | `Malgun Gothic` |
| `fr` | Tiếng Pháp (French) | Mặc định hệ thống |
| `de` | Tiếng Đức (German) | Mặc định hệ thống |
| `es` | Tiếng Tây Ban Nha (Spanish) | Mặc định hệ thống |

---

## 🔒 An Toàn & Bảo Mật Khi Đẩy Lên Git
* File `.gitignore` đã được cấu hình chặt chẽ để bảo vệ bạn:
  * **Tuyệt đối không đẩy file `.env` chứa API Key lên GitHub.**
  * Không đẩy các file PowerPoint cá nhân đã xuất trong `data/exports/` hoặc database `data/*.db`.
  * Không đẩy thư mục môi trường ảo `venv/` và các file cache rác.

---

## 📜 Giấy phép (License)
Dự án được phân phối dưới giấy phép **MIT License**. Mọi người đều có quyền tự do sử dụng, chỉnh sửa và đóng góp cho dự án.

---
*Phát triển với sự tận tâm bởi **Lương Văn Sơn (2002)** • 2026*
