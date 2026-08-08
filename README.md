# Hệ thống Tin - Sinh học nhận diện biểu cảm khuôn mặt và tư thế học sinh. Phiên bản 5.0
# I. Nhóm tác giả
- Học sinh: Nguyễn Tấn Phú, Nguyễn Xuân Trụ
- Giáo viên hướng dẫn: Huỳnh Thị Khánh Nga
- Lớp: 12/5
- Trường: THPT Nguyễn Trãi
- Năm học 2025-2026
# II. Chức năng

Hệ thống hỗ trợ theo dõi **biểu cảm khuôn mặt** và **tư thế ngồi** thông qua camera hoặc màn hình máy tính.

## 1. Tính năng

- 📷 **Quét bằng Camera**: nhận diện biểu cảm và tư thế thông qua camera.
- 🖥️ **Quét bằng màn hình**: sử dụng hình ảnh chụp từ màn hình để thực hiện nhận diện.
- 🎯 **Theo dõi ROI** (*Region of Interest*): khoanh vùng và theo dõi một học sinh cụ thể.
- 🤖 **AI Smart Monitor**: mở giao diện giám sát thông minh.
- 📊 **Xuất báo cáo**: tạo báo cáo kết quả sau phiên quét.

---

## 2. Cấu trúc các thành phần

Các file quan trọng của dự án:

```text
Project/
│
├── emotion_posture_detector_v5.0.py
├── client.py
├── AISmartMonitor.exe
│
├── model.h5
├── haarcascade_frontalface_default.xml
│
├── *.ico
├── *.ttf
└── ...
```

> **Lưu ý:** `AISmartMonitor.exe` là file executable được **build từ `client.py`**. Người dùng thông thường không cần chạy `client.py` trực tiếp mà chỉ cần chạy `AISmartMonitor.exe`.

---

## 3. Khởi động chương trình

Nếu chạy trực tiếp bằng Python:

```bash
python emotion_posture_detector_v5.0.py
```

Nếu sử dụng phiên bản đã build thành `.exe`, chỉ cần chạy file executable tương ứng.

Sau khi khởi động, chương trình hiển thị giao diện lựa chọn chế độ.

### 📷 Quét bằng Camera

Sử dụng camera để nhận diện:

- Biểu cảm khuôn mặt.
- Tư thế ngồi.
- Số lượng khuôn mặt.
- Trạng thái tư thế.

### 🖥️ Quét bằng màn hình

Sử dụng hình ảnh chụp từ màn hình để thực hiện nhận diện.

### 🤖 AI Smart Monitor

Chọn **AI Smart Monitor** để mở chương trình giám sát thông minh.

Chương trình chính sẽ tìm và khởi chạy:

`AISmartMonitor.exe`

> **Lưu ý:** `AISmartMonitor.exe` được build từ `client.py`. Khi sử dụng chức năng **AI Smart Monitor**, hãy đảm bảo `AISmartMonitor.exe` nằm đúng vị trí để chương trình chính có thể tìm và khởi chạy file.

### 📁 Xuất file log tại...

Cho phép lựa chọn thư mục lưu các file kết quả của quá trình quét.

---

## 4. Quét bằng Camera

### Bước 1 — Chọn Camera

Chọn:

**📷 QUÉT BẰNG CAMERA**

Chương trình sẽ hiển thị danh sách các camera được hệ thống nhận diện.

Chọn camera muốn sử dụng và kiểm tra hình ảnh preview trước khi bắt đầu.

### Bước 2 — Chọn lớp

Chọn lớp học tương ứng, sau đó nhấn **Mở camera**.

Sau khi camera được mở, hệ thống bắt đầu quá trình nhận diện.

---

## 5. Các phím điều khiển

Trong cửa sổ nhận diện, sử dụng các phím sau:

| Phím | Chức năng |
|---|---|
| `V` | Bật/tắt chế độ vẽ vùng ROI |
| `S` | Bắt đầu quét ROI |
| `E` | Kết thúc quét ROI và xuất báo cáo |
| `Q` | Dừng phiên quét |
| `M` | Phóng to hình ảnh |
| `N` | Thu nhỏ hình ảnh |

---

## 6. Quét ROI — theo dõi một học sinh

ROI (**Region of Interest**) được sử dụng để khoanh vùng một học sinh cụ thể trong hình ảnh.

### Bước 1 — Bật chế độ vẽ ROI

Nhấn `V` để bật chế độ vẽ ROI.

### Bước 2 — Vẽ vùng ROI

Dùng chuột kéo từ điểm bắt đầu đến điểm kết thúc để tạo khung quanh học sinh cần theo dõi.

### Bước 3 — Bắt đầu quét ROI

Nhấn `S` để bắt đầu quét vùng ROI.

Nếu chương trình yêu cầu, nhập **Student ID** trước khi bắt đầu.

### Bước 4 — Kết thúc quét ROI

Khi muốn kết thúc, nhấn `E`.

Hệ thống sẽ xử lý dữ liệu và tạo báo cáo ROI.

---

## 7. Quét bằng màn hình

Chọn:

**🖥️ QUÉT BẰNG MÀN HÌNH**

Sau đó chọn lớp học.

Hệ thống sử dụng hình ảnh chụp từ màn hình để thực hiện nhận diện.

Các phím điều khiển vẫn được hỗ trợ:

- `V` — Vẽ ROI
- `S` — Bắt đầu quét ROI
- `E` — Kết thúc quét ROI
- `M` — Phóng to
- `N` — Thu nhỏ
- `Q` — Dừng phiên quét

---

## 8. Thời gian quét tối thiểu

Một phiên quét chính yêu cầu thời gian tối thiểu:

**1800 giây (30 phút)**

Để có báo cáo tổng hợp đầy đủ, nên duy trì phiên quét trong ít nhất **30 phút**.

Nếu dừng phiên trước thời gian tối thiểu, chương trình sẽ thông báo thời gian còn thiếu.

> **Lưu ý:** Nếu xác nhận dừng phiên trước khi đủ thời gian tối thiểu, phiên đó sẽ không xuất báo cáo tổng hợp.

---

## 9. Nhận diện tư thế

Hệ thống sử dụng **MediaPipe Pose** để xác định tư thế dựa trên các điểm trên cơ thể.

Các trạng thái chính:

| Trạng thái | Điều kiện |
|---|---|
| 🟢 **Ngồi thẳng (Good)** | Góc lưng ≥ 170° |
| 🟡 **Hơi cúi (Warning)** | 150° ≤ góc < 170° |
| 🔴 **Cúi nhiều (Bad)** | Góc < 150° |
| ⚪ **Không đủ dữ liệu** | Không xác định được tư thế |

Hệ thống cũng có cơ chế cảnh báo khi phát hiện tư thế xấu kéo dài.

---

## 10. Nhận diện biểu cảm

Hệ thống sử dụng mô hình Keras để nhận diện các trạng thái biểu cảm:

- 😠 **Giận dữ**
- 🤢 **Ghê sợ**
- 😨 **Sợ hãi**
- 😄 **Vui vẻ**
- 😢 **Buồn**
- 😲 **Bất ngờ**
- 😐 **Trung lập**

Hệ thống thực hiện kiểm tra độ tin cậy của kết quả trước khi xác nhận trạng thái biểu cảm.

---

## 11. Xuất báo cáo

Sau khi phiên quét kết thúc hợp lệ, hệ thống tự động tổng hợp dữ liệu.

Báo cáo có thể bao gồm:

- Tổng thời gian quét.
- Thời gian của từng trạng thái biểu cảm.
- Tỷ lệ các trạng thái biểu cảm.
- Phân tích tư thế.
- Tỷ lệ tư thế.
- Các tín hiệu cảnh báo.
- Đánh giá tổng hợp của phiên quét.

### File CSV

Báo cáo được lưu với định dạng:

```text
Emotion_Posture_Report_YYYYMMDD_HHMMSS.csv
```

File sử dụng encoding **UTF-8-SIG** để hạn chế lỗi hiển thị tiếng Việt khi mở bằng Excel.

### File SUMMARY

Ngoài file CSV, hệ thống tạo thêm:

```text
Emotion_Posture_Report_YYYYMMDD_HHMMSS_SUMMARY.txt
```

File này chứa báo cáo tổng hợp dạng văn bản.

---

## 12. Báo cáo ROI

Khi kết thúc phiên ROI bằng phím `E`, hệ thống tạo báo cáo:

```text
ROI_Report_YYYYMMDD_HHMMSS.docx
```

Báo cáo chứa các thông tin phân tích của vùng ROI đã theo dõi.

---

## 13. AI Smart Monitor

Từ giao diện chính, chọn **AI Smart Monitor**.

Chương trình sẽ tìm và khởi chạy:

```text
AISmartMonitor.exe
```

### ⚠️ Lưu ý quan trọng

`AISmartMonitor.exe` là **file executable được build từ `client.py`**:

```text
client.py
    │
    │ Build
    ▼
AISmartMonitor.exe
```

Vì vậy:

- `client.py` là **source code**.
- `AISmartMonitor.exe` là **file executable được build từ `client.py`**.
- Người dùng cuối **không cần chạy `client.py` trực tiếp**.
- Khi sử dụng chức năng **AI Smart Monitor**, cần đảm bảo `AISmartMonitor.exe` nằm đúng vị trí để chương trình chính có thể tìm và khởi chạy file.

---

## 14. Live Stream

Trong quá trình quét, hệ thống có thể khởi động server để cung cấp hình ảnh nhận diện trực tiếp.

Server sử dụng:

`Port: 5000`

Địa chỉ truy cập có dạng:

```text
http://<IP-máy-chạy-chương-trình>:5000/
```

Thiết bị truy cập cần có kết nối mạng phù hợp với máy đang chạy chương trình.

---

## 15. Lưu ý khi sử dụng

### Camera

Đảm bảo camera đã được Windows nhận diện và có thể sử dụng trước khi chạy chương trình.

### Các file phụ thuộc

Khi chạy phiên bản Python, cần đảm bảo các file tài nguyên cần thiết nằm đúng vị trí, bao gồm:

- Model AI.
- Haar Cascade.
- Font.
- Icon.
- Các file tài nguyên khác.

### Không chạy đồng thời nhiều chế độ

Không nên chạy Camera và Fullscreen Capture cùng lúc.

### Dừng phiên sớm

Nếu phiên quét chưa đạt thời gian tối thiểu và người dùng xác nhận dừng, báo cáo tổng hợp sẽ không được xuất.

### AI Smart Monitor

Đảm bảo `AISmartMonitor.exe` được đặt đúng vị trí.

Đặc biệt, **không nhầm `AISmartMonitor.exe` với source code `client.py`**:

- `client.py` → **source code**
- `AISmartMonitor.exe` → **file đã build từ `client.py`**

---

## 16. Tóm tắt thao tác nhanh

| Chức năng | Thao tác |
|---|---|
| 📷 Camera | Chọn **QUÉT BẰNG CAMERA** |
| 🖥️ Màn hình | Chọn **QUÉT BẰNG MÀN HÌNH** |
| 🎯 Bật ROI | Nhấn `V` |
| ▶️ Bắt đầu ROI | Nhấn `S` |
| ⏹️ Kết thúc ROI | Nhấn `E` |
| 🔍 Phóng to | Nhấn `M` |
| 🔎 Thu nhỏ | Nhấn `N` |
| ⛔ Dừng phiên | Nhấn `Q` |
| 🤖 AI Smart Monitor | Chạy `AISmartMonitor.exe` |
| 📊 Báo cáo chính | `.CSV` + `.TXT` |
| 📄 Báo cáo ROI | `.DOCX` |

---

## 17. Quick Start

Quy trình sử dụng cơ bản:

**1.** Khởi động `emotion_posture_detector_v5.0.py` hoặc phiên bản `.exe`.

**2.** Chọn **Camera** hoặc **Màn hình**.

**3.** Chọn lớp học.

**4.** Bắt đầu phiên nhận diện.

**5.** Nếu cần theo dõi một học sinh cụ thể, sử dụng chức năng **ROI** với phím `V` → `S`.

**6.** Kết thúc ROI bằng phím `E` hoặc kết thúc phiên chính bằng phím `Q`.

**7.** Kiểm tra các file báo cáo được tạo trong thư mục đã chọn.

**8.** Để sử dụng AI Smart Monitor, chạy chức năng **AI Smart Monitor** và đảm bảo `AISmartMonitor.exe` có sẵn trong đúng thư mục.

# III. Tải phần mềm
[Bấm vào đây](https://epd-info.onrender.com)
# IV. Giải thưởng
Sản phẩm đã đạt giải Nhất tại Cuộc thi Khoa học kĩ thuật cấp trường THPT Nguyễn Trãi và Giải Ba cấp thành phố Đà Nẵng năm học 2025-2026
- Ngày 27/11/2025(Cấp trường):
![z7267413454464_05ab5d4d965537be988674d64058b4c5](https://github.com/user-attachments/assets/c2cf1272-0a45-4ad5-b7ca-1918e65e1780)

- Ngày 1/10/2026(Cấp thành phố):
![614478204_1434528832005855_509348093538316051_n](https://github.com/user-attachments/assets/d92ab5c2-1f92-4965-8b81-effe1bc57df3)
![614541791_1434528878672517_2539231648361761621_n](https://github.com/user-attachments/assets/834292c8-bc4c-45f2-83d0-6643bdc85822)

