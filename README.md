# Playlist Checker

## 📋 Giới Thiệu

Playlist Checker là công cụ kiểm tra và so sánh tự động giữa file HDPS (Hồ Sơ Phát Sóng) và Playlist thực tế của kênh THVL3 và THVL4. Chương trình giúp phát hiện các sai lệch về thời gian, thời lượng, tên chương trình và quảng cáo.

## ✨ Tính Năng

### 🎯 Kiểm Tra Playlist
- So sánh tự động giữa file HDPS và Playlist
- Kiểm tra độ chính xác về:
  - ⏰ Thời gian phát sóng (cho phép sai lệch tối đa 180 giây)
  - ⌛ Thời lượng chương trình (cho phép sai lệch tối đa 10 giây)
  - 📝 Tên chương trình (độ tương đồng tối thiểu 70%)

### 📺 Kiểm Tra Quảng Cáo
- Tự động kiểm tra các file quảng cáo trong thư mục được cấu hình
- Xác minh quảng cáo có trong playlist và đúng khung giờ phát sóng
- Hỗ trợ cả kênh THVL3 và THVL4

### 🎬 Xử Lý Phim Đa Tập
- Tự động tách và kiểm tra các phần của phim đa tập
- Phát hiện lỗi thứ tự Part trong playlist
- Hỗ trợ format: "Tên phim - T.01", "Tên phim (Part 1)"

### 🔧 Chuẩn Hóa Dữ Liệu
- Tự động loại bỏ dấu tiếng Việt
- Chuẩn hóa tên chương trình để so sánh chính xác
- Hỗ trợ thay thế tên chương trình qua file cấu hình

## 🛠️ Yêu Cầu Hệ Thống

### Phần Mềm
- Python 3.7 trở lên
- Microsoft Excel (nếu cần xử lý file .xls)

### Thư Viện Python
```bash
pip install pandas openpyxl rapidfuzz unidecode pywin32
```

## 📦 Cài Đặt

1. **Clone repository:**
   ```bash
   git clone https://github.com/yourusername/playlist-checker.git
   cd playlist-checker
   ```

2. **Cài đặt các thư viện cần thiết:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Tạo file requirements.txt:**
   ```
   pandas>=1.3.0
   openpyxl>=3.0.0
   rapidfuzz>=2.0.0
   unidecode>=1.3.0
   pywin32>=301
   ```

4. **Tạo file cấu hình** (xem phần Cấu Hình bên dưới)

## ⚙️ Cấu Hình

### 1. File `ad.json`
Cấu hình khung giờ phát sóng quảng cáo cho từng kênh:

```json
{
  "3CMSP": {
    "start_time": "06:00:00",
    "end_time": "23:59:59"
  },
  "3QC": {
    "start_time": "06:00:00",
    "end_time": "23:59:59"
  },
  "3TB": {
    "start_time": "06:00:00",
    "end_time": "23:59:59"
  },
  "4CMSP": {
    "start_time": "06:00:00",
    "end_time": "23:59:59"
  },
  "4TB": {
    "start_time": "06:00:00",
    "end_time": "23:59:59"
  }
}
```

### 2. File `replacements.json`
Cấu hình thay thế tên chương trình:

```json
{
  "Tên chương trình gốc": {
    "replace_type": "full",
    "value": "Tên chương trình mới"
  },
  "Từ khóa cần thay": {
    "replace_type": "partial",
    "value": "Từ khóa mới"
  }
}
```

**Loại thay thế:**
- `"full"`: Thay thế toàn bộ tên chương trình khi khớp hoàn toàn
- `"partial"`: Thay thế một phần của tên chương trình

**Ví dụ:**
```json
{
  "Chương trình A": {
    "replace_type": "full",
    "value": "Chương trình B"
  },
  "VN": {
    "replace_type": "partial",
    "value": "Việt Nam"
  }
}
```

## 🚀 Hướng Dẫn Sử Dụng

### Khởi Động Chương Trình
```bash
python playlist_checker_v1.py
```

### Các Bước Thực Hiện

1. **Chọn File Excel HDPS:**
   - Click nút "Browse" ở dòng "File Excel HDPS"
   - Chọn file HDPS (format: `THVL3_DDMMYYYY.xls` hoặc `THVL4_DDMMYYYY.xls`)
   - Chương trình tự động nhận diện kênh và ngày

2. **Chọn File Excel Playlist:**
   - Click nút "Browse" ở dòng "File Excel Playlist"
   - Chọn file playlist tương ứng

3. **Chọn Khung Giờ:**
   - **Sáng**: Từ LOGO ĐẢO đầu tiên đến LOGO ĐẢO thứ hai
   - **Trưa**: Từ LOGO ĐẢO thứ hai đến LOGO ĐẢO thứ ba
   - **Chiều**: Từ LOGO ĐẢO thứ ba đến hết
   - **All**: Toàn bộ playlist (từ LOGO ĐẢO đầu tiên)

4. **Chạy Kiểm Tra:**
   - Click nút "Check Playlist"
   - Xem kết quả trong khung text bên dưới

## 📄 Format File Đầu Vào

### File HDPS
- **Format**: Excel (.xls hoặc .xlsx)
- **Header**: Dòng 2
- **Tên file**: `THVL3_01012025.xls` hoặc `THVL4_01012025.xls`
- **Các cột**:
  - Cột 1: Giờ (HH:MM:SS hoặc HH:MM:SS:FF)
  - Cột 2: Thời lượng
  - Cột 3: Tên chương trình

**Lưu ý về format tên chương trình trong HDPS:**
- `Phim VN: Tên phim - T.01` → Tự động tách thành Đầu phim, Nội dung, Hết tập, Đuôi phim
- `Phim sitcom: Tên phim` → Thêm hinh hiệu "HH_PHIM TRUYEN VIET NAM"
- `Cải lương: Tên vở` → Thêm hinh hiệu "San Khau Cai Luong_HD"

### File Playlist
- **Format**: Excel (.xls hoặc .xlsx)
- **Header**: Dòng 1
- **Các cột quan trọng**:
  - Cột 2: Giờ
  - Cột 3: Thời lượng
  - Cột 6: Tên chương trình

**Lưu ý:**
- Tự động lọc bỏ các dòng [Note], [Event], [Stop], [Gap]
- Tự động gộp các Part liên tiếp: `Tên CT (Part 1)`, `Tên CT (Part 2)` → `Tên CT`

## 📊 Kết Quả Kiểm Tra

### Loại Lỗi

#### 1. `[PLAYLIST LOI]`
Sai lệch giữa HDPS và Playlist

```
[PLAYLIST LOI] File: Phim_VN_T01
   Gio: HDPS=06:30:00, Playlist=06:30:45, chenh=45.0s
   Thoi luong: HDPS=00:45:00, Playlist=00:44:50, chenh=10.0s
   Ten CT: match=65% (< 70%)
   HDPS: phim vn ten phim tap 01
   Playlist: phim viet nam ten phim t01
```

#### 2. `[QUANG CAO LOI]`
Lỗi quảng cáo

```
[QUANG CAO LOI] File: CMSP_THVL3_01012025.mp4
   Khong tim thay trong playlist

[QUANG CAO LOI] File: QC-THVL3__01-01-2025_spot1.mp4
   Thoi khoang khong khop - thoi khoang chuan: 06:00:00 - 23:59:59
   Thoi khoang thuc te: 05:55:00
```

#### 3. `[LOI THU TU PART]`
Lỗi thứ tự phần phim

```
[LOI THU TU PART] Phim Hay Episode
   (Phát hiện Part 1, 3, 4 - thiếu Part 2)
```

### Ví Dụ Kết Quả Hoàn Chỉnh

```
=== KET QUA SO SANH - SECTION: SANG ===

Tong so dong HDPS (section sang): 45
Tong so dong Playlist: 45

============================================================
KIEM TRA PLAYLIST:
============================================================

[PLAYLIST LOI] File: Tin tuc sang
   Gio: HDPS=06:00:00, Playlist=06:00:15, chenh=15.0s

[PLAYLIST LOI] File: Phim bo VN tap 5
   Ten CT: match=68% (< 70%)
   HDPS: phim bo viet nam tap 05
   Playlist: phim bo vn t5

============================================================
KIEM TRA QUANG CAO:
============================================================

[QUANG CAO LOI] File: CMSP_THVL3_01012025.mp4
   Khong tim thay trong playlist

============================================================
LOI THU TU PART:
============================================================

[LOI THU TU PART] Chuong trinh giai tri
   (Phát hiện bất thường trong thứ tự Part)
```

## 🔧 Xử Lý Lỗi Thường Gặp

### 1. "Chi chon HDPS Kenh 3 hoac Kenh 4!"
**Nguyên nhân**: File HDPS không đúng format tên

**Giải pháp**: Đảm bảo tên file có chứa "THVL3" hoặc "THVL4"
```
✅ THVL3_01012025.xls
✅ HDPS_THVL4_31122024.xlsx
❌ kenh3_01012025.xls
```

### 2. "Dinh dang ngay khong hop le!"
**Nguyên nhân**: Ngày trong tên file không đúng format DDMMYYYY

**Giải pháp**: Đổi tên file theo format đúng
```
✅ THVL3_01012025.xls (01/01/2025)
❌ THVL3_2025-01-01.xls
❌ THVL3_1-1-2025.xls
```

### 3. "Khong tim thay du 3 'LOGO DAO'"
**Nguyên nhân**: File HDPS thiếu điểm đánh dấu phân chia khung giờ

**Giải pháp**: 
- Kiểm tra file HDPS có đủ 3 dòng "LOGO ĐẢO" không
- LOGO ĐẢO 1: Đầu khung sáng
- LOGO ĐẢO 2: Đầu khung trưa
- LOGO ĐẢO 3: Đầu khung chiều

### 4. Lỗi đọc file .xls
**Nguyên nhân**: Thiếu thư viện pywin32 hoặc Excel không được cài đặt

**Giải pháp**: 
```bash
pip install pywin32
```
Hoặc chuyển đổi file .xls sang .xlsx thủ công bằng Excel

### 5. Lỗi đường dẫn quảng cáo
**Nguyên nhân**: Không truy cập được thư mục network

**Giải pháp**:
- Kiểm tra kết nối mạng
- Đảm bảo có quyền truy cập vào `\\server-40t02\thanhpham$\`
- Kiểm tra thư mục quảng cáo tồn tại với đúng format ngày

## 📂 Cấu Trúc Thư Mục

```
playlist-checker/
│
├── test_playlist_checker_v8_add_part_error.py    # File chương trình chính
├── ad.json                                        # Cấu hình quảng cáo
├── replacements.json                              # Cấu hình thay thế tên CT
├── requirements.txt                               # Danh sách thư viện
├── README.md                                      # File hướng dẫn này
│
├── input/                                         # Thư mục chứa file đầu vào
│   ├── HDPS/
│   └── Playlist/
│
└── output/                                        # Thư mục chứa kết quả
```

## 🗂️ Đường Dẫn Quảng Cáo

### THVL3
```
\\server-40t02\thanhpham$\P. Quang cao\THVL3\CMSP_VL3\CMSP_THVL3_%d.%m.%Y
\\server-40t02\thanhpham$\P. Quang cao\THVL3\Quang cao VL3\QC-THVL3__%d-%m-%Y
\\server-40t02\thanhpham$\P. Quang cao\THVL3\ThongBaoVL3\%d-%m-%Y
```

### THVL4
```
\\server-40t02\thanhpham$\P. Quang cao\THVL4\CMSP_VL4\CMSP_THVL4_%d.%m.%Y
\\server-40t02\thanhpham$\P. Quang cao\THVL4\ThongBao VL4\%d-%m-%Y
```

**Format ngày**: `%d.%m.%Y` (VD: 01.01.2025) hoặc `%d-%m-%Y` (VD: 01-01-2025)

## 📝 Lưu Ý Quan Trọng

1. **Format thời gian**: 
   - Hỗ trợ cả `HH:MM:SS` và `HH:MM:SS:FF` (frame, 25fps)
   - Frame được chuyển đổi: 25 frame = 1 giây

2. **Chuẩn hóa tên**: 
   - Tự động loại bỏ dấu tiếng Việt
   - Chuyển thành chữ thường
   - Loại bỏ khoảng trắng thừa

3. **Phim đa tập**: 
   - Tự động phát hiện format: `T.01`, `Tap 01`, `HDGC 01`
   - Hỗ trợ cả đơn tập và đa tập: `19-20`, `19_20`

4. **Lọc tự động**:
   - Loại bỏ: Trailer, QC, Thông báo, CMSP từ HDPS
   - Loại bỏ: [Note], [Event], [Stop], [Gap] từ Playlist

5. **Threshold kiểm tra**:
   - Thời gian: ±180 giây (3 phút)
   - Thời lượng: ±10 giây
   - Độ tương đồng tên: ≥70%
   - "HET TAP": Yêu cầu 100% khớp chính xác

**Phiên bản**: 8.0  
**Cập nhật cuối**: 2025  
**Tính năng mới**: Kiểm tra lỗi thứ tự Part trong playlist
