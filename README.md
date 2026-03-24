# Auto Assemble CapCut

Công cụ tự động hóa việc tạo project CapCut từ audio, SRT và script.

## Tính năng

- **Tự động tạo project CapCut** — Tạo project CapCut hoàn chỉnh từ scratch, không cần mở CapCut trước
- **SRT → Timeline** — Phân tích file SRT để xác định timeline chính xác cho từng scene
- **Script matching** — Sử dụng Deep Fuzzy Logic để match scenes từ script với SRT entries
- **Auto captions** — Import SRT thành **captions** (subtitles) trong CapCut, không phải text
- **Hỗ trợ nhiều Aspect Ratio** — 9:16 (TikTok), 16:9 (YouTube), 1:1, 4:5
- **Groq Whisper API** — Tích hợp transcription audio → SRT qua Groq API
- **Tắt Main Track Magnet** — Tự động disable `maintrack_adsorb` để giữ timeline chính xác

## Cách sử dụng

### Yêu cầu
- Python 3.10+
- CapCut Desktop (Windows)

### Chạy từ source
```bash
pip install pyinstaller
python auto_assemble_v2.py
```

### Build exe
```bash
pyinstaller auto_assemble_v2.spec
```

### Workflow
1. Chọn thư mục chứa audio (.mp3), SRT (.srt), script (.txt) và ảnh
2. Chọn Aspect Ratio phù hợp
3. Chọn **"Create New Project"** và đặt tên project
4. Nhấn **Run** — tool sẽ:
   - Parse SRT & script
   - Match scenes với timeline
   - Tạo project CapCut hoàn chỉnh
   - Khởi động CapCut với project mới

## Cấu trúc file

| File | Mô tả |
|------|--------|
| `auto_assemble_v2.py` | Source code chính (v2 — tạo project mới) |
| `auto_assemble.py` | Source code v1 (inject vào project có sẵn) |

## Changelog

### v2 (2026-03-24)
- Tạo project CapCut từ scratch (không cần project có sẵn)
- SRT import dạng **captions/subtitles** thay vì text
- Fix Aspect Ratio — canvas_config được apply đúng từ user selection
- Sử dụng template base từ project có sẵn để đảm bảo schema đầy đủ
- Kill-Write-Launch workflow cho injection đáng tin cậy
