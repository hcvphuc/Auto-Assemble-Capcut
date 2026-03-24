# Auto Video Assembly v4 — Whisper + Gemini Hybrid

Tự động tạo CapCut project từ: Audio + Script + Images → Timeline hoàn chỉnh.

## 🚀 Cách sử dụng

### Bước 1: Chuẩn bị folder project
Tạo 1 folder chứa tất cả file cần thiết:
```
📁 MyProject/
├── script_voiceover.txt    ← Script chia scene ([SCENE 1], [SCENE 2]...)
├── my_script.md            ← Script gốc (raw, để sửa chính tả)
├── audio.mp3               ← File voice-over
└── 📁 img/                 ← Folder chứa ảnh (1.png, 2.png...)
```

### Bước 2: Mở tool
Chạy `auto_assemble_v4.exe`

### Bước 3: Import
1. Click **Browse** → chọn folder project
2. Tool tự động detect tất cả files
3. Nhập **Groq API Key** (bắt buộc cho transcription)
4. Click **💾** để lưu key

### Bước 4: Transcribe (nếu chưa có SRT)
1. Click **🎤 Transcribe→SRT**
2. Whisper sẽ tạo file `.srt` trong ~20-30 giây
3. Nếu có Script (raw) `.md` → tự động sửa chính tả proper nouns

### Bước 5: Generate CapCut Project
1. Chọn **✨ Create New CapCut Project**
2. Đặt tên project
3. Click **⚡ Generate Draft & Inject**
4. CapCut sẽ tự mở với project mới!

## 🔑 API Keys

| Key | Lấy ở đâu | Dùng cho |
|-----|-----------|----------|
| **Groq API Key** | https://console.groq.com | Whisper transcription (bắt buộc) |
| **2BRAIN Key** | https://api-v2.itera102.cloud | Gemini correction (optional) |

## 📁 Cấu trúc file project

| File | Extension | Mô tả |
|------|-----------|-------|
| Scene Script | `.txt` | Script chia scene với `[SCENE n]` markers |
| Raw Script | `.md` | Script gốc (markdown) để sửa chính tả SRT |
| Audio | `.mp3/.wav/.m4a` | File voice-over |
| Images | `img/*.png` | Ảnh theo số scene (1.png, 2.png...) |
| SRT | `.srt` | Captions (tự tạo hoặc import) |

## ⚙️ Yêu cầu hệ thống
- Windows 10/11
- CapCut Desktop (đã cài đặt)
- Internet (cho Groq API)

## 💡 Tips
- **Aspect Ratio**: Chọn 9:16 cho TikTok/Shorts, 16:9 cho YouTube
- **Scene Offset**: Dùng khi SRT không bắt đầu từ scene 1
- **Auto-rename**: Tự đổi tên ảnh thành số scene
- **Auto-reload CapCut**: Tự kill → inject → relaunch CapCut

---
Made with ❤️ by Auto Assemble Team
