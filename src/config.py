"""Cấu hình toàn hệ thống PPTX Translator."""
import os
from pathlib import Path

# Thư mục gốc dự án
ROOT_DIR = Path(__file__).resolve().parent.parent

# Tự động nạp file .env từ thư mục gốc
try:
    from dotenv import load_dotenv
    load_dotenv(ROOT_DIR / ".env", override=True)
except ImportError:
    env_file = ROOT_DIR / ".env"
    if env_file.exists():
        for line in env_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())

# Thư mục dữ liệu & bộ nhớ
DATA_DIR = ROOT_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
MEMORY_JSON = DATA_DIR / "translation_memory.json"
MEMORY_MD = DATA_DIR / "translation_memory.md"

SAMPLES_DIR = ROOT_DIR / "samples"
SAMPLE_PPTX = SAMPLES_DIR / "sample_presentation.pptx"

# Danh sách ngôn ngữ hỗ trợ: (Tên hiển thị, Locale, Font Đông Á CJK nếu có)
LANGS = {
    "vi": ("Tiếng Việt", "vi-VN", None),
    "en": ("English", "en-US", None),
    "zh": ("Trung Giản thể (Simplified)", "zh-CN", "Microsoft YaHei"),
    "zh-tw": ("Trung Phồn thể (Traditional)", "zh-TW", "Microsoft JhengHei"),
    "ja": ("Tiếng Nhật (Japanese)", "ja-JP", "Meiryo"),
    "ko": ("Tiếng Hàn (Korean)", "ko-KR", "Malgun Gothic"),
    "fr": ("Tiếng Pháp (French)", "fr-FR", None),
    "de": ("Tiếng Đức (German)", "de-DE", None),
    "es": ("Tây Ban Nha (Spanish)", "es-ES", None),
}

# Các model LLM hỗ trợ
AVAILABLE_MODELS = [
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
    "gemini-3.8-flash",
    "Tùy chỉnh khác..."
]

FALLBACK_MODELS = [
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
    "gemini-3.8-flash"
]

