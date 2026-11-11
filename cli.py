"""Giao diện dòng lệnh (CLI) cho PPTX Translator."""
import argparse
import os
import sys
from pathlib import Path

from src.config import LANGS
from src.core.translator import translate_document
from src.services.gemini_service import GeminiTranslator, load_glossary


def main():
    ap = argparse.ArgumentParser(description="Dịch PPTX / Word / Excel / PDF giữ nguyên định dạng (Gemini AI)")
    ap.add_argument("input", help="Đường dẫn file cần dịch (.pptx, .docx, .xlsx, .pdf)")
    ap.add_argument("-s", "--source", default="auto", help="Ngôn ngữ gốc (mặc định: auto)")
    ap.add_argument("-t", "--target", choices=LANGS, required=True, help="Ngôn ngữ đích")
    ap.add_argument("-o", "--output", help="Đường dẫn file xuất ra")
    ap.add_argument("--model", default="gemini-3.5-flash", help="Mã model Gemini AI")
    ap.add_argument("--glossary", help="File txt từ điển, mỗi dòng: thuật ngữ = bản dịch")
    ap.add_argument("--no-fit", action="store_true", help="Không tự co chữ chống tràn")
    ap.add_argument("--no-cache", action="store_true", help="Không dùng bộ nhớ từ điển")
    ap.add_argument("--min-scale", type=float, default=0.65)
    a = ap.parse_args()

    key = os.getenv("GEMINI_API_KEY")
    if not key:
        sys.exit("❌ Thiếu biến môi trường GEMINI_API_KEY trong file .env")

    in_path = Path(a.input)
    out = a.output or f"{in_path.with_suffix('')}.{a.target}{in_path.suffix}"
    gl = load_glossary(Path(a.glossary).read_text(encoding="utf-8")) if a.glossary else None

    rep = translate_document(
        src=a.input,
        dst=out,
        target=a.target,
        translator=GeminiTranslator(key, a.model),
        source_lang=a.source,
        glossary=gl,
        fit=not a.no_fit,
        min_scale=a.min_scale,
        use_cache=not a.no_cache,
        progress=lambda p, m: print(f"[{int(p*100):3d}%] {m}")
    )
    print("\n" + rep.text(), f"\n→ Đã xuất file thành công: {out}")


if __name__ == "__main__":
    main()
