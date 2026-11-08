"""Quản trị Bộ nhớ Dịch thuật Song ngữ (Translation Memory - TM).
100% dữ liệu được lưu trữ trong bảng `translation_memory` của SQLite cục bộ (data/pptx_translator.db).
Hoàn toàn không sinh file rác JSON hay MD ra ổ cứng; xuất Markdown động trong RAM khi người dùng tải.
"""
from __future__ import annotations

import threading
from datetime import datetime
from src.config import DATA_DIR
from src.services.db_service import get_connection, init_db, DB_PATH

_mem_lock = threading.RLock()


def get_cached_translations(sources: list[str], target: str) -> tuple[dict[str, str], list[str]]:
    """Tách sources thành cached (đã có trong SQLite TM) và remaining (cần dịch mới)."""
    init_db()
    cached: dict[str, str] = {}
    remaining: list[str] = []

    cleaned_sources = [s.strip() for s in sources if s.strip()]
    if not cleaned_sources:
        return {}, sources

    found_map: dict[str, str] = {}
    chunk_size = 400

    with _mem_lock:
        with get_connection() as conn:
            for i in range(0, len(cleaned_sources), chunk_size):
                chunk = cleaned_sources[i:i + chunk_size]
                placeholders = ",".join("?" for _ in chunk)
                query = f"""
                SELECT source_text, translated_text 
                FROM translation_memory 
                WHERE target_lang = ? AND source_text IN ({placeholders});
                """
                params = [target] + chunk
                rows = conn.execute(query, params).fetchall()
                for r in rows:
                    found_map[r["source_text"]] = r["translated_text"]

    for s in sources:
        key = s.strip()
        if key in found_map:
            cached[s] = found_map[key]
        else:
            remaining.append(s)

    return cached, remaining


def update_cache(translations: dict[str, str], target: str):
    """Cập nhật các câu mới vào bảng translation_memory trong SQLite."""
    if not translations:
        return
    init_db()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []
    for s, d in translations.items():
        if s and s.strip() and d and d.strip():
            rows.append((target, s.strip(), d.strip(), now_str, now_str))

    if not rows:
        return

    with _mem_lock:
        with get_connection() as conn:
            conn.executemany("""
            INSERT INTO translation_memory (target_lang, source_text, translated_text, used_count, created_at, updated_at)
            VALUES (?, ?, ?, 1, ?, ?)
            ON CONFLICT(target_lang, source_text) DO UPDATE SET
                translated_text = excluded.translated_text,
                used_count = used_count + 1,
                updated_at = excluded.updated_at;
            """, rows)
            conn.commit()


def clear_cache() -> int:
    """Xóa toàn bộ các câu trong bộ nhớ từ điển SQLite."""
    init_db()
    with _mem_lock:
        with get_connection() as conn:
            row = conn.execute("SELECT COUNT(*) FROM translation_memory;").fetchone()
            count = row[0] if row else 0
            conn.execute("DELETE FROM translation_memory;")
            conn.execute("VACUUM;")
            conn.commit()
            return count


def get_cache_stats() -> dict:
    """Lấy số lượng câu đã lưu trong SQLite và dung lượng file DB."""
    init_db()
    with _mem_lock:
        with get_connection() as conn:
            rows = conn.execute("SELECT target_lang, COUNT(*) as cnt FROM translation_memory GROUP BY target_lang;").fetchall()
            counts = {r["target_lang"]: r["cnt"] for r in rows}
            total = sum(counts.values())

    db_size_kb = 0.0
    if DB_PATH.exists():
        db_size_kb = round(DB_PATH.stat().st_size / 1024, 1)

    return {
        "counts": counts,
        "total": total,
        "db_size_kb": db_size_kb,
    }


def export_memory_markdown() -> str:
    """Tạo báo cáo Markdown từ điển song ngữ trực tiếp trong RAM (không tạo file rác)."""
    init_db()
    lines = [
        "# 📚 Bộ Nhớ Dịch Thuật Song Ngữ (SQLite Translation Memory)",
        f"*Dữ liệu trích xuất lúc {datetime.now().strftime('%H:%M:%S • %d/%m/%Y')} từ SQLite Database nội bộ (`data/pptx_translator.db`).*",
        "",
        "---",
        ""
    ]
    with _mem_lock:
        with get_connection() as conn:
            langs = [r[0] for r in conn.execute("SELECT DISTINCT target_lang FROM translation_memory;").fetchall()]
            if not langs:
                lines.append("*Bộ nhớ từ điển hiện đang trống. Hãy dịch một tệp PowerPoint để tự động học câu mới!*")
            for lang in langs:
                rows = conn.execute("""
                SELECT source_text, translated_text, used_count 
                FROM translation_memory 
                WHERE target_lang = ? 
                ORDER BY id DESC LIMIT 500;
                """, (lang,)).fetchall()

                lines.append(f"## 🌐 Ngôn ngữ: `{lang}` ({len(rows)} mục)")
                lines.append("")
                for r in rows:
                    src_clean = r["source_text"].replace("\n", " ").strip()
                    dst_clean = r["translated_text"].replace("\n", " ").strip()
                    lines.append(f"- **Gốc:** `{src_clean}`")
                    lines.append(f"  → **Dịch:** *{dst_clean}* *(Tái sử dụng: {r['used_count']} lần)*")
                lines.append("")

    return "\n".join(lines)
