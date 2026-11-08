"""Dịch vụ lưu trữ cơ sở dữ liệu SQLite cho PPTX Translator.
Quản trị tập trung 100% dữ liệu (Tasks History, Translation Memory, API Quota Usage).
Hoàn toàn lưu trữ cục bộ (Local-only) trong máy người dùng, hiệu năng cao với WAL mode.
"""
from __future__ import annotations

import json
import os
import sqlite3
import threading
import uuid
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

from src.config import DATA_DIR, MEMORY_JSON, MEMORY_MD

DB_PATH = DATA_DIR / "pptx_translator.db"
EXPORTS_DIR = DATA_DIR / "exports"
EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
USAGE_JSON = DATA_DIR / "daily_usage.json"

_db_lock = threading.RLock()
MAX_EXPORTS_RETAINED = 20  # Tự động lưu tối đa 20 file gần nhất để bảo vệ ổ cứng


def get_connection() -> sqlite3.Connection:
    """Tạo kết nối tới SQLite với WAL mode để đọc/ghi đồng thời cực nhanh và an toàn."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA synchronous = NORMAL;")
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Khởi tạo toàn bộ cấu trúc bảng và tự động di chuyển dữ liệu cũ từ JSON vào SQLite."""
    with _db_lock:
        with get_connection() as conn:
            # 1. Bảng lịch sử tác vụ dịch file PPTX
            conn.execute("""
            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                task_id TEXT UNIQUE NOT NULL,
                source_file TEXT NOT NULL,
                translated_file TEXT NOT NULL,
                target_lang TEXT NOT NULL,
                slides_count INTEGER DEFAULT 0,
                units_count INTEGER DEFAULT 0,
                file_size_kb REAL DEFAULT 0,
                status TEXT DEFAULT 'SUCCESS',
                file_path TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            """)

            # 2. Bảng bộ nhớ song ngữ (Translation Memory - TM)
            conn.execute("""
            CREATE TABLE IF NOT EXISTS translation_memory (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                target_lang TEXT NOT NULL,
                source_text TEXT NOT NULL,
                translated_text TEXT NOT NULL,
                used_count INTEGER DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(target_lang, source_text)
            );
            """)
            conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_tm_lookup ON translation_memory(target_lang, source_text);
            """)

            # 3. Bảng theo dõi hạn mức Quota API theo chu kỳ ngày
            conn.execute("""
            CREATE TABLE IF NOT EXISTS api_usage (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cycle_date TEXT NOT NULL,
                model_name TEXT NOT NULL,
                request_count INTEGER DEFAULT 0,
                is_exhausted INTEGER DEFAULT 0,
                updated_at TEXT NOT NULL,
                UNIQUE(cycle_date, model_name)
            );
            """)
            conn.commit()

        # Tự động di chuyển dữ liệu cũ từ JSON/MD sang SQLite nếu có
        _migrate_legacy_data_if_needed()


def _migrate_legacy_data_if_needed():
    """Chuyển đổi các file JSON/MD cũ vào SQLite rồi tự động dọn dẹp sạch sẽ."""
    # Di chuyển translation_memory.json
    if MEMORY_JSON.exists():
        try:
            with open(MEMORY_JSON, "r", encoding="utf-8") as f:
                old_mem = json.load(f)
            now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            rows = []
            for lang, entries in old_mem.items():
                if isinstance(entries, dict):
                    for s, d in entries.items():
                        if s and d:
                            rows.append((lang, s.strip(), d.strip(), now_iso, now_iso))
            if rows:
                with get_connection() as conn:
                    conn.executemany("""
                    INSERT INTO translation_memory (target_lang, source_text, translated_text, used_count, created_at, updated_at)
                    VALUES (?, ?, ?, 1, ?, ?)
                    ON CONFLICT(target_lang, source_text) DO UPDATE SET
                        translated_text = excluded.translated_text,
                        updated_at = excluded.updated_at;
                    """, rows)
                    conn.commit()
            # Xóa các file cũ để không lưu trữ thừa thãi trên máy
            MEMORY_JSON.unlink(missing_ok=True)
            if MEMORY_MD.exists():
                MEMORY_MD.unlink(missing_ok=True)
        except Exception:
            pass

    # Di chuyển daily_usage.json
    if USAGE_JSON.exists():
        try:
            with open(USAGE_JSON, "r", encoding="utf-8") as f:
                old_usage = json.load(f)
            cycle = old_usage.get("cycle")
            models = old_usage.get("models", {})
            if cycle and models:
                now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with get_connection() as conn:
                    for m, cnt in models.items():
                        conn.execute("""
                        INSERT INTO api_usage (cycle_date, model_name, request_count, is_exhausted, updated_at)
                        VALUES (?, ?, ?, 0, ?)
                        ON CONFLICT(cycle_date, model_name) DO UPDATE SET
                            request_count = excluded.request_count,
                            updated_at = excluded.updated_at;
                        """, (cycle, m, cnt, now_iso))
                    conn.commit()
            USAGE_JSON.unlink(missing_ok=True)
        except Exception:
            pass


# ----------------- QUẢN LÝ TASK DỊCH VÀ AUTO-RETENTION -----------------

def prune_old_exports(max_retained: int = MAX_EXPORTS_RETAINED):
    """Tự động xóa các file vật lý xuất ra cũ hơn 20 file gần nhất để bảo vệ dung lượng ổ cứng."""
    with _db_lock:
        with get_connection() as conn:
            rows = conn.execute("SELECT id, file_path FROM tasks ORDER BY id DESC;").fetchall()
            if len(rows) > max_retained:
                excess_rows = rows[max_retained:]
                for r in excess_rows:
                    try:
                        p = Path(r["file_path"])
                        if p.exists():
                            p.unlink()
                    except Exception:
                        pass
                excess_ids = [r["id"] for r in excess_rows]
                placeholders = ",".join("?" for _ in excess_ids)
                conn.execute(f"DELETE FROM tasks WHERE id IN ({placeholders});", excess_ids)
                conn.commit()


def save_translated_task(
    source_name: str,
    translated_name: str,
    target_lang: str,
    slides_count: int,
    units_count: int,
    file_bytes: bytes,
) -> dict:
    """Lưu tệp dịch ra thư mục exports và ghi nhận vào cơ sở dữ liệu SQLite."""
    init_db()
    task_id = str(uuid.uuid4())[:8]
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Đặt tên file xuất an toàn chống trùng lặp
    export_filename = f"{Path(translated_name).stem}_{timestamp_str}{Path(translated_name).suffix}"
    export_path = EXPORTS_DIR / export_filename
    
    with open(export_path, "wb") as f:
        f.write(file_bytes)
        
    size_kb = round(len(file_bytes) / 1024, 1)
    now_iso = datetime.now().strftime("%H:%M • %d/%m/%Y")
    
    with _db_lock:
        with get_connection() as conn:
            conn.execute("""
            INSERT INTO tasks (
                task_id, source_file, translated_file, target_lang,
                slides_count, units_count, file_size_kb, status, file_path, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, 'SUCCESS', ?, ?);
            """, (
                task_id, source_name, translated_name, target_lang,
                slides_count, units_count, size_kb, str(export_path), now_iso
            ))
            conn.commit()
            
        # Áp dụng cơ chế dọn dẹp xoay vòng tự động (giữ 20 file gần nhất)
        prune_old_exports()
            
    return {
        "task_id": task_id,
        "source_file": source_name,
        "translated_file": translated_name,
        "target_lang": target_lang,
        "slides_count": slides_count,
        "units_count": units_count,
        "file_size_kb": size_kb,
        "file_path": str(export_path),
        "created_at": now_iso
    }


def get_recent_tasks(limit: int = 30) -> list[dict]:
    """Lấy danh sách các tệp đã dịch từ SQLite."""
    init_db()
    with _db_lock:
        with get_connection() as conn:
            cursor = conn.execute("""
            SELECT id, task_id, source_file, translated_file, target_lang,
                   slides_count, units_count, file_size_kb, status, file_path, created_at
            FROM tasks
            ORDER BY id DESC
            LIMIT ?;
            """, (limit,))
            rows = cursor.fetchall()
            
            tasks = []
            for r in rows:
                p = Path(r["file_path"])
                tasks.append({
                    "id": r["id"],
                    "task_id": r["task_id"],
                    "source_file": r["source_file"],
                    "translated_file": r["translated_file"],
                    "target_lang": r["target_lang"],
                    "slides_count": r["slides_count"],
                    "units_count": r["units_count"],
                    "file_size_kb": r["file_size_kb"],
                    "status": r["status"],
                    "file_path": r["file_path"],
                    "file_exists": p.exists(),
                    "created_at": r["created_at"],
                })
            return tasks


def delete_task(task_id: str):
    """Xóa 1 bản ghi và tệp xuất tương ứng."""
    with _db_lock:
        with get_connection() as conn:
            row = conn.execute("SELECT file_path FROM tasks WHERE task_id = ?;", (task_id,)).fetchone()
            if row and row["file_path"]:
                try:
                    p = Path(row["file_path"])
                    if p.exists():
                        p.unlink()
                except Exception:
                    pass
            conn.execute("DELETE FROM tasks WHERE task_id = ?;", (task_id,))
            conn.commit()


def clear_all_tasks():
    """Dọn sạch lịch sử tác vụ và các file đã xuất."""
    with _db_lock:
        with get_connection() as conn:
            rows = conn.execute("SELECT file_path FROM tasks;").fetchall()
            for r in rows:
                try:
                    p = Path(r["file_path"])
                    if p.exists():
                        p.unlink()
                except Exception:
                    pass
            conn.execute("DELETE FROM tasks;")
            conn.commit()
