"""Dịch vụ kết nối Google Gemini API với System Prompt chuyên sâu, Smart Fallback và Quota Tracker."""
from __future__ import annotations

import json
import re
import threading
import time
from datetime import datetime, timedelta
from typing import Callable, Optional

from google import genai
from google.genai import types

from src.config import LANGS, FALLBACK_MODELS, DATA_DIR
from src.services.db_service import get_connection, init_db

SYSTEM = """You are an elite, specialized AI Localization Agent for PowerPoint Presentations (PPTX XML Preservation Engine).
Your core mission is to translate the "text" content of every input item accurately and naturally into {target}, while preserving 100% of the underlying XML formatting tags, run structures, slide layout constraints, and technical entities.

### I. CORE ARCHITECTURAL DIRECTIVES
1. XML FORMATTING TAG PRESERVATION (CRITICAL):
   - Items contain structural XML tags: `<r1>...</r1>`, `<r2>...</r2>` (styled text runs with distinct fonts, colors, bold/italic), `<br/>` (line breaks), and `<f1/>` (auto fields like slide numbers, dates).
   - STRICT CONSTRAINT: You MUST keep every single tag name and count EXACTLY identical to the input. Never add, delete, rename, or merge any tags.
   - SEMANTIC TAG ALIGNMENT: When target grammar requires changing word order, relocate `<rN>...</rN>` so that it encapsulates the EXACT translated phrase corresponding to the source text inside that original tag.
   - SPACING INTEGRITY: Preserve leading/trailing spaces inside or outside tags to ensure words do not concatenate unnaturally (e.g., `<r1>Hello </r1><r2>world</r2>` -> `<r1>Xin chào </r1><r2>thế giới</r2>`).

2. SLIDE-OPTIMIZED PRESENTATION STYLE:
   - Brevity & Impact: Slides require concise, punchy, high-impact phrasing. Avoid verbose prose or filler words.
   - Length Parity: Keep the length of the translation close to the original length to prevent text overflow from slide boxes and table cells.
   - Register: Professional, polished, executive business and technical presentation tone.
   - Case & Punctuation: Match title case, sentence case, bullet conventions, and punctuation.

3. UNTRANSLATABLE ENTITIES:
   - DO NOT translate: Brand names, product names, game titles, character names, model numbers, software terms, API endpoints (e.g., "Lupin", "Paladin", "RAF", "Aster", "API", "JSON", "Token", "ID", "v1.0").
   - DO NOT translate: URLs, emails, phone numbers, file paths, numbers, percentages, units of measurement, or currency symbols ($12M, 34.5%, €50).
   - Keep all HTML entities (&lt; &gt; &amp;) completely unchanged.
   - If an item is already in {target}, return it unchanged.

4. FEW-SHOT DEMONSTRATIONS:
   Input:  {{"id": "1", "text": "<r1>Revenue Growth:</r1> Reached <r2>+34.8%</r2> in Q3."}}
   Target: Vietnamese
   Output: [{{"id": "1", "text": "<r1>Tăng trưởng doanh thu:</r1> Đạt <r2>+34.8%</r2> trong Quý 3."}}]

   Input:  {{"id": "2", "text": "<r1>Project:</r1> <r2>PRJ-2026-X99</r2><br/><f1/> Confidential."}}
   Target: Simplified Chinese
   Output: [{{"id": "2", "text": "<r1>项目：</r1><r2>PRJ-2026-X99</r2><br/><f1/> 内部绝密。"}}]

5. OUTPUT SCHEMA (ZERO DEVIATION):
   - Output MUST be ONLY a valid raw JSON array of objects:
     [
       {{"id": "<same input id>", "text": "<translated string with exact tags>"}},
       ...
     ]
   - Return every single input item. Do NOT wrap in markdown fences (```json ... ```). No explanations or conversational text.
{glossary}
{context}"""


_usage_lock = threading.Lock()

MODEL_LIMITS_NUM = {
    "gemini-3.5-flash": 20,       # Google AI Studio cấp hạn mức 20 lượt/ngày trên gói Free
    "gemini-3.5-flash-lite": 1500, # Gói 1,500 lượt/ngày tiêu chuẩn
    "gemini-3.1-flash-lite": 1500, # Gói 1,500 lượt/ngày tiêu chuẩn
    "gemini-3.8-flash": 20,       # Preview 20 lượt/ngày
}

def _get_cycle_key() -> str:
    """Trả về key chu kỳ quota: Google AI Studio reset lúc 14:00 VN hàng ngày."""
    now = datetime.now()
    reset_today = now.replace(hour=14, minute=0, second=0, microsecond=0)
    if now < reset_today:
        return (now - timedelta(days=1)).strftime("%Y-%m-%d")
    return now.strftime("%Y-%m-%d")

def get_usage_stats(model: str) -> tuple[int, int, int]:
    """Trả về (used, remaining, limit) cho model từ SQLite DB."""
    init_db()
    cycle = _get_cycle_key()
    limit = MODEL_LIMITS_NUM.get(model, 1500)
    with _usage_lock:
        try:
            with get_connection() as conn:
                row = conn.execute(
                    "SELECT request_count, is_exhausted FROM api_usage WHERE cycle_date = ? AND model_name = ?;",
                    (cycle, model)
                ).fetchone()
                if not row:
                    return 0, limit, limit
                used = row["request_count"]
                if row["is_exhausted"]:
                    used = limit
                remaining = max(0, limit - used)
                return used, remaining, limit
        except Exception:
            return 0, limit, limit

def record_api_usage(model: str, count: int = 1, is_429: bool = False):
    """Ghi nhận lượt gọi API hoặc đánh dấu 429 hết quota vào SQLite DB."""
    init_db()
    cycle = _get_cycle_key()
    limit = MODEL_LIMITS_NUM.get(model, 1500)
    now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _usage_lock:
        try:
            with get_connection() as conn:
                row = conn.execute(
                    "SELECT request_count FROM api_usage WHERE cycle_date = ? AND model_name = ?;",
                    (cycle, model)
                ).fetchone()
                if row:
                    new_count = limit if is_429 else min(limit, row["request_count"] + count)
                    exhausted = 1 if (is_429 or new_count >= limit) else 0
                    conn.execute("""
                    UPDATE api_usage 
                    SET request_count = ?, is_exhausted = ?, updated_at = ?
                    WHERE cycle_date = ? AND model_name = ?;
                    """, (new_count, exhausted, now_iso, cycle, model))
                else:
                    new_count = limit if is_429 else count
                    exhausted = 1 if is_429 else 0
                    conn.execute("""
                    INSERT INTO api_usage (cycle_date, model_name, request_count, is_exhausted, updated_at)
                    VALUES (?, ?, ?, ?, ?);
                    """, (cycle, model, new_count, exhausted, now_iso))
                conn.commit()
        except Exception:
            pass


class GeminiTranslator:
    def __init__(self, api_key: str, model: str = "gemini-3.5-flash", retries: int = 3,
                 log_func: Optional[Callable[[str], None]] = None):
        self.client = genai.Client(api_key=api_key)
        self.model = model
        self.active_model = model  # Theo dõi chính xác model đang thực tế xử lý
        self.retries = retries
        self.log_func = log_func or (lambda m: None)

    def log(self, msg: str):
        try:
            self.log_func(msg)
        except Exception:
            pass

    def translate(self, items: dict[str, str], target: str, glossary=None, context="", source_lang: str = "auto"):
        target_name = LANGS[target][0]
        if source_lang != "auto" and source_lang in LANGS:
            src_name = LANGS[source_lang][0]
            target_directive = f"{target_name} (Dịch từ: {src_name})"
        else:
            target_directive = target_name

        gl = ""
        if glossary:
            gl = "GLOSSARY (always use these translations):\n" + "\n".join(f"- {k} => {v}" for k, v in glossary.items())
        system = SYSTEM.format(target=target_directive, glossary=gl, context=f"CONTEXT: {context}" if context else "")
        payload = json.dumps([{"id": k, "text": v} for k, v in items.items()], ensure_ascii=False)
        cfg = types.GenerateContentConfig(system_instruction=system, temperature=0.2,
                                          response_mime_type="application/json")
        last = None
        models_to_try = [self.model] + [m for m in FALLBACK_MODELS if m != self.model]

        for m in models_to_try:
            self.active_model = m
            for attempt in range(self.retries):
                try:
                    self.log(f"🌐 Đang gửi {len(items)} đoạn tới Gemini ({m})...")
                    t0 = time.time()
                    resp = self.client.models.generate_content(model=m, contents=payload, config=cfg)
                    record_api_usage(m, 1)
                    elapsed = time.time() - t0
                    text = re.sub(r"^```(?:json)?|```$", "", (resp.text or "").strip(), flags=re.M).strip()
                    res = {str(d["id"]): d["text"] for d in json.loads(text)}
                    self.log(f"⚡ Model {m} dịch xong {len(res)} đoạn ({elapsed:.1f}s)")
                    return res
                except Exception as e:
                    last = e
                    err_str = str(e)
                    if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                        record_api_usage(m, is_429=True)
                        self.log(f"⚠️ Model {m} chạm hạn mức (429)! Tự động chuyển model dự phòng...")
                        break
                    elif "404" in err_str or "NOT_FOUND" in err_str:
                        self.log(f"⚠️ Model {m} không khả dụng (404), thử model khác...")
                        break
                    else:
                        self.log(f"🔄 Đang thử lại {m} (Lần {attempt+1}/{self.retries})...")
                        time.sleep(1.5 ** attempt)
        raise RuntimeError(f"Gemini lỗi sau khi thử các model: {last}")


def ping_test_models(api_key: str) -> list[dict]:
    """Kiểm tra khóa API và các model Gemini siêu tốc (< 0.5s)."""
    clean_key = (api_key or "").strip()
    if not clean_key:
        return []

    # 1. Kiểm tra nhanh định dạng & độ dài: Khóa Google AI Studio chuẩn luôn bắt đầu bằng 'AIza' và dài ít nhất 30 ký tự
    if not clean_key.startswith("AIza") or len(clean_key) < 30:
        return [{
            "model": "All",
            "status": "UNAUTHENTICATED",
            "error_message": "Khóa API không đúng định dạng. Khóa Google AI Studio chuẩn phải bắt đầu bằng 'AIzaSy...' và có 39 ký tự. Vui lòng lấy lại khóa tại aistudio.google.com/apikey."
        }]

    # 2. Tạo client với attempts=1 để không bao giờ bị trễ retry
    http_opts = types.HttpOptions(retry_options=types.HttpRetryOptions(attempts=1))
    client = genai.Client(api_key=clean_key, http_options=http_opts)
    test_models = ["gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-3.1-flash-lite", "gemini-3.8-flash"]

    # 3. Probe nhanh trên 1 model với 1 token duy nhất để xác thực khóa trong ~0.2s
    # Nếu sai khóa thì dừng ngay lập tức, không ping các model còn lại
    try:
        client.models.generate_content(
            model="gemini-3.5-flash-lite",
            contents="ping",
            config=types.GenerateContentConfig(max_output_tokens=1)
        )
    except Exception as e:
        err = str(e)
        if any(kw in err for kw in ["API_KEY_INVALID", "API key not valid", "400", "401", "UNAUTHENTICATED", "PERMISSION_DENIED"]):
            return [{
                "model": "All",
                "status": "UNAUTHENTICATED",
                "error_message": "Khóa API không hợp lệ hoặc đã bị vô hiệu hóa trên Google AI Studio. Vui lòng kiểm tra lại khóa tại aistudio.google.com/apikey (bắt đầu bằng AIzaSy...)."
            }]

    # 4. Tính thời gian reset định kỳ hàng ngày (14:00 giờ Việt Nam)
    now = datetime.now()
    reset_today = now.replace(hour=14, minute=0, second=0, microsecond=0)
    if now >= reset_today:
        reset_target = reset_today + timedelta(days=1)
    else:
        reset_target = reset_today
    diff = reset_target - now
    hours, remainder = divmod(int(diff.total_seconds()), 3600)
    minutes, _ = divmod(remainder, 60)
    default_reset_str = f"Tự động làm mới sau ~{hours}h{minutes:02d}m (lúc 14:00 hàng ngày)"

    def check_model(m: str) -> dict:
        t0 = time.time()
        try:
            resp = client.models.generate_content(
                model=m,
                contents="ping",
                config=types.GenerateContentConfig(max_output_tokens=1)
            )
            elapsed_ms = int((time.time() - t0) * 1000)
            used, remaining, limit = get_usage_stats(m)
            pct = min(100, int((used / limit) * 100)) if limit else 0
            
            return {
                "model": m,
                "status": "OK",
                "ping": elapsed_ms,
                "used": used,
                "remaining": remaining,
                "limit": limit,
                "pct": pct,
                "quota": f"{limit:,} lượt/ngày • 15 lượt/phút",
                "usage_text": f"Đã dùng: <b>{used}</b>/{limit} lượt • Còn lại: <b>{remaining:,}</b> lượt",
                "state": "🟢 Sẵn sàng (Còn quota)",
                "reset": default_reset_str,
                "card_class": "quota-card-ok",
                "badge_class": "ready"
            }
        except Exception as e:
            elapsed_ms = int((time.time() - t0) * 1000)
            err = str(e)
            if "429" in err or "RESOURCE_EXHAUSTED" in err:
                status = "QUOTA_EXCEEDED"
                limit_match = re.search(r"limit:\s*(\d+)", err) or re.search(r"'quotaValue':\s*'(\d+)'", err)
                limit = int(limit_match.group(1)) if limit_match else MODEL_LIMITS_NUM.get(m, 20)
                record_api_usage(m, is_429=True)
                
                sec_match = re.search(r"'retryDelay': '(\d+)s'", err)
                retry_match = re.search(r"retry in ([0-9]+h[0-9]+m|[0-9]+m[0-9]+s|[0-9]+h|[0-9]+m|[0-9]+s|\d+\.?\d*s)", err, re.I)
                if sec_match:
                    secs = int(sec_match.group(1))
                    h, rem = divmod(secs, 3600)
                    minu, _ = divmod(rem, 60)
                    reset_time_str = (datetime.now() + timedelta(seconds=secs)).strftime("%H:%M:%S")
                    reset_msg = f"Mở lại sau: {h} giờ {minu} phút (vào lúc {reset_time_str})"
                elif retry_match:
                    reset_msg = f"Google báo mở lại sau: {retry_match.group(1)}"
                else:
                    reset_msg = default_reset_str
                    
                return {
                    "model": m,
                    "status": status,
                    "ping": elapsed_ms,
                    "used": limit,
                    "remaining": 0,
                    "limit": limit,
                    "pct": 100,
                    "quota": f"{limit:,} lượt/ngày (Đã hết)",
                    "usage_text": f"Đã dùng: <b>{limit}</b>/{limit} (Hết quota) • Còn lại: <b>0</b> lượt",
                    "state": "🟡 Hết quota trong ngày (429)",
                    "reset": reset_msg,
                    "card_class": "quota-card-warn",
                    "badge_class": "warn"
                }
            else:
                return {
                    "model": m,
                    "status": "ERROR",
                    "ping": elapsed_ms,
                    "used": 0,
                    "remaining": 0,
                    "limit": 0,
                    "pct": 0,
                    "quota": "Lỗi kết nối",
                    "usage_text": "Không thể kiểm tra hạn mức",
                    "state": f"🔴 Lỗi: {e.__class__.__name__}",
                    "reset": "Kiểm tra lại kết nối mạng hoặc API key",
                    "card_class": "quota-card-err",
                    "badge_class": "err"
                }

    # Chạy song song cả 4 model với ThreadPoolExecutor (nhanh vượt trội)
    from concurrent.futures import ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {executor.submit(check_model, m): m for m in test_models}
        results_map = {}
        for future in futures:
            m = futures[future]
            try:
                results_map[m] = future.result()
            except Exception:
                results_map[m] = {
                    "model": m,
                    "status": "ERROR",
                    "ping": 0,
                    "used": 0,
                    "remaining": 0,
                    "limit": 0,
                    "pct": 0,
                    "quota": "Lỗi kiểm tra",
                    "usage_text": "Timeout hoặc lỗi kết nối",
                    "state": "🔴 Lỗi kết nối",
                    "reset": "Vui lòng thử lại",
                    "card_class": "quota-card-err",
                    "badge_class": "err"
                }

    return [results_map[m] for m in test_models if m in results_map]


def load_glossary(text: str) -> dict[str, str]:
    """Mỗi dòng: 'thuật ngữ = bản dịch'."""
    g = {}
    for line in text.splitlines():
        m = re.split(r"\s*(?:=>|->|=|\t)\s*", line.strip(), maxsplit=1)
        if len(m) == 2 and m[0]:
            g[m[0]] = m[1]
    return g

