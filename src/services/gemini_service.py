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

DOC_STYLE_GUIDES = {
    "pptx": """2. SLIDE-OPTIMIZED PRESENTATION STYLE:
   - Brevity & Impact: Slides require concise, punchy, high-impact phrasing. Avoid verbose prose or filler words.
   - Length Parity: Keep the length of the translation close to the original length to prevent text overflow from slide boxes and table cells.
   - Register: Professional, polished, executive business and technical presentation tone.
   - Case & Punctuation: Match title case, sentence case, bullet conventions, and punctuation.""",
    "docx": """2. DOCUMENT LOCALIZATION STYLE (WORD):
   - Fluency & Coherence: Translate paragraphs naturally with smooth grammar while preserving exact inline formatting tags.
   - Structural Consistency: Match headings, numbered lists, bullet conventions, form labels, and table header style.
   - Register: Clear, authoritative business and technical document prose.
   - Case & Punctuation: Preserve capitalization style and punctuation.""",
    "xlsx": """2. SPREADSHEET LOCALIZATION STYLE (EXCEL):
   - Precision & Conciseness: Translate column headers, labels, statuses, KPIs, and cell notes cleanly so spreadsheets remain compact and scannable.
   - Metric Alignment: Keep units, percentages, currency codes, ticker/account codes, and short abbreviations intact.
   - Register: Standard financial, operational, and business terminology.
   - Preserve `<br/>` line breaks exactly where multi-line cell text uses them.""",
    "pdf": """2. PDF LAYOUT-PRESERVED LOCALIZATION STYLE:
   - Fixed-Box Awareness: Keep translated text concise and proportional in length to the original block so it fits inside fixed PDF bounding boxes without overflowing.
   - Semantic Tag Alignment: Preserve inline `<rN>...</rN>` bold, italic, and color emphasis accurately.
   - Register: Clean, accurate professional publication/business tone.""",
}

SYSTEM = """You are an elite, specialized AI Localization Agent for Enterprise Documents (PowerPoint, Word, Excel, PDF Preservation Engine).
Your core mission is to translate the "text" content of every input item accurately and naturally into {target}, while preserving 100% of the underlying formatting tags, run structures, layout constraints, and technical entities.

### I. CORE ARCHITECTURAL DIRECTIVES
1. FORMATTING TAG PRESERVATION (CRITICAL):
   - Items may contain structural tags: `<r1>...</r1>`, `<r2>...</r2>` (styled text runs with distinct fonts, colors, bold/italic), `<br/>` (line breaks), and `<f1/>` (auto fields, inline images, page numbers, drawings).
   - STRICT CONSTRAINT: You MUST keep every single tag name and count EXACTLY identical to the input. Never add, delete, rename, or merge any tags.
   - SEMANTIC TAG ALIGNMENT: When target grammar requires changing word order, relocate `<rN>...</rN>` so that it encapsulates the EXACT translated phrase corresponding to the source text inside that original tag.
   - SPACING INTEGRITY: Preserve leading/trailing spaces inside or outside tags to ensure words do not concatenate unnaturally (e.g., `<r1>Hello </r1><r2>world</r2>` -> `<r1>Xin chào </r1><r2>thế giới</r2>`).

{doc_style}

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

    def _target_directive(self, target: str, source_lang: str = "auto") -> str:
        target_name = LANGS[target][0]
        if source_lang != "auto" and source_lang in LANGS:
            return f"{target_name} (Dịch từ: {LANGS[source_lang][0]})"
        return target_name

    def _generate_with_fallback(self, contents, cfg: types.GenerateContentConfig, label: str):
        """Gửi request tới Gemini có cơ chế tự động retry và chuyển model dự phòng."""
        last = None
        models_to_try = [self.model] + [m for m in FALLBACK_MODELS if m != self.model]

        for m in models_to_try:
            self.active_model = m
            for attempt in range(self.retries):
                try:
                    self.log(f"🌐 {label} ({m})...")
                    t0 = time.time()
                    resp = self.client.models.generate_content(model=m, contents=contents, config=cfg)
                    record_api_usage(m, 1)
                    elapsed = time.time() - t0
                    text = re.sub(r"^```(?:json)?|```$", "", (resp.text or "").strip(), flags=re.M).strip()
                    data = json.loads(text)
                    return m, data, elapsed
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

    def translate(self, items: dict[str, str], target: str, glossary=None, context="",
                  source_lang: str = "auto", doc_type: str = "pptx"):
        target_directive = self._target_directive(target, source_lang)
        gl = ""
        if glossary:
            gl = "GLOSSARY (always use these translations):\n" + "\n".join(f"- {k} => {v}" for k, v in glossary.items())
        doc_style = DOC_STYLE_GUIDES.get(doc_type, DOC_STYLE_GUIDES["pptx"])
        system = SYSTEM.format(
            target=target_directive,
            doc_style=doc_style,
            glossary=gl,
            context=f"CONTEXT: {context}" if context else ""
        )
        payload = json.dumps([{"id": k, "text": v} for k, v in items.items()], ensure_ascii=False)
        cfg = types.GenerateContentConfig(system_instruction=system, temperature=0.2,
                                          response_mime_type="application/json")
        m, parsed, elapsed = self._generate_with_fallback(payload, cfg, f"Đang gửi {len(items)} đoạn tới Gemini")
        res = {str(d["id"]): d["text"] for d in parsed}
        self.log(f"⚡ Model {m} dịch xong {len(res)} đoạn ({elapsed:.1f}s)")
        return res

    def translate_image_page(self, png_bytes: bytes, target: str, glossary=None,
                             source_lang: str = "auto") -> list[dict]:
        """OCR & dịch trang PDF dạng ảnh scan bằng Gemini Vision."""
        target_directive = self._target_directive(target, source_lang)
        gl = ""
        if glossary:
            gl = "\nGLOSSARY:\n" + "\n".join(f"- {k} => {v}" for k, v in glossary.items())
        prompt = (
            f"Read all visible text in this scanned document page image in natural reading order, "
            f"and translate it accurately into {target_directive}. Keep numbers, codes, URLs intact.{gl}\n"
            f'Return ONLY a JSON array of objects: [{{"type": "heading"|"text"|"table", "content": "<translated text>"}}].'
        )
        cfg = types.GenerateContentConfig(temperature=0.2, response_mime_type="application/json")
        contents = [types.Part.from_bytes(data=png_bytes, mime_type="image/png"), prompt]
        _, parsed, _ = self._generate_with_fallback(contents, cfg, "Đang OCR & dịch trang PDF scan")
        return parsed if isinstance(parsed, list) else []


def ping_test_models(api_key: str) -> list[dict]:
    """Kiểm tra độ trễ, số lượt đã dùng/còn lại và thời gian reset của từng model Gemini (Chạy song song siêu tốc)."""
    clean_key = (api_key or "").strip()
    if not clean_key:
        return []

    # Sử dụng attempts=1 để kiểm tra cực nhanh, không bị trễ retry
    http_opts = types.HttpOptions(retry_options=types.HttpRetryOptions(attempts=1))
    client = genai.Client(api_key=clean_key, http_options=http_opts)
    test_models = ["gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-3.1-flash-lite", "gemini-3.8-flash"]
    
    # Tính thời gian reset định kỳ hàng ngày (14:00 giờ Việt Nam)
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
                
                # Trích xuất số giây retryDelay nếu có
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
            elif "401" in err or "UNAUTHENTICATED" in err or "API key not valid" in err or "API_KEY_INVALID" in err:
                return {
                    "model": m,
                    "status": "UNAUTHENTICATED",
                    "ping": elapsed_ms,
                    "used": 0,
                    "remaining": 0,
                    "limit": 0,
                    "pct": 100,
                    "quota": "Sai Khóa API",
                    "usage_text": "API Key không hợp lệ (401)",
                    "state": "🔴 Sai khóa API (UNAUTHENTICATED)",
                    "reset": "Vui lòng kiểm tra lại API Key",
                    "card_class": "quota-card-err",
                    "badge_class": "err"
                }
            elif "404" in err or "NOT_FOUND" in err:
                return {
                    "model": m,
                    "status": "NOT_AVAILABLE",
                    "ping": elapsed_ms,
                    "used": 0,
                    "remaining": 0,
                    "limit": 0,
                    "pct": 100,
                    "quota": "Đã ngừng cung cấp",
                    "usage_text": "Model không khả dụng trên API này",
                    "state": "🔴 Không khả dụng (404)",
                    "reset": "Google đã thay thế bằng model mới hơn",
                    "card_class": "quota-card-err",
                    "badge_class": "err"
                }
            else:
                return {
                    "model": m,
                    "status": "ERROR",
                    "ping": elapsed_ms,
                    "used": 0,
                    "remaining": 0,
                    "limit": 0,
                    "pct": 100,
                    "quota": "Lỗi kết nối",
                    "usage_text": f"Lỗi: {err[:60]}...",
                    "state": f"🔴 Lỗi: {err[:40]}",
                    "reset": "Vui lòng kiểm tra lại đường truyền hoặc API Key",
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

    # Trả về theo đúng thứ tự ưu tiên
    return [results_map[m] for m in test_models if m in results_map]


def load_glossary(text: str) -> dict[str, str]:
    """Mỗi dòng: 'thuật ngữ = bản dịch'."""
    g = {}
    for line in text.splitlines():
        m = re.split(r"\s*(?:=>|->|=|\t)\s*", line.strip(), maxsplit=1)
        if len(m) == 2 and m[0]:
            g[m[0]] = m[1]
    return g

