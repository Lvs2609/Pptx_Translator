"""Pipeline xử lý dịch toàn diện bài thuyết trình PowerPoint."""
from __future__ import annotations

import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Callable, Optional, Protocol

from pptx import Presentation

from src.config import LANGS
from src.core.autofit import fit_text, _width
from src.core.parser import (
    STYLE_TAG_RE,
    Report,
    apply_translation,
    collect_units,
    valid,
)
from src.services.memory_service import get_cached_translations, update_cache


class Translator(Protocol):
    def translate(self, items: dict[str, str], target: str,
                  glossary: dict[str, str] | None, context: str) -> dict[str, str]: ...


def _batches(keys, max_chars, max_items):
    cur, size = [], 0
    for k in keys:
        if cur and (size + len(k) > max_chars or len(cur) >= max_items):
            yield cur
            cur, size = [], 0
        cur.append(k)
        size += len(k)
    if cur:
        yield cur


def translate_all(tr: Translator, sources: list[str], target: str, glossary, context,
                  max_chars=6000, max_items=50, workers=4, progress=None, log_func=None,
                  use_cache=True, rate_delay=0.3):
    """Xử lý dịch toàn bộ các đoạn văn bản (kết hợp bộ nhớ song ngữ và Gemini AI)."""
    result, degraded = {}, set()
    cached_count = 0

    def _log(msg):
        if log_func:
            try: log_func(msg)
            except Exception: pass

    # 1. Kiểm tra bộ nhớ từ điển trước
    to_translate = sources
    if use_cache:
        cached_dict, to_translate = get_cached_translations(sources, target)
        result.update(cached_dict)
        cached_count = len(cached_dict)
        if cached_count > 0:
            _log(f"⚡ Đã nạp {cached_count}/{len(sources)} đoạn từ SQLite TM (0s, 0 quota)! Cần dịch mới: {len(to_translate)} đoạn.")

    if not to_translate:
        _log("✨ 100% nội dung đã có trong SQLite TM! Hoàn thành tức thì.")
        return result, degraded, cached_count

    batches = list(_batches(to_translate, max_chars, max_items))
    total_batches = len(batches)
    total_items = len(to_translate)
    done_batches = 0
    done_items = 0

    _log(f"🚀 Chia làm {total_batches} gói dữ liệu (~{max_items} đoạn/gói), bắt đầu gửi Gemini AI...")

    def run_batch(batch):
        if rate_delay > 0:
            time.sleep(rate_delay)
        items = {str(i): s for i, s in enumerate(batch)}
        out = tr.translate(items, target, glossary, context)
        res = {}
        for i, s in enumerate(batch):
            d = out.get(str(i))
            if d is not None and valid(s, d):
                res[s] = d
        return res

    with ThreadPoolExecutor(workers) as ex:
        futures = {ex.submit(run_batch, b): b for b in batches}
        for fut in as_completed(futures):
            batch_res = fut.result()
            result.update(batch_res)
            done_batches += 1
            done_items += len(futures[fut])
            _log(f"📦 Hoàn thành gói {done_batches}/{total_batches} (Đã xong {done_items}/{total_items} đoạn mới)")
            if progress:
                total_all = len(sources)
                curr_done = cached_count + done_items
                pct = min(0.85, 0.10 + 0.75 * (curr_done / max(1, total_all)))
                progress(pct, f"Đang dịch AI: {curr_done}/{total_all} đoạn ({int((curr_done/total_all)*100)}%) • Gói {done_batches}/{total_batches}")

    # Xử lý lại các đoạn lỗi thẻ định dạng
    failed_sources = [s for s in to_translate if s not in result]
    if failed_sources:
        _log(f"🔧 Đang tinh chỉnh {len(failed_sources)} đoạn có cấu trúc phức tạp...")
        if progress:
            progress(0.86, f"Đang tinh chỉnh {len(failed_sources)} đoạn...")
        for idx, s in enumerate(failed_sources, 1):
            for _ in range(2):
                d = tr.translate({"0": s}, target, glossary, context).get("0")
                if d is not None and valid(s, d):
                    result[s] = d
                    break
            else:
                flat = STYLE_TAG_RE.sub("", s)
                d = tr.translate({"0": flat}, target, glossary, context).get("0") or flat
                result[s] = d if valid(flat, d, style=False) else flat
                degraded.add(s)
            if progress:
                pct = min(0.90, 0.86 + 0.04 * (idx / len(failed_sources)))
                progress(pct, f"Đang tinh chỉnh định dạng ({idx}/{len(failed_sources)})...")

    # Lưu kết quả mới vào SQLite TM
    if use_cache:
        new_translations = {s: result[s] for s in to_translate if s in result}
        update_cache(new_translations, target)
        _log(f"💾 Đã lưu {len(new_translations)} câu dịch mới vào SQLite Translation Memory.")

    return result, degraded, cached_count


def translate_pptx(src, dst, target, translator: Translator, glossary=None, fit=True,
                   min_scale=0.65, include_layouts=True,
                   progress: Callable[[float, str], None] | None = None,
                   log_func: Callable[[str], None] | None = None,
                   use_cache=True, batch_size=50, rate_delay=0.3) -> Report:
    if target not in LANGS:
        raise ValueError(f"target phải thuộc {list(LANGS)}")

    def _log(msg):
        if log_func:
            try: log_func(msg)
            except Exception: pass

    prs = Presentation(src)
    num_slides = len(prs.slides)
    _log(f"🔍 Bắt đầu phân tích {num_slides} slide trong tệp...")
    if progress:
        progress(0.04, f"Đang quét {num_slides} slide...")
    units, notes = collect_units(prs, include_layouts)
    rep = Report(units=len(units), warnings=notes)

    uniq = list(dict.fromkeys(u.tagged for u in units))
    rep.unique = len(uniq)
    _log(f"📊 Tìm thấy {len(units)} đoạn văn ({rep.unique} đoạn duy nhất sau khi lọc trùng).")
    first = " | ".join(u.plain.strip().replace("\n", " ") for u in units[:5])
    context = f"Deck bắt đầu bằng: {first[:400]}"
    if progress:
        progress(0.10, f"Đã quét {num_slides} slide ({len(units)} đoạn văn, {rep.unique} đoạn duy nhất). Đang kết nối Gemini AI...")
    trans, degraded, cached_count = translate_all(
        translator, uniq, target, glossary, context,
        max_items=batch_size, rate_delay=rate_delay, use_cache=use_cache,
        progress=progress, log_func=log_func
    )
    rep.cached = cached_count

    _log(f"🎨 Đang phục hồi màu sắc, in đậm/nghiêng và gắn font cho {len(units)} đoạn văn...")
    if progress:
        progress(0.91, f"Đang phục hồi màu sắc & định dạng cho {len(units)} đoạn văn...")
    by_owner: dict = {}
    for u in units:
        d = trans[u.tagged]
        before = _width(u.plain)
        apply_translation(u, d, target)
        if u.tagged in degraded:
            rep.degraded.append(f"{u.where}: “{u.plain[:30].strip()}…”")
        rec = by_owner.setdefault(id(u.owner), [u.owner, u.kind, u.where, 0.0, 0.0])
        rec[3] += before
        rec[4] += _width(u.translated_plain)

    if fit:
        _log(f"📐 Tự động nhận diện bố cục: Kiểm tra co chữ chống tràn cho {len(by_owner)} khung nội dung...")
        if progress:
            progress(0.95, "Đang tự động nhận diện bố cục & tối ưu kích thước chữ chống tràn...")
        for owner, kind, where, w_src, w_dst in by_owner.values():
            if kind != "notes" and w_src > 0 and w_dst / w_src > 1.2:
                ratio = w_dst / w_src
                s = fit_text(owner, kind, ratio, min_scale)
                if s is None:
                    rep.warnings.append(f"{where}: text dài ×{ratio:.1f}, không tự co được – kiểm tra tay.")
                else:
                    rep.fitted.append(f"{where} ({int(s*100)}%)")
                    if ratio ** -0.5 < min_scale:
                        rep.warnings.append(f"{where}: đã co tối đa {int(min_scale*100)}%, vẫn có thể tràn.")

    _log("💾 Đang đóng gói và lưu tệp PowerPoint hoàn chỉnh...")
    if progress:
        progress(0.98, "Đang hoàn tất đóng gói tệp PowerPoint...")
    prs.save(dst)
    _log("🎉 Hoàn tất 100%! Tệp PowerPoint đã sẵn sàng tải về máy.")
    if progress:
        progress(1.0, "Hoàn tất 100%! Tệp sẵn sàng tải về.")
    return rep

