"""Bóc tách cấu trúc văn bản PDF và dịch thuật giữ nguyên bố cục, bảng biểu, hình ảnh (PyMuPDF)."""
from __future__ import annotations

import html
import re
from dataclasses import dataclass, field
from typing import Callable, Optional

import pymupdf

from src.config import LANGS
from src.core.parser import STYLE_TAG_RE, TAG_RE, Report, needs_translation_text
from src.core.translator import Translator, translate_all


@dataclass
class PdfSpanStyle:
    bold: bool
    italic: bool
    color: str
    size: float

    def sig(self) -> tuple:
        return (self.bold, self.italic, self.color, self.size)


@dataclass
class PdfUnit:
    """Một khối/đoạn văn bản PDF kèm tọa độ bounding box chính xác."""
    page_idx: int
    where: str
    bbox: tuple[float, float, float, float]
    span_rects: list[pymupdf.Rect] = field(default_factory=list)
    styles_by_id: dict[int, PdfSpanStyle] = field(default_factory=dict)
    default_id: int = 1
    avg_size: float = 10.5
    rotate: int = 0
    use_tags: bool = False
    tagged: str = ""
    plain: str = ""


def _dir_to_rotate(dir_vec) -> int:
    dx, dy = float(dir_vec[0]), float(dir_vec[1])
    if abs(dx - 1.0) < 0.25 and abs(dy) < 0.25:
        return 0
    if abs(dx) < 0.25 and abs(dy + 1.0) < 0.25:
        return 90
    if abs(dx + 1.0) < 0.25 and abs(dy) < 0.25:
        return 180
    if abs(dx) < 0.25 and abs(dy - 1.0) < 0.25:
        return 270
    return 0


def _span_style(span: dict) -> PdfSpanStyle:
    flags = int(span.get("flags", 0))
    font_lower = (span.get("font") or "").lower()
    bold = bool(flags & 16) or ("bold" in font_lower) or ("black" in font_lower) or ("heavy" in font_lower)
    italic = bool(flags & 2) or ("italic" in font_lower) or ("oblique" in font_lower)
    color_int = int(span.get("color", 0)) & 0xFFFFFF
    color_hex = f"#{color_int:06x}"
    raw_size = float(span.get("size", 10.0))
    # Làm tròn về bội 0.5pt để tránh tách run thừa vì sai số dấu phẩy động
    size = max(5.0, round(raw_size * 2.0) / 2.0)
    return PdfSpanStyle(bold=bold, italic=italic, color=color_hex, size=size)


def _split_line_columns(line: dict) -> list[dict]:
    """Tách các span trên cùng 1 dòng nếu có khoảng trống ngang lớn (ô bảng / nhiều cột)."""
    spans = [s for s in line.get("spans", []) if s.get("text")]
    if not spans:
        return []

    spans.sort(key=lambda s: s["bbox"][0])
    groups: list[list[dict]] = [[spans[0]]]
    for sp in spans[1:]:
        prev = groups[-1][-1]
        gap = float(sp["bbox"][0]) - float(prev["bbox"][2])
        ref_size = max(float(prev.get("size", 10.0)), float(sp.get("size", 10.0)), 6.0)
        if gap > max(14.0, ref_size * 2.2):
            groups.append([sp])
        else:
            groups[-1].append(sp)

    segments = []
    for grp in groups:
        x0 = min(s["bbox"][0] for s in grp)
        y0 = min(s["bbox"][1] for s in grp)
        x1 = max(s["bbox"][2] for s in grp)
        y1 = max(s["bbox"][3] for s in grp)
        segments.append({
            "spans": grp,
            "bbox": (x0, y0, x1, y1),
            "dir": line.get("dir", (1.0, 0.0)),
            "has_split": len(groups) > 1,
        })
    return segments


def _needs_join_space(prev_text: str, next_text: str) -> bool:
    if not prev_text or not next_text:
        return False
    if prev_text[-1].isspace() or next_text[0].isspace():
        return False
    # Không chèn dấu cách giữa hai ký tự Đông Á (CJK)
    if ord(prev_text[-1]) >= 0x2E80 and ord(next_text[0]) >= 0x2E80:
        return False
    return True


def _build_pdf_unit(page_idx: int, where: str, segments: list[dict]) -> Optional[PdfUnit]:
    """Gom nhóm các segment dòng thành một PdfUnit hoàn chỉnh kèm thẻ <rN>."""
    parts: list[list] = []
    style_by_sig: dict[tuple, PdfSpanStyle] = {}
    span_rects: list[pymupdf.Rect] = []
    sizes: list[float] = []

    rot = _dir_to_rotate(segments[0]["dir"])
    x0 = min(seg["bbox"][0] for seg in segments)
    y0 = min(seg["bbox"][1] for seg in segments)
    x1 = max(seg["bbox"][2] for seg in segments)
    y1 = max(seg["bbox"][3] for seg in segments)

    prev_line_text = ""
    for seg_idx, seg in enumerate(segments):
        line_buf = ""
        for sp in seg["spans"]:
            txt = sp.get("text", "")
            if not txt:
                continue
            span_rects.append(pymupdf.Rect(sp["bbox"]))
            st = _span_style(sp)
            sizes.append(st.size)
            sig = st.sig()
            style_by_sig.setdefault(sig, st)

            prefix = ""
            if seg_idx > 0 and not line_buf and _needs_join_space(prev_line_text, txt):
                prefix = " "
            line_buf += txt
            combined = prefix + txt

            if parts and parts[-1][1] == sig:
                parts[-1][0] += combined
            else:
                parts.append([combined, sig])
        if line_buf:
            prev_line_text = line_buf

    if not parts or not span_rects:
        return None

    # Gom khoảng trắng đứng lẻ vào kiểu định dạng lân cận
    for i, p in enumerate(parts):
        if not p[0].strip():
            nxt = next((y for y in parts[i + 1:] if y[0].strip()), None)
            prv = next((y for y in reversed(parts[:i]) if y[0].strip()), None)
            ref = nxt or prv
            if ref:
                p[1] = ref[1]

    sig_to_id: dict[tuple, int] = {}
    styles_by_id: dict[int, PdfSpanStyle] = {}
    for _txt, sig in parts:
        if sig not in sig_to_id:
            rid = len(sig_to_id) + 1
            sig_to_id[sig] = rid
            styles_by_id[rid] = style_by_sig[sig]

    use_tags = len(sig_to_id) > 1
    out, plain = [], []
    longest: dict[int, int] = {}
    for txt, sig in parts:
        rid = sig_to_id[sig]
        longest[rid] = longest.get(rid, 0) + len(txt)
        plain.append(txt)
        esc = html.escape(txt, quote=False)
        out.append(f"<r{rid}>{esc}</r{rid}>" if use_tags else esc)

    plain_str = "".join(plain)
    if not needs_translation_text(plain_str):
        return None

    return PdfUnit(
        page_idx=page_idx,
        where=where,
        bbox=(x0, y0, x1, y1),
        span_rects=span_rects,
        styles_by_id=styles_by_id,
        default_id=max(longest, key=longest.get),
        avg_size=(sum(sizes) / len(sizes)) if sizes else 10.5,
        rotate=rot,
        use_tags=use_tags,
        tagged="".join(out),
        plain=plain_str,
    )


def _collect_pdf_units(doc: pymupdf.Document) -> dict[int, list[PdfUnit]]:
    """Quét tất cả các trang PDF, tách bảng/cột và gom đoạn văn bảo toàn vị trí."""
    units_by_page: dict[int, list[PdfUnit]] = {}
    flags = pymupdf.TEXTFLAGS_DICT & ~pymupdf.TEXT_DEHYPHENATE

    for page_idx in range(len(doc)):
        page = doc[page_idx]
        p_dict = page.get_text("dict", flags=flags)
        page_units: list[PdfUnit] = []
        unit_counter = 0

        for block in p_dict.get("blocks", []):
            if block.get("type") != 0:
                continue

            line_groups: list[list[dict]] = []
            for line in block.get("lines", []):
                segs = _split_line_columns(line)
                if not segs:
                    continue
                # Nếu dòng bị tách nhiều cột ngang -> mỗi cột là 1 unit riêng biệt
                if len(segs) > 1 or _dir_to_rotate(line.get("dir", (1, 0))) != 0:
                    for seg in segs:
                        unit_counter += 1
                        u = _build_pdf_unit(page_idx, f"Trang {page_idx + 1}", [seg])
                        if u:
                            page_units.append(u)
                    line_groups.append([])  # Ngắt nhóm đoạn
                    continue

                seg = segs[0]
                if line_groups and line_groups[-1]:
                    prev_seg = line_groups[-1][-1]
                    prev_y1 = prev_seg["bbox"][3]
                    curr_y0 = seg["bbox"][1]
                    prev_h = max(8.0, prev_y1 - prev_seg["bbox"][1])
                    v_gap = curr_y0 - prev_y1
                    x_aligned = abs(seg["bbox"][0] - prev_seg["bbox"][0]) <= 28.0
                    # Chỉ gộp nếu nằm gần nhau theo chiều dọc và thẳng lề
                    if -2.0 <= v_gap <= prev_h * 0.85 and x_aligned:
                        line_groups[-1].append(seg)
                        continue
                line_groups.append([seg])

            for grp in line_groups:
                if not grp:
                    continue
                unit_counter += 1
                u = _build_pdf_unit(page_idx, f"Trang {page_idx + 1}", grp)
                if u:
                    page_units.append(u)

        if page_units:
            units_by_page[page_idx] = page_units

    return units_by_page


def _unit_to_html(unit: PdfUnit, translated: str) -> str:
    """Chuyển chuỗi đã dịch (kèm thẻ <rN>, <br/>) sang HTML để PyMuPDF render Unicode & style."""
    cur = unit.default_id
    chunks: list[str] = []

    def _span_html(style: Optional[PdfSpanStyle], text: str) -> str:
        safe = html.escape(text, quote=True)
        if not style:
            return safe
        weight = "bold" if style.bold else "normal"
        fstyle = "italic" if style.italic else "normal"
        return (
            f'<span style="font-weight:{weight}; font-style:{fstyle}; '
            f'color:{style.color}; font-size:{style.size:.1f}pt;">{safe}</span>'
        )

    for tok in TAG_RE.split(translated):
        if not tok:
            continue
        if m := re.fullmatch(r"<r(\d+)>", tok):
            rid = int(m.group(1))
            cur = rid if rid in unit.styles_by_id else unit.default_id
        elif re.fullmatch(r"</r\d+>", tok):
            cur = unit.default_id
        elif tok == "<br/>":
            chunks.append("<br/>")
        elif re.fullmatch(r"<f\d+/>", tok):
            continue
        else:
            chunks.append(_span_html(unit.styles_by_id.get(cur), html.unescape(tok)))

    def_style = unit.styles_by_id.get(unit.default_id)
    base_sz = def_style.size if def_style else unit.avg_size
    return (
        f'<div style="font-family: sans-serif; font-size: {base_sz:.1f}pt; '
        f'line-height: 1.18; margin: 0; padding: 0;">{"".join(chunks)}</div>'
    )


def _compute_safe_rect(unit: PdfUnit, all_units: list[PdfUnit], page_rect: pymupdf.Rect) -> pymupdf.Rect:
    """Mở rộng nhẹ vùng chèn chữ nếu không chạm khối lân cận, giúp câu dịch không bị cắt."""
    x0, y0, x1, y1 = unit.bbox
    if unit.rotate != 0:
        return pymupdf.Rect(x0, y0, x1, y1)

    # Tìm giới hạn y1 gần nhất phía dưới có chồng lấn theo chiều ngang
    max_y1 = page_rect.y1 - 10.0
    max_x1 = page_rect.x1 - 10.0
    for other in all_units:
        if other is unit:
            continue
        ox0, oy0, ox1, oy1 = other.bbox
        # Kiểm tra chồng lấn ngang cho mở rộng dọc
        if min(x1, ox1) > max(x0, ox0) - 4.0 and oy0 >= y1 - 2.0:
            max_y1 = min(max_y1, oy0 - 1.5)
        # Kiểm tra chồng lấn dọc cho mở rộng ngang nhẹ
        if min(y1, oy1) > max(y0, oy0) and ox0 >= x1 - 2.0:
            max_x1 = min(max_x1, ox0 - 2.0)

    height = y1 - y0
    safe_y1 = max(y1 + 2.0, min(max_y1, y0 + max(height * 1.65, unit.avg_size * 1.9)))
    extra_w = min(24.0, max(6.0, (x1 - x0) * 0.2))
    safe_x1 = min(max_x1, x1 + extra_w)
    return pymupdf.Rect(max(page_rect.x0, x0), max(page_rect.y0, y0 - 0.5), safe_x1, safe_y1)


def _translate_scanned_pdf(doc: pymupdf.Document, dst, target: str, translator: Translator,
                           source_lang: str, glossary, progress, _log) -> Report:
    """Fallback thông minh bằng Gemini Vision OCR khi gặp PDF dạng ảnh scan."""
    total_pages = len(doc)
    out_doc = pymupdf.open()
    rep = Report(units=total_pages, unique=total_pages)
    rep.warnings.append("PDF dạng ảnh scan (không có văn bản gốc) — đã áp dụng Gemini Vision OCR.")
    _log(f"📸 Phát hiện PDF dạng ảnh scan ({total_pages} trang) — AI đang OCR và dịch từng trang...")

    for i in range(total_pages):
        src_page = doc[i]
        rect = src_page.rect
        new_page = out_doc.new_page(width=rect.width, height=rect.height)
        if hasattr(translator, "translate_image_page"):
            png = src_page.get_pixmap(dpi=150).tobytes("png")
            blocks = translator.translate_image_page(png, target, glossary=glossary, source_lang=source_lang)
            html_parts = []
            for b in blocks:
                content = html.escape(str(b.get("content", "")).strip())
                if not content:
                    continue
                if b.get("type") == "heading":
                    html_parts.append(f'<p style="font-size:13pt; font-weight:bold; margin:6px 0 4px 0;">{content}</p>')
                else:
                    html_parts.append(f'<p style="font-size:10.5pt; margin:0 0 5px 0;">{content}</p>')
            if html_parts:
                box = pymupdf.Rect(36, 36, rect.width - 36, rect.height - 36)
                html_doc = f'<div style="font-family:sans-serif; line-height:1.35;">{"".join(html_parts)}</div>'
                new_page.insert_htmlbox(box, html_doc, scale_low=0.35)
        if progress:
            progress(min(0.95, 0.15 + 0.80 * ((i + 1) / total_pages)), f"Đang OCR & dịch trang {i + 1}/{total_pages}...")
        _log(f"📄 Hoàn thành OCR & dịch trang {i + 1}/{total_pages}")

    try:
        out_doc.subset_fonts()
    except Exception:
        pass
    out_doc.save(dst, garbage=3, deflate=True)
    out_doc.close()
    return rep


def translate_pdf(src, dst, target, translator: Translator, source_lang: str = "auto",
                  glossary=None, fit: bool = True, min_scale: float = 0.65,
                  progress: Callable[[float, str], None] | None = None,
                  log_func: Callable[[str], None] | None = None,
                  use_cache: bool = True, batch_size: int = 50, rate_delay: float = 0.3) -> Report:
    """Dịch PDF giữ nguyên 100% hình ảnh, đồ họa, bảng biểu và tọa độ văn bản."""
    if target not in LANGS:
        raise ValueError(f"target phải thuộc {list(LANGS)}")

    def _log(msg: str):
        if log_func:
            try:
                log_func(msg)
            except Exception:
                pass

    doc = pymupdf.open(src)
    try:
        if doc.is_encrypted and not doc.authenticate(""):
            raise ValueError("Tệp PDF đang được bảo vệ bằng mật khẩu. Vui lòng gỡ mật khẩu trước khi dịch.")

        num_pages = len(doc)
        _log(f"🔍 Bắt đầu phân tích {num_pages} trang trong tệp PDF...")
        if progress:
            progress(0.05, f"Đang phân tích {num_pages} trang PDF...")

        units_by_page = _collect_pdf_units(doc)
        all_units = [u for page_units in units_by_page.values() for u in page_units]

        if not all_units and num_pages > 0:
            rep = _translate_scanned_pdf(doc, dst, target, translator, source_lang, glossary, progress, _log)
            _log("🎉 Hoàn tất 100%! Tệp PDF đã sẵn sàng tải về máy.")
            if progress:
                progress(1.0, "Hoàn tất 100%! Tệp sẵn sàng tải về.")
            return rep

        rep = Report(units=len(all_units))
        uniq = list(dict.fromkeys(u.tagged for u in all_units))
        rep.unique = len(uniq)

        src_label = "Tự động nhận diện" if source_lang == "auto" else LANGS.get(source_lang, (source_lang,))[0]
        dst_label = LANGS[target][0]
        _log(f"📊 Tìm thấy {len(all_units)} khối văn bản ({rep.unique} khối duy nhất) trên {num_pages} trang. Hướng dịch: {src_label} ➔ {dst_label}")

        first = " | ".join(u.plain.strip().replace("\n", " ") for u in all_units[:5])
        context = f"Tài liệu PDF bắt đầu bằng: {first[:400]}"
        if progress:
            progress(0.10, f"Đã quét {num_pages} trang ({len(all_units)} khối văn bản). Đang kết nối Gemini AI...")

        trans, degraded, cached_count = translate_all(
            translator, uniq, target, glossary, context,
            source_lang=source_lang, doc_type="pdf",
            max_items=batch_size, rate_delay=rate_delay, use_cache=use_cache,
            progress=progress, log_func=log_func,
        )
        rep.cached = cached_count

        _log(f"🎨 Đang tái dựng chữ Unicode, màu sắc & chống tràn cho {len(all_units)} khối văn bản PDF...")
        if progress:
            progress(0.92, "Đang chèn bản dịch vào bố cục PDF gốc...")

        scale_floor = min(0.45, max(0.15, min_scale - 0.15)) if fit else 1.0
        for page_idx, page_units in units_by_page.items():
            page = doc[page_idx]
            # 1. Xóa sạch chữ gốc theo từng span bbox (không xóa ảnh, nền màu, đường viền)
            for u in page_units:
                for r in u.span_rects:
                    page.add_redact_annot(r, fill=None)
            page.apply_redactions(
                images=pymupdf.PDF_REDACT_IMAGE_NONE,
                graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
            )

            # 2. Chèn văn bản đã dịch kèm font Unicode vào đúng tọa độ cũ
            for u in page_units:
                d = trans.get(u.tagged, u.tagged)
                if u.tagged in degraded:
                    rep.degraded.append(f"{u.where}: “{u.plain[:25].strip()}…”")

                html_str = _unit_to_html(u, d)
                rect = _compute_safe_rect(u, page_units, page.rect)
                spare, scale = page.insert_htmlbox(rect, html_str, scale_low=scale_floor, rotate=u.rotate)
                if spare < 0 and fit:
                    # Fallback an toàn: hạ scale tối thiểu hoặc nới rộng hộp để chữ không bị mất
                    spare2, scale = page.insert_htmlbox(rect, html_str, scale_low=0.1, rotate=u.rotate)
                    if spare2 < 0:
                        wider = pymupdf.Rect(rect.x0, rect.y0, max(rect.x1, page.rect.x1 - 12), max(rect.y1, rect.y0 + rect.height * 2.5))
                        page.insert_htmlbox(wider, html_str, scale_low=0.1, rotate=u.rotate)
                        rep.warnings.append(f"{u.where}: khối chữ rất dài, đã mở rộng khung chứa.")
                if scale < 0.98 and fit:
                    rep.fitted.append(f"{u.where} ({int(scale * 100)}%)")

        _log("💾 Đang nén font Unicode (subset) & lưu tệp PDF hoàn chỉnh...")
        if progress:
            progress(0.98, "Đang tối ưu dung lượng & lưu tệp PDF...")
        try:
            doc.subset_fonts()
        except Exception:
            pass
        doc.save(dst, garbage=3, deflate=True)

        _log("🎉 Hoàn tất 100%! Tệp PDF đã sẵn sàng tải về máy.")
        if progress:
            progress(1.0, "Hoàn tất 100%! Tệp sẵn sàng tải về.")
        return rep
    finally:
        doc.close()
