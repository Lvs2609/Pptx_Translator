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


def _extract_page_walls(page: pymupdf.Page) -> tuple[list[tuple[float, float, float]], list[tuple[float, float, float]]]:
    """Trích xuất toàn bộ đường viền ngang/dọc của bảng biểu, hộp Step và callout trên trang PDF."""
    h_lines: list[tuple[float, float, float]] = []
    v_lines: list[tuple[float, float, float]] = []
    pw, ph = page.rect.width, page.rect.height

    def _add_seg(p1, p2):
        x1, y1 = float(p1.x), float(p1.y)
        x2, y2 = float(p2.x), float(p2.y)
        if abs(y1 - y2) <= 1.5 and abs(x2 - x1) >= 6.0:
            h_lines.append(((y1 + y2) * 0.5, min(x1, x2), max(x1, x2)))
        elif abs(x1 - x2) <= 1.5 and abs(y2 - y1) >= 6.0:
            v_lines.append(((x1 + x2) * 0.5, min(y1, y2), max(y1, y2)))

    def _add_box_edges(r: pymupdf.Rect):
        w, h = r.width, r.height
        if w > pw * 0.98 and h > ph * 0.98:
            return
        if h <= 3.5 and w >= 6.0:
            h_lines.append(((r.y0 + r.y1) * 0.5, r.x0, r.x1))
        elif w <= 3.5 and h >= 6.0:
            v_lines.append(((r.x0 + r.x1) * 0.5, r.y0, r.y1))
        elif w >= 6.0 and h >= 6.0:
            h_lines.append((r.y0, r.x0, r.x1))
            h_lines.append((r.y1, r.x0, r.x1))
            v_lines.append((r.x0, r.y0, r.y1))
            v_lines.append((r.x1, r.y0, r.y1))

    for d in page.get_drawings():
        has_straight = False
        for item in d.get("items", []):
            kind = item[0]
            if kind == "l":
                _add_seg(item[1], item[2])
                if abs(item[1].x - item[2].x) >= 6.0 or abs(item[1].y - item[2].y) >= 6.0:
                    has_straight = True
            elif kind == "re":
                _add_box_edges(pymupdf.Rect(item[1]))
                has_straight = True
            elif kind == "qu":
                q = item[1]
                _add_seg(q.ul, q.ur)
                _add_seg(q.ll, q.lr)
                _add_seg(q.ul, q.ll)
                _add_seg(q.ur, q.lr)
                has_straight = True

        # Fallback cho callout bo tròn vẽ toàn bằng đường cong Bezier
        if not has_straight and (d.get("color") is not None or d.get("fill") is not None):
            r = d.get("rect")
            if r is not None and 12.0 <= r.width <= pw * 0.85 and 10.0 <= r.height <= ph * 0.65:
                _add_box_edges(pymupdf.Rect(r))

    return h_lines, v_lines


def _has_h_wall_between(y_top: float, y_bot: float, x0: float, x1: float,
                        h_lines: list[tuple[float, float, float]]) -> bool:
    """Kiểm tra có đường kẻ ngang ngăn cách hai dòng bảng hay không."""
    for hy, lx0, lx1 in h_lines:
        if y_top - 1.5 <= hy <= y_bot + 1.5:
            if min(x1, lx1) > max(x0, lx0) + 2.0:
                return True
    return False


def _collect_pdf_units(doc: pymupdf.Document) -> tuple[dict[int, list[PdfUnit]], dict[int, tuple]]:
    """Quét tất cả các trang PDF, trích xuất ranh giới khung/bảng và gom đoạn văn bảo toàn ô."""
    units_by_page: dict[int, list[PdfUnit]] = {}
    walls_by_page: dict[int, tuple] = {}
    flags = pymupdf.TEXTFLAGS_DICT & ~pymupdf.TEXT_DEHYPHENATE

    for page_idx in range(len(doc)):
        page = doc[page_idx]
        h_lines, v_lines = _extract_page_walls(page)
        walls_by_page[page_idx] = (h_lines, v_lines)
        p_dict = page.get_text("dict", flags=flags)
        page_units: list[PdfUnit] = []

        for block in p_dict.get("blocks", []):
            if block.get("type") != 0:
                continue

            line_groups: list[list[dict]] = []
            for line in block.get("lines", []):
                segs = _split_line_columns(line)
                if not segs:
                    continue
                # Nếu dòng bị tách nhiều cột ngang hoặc xoay góc -> mỗi cột là 1 unit riêng biệt
                if len(segs) > 1 or _dir_to_rotate(line.get("dir", (1, 0))) != 0:
                    for seg in segs:
                        u = _build_pdf_unit(page_idx, f"Trang {page_idx + 1}", [seg])
                        if u:
                            page_units.append(u)
                    line_groups.append([])
                    continue

                seg = segs[0]
                if line_groups and line_groups[-1]:
                    prev_seg = line_groups[-1][-1]
                    prev_y1 = prev_seg["bbox"][3]
                    curr_y0 = seg["bbox"][1]
                    prev_h = max(8.0, prev_y1 - prev_seg["bbox"][1])
                    v_gap = curr_y0 - prev_y1
                    prev_cx = (prev_seg["bbox"][0] + prev_seg["bbox"][2]) * 0.5
                    curr_cx = (seg["bbox"][0] + seg["bbox"][2]) * 0.5
                    x_aligned = abs(seg["bbox"][0] - prev_seg["bbox"][0]) <= 28.0 or abs(curr_cx - prev_cx) <= 16.0
                    crossed_wall = _has_h_wall_between(
                        prev_y1, curr_y0,
                        min(seg["bbox"][0], prev_seg["bbox"][0]),
                        max(seg["bbox"][2], prev_seg["bbox"][2]),
                        h_lines,
                    )
                    if -2.0 <= v_gap <= prev_h * 0.85 and x_aligned and not crossed_wall:
                        line_groups[-1].append(seg)
                        continue
                line_groups.append([seg])

            for grp in line_groups:
                if not grp:
                    continue
                u = _build_pdf_unit(page_idx, f"Trang {page_idx + 1}", grp)
                if u:
                    page_units.append(u)

        if page_units:
            units_by_page[page_idx] = page_units

    return units_by_page, walls_by_page


def _unit_to_html(unit: PdfUnit, translated: str, align: str = "left", tight: bool = False) -> str:
    """Chuyển chuỗi đã dịch sang HTML với kích thước tương đối để PyMuPDF tự co chống tràn."""
    cur = unit.default_id
    chunks: list[str] = []
    def_style = unit.styles_by_id.get(unit.default_id)
    raw_base_sz = def_style.size if def_style else unit.avg_size
    # Trong ô/hộp hẹp giảm nhẹ ~6% cỡ gốc và siết line-height để chữ không chạm viền
    base_sz = max(5.0, raw_base_sz * (0.93 if tight else 0.97))

    def _span_html(style: Optional[PdfSpanStyle], text: str) -> str:
        safe = html.escape(text, quote=True)
        if not style:
            return safe
        weight = "bold" if style.bold else "normal"
        fstyle = "italic" if style.italic else "normal"
        pct = max(65.0, min(160.0, (style.size / max(1.0, raw_base_sz)) * 100.0))
        sz_css = f" font-size:{pct:.0f}%;" if abs(pct - 100.0) >= 5.0 else ""
        return (
            f'<span style="font-weight:{weight}; font-style:{fstyle}; '
            f'color:{style.color};{sz_css}">{safe}</span>'
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

    line_h = "1.06" if tight else "1.12"
    return (
        f'<div style="font-family: sans-serif; font-size: {base_sz:.1f}pt; '
        f'line-height: {line_h}; text-align: {align}; overflow-wrap: break-word; '
        f'margin: 0; padding: 0.5pt 1pt;">{"".join(chunks)}</div>'
    )


def _compute_safe_rect(
    unit: PdfUnit,
    all_units: list[PdfUnit],
    h_lines: list[tuple[float, float, float]],
    v_lines: list[tuple[float, float, float]],
    page_rect: pymupdf.Rect,
) -> tuple[pymupdf.Rect, str, bool, str]:
    """Khóa chặt vùng chèn chữ bên trong ranh giới đường viền ô/callout và tránh khối lân cận."""
    ux0, uy0, ux1, uy1 = unit.bbox
    if unit.rotate != 0:
        return pymupdf.Rect(ux0, uy0, ux1, uy1), "left", True, "top"

    cx = (ux0 + ux1) * 0.5
    cy = (uy0 + uy1) * 0.5
    u_w = max(1.0, ux1 - ux0)
    u_h = max(1.0, uy1 - uy0)

    # 1. Tìm tường dọc trái/phải của ô bảng hoặc callout chứa đoạn văn
    left_wall: Optional[float] = None
    right_wall: Optional[float] = None
    for vx, ly0, ly1 in v_lines:
        if min(uy1 - 0.5, ly1) >= max(uy0 + 0.5, ly0) or (ly0 - 1.5 <= cy <= ly1 + 1.5):
            if vx <= ux0 + 2.5 and (left_wall is None or vx > left_wall):
                left_wall = vx
            if vx >= ux1 - 2.5 and (right_wall is None or vx < right_wall):
                right_wall = vx

    # 2. Tìm tường ngang trên/dưới của ô bảng hoặc callout
    top_wall: Optional[float] = None
    bottom_wall: Optional[float] = None
    for hy, lx0, lx1 in h_lines:
        if min(ux1 - 0.5, lx1) >= max(ux0 + 0.5, lx0) or (lx0 - 1.5 <= cx <= lx1 + 1.5):
            if hy <= uy0 + 2.5 and (top_wall is None or hy > top_wall):
                top_wall = hy
            if hy >= uy1 - 1.0 and (bottom_wall is None or hy < bottom_wall):
                bottom_wall = hy

    wall_x0 = (left_wall + 2.0) if left_wall is not None else page_rect.x0 + 4.0
    wall_x1 = (right_wall - 2.0) if right_wall is not None else page_rect.x1 - 4.0
    wall_y0 = (top_wall + 1.5) if top_wall is not None else page_rect.y0 + 4.0
    wall_y1 = (bottom_wall - 1.5) if bottom_wall is not None else page_rect.y1 - 4.0

    # 3. Tìm chữ lân cận theo chiều ngang trên cùng dải dòng
    nb_left = page_rect.x0 + 4.0
    nb_right = page_rect.x1 - 4.0
    same_row_siblings = 0
    for other in all_units:
        if other is unit:
            continue
        ox0, oy0, ox1, oy1 = other.bbox
        if min(uy1 - 1.0, oy1) > max(uy0 + 1.0, oy0):
            if ox1 <= ux0 + 1.5:
                nb_left = max(nb_left, ox1 + 2.5)
            elif ox0 >= ux1 - 1.5:
                nb_right = min(nb_right, ox0 - 2.5)
            if left_wall is not None and right_wall is not None:
                ocx = (ox0 + ox1) * 0.5
                if left_wall <= ocx <= right_wall:
                    same_row_siblings += 1

    limit_x0 = max(wall_x0, nb_left)
    limit_x1 = min(wall_x1, nb_right)

    align = "left"
    enclosed_h = (left_wall is not None and right_wall is not None and (right_wall - left_wall) <= 440.0)
    if enclosed_h and same_row_siblings == 0 and limit_x1 > limit_x0 + 8.0:
        # Một chữ duy nhất nằm trong ô bảng/callout: tận dụng toàn bộ chiều ngang bên trong ô
        left_gap = ux0 - wall_x0
        right_gap = wall_x1 - ux1
        if abs(left_gap - right_gap) <= 10.0 and left_gap >= 3.0 and right_gap >= 3.0:
            align = "center"
            final_x0, final_x1 = limit_x0, limit_x1
        elif right_gap <= 6.0 and left_gap > right_gap + 12.0:
            align = "right"
            final_x0, final_x1 = limit_x0, limit_x1
        else:
            final_x0 = limit_x0 if left_gap <= 8.0 else max(limit_x0, ux0)
            final_x1 = limit_x1
    elif right_wall is not None and (right_wall - ux1) <= 160.0:
        final_x0 = max(limit_x0, ux0)
        final_x1 = limit_x1
    else:
        # Vùng mở không có khung viền sát bên phải: chỉ mở rộng cực nhẹ để không tràn vào ảnh
        final_x0 = max(limit_x0, ux0)
        final_x1 = min(limit_x1, ux1 + min(10.0, max(3.0, u_w * 0.08)))

    if final_x1 <= final_x0 + 6.0:
        final_x0 = max(page_rect.x0 + 2.0, min(ux0, wall_x0))
        final_x1 = max(final_x0 + 8.0, min(limit_x1, ux1 + 4.0))

    # 4. Tìm chữ lân cận theo chiều dọc và anh em cùng cột trong ô [wall_y0, wall_y1]
    nb_bottom = page_rect.y1 - 4.0
    same_col_siblings = 0
    for other in all_units:
        if other is unit:
            continue
        ox0, oy0, ox1, oy1 = other.bbox
        if min(final_x1, ox1) > max(final_x0, ox0) - 2.0:
            if oy0 >= uy1 - 2.0:
                nb_bottom = min(nb_bottom, oy0 - 1.5)
            if top_wall is not None and bottom_wall is not None:
                ocy = (oy0 + oy1) * 0.5
                if top_wall <= ocy <= bottom_wall:
                    same_col_siblings += 1

    hard_y1 = min(wall_y1, nb_bottom)
    tight_box = False
    v_align = "top"

    enclosed_v = (top_wall is not None and bottom_wall is not None and (wall_y1 - uy1) <= max(u_h * 1.4, 42.0))
    if enclosed_v:
        # Đóng khung kín phía dưới (ô bảng, hộp Step, callout): TUYỆT ĐỐI không vượt wall_y1
        tight_box = (wall_y1 - wall_y0) <= max(u_h * 1.6, 34.0)
        if same_col_siblings == 0:
            top_gap = uy0 - wall_y0
            bot_gap = wall_y1 - uy1
            if abs(top_gap - bot_gap) <= 7.0 and top_gap >= 1.5 and bot_gap >= 1.5:
                v_align = "middle"
                final_y0 = wall_y0
            elif bot_gap <= 5.0 and top_gap > bot_gap + 10.0:
                v_align = "bottom"
                final_y0 = wall_y0
            else:
                final_y0 = max(wall_y0, uy0 - 0.5)
        else:
            final_y0 = max(wall_y0, uy0 - 0.5)
        final_y1 = max(final_y0 + 6.0, hard_y1)
    else:
        # Không có đường viền sát bên dưới -> giữ nguyên y0 gốc, chỉ cho phép nới dọc vừa phải
        final_y0 = max(wall_y0, uy0 - 0.5)
        open_cap = uy0 + max(u_h * 1.25, unit.avg_size * 1.55)
        final_y1 = max(final_y0 + 6.0, min(hard_y1, open_cap))

    return pymupdf.Rect(final_x0, final_y0, final_x1, final_y1), align, tight_box, v_align


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

        units_by_page, walls_by_page = _collect_pdf_units(doc)
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

        scale_floor = min(0.35, max(0.18, min_scale * 0.4)) if fit else 1.0
        scratch_doc = pymupdf.open()
        try:
            for page_idx, page_units in units_by_page.items():
                page = doc[page_idx]
                h_lines, v_lines = walls_by_page.get(page_idx, ([], []))

                # Lọc các khối thực sự thay đổi nội dung (giữ nguyên 100% mã hiệu/số/ID không đổi)
                changed_units: list[tuple[PdfUnit, str]] = []
                for u in page_units:
                    d = trans.get(u.tagged, u.tagged)
                    dst_plain = html.unescape(STYLE_TAG_RE.sub("", d).replace("<br/>", "\n")).strip()
                    if dst_plain == u.plain.strip():
                        continue
                    changed_units.append((u, d))

                if not changed_units:
                    continue

                # 1. Xóa sạch chữ gốc theo từng span bbox (không xóa ảnh, nền màu, đường viền)
                for u, _d in changed_units:
                    for r in u.span_rects:
                        page.add_redact_annot(r, fill=None)
                page.apply_redactions(
                    images=pymupdf.PDF_REDACT_IMAGE_NONE,
                    graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                )

                spage = scratch_doc.new_page(width=page.rect.width, height=page.rect.height)

                # 2. Chèn văn bản đã dịch giữ đúng vị trí gốc (căn giữa dọc/ngang chuẩn xác)
                for u, d in changed_units:
                    if u.tagged in degraded:
                        rep.degraded.append(f"{u.where}: “{u.plain[:25].strip()}…”")

                    rect, align, tight, v_align = _compute_safe_rect(u, page_units, h_lines, v_lines, page.rect)
                    html_str = _unit_to_html(u, d, align=align, tight=tight)

                    if v_align in ("middle", "bottom"):
                        m_spare, m_scale = spage.insert_htmlbox(rect, html_str, scale_low=scale_floor, rotate=u.rotate)
                        if m_spare < 0:
                            m_spare, m_scale = spage.insert_htmlbox(rect, html_str, scale_low=0.08, rotate=u.rotate)
                        if m_spare > 1.0:
                            offset = m_spare * 0.48 if v_align == "middle" else max(0.0, m_spare - 0.8)
                            rect = pymupdf.Rect(rect.x0, min(rect.y1 - 4.0, rect.y0 + offset), rect.x1, rect.y1)

                    spare, scale = page.insert_htmlbox(rect, html_str, scale_low=scale_floor, rotate=u.rotate)
                    if spare < 0:
                        spare2, scale = page.insert_htmlbox(rect, html_str, scale_low=0.08, rotate=u.rotate)
                        if spare2 < 0:
                            padded = pymupdf.Rect(rect.x0, max(page.rect.y0, rect.y0 - 2.0), rect.x1, min(page.rect.y1, rect.y1 + max(4.0, rect.height * 0.4)))
                            page.insert_htmlbox(padded, html_str, scale_low=0.05, rotate=u.rotate)
                            rep.warnings.append(f"{u.where}: khối chữ quá hẹp, đã co tối đa ({int(scale * 100)}%).")
                    if scale < 0.98 and fit:
                        rep.fitted.append(f"{u.where} ({int(scale * 100)}%)")
        finally:
            scratch_doc.close()

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
