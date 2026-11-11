"""Bóc tách và dịch thuật tệp Excel (.xlsx) bảo toàn 100% công thức, định dạng ô, rich text và ghi chú."""
from __future__ import annotations

import copy
import html
import math
import re
from dataclasses import dataclass, field
from typing import Callable, Optional

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.utils import get_column_letter

from src.config import LANGS
from src.core.autofit import _width
from src.core.parser import STYLE_TAG_RE, TAG_RE, Report, needs_translation_text
from src.core.translator import Translator, translate_all


@dataclass
class XlsxUnit:
    """Một ô hoặc ghi chú Excel cần dịch."""
    ws: object
    cell: object
    kind: str                         # "cell" | "richtext" | "comment"
    where: str
    font_by_id: dict = field(default_factory=dict)
    default_id: int = 1
    use_tags: bool = False
    tagged: str = ""
    plain: str = ""


def _encode_newlines(text: str) -> str:
    esc = html.escape(text.replace("\r\n", "\n").replace("\r", "\n"), quote=False)
    return esc.replace("\n", "<br/>")


def _decode_to_plain(translated: str) -> str:
    no_tags = STYLE_TAG_RE.sub("", translated)
    no_br = no_tags.replace("<br/>", "\n")
    return html.unescape(no_br)


def _font_sig(font_obj: Optional[InlineFont]) -> str:
    if font_obj is None:
        return ""
    return repr(font_obj)


def _parse_richtext_cell(ws, cell, rt: CellRichText, where: str) -> Optional[XlsxUnit]:
    """Phân giải CellRichText thành thẻ <rN> để bảo toàn in đậm/màu chữ từng đoạn trong ô Excel."""
    parts: list[list] = []
    font_by_sig: dict[str, Optional[InlineFont]] = {}

    for item in rt:
        if isinstance(item, TextBlock):
            text = item.text or ""
            font = item.font
        else:
            text = str(item) if item is not None else ""
            font = None
        if not text:
            continue
        sig = _font_sig(font)
        font_by_sig.setdefault(sig, font)
        if parts and parts[-1][0] == sig:
            parts[-1][1] += text
        else:
            parts.append([sig, text])

    if not parts:
        return None

    # Gom khoảng trắng đứng lẻ vào font lân cận
    for i, p in enumerate(parts):
        if not p[1].strip():
            nxt = next((y for y in parts[i + 1:] if y[1].strip()), None)
            prv = next((y for y in reversed(parts[:i]) if y[1].strip()), None)
            ref = nxt or prv
            if ref:
                p[0] = ref[0]

    sig_to_id: dict[str, int] = {}
    font_by_id: dict[int, Optional[InlineFont]] = {}
    for sig, _txt in parts:
        if sig not in sig_to_id:
            rid = len(sig_to_id) + 1
            sig_to_id[sig] = rid
            font_by_id[rid] = font_by_sig.get(sig)

    use_tags = len(sig_to_id) > 1
    out, plain = [], []
    longest: dict[int, int] = {}
    for sig, txt in parts:
        rid = sig_to_id[sig]
        longest[rid] = longest.get(rid, 0) + len(txt)
        plain.append(txt)
        enc = _encode_newlines(txt)
        out.append(f"<r{rid}>{enc}</r{rid}>" if use_tags else enc)

    plain_str = "".join(plain)
    if not needs_translation_text(plain_str):
        return None

    return XlsxUnit(
        ws=ws,
        cell=cell,
        kind="richtext",
        where=where,
        font_by_id=font_by_id,
        default_id=max(longest, key=longest.get),
        use_tags=use_tags,
        tagged="".join(out),
        plain=plain_str,
    )


def _apply_richtext_translation(unit: XlsxUnit, translated: str) -> str:
    """Phục hồi CellRichText từ chuỗi đã dịch có gắn thẻ <rN>."""
    cur = unit.default_id
    blocks = []
    plain_parts = []

    for tok in TAG_RE.split(translated):
        if not tok:
            continue
        if m := re.fullmatch(r"<r(\d+)>", tok):
            rid = int(m.group(1))
            cur = rid if rid in unit.font_by_id else unit.default_id
        elif re.fullmatch(r"</r\d+>", tok):
            cur = unit.default_id
        elif tok == "<br/>":
            plain_parts.append("\n")
            font = unit.font_by_id.get(cur)
            if font is not None:
                blocks.append(TextBlock(copy.copy(font), "\n"))
            else:
                blocks.append("\n")
        elif re.fullmatch(r"<f\d+/>", tok):
            continue
        else:
            txt = html.unescape(tok)
            plain_parts.append(txt)
            font = unit.font_by_id.get(cur)
            if font is not None:
                blocks.append(TextBlock(copy.copy(font), txt))
            else:
                blocks.append(txt)

    if blocks:
        unit.cell.value = CellRichText(blocks)
    return "".join(plain_parts)


def _collect_xlsx_units(wb: openpyxl.Workbook) -> list[XlsxUnit]:
    """Thu thập mọi ô chuỗi, rich text và comment trong toàn bộ workbook Excel (bỏ qua công thức)."""
    units: list[XlsxUnit] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                where = f"{ws.title}!{cell.coordinate}"
                val = cell.value

                # Tuyệt đối bỏ qua công thức Excel (=SUM, =IF...) và giá trị số/ngày
                if val is not None and cell.data_type != "f":
                    if isinstance(val, CellRichText):
                        u = _parse_richtext_cell(ws, cell, val, where)
                        if u:
                            units.append(u)
                    elif isinstance(val, str) and not val.startswith("="):
                        if needs_translation_text(val):
                            units.append(XlsxUnit(
                                ws=ws,
                                cell=cell,
                                kind="cell",
                                where=where,
                                tagged=_encode_newlines(val),
                                plain=val,
                            ))

                # Ghi chú ô (Cell comment / Note)
                if cell.comment and cell.comment.text and needs_translation_text(cell.comment.text):
                    ctext = cell.comment.text
                    units.append(XlsxUnit(
                        ws=ws,
                        cell=cell,
                        kind="comment",
                        where=f"{where} (Comment)",
                        tagged=_encode_newlines(ctext),
                        plain=ctext,
                    ))
    return units


def translate_xlsx(src, dst, target, translator: Translator, source_lang: str = "auto",
                   glossary=None, fit: bool = True,
                   progress: Callable[[float, str], None] | None = None,
                   log_func: Callable[[str], None] | None = None,
                   use_cache: bool = True, batch_size: int = 50, rate_delay: float = 0.3) -> Report:
    """Dịch bảng tính Excel (.xlsx) bảo toàn 100% công thức, màu ô, viền, merged cells và ghi chú."""
    if target not in LANGS:
        raise ValueError(f"target phải thuộc {list(LANGS)}")

    def _log(msg: str):
        if log_func:
            try:
                log_func(msg)
            except Exception:
                pass

    wb = openpyxl.load_workbook(src, rich_text=True)
    num_sheets = len(wb.worksheets)
    _log(f"🔍 Bắt đầu phân tích {num_sheets} trang tính (sheet) trong tệp Excel...")
    if progress:
        progress(0.05, f"Đang quét {num_sheets} sheet trong tệp Excel...")

    units = _collect_xlsx_units(wb)
    rep = Report(units=len(units))
    uniq = list(dict.fromkeys(u.tagged for u in units))
    rep.unique = len(uniq)

    if not units:
        rep.warnings.append("Không tìm thấy ô văn bản cần dịch trong tệp Excel (các công thức được giữ nguyên).")
        wb.save(dst)
        if progress:
            progress(1.0, "Hoàn tất!")
        return rep

    src_label = "Tự động nhận diện" if source_lang == "auto" else LANGS.get(source_lang, (source_lang,))[0]
    dst_label = LANGS[target][0]
    _log(f"📊 Tìm thấy {len(units)} mục văn bản ({rep.unique} mục duy nhất, giữ nguyên công thức). Hướng dịch: {src_label} ➔ {dst_label}")

    first = " | ".join(u.plain.strip().replace("\n", " ") for u in units[:8])
    context = f"Bảng tính Excel bắt đầu bằng: {first[:400]}"
    if progress:
        progress(0.10, f"Đã quét {num_sheets} sheet ({len(units)} ô văn bản). Đang kết nối Gemini AI...")

    trans, degraded, cached_count = translate_all(
        translator, uniq, target, glossary, context,
        source_lang=source_lang, doc_type="xlsx",
        max_items=batch_size, rate_delay=rate_delay, use_cache=use_cache,
        progress=progress, log_func=log_func,
    )
    rep.cached = cached_count

    _log(f"🎨 Đang ghi bản dịch vào {len(units)} ô Excel & tối ưu hiển thị chống tràn...")
    if progress:
        progress(0.92, "Đang phục hồi định dạng ô & căn chỉnh Excel...")

    # Tính tổng chiều rộng các vùng merged cell để ước lượng đúng số dòng ngắt
    merged_span_cols: dict[tuple[str, str], tuple[int, int]] = {}
    for ws in wb.worksheets:
        for mr in ws.merged_cells.ranges:
            coord = ws.cell(mr.min_row, mr.min_col).coordinate
            merged_span_cols[(ws.title, coord)] = (mr.min_col, mr.max_col)

    col_widened: set[tuple[str, int]] = set()
    rows_fitted: set[tuple[str, int]] = set()
    for u in units:
        d = trans.get(u.tagged, u.tagged)
        if u.tagged in degraded:
            rep.degraded.append(f"{u.where}: “{u.plain[:25].strip()}…”")

        if u.kind == "comment":
            u.cell.comment.text = _decode_to_plain(d)
            continue

        if u.kind == "richtext":
            translated_plain = _apply_richtext_translation(u, d)
        else:
            translated_plain = _decode_to_plain(d)
            # Chống Excel hiểu nhầm chuỗi bắt đầu bằng '=' là công thức lỗi
            if translated_plain.startswith("="):
                translated_plain = "'" + translated_plain
            u.cell.value = translated_plain

        if fit and u.cell.coordinate:
            try:
                c_min, c_max = merged_span_cols.get((u.ws.title, u.cell.coordinate), (u.cell.column, u.cell.column))
                total_col_w = sum(
                    (u.ws.column_dimensions[get_column_letter(ci)].width or 10.5)
                    for ci in range(c_min, c_max + 1)
                )
                w_src = _width(u.plain)
                w_dst = _width(translated_plain)
                # Quy đổi tương đương độ dài ký tự hiển thị Excel
                dst_chars = w_dst * 2.0

                if dst_chars > total_col_w * 0.9 or "\n" in translated_plain or (w_src > 0 and w_dst > w_src * 1.1):
                    align = copy.copy(u.cell.alignment) if u.cell.alignment else openpyxl.styles.Alignment()
                    if not align.wrap_text:
                        align.wrap_text = True
                        u.cell.alignment = align

                    # Mở rộng cột đơn vừa phải nếu hẹp (< 42)
                    if c_min == c_max:
                        col_key = (u.ws.title, u.cell.column)
                        if col_key not in col_widened:
                            col_letter = get_column_letter(u.cell.column)
                            cur_w = u.ws.column_dimensions[col_letter].width or 10.5
                            if cur_w < 42 and dst_chars > cur_w * 0.95:
                                new_w = min(46.0, round(max(cur_w * min(1.35, max(1.1, w_dst / max(0.5, w_src))), min(38.0, dst_chars * 1.08)), 1))
                                u.ws.column_dimensions[col_letter].width = new_w
                                total_col_w = new_w
                                col_widened.add(col_key)

                    # Tự động tăng chiều cao dòng nếu dòng đã gán chiều cao cố định (chống cắt chữ theo chiều dọc)
                    row_dim = u.ws.row_dimensions[u.cell.row]
                    if row_dim.height is not None:
                        font_sz = float(u.cell.font.size) if (u.cell.font and u.cell.font.size) else 11.0
                        lines = translated_plain.splitlines() or [""]
                        est_lines = sum(max(1, math.ceil((_width(ln) * 2.0) / max(4.0, total_col_w * 0.92))) for ln in lines)
                        needed_h = round(est_lines * max(13.5, font_sz * 1.38) + 4.0, 1)
                        if needed_h > row_dim.height:
                            row_dim.height = min(409.0, needed_h)
                            rows_fitted.add((u.ws.title, u.cell.row))
            except Exception:
                pass

    if col_widened or rows_fitted:
        rep.fitted.append(f"Chống tràn Excel ({len(col_widened)} cột, {len(rows_fitted)} dòng)")

    _log("💾 Đang đóng gói và lưu tệp Excel (.xlsx) hoàn chỉnh...")
    if progress:
        progress(0.98, "Đang hoàn tất lưu tệp Excel...")
    wb.save(dst)
    _log("🎉 Hoàn tất 100%! Tệp Excel đã sẵn sàng tải về máy.")
    if progress:
        progress(1.0, "Hoàn tất 100%! Tệp sẵn sàng tải về.")
    return rep
