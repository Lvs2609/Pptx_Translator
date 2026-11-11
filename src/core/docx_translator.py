"""Bóc tách cấu trúc XML Word (.docx) và dịch thuật bảo toàn 100% định dạng, bảng biểu, hình ảnh."""
from __future__ import annotations

import copy
import html
import re
from dataclasses import dataclass, field
from typing import Callable, Optional

import docx
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from lxml import etree

from src.config import LANGS
from src.core.parser import STYLE_TAG_RE, TAG_RE, Report, needs_translation_text
from src.core.translator import Translator, translate_all

_XML_SPACE = qn("xml:space")
_SKIP_P_ANCESTORS = {qn("w:del"), qn("w:moveFrom")}
_IGNORED_P_CHILDREN = {
    qn("w:pPr"),
    qn("w:bookmarkStart"),
    qn("w:bookmarkEnd"),
    qn("w:commentRangeStart"),
    qn("w:commentRangeEnd"),
    qn("w:proofErr"),
    qn("w:permStart"),
    qn("w:permEnd"),
    qn("w:del"),
    qn("w:moveFrom"),
}
_OPAQUE_RUN_CHILDREN = {
    qn("w:drawing"),
    qn("w:pict"),
    qn("w:object"),
    qn("w:fldChar"),
    qn("w:instrText"),
    qn("w:sym"),
    qn("w:footnoteReference"),
    qn("w:endnoteReference"),
    qn("w:commentReference"),
    qn("w:tab"),
    qn("w:ptab"),
    qn("w:pgNum"),
    qn("w:dayShort"),
    qn("w:monthShort"),
    qn("w:yearShort"),
    qn("w:lastRenderedPageBreak"),
    qn("w:ruby"),
}


@dataclass
class DocxUnit:
    """Một đoạn văn bản Word (<w:p>) đã phân giải run định dạng và đối tượng nhúng."""
    p: etree._Element
    where: str
    replaced_nodes: list = field(default_factory=list)
    rpr_by_id: dict = field(default_factory=dict)
    link_by_id: dict = field(default_factory=dict)
    sig_to_id: dict = field(default_factory=dict)
    fields: dict = field(default_factory=dict)
    brs: list = field(default_factory=list)
    default_id: int = 1
    use_tags: bool = False
    tagged: str = ""
    plain: str = ""


def _rpr_sig(rPr, link_el=None) -> str:
    """Tạo chữ ký duy nhất cho cặp định dạng <w:rPr> + <w:hyperlink> (nếu có)."""
    rpr_str = ""
    if rPr is not None:
        c = copy.deepcopy(rPr)
        for tag_name in ("w:noProof", "w:lang"):
            for old in c.findall(qn(tag_name)):
                c.remove(old)
        rpr_str = etree.tostring(c, method="c14n").decode()

    link_str = ""
    if link_el is not None:
        attrs = sorted((k, v) for k, v in link_el.attrib.items())
        link_str = "|".join(f"{k}={v}" for k, v in attrs)

    return f"{rpr_str}||{link_str}"


def _has_ancestor(el, tags: set[str]) -> bool:
    cur = el.getparent()
    while cur is not None:
        if cur.tag in tags:
            return True
        cur = cur.getparent()
    return False


def _is_soft_break(el) -> bool:
    if el.tag == qn("w:cr"):
        return True
    if el.tag == qn("w:br"):
        br_type = el.get(qn("w:type"))
        return br_type in (None, "", "textWrapping")
    return False


def _wrap_opaque_child(src_rPr, child) -> etree._Element:
    r = OxmlElement("w:r")
    if src_rPr is not None:
        r.append(copy.deepcopy(src_rPr))
    r.append(copy.deepcopy(child))
    return r


def _parse_docx_paragraph(p: etree._Element, where: str) -> Optional[DocxUnit]:
    """Tách đoạn văn <w:p> thành chuỗi có gắn thẻ <rN>, <br/>, <fN/> bảo toàn ảnh & trường."""
    parts: list[list] = []
    meta_by_sig: dict[str, tuple] = {}
    replaced_nodes: list = []

    def _process_run(r_el, link_el=None):
        rPr = r_el.find(qn("w:rPr"))
        sig = _rpr_sig(rPr, link_el)
        meta_by_sig.setdefault(sig, (rPr, link_el))

        children = [c for c in r_el if c.tag != qn("w:rPr")]
        if not children:
            return

        # Nếu run không chứa text mà chỉ là ảnh/trường/phím tab -> giữ nguyên toàn bộ run
        has_text = any(c.tag in (qn("w:t"), qn("w:delText")) for c in children)
        if not has_text and all(not _is_soft_break(c) for c in children):
            parts.append(["f", copy.deepcopy(r_el), sig, link_el])
            return

        for child in children:
            if child.tag == qn("w:t"):
                text = child.text or ""
                if not text:
                    continue
                if parts and parts[-1][0] == "seg" and (parts[-1][2] == sig or (not text.strip() and link_el is None)):
                    parts[-1][1] += text
                else:
                    parts.append(["seg", text, sig, link_el])
            elif _is_soft_break(child):
                parts.append(["br", copy.deepcopy(child), "", link_el])
            elif child.tag == qn("w:noBreakHyphen"):
                text = "-"
                if parts and parts[-1][0] == "seg":
                    parts[-1][1] += text
                else:
                    parts.append(["seg", text, sig, link_el])
            else:
                # Bảo toàn nguyên vẹn <w:drawing>, <w:tab/>, <w:br w:type="page"/>, <w:fldChar>...
                parts.append(["f", _wrap_opaque_child(rPr, child), sig, link_el])

    for child in p:
        if child.tag in _IGNORED_P_CHILDREN:
            continue
        if child.tag == qn("w:r"):
            replaced_nodes.append(child)
            _process_run(child, None)
        elif child.tag == qn("w:hyperlink"):
            replaced_nodes.append(child)
            for sub in child:
                if sub.tag == qn("w:r"):
                    _process_run(sub, child)
        elif child.tag == qn("w:ins"):
            # Track changes accepted insert: vẫn dịch nội dung bên trong
            replaced_nodes.append(child)
            for sub in child:
                if sub.tag == qn("w:r"):
                    _process_run(sub, None)
                elif sub.tag == qn("w:hyperlink"):
                    for hr in sub:
                        if hr.tag == qn("w:r"):
                            _process_run(hr, sub)
        elif child.tag.endswith("oMath") or child.tag.endswith("oMathPara"):
            replaced_nodes.append(child)
            parts.append(["f", copy.deepcopy(child), "", None])

    segs = [x for x in parts if x[0] == "seg"]
    if not segs or not replaced_nodes:
        return None

    # Gom khoảng trắng đứng lẻ vào kiểu định dạng lân cận
    for i, x in enumerate(parts):
        if x[0] == "seg" and not x[1].strip():
            nxt = next((y for y in parts[i + 1:] if y[0] == "seg" and y[1].strip()), None)
            prv = next((y for y in reversed(parts[:i]) if y[0] == "seg" and y[1].strip()), None)
            ref = nxt or prv
            if ref:
                x[2] = ref[2]
                x[3] = ref[3]

    unit = DocxUnit(p=p, where=where, replaced_nodes=replaced_nodes)
    for x in parts:
        if x[0] == "seg" and x[2] not in unit.sig_to_id:
            rid = len(unit.sig_to_id) + 1
            unit.sig_to_id[x[2]] = rid
            rPr, link_el = meta_by_sig.get(x[2], (None, None))
            unit.rpr_by_id[rid] = rPr
            unit.link_by_id[rid] = link_el

    unit.use_tags = len(unit.sig_to_id) > 1
    out, plain, fn = [], [], 0
    longest: dict[int, int] = {}

    for kind_, payload, sig, _link in parts:
        if kind_ == "seg":
            rid = unit.sig_to_id[sig]
            longest[rid] = longest.get(rid, 0) + len(payload)
            plain.append(payload)
            esc = html.escape(payload, quote=False)
            out.append(f"<r{rid}>{esc}</r{rid}>" if unit.use_tags else esc)
        elif kind_ == "br":
            unit.brs.append(payload)
            plain.append("\n")
            out.append("<br/>")
        else:
            fn += 1
            unit.fields[fn] = payload
            out.append(f"<f{fn}/>")

    unit.default_id = max(longest, key=longest.get)
    unit.tagged = "".join(out)
    unit.plain = "".join(plain)
    return unit


def _collect_docx_units(doc: docx.Document, include_headers_footers: bool = True) -> list[DocxUnit]:
    """Thu thập mọi đoạn văn cần dịch trong Word: Body, Bảng, Textbox, Header & Footer."""
    units: list[DocxUnit] = []
    seen_roots: set[int] = set()

    def scan_root(root_el, label: str):
        if root_el is None or id(root_el) in seen_roots:
            return
        seen_roots.add(id(root_el))
        for idx, p in enumerate(root_el.iter(qn("w:p")), 1):
            if _has_ancestor(p, _SKIP_P_ANCESTORS):
                continue
            where = f"{label} #{idx}" if len(label) < 20 else label
            u = _parse_docx_paragraph(p, where)
            if u and needs_translation_text(u.plain):
                units.append(u)

    scan_root(doc.element.body, "Đoạn văn")

    if include_headers_footers:
        for s_idx, section in enumerate(doc.sections, 1):
            for kind_name, hf in (
                ("Header", section.header),
                ("First Header", section.first_page_header),
                ("Even Header", section.even_page_header),
                ("Footer", section.footer),
                ("First Footer", section.first_page_footer),
                ("Even Footer", section.even_page_footer),
            ):
                try:
                    if not hf.is_linked_to_previous:
                        scan_root(hf._element, f"Trang {s_idx} ({kind_name})")
                except Exception:
                    pass

    return units


def _make_docx_run(src_rPr, text: str, lang: str, ea_font: Optional[str]) -> etree._Element:
    """Tạo <w:r> mới giữ nguyên định dạng <w:rPr> gốc và gán font CJK nếu cần."""
    r = OxmlElement("w:r")
    if src_rPr is not None:
        rPr = copy.deepcopy(src_rPr)
        for existing_lang in rPr.findall(qn("w:lang")):
            existing_lang.set(qn("w:val"), lang)
        if ea_font:
            rf = rPr.get_or_add_rFonts()
            rf.set(qn("w:eastAsia"), ea_font)
        r.append(rPr)
    elif ea_font:
        rPr = r.get_or_add_rPr()
        rPr.get_or_add_rFonts().set(qn("w:eastAsia"), ea_font)

    t = OxmlElement("w:t")
    if text and (text[0].isspace() or text[-1].isspace() or "  " in text):
        t.set(_XML_SPACE, "preserve")
    t.text = text
    r.append(t)
    return r


def _clone_hyperlink_shell(link_el: etree._Element) -> etree._Element:
    shell = OxmlElement("w:hyperlink")
    for k, v in link_el.attrib.items():
        shell.set(k, v)
    return shell


def _apply_docx_translation(u: DocxUnit, translated: str, target: str):
    """Ghi văn bản đã dịch trở lại <w:p>, bảo toàn chính xác font, màu sắc, link và hình ảnh."""
    _, lang, ea_font = LANGS[target]
    cur = u.default_id
    brs = list(u.brs)
    built: list[tuple[etree._Element, Optional[etree._Element]]] = []

    for tok in TAG_RE.split(translated):
        if not tok:
            continue
        if m := re.fullmatch(r"<r(\d+)>", tok):
            rid = int(m.group(1))
            cur = rid if rid in u.rpr_by_id else u.default_id
        elif re.fullmatch(r"</r\d+>", tok):
            cur = u.default_id
        elif tok == "<br/>":
            br_node = copy.deepcopy(brs.pop(0)) if brs else OxmlElement("w:br")
            r = OxmlElement("w:r")
            rPr = u.rpr_by_id.get(cur)
            if rPr is not None:
                r.append(copy.deepcopy(rPr))
            r.append(br_node)
            built.append((r, u.link_by_id.get(cur)))
        elif m := re.fullmatch(r"<f(\d+)/>", tok):
            fid = int(m.group(1))
            if fid in u.fields:
                built.append((u.fields[fid], None))
        else:
            r = _make_docx_run(u.rpr_by_id.get(cur), html.unescape(tok), lang, ea_font)
            built.append((r, u.link_by_id.get(cur)))

    # Xác định vị trí chèn sau node thay thế cuối cùng để không làm xáo trộn bookmark/pPr
    anchor_next = None
    last_replaced = u.replaced_nodes[-1]
    nxt = last_replaced.getnext()
    while nxt is not None and nxt in u.replaced_nodes:
        nxt = nxt.getnext()
    anchor_next = nxt

    for node in u.replaced_nodes:
        if node.getparent() is u.p:
            u.p.remove(node)

    # Nhóm các run cùng hyperlink shell liền kề thành một <w:hyperlink>
    final_elements: list[etree._Element] = []
    cur_link_key: Optional[str] = None
    cur_link_box: Optional[etree._Element] = None

    for el, link_el in built:
        if link_el is not None and el.tag == qn("w:r"):
            key = "|".join(f"{k}={v}" for k, v in sorted(link_el.attrib.items()))
            if cur_link_box is None or key != cur_link_key:
                cur_link_box = _clone_hyperlink_shell(link_el)
                cur_link_key = key
                final_elements.append(cur_link_box)
            cur_link_box.append(el)
        else:
            cur_link_box = None
            cur_link_key = None
            final_elements.append(el)

    for el in final_elements:
        if anchor_next is not None:
            anchor_next.addprevious(el)
        else:
            u.p.append(el)


def translate_docx(src, dst, target, translator: Translator, source_lang: str = "auto",
                   glossary=None, include_headers_footers: bool = True,
                   progress: Callable[[float, str], None] | None = None,
                   log_func: Callable[[str], None] | None = None,
                   use_cache: bool = True, batch_size: int = 50, rate_delay: float = 0.3) -> Report:
    """Dịch tài liệu Word (.docx) giữ nguyên 100% bố cục, bảng biểu, hình ảnh và định dạng chữ."""
    if target not in LANGS:
        raise ValueError(f"target phải thuộc {list(LANGS)}")

    def _log(msg: str):
        if log_func:
            try:
                log_func(msg)
            except Exception:
                pass

    doc = docx.Document(src)
    _log("🔍 Bắt đầu phân tích cấu trúc XML tài liệu Word (.docx)...")
    if progress:
        progress(0.05, "Đang quét nội dung tài liệu Word...")

    units = _collect_docx_units(doc, include_headers_footers=include_headers_footers)
    rep = Report(units=len(units))
    uniq = list(dict.fromkeys(u.tagged for u in units))
    rep.unique = len(uniq)

    if not units:
        rep.warnings.append("Không tìm thấy văn bản cần dịch trong tệp Word.")
        doc.save(dst)
        if progress:
            progress(1.0, "Hoàn tất!")
        return rep

    src_label = "Tự động nhận diện" if source_lang == "auto" else LANGS.get(source_lang, (source_lang,))[0]
    dst_label = LANGS[target][0]
    _log(f"📊 Tìm thấy {len(units)} đoạn văn ({rep.unique} đoạn duy nhất). Hướng dịch: {src_label} ➔ {dst_label}")

    first = " | ".join(u.plain.strip().replace("\n", " ") for u in units[:5])
    context = f"Tài liệu Word bắt đầu bằng: {first[:400]}"
    if progress:
        progress(0.10, f"Đã quét {len(units)} đoạn văn ({rep.unique} đoạn duy nhất). Đang kết nối Gemini AI...")

    trans, degraded, cached_count = translate_all(
        translator, uniq, target, glossary, context,
        source_lang=source_lang, doc_type="docx",
        max_items=batch_size, rate_delay=rate_delay, use_cache=use_cache,
        progress=progress, log_func=log_func,
    )
    rep.cached = cached_count

    _log(f"🎨 Đang phục hồi in đậm/nghiêng, màu chữ, bảng biểu & hyperlink cho {len(units)} đoạn văn...")
    if progress:
        progress(0.92, f"Đang khôi phục định dạng Word cho {len(units)} đoạn văn...")

    for u in units:
        d = trans.get(u.tagged, u.tagged)
        _apply_docx_translation(u, d, target)
        if u.tagged in degraded:
            rep.degraded.append(f"{u.where}: “{u.plain[:30].strip()}…”")

    _log("💾 Đang đóng gói và lưu tệp Word (.docx) hoàn chỉnh...")
    if progress:
        progress(0.98, "Đang hoàn tất lưu tệp Word...")
    doc.save(dst)
    _log("🎉 Hoàn tất 100%! Tệp Word đã sẵn sàng tải về máy.")
    if progress:
        progress(1.0, "Hoàn tất 100%! Tệp sẵn sàng tải về.")
    return rep
