"""Bóc tách cấu trúc XML PowerPoint và phục hồi định dạng chi tiết."""
from __future__ import annotations

import copy
import html
import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Optional

from lxml import etree
from pptx.oxml.ns import qn
from pptx.shapes.group import GroupShape

from src.config import LANGS

TAG_RE = re.compile(r"(</?r\d+>|<br/>|<f\d+/>)")
STYLE_TAG_RE = re.compile(r"</?r\d+>")
_LETTER = re.compile(r"[^\W\d_]")
_SKIP = re.compile(r"^(\S+@\S+|https?://\S+|www\.\S+)$")
_EA_AFTER = ("cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst")


@dataclass
class Unit:
    p: etree._Element                 # <a:p>
    owner: etree._Element             # <p:txBody> / <a:txBody>
    kind: str                         # shape | cell | notes | layout
    where: str                        # "Slide 3" ...
    parts: list = field(default_factory=list)
    rpr_by_id: dict = field(default_factory=dict)
    sig_to_id: dict = field(default_factory=dict)
    fields: dict = field(default_factory=dict)
    brs: list = field(default_factory=list)
    default_id: int = 1
    use_tags: bool = False
    tagged: str = ""
    plain: str = ""
    translated_plain: str = ""
    degraded: bool = False


@dataclass
class Report:
    units: int = 0
    unique: int = 0
    cached: int = 0
    degraded: list = field(default_factory=list)
    fitted: list = field(default_factory=list)
    warnings: list = field(default_factory=list)

    def text(self) -> str:
        lines = [f"Đã dịch {self.units} đoạn ({self.unique} đoạn duy nhất)."]
        if self.cached:
            lines.append(f"⚡ Tái sử dụng từ Bộ nhớ: {self.cached} đoạn (Tiết kiệm {self.cached} lượt gọi API).")
        if self.fitted:
            lines.append(f"Co chữ chống tràn ở {len(self.fitted)} khung: " + "; ".join(self.fitted[:15]))
        if self.degraded:
            lines.append("Mất định dạng từng-từ (fallback) ở: " + "; ".join(self.degraded[:15]))
        lines += self.warnings
        return "\n".join(lines)


def _sig(rPr) -> str:
    if rPr is None:
        return ""
    c = copy.deepcopy(rPr)
    for a in ("lang", "altLang", "dirty", "err", "noProof", "smtClean", "smtId"):
        c.attrib.pop(a, None)
    return etree.tostring(c, method="c14n").decode()


def parse_paragraph(p, owner, kind, where) -> Optional[Unit]:
    parts: list[list] = []
    rpr_by_sig: dict[str, object] = {}
    for c in p:
        if c.tag == qn("a:r"):
            t = c.find(qn("a:t"))
            text = (t.text or "") if t is not None else ""
            if text == "":
                continue
            rPr = c.find(qn("a:rPr"))
            sig = _sig(rPr)
            rpr_by_sig.setdefault(sig, rPr)
            if parts and parts[-1][0] == "seg" and (parts[-1][2] == sig or not text.strip()):
                parts[-1][1] += text
            else:
                parts.append(["seg", text, sig])
        elif c.tag == qn("a:br"):
            parts.append(["br", c, ""])
        elif c.tag == qn("a:fld"):
            parts.append(["f", c, ""])

    segs = [x for x in parts if x[0] == "seg"]
    if not segs:
        return None
    for i, x in enumerate(parts):
        if x[0] == "seg" and not x[1].strip():
            nxt = next((y for y in parts[i + 1:] if y[0] == "seg" and y[1].strip()), None)
            prv = next((y for y in reversed(parts[:i]) if y[0] == "seg" and y[1].strip()), None)
            ref = nxt or prv
            if ref:
                x[2] = ref[2]

    u = Unit(p=p, owner=owner, kind=kind, where=where)
    for x in parts:
        if x[0] == "seg" and x[2] not in u.sig_to_id:
            n = len(u.sig_to_id) + 1
            u.sig_to_id[x[2]] = n
            u.rpr_by_id[n] = rpr_by_sig.get(x[2])
    u.use_tags = len(u.sig_to_id) > 1

    out, plain, fn = [], [], 0
    longest = {}
    for kind_, payload, sig in parts:
        if kind_ == "seg":
            rid = u.sig_to_id[sig]
            longest[rid] = longest.get(rid, 0) + len(payload)
            plain.append(payload)
            esc = html.escape(payload, quote=False)
            out.append(f"<r{rid}>{esc}</r{rid}>" if u.use_tags else esc)
        elif kind_ == "br":
            u.brs.append(payload)
            plain.append("\n")
            out.append("<br/>")
        else:
            fn += 1
            u.fields[fn] = payload
            out.append(f"<f{fn}/>")
    u.default_id = max(longest, key=longest.get)
    u.tagged = "".join(out)
    u.plain = "".join(plain)
    return u


def needs_translation(u: Unit) -> bool:
    t = u.plain.strip()
    return bool(_LETTER.search(t)) and not _SKIP.match(t)


def _walk(shapes):
    for shp in shapes:
        if isinstance(shp, GroupShape):
            yield from _walk(shp.shapes)
            continue
        if shp.has_text_frame:
            yield "shape", shp.text_frame._txBody
        if getattr(shp, "has_table", False) and shp.has_table:
            for row in shp.table.rows:
                for cell in row.cells:
                    if not cell.is_spanned:
                        yield "cell", cell.text_frame._txBody
        if getattr(shp, "has_chart", False) and shp.has_chart:
            yield "chart", shp
        elif shp.element.tag.endswith("graphicFrame") and b"diagram" in etree.tostring(shp.element):
            yield "diagram", shp
        elif shp.shape_type is not None and "PICTURE" in str(shp.shape_type):
            yield "picture", shp


def collect_units(prs, include_layouts=False):
    units: list[Unit] = []
    notes: list[str] = []
    counts = Counter()

    def add(tx, kind, where):
        for p in tx.findall(qn("a:p")):
            u = parse_paragraph(p, tx, kind, where)
            if u and needs_translation(u):
                units.append(u)

    for i, slide in enumerate(prs.slides, 1):
        for kind, obj in _walk(slide.shapes):
            if kind in ("shape", "cell"):
                add(obj, kind, f"Slide {i}")
            else:
                counts[kind] += 1
                if kind in ("chart", "diagram"):
                    notes.append(f"Slide {i}: có {kind} chưa được dịch.")
        if slide.has_notes_slide:
            add(slide.notes_slide.notes_text_frame._txBody, "notes", f"Slide {i} (ghi chú)")

    if include_layouts:
        for holder in list(prs.slide_masters) + [l for m in prs.slide_masters for l in m.slide_layouts]:
            for shp in holder.shapes:
                if shp.has_text_frame and not shp.is_placeholder:
                    add(shp.text_frame._txBody, "layout", "Master/Layout")
    return units, notes


def _tags(s: str, style=True) -> Counter:
    tags = TAG_RE.findall(s)
    if not style:
        tags = [t for t in tags if not STYLE_TAG_RE.fullmatch(t)]
    return Counter(tags)


def valid(src: str, dst: str, style=True) -> bool:
    return bool(dst and dst.strip()) and _tags(src, style) == _tags(dst, style)


def _set_ea(rPr, typeface):
    for old in rPr.findall(qn("a:ea")):
        rPr.remove(old)
    ea = etree.Element(qn("a:ea"))
    ea.set("typeface", typeface)
    latin = rPr.find(qn("a:latin"))
    if latin is not None:
        latin.addnext(ea)
        return
    for tag in _EA_AFTER:
        nxt = rPr.find(qn(f"a:{tag}"))
        if nxt is not None:
            nxt.addprevious(ea)
            return
    rPr.append(ea)


def _make_run(src_rPr, text, lang, ea_font):
    r = etree.Element(qn("a:r"))
    rPr = copy.deepcopy(src_rPr) if src_rPr is not None else etree.Element(qn("a:rPr"))
    for a in ("err", "dirty"):
        rPr.attrib.pop(a, None)
    rPr.set("lang", lang)
    if ea_font:
        rPr.set("altLang", "en-US")
        _set_ea(rPr, ea_font)
    r.append(rPr)
    t = etree.SubElement(r, qn("a:t"))
    t.text = text
    return r


def apply_translation(u: Unit, translated: str, target: str):
    _, lang, ea_font = LANGS[target]
    cur, new = u.default_id, []
    brs = list(u.brs)
    for tok in TAG_RE.split(translated):
        if not tok:
            continue
        if m := re.fullmatch(r"<r(\d+)>", tok):
            cur = int(m.group(1)) if int(m.group(1)) in u.rpr_by_id else u.default_id
        elif re.fullmatch(r"</r\d+>", tok):
            cur = u.default_id
        elif tok == "<br/>":
            new.append(copy.deepcopy(brs.pop(0)) if brs else etree.Element(qn("a:br")))
        elif m := re.fullmatch(r"<f(\d+)/>", tok):
            if int(m.group(1)) in u.fields:
                new.append(u.fields[int(m.group(1))])
        else:
            new.append(_make_run(u.rpr_by_id.get(cur), html.unescape(tok), lang, ea_font))
            
    for c in list(u.p):
        if c.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
            u.p.remove(c)
    end = u.p.find(qn("a:endParaRPr"))
    for el in new:
        end.addprevious(el) if end is not None else u.p.append(el)
    u.translated_plain = html.unescape(STYLE_TAG_RE.sub("", TAG_RE.sub(lambda m: "" if m.group(0).startswith("<f") else (" " if m.group(0) == "<br/>" else m.group(0)), translated)))

