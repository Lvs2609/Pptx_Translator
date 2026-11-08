"""Tự động nhận diện bố cục và co chữ thông minh (Auto-fitting)."""
from typing import Optional
from lxml import etree
from pptx.oxml.ns import qn


def _width(s: str) -> float:
    """Đo độ dài chuỗi (ký tự Đông Á CJK tính độ rộng 1.0, ký tự Latin tính 0.5)."""
    return sum(1.0 if ord(ch) >= 0x2E80 else 0.5 for ch in s)


def fit_text(owner, kind: str, ratio: float, min_scale: float) -> Optional[float]:
    """Tự động co cỡ font chữ vừa khít khung khi bản dịch bị dài ra."""
    s = max(min_scale, ratio ** -0.5)
    s = round(s * 40) / 40
    
    if kind in ("shape", "layout"):
        bodyPr = owner.find(qn("a:bodyPr"))
        if bodyPr is None:
            return None
        for tag in ("noAutofit", "normAutofit", "spAutoFit"):
            for el in bodyPr.findall(qn(f"a:{tag}")):
                bodyPr.remove(el)
        fit = etree.Element(qn("a:normAutofit"))
        fit.set("fontScale", str(int(s * 100000)))
        bodyPr.insert(1 if bodyPr.find(qn("a:prstTxWarp")) is not None else 0, fit)
        return s
        
    if kind == "cell":
        found = False
        for rPr in owner.iter(qn("a:rPr"), qn("a:endParaRPr")):
            current_sz = int(rPr.get("sz")) if rPr.get("sz") else 1400
            rPr.set("sz", str(max(100, int(current_sz * s))))
            found = True
        return s if found else None
        
    return None

