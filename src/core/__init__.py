"""Core PPTX translation & layout engine."""
from src.core.autofit import fit_text, _width
from src.core.parser import Unit, Report, collect_units, apply_translation
from src.core.translator import translate_pptx, translate_all

