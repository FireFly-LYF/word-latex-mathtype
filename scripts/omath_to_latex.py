# -*- coding: utf-8 -*-
"""Convert Word OMML (OMath) fragments to LaTeX via Office XSLT + pandoc.

Used by the OMath → MathType path: OMML → MathML → LaTeX text → TeX Toggle.
"""
from __future__ import annotations

import re
import shutil
import subprocess
from functools import lru_cache
from pathlib import Path

from lxml import etree

OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
OMML_NSMAP = {"m": OMML_NS}

OMML2MML_CANDIDATES = (
    Path(r"C:\Program Files\Microsoft Office\root\Office16\OMML2MML.XSL"),
    Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\OMML2MML.XSL"),
    Path(r"C:\Program Files\Microsoft Office\Office16\OMML2MML.XSL"),
)


def resolve_omml2mml_xsl(explicit: Path | None = None) -> Path:
    if explicit is not None:
        explicit = explicit.resolve()
        if not explicit.is_file():
            raise FileNotFoundError(f"OMML2MML.XSL not found: {explicit}")
        return explicit
    for candidate in OMML2MML_CANDIDATES:
        if candidate.is_file():
            return candidate
    raise FileNotFoundError(
        "Could not locate OMML2MML.XSL under Microsoft Office. Pass --omml2mml-xsl."
    )


def require_pandoc() -> str:
    path = shutil.which("pandoc")
    if not path:
        raise RuntimeError(
            "pandoc not found on PATH. Install pandoc for OMath→LaTeX conversion "
            "(https://pandoc.org/), or add it to PATH."
        )
    return path


@lru_cache(maxsize=4)
def _xslt_transform(xsl_path: str):
    return etree.XSLT(etree.parse(xsl_path))


def extract_omath_elements(word_open_xml: str) -> list[etree._Element]:
    root = etree.fromstring(word_open_xml.encode("utf-8"))
    return list(root.xpath("//m:oMath", namespaces=OMML_NSMAP))


def omml_element_to_mathml(om_el: etree._Element, *, xsl_path: Path | None = None) -> str:
    xsl = resolve_omml2mml_xsl(xsl_path)
    transform = _xslt_transform(str(xsl))
    result = transform(om_el)
    return etree.tostring(result.getroot(), encoding="unicode")


def mathml_to_latex(mathml_xml: str, *, pandoc_exe: str | None = None) -> str:
    pandoc = pandoc_exe or require_pandoc()
    html = f"<!DOCTYPE html><html><body>{mathml_xml}</body></html>"
    proc = subprocess.run(
        [pandoc, "-f", "html", "-t", "latex"],
        input=html,
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=False,
    )
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "").strip()
        raise RuntimeError(f"pandoc MathML→LaTeX failed: {err or proc.returncode}")
    return clean_pandoc_latex(proc.stdout or "")


def clean_pandoc_latex(tex: str) -> str:
    """Normalize pandoc latex math into a bare formula body for $...$ / $$...$$."""
    text = (tex or "").strip()
    if not text:
        return ""
    for _ in range(3):
        nxt = re.sub(r"^\\\((.*)\\\)$", r"\1", text, flags=re.S)
        nxt = re.sub(r"^\\\[(.*)\\\]$", r"\1", nxt, flags=re.S)
        nxt = re.sub(r"^\$\$(.*)\$\$$", r"\1", nxt, flags=re.S)
        nxt = re.sub(r"^\$(.*)\$$", r"\1", nxt, flags=re.S)
        nxt = nxt.strip()
        if nxt == text:
            break
        text = nxt
    text = re.sub(r"(\\\$)+\s*$", "", text).rstrip()
    text = text.rstrip("$").strip()
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def sanitize_mathtype_latex(tex: str) -> str:
    """Make pandoc OMML→LaTeX output friendlier for MathType TeX Toggle."""
    t = clean_pandoc_latex(tex)
    if not t:
        return t
    # pandoc accent artifacts: \left. x \right.\hat{} -> \hat{x}
    t = re.sub(
        r"\\left\.\s*([A-Za-z\\][A-Za-z0-9\\{\}^_]*)\s*\\right\.\\hat\{\}",
        r"\\hat{\1}",
        t,
    )
    t = t.replace(r"{\mathcal{l}}", r"\ell").replace(r"\mathcal{l}", r"\ell")
    t = t.replace(r"\middle|", r"|")
    t = re.sub(r"\\left\.\s*", "", t)
    t = re.sub(r"\s*\\right\.", "", t)
    t = t.replace(r"~~", r"\quad ")
    # Prefer plain fences; MathType TeX accepts these reliably
    t = t.replace(r"\left(", "(").replace(r"\right)", ")")
    t = t.replace(r"\left[", "[").replace(r"\right]", "]")
    t = t.replace(r"\left\{", r"\{").replace(r"\right\}", r"\}")
    t = t.replace(r"\left|", "|").replace(r"\right|", "|")
    t = re.sub(r"\(\s+", "(", t)
    t = re.sub(r"\s+\)", ")", t)
    t = re.sub(r"[ \t]+", " ", t).strip()
    return t


def word_open_xml_to_latex(
    word_open_xml: str,
    *,
    xsl_path: Path | None = None,
    pandoc_exe: str | None = None,
    sanitize: bool = True,
) -> str:
    elements = extract_omath_elements(word_open_xml)
    if not elements:
        raise ValueError("No m:oMath element found in WordOpenXML.")
    mathml = omml_element_to_mathml(elements[0], xsl_path=xsl_path)
    latex = mathml_to_latex(mathml, pandoc_exe=pandoc_exe)
    if sanitize:
        latex = sanitize_mathtype_latex(latex)
    if not latex:
        raise ValueError("Empty LaTeX after OMML→MathML→pandoc.")
    return latex
