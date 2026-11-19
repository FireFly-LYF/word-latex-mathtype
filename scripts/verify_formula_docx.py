# -*- coding: utf-8 -*-
"""Verify a converted .docx contains MathType OLE and few leftovers.

Exit codes:
  0 = pass
  2 = failed gates
  1 = usage / IO error
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path


def inspect_docx(path: Path) -> dict:
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", errors="replace")
        embeds = [n for n in z.namelist() if n.startswith("word/embeddings/")]
    text = re.sub(r"<[^>]+>", "", xml)
    mteqn = len(re.findall(r"SEQ\s+MTEqn", xml, flags=re.I))
    # Plain Word-typed (n) near equations: digits in parentheses in visible text,
    # excluding MathType field scaffolding. Rough count of leftover typed numbers.
    word_eqnum = len(re.findall(r"[\(（]\s*\d+\s*[\)）]", text))
    return {
        "path": str(path.resolve()),
        "size": path.stat().st_size,
        "omath": len(re.findall(r"<m:oMath[\s>]", xml)),
        "omathPara": len(re.findall(r"<m:oMathPara[\s>]", xml)),
        "dsmt4": len(re.findall(r"Equation\.DSMT4", xml)),
        "embeddings": len(embeds),
        "dollar_pairs": len(re.findall(r"\$[^$\n]{1,200}\$", text)),
        "helper_marker": "⟦WLMT⟧" in text,
        "mteqn_seq": mteqn,
        "word_eqnum_text": word_eqnum,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Verify MathType conversion output docx.")
    parser.add_argument("--docx", type=Path, required=True)
    parser.add_argument("--min-dsmt4", type=int, default=1, help="Minimum Equation.DSMT4 count")
    parser.add_argument("--max-omath", type=int, default=0, help="Maximum remaining m:oMath")
    parser.add_argument("--max-dollar-pairs", type=int, default=0, help="Maximum leftover $...$ pairs")
    parser.add_argument(
        "--min-mteqn",
        type=int,
        default=0,
        help="Minimum MathType SEQ MTEqn field count (right numbers)",
    )
    parser.add_argument(
        "--max-word-eqnum",
        type=int,
        default=None,
        help="Maximum plain Word-typed (n) leftovers (optional)",
    )
    parser.add_argument("--json-out", type=Path, help="Optional JSON report path")
    args = parser.parse_args(argv)

    docx = args.docx.resolve()
    if not docx.is_file():
        print(f"not found: {docx}", file=sys.stderr)
        return 1

    info = inspect_docx(docx)
    gates = {
        "dsmt4_ok": info["dsmt4"] >= args.min_dsmt4,
        "omath_ok": info["omath"] <= args.max_omath,
        "dollar_ok": info["dollar_pairs"] <= args.max_dollar_pairs,
        "helper_ok": not info["helper_marker"],
        "mteqn_ok": info["mteqn_seq"] >= args.min_mteqn,
    }
    if args.max_word_eqnum is not None:
        gates["word_eqnum_ok"] = info["word_eqnum_text"] <= args.max_word_eqnum
    report = {"inspect": info, "gates": gates, "pass": all(gates.values())}
    if args.json_out:
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        args.json_out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["pass"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
