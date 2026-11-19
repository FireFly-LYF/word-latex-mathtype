# -*- coding: utf-8 -*-
"""Unit tests for delimited-math signature gate (no Word required)."""
from __future__ import annotations

import importlib.util
import unittest
from pathlib import Path


def _load():
    path = Path(__file__).with_name("convert_docx_latex_to_equations.py")
    spec = importlib.util.spec_from_file_location("convert_docx_latex_to_equations", path)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    return mod


class FormulaSignatureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.m = _load()

    def test_delimited_simple_idents_in_cjk(self):
        body = "参照回波建模思路，第$i$部雷达对第$m$个脉冲下变频后的接收信号可表示为"
        self.assertTrue(self.m.has_delimited_math(body))
        self.assertTrue(self.m.has_formula_signature(body))
        matches = self.m.get_inline_formula_matches(body)
        self.assertEqual({c["text"] for c in matches}, {"i", "m"})
        self.assertFalse(self.m.is_display_formula_paragraph(body))

    def test_delimited_k_i_short_sentence(self):
        body = "同理，友邻雷达$k$经目标散射与直达路径进入雷达$i$的互扰瞄准概率分别写为"
        self.assertTrue(self.m.has_formula_signature(body))
        matches = self.m.get_inline_formula_matches(body)
        self.assertEqual({c["text"] for c in matches}, {"k", "i"})

    def test_bare_dollar_i_not_display(self):
        self.assertTrue(self.m.has_formula_signature("$i$"))
        self.assertFalse(self.m.is_display_formula_paragraph("$i$"))

    def test_plain_prose_still_false(self):
        self.assertFalse(self.m.has_formula_signature("无公式的中文句子。"))
        self.assertFalse(self.m.has_delimited_math("无公式的中文句子。"))

    def test_backslash_signature_unchanged(self):
        self.assertTrue(self.m.has_formula_signature(r"$\alpha_{i}$"))
        self.assertTrue(self.m.has_formula_signature(r"$$\mathrm{SIJNR}_{i,m}=1$$"))


if __name__ == "__main__":
    unittest.main()
