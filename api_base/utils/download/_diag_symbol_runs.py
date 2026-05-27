# -*- coding: utf-8 -*-
import os
import re
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

import docx
from pdf2docx import Converter
from app.services.document_v2.pdf_docx_pipeline.layout_recovery import (
    sanitize_converted_docx,
    _read_run_font_name,
)

pdf_path = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "app", "uploads", "HoangPhuc222683.pdf")
)

out = open(os.path.join(os.path.dirname(__file__), "_diag_symbol_runs.txt"), "w", encoding="utf-8")

with tempfile.TemporaryDirectory() as td:
    raw_docx = os.path.join(td, "raw.docx")
    san_docx = os.path.join(td, "san.docx")
    cv = Converter(pdf_path)
    cv.convert(raw_docx)
    cv.close()

    import shutil
    shutil.copy(raw_docx, san_docx)
    sanitize_converted_docx(san_docx, pdf_path=pdf_path, layout_mode="academic")
    doc = docx.Document(san_docx)

    for i, p in enumerate(doc.paragraphs):
        txt = "".join(r.text or "" for r in p.runs)
        if not txt.strip():
            continue
        if "KhachHang" in txt or "ThanhToan" in txt or "OnlinePayment" in txt or "COD" in txt:
            out.write(f"\n=== p{i} {txt[:100]!r} ===\n")
            for j, r in enumerate(p.runs):
                t = r.text or ""
                if not t:
                    continue
                fn = _read_run_font_name(r)
                pua = [hex(ord(c)) for c in t if 0xF020 <= ord(c) <= 0xF0FF]
                out.write(f"  r{j} font={fn!r} text={t!r} pua={pua}\n")

out.close()
print("done")
