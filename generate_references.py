#!/usr/bin/env python3
"""Generate a References doc styled to match the Basic resume header, from
a person's resume-data.json.

Usage:
    uv run --with python-docx generate_references.py "Karleigh Henegar"
"""
import json
import sys
from pathlib import Path

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

if len(sys.argv) < 2:
    sys.exit("Usage: generate_references.py <person-folder>  (folder must contain resume-data.json)")

BASE = Path(sys.argv[1])
DATA = json.loads((BASE / "resume-data.json").read_text())

ACCENT = RGBColor(0x1F, 0x3A, 0x5F)
GRAY = RGBColor(0x55, 0x55, 0x55)


def build_references():
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    name = doc.add_paragraph()
    name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = name.add_run(DATA["name"])
    r.bold = True
    r.font.size = Pt(18)

    contact = doc.add_paragraph()
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    loc = DATA["contact"].get("location")
    parts = [loc, DATA["contact"]["phone"], DATA["contact"]["email"]]
    contact.add_run("  |  ".join(p for p in parts if p))

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(14)
    title.paragraph_format.space_after = Pt(10)
    r = title.add_run("PROFESSIONAL REFERENCES")
    r.bold = True
    r.underline = True
    r.font.size = Pt(13)

    for i, ref in enumerate(DATA["references"]):
        if i > 0:
            spacer = doc.add_paragraph()
            spacer.paragraph_format.space_after = Pt(6)

        name_p = doc.add_paragraph()
        name_p.paragraph_format.space_after = Pt(0)
        r = name_p.add_run(ref["name"])
        r.bold = True
        r.font.size = Pt(12)

        role_p = doc.add_paragraph()
        role_p.paragraph_format.space_after = Pt(0)
        role_line = ref["title"]
        if ref.get("organization"):
            role_line += f', {ref["organization"]}'
        role_p.add_run(role_line)

        rel_p = doc.add_paragraph()
        rel_p.paragraph_format.space_after = Pt(0)
        r = rel_p.add_run(f'Relationship: {ref["relationship"]}')
        r.italic = True
        r.font.color.rgb = GRAY

        contact_bits = []
        if ref.get("email"):
            contact_bits.append(f'Email: {ref["email"]}')
        if ref.get("phone"):
            contact_bits.append(f'Phone: {ref["phone"]}')
        if contact_bits:
            c_p = doc.add_paragraph()
            c_p.add_run("   |   ".join(contact_bits))

    doc.save(BASE / f'{DATA["name"]} - References.docx')


if __name__ == "__main__":
    build_references()
    print("done")
