#!/usr/bin/env python3
"""Generate a reusable cover letter template, styled to match the resumes,
from a person's resume-data.json. The Johnson & Johnson Field Service
Engineer letter in this file is tailored content specific to that one
application (Karleigh's) — reuse the template builder for other people/roles.

Usage:
    uv run --with python-docx generate_cover_letters.py "Karleigh Henegar"
"""
import json
import sys
from datetime import date
from pathlib import Path

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

if len(sys.argv) < 2:
    sys.exit("Usage: generate_cover_letters.py <person-folder>  (folder must contain resume-data.json)")

BASE = Path(sys.argv[1])
DATA = json.loads((BASE / "resume-data.json").read_text())

ACCENT = RGBColor(0x1F, 0x3A, 0x5F)
GRAY = RGBColor(0x55, 0x55, 0x55)

TODAY = date.today().strftime("%B %d, %Y")


def new_letter_doc():
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    name_p = doc.add_paragraph()
    name_p.paragraph_format.space_after = Pt(0)
    r = name_p.add_run(DATA["name"])
    r.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = ACCENT

    contact_p = doc.add_paragraph()
    contact_p.paragraph_format.space_after = Pt(16)
    parts = [DATA["contact"].get("location"), DATA["contact"]["phone"], DATA["contact"]["email"]]
    r = contact_p.add_run("   |   ".join(p for p in parts if p))
    r.font.size = Pt(10)
    r.font.color.rgb = GRAY

    return doc


def add_note(doc, text):
    """A bracketed, italic instructional note meant to be deleted before sending."""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(text)
    r.italic = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0x99, 0x33, 0x33)


def add_body_paragraph(doc, text, space_after=10):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.add_run(text)
    return p


def build_template():
    doc = new_letter_doc()

    add_note(
        doc,
        "[TEMPLATE — replace every bracketed placeholder below, then delete this note and any "
        "unused instructions before sending. Keep it to one page.]",
    )

    doc.add_paragraph("[Date]")
    doc.add_paragraph()
    doc.add_paragraph("[Hiring Manager Name, or \"Hiring Team\" if unknown]")
    doc.add_paragraph("[Company Name]")
    doc.add_paragraph("[Company Address, if known]")
    doc.add_paragraph()

    doc.add_paragraph("Dear [Hiring Manager Name / Hiring Team],")
    doc.add_paragraph()

    add_body_paragraph(
        doc,
        "I am writing to express my interest in the [Position Title] role at [Company Name], "
        "as advertised on [where you found the posting]. As a Biomedical Engineering graduate "
        "from the University of Alabama at Birmingham with hands-on experience in [1-2 skills "
        "most relevant to this specific role], I am excited about the opportunity to bring my "
        "technical background and problem-solving skills to your team.",
    )

    add_body_paragraph(
        doc,
        "[Paragraph 2 — pick ONE accomplishment that best matches this job's requirements and "
        "describe it concretely: what you did, what tools/skills you used, and the measurable "
        "result. Good source material: the Capstone Project with Children's of Alabama "
        "(fabrication, SolidWorks, welding, machining, working directly with a clinical client), "
        "the Honors Thesis (independent research, data analysis, technical writing, presenting "
        "to a faculty committee), or your CNA/Resident Assistant experience (patient care, "
        "communication, working under pressure in clinical or high-contact environments).]",
    )

    add_body_paragraph(
        doc,
        "[Paragraph 3 — connect your skills to what this specific employer does or values. "
        "Reference something specific about the company/product/mission from the job posting "
        "or their website, and explain why it appeals to you. Avoid generic statements — this "
        "paragraph should read differently for every application.]",
    )

    add_body_paragraph(
        doc,
        "Thank you for considering my application. I would welcome the opportunity to discuss "
        "how my background in [field/skill] can contribute to [Company Name]'s [team/goal/"
        "mission]. I am available at your convenience for an interview and can be reached at "
        f"{DATA['contact']['phone']} or {DATA['contact']['email']}.",
    )

    doc.add_paragraph()
    doc.add_paragraph("Sincerely,")
    doc.add_paragraph(DATA["name"])

    doc.save(BASE / "Cover Letter - Template.docx")


def build_jnj_fse_letter():
    """Field Service Engineer - Medical Device Equipment, Minneapolis, MN.
    Tailored to the actual JD shared via a Real Staffing (SThree) recruiter,
    which named Johnson & Johnson explicitly (2026-09-02). Framed around the
    JD's three stated pillars: install/maintain, troubleshoot/repair, and
    customer service."""
    doc = new_letter_doc()

    doc.add_paragraph(TODAY)
    doc.add_paragraph()
    doc.add_paragraph("Hiring Team")
    doc.add_paragraph("Johnson & Johnson")
    doc.add_paragraph()

    doc.add_paragraph("Dear Hiring Team,")
    doc.add_paragraph()

    add_body_paragraph(
        doc,
        "I am writing to express my interest in the Field Service Engineer - Medical Device "
        "Equipment position in Minneapolis. As a Biomedical Engineering graduate from the "
        "University of Alabama at Birmingham with hands-on experience installing, fabricating, "
        "and troubleshooting equipment in real clinical settings, I am drawn to a role that "
        "combines technical problem-solving with direct, in-person support for the healthcare "
        "professionals and patients who rely on that equipment every day. I am also based just "
        "outside Minneapolis in Woodbury, MN, and available to start on-site work right away.",
    )

    add_body_paragraph(
        doc,
        "During my senior capstone project, I worked directly with a clinical client from "
        "Children's of Alabama to design and fabricate a support bracket for operating-room "
        "medical devices. I led the fabrication and prototyping logistics using SolidWorks, "
        "carpentry, welding, and machining, and delivered a bracket that holds two Abbott "
        "CentriMags in a more compact configuration for OR use. That project required exactly "
        "the kind of work I understand this role involves: installing and configuring equipment "
        "so it performs reliably in the room where clinicians actually use it, and communicating "
        "clearly with clinical staff throughout the process.",
    )

    add_body_paragraph(
        doc,
        "That same project also demanded the diagnostic mindset this role calls for — working "
        "through mechanical constraints, iterating on a design under real physical limitations, "
        "and troubleshooting until the solution worked in practice, not just on paper. My Honors "
        "Thesis research added a further layer of analytical rigor: collecting and interpreting "
        "data, documenting my process, and defending my conclusions to a faculty committee. "
        "Between that project work and my proficiency in R, MATLAB, SolidWorks, and LTSpice, I "
        "am comfortable diagnosing and resolving technical problems methodically and "
        "documenting them clearly.",
    )

    add_body_paragraph(
        doc,
        "I also bring direct experience managing customer and patient expectations under "
        "pressure: I spent several summers as a Certified Nursing Assistant, and I currently "
        "serve as a Resident Assistant, both roles that required clear communication before and "
        "after every interaction, composure under pressure, and accountability to the people I "
        "was serving. My CPR and BLS certifications, along with my wet lab and engineering "
        "design experience, give me a practical foundation for working safely and effectively "
        "around medical equipment and the clinicians who depend on it.",
    )

    add_body_paragraph(
        doc,
        "Thank you for considering my application. I would welcome the opportunity to discuss "
        "how my background in hands-on medical device fabrication, technical troubleshooting, "
        "and patient- and customer-facing service can contribute to your field service team. I "
        f"am available at your convenience for an interview and can be reached at "
        f"{DATA['contact']['phone']} or {DATA['contact']['email']}.",
    )

    doc.add_paragraph()
    doc.add_paragraph("Sincerely,")
    doc.add_paragraph(DATA["name"])

    doc.save(BASE / "Cover Letter - Field Service Engineer (Johnson & Johnson - Medical Devices).docx")


if __name__ == "__main__":
    build_template()
    build_jnj_fse_letter()
    print("done")
