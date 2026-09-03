#!/usr/bin/env python3
"""Generate Basic/Styled/OnePage DOCX resumes from a person's resume-data.json.

Reusable across people: pass the person's folder (containing resume-data.json)
as the first CLI argument. Outputs are written into that same folder.

Usage:
    uv run --with python-docx generate_resumes.py "Karleigh Henegar"
"""
import json
import sys
from datetime import date
from pathlib import Path

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

if len(sys.argv) < 2:
    sys.exit("Usage: generate_resumes.py <person-folder>  (folder must contain resume-data.json)")

BASE = Path(sys.argv[1])
DATA = json.loads((BASE / "resume-data.json").read_text())

ACCENT = RGBColor(0x1F, 0x3A, 0x5F)
GRAY = RGBColor(0x55, 0x55, 0x55)


def set_cell_border_bottom(paragraph, color="1F3A5F", sz=6):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(sz))
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def daterange(item):
    start = item.get("start", "") or ""
    end = item.get("end", "") or ""
    if start and end and start == end:
        return start
    if start or end:
        return f"{start} – {end}"
    return ""


def item_title(item):
    if item.get("subtitle"):
        return f'{item["title"]} — {item["subtitle"]}'
    return item["title"]


def add_footer(doc, text, size=8.5, color=None):
    footer = doc.sections[0].footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.text = ""
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.italic = True
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color


CURRENT_YEAR = date.today().year


def job_sort_key(job):
    end = job.get("end") or job.get("start") or "0"
    if str(end).lower() == "present":
        return 9999
    try:
        return int(end)
    except ValueError:
        return 0


def is_condensed(job, ignore_override=False):
    """A basic-resume job is condensed to a header line when it's stale
    (ended more than 4 years ago) or has no engineering relevance, unless
    explicitly kept expanded (e.g. most-recent jobs worth showing in full).
    ignore_override=True forces full condensing regardless of that
    exception, for the maximally-condensed one-page variant."""
    if job.get("forceCondensed"):
        return True
    if job.get("keepBulletsOnBasic") and not ignore_override:
        return False
    if job.get("engineeringRelevant") is False:
        return True
    end = job.get("end") or job.get("start")
    try:
        return (CURRENT_YEAR - int(end)) > 4
    except (TypeError, ValueError):
        return False


def has_skill_categories():
    return "categories" in DATA.get("skills", {})


def skill_lines(skills_format):
    """Yield (heading, content_lines) pairs for the Skills section, honoring
    either the legacy technical/certifications lists or the newer
    categories[] schema (each with narrative + tags), rendered per
    skills_format ('narrative' or 'tags')."""
    skills = DATA["skills"]
    if not has_skill_categories():
        yield ("Technical Skills", [", ".join(skills["technical"])])
        return
    for cat in skills["categories"]:
        if skills_format == "tags":
            yield (cat["name"], cat["tags"])
        else:
            yield (cat["name"], [cat["narrative"]])


# ---------------------------------------------------------------------------
# BASIC VERSION
# ---------------------------------------------------------------------------

def build_basic(filename=None, combine_activities=False, full_condense=False, skills_format="narrative"):
    filename = filename or f'{DATA["name"]} Resume - Basic.docx'
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(12.5 if full_condense else 11)

    margin = 0.4 if full_condense else 0.5
    for section in doc.sections:
        section.top_margin = Inches(margin)
        section.bottom_margin = Inches(margin)
        section.left_margin = Inches(margin)
        section.right_margin = Inches(margin)

    name = doc.add_paragraph()
    name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name.paragraph_format.space_after = Pt(0)
    r = name.add_run(DATA["name"])
    r.bold = True
    r.font.size = Pt(15 if full_condense else 18)

    contact = doc.add_paragraph()
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if full_condense:
        contact.paragraph_format.space_after = Pt(4)
    loc = DATA["contact"].get("location")
    parts = [loc, DATA["contact"]["phone"], DATA["contact"]["email"]]
    contact.add_run("  |  ".join(p for p in parts if p))

    heading_before, heading_after = (6, 1) if full_condense else (12, 2)
    heading_size = 13.5 if full_condense else 12

    def heading(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(heading_before)
        p.paragraph_format.space_after = Pt(heading_after)
        r = p.add_run(text.upper())
        r.bold = True
        r.underline = True
        r.font.size = Pt(heading_size)

    def bullet(text):
        p = doc.add_paragraph(style="List Bullet")
        if full_condense:
            p.paragraph_format.space_after = Pt(1)
        p.add_run(text)

    def entry_line(title, dates):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.add_run(title).bold = True
        if dates:
            r = p.add_run(f"   ({dates})")
            r.italic = True
        return p

    if DATA.get("summary"):
        heading("Summary")
        doc.add_paragraph().add_run(DATA["summary"])

    heading("Skills")
    for name, lines in skill_lines(skills_format):
        p = doc.add_paragraph()
        p.add_run(f"{name}: ").bold = True
        if skills_format == "tags" and len(lines) > 1:
            p.add_run(" | ".join(lines))
        else:
            p.add_run(lines[0] if lines else "")
    p = doc.add_paragraph()
    p.add_run("Certifications: ").bold = True
    p.add_run(", ".join(DATA["skills"]["certifications"]))

    heading("Experience")
    for job in sorted(DATA["experience"], key=job_sort_key, reverse=True):
        entry_line(f'{job["title"]} — {job["organization"]}, {job["location"]}', daterange(job))
        condensed = is_condensed(job, ignore_override=full_condense)
        if not condensed:
            bullets = job.get("compactBullets", job["bullets"]) if full_condense else job["bullets"]
            for b in bullets:
                bullet(b)

    if DATA.get("capstoneAndResearch"):
        heading("Capstone & Research")
        for item in DATA["capstoneAndResearch"]:
            entry_line(item_title(item), daterange(item))
            bullets = item.get("compactBullets", item["bullets"]) if full_condense else item["bullets"]
            for b in bullets:
                bullet(b)

    heading("Education")
    for edu in DATA["education"]:
        entry_line(f'{edu["institution"]}, {edu["location"]}', daterange(edu))
        details = edu.get("compactDetails", edu["details"]) if full_condense else edu["details"]
        for d in details:
            bullet(d)

    if DATA.get("leadership"):
        if combine_activities:
            heading("Leadership & Volunteer")
            for role in sorted(DATA["leadership"], key=job_sort_key, reverse=True):
                bullet(f'{role["title"]} — {role["organization"]}')
            for v in DATA.get("volunteer", []):
                bullet(f'{v["role"]} — {v["organization"]}')
        else:
            heading("Leadership Roles")
            for role in DATA["leadership"]:
                entry_line(f'{role["title"]} — {role["organization"]}', daterange(role))
                for b in role["bullets"]:
                    bullet(b)

            for v in DATA.get("volunteer", []):
                bullet(f'{v["role"]}, {v["organization"]} — {v["hours"]}')

    if DATA.get("footer"):
        add_footer(doc, DATA["footer"])

    doc.save(BASE / filename)


# ---------------------------------------------------------------------------
# TWO-COLUMN LAYOUT (shared by Styled and One-Page)
# ---------------------------------------------------------------------------

TWIPS_PER_IN = 1440
SIDEBAR_IN = 2.5


def set_tc_width(cell, width_in):
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(width_in * TWIPS_PER_IN)))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)
    cell.width = Inches(width_in)


def set_tc_right_rule(cell, hex_color, sz=18):
    """Thin vertical divider instead of a solid fill, to stay printer/ink friendly."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    right = OxmlElement('w:right')
    right.set(qn('w:val'), 'single')
    right.set(qn('w:sz'), str(sz))
    right.set(qn('w:space'), '0')
    right.set(qn('w:color'), hex_color)
    tcBorders.append(right)
    tcPr.append(tcBorders)


def set_tc_margins(cell, top=180, start=220, bottom=180, end=220):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement('w:tcMar')
    for tag, val in (('top', top), ('start', start), ('bottom', bottom), ('end', end)):
        node = OxmlElement(f'w:{tag}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        mar.append(node)
    tcPr.append(mar)


def cell_para(cell, space_before=0, space_after=4):
    p0 = cell.paragraphs[-1]
    p = p0 if (p0.text == "" and not p0.runs) else cell.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    return p


def col_heading(cell, text, size=12):
    p = cell_para(cell, space_before=10, space_after=4)
    r = p.add_run(text.upper())
    r.bold = True
    r.font.size = Pt(size)
    r.font.color.rgb = ACCENT
    r.font.name = "Cambria"
    set_cell_border_bottom(p, color="B7C4D6", sz=6)


def col_entry_header(cell, title, dates, title_size=10.5, date_size=9.5):
    p = cell_para(cell, space_before=6, space_after=0)
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(title_size)
    if dates:
        r2 = p.add_run(f"   {dates}")
        r2.italic = True
        r2.font.size = Pt(date_size)
        r2.font.color.rgb = GRAY


def col_bullet(doc, cell, text, size=9.5, indent=14):
    p = cell_para(cell, space_after=2)
    p.style = doc.styles["List Bullet"]
    p.paragraph_format.left_indent = Pt(indent)
    r = p.add_run(text)
    r.font.size = Pt(size)


def col_dash(cell, text, size=9.5, color=None):
    p = cell_para(cell, space_after=2)
    p.paragraph_format.left_indent = Pt(10)
    r = p.add_run(f"–  {text}")
    r.font.size = Pt(size)
    r.font.color.rgb = color or RGBColor(0x33, 0x33, 0x33)


def col_line(cell, text, size=9.5, bold=False, italic=False, color=None, space_after=2):
    p = cell_para(cell, space_after=space_after)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    r.font.color.rgb = color or RGBColor(0x22, 0x22, 0x22)


def build_two_column(filename, combine_activities=False, condense_experience=False, skills_format="narrative"):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)

    section = doc.sections[0]
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    # --- Full-width header: name, title, contact all together ---
    name_p = doc.add_paragraph()
    name_p.paragraph_format.space_after = Pt(0)
    r = name_p.add_run(DATA["name"])
    r.bold = True
    r.font.size = Pt(19)
    r.font.color.rgb = ACCENT
    r.font.name = "Cambria"

    contact_p = doc.add_paragraph()
    contact_p.paragraph_format.space_after = Pt(8)
    parts = [DATA["contact"].get("headline"), DATA["contact"].get("location"), DATA["contact"]["phone"], DATA["contact"]["email"]]
    r = contact_p.add_run("   •   ".join(p for p in parts if p))
    r.italic = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = GRAY

    if DATA.get("summary"):
        summary_p = doc.add_paragraph()
        summary_p.paragraph_format.space_after = Pt(10)
        r = summary_p.add_run(DATA["summary"])
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0x22, 0x22, 0x22)

    # --- Two-column table: Skills/Certifications/Education/Experience on the left (narrow),
    #     Capstone/Leadership/Volunteer on the right (wide) ---
    usable_in = 8.5 - 1.0
    main_in = usable_in - SIDEBAR_IN

    table = doc.add_table(rows=1, cols=2)
    tblPr = table._tbl.tblPr
    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    tblPr.append(layout)
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'none')
        borders.append(el)
    tblPr.append(borders)

    sidebar, main = table.rows[0].cells
    set_tc_width(sidebar, SIDEBAR_IN)
    set_tc_width(main, main_in)
    set_tc_right_rule(sidebar, "1F3A5F")
    set_tc_margins(sidebar, start=0, end=200)
    set_tc_margins(main, start=420)

    # --- Left column: Skills, Certifications, Education, then Experience at the bottom ---
    if has_skill_categories():
        for name, lines in skill_lines(skills_format):
            col_heading(sidebar, name, size=10.5)
            if skills_format == "tags":
                for s in lines:
                    col_dash(sidebar, s)
            else:
                col_line(sidebar, lines[0], size=8.5)
    else:
        col_heading(sidebar, "Technical Skills", size=10.5)
        for s in DATA["skills"]["technicalGrouped"]:
            col_dash(sidebar, s)

    col_heading(sidebar, "Certifications", size=10.5)
    for c in DATA["skills"]["certifications"]:
        col_dash(sidebar, c)

    col_heading(sidebar, "Education", size=10.5)
    for edu in DATA["education"]:
        col_entry_header(sidebar, edu["institution"], daterange(edu), title_size=9.5, date_size=8.5)
        for d in edu["details"]:
            col_line(sidebar, d, size=8.5)

    has_extras = DATA.get("capstoneAndResearch") or DATA.get("leadership")

    def render_experience(cell, size=8.5, indent=10):
        for job in sorted(DATA["experience"], key=job_sort_key, reverse=True):
            col_entry_header(cell, f'{job["title"]} — {job["organization"]}, {job["location"]}', daterange(job), title_size=max(size, 9.5), date_size=max(size - 1, 8.5))
            condensed = condense_experience and is_condensed(job)
            if not condensed:
                bullets = job.get("compactBullets", job["bullets"]) if condense_experience else job["bullets"]
                for b in bullets:
                    col_bullet(doc, cell, b, size=size, indent=indent)

    if has_extras:
        # --- Student-style layout: Experience stays in the narrow sidebar; the
        #     wide main column showcases Capstone/Research and Leadership. ---
        col_heading(sidebar, "Experience", size=10.5)
        render_experience(sidebar, size=8.5, indent=10)

        if DATA.get("capstoneAndResearch"):
            col_heading(main, "Capstone & Research")
            for item in DATA["capstoneAndResearch"]:
                col_entry_header(main, item_title(item), daterange(item))
                for b in item["bullets"]:
                    col_bullet(doc, main, b)

        if combine_activities:
            col_heading(main, "Leadership & Volunteer")
            for role in sorted(DATA["leadership"], key=job_sort_key, reverse=True):
                col_bullet(doc, main, f'{role["title"]} — {role["organization"]}', size=9.5)
            for v in DATA.get("volunteer", []):
                col_bullet(doc, main, f'{v["role"]} — {v["organization"]}', size=9.5)
        else:
            col_heading(main, "Leadership & Activities")
            for role in sorted(DATA["leadership"], key=job_sort_key, reverse=True):
                col_entry_header(main, f'{role["title"]} — {role["organization"]}', daterange(role))
                for b in role["bullets"]:
                    col_bullet(doc, main, b)
            for v in DATA.get("volunteer", []):
                col_bullet(doc, main, f'{v["role"]}, {v["organization"]} — {v["hours"]}')
    else:
        # --- Professional-experience layout: Experience is the centerpiece,
        #     so it runs in the wide main column instead of the narrow sidebar. ---
        col_heading(main, "Experience")
        render_experience(main, size=9.5, indent=14)

    if DATA.get("footer"):
        add_footer(doc, DATA["footer"], color=GRAY)

    doc.save(BASE / filename)


if __name__ == "__main__":
    name = DATA["name"]
    if has_skill_categories():
        for fmt, suffix in (("narrative", "Narrative"), ("tags", "Tags")):
            build_basic(f"{name} Resume - Basic - {suffix}.docx", combine_activities=False, skills_format=fmt)
            build_basic(f"{name} Resume - Basic OnePage - {suffix}.docx", combine_activities=True, full_condense=True, skills_format=fmt)
            build_two_column(f"{name} Resume - Styled - {suffix}.docx", combine_activities=False, condense_experience=True, skills_format=fmt)
            build_two_column(f"{name} Resume - Styled OnePage - {suffix}.docx", combine_activities=True, condense_experience=True, skills_format=fmt)
    else:
        build_basic(f"{name} Resume - Basic.docx", combine_activities=False)
        build_basic(f"{name} Resume - Basic OnePage.docx", combine_activities=True, full_condense=True)
        build_two_column(f"{name} Resume - Styled.docx", combine_activities=False, condense_experience=True)
        build_two_column(f"{name} Resume - Styled OnePage.docx", combine_activities=True, condense_experience=True)
    print("done")
