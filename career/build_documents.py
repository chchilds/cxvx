#!/usr/bin/env python3
"""Build the aligned Chris Childs career document set from career/facts.py."""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as pdfcanvas

import facts as F

HERE = Path(__file__).resolve().parent
NAVY = RGBColor(0x1B, 0x36, 0x5D)
BLACK = RGBColor(0x22, 0x22, 0x22)
GRAY = RGBColor(0x44, 0x44, 0x44)
RULE = "1B365D"
PDF_NAVY = HexColor("#1B365D")
PDF_BLACK = HexColor("#222222")
PDF_GRAY = HexColor("#444444")

FONT_DIR = Path("/usr/share/fonts/truetype/liberation")
if (FONT_DIR / "LiberationSans-Regular.ttf").exists():
    pdfmetrics.registerFont(TTFont("Sans", str(FONT_DIR / "LiberationSans-Regular.ttf")))
    pdfmetrics.registerFont(TTFont("Sans-Bold", str(FONT_DIR / "LiberationSans-Bold.ttf")))
    pdfmetrics.registerFont(TTFont("Sans-Italic", str(FONT_DIR / "LiberationSans-Italic.ttf")))
    pdfmetrics.registerFont(TTFont("Sans-BoldItalic", str(FONT_DIR / "LiberationSans-BoldItalic.ttf")))
else:
    # Fallback names still resolve to Helvetica if registration fails later.
    pass


def set_run_font(run, name="Calibri", size=11, bold=False, italic=False, color=BLACK):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color


def set_paragraph_spacing(p, before=0, after=0, exact=None):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if exact is not None:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(exact)
    pf.widow_control = True


def add_bottom_border(paragraph, color=RULE, size="12"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_page_field(paragraph, instr, placeholder):
    def run_el():
        r = paragraph.add_run()
        set_run_font(r, size=8.5, color=GRAY)
        return r._r

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    run_el().append(begin)

    inst = OxmlElement("w:instrText")
    inst.set(qn("xml:space"), "preserve")
    inst.text = instr
    run_el().append(inst)

    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    r = run_el()
    r.append(sep)
    t = OxmlElement("w:t")
    t.text = placeholder
    r.append(t)

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run_el().append(end)


def add_hyperlink(paragraph, text, url, size=10.5, color="1B365D"):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), "Calibri")
    rFonts.set(qn("w:hAnsi"), "Calibri")
    rPr.append(rFonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rPr.append(sz)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(size * 2)))
    rPr.append(szCs)
    color_el = OxmlElement("w:color")
    color_el.set(qn("w:val"), color)
    rPr.append(color_el)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)
    new_run.append(rPr)
    text_el = OxmlElement("w:t")
    text_el.set(qn("xml:space"), "preserve")
    text_el.text = text
    new_run.append(text_el)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def new_doc(footer_label: str) -> Document:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = BLACK
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")

    bullet_style = doc.styles["List Bullet"]
    bullet_style.font.name = "Calibri"
    bullet_style.font.size = Pt(10.5)

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(fp, before=0, after=0)
    r = fp.add_run(f"{F.NAME}  ·  {footer_label}  ·  Page ")
    set_run_font(r, size=8.5, color=GRAY)
    add_page_field(fp, " PAGE ", "1")
    r = fp.add_run(" of ")
    set_run_font(r, size=8.5, color=GRAY)
    add_page_field(fp, " NUMPAGES ", "2")
    return doc


def section_heading(doc, text):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=8, after=3, exact=15)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text.upper())
    set_run_font(run, size=12, bold=True, color=NAVY)
    add_bottom_border(p)
    return p


def job_header(doc, title, dates):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=5, after=0, exact=14)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.tab_stops.add_tab_stop(Inches(7.2), WD_TAB_ALIGNMENT.RIGHT)
    r = p.add_run(title)
    set_run_font(r, size=11, bold=True, color=BLACK)
    r = p.add_run("\t" + dates)
    set_run_font(r, size=10.5, italic=True, color=GRAY)
    return p


def company_line(doc, company, location):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=0, after=1, exact=13)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(company)
    set_run_font(r, size=10.5, italic=True, color=NAVY)
    r = p.add_run(f"  |  {location}")
    set_run_font(r, size=10.5, color=GRAY)
    return p


def bullet(doc, text, keep_next=True):
    p = doc.add_paragraph(style="List Bullet")
    set_paragraph_spacing(p, before=0, after=0, exact=13)
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.15)
    p.paragraph_format.keep_with_next = keep_next
    if p.runs:
        p.runs[0].text = ""
    r = p.add_run(text)
    set_run_font(r, size=10.5, color=BLACK)
    return p


def body(doc, text, size=10.5, italic=False, before=2, after=4):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=before, after=after, exact=14)
    r = p.add_run(text)
    set_run_font(r, size=size, italic=italic, color=BLACK)
    return p


def edu_line(doc, left, right):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=2, after=0, exact=14)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.tab_stops.add_tab_stop(Inches(7.2), WD_TAB_ALIGNMENT.RIGHT)
    r = p.add_run(left)
    set_run_font(r, size=10.5, color=BLACK)
    r = p.add_run("\t" + right)
    set_run_font(r, size=10.5, italic=True, color=GRAY)
    return p


def skill_row(doc, label, items):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=1, after=1, exact=14)
    r = p.add_run(f"{label}:  ")
    set_run_font(r, size=10.5, bold=True, color=NAVY)
    r = p.add_run(items)
    set_run_font(r, size=10.5, color=BLACK)
    return p


def add_header_block(doc):
    name = doc.add_paragraph()
    name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(name, before=0, after=1, exact=28)
    r = name.add_run(F.NAME.upper())
    set_run_font(r, size=22, bold=True, color=NAVY)

    tag = doc.add_paragraph()
    tag.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(tag, before=0, after=4, exact=14)
    r = tag.add_run(F.HEADLINE)
    set_run_font(r, size=11, italic=True, color=GRAY)

    contact = doc.add_paragraph()
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(contact, before=0, after=2, exact=13)
    r = contact.add_run(f"{F.ADDRESS}  ·  {F.PHONE}")
    set_run_font(r, size=10.5, color=BLACK)

    emails = doc.add_paragraph()
    emails.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(emails, before=0, after=2, exact=13)
    add_hyperlink(emails, F.EMAIL_CISCO, f"mailto:{F.EMAIL_CISCO}")
    r = emails.add_run("  ·  ")
    set_run_font(r, size=10.5, color=BLACK)
    add_hyperlink(emails, F.EMAIL_PERSONAL, f"mailto:{F.EMAIL_PERSONAL}")
    r = emails.add_run("  ·  ")
    set_run_font(r, size=10.5, color=BLACK)
    add_hyperlink(emails, F.LINKEDIN_TEXT, F.LINKEDIN_URL)


def add_experience(doc, *, resume: bool):
    section_heading(doc, "Experience")
    company_line(doc, F.COMPANY, F.LOCATION_WORK)
    cisco_roles = F.ROLES
    for i, role in enumerate(cisco_roles):
        bullets = role["bullets"]
        if resume and role["title"].startswith("Network Consulting"):
            bullets = bullets[:1]
        if resume and role["title"].startswith("Network Engineer"):
            bullets = bullets[:1]
        job_header(doc, role["title"], role["dates"])
        if (not resume) and role.get("note"):
            body(doc, role["note"], size=10, italic=True, before=0, after=1)
        for j, text in enumerate(bullets):
            bullet(doc, text, keep_next=not (j == len(bullets) - 1 and i == len(cisco_roles) - 1))

    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=6, after=1, exact=14)
    p.paragraph_format.keep_with_next = True
    r = p.add_run("Earlier career")
    set_run_font(r, size=11, bold=True, color=BLACK)
    for j, job in enumerate(F.EARLIER):
        bullet(
            doc,
            f"{job['title']}, {job['org']}, {job['location']} ({job['dates']}) — {job['text']}",
            keep_next=j != len(F.EARLIER) - 1,
        )


def add_education_certs_skills(doc):
    section_heading(doc, "Education")
    for edu in F.EDUCATION:
        edu_line(doc, f"{edu['school']} — {edu['credential']}", edu["dates"])
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=1, after=0, exact=13)
    r = p.add_run("  ".join(edu["note"] for edu in reversed(F.EDUCATION)))
    set_run_font(r, size=10, italic=True, color=GRAY)

    section_heading(doc, "Certifications & Training")
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=3, after=1, exact=14)
    r = p.add_run(f"{F.CERTS_CURRENT_LABEL}:  ")
    set_run_font(r, size=10.5, bold=True, color=NAVY)
    r = p.add_run(F.CERTS_CURRENT)
    set_run_font(r, size=10.5, color=BLACK)

    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=1, after=1, exact=14)
    r = p.add_run("Sales enablement:  ")
    set_run_font(r, size=10.5, bold=True, color=NAVY)
    r = p.add_run(F.SALES_ENABLEMENT)
    set_run_font(r, size=10.5, color=BLACK)

    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=1, after=1, exact=14)
    r = p.add_run("Additional:  ")
    set_run_font(r, size=10.5, bold=True, color=NAVY)
    r = p.add_run(F.CERTS_ADDITIONAL)
    set_run_font(r, size=10.5, color=BLACK)

    section_heading(doc, "Skills")
    for label, items in F.SKILLS:
        skill_row(doc, label, items)


def stamp_core(doc, title: str, subject: str, comments: str):
    core = doc.core_properties
    core.author = F.NAME
    core.last_modified_by = F.NAME
    core.title = title
    core.subject = subject
    core.comments = comments


def rewrite_with_libreoffice(path: Path) -> None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.check_call(
            [soffice, "--headless", "--convert-to", "docx", "--outdir", tmp, str(path)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        produced = Path(tmp) / path.name
        if not produced.exists():
            matches = list(Path(tmp).glob("*.docx"))
            if not matches:
                raise SystemExit("LibreOffice did not produce a .docx")
            produced = matches[0]
        shutil.copy2(produced, path)


def build_resume_docx() -> Path:
    path = HERE / f"Resume - Chris Childs 2026 ({F.VERSION}).docx"
    doc = new_doc(f"Resume 2026 ({F.VERSION})")
    add_header_block(doc)
    section_heading(doc, "Professional Summary")
    body(doc, F.SUMMARY, before=3, after=1)
    add_experience(doc, resume=True)
    add_education_certs_skills(doc)
    stamp_core(
        doc,
        f"Resume — {F.NAME} 2026 ({F.VERSION})",
        "Cisco EA Architect, Premium Service Seller, and Strategic Software Sales Specialist",
        f"Word-compatible resume; Cisco certs listed as valid through December 2026; {F.LOCATION_WORK}.",
    )
    doc.save(path)
    rewrite_with_libreoffice(path)
    return path


def build_cv_docx() -> Path:
    path = HERE / f"CV - Chris Childs 2026 ({F.VERSION}).docx"
    doc = new_doc(f"CV 2026 ({F.VERSION})")
    add_header_block(doc)
    section_heading(doc, "Professional Summary")
    body(doc, F.SUMMARY, before=3, after=1)
    body(
        doc,
        "The Combined Seller (August 2023–present) and Strategic Software Sales Specialist "
        "(October 2019–present) titles are concurrent. Both are current on the 2026 resume and on LinkedIn.",
        size=10,
        italic=True,
        before=1,
        after=2,
    )
    add_experience(doc, resume=False)
    add_education_certs_skills(doc)
    stamp_core(
        doc,
        f"CV — {F.NAME} 2026 ({F.VERSION})",
        "Longer-form curriculum vitae aligned to the 2026 resume",
        "Same employment history as the resume; expanded Cisco consulting and NOC detail from the labeled sources.",
    )
    doc.save(path)
    rewrite_with_libreoffice(path)
    return path


def build_bio_docx() -> Path:
    path = HERE / f"Executive-Bio - Chris Childs 2026 ({F.VERSION}).docx"
    doc = new_doc(f"Executive bio 2026 ({F.VERSION})")
    add_header_block(doc)
    section_heading(doc, "Executive biography")
    for para in F.BIO_PARAS:
        body(doc, para, before=3, after=4)
    section_heading(doc, "Contact")
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=3, after=1, exact=14)
    r = p.add_run(f"{F.LOCATION_WORK}  ·  {F.PHONE}  ·  ")
    set_run_font(r, size=10.5, color=BLACK)
    add_hyperlink(p, F.EMAIL_CISCO, f"mailto:{F.EMAIL_CISCO}")
    r = p.add_run("  ·  ")
    set_run_font(r, size=10.5, color=BLACK)
    add_hyperlink(p, F.LINKEDIN_TEXT, F.LINKEDIN_URL)
    stamp_core(
        doc,
        f"Executive bio — {F.NAME} 2026 ({F.VERSION})",
        "Third-person biography aligned to the 2026 resume",
        "For speaker intros, internal bios, and partner introductions.",
    )
    doc.save(path)
    rewrite_with_libreoffice(path)
    return path


def build_portfolio_docx() -> Path:
    path = HERE / f"Commercial-Portfolio - Chris Childs 2026 ({F.VERSION}).docx"
    doc = new_doc(f"Commercial portfolio 2026 ({F.VERSION})")
    add_header_block(doc)
    section_heading(doc, "Purpose")
    body(
        doc,
        "Capability portfolio for Cisco software-plus-services commercial work. "
        "Employment claims match the 2026 resume. Customer decks and regional action plans "
        "in this repository are indexed as seller collateral — Cisco Confidential, indicative "
        "pricing or pipeline, not attested closed wins. Dollar amounts from those PRs are not repeated here.",
        before=3,
        after=2,
    )
    section_heading(doc, "Commercial motions")
    for motion in F.MOTIONS:
        p = doc.add_paragraph()
        set_paragraph_spacing(p, before=3, after=0, exact=14)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(motion["name"])
        set_run_font(r, size=11, bold=True, color=NAVY)
        bullet(doc, motion["text"], keep_next=True)
    section_heading(doc, "Representative work in this repository")
    body(
        doc,
        "Open pull requests on github.com/chchilds/cxvx. Treat customer-named items as Cisco Confidential.",
        size=10,
        italic=True,
        before=2,
        after=2,
    )
    for sample in F.WORK_SAMPLES:
        p = doc.add_paragraph()
        set_paragraph_spacing(p, before=2, after=0, exact=14)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(f"PR #{sample['pr']}  ·  {sample['kind']}")
        set_run_font(r, size=10.5, bold=True, color=NAVY)
        p = doc.add_paragraph()
        set_paragraph_spacing(p, before=0, after=1, exact=13)
        add_hyperlink(p, sample["title"], sample["url"])
    section_heading(doc, "How this set is used")
    bullet(doc, "Resume — two-page external CV of record.")
    bullet(doc, "CV — same facts with fuller Cisco consulting and NOC detail.")
    bullet(doc, "Executive bio — third-person introduction.")
    bullet(
        doc,
        "This portfolio — motions and a map to commercial collateral already in the repo.",
        keep_next=False,
    )
    stamp_core(
        doc,
        f"Commercial portfolio — {F.NAME} 2026 ({F.VERSION})",
        "Capability portfolio aligned to the 2026 resume",
        "Does not copy customer TCV or pipeline figures from sibling PRs.",
    )
    doc.save(path)
    rewrite_with_libreoffice(path)
    return path


def write_markdown():
    resume_md = HERE / f"Resume - Chris Childs 2026 ({F.VERSION}).md"
    cv_md = HERE / f"CV - Chris Childs 2026 ({F.VERSION}).md"
    bio_md = HERE / f"Executive-Bio - Chris Childs 2026 ({F.VERSION}).md"
    port_md = HERE / f"Commercial-Portfolio - Chris Childs 2026 ({F.VERSION}).md"
    notes = HERE / "source-notes.md"

    def header():
        return "\n".join(
            [
                f"# {F.NAME}",
                "",
                f"*{F.HEADLINE}*",
                "",
                f"{F.ADDRESS}  ·  {F.PHONE}",
                "",
                f"[{F.EMAIL_CISCO}](mailto:{F.EMAIL_CISCO})  ·  "
                f"[{F.EMAIL_PERSONAL}](mailto:{F.EMAIL_PERSONAL})  ·  "
                f"[{F.LINKEDIN_TEXT}]({F.LINKEDIN_URL})",
                "",
                "",
            ]
        )

    def experience(resume: bool) -> str:
        lines = ["## Experience", "", f"**{F.COMPANY}**  |  {F.LOCATION_WORK}", ""]
        for role in F.ROLES:
            lines.append(f"### {role['title']}")
            lines.append(f"*{role['dates']}*")
            lines.append("")
            if (not resume) and role.get("note"):
                lines.append(f"_{role['note']}_")
                lines.append("")
            bullets = role["bullets"]
            if resume and role["title"].startswith("Network Consulting"):
                bullets = bullets[:1]
            if resume and role["title"].startswith("Network Engineer"):
                bullets = bullets[:1]
            for b in bullets:
                lines.append(f"- {b}")
            lines.append("")
        lines.append("### Earlier career")
        lines.append("")
        for job in F.EARLIER:
            lines.append(
                f"- **{job['title']}**, {job['org']}, {job['location']} ({job['dates']}) — {job['text']}"
            )
        lines.append("")
        return "\n".join(lines)

    def edu_certs_skills() -> str:
        lines = ["## Education", ""]
        for edu in F.EDUCATION:
            lines.append(f"- **{edu['school']}** — {edu['credential']}  \n  *{edu['dates']}*. {edu['note']}")
        lines += [
            "",
            "## Certifications & Training",
            "",
            f"- **{F.CERTS_CURRENT_LABEL}:** {F.CERTS_CURRENT}",
            f"- **Sales enablement:** {F.SALES_ENABLEMENT}",
            f"- **Additional:** {F.CERTS_ADDITIONAL}",
            "",
            "## Skills",
            "",
        ]
        for label, items in F.SKILLS:
            lines.append(f"- **{label}:** {items}")
        lines.append("")
        return "\n".join(lines)

    resume_md.write_text(
        header()
        + f"_Resume 2026 ({F.VERSION}) · {F.AS_OF}_\n\n"
        + "## Professional Summary\n\n"
        + F.SUMMARY
        + "\n\n"
        + experience(True)
        + edu_certs_skills(),
        encoding="utf-8",
    )

    cv_md.write_text(
        header()
        + f"_Curriculum vitae 2026 ({F.VERSION}) · {F.AS_OF}_\n\n"
        + "## Professional Summary\n\n"
        + F.SUMMARY
        + "\n\n"
        + "The Combined Seller (August 2023–present) and Strategic Software Sales Specialist "
        "(October 2019–present) titles are concurrent. Both are current on the 2026 resume and on LinkedIn.\n\n"
        + experience(False)
        + edu_certs_skills(),
        encoding="utf-8",
    )

    bio_md.write_text(
        header()
        + f"_Executive biography 2026 ({F.VERSION}) · {F.AS_OF}_\n\n"
        + "\n\n".join(F.BIO_PARAS)
        + "\n",
        encoding="utf-8",
    )

    port_lines = [
        header().rstrip(),
        "",
        f"_Commercial portfolio 2026 ({F.VERSION}) · {F.AS_OF}_",
        "",
        "## Purpose",
        "",
        "Capability portfolio for Cisco software-plus-services commercial work. "
        "Employment claims match the 2026 resume. Customer decks and regional action plans "
        "in this repository are indexed as seller collateral — Cisco Confidential, indicative "
        "pricing or pipeline, not attested closed wins. Dollar amounts from those PRs are not repeated here.",
        "",
        "## Commercial motions",
        "",
    ]
    for motion in F.MOTIONS:
        port_lines.append(f"### {motion['name']}")
        port_lines.append("")
        port_lines.append(motion["text"])
        port_lines.append("")
    port_lines += [
        "## Representative work in this repository",
        "",
        "Open pull requests on [github.com/chchilds/cxvx](https://github.com/chchilds/cxvx). "
        "Treat customer-named items as Cisco Confidential.",
        "",
    ]
    for sample in F.WORK_SAMPLES:
        port_lines.append(f"- **PR #{sample['pr']}** · {sample['kind']} — [{sample['title']}]({sample['url']})")
    port_lines += [
        "",
        "## How this set is used",
        "",
        "- Resume — two-page external CV of record.",
        "- CV — same facts with fuller Cisco consulting and NOC detail.",
        "- Executive bio — third-person introduction.",
        "- This portfolio — motions and a map to commercial collateral already in the repo.",
        "",
    ]
    port_md.write_text("\n".join(port_lines), encoding="utf-8")

    notes.write_text(
        "\n".join(
            [
                f"# Source notes — {F.NAME} document set ({F.VERSION})",
                "",
                f"As of {F.AS_OF}. Do not invent employment history or claims beyond these sources.",
                "",
                "## Sources used",
                "",
                "- Resume v14 and generator on [PR #9](https://github.com/chchilds/cxvx/pull/9) (`cursor/resume-update-b3b4`). Not merged to `main`.",
                f"- Public LinkedIn: [{F.LINKEDIN_TEXT}]({F.LINKEDIN_URL}). Used only to confirm concurrent titles and to restore NCE/NOC detail that v14 had already compressed.",
                "- Open commercial PRs on github.com/chchilds/cxvx (EA decks, PS-Sub/LCS one-pager, Tallgrass action plan). Indexed as work samples; TCV and pipeline figures are not copied into the public resume set.",
                "",
                "## What was stale or missing",
                "",
                "- `main` has no resume, CV, bio, or portfolio — only the Spark Operations Bot.",
                "- Resume v14 lived only on PR #9 and was never paired with a CV, executive bio, or commercial portfolio.",
                "- Dual current titles (Combined Seller from Aug 2023 and Strategic Software Sales Specialist from Oct 2019) were both listed as Present but not labeled as concurrent.",
                "- Commercial EA / services-attach work in later PRs was not reflected as a capability on the resume.",
                f"- {F.LOCATION_NOTE}",
                "- Certifications are listed as valid through December 2026; recertification after that date is not in the sources.",
                "",
                "## Open questions for Chris",
                "",
            ]
            + [f"- {q}" for q in F.OPEN_QUESTIONS]
            + ["", ""]
        ),
        encoding="utf-8",
    )


def wrap_pdf(c, text, font, size, max_w):
    words = text.split()
    lines, cur = [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if c.stringWidth(trial, font, size) <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


class PdfDoc:
    def __init__(self, path: Path, footer: str):
        self.path = path
        self.footer = footer
        self.c = pdfcanvas.Canvas(str(path), pagesize=letter)
        self.W, self.H = letter
        self.x0 = 0.7 * inch
        self.x1 = self.W - 0.7 * inch
        self.y = self.H - 0.5 * inch
        self.page = 1
        self.max_w = self.x1 - self.x0
        self.font = "Sans"
        self.font_b = "Sans-Bold"
        self.font_i = "Sans-Italic"

    def new_page(self):
        self._footer()
        self.c.showPage()
        self.page += 1
        self.y = self.H - 0.5 * inch

    def _footer(self):
        self.c.setFillColor(PDF_GRAY)
        self.c.setFont(self.font, 8)
        self.c.drawCentredString(self.W / 2, 0.32 * inch, f"{F.NAME}  ·  {self.footer}  ·  Page {self.page}")

    def ensure(self, need=16):
        if self.y < 0.7 * inch + need:
            self.new_page()

    def heading(self, text):
        self.ensure(22)
        self.y -= 10
        self.c.setFillColor(PDF_NAVY)
        self.c.setFont(self.font_b, 12)
        self.c.drawString(self.x0, self.y, text.upper())
        self.y -= 4
        self.c.setStrokeColor(PDF_NAVY)
        self.c.setLineWidth(0.8)
        self.c.line(self.x0, self.y, self.x1, self.y)
        self.y -= 12

    def para(self, text, size=10.5, leading=13, italic=False, color=PDF_BLACK, before=2, after=4):
        font = self.font_i if italic else self.font
        lines = wrap_pdf(self.c, text, font, size, self.max_w)
        self.ensure(len(lines) * leading + before + after)
        self.y -= before
        self.c.setFillColor(color)
        self.c.setFont(font, size)
        for line in lines:
            self.c.drawString(self.x0, self.y, line)
            self.y -= leading
        self.y -= after

    def job(self, title, dates):
        # Keep the title with the first bullet (avoid orphan headers at page bottom).
        self.ensure(72)
        self.c.setFillColor(PDF_BLACK)
        self.c.setFont(self.font_b, 11)
        self.c.drawString(self.x0, self.y, title)
        self.c.setFillColor(PDF_GRAY)
        self.c.setFont(self.font_i, 10.5)
        self.c.drawRightString(self.x1, self.y, dates)
        self.y -= 13

    def company(self, company, location):
        self.ensure(16)
        self.c.setFillColor(PDF_NAVY)
        self.c.setFont(self.font_i, 10.5)
        left = f"{company}"
        self.c.drawString(self.x0, self.y, left)
        w = self.c.stringWidth(left, self.font_i, 10.5)
        self.c.setFillColor(PDF_GRAY)
        self.c.setFont(self.font, 10.5)
        self.c.drawString(self.x0 + w, self.y, f"  |  {location}")
        self.y -= 13

    def bullet(self, text):
        font = self.font
        size = 10.5
        leading = 13
        indent = 14
        lines = wrap_pdf(self.c, text, font, size, self.max_w - indent)
        self.ensure(len(lines) * leading)
        self.c.setFillColor(PDF_BLACK)
        self.c.circle(self.x0 + 4, self.y + 3, 1.4, fill=1, stroke=0)
        self.c.setFont(font, size)
        for i, line in enumerate(lines):
            self.c.drawString(self.x0 + indent, self.y, line)
            self.y -= leading

    def skill(self, label, items):
        prefix = f"{label}:  "
        self.c.setFont(self.font_b, 10.5)
        pw = self.c.stringWidth(prefix, self.font_b, 10.5)
        rest_lines = wrap_pdf(self.c, items, self.font, 10.5, self.max_w - pw)
        self.ensure(len(rest_lines) * 13)
        self.c.setFillColor(PDF_NAVY)
        self.c.setFont(self.font_b, 10.5)
        self.c.drawString(self.x0, self.y, prefix)
        self.c.setFillColor(PDF_BLACK)
        self.c.setFont(self.font, 10.5)
        if rest_lines:
            self.c.drawString(self.x0 + pw, self.y, rest_lines[0])
            self.y -= 13
            for line in rest_lines[1:]:
                self.c.drawString(self.x0 + pw, self.y, line)
                self.y -= 13
        else:
            self.y -= 13

    def header_block(self):
        self.c.setFillColor(PDF_NAVY)
        self.c.setFont(self.font_b, 20)
        self.c.drawCentredString(self.W / 2, self.y, F.NAME.upper())
        self.y -= 16
        self.c.setFillColor(PDF_GRAY)
        self.c.setFont(self.font_i, 11)
        self.c.drawCentredString(self.W / 2, self.y, F.HEADLINE)
        self.y -= 14
        self.c.setFillColor(PDF_BLACK)
        self.c.setFont(self.font, 10.5)
        self.c.drawCentredString(self.W / 2, self.y, f"{F.ADDRESS}  ·  {F.PHONE}")
        self.y -= 13
        self.c.setFillColor(PDF_NAVY)
        self.c.setFont(self.font, 10.5)
        self.c.drawCentredString(
            self.W / 2,
            self.y,
            f"{F.EMAIL_CISCO}  ·  {F.EMAIL_PERSONAL}  ·  {F.LINKEDIN_TEXT}",
        )
        self.y -= 8

    def save(self):
        self._footer()
        self.c.save()


def pdf_experience(doc: PdfDoc, resume: bool):
    doc.heading("Experience")
    doc.company(F.COMPANY, F.LOCATION_WORK)
    for role in F.ROLES:
        bullets = role["bullets"]
        if resume and role["title"].startswith("Network Consulting"):
            bullets = bullets[:1]
        if resume and role["title"].startswith("Network Engineer"):
            bullets = bullets[:1]
        doc.job(role["title"], role["dates"])
        if (not resume) and role.get("note"):
            doc.para(role["note"], size=10, italic=True, before=0, after=2)
        for text in bullets:
            doc.bullet(text)
    doc.y -= 4
    doc.ensure(72)
    doc.c.setFillColor(PDF_BLACK)
    doc.c.setFont(doc.font_b, 11)
    doc.c.drawString(doc.x0, doc.y, "Earlier career")
    doc.y -= 14
    for job in F.EARLIER:
        doc.bullet(
            f"{job['title']}, {job['org']}, {job['location']} ({job['dates']}) — {job['text']}"
        )


def pdf_edu_certs_skills(doc: PdfDoc):
    doc.heading("Education")
    for edu in F.EDUCATION:
        doc.job(f"{edu['school']} — {edu['credential']}", edu["dates"])
    doc.para(
        "  ".join(edu["note"] for edu in reversed(F.EDUCATION)),
        size=10,
        italic=True,
        color=PDF_GRAY,
        before=1,
        after=2,
    )
    doc.heading("Certifications & Training")
    doc.skill(F.CERTS_CURRENT_LABEL, F.CERTS_CURRENT)
    doc.skill("Sales enablement", F.SALES_ENABLEMENT)
    doc.skill("Additional", F.CERTS_ADDITIONAL)
    doc.heading("Skills")
    for label, items in F.SKILLS:
        doc.skill(label, items)


def build_resume_pdf() -> Path:
    path = HERE / f"Resume - Chris Childs 2026 ({F.VERSION}).pdf"
    doc = PdfDoc(path, f"Resume 2026 ({F.VERSION})")
    doc.header_block()
    doc.heading("Professional Summary")
    doc.para(F.SUMMARY, before=2, after=2)
    pdf_experience(doc, resume=True)
    pdf_edu_certs_skills(doc)
    doc.save()
    return path


def build_cv_pdf() -> Path:
    path = HERE / f"CV - Chris Childs 2026 ({F.VERSION}).pdf"
    doc = PdfDoc(path, f"CV 2026 ({F.VERSION})")
    doc.header_block()
    doc.heading("Professional Summary")
    doc.para(F.SUMMARY, before=2, after=2)
    doc.para(
        "The Combined Seller (August 2023–present) and Strategic Software Sales Specialist "
        "(October 2019–present) titles are concurrent. Both are current on the 2026 resume and on LinkedIn.",
        size=10,
        italic=True,
        before=1,
        after=4,
    )
    pdf_experience(doc, resume=False)
    pdf_edu_certs_skills(doc)
    doc.save()
    return path


def build_bio_pdf() -> Path:
    path = HERE / f"Executive-Bio - Chris Childs 2026 ({F.VERSION}).pdf"
    doc = PdfDoc(path, f"Executive bio 2026 ({F.VERSION})")
    doc.header_block()
    doc.heading("Executive biography")
    for para in F.BIO_PARAS:
        doc.para(para, before=3, after=6)
    doc.heading("Contact")
    doc.para(f"{F.LOCATION_WORK}  ·  {F.PHONE}  ·  {F.EMAIL_CISCO}  ·  {F.LINKEDIN_TEXT}")
    doc.save()
    return path


def build_portfolio_pdf() -> Path:
    path = HERE / f"Commercial-Portfolio - Chris Childs 2026 ({F.VERSION}).pdf"
    doc = PdfDoc(path, f"Commercial portfolio 2026 ({F.VERSION})")
    doc.header_block()
    doc.heading("Purpose")
    doc.para(
        "Capability portfolio for Cisco software-plus-services commercial work. "
        "Employment claims match the 2026 resume. Customer decks and regional action plans "
        "in this repository are indexed as seller collateral — Cisco Confidential, indicative "
        "pricing or pipeline, not attested closed wins. Dollar amounts from those PRs are not repeated here."
    )
    doc.heading("Commercial motions")
    for motion in F.MOTIONS:
        doc.job(motion["name"], "")
        doc.bullet(motion["text"])
    doc.heading("Representative work in this repository")
    doc.para(
        "Open pull requests on github.com/chchilds/cxvx. Treat customer-named items as Cisco Confidential.",
        size=10,
        italic=True,
    )
    for sample in F.WORK_SAMPLES:
        doc.bullet(f"PR #{sample['pr']} · {sample['kind']} — {sample['title']}")
    doc.heading("How this set is used")
    doc.bullet("Resume — two-page external CV of record.")
    doc.bullet("CV — same facts with fuller Cisco consulting and NOC detail.")
    doc.bullet("Executive bio — third-person introduction.")
    doc.bullet("This portfolio — motions and a map to commercial collateral already in the repo.")
    doc.save()
    return path


def write_readme():
    (HERE / "README.md").write_text(
        "\n".join(
            [
                f"# {F.NAME} — career documents ({F.VERSION}, {F.AS_OF})",
                "",
                "Aligned resume, CV, executive bio, and commercial portfolio. "
                "Facts come from resume v14 ([PR #9](https://github.com/chchilds/cxvx/pull/9)), "
                f"the public LinkedIn profile [{F.LINKEDIN_TEXT}]({F.LINKEDIN_URL}), "
                "and open commercial PRs in this repository. Customer TCV and pipeline figures "
                "from those PRs are not copied here.",
                "",
                "## Documents",
                "",
                f"| Use | Markdown | Word | PDF |",
                f"| --- | --- | --- | --- |",
                f"| Resume (external) | [Resume - Chris Childs 2026 ({F.VERSION}).md](Resume%20-%20Chris%20Childs%202026%20({F.VERSION}).md) | [docx](Resume%20-%20Chris%20Childs%202026%20({F.VERSION}).docx) | [pdf](Resume%20-%20Chris%20Childs%202026%20({F.VERSION}).pdf) |",
                f"| CV (longer) | [CV - Chris Childs 2026 ({F.VERSION}).md](CV%20-%20Chris%20Childs%202026%20({F.VERSION}).md) | [docx](CV%20-%20Chris%20Childs%202026%20({F.VERSION}).docx) | [pdf](CV%20-%20Chris%20Childs%202026%20({F.VERSION}).pdf) |",
                f"| Executive bio | [Executive-Bio - Chris Childs 2026 ({F.VERSION}).md](Executive-Bio%20-%20Chris%20Childs%202026%20({F.VERSION}).md) | [docx](Executive-Bio%20-%20Chris%20Childs%202026%20({F.VERSION}).docx) | [pdf](Executive-Bio%20-%20Chris%20Childs%202026%20({F.VERSION}).pdf) |",
                f"| Commercial portfolio | [Commercial-Portfolio - Chris Childs 2026 ({F.VERSION}).md](Commercial-Portfolio%20-%20Chris%20Childs%202026%20({F.VERSION}).md) | [docx](Commercial-Portfolio%20-%20Chris%20Childs%202026%20({F.VERSION}).docx) | [pdf](Commercial-Portfolio%20-%20Chris%20Childs%202026%20({F.VERSION}).pdf) |",
                "",
                "Source notes and open questions: [source-notes.md](source-notes.md).",
                "",
                "## Rebuild",
                "",
                "```bash",
                "pip install -r career/requirements.txt",
                "python3 career/build_documents.py",
                "```",
                "",
                "Edit `career/facts.py` and rebuild. Do not invent employment history.",
                "",
            ]
        ),
        encoding="utf-8",
    )


def main() -> None:
    os.chdir(HERE)
    write_markdown()
    write_readme()
    built = [
        build_resume_docx(),
        build_cv_docx(),
        build_bio_docx(),
        build_portfolio_docx(),
        build_resume_pdf(),
        build_cv_pdf(),
        build_bio_pdf(),
        build_portfolio_pdf(),
    ]
    for p in built:
        print("Wrote", p)


if __name__ == "__main__":
    main()
