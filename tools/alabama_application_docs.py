#!/usr/bin/env python3
"""Generate University of Alabama CS Assistant Professor (530199) application
DOCX files, styled to match Debasish Chakroborti academic CV (amber/green accents).

Position: Computer Science — Assistant Professor — 530199
Department of Computer Science, Lee J. Styslinger Jr. College of Engineering,
The University of Alabama, Tuscaloosa. Typical start 16 August 2027.
Preferred for fullest consideration: 30 September 2026.
Questions: coe-csdept@ua.edu

Required by the posting:
  (1) cover letter naming the department, Assistant Professor rank, and
      area(s) of specialization
  (2) resume / curriculum vitae
  (3) 3–5 page research statement
  (4) 1–2 page statement of three teaching interests and plans
  (5) names and contact information for at least four professional references

Output is written to a gitignored folder.
"""
import os
import shutil
import subprocess
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "University-of-Alabama-CS-TT-2026"
CV_SRC = ROOT / "DebasishChakrobortiCV_I"
FIG_DIR = OUT / "figures"

AMBER = RGBColor(0xB9, 0x77, 0x0E)
GREEN = RGBColor(0x1E, 0x84, 0x49)
DARK = RGBColor(0x22, 0x22, 0x22)
MUTED = RGBColor(0x44, 0x44, 0x44)

NAME = "Debasish Chakroborti (Joy), Ph.D."
EMAIL = "debasish.chakroborti@usask.ca"
PHONE = "306 914 6212"
ADDR = "410 Stensrud Rd #118, Saskatoon, SK"
WEB = "https://joydeba.github.io/"
ORCID = "https://orcid.org/0000-0002-1597-8162"
GITHUB = "https://github.com/joydeba"
SCHOLAR = "https://scholar.google.com/citations?user=VdrsyT8AAAAJ"

DATE_LINE = "September 6, 2026"
POSTING = "530199"
SEARCH_EMAIL = "coe-csdept@ua.edu"
PREFERRED = "September 30, 2026"
START = "August 16, 2027"

REFEREES = [
    {
        "name": "Professor Chanchal K. Roy",
        "rank": "Professor of Software Engineering",
        "dept": "Department of Computer Science",
        "university": "University of Saskatchewan",
        "street": "176 Thorvaldson Building, 110 Science Place",
        "city": "Saskatoon, SK  S7N 5C9, Canada",
        "email": "chanchal.roy@usask.ca",
        "phone": "+1 306 715 0600",
        "rel": (
            "Professor Roy was my Ph.D. and M.Sc. co advisor. He is also my current "
            "postdoctoral supervisor in the Software Research Lab. He can address my "
            "research, publications, mentoring, and industry collaboration."
        ),
    },
    {
        "name": "Professor Kevin A. Schneider",
        "rank": "Professor of Computer Science",
        "dept": "Department of Computer Science",
        "university": "University of Saskatchewan",
        "street": "176 Thorvaldson Building, 110 Science Place",
        "city": "Saskatoon, SK  S7N 5C9, Canada",
        "email": "kevin.schneider@usask.ca",
        "phone": "+1 306 966 4891",
        "rel": (
            "Professor Schneider was my Ph.D. and M.Sc. co advisor. He can address the "
            "quality of my research, my contributions in software engineering, and my "
            "work with students."
        ),
    },
    {
        "name": "Dr. Banani Roy",
        "rank": "Associate Professor",
        "dept": "Department of Computer Science",
        "university": "University of Saskatchewan",
        "street": "176 Thorvaldson Building, 110 Science Place",
        "city": "Saskatoon, SK  S7N 5C9, Canada",
        "email": "banani.roy@usask.ca",
        "phone": "+1 306 850 5630",
        "rel": (
            "Dr. Roy mentored my M.Sc. research on scientific workflow systems at P2IRC. "
            "She also served on my Ph.D. advisory committee. We have coauthored papers. "
            "She can address my applied AI research, collaboration, and mentoring."
        ),
    },
    {
        "name": "Joseph Herbert",
        "rank": "Academic Chair",
        "dept": "Faculty of Digital Innovation, Arts and Sciences",
        "university": "Saskatchewan Polytechnic, Regina Campus",
        "street": "4500 Wascana Parkway, Box 556",
        "city": "Regina, SK  S4P 3A3, Canada",
        "email": "herbertj@saskpolytech.ca",
        "phone": "+1 306 775 7742",
        "rel": (
            "Mr. Herbert is Academic Chair in the faculty where I teach Computer Systems "
            "Technology and Cloud Computing and Blockchain. He previously served as "
            "Instructor and Program Head of Intelligent Computing Systems from 2019 to "
            "2026. He can address my teaching, laboratory and curriculum design, and my "
            "work with colleagues and students."
        ),
    },
]


def set_run(run, *, size=10, bold=False, italic=False, color=DARK, font="Georgia"):
    run.font.name = font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color


def add_horizontal_line(paragraph, color_hex="B9770E"):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "18")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color_hex)
    pBdr.append(bottom)
    pPr.append(pBdr)


def setup_doc(margins=0.7):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(margins)
        section.bottom_margin = Inches(margins)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Georgia"
    normal.font.size = Pt(10)
    normal.font.color.rgb = DARK
    pf = normal.paragraph_format
    pf.space_after = Pt(4)
    pf.space_before = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    return doc


def header_block(doc, subtitle=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(NAME)
    set_run(r, size=16, bold=True, color=DARK, font="Calibri")

    if subtitle:
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after = Pt(2)
        r2 = p2.add_run(subtitle)
        set_run(r2, size=11, bold=True, color=AMBER, font="Calibri")

    contact = doc.add_paragraph()
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact.paragraph_format.space_after = Pt(2)
    r = contact.add_run(f"{ADDR}  ·  {PHONE}  ·  {EMAIL}")
    set_run(r, size=9, color=MUTED, font="Calibri")

    links = doc.add_paragraph()
    links.alignment = WD_ALIGN_PARAGRAPH.CENTER
    links.paragraph_format.space_after = Pt(6)
    r = links.add_run(f"{WEB}  ·  {GITHUB}  ·  {ORCID}")
    set_run(r, size=9, color=GREEN, font="Calibri")
    add_horizontal_line(links)


def meta_line(doc, text):
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = meta.add_run(text)
    set_run(r, size=9, italic=True, color=MUTED, font="Calibri")
    return meta


def section_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text.upper())
    set_run(r, size=11, bold=True, color=AMBER, font="Calibri")
    add_horizontal_line(p, "B9770E")
    return p


def sub_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    set_run(r, size=10, bold=True, color=GREEN, font="Calibri")
    return p


def body_para(doc, text, *, first_indent=False, space_after=6, size=10, bold_lead=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    if first_indent:
        p.paragraph_format.first_line_indent = Inches(0.2)
    if bold_lead:
        r0 = p.add_run(bold_lead)
        set_run(r0, size=size, bold=True, color=DARK, font="Georgia")
        r = p.add_run(text)
        set_run(r, size=size, color=DARK, font="Georgia")
    else:
        r = p.add_run(text)
        set_run(r, size=size, color=DARK, font="Georgia")
    return p


def bullet(doc, text, *, bold_prefix=None, size=10):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(0)
    if bold_prefix:
        r0 = p.add_run(bold_prefix)
        set_run(r0, size=size, bold=True, color=DARK, font="Georgia")
        r1 = p.add_run(text)
        set_run(r1, size=size, color=DARK, font="Georgia")
    else:
        r = p.add_run(text)
        set_run(r, size=size, color=DARK, font="Georgia")
    return p


def entry(doc, left, right_title, body=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.space_before = Pt(3)
    r1 = p.add_run(f"{left}  ")
    set_run(r1, size=9, bold=True, color=GREEN, font="Calibri")
    r2 = p.add_run(right_title)
    set_run(r2, size=10, bold=True, color=DARK, font="Georgia")
    if body:
        p2 = doc.add_paragraph()
        p2.paragraph_format.space_after = Pt(4)
        p2.paragraph_format.left_indent = Inches(0.15)
        r = p2.add_run(body)
        set_run(r, size=9.5, color=MUTED, font="Georgia")


def write_text(name, content):
    path = OUT / name
    path.write_text(content.strip() + "\n", encoding="utf-8")
    return path


def add_centered_figure(doc, image_path, caption, *, width=6.5):
    pic = doc.add_paragraph()
    pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic.paragraph_format.space_before = Pt(4)
    pic.paragraph_format.space_after = Pt(2)
    run = pic.add_run()
    run.add_picture(str(image_path), width=Inches(width))
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(8)
    cap.paragraph_format.space_before = Pt(0)
    r = cap.add_run(caption)
    set_run(r, size=8.5, italic=True, color=MUTED, font="Calibri")
    return pic


def _rounded(ax, x, y, w, h, facecolor, edgecolor, lw=1.15):
    box = FancyBboxPatch(
        (x, y),
        w,
        h,
        boxstyle="round,pad=0.008,rounding_size=0.018",
        linewidth=lw,
        facecolor=facecolor,
        edgecolor=edgecolor,
        mutation_aspect=0.6,
    )
    ax.add_patch(box)


def _text(ax, x, y, s, *, size=8.5, weight="regular", color="#222222", ha="center", va="center"):
    ax.text(x, y, s, fontsize=size, fontweight=weight, color=color, ha=ha, va=va, family="DejaVu Sans")


def write_research_figure():
    FIG_DIR.mkdir(parents=True, exist_ok=True)
    path = FIG_DIR / "research-overview.png"
    fig, ax = plt.subplots(figsize=(9.4, 4.55), dpi=220)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.patch.set_facecolor("white")

    _rounded(ax, 0.02, 0.86, 0.96, 0.12, "#FBF4E6", "#B9770E", lw=1.4)
    _text(ax, 0.50, 0.94, "IRE Lab  ·  Artificial Intelligence for Software Engineering", size=11, weight="bold", color="#7A4E08")
    _text(ax, 0.50, 0.885, "Goal. Issue free software on the versions people actually use", size=8.4, color="#444444")

    _text(ax, 0.50, 0.825, "How a fix reaches a stable release", size=8.2, weight="bold", color="#1E8449")

    steps = [
        (0.035, "Baczer", "Characterize\nchangesets"),
        (0.235, "ReBack", "Recommend\nwhat to backport"),
        (0.435, "BackSlice", "Keep only the\nessential edit"),
        (0.635, "BackTrans", "Adapt the patch\nwith LLMs"),
        (0.835, "BranchBench", "Evaluate agents\nvs. humans"),
    ]
    for x, title, body in steps:
        _rounded(ax, x, 0.62, 0.145, 0.18, "#F4FBF6", "#1E8449")
        _text(ax, x + 0.0725, 0.755, title, size=8.4, weight="bold", color="#146338")
        _text(ax, x + 0.0725, 0.675, body, size=7.1, color="#333333")
    for x in (0.18, 0.38, 0.58, 0.78):
        ax.add_patch(
            FancyArrowPatch(
                (x, 0.71),
                (x + 0.052, 0.71),
                arrowstyle="-|>",
                mutation_scale=10,
                color="#B9770E",
                lw=1.3,
            )
        )

    _text(ax, 0.50, 0.575, "Five year program at Alabama", size=8.2, weight="bold", color="#7A4E08")
    themes = [
        (0.035, "Theme 1", "Intelligent release\nengineering and QA", "Risk models · BranchBench\nrelease engineering assistant"),
        (0.345, "Theme 2", "Trustworthy AI assisted\nworkflows and debt", "Agent and human audits\nPeople decide. Agents propose."),
        (0.655, "Theme 3", "Industry connected\nsecure maintenance", "Cloud Linux · CVE patches\nlegacy and stable branch security"),
    ]
    for x, label, title, body in themes:
        _rounded(ax, x, 0.27, 0.31, 0.28, "#FFFCF6", "#B9770E")
        _text(ax, x + 0.155, 0.505, label, size=7.4, weight="bold", color="#B9770E")
        _text(ax, x + 0.155, 0.445, title, size=8.1, weight="bold", color="#222222")
        _text(ax, x + 0.155, 0.345, body, size=7.0, color="#444444")

    _rounded(ax, 0.02, 0.03, 0.96, 0.21, "#F7F3EA", "#B9770E")
    _text(ax, 0.50, 0.195, "At The University of Alabama", size=8.2, weight="bold", color="#7A4E08")
    _text(ax, 0.18, 0.105, "Collaborate\nProf. Carver · Prof. Gray\nCyber · ALAAI · HPC", size=7.2, color="#333333")
    _text(ax, 0.50, 0.105, "Fund\nYear 1 NSF CCF or CRII then CAREER then SaTC", size=7.2, color="#333333")
    _text(ax, 0.82, 0.105, "Train\nPh.D. · M.S. · undergraduates\nCS · AI · Cyber Security", size=7.2, color="#333333")

    fig.tight_layout(pad=0.15)
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def write_teaching_figure():
    FIG_DIR.mkdir(parents=True, exist_ok=True)
    path = FIG_DIR / "teaching-overview.png"
    fig, ax = plt.subplots(figsize=(9.4, 2.85), dpi=220)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.patch.set_facecolor("white")

    _rounded(ax, 0.02, 0.82, 0.96, 0.15, "#FBF4E6", "#B9770E", lw=1.4)
    _text(ax, 0.50, 0.935, "Three teaching interests  ·  Department of Computer Science, UA", size=10.5, weight="bold", color="#7A4E08")
    _text(ax, 0.50, 0.865, "Students design, build, test, and maintain software, including software with AI components", size=7.6, color="#444444")

    cols = [
        (
            0.03,
            "1",
            "Software design and evolution",
            "CS 200  ·  CS 415\nCS 420  ·  CS 416",
            "Ready in year 1\nLabs from backporting tools",
        ),
        (
            0.35,
            "2",
            "AI for software engineering",
            "New elective serving\nCS · AI · Cyber Security",
            "Develop in years 1 to 2\nMeasure LLM and agent patches",
        ),
        (
            0.67,
            "3",
            "Core computing and capstone",
            "CS 100 and CS 101 · CS 301\nCS 495 Capstone",
            "Ready in year 1\nWriting and staged delivery",
        ),
    ]
    for x, num, title, courses, plan in cols:
        _rounded(ax, x, 0.06, 0.30, 0.72, "#F4FBF6", "#1E8449")
        _rounded(ax, x + 0.125, 0.645, 0.05, 0.10, "#1E8449", "#1E8449")
        _text(ax, x + 0.15, 0.695, num, size=10, weight="bold", color="white")
        _text(ax, x + 0.15, 0.555, title, size=8.2, weight="bold", color="#146338")
        _text(ax, x + 0.15, 0.395, courses, size=7.6, color="#222222")
        _text(ax, x + 0.15, 0.175, plan, size=7.2, color="#444444")

    fig.tight_layout(pad=0.15)
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


# ---------------------------------------------------------------------------
# 1. COVER LETTER
# ---------------------------------------------------------------------------

def write_cover_letter():
    doc = setup_doc(0.8)
    header_block(doc, "Cover Letter. Assistant Professor of Computer Science")
    meta_line(
        doc,
        "Department of Computer Science  ·  Lee J. Styslinger Jr. College of Engineering\n"
        f"The University of Alabama  ·  Posting {POSTING}  ·  Typical start {START}",
    )

    body_para(doc, DATE_LINE, space_after=8)
    body_para(
        doc,
        "Search Committee\n"
        "Department of Computer Science\n"
        "Lee J. Styslinger Jr. College of Engineering\n"
        "The University of Alabama\n"
        "Tuscaloosa, AL  35487\n"
        f"{SEARCH_EMAIL}",
        space_after=8,
    )
    body_para(doc, "Dear Members of the Search Committee,", space_after=8)

    body_para(
        doc,
        "I am writing to apply for the tenure track Assistant Professor position in the "
        "Department of Computer Science at The University of Alabama (posting 530199). My "
        "research focuses on artificial intelligence for software engineering, with "
        "particular attention to software maintenance, automated backporting, and release "
        "engineering. I earned a Ph.D. in Computer Science from the University of "
        "Saskatchewan in 2024 and continue there as a Living Skies Postdoctoral Fellow. I "
        "also serve as an Instructor at Saskatchewan Polytechnic and as an AI Automation "
        "Consultant with Cloud Linux Software in the United States. I would welcome the "
        "opportunity to join the department in August 2027.",
        first_indent=True,
    )
    body_para(
        doc,
        "My program aligns closely with the department’s interest in artificial intelligence "
        "for software engineering. Most production software does not run on the newest "
        "development branch. It runs on a stable release that must still receive security "
        "fixes and quality repairs. Moving such a change onto a stable branch, known as "
        "backporting, requires expertise and remains error prone when done by hand. During "
        "my doctorate with Professors Chanchal K. Roy and Kevin A. Schneider, I developed "
        "Baczer, ReBack, BackSlice, and BackTrans, systems that combine repository mining, "
        "machine learning, and program analysis to study and automate backporting at scale. "
        "This work has appeared in Automated Software Engineering, the Journal of Software "
        "Evolution and Process, Empirical Software Engineering, and ACM PACMHCI, and I have "
        "presented at ICPC, ICSE, MSR, and CASCON. My current research asks whether an LLM "
        "agent maintenance change is only superficially correct, or whether it leaves "
        "technical debt, missing tests, or extra cost that a human maintainer would not "
        "have accepted.",
        first_indent=True,
    )
    body_para(
        doc,
        "I am prepared to help the department build on existing strengths while expanding "
        "its research scope. I would join the software engineering community and the "
        "growing AI faculty, collaborating with Professor Jeffrey Carver on empirical "
        "methods, research software quality, and secure maintenance, with Professor Jeff "
        "Gray on software engineering and computer science education, with cybersecurity "
        "colleagues on "
        "patch propagation and long lived systems, and with AI colleagues on language "
        "models and agent based tools. I propose to establish a laboratory in intelligent "
        "release engineering focused on risk aware maintenance of stable software, "
        "trustworthy AI assisted developer workflows, and industry connected automation. "
        "The planned campus HPC and data center would support large scale repository and "
        "model evaluation studies, and Cloud Linux, which maintains long lived Linux "
        "distributions, is an existing United States industry partner I can bring to this "
        "work.",
        first_indent=True,
    )
    body_para(
        doc,
        "I treat teaching, research, and service as complementary parts of faculty work. I "
        "am prepared to teach required undergraduate courses that support accreditation, "
        "including CS 200 Software Design and Engineering, introductory programming in CS "
        "100 and CS 101, and CS 495 Capstone Computing, and to contribute to CS 415, CS "
        "416, and CS 420 Software Evolution. I would also develop an elective in AI for "
        "software engineering to serve students in Computer Science, Artificial "
        "Intelligence, and Cyber Security. I held full responsibility for two sections of "
        "CMPT 145 at the University of Saskatchewan, taught programming and software "
        "engineering for four years as a Lecturer at Khulna University, and currently "
        "teach systems, cloud, DevOps, and project courses at Saskatchewan Polytechnic, "
        "where I have rewritten most of the laboratory materials I use. I mentored Jarin "
        "Tasnim, an M.Sc. student, to coauthored publications, supervised three "
        "undergraduate summer researchers, and completed more than 250 hours of formal "
        "instructional training.",
        first_indent=True,
    )
    body_para(
        doc,
        "I also value the collegial culture described in the advertisement. As President of "
        "the Computer Science Graduate Council I led the ICSAC Research Fest symposium, "
        "and I currently serve on the MSR 2026 Data and Challenge program committee while "
        "reviewing for JSS, JSME, ASE, FSE, ICPC, and related venues. I work effectively in "
        "laboratories, classrooms, and committees, and I would contribute to the same "
        "collegial spirit at Alabama.",
        first_indent=True,
    )
    body_para(
        doc,
        "Enclosed please find my curriculum vitae, a statement of research interests and "
        "plans, a statement of three teaching interests and plans, and the names and contact "
        "information of four professional references. I would welcome the opportunity to "
        "discuss how an externally funded laboratory in AI for software engineering can "
        "support the growth of the department and the ambitions of the College. Thank you "
        "for your consideration.",
        first_indent=True,
    )
    body_para(doc, "Sincerely,", space_after=18)
    body_para(doc, "Debasish Chakroborti (Joy), Ph.D.", space_after=2)
    body_para(doc, f"{EMAIL}  ·  {PHONE}  ·  {WEB}", space_after=2)

    path = OUT / "01_CoverLetter.docx"
    doc.save(path)
    return path


# ---------------------------------------------------------------------------
# 2. CURRICULUM VITAE (original LaTeX academic CV; UA objective only)
# ---------------------------------------------------------------------------

UA_OBJECTIVES = r"""\begin{rubric}{Career Objectives}
\entry*[]%
My research focuses on \textbf{AI assisted software maintenance}, with the goal of delivering
\textbf{issue free software on the versions people actually use}.
I pursue and continue work in legacy system monitoring, technical debt management,
continuous maintenance, cybersecurity maintenance, resource optimization, and software
sustainability, in close collaboration with industry.
I also train the next generation of computing professionals through teaching, and I aim to
expand that impact further as an Assistant Professor in the Department of Computer Science
at The University of Alabama, building an externally funded program in
\textbf{artificial intelligence for software engineering}.
\end{rubric}
"""


def write_cv():
    """Compile the original academic LaTeX CV with a UA-specific objective.

    Build in /tmp so OneDrive does not lock the PDF mid-write, then copy sources
    and the finished PDF into the application folder.
    """
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "objectives.tex").write_text(UA_OBJECTIVES, encoding="utf-8")
    shutil.copy(CV_SRC / "cv-academic.tex", OUT / "cv-academic.tex")

    build = Path("/tmp/ua-cs-tt-2026-cv")
    if build.exists():
        shutil.rmtree(build)
    build.mkdir(parents=True)
    shutil.copy(CV_SRC / "cv-academic.tex", build / "cv-academic.tex")
    (build / "objectives.tex").write_text(UA_OBJECTIVES, encoding="utf-8")
    shutil.copy(CV_SRC / "own-bib.bib", build / "own-bib.bib")

    env = os.environ.copy()
    env["TEXINPUTS"] = f"{build}:{CV_SRC}:"
    env["BIBINPUTS"] = f"{build}:{CV_SRC}:"
    env["BSTINPUTS"] = f"{CV_SRC}:"

    def run(cmd):
        result = subprocess.run(
            cmd,
            cwd=build,
            env=env,
            check=False,
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            tail = (result.stdout + "\n" + result.stderr)[-4000:]
            raise RuntimeError(f"{' '.join(cmd)} failed:\n{tail}")
        return result

    run(["pdflatex", "-interaction=nonstopmode", "-halt-on-error", "cv-academic"])
    biber = shutil.which("biber")
    thin = Path("/tmp/biber-arm64")
    if biber and not thin.exists():
        subprocess.run(
            ["lipo", biber, "-thin", "arm64", "-output", str(thin)],
            check=False,
            capture_output=True,
        )
    biber_cmd = str(thin) if thin.exists() else biber
    if biber_cmd:
        run([biber_cmd, "cv-academic"])
    run(["pdflatex", "-interaction=nonstopmode", "-halt-on-error", "cv-academic"])
    run(["pdflatex", "-interaction=nonstopmode", "-halt-on-error", "cv-academic"])

    dest = OUT / "02_CV.pdf"
    shutil.copy(build / "cv-academic.pdf", dest)
    shutil.copy(build / "cv-academic.pdf", OUT / "cv-academic.pdf")
    return dest


# ---------------------------------------------------------------------------
# 3. RESEARCH STATEMENT (3–5 pages)
# ---------------------------------------------------------------------------

def write_research_statement():
    doc = setup_doc(0.75)
    header_block(doc, "Statement of Research Interests and Plans")
    meta_line(
        doc,
        "Assistant Professor of Computer Science · Posting 530199 · The University of Alabama\n"
        "Department of Computer Science, Lee J. Styslinger Jr. College of Engineering · September 2026\n"
        "(3 to 5 pages, as requested)",
    )

    section_heading(doc, "1. Vision")
    body_para(
        doc,
        "My research lies at the intersection of artificial intelligence and software "
        "engineering. The central goal is issue free software on the versions that people "
        "actually use. Most critical software does not run on the newest development head. "
        "It runs on stable releases and older branches. That is where security patches, bug "
        "fixes, and quality problems are most difficult to manage. Moving a change onto a "
        "stable branch is known as backporting. The work remains largely manual and requires "
        "expert judgment. LLM agents now produce still more branches, pull requests, and "
        "candidate patches. That increases the maintenance burden. I develop methods and "
        "tools that help developers assess risk, migrate safe changes, manage technical debt, "
        "and evolve software with care. I evaluate these ideas on large open source projects "
        "and with industry partners.",
        first_indent=True,
        space_after=5,
    )
    body_para(
        doc,
        "This agenda aligns with the department’s stated priorities. It builds on existing "
        "strength in software engineering, empirical methods, cybersecurity, and AI, while "
        "opening a distinctive research space in intelligent release engineering and "
        "trustworthy AI assisted maintenance. The questions remain important as tools "
        "change, the resulting data and software can be released openly, and the work maps "
        "naturally onto NSF programs in software engineering and trustworthy AI. I propose "
        "to establish the Intelligent Release Engineering Lab at The University of Alabama "
        "and to recruit Ph.D., M.S., and undergraduate students from Computer Science, "
        "Artificial Intelligence, and Cyber Security.",
        first_indent=True,
        space_after=4,
    )
    add_centered_figure(
        doc,
        write_research_figure(),
        "Figure 1. Overview of the research program. The figure shows the backporting tools, "
        "three five year themes, and how the laboratory would operate at Alabama.",
        width=6.55,
    )

    section_heading(doc, "2. Research to date")
    body_para(
        doc,
        "During my Ph.D. at the University of Saskatchewan I worked with Professors "
        "Chanchal K. Roy and Kevin A. Schneider. We developed a family of tools around a "
        "single maintenance problem. Baczer characterizes backporting changes at ecosystem "
        "scale. That work appeared at ICPC in 2022 and as an ICSE Companion paper in 2024. "
        "ReBack applies neural models over pull requests, reviews, and discussion. It "
        "recommends which changes should be backported. That paper appeared in Automated "
        "Software Engineering in 2024. BackSlice uses program slicing so that only the "
        "necessary parts of a change are migrated. BackTrans uses large language models to "
        "adapt patches when branches have diverged. I also studied how open source projects "
        "manage release cutting. That paper appeared in the Journal of Software Evolution "
        "and Process in 2022. I contributed to a widely used data set of tangled bug fixing "
        "commits. That paper appeared in Empirical Software Engineering in 2021. I am "
        "completing a review of how code moves under stable version management.",
        first_indent=True,
        space_after=5,
    )
    body_para(
        doc,
        "I also mentor students as coauthors. With M.Sc. student Jarin Tasnim I studied how "
        "quality evolves in stable releases after backporting. We presented that work at "
        "ICSE in 2023. We also studied how maintainers balance a needed fix against the "
        "technical debt they incur. That paper appeared at CASCON in 2025. Her ICSAC poster "
        "on quality in backporting pull requests received a best poster award. I have "
        "supervised three undergraduate summer researchers. This mentoring model is the one "
        "I would bring to Alabama. Each project should be scoped to completion. Each student "
        "should be able to defend the data. We write early. We release a public artifact or "
        "present the work.",
        first_indent=True,
        space_after=5,
    )
    body_para(
        doc,
        "My current research examines the trustworthiness of AI based maintainers. As a "
        "Living Skies Postdoctoral Fellow I compare LLM agent backports with human "
        "backports. A change that appears correct can still introduce technical debt, omit "
        "tests, or increase energy cost. A 2026 ICSE submission compares agent fixes with "
        "human fixes side by side. I am also developing BranchBench and BranchBridge. These "
        "are patch level benchmarks for CVE backport adaptation. I am not chasing every new "
        "model. I aim for evaluation methods that remain reliable as the tools change.",
        first_indent=True,
        space_after=5,
    )
    body_para(
        doc,
        "I have also built tools that others can run. During my M.Sc. and as a research "
        "assistant at the Plant Phenotyping and Imaging Research Centre I developed RISP "
        "and RISPts. They help scientists find and reuse intermediate workflow data. That "
        "work appeared at IEEE Big Data in 2018, in ACM PACMHCI in 2021, and as a Springer "
        "chapter. I also developed InsCount, an imaging tool for phenotyping. As an AI "
        "Automation Consultant with Cloud Linux Software I apply backporting research to "
        "production Linux. I also contribute to an NSERC Alliance Advantage project with "
        "industry. These experiences matter for Alabama. The research already moves from "
        "papers to partners. I can deliver research software that others can use.",
        first_indent=True,
        space_after=5,
    )

    section_heading(doc, "3. A five year research program")
    body_para(
        doc,
        "The IRE Lab will pursue three complementary themes. Each theme includes near term "
        "work that can yield papers and open artifacts within 12 to 18 months. Each theme "
        "also includes a longer horizon that can support a CAREER scale program.",
        first_indent=True,
        space_after=5,
    )
    body_para(
        doc,
        "Theme 1 addresses intelligent release engineering and quality assurance. In the "
        "near term I will build risk models on Baczer and ReBack. At pull request time the "
        "models will estimate whether a change will require backporting and how it may "
        "affect stable releases. I will also release BranchBench as a public benchmark for "
        "patch level backport adaptation, including CVE fixes. Later I will develop a "
        "release engineering assistant. It will recommend testing and repair paths for each "
        "branch. We will evaluate it with industry partners and with campus HPC. We will "
        "share open data, benchmarks, risk models, and tools. Target venues include ICSE, "
        "FSE, ASE, EMSE, and TSE.",
        first_indent=True,
        space_after=5,
    )
    body_para(
        doc,
        "Theme 2 addresses trustworthy AI assisted developer workflows and technical debt. "
        "In the near term I will turn the agent versus human studies into methods for "
        "auditing LLM maintenance. We will examine correctness, maintainability, debt, "
        "reproducibility, and energy cost. Later I will study workflows in which agents "
        "propose changes and people decide. The guardrails will rest on evidence. "
        "Practitioners and educators will be able to use them. This theme also informs "
        "teaching. An AI for software engineering course and capstone projects can treat "
        "generated code as material to evaluate rather than as a finished product.",
        first_indent=True,
        space_after=5,
    )
    body_para(
        doc,
        "Theme 3 addresses secure maintenance with industry. In the near term I will "
        "continue the Cloud Linux work on long lived distributions. I will also seek "
        "additional United States partners who manage stable branch security daily. Later "
        "I will pursue funded projects on legacy modernization, safe patch migration, and "
        "automation. The work should produce both papers and usable tools. This theme "
        "aligns with the Cyber Security degree. It also aligns with colleagues who work on "
        "software security and research software quality.",
        first_indent=True,
        space_after=5,
    )

    section_heading(doc, "4. Why Alabama")
    body_para(
        doc,
        "I am not asking the department to invent a new area. Software engineering is "
        "already strong here. The college is also investing in AI. I would contribute to "
        "that work rather than duplicate it. Professor Jeffrey Carver is the closest "
        "faculty match. We share a commitment to measurement, people, and evidence. We "
        "both care about software that scientists and operators actually run. Professor "
        "Jeff Gray is a natural "
        "partner for mentoring students and for connecting maintenance research to the "
        "undergraduate curriculum. Cybersecurity faculty are partners for Theme 3. AI "
        "faculty, including colleagues associated with ALAAI, are partners for Theme 2. I "
        "would not bring another foundation model. I would bring careful evaluation of "
        "models used as software maintainers. Newer software engineering faculty are also "
        "welcome collaborators on shared data and courses.",
        first_indent=True,
        space_after=5,
    )
    body_para(
        doc,
        "The planned HPC and data center is a concrete resource. Repository mining, large "
        "patch studies, and repeated LLM agent experiments require data and compute. They "
        "still cost less than training a frontier model. That mix is appropriate for a new "
        "laboratory. Approximately 800 undergraduates and approximately 100 graduate "
        "students are also a strength. Undergraduates can join benchmark and reproduction "
        "projects. M.S. and Ph.D. students can lead themes. The AI and Cyber Security "
        "degrees already attract students who care about models and about safe long lived "
        "systems.",
        first_indent=True,
        space_after=5,
    )

    section_heading(doc, "5. Funding plan")
    body_para(
        doc,
        "The college seeks research growth toward a top 30 ranking by 2030. A new hire "
        "should therefore launch a visible and sustainable program promptly. My plan uses "
        "United States agency language. My Canadian NSERC experience is background only. I "
        "already understand sponsored and partnered research. NSERC is not the funding "
        "target at Alabama.",
        first_indent=True,
        space_after=5,
    )
    body_para(
        doc,
        "In year 1 I will use start up funds to recruit the first Ph.D. student and to "
        "establish shared artifacts. I will submit an NSF proposal to CISE CCF Software and "
        "Hardware Foundations on intelligent release engineering. I will continue industry "
        "support with Cloud Linux where possible. I will seek internal seed funds if they "
        "are available. If I am eligible I will also consider NSF CRII.",
        first_indent=True,
        space_after=5,
    )
    body_para(
        doc,
        "In years 2 and 3 I will prepare an NSF CAREER proposal on trustworthy AI assisted "
        "maintenance of stable software. Teaching will be part of that plan through CS 200, "
        "CS 420, and an AI for software engineering elective. If the CVE work is ready I "
        "will also consider a SaTC proposal. I will seek additional industry partners and "
        "undergraduate research placements.",
        first_indent=True,
        space_after=5,
    )
    body_para(
        doc,
        "In years 4 and 5 I will execute the CAREER project and prepare a second CISE or "
        "team proposal. The laboratory should then include two or three Ph.D. students, "
        "M.S. students, and undergraduates. Other groups should be using our open "
        "benchmarks. The compute needs remain moderate. We mine repositories, fine tune "
        "when needed, and evaluate agents. That keeps budgets sustainable and results "
        "reproducible on campus HPC.",
        first_indent=True,
        space_after=5,
    )

    section_heading(doc, "6. Students, open science, and impact")
    body_para(
        doc,
        "Students are the primary output of the program. Every thesis should produce an "
        "open artifact. That may be a data set, a benchmark, or a tool. It should also "
        "produce writing that a committee can evaluate. I already used this model with my "
        "M.Sc. mentee. We coauthored ICSE related and CASCON papers. I will supervise "
        "Ph.D. and M.S. students in Computer Science and Artificial Intelligence. I will "
        "involve undergraduates through summer research and CS 495. I will bring research "
        "modules into the courses named in my teaching statement. Tools and data will be "
        "released openly, as mine already are. Studies involving people will follow IRB "
        "practice.",
        first_indent=True,
        space_after=5,
    )
    body_para(
        doc,
        "The broader impact is concrete. Unpatched stable software is how known "
        "vulnerabilities remain in systems that people depend on. Better decisions about "
        "what to backport can reduce outages and residual risk. So can better methods for "
        "adapting a fix. So can knowing when not to trust an automated suggestion. There is "
        "a teaching impact as well. Students who learn to measure generated maintenance "
        "will become stronger engineers and stronger researchers. I wish to pursue that "
        "work in a growing department that already has software engineering and AI "
        "colleagues and that is building the computing infrastructure such a laboratory "
        "requires.",
        first_indent=True,
        space_after=5,
    )

    path = OUT / "03_ResearchStatement.docx"
    doc.save(path)
    return path


# ---------------------------------------------------------------------------
# 4. TEACHING STATEMENT (1–2 pages; three interests)
# ---------------------------------------------------------------------------

def write_teaching_statement():
    doc = setup_doc(0.8)
    header_block(doc, "Statement of Teaching Interests and Plans")
    meta_line(
        doc,
        "Assistant Professor of Computer Science · Posting 530199 · The University of Alabama\n"
        "Three teaching interests and plans · 1 to 2 pages, as requested · September 2026",
    )

    body_para(
        doc,
        "I teach so that students can design, build, test, and maintain software. That "
        "includes software with AI components. Students learn by doing the work. They use "
        "version control, code review, automated testing, and staged team projects. They "
        "also learn to evaluate tools, including large language models. I publish clear "
        "learning outcomes. I provide frequent low stakes feedback. I revise a laboratory "
        "when the help desk queue shows a bottleneck. I developed this practice as "
        "instructor of record for two sections of CMPT 145. I also taught for four years as "
        "a Lecturer at Khulna University. I currently teach systems and DevOps at "
        "Saskatchewan Polytechnic. I have completed more than 250 hours of instructional "
        "training. Below are the three teaching interests I would bring. Each maps to "
        "courses the department already offers or needs.",
        first_indent=True,
        space_after=4,
    )
    add_centered_figure(
        doc,
        write_teaching_figure(),
        "Figure 1. Three teaching interests and the University of Alabama CS courses they "
        "support.",
        width=6.55,
    )

    sub_heading(doc, "Interest 1. Software design and evolution")
    body_para(
        doc,
        "Software design does not end at first release. My first interest is the software "
        "engineering core. That includes CS 200 Software Design and Engineering, CS 415 "
        "Software Design and Development, and CS 420 Software Evolution. I can teach CS 200 "
        "immediately. I have taught software engineering and team projects as a university "
        "lecturer. I currently lead project and DevOps courses in which students configure "
        "pipelines, review one another, and deliver work in stages. CS 420 is the closest "
        "match to my research. Students would study how long lived systems change, how "
        "patches move across branches, and how adaptation fails in practice. They would "
        "use public repositories. They would use components of Baczer, ReBack, and "
        "BackSlice as laboratory artifacts rather than as a tour of my curriculum vitae. I "
        "can also contribute to CS 416 Testing and Quality Assurance. Students should leave "
        "able to specify, design, test, and reason about change. Completing an assignment "
        "is not sufficient.",
        first_indent=True,
        space_after=5,
    )

    sub_heading(doc, "Interest 2. Artificial intelligence for software engineering")
    body_para(
        doc,
        "My second interest is a new elective titled AI for Software Engineering. It can "
        "serve undergraduate and graduate students in Computer Science, Artificial "
        "Intelligence, and Cyber Security. The course would treat LLMs and agents as "
        "developer tools that must be measured. What do they repair? What debt do they "
        "introduce? When must a person govern the change? How do we test a suggested patch "
        "on a stable branch? The course would complement existing AI classes such as CS "
        "465. It would not duplicate them. I can also offer modules or guest lectures in "
        "graduate software analytics. Students may use AI assistants. They must disclose, "
        "verify, test, and take responsibility for what they submit. That is already how I "
        "conduct current development courses.",
        first_indent=True,
        space_after=5,
    )

    sub_heading(doc, "Interest 3. Core computing and capstone")
    body_para(
        doc,
        "A research active hire must also staff the core curriculum. My third interest is "
        "introductory programming, databases, and the capstone. I taught two on campus "
        "sections of Principles of Computer Science, CMPT 145. I conducted the lectures, "
        "the examinations, and supervision of three teaching assistants. I have also served "
        "as a teaching assistant across the undergraduate core, including data structures, "
        "systems programming, and databases. I am prepared to teach CS 100 and CS 101. I "
        "can contribute to CS 301. I wish to teach CS 495 Capstone Computing. It is not a "
        "residual assignment. I already supervise industry project courses and "
        "undergraduate research. I treat writing and staged delivery as part of "
        "engineering practice. That aligns with the capstone writing requirement. The "
        "department enrolls approximately 800 undergraduates. Steady core teaching is how "
        "a new assistant professor earns collegial trust.",
        first_indent=True,
        space_after=5,
    )

    body_para(
        doc,
        "Mentoring is part of teaching. I mentored Jarin Tasnim from the first research "
        "question to publication. I supervised three undergraduate summer researchers. I "
        "will recruit Ph.D., M.S., and undergraduate students into the IRE Lab from the "
        "first year. I have taught CS majors, BA students who were not majors, high school "
        "teachers, and international cohorts. That includes a selected teaching assignment "
        "in the Saskatchewan Polytechnic China Henan partnership. Students arrive with "
        "different preparation. I publish weekly targets. I live code and intentionally "
        "demonstrate errors. I revise laboratories from teaching assistant queue data. I "
        "would be glad to help teach and support accreditation of the core programs. I "
        "would also develop the AI for software engineering instruction that the search "
        "identified as a priority.",
        first_indent=True,
        space_after=4,
    )

    path = OUT / "04_TeachingStatement.docx"
    doc.save(path)
    return path


# ---------------------------------------------------------------------------
# 5. REFERENCES (at least four)
# ---------------------------------------------------------------------------

def write_references():
    doc = setup_doc(0.85)
    header_block(doc, "Professional References")
    meta_line(
        doc,
        f"Assistant Professor of Computer Science · Posting {POSTING} · The University of Alabama\n"
        "Names and complete contact information for four professional references, as required.",
    )

    for i, ref in enumerate(REFEREES, start=1):
        section_heading(doc, f"Referee {i}. {ref['name']}")
        body_para(doc, ref["rank"], space_after=1)
        body_para(doc, ref["dept"], space_after=1)
        body_para(doc, ref["university"], space_after=1)
        body_para(doc, ref["street"], space_after=1)
        body_para(doc, ref["city"], space_after=3)
        body_para(
            doc,
            f"Email {ref['email']}        Telephone {ref['phone']}",
            space_after=3,
        )
        body_para(doc, ref["rel"], space_after=8)

    path = OUT / "05_References.docx"
    doc.save(path)
    return path


# ---------------------------------------------------------------------------
# 6. LETTER-WRITER BRIEFING (internal)
# ---------------------------------------------------------------------------

def write_letter_briefing():
    doc = setup_doc(0.8)
    header_block(doc, "Letter-Writer Briefing (internal — do not submit)")
    meta_line(
        doc,
        "The University of Alabama · CS Assistant Professor · Posting 530199\n"
        f"Preferred for fullest consideration: {PREFERRED}  ·  Typical start {START}",
    )

    section_heading(doc, "What the committee requires")
    body_para(
        doc,
        "Names and contact information for at least four professional references (uploaded with the "
        "application). The portal may later request letters. Ask all four people now so they are not "
        "surprised by a short deadline later. This is an R1, research-growing College of Engineering "
        "search. Letters should speak to potential to fund and lead a lab, to teach undergraduate and "
        "graduate courses, and to be a collegial colleague.",
        first_indent=True,
    )
    bullet(doc, SEARCH_EMAIL, bold_prefix="Search email: ", size=9.5)
    bullet(doc, f"{START} (flexibility possible).", bold_prefix="Start: ", size=9.5)
    bullet(doc, PREFERRED, bold_prefix="Preferred complete application: ", size=9.5)

    section_heading(doc, "Suggested slate (four people)")
    bullet(
        doc,
        "Prof. Chanchal K. Roy — chanchal.roy@usask.ca — research independence, mentoring of "
        "Jarin Tasnim, industry-connected agenda, readiness to run a lab.",
        bold_prefix="1. Research: ",
        size=9.5,
    )
    bullet(
        doc,
        "Prof. Kevin A. Schneider — kevin.schneider@usask.ca — scholarly quality, SE contributions, "
        "graduate training, communication.",
        bold_prefix="2. Research: ",
        size=9.5,
    )
    bullet(
        doc,
        "Dr. Banani Roy — banani.roy@usask.ca — applied AI / workflow systems, interdisciplinary "
        "collaboration, mentoring. Do not use her as the only teaching letter unless she observed teaching.",
        bold_prefix="3. Research / collaboration: ",
        size=9.5,
    )
    bullet(
        doc,
        "Joseph Herbert — herbertj@saskpolytech.ca — Academic Chair, Faculty of Digital "
        "Innovation, Arts and Sciences, Saskatchewan Polytechnic (Regina). Former Program Head, "
        "Intelligent Computing Systems. Ask him to speak to teaching in CST and Cloud Computing "
        "& Blockchain, curriculum/lab rewrite, collegiality, and readiness to staff core courses.",
        bold_prefix="4. Teaching: ",
        size=9.5,
    )

    section_heading(doc, "Talking points to send each writer")
    bullet(doc, "R1 CS department; multiple hires; particular interest in AI with a focus on software engineering.", size=9.5)
    bullet(doc, "They want someone who can fund a lab (NSF CAREER trajectory) and teach accredited core courses.", size=9.5)
    bullet(doc, "Collaborators on campus: Carver (empirical SE), Gray (SE/CSEd), cyber faculty, ALAAI, campus HPC.", size=9.5)
    bullet(doc, "Teaching map: CS 200, CS 420, CS 495, new AI-for-SE elective; also CS 100/101.", size=9.5)
    bullet(doc, "Mentee Jarin Tasnim: ICSE 2023 presentation and CASCON 2025 paper.", size=9.5)
    bullet(doc, "Existing U.S. industry partner: Cloud Linux Software (automated backporting).", size=9.5)
    bullet(doc, "Collegiality is named in the ad: CSGC presidency and ICSAC organizing are useful evidence.", size=9.5)
    bullet(doc, "For Joe Herbert: CST and Cloud Computing & Blockchain teaching, lab/curriculum rewrite, ZJTIE selection, LIFT, how you work with colleagues and students at Regina campus.", size=9.5)

    section_heading(doc, "Draft request (adapt before sending)")
    body_para(
        doc,
        "I am applying for a tenure-track Assistant Professor position in Computer Science at The "
        "University of Alabama (posting 530199; AI for software engineering is a named interest; "
        "typical start 16 August 2027). Applications are preferred by 30 September 2026 for fullest "
        "consideration. The committee asks for four professional references with the application and "
        "may later request letters. I am asking you to serve as a reference and, if contacted, to "
        "comment on [research independence / scholarly quality / teaching / industry collaboration]. "
        "I will send a CV, the job ad, and a short paragraph on why this department. Thank you — I "
        "know the preferred date is close.",
        first_indent=True,
    )

    path = OUT / "06_LetterWriterBriefing_DO_NOT_SUBMIT.docx"
    doc.save(path)
    return path


# ---------------------------------------------------------------------------
# 7. CHECKLIST / README
# ---------------------------------------------------------------------------

def write_readme_docx():
    doc = setup_doc(0.8)
    header_block(doc, "Application checklist (internal)")
    meta_line(
        doc,
        f"The University of Alabama · CS Assistant Professor · {POSTING} · "
        f"Preferred {PREFERRED}",
    )

    section_heading(doc, "Submit")
    bullet(doc, "01_CoverLetter.docx (export PDF if the portal prefers PDF).", size=9.5)
    bullet(doc, "02_CV.pdf — original LaTeX academic CV; only the Career Objectives rubric is UA-specific.", size=9.5)
    bullet(doc, "03_ResearchStatement.docx — keep at 3–5 pages after your edits.", size=9.5)
    bullet(doc, "04_TeachingStatement.docx — keep at 1–2 pages; it is built around three interests.", size=9.5)
    bullet(doc, "05_References.docx — four named contacts, including Joseph Herbert (Sask Poly).", size=9.5)

    section_heading(doc, "Do not submit")
    bullet(doc, "06_LetterWriterBriefing_DO_NOT_SUBMIT.docx", size=9.5)
    bullet(doc, "This checklist, README.md, or 00_JobNotes_DO_NOT_SUBMIT.md", size=9.5)
    bullet(doc, "EDIA / statement of commitment (not requested; UA emphasizes institutional neutrality).", size=9.5)
    bullet(doc, "Haverford, Guelph, or other institutional packets. Do not reuse those files.", size=9.5)
    bullet(doc, "The Associate or Full Professor UA postings. This file is Assistant Professor only.", size=9.5)

    section_heading(doc, "Before you upload")
    body_para(
        doc,
        "Read every document aloud. Put back any sentence that does not sound like you. Confirm "
        "page counts (research 3–5; teaching 1–2). Ask all four people this week, including Joe Herbert. "
        "Apply on or before 30 September 2026.",
        first_indent=True,
    )
    body_para(
        doc,
        f"Questions: {SEARCH_EMAIL}\n"
        "Department: https://cs.ua.edu/\n"
        "College: https://eng.ua.edu/\n"
        f"Typical start: {START}",
        space_after=6,
    )

    path = OUT / "00_README_checklist.docx"
    doc.save(path)
    return path


def write_readme_md():
    text = f"""# University of Alabama — CS Assistant Professor ({POSTING})

**Do not commit this folder.** It is gitignored.

Tenure-track Assistant Professor, Department of Computer Science, Lee J. Styslinger Jr.
College of Engineering, The University of Alabama (Tuscaloosa). Particular interest:
**Artificial Intelligence with a focus on Software Engineering**.

| Item | File | Submit? |
|---|---|---|
| Cover letter | `01_CoverLetter.docx` | Yes |
| CV | `02_CV.pdf` | Yes — original LaTeX academic CV (UA objective only) |
| Research statement (3–5 pp.) | `03_ResearchStatement.docx` | Yes (includes summary figure) |
| Teaching statement (1–2 pp., three interests) | `04_TeachingStatement.docx` | Yes (includes summary figure) |
| Four professional references | `05_References.docx` | Yes |
| Letter-writer briefing | `06_LetterWriterBriefing_DO_NOT_SUBMIT.docx` | **No** |
| Job notes | `00_JobNotes_DO_NOT_SUBMIT.md` | **No** |
| This checklist | `00_README_checklist.docx` / `README.md` | **No** |

Editable Markdown drafts of the four narrative documents are also in this folder
(`01`–`05` `.md` files) if you prefer to revise in text and then re-export.

- **Preferred for fullest consideration:** {PREFERRED}
- **Posted close:** 31 March 2027, 22:55 CDT (do not wait)
- **Typical start:** {START} (flexibility possible)
- **Questions:** {SEARCH_EMAIL}
- **Apply only to this Assistant Professor posting.** Do not also apply to the Associate/Full ads.

## Required by the ad

1. Cover letter that includes the department name, the rank (Assistant Professor), and a brief
   description of specialization.
2. Resume / curriculum vitae.
3. Three-to-five-page statement of research interests and plans.
4. One-to-two-page statement of **three** teaching interests and plans.
5. Names and contact information for **at least four** professional references.

Do **not** attach an EDIA statement unless they ask.

## First actions this week

1. Email all four referees, including **Joseph Herbert** (`herbertj@saskpolytech.ca`), using the draft in `06_LetterWriterBriefing_DO_NOT_SUBMIT.docx`.
2. Read the cover letter, research statement, and teaching statement aloud and revise into your voice.
3. Export to PDF if the UA portal wants PDF.
4. Submit on or before **{PREFERRED}**.

## Regenerate

From the repository root:

```
python3 tools/alabama_application_docs.py
```

This overwrites the generated `.docx` files. It does not overwrite your hand-edits unless you
re-run the script. Keep a copy if you have already customized the Word files.
"""
    return write_text("README.md", text)


def write_job_notes():
    text = f"""# Job notes (internal — do not submit)

**Computer Science — Assistant Professor — {POSTING}**
The University of Alabama · Tuscaloosa · Department of Computer Science (214241)
Lee J. Styslinger Jr. College of Engineering
Tenure/Tenure-Track Faculty · Regular full-time

- **Preferred complete application:** {PREFERRED} (review already underway)
- **Hard close listed:** 31 March 2027, 22:55 CDT
- **Typical start:** {START}
- **Questions:** {SEARCH_EMAIL}
- **Department:** https://cs.ua.edu/
- **College:** https://eng.ua.edu/

## What they asked for

Cover letter (department + Assistant Professor + specialization), CV, 3–5 page research
statement, 1–2 page statement of **three** teaching interests and plans, **four** references.

## Why this search

Open to all CS areas; **AI with a focus on Software Engineering** is named. Growing R1
college (Top-30+ by ’30), new HPC/data center, ~800 CS undergrads, ~100 graduate students,
BS/MS in AI, BS in Cyber Security. Collegiality is unusually emphasized.

## Campus hooks used in the drafts

- Jeffrey Carver — empirical SE, research software, security, education
- Jeff Gray — SE, CS education
- Cybersecurity faculty / Cyber Security degree — secure maintenance
- ALAAI / AI faculty — LLM/agent evaluation, not competing foundation-model work
- CS 200, CS 415, CS 416, CS 420, CS 495, CS 100/101
- Cloud Linux as existing U.S. industry partner
- Campus HPC for repository-scale and agent-evaluation studies

## Do not

- Apply to the Associate/Full postings
- Attach EDIA materials unless asked
- Overclaim ICSE 2023 as a research-track paper (it is listed as poster / related presentation)
- Lead the funding plan with NSERC Discovery / Mitacs
"""
    return write_text("00_JobNotes_DO_NOT_SUBMIT.md", text)


def write_markdown_drafts():
    """Short pointers. The submitted Word files are the source of truth."""
    return [
        write_text(
            "01_CoverLetter.md",
            "See 01_CoverLetter.docx for the submitted letter.\n",
        )
    ]

def main():
    OUT.mkdir(parents=True, exist_ok=True)
    paths = [
        write_readme_docx(),
        write_readme_md(),
        write_job_notes(),
        write_cover_letter(),
        write_cv(),
        write_research_statement(),
        write_teaching_statement(),
        write_references(),
        write_letter_briefing(),
    ]
    paths.extend(write_markdown_drafts())
    for p in paths:
        print(p)


if __name__ == "__main__":
    main()
