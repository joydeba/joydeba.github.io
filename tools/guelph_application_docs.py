#!/usr/bin/env python3
"""Generate University of Guelph (Req. 2649) application DOCX files
styled to match Debasish Chakroborti academic CV (amber/green accents).
"""
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement

OUT = Path(__file__).resolve().parents[1] / "university-of-guelph"

# Match cv-academic.tex colour overrides
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


def section_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text.upper())
    set_run(r, size=11, bold=True, color=AMBER, font="Calibri")
    add_horizontal_line(p, "B9770E")
    return p


def body_para(doc, text, *, first_indent=False, space_after=6, size=10):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    if first_indent:
        p.paragraph_format.first_line_indent = Inches(0.2)
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
    """Date-left style entry matching curve CV layout (simplified)."""
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


def write_cover_letter():
    doc = setup_doc(0.85)
    header_block(doc, "Cover Letter — Assistant Professor Application")

    meta = doc.add_paragraph()
    r = meta.add_run("Requisition ID 2649  ·  School of Computer Science  ·  University of Guelph")
    set_run(r, size=9, italic=True, color=MUTED, font="Calibri")
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER

    body_para(doc, "August 2026", space_after=8)
    body_para(
        doc,
        "Dr. Minglun Gong\nDirector, School of Computer Science\nUniversity of Guelph\nfacultyjobs@socs.uoguelph.ca",
        space_after=8,
    )
    body_para(doc, "Dear Dr. Gong and Members of the Search Committee,", space_after=8)

    body_para(
        doc,
        "I am applying for the tenure-track Assistant Professor position in Software Engineering & AI / Applied AI "
        "in the School of Computer Science at the University of Guelph (Requisition ID 2649; earliest start "
        "January 1, 2027). I hold a Ph.D. in Computer Science / Software Engineering from the University of "
        "Saskatchewan and currently serve as a Living Skies Postdoctoral Fellow (NSERC CREATE SOAR / Alliance "
        "Advantage) and as an Instructor at Saskatchewan Polytechnic. My research and teaching sit squarely in "
        "software engineering and applied AI: I build intelligent methods and tools so that teams can maintain "
        "and evolve software safely on the versions people actually use.",
        first_indent=True,
    )
    body_para(
        doc,
        "I develop AI-assisted approaches for release engineering, backporting, and quality-aware software evolution. "
        "In my doctoral and postdoctoral work I created systems—Baczer, ReBack, BackSlice, and BackTrans—that combine "
        "repository mining, machine learning (including CNNs and LLMs), and program analysis, evaluated on large "
        "open-source ecosystems. This work has appeared in ICSE, ICPC, MSR, CASCON, Automated Software Engineering, "
        "and Empirical Software Engineering. As AI Automation Consultant with Cloud Linux Software, Inc., I connect "
        "the same agenda to industry maintenance practice. Earlier applied systems work at P2IRC/GIFS (RISP/RISPts) "
        "strengthens my capacity to deliver usable research software. At Guelph I will lead an independent, "
        "fundable scholarship program in intelligent release engineering, AI-assisted developer workflows and "
        "technical debt, and industry-connected applied AI—targeting NSERC Discovery, Alliance/CRD, Mitacs, and "
        "related opportunities, and aiming for internationally visible publications and open artifacts.",
        first_indent=True,
    )
    body_para(
        doc,
        "I have a strong record of undergraduate and graduate teaching and mentoring. At Saskatchewan Polytechnic "
        "I teach and redesign courses across systems, cloud, DevOps, web, and project work; at the University of "
        "Saskatchewan I was Sessional Lecturer for CMPT 145 (two sections, Spring/Summer 2025) and a Teaching "
        "Assistant for CMPT 141, 145, 214, 270, 280, 370, and Database Management Systems. As Lecturer in CSE at "
        "Khulna University I taught Structured Programming, Software Engineering and Information System, Advanced "
        "Programming Laboratory, and Software Development Project. I have mentored an M.Sc. student with co-authored "
        "publications, supervised undergraduate summer researchers, and guided teaching assistants—experience that "
        "matches the School’s preference for supervising students on scholarship-oriented projects. I am prepared "
        "to teach core computer science courses at all levels, offer graduate courses in software engineering and "
        "applied AI, supervise MSc and PhD students, and contribute to a typical three-course annual load.",
        first_indent=True,
    )
    body_para(
        doc,
        "I also bring clear evidence of community engagement and professional service: President of the Computer "
        "Science Graduate Council (CSGC); Lead Organizer of ICSAC/Research Fest; PC member (MSR 2026; SOAR 2025); "
        "reviewer for JSS, JSME/SMR, ASE, FSE, ICPC, and ICSE SEIS; outreach as Volunteer Lead for Digitized "
        "Technology (high school students); and sustained instructional development including LIFT (90 hours), "
        "Reconciliation Education, and Four Seasons of Reconciliation. These roles show I can serve the School, "
        "College, University, and broader research community.",
        first_indent=True,
    )
    body_para(\n+        doc,\n+        \"I am enclosing my curriculum vitae, Teaching Statement, Research Statement (including a funding plan), \"\n+        \"two representative papers, and contact information for three referees. I would welcome the opportunity to \"\n+        \"discuss my candidacy and how my research and teaching can strengthen the School of Computer Science. \"\n+        \"I am available for interviews and can provide additional materials or references on request. Thank you \"\n+        \"for your time and consideration.\",\n+        first_indent=True,\n+    )
    body_para(doc, "Sincerely,", space_after=18)
    body_para(doc, "Debasish Chakroborti (Joy), Ph.D.", space_after=2)
    body_para(doc, f"{EMAIL}  ·  {PHONE}", space_after=2)

    path = OUT / "2649_CoverLetter.docx"
    doc.save(path)
    return path


def write_cv():
    doc = setup_doc(0.65)
    header_block(doc, "Curriculum Vitae (Academic)")

    section_heading(doc, "Research Interests")
    body_para(
        doc,
        "Software Engineering & Applied AI — AI-assisted software maintenance and evolution; automated backporting "
        "and release engineering; technical debt and quality in stable releases; mining software repositories; "
        "LLM/ML for code analysis and recommendation; scientific workflow / data reusability; industry-connected "
        "software automation. Goal: deliver issue-free software on the versions people actually use, while training "
        "students through research-informed teaching and mentoring.",
        size=9.5,
    )

    section_heading(doc, "Education")
    entry(
        doc,
        "2020–2024",
        "Ph.D., Computer Science / Software Engineering — University of Saskatchewan, Canada",
        "Thesis: Automated Backporting for Efficient Version Management in Software Repositories. "
        "Advisors: Prof. Chanchal K. Roy and Prof. Kevin Schneider. Course mark 89%. Awards include Carl McCrosky "
        "Innovation Scholarship (2023), GSA Mark Kroeker Exceptional Student Leadership Award (2023), Ph.D. Citizenship "
        "Award (2024), Education Graduate Award (Friends Club of Saskatoon, 2024), Faculty/Departmental Scholarship, "
        "best poster awards (SOAR, SEMLA, CSGC), and Graduate Travel Awards.",
    )
    entry(
        doc,
        "2017–2019",
        "M.Sc., Computer Science / Big Data — University of Saskatchewan, Canada",
        "Thesis: An Intermediate Data-driven Methodology for Scientific Workflow Management System to Support Reusability. "
        "Advisors: Prof. Chanchal K. Roy and Prof. Kevin Schneider. Course mark 89%. Saskatchewan Innovation and "
        "Opportunity Scholarship; University of Saskatchewan Faculty Stipend.",
    )
    entry(
        doc,
        "2008–2012",
        "B.Sc., Computer Science and Engineering — Khulna University, Bangladesh",
        "Thesis: Web Service Performance Enhancement for Portable Devices Modifying SOAP Security Principle "
        "(Supervisor: Dr. Masud Rahman). CGPA 3.73/4.00 (1st class 3rd). Class Representative (Term 3.1). "
        "Innovation Award; Dean’s Merit / Merit List Scholarships (2008–2012).",
    )

    section_heading(doc, "Employment History")
    entry(
        doc,
        "2025–2026",
        "Consultant (AI Automation Team), Cloud Linux Software, Inc., USA",
        "Research advisory on AI-based automated backporting tooling.",
    )
    entry(
        doc,
        "2024–Present",
        "Instructor, School of Computing and Digital Innovation, Saskatchewan Polytechnic (Regina Campus)",
        "Computer Systems Technology and Cloud Computing & Blockchain. Selected faculty participant, China Henan "
        "(ZJTIE) Transnational Education Project (2026).",
    )
    entry(
        doc,
        "2024–Present",
        "Living Skies Postdoctoral Fellow, Department of Computer Science, University of Saskatchewan",
        "NSERC CREATE Software Analytics Research (SOAR). Automated software backporting and AI-driven development tools. "
        "Contributing research on an NSERC Alliance Advantage industry–academia project.",
    )
    entry(
        doc,
        "2025",
        "Sessional Lecturer, Department of Computer Science, University of Saskatchewan",
        "CMPT 145.3 Principles of Computer Science (Spring/Summer 2025, sections MT1 and MT2). Guided 3 TAs.",
    )
    entry(
        doc,
        "2024",
        "Software Engineer (Intern), Push Interactions",
        "Enhanced https://www.wilger.net.",
    )
    entry(
        doc,
        "2023",
        "PhD Intern (Remote), Synesis IT GmbH, Germany",
        "API planning and big data databases for Traffic Management Intelligent Transportation Systems.",
    )
    entry(
        doc,
        "2017–2024",
        "TA / RA (M.Sc. and Ph.D.), University of Saskatchewan",
        "Teaching Assistant / Student Assistantship (PSAC 40004) for CMPT 214, 145, 370, 280, 270, 141, and Database "
        "Management Systems. P2IRC/GIFS: RISP and RISPts (Hadoop/Spark). Built SE and imaging tools Baczer, ReBack (CNN), "
        "BackTrans (LLM), BackSlice, and InsCount (UNet).",
    )
    entry(
        doc,
        "2013–2017",
        "Lecturer, CSE Discipline, Khulna University, Bangladesh (2017–2024 on study leave)",
        "Teaching and research. Sites for Khulna University, ICCIT, and Convocation 2016 Registration.",
    )
    entry(
        doc,
        "2014–2016",
        "Part-time Lecturer, North Western University; Instructor (Part-time), HSTTI, Khulna",
        "Computer Science (BA program) and ICT courses for high school teachers.",
    )
    entry(
        doc,
        "2012–2013",
        "Software Engineer — Divine IT Limited; Nascenia Limited, Bangladesh",
        "ERP systems (Django/SQL); SnapKnot (Ruby on Rails / ASP.NET).",
    )

    section_heading(doc, "Selected Research Publications")
    body_para(doc, "Journal Articles", size=10)
    bullet(
        doc,
        "Chakroborti, D., Schneider, K. A., & Roy, C. K. (2024). ReBack: Recommending backports in social coding "
        "environments. Automated Software Engineering, 31(1), 18. https://doi.org/10.1007/s10515-024-00416-1",
        size=9,
    )
    bullet(
        doc,
        "Chakroborti, D., Nath, S., Schneider, K., & Roy, C. (2022). Release conventions of open-source software: "
        "An exploratory study. Journal of Software: Evolution and Process, e2499. https://doi.org/10.1002/smr.2499",
        size=9,
    )
    bullet(
        doc,
        "Chakroborti, D., Roy, B., & Nath, S. S. (2021). Designing for recommending intermediate states in a scientific "
        "workflow management system. Proc. ACM Hum.-Comput. Interact., 5(EICS). https://doi.org/10.1145/3457145",
        size=9,
    )
    bullet(
        doc,
        "Herbold, S., et al., including Chakroborti, D. (2021). A fine-grained data set and analysis of tangling in "
        "bug fixing commits. Empirical Software Engineering.",
        size=9,
    )

    body_para(doc, "Conference Proceedings (selected)", size=10)
    bullet(
        doc,
        "Tasnim, J., Chakroborti, D., Roy, C., & Schneider, K. (2025). An insight into the technical debt–fix "
        "trade-off in software backporting. CASCON 2025, Toronto, ON, Canada: IEEE.",
        size=9,
    )
    bullet(
        doc,
        "Chakroborti, D., Roy, C. K., & Schneider, K. A. (2024). A study of backporting code in open-source software "
        "for characterizing changesets. ICSE-Companion ’24. https://doi.org/10.1145/3639478.3643079",
        size=9,
    )
    bullet(
        doc,
        "Tasnim, J., Chakroborti, D., Roy, C. K., & Schneider, K. A. (2023). How does quality deviate in stable "
        "releases by backporting? ICSE 2023, Melbourne, Australia.",
        size=9,
    )
    bullet(
        doc,
        "Chakroborti, D., Schneider, K. A., & Roy, C. K. (2022). Backports: Change types, challenges and strategies. "
        "ICPC 2022. https://doi.org/10.1145/3524610.3527920",
        size=9,
    )
    bullet(
        doc,
        "Bhattacharjee, A., et al., including Chakroborti, D. (2020). An exploratory study to find motives behind "
        "cross-platform forks from Software Heritage dataset. MSR 2020. https://doi.org/10.1145/3379597.3387512",
        size=9,
    )
    bullet(
        doc,
        "Chakroborti, D., Mondal, M., Roy, B., Roy, C. K., & Schneider, K. A. (2018). Optimized storing of workflow "
        "outputs through mining association rules. IEEE Big Data 2018. https://doi.org/10.1109/BigData.2018.8622351",
        size=9,
    )
    body_para(
        doc,
        "Full list (journals, conferences, under submission, posters, book chapter): https://joydeba.github.io/research",
        size=9,
    )

    section_heading(doc, "Talks, Posters & Conference Visits")
    bullet(doc, "CASCON 2025 (Toronto) — technical debt–fix trade-off in software backporting.", size=9)
    bullet(doc, "ICSE 2023 (Melbourne) — poster on quality deviation in stable releases by backporting.", size=9)
    bullet(doc, "SEMLA 2022 (Montréal) — Best Poster (3rd): Backporting for Version Management.", size=9)
    bullet(doc, "ICPC 2022 — Backports: Change Types, Challenges and Strategies.", size=9)
    bullet(doc, "IEEE Big Data 2018 (Seattle) — Optimized Storing of Workflow Outputs…", size=9)
    bullet(doc, "SOAR 2021 — Best Poster; CSGC Research Fest / ICSAC speaker recognition and lead organizer.", size=9)

    section_heading(doc, "Research Grants & Funding")
    bullet(
        doc,
        "NSERC Alliance Advantage Grant — Postdoctoral Fellow (ongoing); Living Skies Postdoctoral Fellow contribution.",
        size=9,
    )
    bullet(doc, "Cloud Linux Software Consultancy Fund — AI Automation Consultant (AI-based automated backporting).", size=9)
    bullet(doc, "Saskatchewan Polytechnic Professional Development Fund — conference presentation support (Toronto).", size=9)
    bullet(doc, "Funding application support — assisted M.Sc. mentor with drafting a portion of a research funding application.", size=9)

    section_heading(doc, "Collaborations")
    bullet(doc, "Software Research Lab (SRLab), University of Saskatchewan — AI-assisted software maintenance / backporting.", size=9)
    bullet(doc, "Interactive Software Engineering Lab (iSE Lab), University of Saskatchewan.", size=9)
    bullet(doc, "Cloud Linux Software, Inc.; Push Interactions (industry).", size=9)
    bullet(doc, "P2IRC / GIFS — scientific data management and imaging tooling (RISP/RISPts, InsCount).", size=9)
    bullet(doc, "Water Security project (USask CFREF) — scientific computing / software analytics support.", size=9)

    section_heading(doc, "Student Supervision & Mentoring")
    bullet(doc, "M.Sc. mentoring — Jarin Tasnim (USask) on software backporting and quality analysis (co-authored publications).", size=9)
    bullet(doc, "Undergraduate summer research students — supervised 3 students.", size=9)
    bullet(doc, "Teaching assistants — guided 3 TAs for CMPT 145 (Sessional Lecturer, Spring/Summer 2025).", size=9)

    section_heading(doc, "Teaching & Curriculum")
    entry(
        doc,
        "2024–Present",
        "Instructor, Saskatchewan Polytechnic — CST and Cloud Computing & Blockchain",
        "COOS 190/294/295, COET 295, CNET 184, CWEB 280, CMPT 145; CCMP 600–606, DEVP 600, PROJ 611. "
        "Designed/rewrote course and lab materials for most courses taught. LIFT Certificate (90 hours, May 2026).",
    )
    entry(
        doc,
        "2025",
        "Sessional Lecturer, University of Saskatchewan — CMPT 145.3 (two sections)",
        "SLEQ end-of-course report available (below response threshold for closed-ended scores).",
    )
    entry(
        doc,
        "2017–2024",
        "Teaching Assistant / Student Assistantship, University of Saskatchewan",
        "CMPT 214, 145, 370, 280, 270, 141, Database Management Systems.",
    )
    entry(
        doc,
        "2013–2017",
        "Lecturer, Khulna University — CSE",
        "Structured Programming; Software Development Project; Advanced Programming Laboratory; "
        "Software Engineering and Information System.",
    )

    section_heading(doc, "Awards, Leadership & Service (selected)")
    bullet(doc, "Ph.D. Citizenship Award (2024); Education Graduate Award, Friends Club of Saskatoon (2024).", size=9)
    bullet(doc, "Carl McCrosky Innovation Scholarship (2023); GSA Mark Kroeker Exceptional Student Leadership Award (2023).", size=9)
    bullet(doc, "SOAR / SEMLA / CSGC Best Poster Awards; Saskatchewan Innovation and Opportunity Scholarship.", size=9)
    bullet(doc, "President, Computer Science Council / CSGC (2022–23); Lead Organizer, 7th ICSAC / Research Fest.", size=9)
    bullet(doc, "PC Member, MSR 2026 (Data/Challenge); PC Member, SOAR Symposium 2025.", size=9)
    bullet(doc, "Reviewer / sub-reviewer: JSS, JSME/SMR, ASE, FSE, ICPC, ICSE SEIS, CHI, SANER, ICSME, CASCON, and others.", size=9)
    bullet(doc, "Judge: Research Fest 2025; Images of Research (2024); GSA Elevator Pitch (2024); GSA Research Conference (2023).", size=9)
    bullet(doc, "Volunteer Lead, Digitized Technology (USask outreach for high school students); Folkfest and community service.", size=9)

    section_heading(doc, "Certifications & Professional Development (selected)")
    bullet(doc, "AWS Academy Educator; AWS Skill Builder badges (Cloud, Blockchain, DevOps, ECS/EKS, Backup, etc.).", size=9)
    bullet(doc, "LIFT (90 hours); New Instructor Orientation (18 hours); Reconciliation Education; 4 Seasons of Reconciliation.", size=9)
    bullet(doc, "GTI Foundation Training for University Teachers (150 hours, 2014); CETL/IQAC pedagogy modules (2014–2016).", size=9)
    bullet(doc, "TCPS 2 CORE 2022; Coursera Machine Learning / Deep Learning Specialization; ACM and IEEE membership.", size=9)

    section_heading(doc, "Skills (selected)")
    body_para(
        doc,
        "Languages: English, Bengali. Coding: Python, Java, C++, C#, Ruby, JavaScript, R, PHP, MATLAB, SQL, LaTeX. "
        "Frameworks: Django, Rails, Flask, ASP.NET, Spark, Hadoop, TensorFlow/PyTorch. Focus: software engineering, "
        "AI/ML for code, backporting, big data, teaching.",
        size=9,
    )

    section_heading(doc, "References")
    body_para(doc, "Names and complete contact information are provided in the accompanying References document.", size=9.5)

    path = OUT / "2649_CV.docx"
    doc.save(path)
    return path


def write_teaching_statement():
    doc = setup_doc(0.8)
    header_block(doc, "Teaching Statement")
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = meta.add_run(
        "Application: Assistant Professor, Software Engineering & AI / Applied AI\n"
        "School of Computer Science, University of Guelph · Requisition ID 2649 · August 2026"
    )
    set_run(r, size=9, italic=True, color=MUTED, font="Calibri")

    section_heading(doc, "1. Teaching Philosophy")
    body_para(
        doc,
        "I teach so that students can design, build, test, and reason about software systems—including systems that "
        "use AI components—with rigor and responsibility. Students learn best when theory meets authentic practice: "
        "version control, testing, code review, staged projects, and reflection on quality, security, and ethics. "
        "My classrooms are active and collaborative. I use outcomes-based design, clear rubrics, and frequent formative "
        "feedback so learners know what success looks like and how to improve.",
        first_indent=True,
    )
    body_para(
        doc,
        "Inclusive teaching is not optional. Completing Reconciliation Education and Four Seasons of Reconciliation "
        "shapes how I set transparent expectations, offer multimodal resources, and keep assessments accessible. I want "
        "every student—across backgrounds and prior preparation—to see a path into core computer science and into "
        "software engineering and applied AI.",
        first_indent=True,
    )

    section_heading(doc, "2. Experience Aligned with SoCS Needs")
    body_para(
        doc,
        "The posting asks for interest and ability in undergraduate and graduate teaching, including core CS at all "
        "levels and graduate courses in the candidate’s scholarship area. My record matches that expectation.",
        first_indent=True,
    )
    body_para(
        doc,
        "Core foundations and software engineering. As Sessional Lecturer at the University of Saskatchewan "
        "(Spring/Summer 2025), I taught CMPT 145.3 Principles of Computer Science (two sections) with full course "
        "responsibility and three TAs. As TA/Student Assistant (2017–2024) I supported CMPT 141, 145, 214, 270, 280, "
        "370, and Database Management Systems. As Lecturer in CSE at Khulna University (2013–2017) I taught Structured "
        "Programming, Software Development Project, Advanced Programming Laboratory, and Software Engineering and "
        "Information System—experience that maps directly to introductory programming, software engineering, and "
        "project courses.",
        first_indent=True,
    )
    body_para(
        doc,
        "Systems, cloud, DevOps, and applied project courses. As Instructor at Saskatchewan Polytechnic I teach and "
        "rewrite materials across Computer Systems Technology and Cloud Computing & Blockchain (e.g., COOS 190/294/295, "
        "COET 295, CNET 184, CWEB 280, CMPT 145; CCMP 600–606, DEVP 600, PROJ 611). This prepares me to contribute "
        "to a three-course annual load that mixes foundations with professionally relevant electives.",
        first_indent=True,
    )
    body_para(
        doc,
        "Courses I am ready to teach at Guelph include: introductory CS / programming; data structures; software "
        "engineering and design; software testing and quality assurance; software project / capstone; mining software "
        "repositories / software analytics; and graduate topics on engineering AI-enabled systems and AI-assisted "
        "software maintenance. I keep materials current with CI/CD, cloud tooling, and responsible use of AI-assisted "
        "developer workflows.",
        first_indent=True,
    )

    section_heading(doc, "3. Mentorship and Scholarship-Oriented Supervision")
    body_para(
        doc,
        "Preference in this search is given to candidates who have supervised students or staff on scholarship-oriented "
        "projects. I have mentored M.Sc. student Jarin Tasnim on backporting and quality analysis (co-authored "
        "publications, including CASCON 2025 and ICSE 2023-related work), supervised three undergraduate summer "
        "research students, guided teaching assistants, and supported student presentations (e.g., SOAR Symposium). "
        "Mentorship emphasizes reproducible methods, clear milestones, ethical reasoning, and publication or artifact "
        "goals. At Guelph I will supervise MSc and PhD students in SoCS programs and involve undergraduates in "
        "research projects tied to my lab’s themes.",
        first_indent=True,
    )

    section_heading(doc, "4. Evidence and Continuous Improvement")
    body_para(
        doc,
        "I invest in teaching excellence: LIFT Certificate (90 hours, May 2026) with instructional ePortfolio; New "
        "Instructor Orientation (18 hours, 2025); GTI Foundation Training for University Teachers (150 hours, 2014); "
        "and CETL/IQAC pedagogy modules. Teaching dossier: "
        "https://joydeba.github.io/teaching-materials/teaching-dossier.pdf. For CMPT 145 (USask, Spr/Sum 2025), the "
        "SLEQ report is available; closed-ended scores were not released (3/38 responses; below threshold)—qualitative "
        "comments only: https://joydeba.github.io/teaching-materials/sleq-cmpt145-2025.pdf.",
        first_indent=True,
    )
    body_para(
        doc,
        "I aim to create classrooms that are rigorous, welcoming, and practice-grounded—preparing SoCS’s large and "
        "growing undergraduate and graduate cohorts for careers and research in software engineering and AI.",
        first_indent=True,
    )

    path = OUT / "2649_TeachingStatement.docx"
    doc.save(path)
    return path


def write_research_statement():
    doc = setup_doc(0.8)
    header_block(doc, "Research Statement (Scholarship Statement)")
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = meta.add_run(
        "Application: Assistant Professor, Software Engineering & AI / Applied AI\n"
        "School of Computer Science, University of Guelph · Requisition ID 2649 · August 2026\n"
        "(Summary of research achievements · proposal for future research · funding plan)"
    )
    set_run(r, size=9, italic=True, color=MUTED, font="Calibri")

    section_heading(doc, "1. Vision and Fit")
    body_para(
        doc,
        "I work at the intersection of software engineering and applied AI. My goal is practical and ambitious: "
        "deliver issue-free software on the versions people actually use. Stable releases, long-lived branches, and "
        "security backports are where users live—and where failures are costly. I design intelligent, data-driven "
        "methods and tools that help developers anticipate risk, propagate safe changes, manage technical debt, and "
        "evolve software efficiently. This agenda matches the School’s restricted search in Software Engineering and "
        "AI, complements growth in SoCS (1,200+ undergraduates; 100+ graduate students), and is built for an "
        "independent, externally fundable faculty program.",
        first_indent=True,
    )

    section_heading(doc, "2. Research Achievements and Impact")
    body_para(
        doc,
        "A. AI-assisted software evolution at scale. In my Ph.D. at the University of Saskatchewan (advisors "
        "Prof. Chanchal K. Roy and Prof. Kevin Schneider), I developed Baczer, ReBack, BackSlice, and BackTrans to "
        "analyze, recommend, and adapt backporting changes. These systems combine repository mining, machine learning "
        "(CNNs and LLMs), and program analysis on large open-source ecosystems. Representative outcomes include "
        "ReBack in Automated Software Engineering (2024) and Backports: Change types, challenges and strategies "
        "(ICPC 2022), with related contributions at ICSE, MSR, and CASCON. The impact is both scholarly (peer-reviewed "
        "evidence and tools) and practical (automation for a high-stakes maintenance task).",
        first_indent=True,
    )
    body_para(
        doc,
        "B. Applied systems and data reusability. At P2IRC / GIFS I developed RISP and RISPts for intermediate data "
        "reusability in scientific workflow management on Hadoop/Spark (e.g., IEEE Big Data 2018; PACMHCI EICS 2021). "
        "This work shows I can build usable research software beyond a single niche.",
        first_indent=True,
    )
    body_para(
        doc,
        "C. Industry-connected applied AI and training. As Living Skies Postdoctoral Fellow (NSERC CREATE SOAR / "
        "Alliance Advantage) and AI Automation Consultant with Cloud Linux Software, Inc., I apply automated "
        "backporting research to production maintenance challenges. Mentoring an M.Sc. student and undergraduate "
        "researchers produced co-authored work on quality deviation and the technical debt–fix trade-off in "
        "backporting (ICSE 2023; CASCON 2025)—evidence that I already run scholarship-oriented student projects.",
        first_indent=True,
    )

    section_heading(doc, "3. Independent Research Program (Years 1–5)")
    body_para(
        doc,
        "Theme 1 — Intelligent release engineering and quality assurance. Forecast backport and update risk; estimate "
        "quality impact on stable releases; recommend testing and remediation pathways. Deliverables: open benchmarks, "
        "risk models, developer-facing tools, and publications in top SE venues.",
        first_indent=True,
    )
    body_para(
        doc,
        "Theme 2 — AI-assisted developer workflows and technical debt. Link issues, pull requests, commits, and reviews "
        "to risk-aware planning; study LLM-agent versus human maintenance; measure maintainability, reproducibility, "
        "and sustainability costs. Deliverables: datasets, tooling, and empirical guidelines for practice and teaching.",
        first_indent=True,
    )
    body_para(
        doc,
        "Theme 3 — Industry-connected applied AI for software maintenance. Partner on legacy modernization, secure "
        "maintenance, and automation with dual scholarly and translational impact—extending existing Cloud Linux ties "
        "and opening new Ontario/Canada partnerships aligned with SoCS and CCMPS strengths.",
        first_indent=True,
    )

    section_heading(doc, "4. Funding Plan (Tri-Council and Related)")
    body_para(
        doc,
        "The posting expects an internationally recognized program secured through NSERC and similarly prestigious "
        "sources. My staged plan:",
        first_indent=True,
    )
    bullet(
        doc,
        "Years 1–2: NSERC Discovery Grant to establish the lab (students, compute, travel); institutional start-up; "
        "Mitacs Accelerate with industry partners building on current collaborations.",
        size=9.5,
    )
    bullet(
        doc,
        "Years 2–4: NSERC Alliance / CRD proposals with industry and academic co-applicants; additional Mitacs and "
        "partner-funded placements; targeted compute/infrastructure support as needed.",
        size=9.5,
    )
    bullet(
        doc,
        "Training: supervise MSc and PhD students in SoCS programs; involve undergraduates in research projects; "
        "integrate research modules into SE and applied AI courses (supports the 40% scholarship / 40% teaching balance).",
        size=9.5,
    )
    bullet(
        doc,
        "Scale: Discovery/Alliance-class budgets typically supporting 1–2 PhD and 1–2 MSc students, infrastructure, "
        "open artifact curation, and dissemination.",
        size=9.5,
    )

    section_heading(doc, "5. Collaboration and Broader Impact")
    body_para(
        doc,
        "I will build collaborations within SoCS and across the College of Computational Mathematical and Physical "
        "Sciences, while maintaining productive links with software analytics and industry partners where appropriate. "
        "Broader impacts include open-source tools and datasets, guidance for safer stable-release maintenance, and "
        "training students in rigorous, reproducible methods—consistent with the University’s purpose To Improve Life.",
        first_indent=True,
    )

    path = OUT / "2649_ResearchStatement.docx"
    doc.save(path)
    return path


def write_references():
    doc = setup_doc(0.85)
    header_block(doc, "References")
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = meta.add_run(
        "Assistant Professor Application · Requisition ID 2649 · School of Computer Science, University of Guelph\n"
        "Please confirm permission to contact before submission; update emails/phones if needed."
    )
    set_run(r, size=9, italic=True, color=MUTED, font="Calibri")

    section_heading(doc, "Referee 1")
    body_para(doc, "Prof. Dr. Chanchal K. Roy", space_after=2)
    body_para(doc, "Professor, Department of Computer Science", space_after=2)
    body_para(doc, "University of Saskatchewan, Saskatoon, SK, Canada", space_after=2)
    body_para(doc, "Email: chanchal.roy@usask.ca", space_after=2)
    body_para(doc, "Relationship: Ph.D. and M.Sc. co-advisor", space_after=8)

    section_heading(doc, "Referee 2")
    body_para(doc, "Prof. Dr. Kevin A. Schneider", space_after=2)
    body_para(doc, "Professor, Department of Computer Science", space_after=2)
    body_para(doc, "University of Saskatchewan, Saskatoon, SK, Canada", space_after=2)
    body_para(doc, "Email: kevin.schneider@usask.ca", space_after=2)
    body_para(doc, "Relationship: Ph.D. and M.Sc. co-advisor", space_after=8)

    section_heading(doc, "Referee 3")
    body_para(doc, "[Name]", space_after=2)
    body_para(doc, "[Title / Position]", space_after=2)
    body_para(doc, "[Institution]", space_after=2)
    body_para(doc, "Email: [email]    Phone: [phone]", space_after=2)
    body_para(
        doc,
        "Relationship: [e.g., teaching supervisor / collaborator / industry mentor — please complete]",
        space_after=8,
    )

    path = OUT / "2649_References.docx"
    doc.save(path)
    return path


def write_papers_included():
    doc = setup_doc(0.85)
    header_block(doc, "Representative Papers")
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = meta.add_run(
        "Assistant Professor Application · Requisition ID 2649 · University of Guelph\n"
        "Attach the two PDF papers after this cover sheet in the integrated application PDF."
    )
    set_run(r, size=9, italic=True, color=MUTED, font="Calibri")

    section_heading(doc, "Paper 1 (recommended)")
    body_para(
        doc,
        "Chakroborti, D., Schneider, K. A., & Roy, C. K. (2024). ReBack: Recommending backports in social coding "
        "environments. Automated Software Engineering, 31(1), 18. https://doi.org/10.1007/s10515-024-00416-1",
        space_after=6,
    )
    body_para(
        doc,
        "Why include: Peer-reviewed journal evidence of applied AI for a core software engineering maintenance "
        "task (recommending backports)—directly matches Software Engineering & AI / Applied AI.",
        first_indent=True,
    )

    section_heading(doc, "Paper 2 (recommended)")
    body_para(
        doc,
        "Chakroborti, D., Schneider, K. A., & Roy, C. K. (2022). Backports: Change types, challenges and strategies. "
        "In 30th IEEE/ACM International Conference on Program Comprehension (ICPC 2022). "
        "https://doi.org/10.1145/3524610.3527920",
        space_after=6,
    )
    body_para(
        doc,
        "Why include: Foundational empirical study of backporting practice (change types, challenges, strategies) "
        "that underpins the AI-assisted maintenance research program.",
        first_indent=True,
    )

    section_heading(doc, "Alternate Paper 2 (optional swap)")
    body_para(
        doc,
        "Tasnim, J., Chakroborti, D., Roy, C., & Schneider, K. (2025). An insight into the technical debt–fix "
        "trade-off in software backporting. CASCON 2025 — highlights mentoring and technical-debt / quality themes.",
        first_indent=True,
    )

    body_para(
        doc,
        "Action for applicant: Place the two selected paper PDFs in this folder (e.g., 2649_Paper1_ReBack.pdf and "
        "2649_Paper2_ICPC2022.pdf) and merge them into the single integrated application PDF after this sheet.",
        space_after=6,
    )

    path = OUT / "2649_PapersIncluded.docx"
    doc.save(path)
    return path


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    paths = [
        write_cover_letter(),
        write_cv(),
        write_teaching_statement(),
        write_research_statement(),
        write_references(),
        write_papers_included(),
    ]
    # Remove checklist if present
    checklist = OUT / "2649_Checklist.docx"
    if checklist.exists():
        checklist.unlink()
    for p in paths:
        print(p)


if __name__ == "__main__":
    main()
