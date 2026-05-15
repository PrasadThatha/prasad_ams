import sys
import os
import re

from flask import Flask, request, send_file, jsonify
import tempfile

app = Flask(__name__)

# ─── TEXT EXTRACTION ──────────────────────────────────────────────────────────

def extract_text_from_pdf(filepath):
    from pypdf import PdfReader
    reader = PdfReader(filepath)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    return text

def extract_text_from_docx(filepath):
    from docx import Document
    doc = Document(filepath)
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

def extract_text(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".pdf":
        return extract_text_from_pdf(filepath)
    elif ext in [".docx", ".doc"]:
        return extract_text_from_docx(filepath)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

# ─── CONTENT PARSING ──────────────────────────────────────────────────────────

def parse_resume(text):
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    resume = {
        "name": "",
        "role": "",
        "mobile": "",
        "email": "",
        "sidebar_sections": [],
        "summary": [],
        "experience": []
    }

    email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', text)
    if email_match:
        resume["email"] = email_match.group()

    phone_match = re.search(r'(\+?\d[\d\s\-]{8,}\d)', text)
    if phone_match:
        resume["mobile"] = phone_match.group().strip()

    for line in lines[:5]:
        if not re.search(r'[@\d]', line) and len(line.split()) <= 5:
            resume["name"] = line
            break

    section_keywords = {
        "summary":        ["summary", "objective", "profile", "about"],
        "experience":     ["experience", "employment", "work history", "career"],
        "education":      ["education", "qualification", "academic"],
        "skills":         ["skills", "technical", "expertise", "competencies"],
        "certifications": ["certification", "certificate", "achievements"],
        "projects":       ["project"]
    }

    current_section = None
    section_data = {k: [] for k in section_keywords}
    role_candidates = []

    for i, line in enumerate(lines):
        low = line.lower()
        matched_section = None
        for sec, keywords in section_keywords.items():
            if any(kw in low for kw in keywords) and len(line.split()) <= 5:
                matched_section = sec
                break
        if matched_section:
            current_section = matched_section
            continue
        if current_section:
            section_data[current_section].append(line)
        else:
            if i > 0 and i < 6:
                role_candidates.append(line)

    title_keywords = ["analyst", "engineer", "developer", "manager", "consultant",
                      "designer", "architect", "lead", "specialist", "associate"]
    for candidate in role_candidates:
        if any(kw in candidate.lower() for kw in title_keywords):
            resume["role"] = candidate
            break
    if not resume["role"] and role_candidates:
        resume["role"] = role_candidates[-1]

    resume["summary"] = [l for l in section_data["summary"] if len(l) > 20][:6]
    if not resume["summary"]:
        resume["summary"] = ["Experienced professional with strong domain expertise."]

    edu_items = [l for l in section_data["education"] if len(l) > 5][:4]
    if edu_items:
        resume["sidebar_sections"].append({"title": "Education", "items": edu_items})

    skills_raw = [l for l in section_data["skills"] if len(l) > 2]
    skills_items = []
    for item in skills_raw:
        parts = re.split(r'[,|•]', item)
        skills_items.extend([p.strip() for p in parts if p.strip()])
    if skills_items:
        resume["sidebar_sections"].append({"title": "Technical Expertise", "items": skills_items[:12]})

    cert_items = [l for l in section_data["certifications"] if len(l) > 5][:5]
    if cert_items:
        resume["sidebar_sections"].append({"title": "Certifications", "items": cert_items})

    if not resume["sidebar_sections"]:
        resume["sidebar_sections"] = [{"title": "Skills", "items": ["Please update skills section"]}]

    exp_lines = section_data["experience"] + section_data["projects"]
    if exp_lines:
        company_name = ""
        projects = []
        current_project = None
        bullet_points = []

        date_pattern = re.compile(
            r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[\s\-]*\d{4}'
            r'|(\d{4}\s*[-–]\s*(\d{4}|Present|present))',
            re.IGNORECASE
        )

        for line in exp_lines:
            if date_pattern.search(line):
                if current_project and bullet_points:
                    current_project["points"] = bullet_points
                    projects.append(current_project)
                    bullet_points = []
                current_project = {"title": line, "duration": "", "points": []}
                date_match = date_pattern.search(line)
                if date_match:
                    current_project["duration"] = date_match.group()
                    current_project["title"] = line[:date_match.start()].strip(" -|:")
            elif line.startswith(("•", "-", "*", "·", "▪", "■")):
                bullet_points.append(re.sub(r'^[•\-\*·▪■]\s*', '', line))
            elif current_project is None and not company_name:
                company_name = line
            else:
                bullet_points.append(line)

        if current_project:
            if bullet_points:
                current_project["points"] = bullet_points
            projects.append(current_project)

        if not company_name:
            company_name = "Professional Experience"

        if projects:
            resume["experience"].append({
                "company": company_name,
                "duration": projects[0].get("duration", ""),
                "projects": projects
            })
        else:
            resume["experience"].append({
                "company": company_name,
                "duration": "",
                "projects": [{"title": "Key Responsibilities", "duration": "",
                               "points": [l for l in exp_lines if len(l) > 10][:10]}]
            })

    if not resume["experience"]:
        resume["experience"] = [{
            "company": "Experience",
            "duration": "",
            "projects": [{"title": "Role", "duration": "",
                          "points": ["Please update experience section"]}]
        }]

    return resume

# ─── PDF GENERATION ───────────────────────────────────────────────────────────

def generate_pdf(resume, output_path):
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate,
                                    Paragraph, Spacer)
    from reportlab.platypus.doctemplate import FrameBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors

    PAGE_WIDTH, PAGE_HEIGHT = A4
    HEADER_HEIGHT = 100

    doc = BaseDocTemplate(output_path, pagesize=A4,
                          leftMargin=0, rightMargin=0, topMargin=0, bottomMargin=0)

    styles = getSampleStyleSheet()

    sidebar_heading = ParagraphStyle('sidebar_heading', fontName='Helvetica-Bold',
                                     fontSize=13, leading=16, textColor=colors.white, spaceAfter=10)
    sidebar_text    = ParagraphStyle('sidebar_text', fontName='Helvetica',
                                     fontSize=9, leading=18, textColor=colors.white)
    section_heading = ParagraphStyle('section_heading', fontName='Helvetica-Bold',
                                     fontSize=15, leading=18,
                                     textColor=colors.HexColor("#333333"), spaceAfter=14)
    bullet_style    = ParagraphStyle('bullet_style', fontName='Helvetica',
                                     fontSize=9, leading=18, leftIndent=15,
                                     textColor=colors.black)
    company_style   = ParagraphStyle('company_style', fontName='Helvetica-Bold',
                                     fontSize=10, leading=14, textColor=colors.black, spaceAfter=8)

    def draw_first_page(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(colors.HexColor("#ECECEC"))
        canvas.rect(0, 0, PAGE_WIDTH, PAGE_HEIGHT, fill=1, stroke=0)
        canvas.setFillColor(colors.HexColor("#1C181B"))
        canvas.rect(0, PAGE_HEIGHT - HEADER_HEIGHT, PAGE_WIDTH, HEADER_HEIGHT, fill=1, stroke=0)
        canvas.setFillColor(colors.HexColor("#006D73"))
        canvas.rect(25, 30, 185, PAGE_HEIGHT - HEADER_HEIGHT - 50, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 18)
        canvas.drawString(25, PAGE_HEIGHT - 30, "U")
        canvas.drawString(48, PAGE_HEIGHT - 30, "S")
        canvas.drawString(25, PAGE_HEIGHT - 52, "T")
        canvas.drawString(48, PAGE_HEIGHT - 52, ".")
        canvas.setFont("Helvetica-Bold", 24)
        canvas.drawString(125, PAGE_HEIGHT - 55, resume["name"])
        canvas.setFillColor(colors.HexColor("#00B894"))
        canvas.setFont("Helvetica-Bold", 10)
        canvas.drawString(125, PAGE_HEIGHT - 73, resume["role"])
        canvas.setFont("Helvetica-Bold", 9)
        canvas.drawString(125, PAGE_HEIGHT - 88,
                          f"Mob: {resume['mobile']} | Email: {resume['email']}")
        canvas.restoreState()

    def draw_later_pages(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(colors.HexColor("#ECECEC"))
        canvas.rect(0, 0, PAGE_WIDTH, PAGE_HEIGHT, fill=1, stroke=0)
        canvas.restoreState()

    sidebar_frame    = Frame(45, 45, 145, PAGE_HEIGHT - HEADER_HEIGHT - 100,
                             showBoundary=0, leftPadding=0, rightPadding=0,
                             topPadding=0, bottomPadding=0)
    main_frame_first = Frame(230, 45, PAGE_WIDTH - 270, PAGE_HEIGHT - HEADER_HEIGHT - 60,
                             showBoundary=0, leftPadding=0, rightPadding=0,
                             topPadding=0, bottomPadding=0)
    later_frame      = Frame(55, 45, PAGE_WIDTH - 110, PAGE_HEIGHT - 70,
                             showBoundary=0, leftPadding=0, rightPadding=0,
                             topPadding=0, bottomPadding=0)

    first_template = PageTemplate(id='First',
                                  frames=[sidebar_frame, main_frame_first],
                                  onPage=draw_first_page)
    later_template = PageTemplate(id='Later',
                                  frames=[later_frame],
                                  onPage=draw_later_pages)
    doc.addPageTemplates([first_template, later_template])

    def build_sidebar():
        items = [Spacer(1, 25)]
        for section in resume["sidebar_sections"]:
            items.append(Paragraph(section["title"], sidebar_heading))
            text = ""
            for item in section["items"]:
                safe = item.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                text += f"&#9632; {safe}<br/><br/>"
            items.append(Paragraph(text, sidebar_text))
            items.append(Spacer(1, 10))
        return items

    def build_main_content():
        content = [Paragraph("Profile Summary:", section_heading)]
        for point in resume["summary"]:
            safe = point.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            content.append(Paragraph(f"&bull; {safe}", bullet_style))
        content.append(Spacer(1, 20))
        content.append(Paragraph("Professional Experience", section_heading))
        for exp in resume["experience"]:
            content.append(Paragraph(f"{exp['company']} ({exp['duration']})", company_style))
            for project in exp["projects"]:
                content.append(Paragraph(
                    f"{project['title']} ({project['duration']})", company_style))
                for point in project["points"]:
                    safe = point.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                    content.append(Paragraph(f"&bull; {safe}", bullet_style))
                content.append(Spacer(1, 12))
        return content

    story = []
    story.extend(build_sidebar())
    story.append(FrameBreak())
    story.extend(build_main_content())
    doc.build(story)

# ─── ROUTES ───────────────────────────────────────────────────────────────────

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "agent": "UST Resume Formatter"})

@app.route("/format-resume", methods=["POST"])
def format_resume():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    file = request.files["file"]
    ext = os.path.splitext(file.filename)[1].lower()

    if ext not in [".pdf", ".docx", ".doc"]:
        return jsonify({"error": "Please upload a PDF or Word file."}), 400

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path  = os.path.join(tmpdir, file.filename)
        output_path = os.path.join(tmpdir, "UST_Formatted_Resume.pdf")
        file.save(input_path)

        try:
            text   = extract_text(input_path)
            resume = parse_resume(text)
            generate_pdf(resume, output_path)
            return send_file(output_path,
                             as_attachment=True,
                             download_name="UST_Formatted_Resume.pdf",
                             mimetype="application/pdf")
        except Exception as e:
            return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
