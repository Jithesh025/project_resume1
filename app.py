from flask import Flask, render_template, request, send_file, jsonify,session, redirect
import os
from dotenv import load_dotenv
import re
from io import BytesIO
from xhtml2pdf import pisa
from google import genai
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import sqlite3
import base64
import html
from io import BytesIO
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from flask import Flask, render_template, request, send_file
from io import BytesIO

from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Image,
     TableStyle,
     Table,
    HRFlowable,
    KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT,TA_CENTER,TA_RIGHT
from reportlab.lib import colors
from reportlab.lib.units import mm

from docx import Document
from docx.shared import Pt, Inches, RGBColor
import base64
from reportlab.platypus import Image



load_dotenv()

app = Flask(__name__)
app.secret_key = "resume_builder_secret_key"



# ==============================
# Gemini AI Configuration
# ==============================

# Recommended:
API_KEY = os.environ.get("GEMINI_API_KEY")

# Or use your API key directly (not recommended for production)
# API_KEY = "YOUR_GEMINI_API_KEY"

client = genai.Client(api_key=API_KEY)

# ==============================
# Home Page
# ==============================
@app.route("/")
def home():
    return render_template("index.html")

#-------fire base sign up and login
@app.route("/login")
def login():
    return render_template("login.html")

@app.route("/register")
def register():
    return render_template("register.html")

@app.route("/main")
def main():
    return render_template("main.html")

@app.route("/feedback")
def feedback():
    return render_template('feedback.html')

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")    
 
    #------------------------------

@app.route("/editor")
def index():
    return render_template("editor.html")


@app.route("/canva-editor")
def canva_editor():
    return render_template("canva_editor.html")

@app.route("/resume1")
def resume1():
    return render_template("resume1.html")

@app.route("/resume2")
def resume2():
    return render_template("resume2.html")

@app.route("/resume3")
def resume3():
    return render_template("resume3.html")

# ==============================
# AI Resume Assistant
# ==============================


#####feedback###

#########3


@app.route("/chat", methods=["POST"])
def chat():

    data = request.get_json()

    if not data:
        return {"response": "No JSON data received."}, 400

    user_message = data.get("message", "")
    section = data.get("section", "general")

    prompt = f"""
You are an expert professional resume writer.

The user is editing the "{section}" section of a resume.

Improve the following text by:
- Correcting grammar.
- Making it ATS-friendly.
- Using professional language.
- Keeping the formatting clean.
- Do NOT invent information.
- Return ONLY the improved text.

Text:
{user_message}
"""

    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt
        )

        return {
            "response": response.text.strip()
        }

    except Exception as e:
        return {
            "response": f"AI Error: {str(e)}"
        }, 500

# ==============================
# Download Resume as PDF
# ==============================

@app.route("/download/pdf", methods=["POST"])
def download_pdf():

    data = request.get_json()

    if not data:
        return "No data received", 400

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="UTF-8">
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin:40px;
        }}

        h1 {{
            text-align:center;
            color:#0d6efd;
        }}

        h2 {{
            color:#0d6efd;
            border-bottom:1px solid #ccc;
            padding-bottom:5px;
        }}

        p {{
            line-height:1.6;
        }}

        .section {{
            margin-top:20px;
        }}
    </style>
    </head>

    <body>

    <h1>{data.get("name","")}</h1>

    <p>
        <b>Contact:</b><br>
        {data.get("contact","")}
    </p>

    <div class="section">
        <h2>Professional Summary</h2>
        <p>{data.get("summary","")}</p>
    </div>

    <div class="section">
        <h2>Education</h2>
        <p>{data.get("education","")}</p>
    </div>

    <div class="section">
        <h2>Technical Skills</h2>
        <p>{data.get("skills","")}</p>
    </div>

    <div class="section">
        <h2>Projects</h2>
        <p>{data.get("projects","")}</p>
    </div>

    <div class="section">
        <h2>Certifications</h2>
        <p>{data.get("certifications","")}</p>
    </div>

    </body>
    </html>
    """

    pdf = BytesIO()

    pisa_status = pisa.CreatePDF(html, dest=pdf)

    if pisa_status.err:
        return "Error while generating PDF.", 500

    pdf.seek(0)

    return send_file(
        pdf,
        download_name="Resume.pdf",
        as_attachment=True,
        mimetype="application/pdf"
    )
# ==============================
# Download Resume as DOCX
# ==============================
@app.route("/download/docx", methods=["POST"])
def download_docx():

    data = request.get_json()

    if not data:
        return "No data received", 400

    document = Document()

    title = document.add_heading(data.get("name",""), level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    contact = document.add_paragraph()
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact.add_run(data.get("contact",""))

    document.add_heading("Professional Summary", level=2)
    document.add_paragraph(data.get("summary",""))

    document.add_heading("Education", level=2)
    document.add_paragraph(data.get("education",""))

    document.add_heading("Technical Skills", level=2)
    document.add_paragraph(data.get("skills",""))

    document.add_heading("Projects", level=2)
    document.add_paragraph(data.get("projects",""))

    document.add_heading("Certifications", level=2)
    document.add_paragraph(data.get("certifications",""))

    file_stream = BytesIO()
    document.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        download_name="Resume.docx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
#####2nd resume pdf route#####
@app.route("/download/professional-pdf", methods=["POST"])
def download_professional_pdf():

    try:
        data = request.get_json()

        if not data:
            return "No resume data received", 400

        # ==========================================
        # CREATE PDF
        # ==========================================

        pdf_stream = BytesIO()

        doc = SimpleDocTemplate(
            pdf_stream,
            pagesize=A4,

            # Reduced page margins
            leftMargin=12 * mm,
            rightMargin=12 * mm,
            topMargin=8 * mm,
            bottomMargin=8 * mm
        )

        styles = getSampleStyleSheet()

        # ==========================================
        # STYLES
        # ==========================================

        name_style = ParagraphStyle(
            "ResumeName",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=19,
            textColor=colors.HexColor("#111111"),
            alignment=TA_CENTER,
            spaceBefore=0,
            spaceAfter=2
        )

        title_style = ParagraphStyle(
            "ResumeTitle",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8,
            leading=9,
            textColor=colors.HexColor("#555555"),
            alignment=TA_CENTER,
            spaceBefore=0,
            spaceAfter=4
        )

        section_style = ParagraphStyle(
            "ResumeSection",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=8.5,
            leading=9,
            textColor=colors.HexColor("#111111"),
            spaceBefore=0,
            spaceAfter=1,
            keepWithNext=True
        )

        body_style = ParagraphStyle(
            "ResumeBody",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=7.2,
            leading=8,
            textColor=colors.HexColor("#222222"),
            spaceBefore=0,
            spaceAfter=0
        )

        small_style = ParagraphStyle(
            "ResumeSmall",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=6.8,
            leading=7.5,
            textColor=colors.HexColor("#222222"),
            spaceBefore=0,
            spaceAfter=0
        )

        # ==========================================
        # CLEAN TEXT
        # ==========================================

        def clean_text(value):

            if value is None:
                return ""

            value = str(value)

            # Remove unnecessary spaces
            value = re.sub(r"[ \t]+", " ", value)

            # Remove multiple blank lines
            value = re.sub(r"\n\s*\n+", "\n", value)

            # Remove spaces around line breaks
            value = re.sub(r"\s*\n\s*", "\n", value)

            value = value.strip()

            # Escape HTML
            value = value.replace("&", "&amp;")
            value = value.replace("<", "&lt;")
            value = value.replace(">", "&gt;")

            # Convert newline to PDF line break
            value = value.replace("\n", "<br/>")

            return value

        # ==========================================
        # PHOTO
        # ==========================================

        photo_element = None

        photo_data = data.get("photo")

        if photo_data:

            try:

                # Remove data URL prefix
                if "," in photo_data:
                    photo_data = photo_data.split(",", 1)[1]

                image_bytes = base64.b64decode(photo_data)

                image_stream = BytesIO(image_bytes)

                photo = Image(
                    image_stream,
                    width=30 * mm,
                    height=36 * mm
                )

                photo.hAlign = "CENTER"

                photo_element = photo

            except Exception as photo_error:

                print("Photo error:", photo_error)

        # ==========================================
        # LEFT COLUMN
        # ==========================================

        left_content = []

        # Add photo ONLY if user uploaded one
        if photo_element:

            left_content.append(photo_element)
            left_content.append(Spacer(1, 1.5 * mm))

        # ------------------------------------------
        # CONTACT
        # ------------------------------------------

        left_content.append(
            Paragraph("CONTACT", section_style)
        )

        contact_parts = []

        if data.get("email"):
            contact_parts.append(clean_text(data.get("email")))

        if data.get("phone"):
            contact_parts.append(clean_text(data.get("phone")))

        if data.get("address"):
            contact_parts.append(clean_text(data.get("address")))

        if data.get("linkedin"):
            contact_parts.append(clean_text(data.get("linkedin")))

        contact_text = "<br/>".join(contact_parts)

        if contact_text:
            left_content.append(
                Paragraph(contact_text, small_style)
            )

        # ------------------------------------------
        # EDUCATION
        # ------------------------------------------

        left_content.append(
            Spacer(1, 1.5 * mm)
        )

        left_content.append(
            Paragraph("EDUCATION", section_style)
        )

        education = clean_text(
            data.get("education", "")
        )

        if education:
            left_content.append(
                Paragraph(education, small_style)
            )

        # ------------------------------------------
        # SKILLS
        # ------------------------------------------

        left_content.append(
            Spacer(1, 1.5 * mm)
        )

        left_content.append(
            Paragraph("SKILLS", section_style)
        )

        skills = clean_text(
            data.get("skills", "")
        )

        if skills:
            left_content.append(
                Paragraph(skills, small_style)
            )

        # ------------------------------------------
        # LANGUAGES
        # ------------------------------------------

        languages = clean_text(
            data.get("languages", "")
        )

        if languages:

            left_content.append(
                Spacer(1, 1.5 * mm)
            )

            left_content.append(
                Paragraph("LANGUAGES", section_style)
            )

            left_content.append(
                Paragraph(languages, small_style)
            )

        # ==========================================
        # RIGHT COLUMN
        # ==========================================

        right_content = []

        # ------------------------------------------
        # NAME
        # ------------------------------------------

        name = clean_text(
            data.get("name", "YOUR NAME")
        )

        right_content.append(
            Paragraph(name, name_style)
        )

        # ------------------------------------------
        # PROFESSIONAL TITLE
        # ------------------------------------------

        professional_title = clean_text(
            data.get(
                "professional_title",
                data.get("jobTitle", "PROFESSIONAL TITLE")
            )
        )

        if professional_title:

            right_content.append(
                Paragraph(
                    professional_title,
                    title_style
                )
            )

        # ------------------------------------------
        # PROFESSIONAL SUMMARY
        # ------------------------------------------

        right_content.append(
            Paragraph(
                "PROFESSIONAL SUMMARY",
                section_style
            )
        )

        summary = clean_text(
            data.get("summary", "")
        )

        if summary:

            right_content.append(
                Paragraph(
                    summary,
                    body_style
                )
            )

        # ------------------------------------------
        # EXPERIENCE
        # ------------------------------------------

        experience = clean_text(
            data.get("experience", "")
        )

        if experience:

            right_content.append(
                Spacer(1, 1.5 * mm)
            )

            right_content.append(
                Paragraph(
                    "EXPERIENCE",
                    section_style
                )
            )

            right_content.append(
                Paragraph(
                    experience,
                    body_style
                )
            )

        # ------------------------------------------
        # PROJECTS
        # ------------------------------------------

        projects = clean_text(
            data.get("projects", "")
        )

        if projects:

            right_content.append(
                Spacer(1, 1.5 * mm)
            )

            right_content.append(
                Paragraph(
                    "PROJECTS",
                    section_style
                )
            )

            right_content.append(
                Paragraph(
                    projects,
                    body_style
                )
            )

        # ------------------------------------------
        # CERTIFICATIONS
        # ------------------------------------------

        certifications = clean_text(
            data.get("certifications", "")
        )

        if certifications:

            right_content.append(
                Spacer(1, 1.5 * mm)
            )

            right_content.append(
                Paragraph(
                    "CERTIFICATIONS &amp; PROFESSIONAL DEVELOPMENT",
                    section_style
                )
            )

            right_content.append(
                Paragraph(
                    certifications,
                    body_style
                )
            )

        # ==========================================
        # TWO COLUMN TABLE
        # ==========================================

        table_data = [
            [
                left_content,
                right_content
            ]
        ]

        resume_table = Table(
            table_data,
            colWidths=[
                52 * mm,
                125 * mm
            ],
            hAlign="CENTER"
        )

        resume_table.setStyle(
            TableStyle([
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "TOP"
                ),

                # Reduced horizontal spacing
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    1
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    2
                ),

                # NO extra vertical padding
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    0
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    0
                ),

                # Divider between columns
                (
                    "LINEAFTER",
                    (0, 0),
                    (0, 0),
                    0.5,
                    colors.HexColor("#cccccc")
                )
            ])
        )

        # ==========================================
        # BUILD PDF
        # ==========================================

        doc.build([
            resume_table
        ])

        pdf_stream.seek(0)

        return send_file(
            pdf_stream,
            as_attachment=True,
            download_name="Professional_Resume.pdf",
            mimetype="application/pdf"
        )

    except Exception as e:

        print("PDF ERROR:", str(e))

        return f"PDF generation error: {str(e)}", 500
####docx    
@app.route("/download/professional-docx", methods=["POST"])
def professional_docx():

    data = request.get_json()

    if not data:
        return "No data received", 400

    # =========================================================
    # CREATE DOCUMENT
    # =========================================================

    document = Document()

    section = document.sections[0]

    section.top_margin = Inches(0.35)
    section.bottom_margin = Inches(0.35)
    section.left_margin = Inches(0.35)
    section.right_margin = Inches(0.35)

    # =========================================================
    # DEFAULT FONT
    # =========================================================

    styles = document.styles

    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(8)

    # =========================================================
    # CREATE TWO COLUMN TABLE
    # =========================================================

    table = document.add_table(
        rows=1,
        cols=2
    )

    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    table.autofit = False

    # Remove table borders
    tbl = table._tbl

    tblPr = tbl.tblPr

    borders = tblPr.first_child_found_in("w:tblBorders")

    if borders is None:

        borders = OxmlElement("w:tblBorders")

        tblPr.append(borders)

    for edge in (
        "top",
        "left",
        "bottom",
        "right",
        "insideH",
        "insideV"
    ):

        tag = "w:" + edge

        element = borders.find(qn(tag))

        if element is None:

            element = OxmlElement(tag)

            borders.append(element)

        element.set(
            qn("w:val"),
            "nil"
        )

    # =========================================================
    # COLUMN WIDTHS
    # =========================================================

    left = table.cell(0, 0)
    right = table.cell(0, 1)

    left.width = Inches(2.05)
    right.width = Inches(5.25)

    left.vertical_alignment = (
        WD_CELL_VERTICAL_ALIGNMENT.TOP
    )

    right.vertical_alignment = (
        WD_CELL_VERTICAL_ALIGNMENT.TOP
    )

    # =========================================================
    # CELL MARGINS
    # =========================================================

    def set_cell_margins(
        cell,
        top=80,
        start=100,
        bottom=80,
        end=100
    ):

        tc = cell._tc

        tcPr = tc.get_or_add_tcPr()

        tcMar = tcPr.first_child_found_in(
            "w:tcMar"
        )

        if tcMar is None:

            tcMar = OxmlElement("w:tcMar")

            tcPr.append(tcMar)

        for margin, value in (
            ("top", top),
            ("start", start),
            ("bottom", bottom),
            ("end", end)
        ):

            node = tcMar.find(
                qn("w:" + margin)
            )

            if node is None:

                node = OxmlElement(
                    "w:" + margin
                )

                tcMar.append(node)

            node.set(
                qn("w:w"),
                str(value)
            )

            node.set(
                qn("w:type"),
                "dxa"
            )

    set_cell_margins(
        left,
        top=100,
        start=120,
        bottom=60,
        end=120
    )

    set_cell_margins(
        right,
        top=100,
        start=180,
        bottom=60,
        end=120
    )

    # =========================================================
    # REMOVE DEFAULT PARAGRAPH
    # =========================================================

    left.paragraphs[0].text = ""
    right.paragraphs[0].text = ""

    # =========================================================
    # PARAGRAPH FORMAT FUNCTION
    # =========================================================

    def format_paragraph(
        paragraph,
        before=0,
        after=2,
        line=1.0
    ):

        paragraph.paragraph_format.space_before = Pt(
            before
        )

        paragraph.paragraph_format.space_after = Pt(
            after
        )

        paragraph.paragraph_format.line_spacing = line

    # =========================================================
    # SECTION HEADING FUNCTION
    # =========================================================

    def add_heading(
        cell,
        title,
        size=9
    ):

        p = cell.add_paragraph()

        format_paragraph(
            p,
            before=6,
            after=3
        )

        run = p.add_run(title)

        run.bold = True

        run.font.name = "Arial"

        run.font.size = Pt(size)

        run.font.color.rgb = None

        # Bottom border
        pPr = p._p.get_or_add_pPr()

        pBdr = OxmlElement("w:pBdr")

        bottom = OxmlElement("w:bottom")

        bottom.set(
            qn("w:val"),
            "single"
        )

        bottom.set(
            qn("w:sz"),
            "4"
        )

        bottom.set(
            qn("w:space"),
            "1"
        )

        bottom.set(
            qn("w:color"),
            "1C27EE"
        )

        pBdr.append(bottom)

        pPr.append(pBdr)

        return p

    # =========================================================
    # CONTENT FUNCTION
    # =========================================================

    def add_content(
        cell,
        value,
        size=8
    ):

        if not value:
            return

        value = str(value)

        lines = value.split("\n")

        for line in lines:

            line = line.strip()

            if not line:
                continue

            p = cell.add_paragraph()

            format_paragraph(
                p,
                before=0,
                after=2,
                line=1.0
            )

            run = p.add_run(line)

            run.font.name = "Arial"

            run.font.size = Pt(size)

    # =========================================================
    # PHOTO
    # =========================================================

    photo = data.get(
        "photo",
        ""
    )

    if photo.startswith("data:image"):

        try:

            image_data = photo.split(
                ",",
                1
            )[1]

            image_bytes = base64.b64decode(
                image_data
            )

            image_stream = BytesIO(
                image_bytes
            )

            p = left.add_paragraph()

            p.alignment = (
                WD_ALIGN_PARAGRAPH.CENTER
            )

            format_paragraph(
                p,
                before=0,
                after=6
            )

            run = p.add_run()

            run.add_picture(
                image_stream,
                width=Inches(1.05),
                height=Inches(1.25)
            )

        except Exception as e:

            print(
                "Photo error:",
                e
            )

    # =========================================================
    # CONTACT
    # =========================================================

    add_heading(
        left,
        "CONTACT"
    )

    contact = ""

    if data.get("email"):
        contact += data.get("email") + "\n"

    if data.get("phone"):
        contact += data.get("phone") + "\n"

    if data.get("address"):
        contact += data.get("address") + "\n"

    if data.get("linkedin"):
        contact += data.get("linkedin")

    add_content(
        left,
        contact,
        7.5
    )

    # =========================================================
    # EDUCATION
    # =========================================================

    add_heading(
        left,
        "EDUCATION"
    )

    add_content(
        left,
        data.get(
            "education",
            ""
        ),
        7.5
    )

    # =========================================================
    # SKILLS
    # =========================================================

    add_heading(
        left,
        "SKILLS"
    )

    add_content(
        left,
        data.get(
            "skills",
            ""
        ),
        7.5
    )

    # =========================================================
    # LANGUAGES
    # =========================================================

    add_heading(
        left,
        "LANGUAGES"
    )

    add_content(
        left,
        data.get(
            "languages",
            ""
        ),
        7.5
    )

    # =========================================================
    # NAME
    # =========================================================

    p = right.add_paragraph()

    p.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    format_paragraph(
        p,
        before=0,
        after=1
    )

    name_run = p.add_run(
        data.get(
            "name",
            "YOUR NAME"
        )
    )

    name_run.bold = True

    name_run.font.name = "Arial"

    name_run.font.size = Pt(20)

    # =========================================================
    # JOB TITLE
    # =========================================================

    p = right.add_paragraph()

    p.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    format_paragraph(
        p,
        before=0,
        after=0
    )

    title_run = p.add_run(
        data.get(
            "jobTitle",
            ""
        )
    )

    title_run.bold = True

    title_run.font.name = "Arial"

    title_run.font.size = Pt(8)

    # =========================================================
    # PROFESSIONAL SUMMARY
    # =========================================================

    add_heading(
        right,
        "PROFESSIONAL SUMMARY"
    )

    add_content(
        right,
        data.get(
            "summary",
            ""
        ),
        8
    )

    # =========================================================
    # EXPERIENCE
    # =========================================================

    add_heading(
        right,
        "EXPERIENCE"
    )

    add_content(
        right,
        data.get(
            "experience",
            ""
        ),
        8
    )

    # =========================================================
    # CERTIFICATIONS
    # =========================================================

    add_heading(
        right,
        "CERTIFICATIONS & PROFESSIONAL DEVELOPMENT"
    )

    add_content(
        right,
        data.get(
            "certifications",
            ""
        ),
        8
    )

    # =========================================================
    # DYNAMIC SECTIONS
    # =========================================================

    dynamic_sections = data.get(
        "dynamicSections",
        []
    )

    for item in dynamic_sections:

        section_name = item.get(
            "name",
            ""
        )

        section_content = item.get(
            "content",
            ""
        )

        if section_name:

            add_heading(
                right,
                section_name.upper()
            )

            add_content(
                right,
                section_content,
                8
            )

    # =========================================================
    # SAVE DOCX
    # =========================================================

    file_stream = BytesIO()

    document.save(
        file_stream
    )

    file_stream.seek(0)

    return send_file(

        file_stream,

        as_attachment=True,

        download_name="Professional_Resume.docx",

        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )

    )

#######resume1 download rote######
@app.route("/download/professional-resume1-pdf", methods=["POST"])
def download_resume1_pdf():

    try:
        data = request.get_json()

        if not data:
            return "No resume data received", 400

        

        # ==========================================
        # GET DATA
        # ==========================================

        name = data.get("name", "YOUR NAME")
        job_title = data.get("jobTitle", "PROFESSIONAL TITLE")

        phone = data.get("phone", "")
        address = data.get("address", "")
        email = data.get("email", "")

        summary = data.get("summary", "")
        experience = data.get("experience", "")
        education = data.get("education", "")
        skills = data.get("skills", "")
        certifications = data.get("certifications", "")

        dynamic_sections = data.get(
            "dynamicSections", []
        )

        # ==========================================
        # CLEAN HTML
        # ==========================================

        def clean_text(value):

            if value is None:
                return ""

            value = str(value)

            value = re.sub(
                r"<br\s*/?>",
                "\n",
                value,
                flags=re.IGNORECASE
            )

            value = re.sub(
                r"<[^>]+>",
                "",
                value
            )

            value = value.replace("&nbsp;", " ")
            value = value.replace("&amp;", "&")
            value = value.replace("&lt;", "<")
            value = value.replace("&gt;", ">")

            return value.strip()

        name = clean_text(name)
        job_title = clean_text(job_title)
        phone = clean_text(phone)
        address = clean_text(address)
        email = clean_text(email)

        summary = clean_text(summary)
        experience = clean_text(experience)
        education = clean_text(education)
        skills = clean_text(skills)
        certifications = clean_text(certifications)

        # ==========================================
        # PDF
        # ==========================================

        pdf_buffer = BytesIO()

        doc = SimpleDocTemplate(
            pdf_buffer,
            pagesize=A4,
            leftMargin=15 * mm,
            rightMargin=15 * mm,
            topMargin=12 * mm,
            bottomMargin=12 * mm
        )

        # ==========================================
        # STYLES
        # ==========================================

        name_style = ParagraphStyle(
            "Name",
            fontName="Helvetica-Bold",
            fontSize=25,
            leading=28,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#1a1a2e"),
            spaceAfter=3
        )

        title_style = ParagraphStyle(
            "Title",
            fontName="Helvetica",
            fontSize=10,
            leading=12,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#4a4a4a"),
            spaceAfter=8
        )

        contact_style = ParagraphStyle(
            "Contact",
            fontName="Helvetica",
            fontSize=7.5,
            leading=9,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#333333")
        )

        heading_style = ParagraphStyle(
            "Heading",
            fontName="Helvetica-Bold",
            fontSize=10.5,
            leading=12,
            textColor=colors.HexColor("#1a1a2e"),
            spaceBefore=5,
            spaceAfter=4
        )

        body_style = ParagraphStyle(
            "Body",
            fontName="Helvetica",
            fontSize=8.5,
            leading=10.5,
            textColor=colors.HexColor("#4a4a4a"),
            spaceBefore=0,
            spaceAfter=2
        )

        # ==========================================
        # CONTENT HELPER
        # ==========================================

        def add_content(story, content):

            content = clean_text(content)

            if not content:
                return

            for line in content.split("\n"):

                line = line.strip()

                if not line:
                    continue

                # Remove HTML bullet entity if present
                line = line.replace("&bull;", "•")

                story.append(
                    Paragraph(
                        line.replace("&", "&amp;")
                            .replace("<", "&lt;")
                            .replace(">", "&gt;"),
                        body_style
                    )
                )

        # ==========================================
        # STORY
        # ==========================================

        story = []

        # NAME
        story.append(
            Paragraph(
                name.upper(),
                name_style
            )
        )

        # JOB TITLE
        story.append(
            Paragraph(
                job_title.upper(),
                title_style
            )
        )

        # ==========================================
        # CONTACT BAR
        # ==========================================

        contacts = []

        if phone:
            contacts.append("Phone: " + phone)

        if address:
            contacts.append("Location: " + address)

        if email:
            contacts.append("Email: " + email)

        contact_text = "   |   ".join(contacts)

        if contact_text:

            contact_table = Table(
                [[
                    Paragraph(
                        contact_text,
                        contact_style
                    )
                ]],
                colWidths=[180 * mm]
            )

            contact_table.setStyle(
                TableStyle([
                    (
                        "LINEABOVE",
                        (0, 0),
                        (-1, 0),
                        0.7,
                        colors.HexColor("#1a1a2e")
                    ),
                    (
                        "LINEBELOW",
                        (0, 0),
                        (-1, 0),
                        0.7,
                        colors.HexColor("#1a1a2e")
                    ),
                    (
                        "TOPPADDING",
                        (0, 0),
                        (-1, 0),
                        5
                    ),
                    (
                        "BOTTOMPADDING",
                        (0, 0),
                        (-1, 0),
                        5
                    )
                ])
            )

            story.append(contact_table)

        # ==========================================
        # ABOUT ME
        # ==========================================

        if summary:

            story.append(
                Paragraph(
                    "ABOUT ME",
                    heading_style
                )
            )

            add_content(
                story,
                summary
            )

        # ==========================================
        # EXPERIENCE
        # ==========================================

        if experience:

            story.append(
                Paragraph(
                    "EXPERIENCE",
                    heading_style
                )
            )

            add_content(
                story,
                experience
            )

        # ==========================================
        # EDUCATION
        # ==========================================

        if education:

            story.append(
                Paragraph(
                    "EDUCATION",
                    heading_style
                )
            )

            add_content(
                story,
                education
            )

        # ==========================================
        # SKILLS
        # ==========================================

        if skills:

            story.append(
                Paragraph(
                    "SKILLS",
                    heading_style
                )
            )

            add_content(
                story,
                skills
            )

        # ==========================================
        # CERTIFICATIONS
        # ==========================================

        if certifications:

            story.append(
                Paragraph(
                    "CERTIFICATIONS & PROFESSIONAL DEVELOPMENT",
                    heading_style
                )
            )

            add_content(
                story,
                certifications
            )

        # ==========================================
        # DYNAMIC SECTIONS
        # ==========================================

        for section in dynamic_sections:

            section_name = clean_text(
                section.get("name", "")
            )

            section_content = clean_text(
                section.get("content", "")
            )

            if not section_name:
                continue

            if not section_content:
                continue

            story.append(
                Paragraph(
                    section_name.upper(),
                    heading_style
                )
            )

            add_content(
                story,
                section_content
            )

        # ==========================================
        # BUILD PDF
        # ==========================================

        doc.build(story)

        pdf_buffer.seek(0)

        return send_file(
            pdf_buffer,
            mimetype="application/pdf",
            as_attachment=True,
            download_name="Professional_Resume.pdf"
        )

    except Exception as e:

        print("PDF ERROR:", e)

        return (
            "PDF generation error: "
            + str(e)
        ), 500


####
@app.route("/download/resume1-download-docx", methods=["POST"])
def download_resume1_docx():

    try:
        data = request.get_json()

        if not data:
            return "No resume data received", 400

        # ==========================================
        # GET DATA
        # ==========================================

        name = data.get("name", "YOUR NAME")
        job_title = data.get(
            "jobTitle",
            "PROFESSIONAL TITLE"
        )

        phone = data.get("phone", "")
        address = data.get("address", "")
        email = data.get("email", "")

        summary = data.get("summary", "")
        experience = data.get("experience", "")
        education = data.get("education", "")
        skills = data.get("skills", "")
        certifications = data.get(
            "certifications",
            ""
        )

        dynamic_sections = data.get(
            "dynamicSections",
            []
        )

        # ==========================================
        # CLEAN TEXT
        # ==========================================

        def clean_text(value):

            if value is None:
                return ""

            value = str(value)

            value = re.sub(
                r"<br\s*/?>",
                "\n",
                value,
                flags=re.IGNORECASE
            )

            value = re.sub(
                r"<[^>]+>",
                "",
                value
            )

            value = value.replace(
                "&nbsp;",
                " "
            )

            return value.strip()

        name = clean_text(name)
        job_title = clean_text(job_title)

        phone = clean_text(phone)
        address = clean_text(address)
        email = clean_text(email)

        summary = clean_text(summary)
        experience = clean_text(experience)
        education = clean_text(education)
        skills = clean_text(skills)
        certifications = clean_text(
            certifications
        )

        # ==========================================
        # CREATE WORD DOCUMENT
        # ==========================================

        document = Document()

        section = document.sections[0]

        section.top_margin = Inches(0.45)
        section.bottom_margin = Inches(0.45)
        section.left_margin = Inches(0.65)
        section.right_margin = Inches(0.65)

        # ==========================================
        # NORMAL FONT
        # ==========================================

        normal_style = document.styles["Normal"]

        normal_style.font.name = "Arial"
        normal_style.font.size = Pt(8.5)

        # ==========================================
        # NAME
        # ==========================================

        p = document.add_paragraph()

        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)

        run = p.add_run(
            name.upper()
        )

        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(24)

        # ==========================================
        # JOB TITLE
        # ==========================================

        p = document.add_paragraph()

        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        p.paragraph_format.space_after = Pt(7)

        run = p.add_run(
            job_title.upper()
        )

        run.font.name = "Arial"
        run.font.size = Pt(10)

        # ==========================================
        # CONTACT
        # ==========================================

        contacts = []

        if phone:
            contacts.append(
                "Phone: " + phone
            )

        if address:
            contacts.append(
                "Location: " + address
            )

        if email:
            contacts.append(
                "Email: " + email
            )

        if contacts:

            p = document.add_paragraph()

            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(8)

            run = p.add_run(
                "   |   ".join(contacts)
            )

            run.font.name = "Arial"
            run.font.size = Pt(8)

            # Bottom border
            pPr = p._p.get_or_add_pPr()

            pBdr = OxmlElement("w:pBdr")

            bottom = OxmlElement("w:bottom")

            bottom.set(
                qn("w:val"),
                "single"
            )

            bottom.set(
                qn("w:sz"),
                "8"
            )

            bottom.set(
                qn("w:space"),
                "4"
            )

            bottom.set(
                qn("w:color"),
                "1A1A2E"
            )

            pBdr.append(bottom)

            pPr.append(pBdr)

        # ==========================================
        # HEADING FUNCTION
        # ==========================================

        def add_heading(title):

            p = document.add_paragraph()

            p.paragraph_format.space_before = Pt(5)
            p.paragraph_format.space_after = Pt(3)

            run = p.add_run(
                title.upper()
            )

            run.bold = True
            run.font.name = "Arial"
            run.font.size = Pt(10.5)

            # Bottom border
            pPr = p._p.get_or_add_pPr()

            pBdr = OxmlElement("w:pBdr")

            bottom = OxmlElement("w:bottom")

            bottom.set(
                qn("w:val"),
                "single"
            )

            bottom.set(
                qn("w:sz"),
                "8"
            )

            bottom.set(
                qn("w:space"),
                "2"
            )

            bottom.set(
                qn("w:color"),
                "1A1A2E"
            )

            pBdr.append(bottom)

            pPr.append(pBdr)

        # ==========================================
        # CONTENT FUNCTION
        # ==========================================

        def add_content(content):

            content = clean_text(content)

            if not content:
                return

            for line in content.split("\n"):

                line = line.strip()

                if not line:
                    continue

                p = document.add_paragraph()

                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.0

                run = p.add_run(line)

                run.font.name = "Arial"
                run.font.size = Pt(8.5)

        # ==========================================
        # ABOUT
        # ==========================================

        if summary:

            add_heading("ABOUT ME")

            add_content(summary)

        # ==========================================
        # EXPERIENCE
        # ==========================================

        if experience:

            add_heading("EXPERIENCE")

            add_content(experience)

        # ==========================================
        # EDUCATION
        # ==========================================

        if education:

            add_heading("EDUCATION")

            add_content(education)

        # ==========================================
        # SKILLS
        # ==========================================

        if skills:

            add_heading("SKILLS")

            add_content(skills)

        # ==========================================
        # CERTIFICATIONS
        # ==========================================

        if certifications:

            add_heading(
                "CERTIFICATIONS & PROFESSIONAL DEVELOPMENT"
            )

            add_content(
                certifications
            )

        # ==========================================
        # DYNAMIC SECTIONS
        # ==========================================

        for section_data in dynamic_sections:

            section_name = clean_text(
                section_data.get(
                    "name",
                    ""
                )
            )

            section_content = clean_text(
                section_data.get(
                    "content",
                    ""
                )
            )

            if not section_name:
                continue

            if not section_content:
                continue

            add_heading(section_name)

            add_content(section_content)

        # ==========================================
        # SAVE DOCX
        # ==========================================

        docx_buffer = BytesIO()

        document.save(docx_buffer)

        docx_buffer.seek(0)

        return send_file(
            docx_buffer,
            mimetype=(
                "application/vnd.openxmlformats-"
                "officedocument.wordprocessingml.document"
            ),
            as_attachment=True,
            download_name="Professional_Resume.docx"
        )

    except Exception as e:

        print("DOCX ERROR:", e)

        return (
            "DOCX generation error: "
            + str(e)
        ), 500    


############### resume 2#####
###pdf route###
@app.route("/download/resume2_professional_pdf", methods=["POST"])
def resume2_professional_pdf():

    data = request.get_json() or {}

    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=32,
        leftMargin=32,
        topMargin=28,
        bottomMargin=28
    )

    styles = getSampleStyleSheet()

    # =====================================================
    # STYLES
    # =====================================================

    name_style = ParagraphStyle(
        "ResumeName",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=17,
        leading=19,
        textColor=colors.HexColor("#1e3a8a"),
        spaceAfter=3,
        spaceBefore=0
    )

    contact_style = ParagraphStyle(
        "Contact",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=9,
        textColor=colors.HexColor("#333333"),
        spaceAfter=0,
        spaceBefore=0
    )

    heading_style = ParagraphStyle(
        "SectionHeading",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=11,
        textColor=colors.HexColor("#1e3a8a"),
        spaceBefore=7,
        spaceAfter=3,
        keepWithNext=True
    )

    body_style = ParagraphStyle(
        "Body",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=10,
        textColor=colors.HexColor("#333333"),
        spaceBefore=0,
        spaceAfter=2
    )

    bullet_style = ParagraphStyle(
        "Bullet",
        parent=body_style,
        leftIndent=9,
        firstLineIndent=-5,
        spaceBefore=0,
        spaceAfter=1
    )

    # =====================================================
    # HELPER FUNCTION
    # =====================================================

    def clean_text(text):
        if not text:
            return ""

        text = str(text)

        # Remove excessive blank lines
        text = re.sub(r'\n\s*\n+', '\n', text)

        return text.strip()

    def add_content(story, content):

        content = clean_text(content)

        if not content:
            return

        lines = content.split("\n")

        for line in lines:

            line = line.strip()

            if not line:
                continue

            # Convert bullet points
            if line.startswith("•"):
                line = line[1:].strip()
                story.append(
                    Paragraph(
                        "• " + line,
                        bullet_style
                    )
                )

            elif line.startswith("-"):
                line = line[1:].strip()
                story.append(
                    Paragraph(
                        "• " + line,
                        bullet_style
                    )
                )

            else:
                story.append(
                    Paragraph(
                        line,
                        body_style
                    )
                )

    # =====================================================
    # HEADER
    # PHOTO LEFT + NAME/CONTACT RIGHT
    # =====================================================

    name = clean_text(
        data.get("name", "YOUR NAME")
    ).upper()

    address = clean_text(
        data.get("address", "")
    )

    phone = clean_text(
        data.get("phone", "")
    )

    email = clean_text(
        data.get("email", "")
    )

    linkedin = clean_text(
        data.get("linkedin", "")
    )

    # -----------------------------------------------------
    # PHOTO
    # -----------------------------------------------------

    photo_element = Paragraph(
        "+ PHOTO",
        ParagraphStyle(
            "PhotoPlaceholder",
            parent=body_style,
            fontSize=6,
            alignment=TA_LEFT,
            textColor=colors.HexColor("#666666")
        )
    )

    photo_data = data.get("photo", "")

    if photo_data and photo_data.startswith("data:image"):

        try:

            image_data = photo_data.split(",", 1)[1]

            image_bytes = base64.b64decode(
                image_data
            )

            photo_buffer = BytesIO(image_bytes)

            photo_element = Image(
                photo_buffer,
                width=55,
                height=65
            )

        except Exception as e:

            print("Photo error:", e)

    # -----------------------------------------------------
    # NAME + CONTACT
    # -----------------------------------------------------

    contact_elements = []

    contact_elements.append(
        Paragraph(
            name,
            name_style
        )
    )

    if address:
        contact_elements.append(
            Paragraph(
                f"<b>Address:</b> {address}",
                contact_style
            )
        )

    if phone:
        contact_elements.append(
            Paragraph(
                f"<b>Phone:</b> {phone}",
                contact_style
            )
        )

    if email:
        contact_elements.append(
            Paragraph(
                f"<b>Email:</b> {email}",
                contact_style
            )
        )

    if linkedin:
        contact_elements.append(
            Paragraph(
                f"<b>Website:</b> {linkedin}",
                contact_style
            )
        )

    # =====================================================
    # HEADER TABLE
    # =====================================================

    header_table = Table(
        [
            [
                photo_element,
                contact_elements
            ]
        ],
        colWidths=[68, 420],
        hAlign="LEFT"
    )

    header_table.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )

    story = []

    story.append(header_table)

    # Small gap after header
    story.append(Spacer(1, 5))

    # =====================================================
    # TOP LINE
    # =====================================================

    story.append(
        HRFlowable(
            width="100%",
            thickness=0.7,
            color=colors.HexColor("#888888"),
            spaceBefore=0,
            spaceAfter=2
        )
    )

    # =====================================================
    # SUMMARY
    # =====================================================

    summary = data.get("summary", "")

    if summary:

        section = []

        section.append(
            Paragraph(
                "SUMMARY",
                heading_style
            )
        )

        add_content(
            section,
            summary
        )

        story.append(
            KeepTogether(section)
        )

    # =====================================================
    # WORK EXPERIENCE
    # =====================================================

    experience = data.get("experience", "")

    if experience:

        section = []

        section.append(
            Paragraph(
                "WORK EXPERIENCE",
                heading_style
            )
        )

        add_content(
            section,
            experience
        )

        story.append(
            KeepTogether(section)
        )

    # =====================================================
    # EDUCATION
    # =====================================================

    education = data.get("education", "")

    if education:

        section = []

        section.append(
            Paragraph(
                "EDUCATION",
                heading_style
            )
        )

        add_content(
            section,
            education
        )

        story.append(
            KeepTogether(section)
        )

    # =====================================================
    # ADDITIONAL INFORMATION
    # =====================================================

    skills = data.get("skills", "")

    if skills:

        section = []

        section.append(
            Paragraph(
                "ADDITIONAL INFORMATION",
                heading_style
            )
        )

        add_content(
            section,
            skills
        )

        story.append(
            KeepTogether(section)
        )

    # =====================================================
    # DYNAMIC SECTIONS
    # =====================================================

    for section_data in data.get(
        "dynamicSections",
        []
    ):

        section_name = clean_text(
            section_data.get("name", "")
        )

        content = section_data.get(
            "content",
            ""
        )

        if section_name and clean_text(content):

            section = []

            section.append(
                Paragraph(
                    section_name.upper(),
                    heading_style
                )
            )

            add_content(
                section,
                content
            )

            story.append(
                KeepTogether(section)
            )

    # =====================================================
    # BUILD PDF
    # =====================================================

    doc.build(story)

    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="Professional_Resume.pdf",
        mimetype="application/pdf"
    )  
#####
####docx####

@app.route("/download/resume2_professional_docx", methods=["POST"])
def resume2_professional_docx():

    data = request.get_json() or {}

    document = Document()

    # =====================================================
    # PAGE SETUP
    # =====================================================

    section = document.sections[0]
    section.top_margin = Inches(0.35)
    section.bottom_margin = Inches(0.35)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    # =====================================================
    # DEFAULT FONT
    # =====================================================

    styles = document.styles

    normal_style = styles["Normal"]

    normal_style.font.name = "Arial"
    normal_style.font.size = Pt(8)

    # =====================================================
    # HELPER
    # =====================================================

    def remove_table_borders(table):

        tbl = table._tbl

        tblPr = tbl.tblPr

        borders = tblPr.first_child_found_in(
            "w:tblBorders"
        )

        if borders is None:

            borders = OxmlElement(
                "w:tblBorders"
            )

            tblPr.append(borders)

        for edge in (
            "top",
            "left",
            "bottom",
            "right",
            "insideH",
            "insideV"
        ):

            tag = "w:" + edge

            element = borders.find(
                qn(tag)
            )

            if element is None:

                element = OxmlElement(tag)

                borders.append(element)

            element.set(
                qn("w:val"),
                "nil"
            )

    def clean_text(text):

        if not text:
            return ""

        text = str(text)

        # Remove excessive blank lines
        text = re.sub(
            r'\n\s*\n+',
            '\n',
            text
        )

        return text.strip()

    # =====================================================
    # HEADER TABLE
    # PHOTO LEFT + NAME RIGHT
    # =====================================================

    header_table = document.add_table(
        rows=1,
        cols=2
    )

    header_table.autofit = False

    header_table.columns[0].width = Inches(0.75)
    header_table.columns[1].width = Inches(6.5)

    remove_table_borders(
        header_table
    )

    left_cell = header_table.cell(
        0,
        0
    )

    right_cell = header_table.cell(
        0,
        1
    )

    left_cell.vertical_alignment = (
        WD_CELL_VERTICAL_ALIGNMENT.TOP
    )

    right_cell.vertical_alignment = (
        WD_CELL_VERTICAL_ALIGNMENT.TOP
    )

    # =====================================================
    # PHOTO
    # =====================================================

    photo_data = data.get(
        "photo",
        ""
    )

    if (
        photo_data
        and photo_data.startswith("data:image")
    ):

        try:

            image_data = photo_data.split(
                ",",
                1
            )[1]

            image_bytes = base64.b64decode(
                image_data
            )

            photo_buffer = BytesIO(
                image_bytes
            )

            photo_para = left_cell.paragraphs[0]

            photo_para.paragraph_format.space_before = Pt(0)
            photo_para.paragraph_format.space_after = Pt(0)

            photo_run = photo_para.add_run()

            photo_run.add_picture(
                photo_buffer,
                width=Inches(0.58),
                height=Inches(0.68)
            )

        except Exception as e:

            print(
                "Photo error:",
                e
            )

    else:

        photo_para = left_cell.paragraphs[0]

        photo_para.paragraph_format.space_before = Pt(0)
        photo_para.paragraph_format.space_after = Pt(0)

        photo_run = photo_para.add_run(
            "+ PHOTO"
        )

        photo_run.font.name = "Arial"
        photo_run.font.size = Pt(7)

    # =====================================================
    # NAME
    # =====================================================

    name = clean_text(
        data.get(
            "name",
            "YOUR NAME"
        )
    ).upper()

    name_para = right_cell.paragraphs[0]

    name_para.paragraph_format.space_before = Pt(0)
    name_para.paragraph_format.space_after = Pt(2)

    name_run = name_para.add_run(
        name
    )

    name_run.bold = True
    name_run.font.name = "Arial"
    name_run.font.size = Pt(17)
    name_run.font.color.rgb = RGBColor(
        30,
        58,
        138
    )

    # =====================================================
    # CONTACT INFORMATION
    # =====================================================

    contact_items = []

    if data.get("address"):

        contact_items.append(
            "Address: " +
            clean_text(
                data["address"]
            )
        )

    if data.get("phone"):

        contact_items.append(
            "Phone: " +
            clean_text(
                data["phone"]
            )
        )

    if data.get("email"):

        contact_items.append(
            "Email: " +
            clean_text(
                data["email"]
            )
        )

    if data.get("linkedin"):

        contact_items.append(
            "Website: " +
            clean_text(
                data["linkedin"]
            )
        )

    for item in contact_items:

        p = right_cell.add_paragraph()

        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1

        run = p.add_run(item)

        run.font.name = "Arial"
        run.font.size = Pt(7.5)

    # =====================================================
    # SMALL GAP
    # =====================================================

    gap = document.add_paragraph()

    gap.paragraph_format.space_before = Pt(0)
    gap.paragraph_format.space_after = Pt(1)
    gap.paragraph_format.line_spacing = 1

    # =====================================================
    # HORIZONTAL LINE
    # =====================================================

    line = document.add_paragraph()

    line.paragraph_format.space_before = Pt(0)
    line.paragraph_format.space_after = Pt(3)

    line_run = line.add_run(
        "_" * 105
    )

    line_run.font.name = "Arial"
    line_run.font.size = Pt(5)

    # =====================================================
    # SECTION FUNCTION
    # =====================================================

    def add_section(
        title,
        content
    ):

        content = clean_text(
            content
        )

        if not content:
            return

        # -------------------------------------------------
        # HEADING
        # -------------------------------------------------

        heading = document.add_paragraph()

        heading.paragraph_format.space_before = Pt(5)
        heading.paragraph_format.space_after = Pt(2)
        heading.paragraph_format.line_spacing = 1

        heading_run = heading.add_run(
            title.upper()
        )

        heading_run.bold = True
        heading_run.font.name = "Arial"
        heading_run.font.size = Pt(9)
        heading_run.font.color.rgb = RGBColor(
            30,
            58,
            138
        )

        # -------------------------------------------------
        # CONTENT
        # -------------------------------------------------

        lines = content.split("\n")

        for line_text in lines:

            line_text = line_text.strip()

            if not line_text:
                continue

            paragraph = document.add_paragraph()

            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(1)
            paragraph.paragraph_format.line_spacing = 1

            # Bullet
            if (
                line_text.startswith("•")
                or line_text.startswith("-")
            ):

                if line_text.startswith("•"):
                    line_text = line_text[1:].strip()

                elif line_text.startswith("-"):
                    line_text = line_text[1:].strip()

                run = paragraph.add_run(
                    "• " + line_text
                )

            else:

                run = paragraph.add_run(
                    line_text
                )

            run.font.name = "Arial"
            run.font.size = Pt(7.5)

    # =====================================================
    # SUMMARY
    # =====================================================

    add_section(
        "SUMMARY",
        data.get(
            "summary",
            ""
        )
    )

    # =====================================================
    # WORK EXPERIENCE
    # =====================================================

    add_section(
        "WORK EXPERIENCE",
        data.get(
            "experience",
            ""
        )
    )

    # =====================================================
    # EDUCATION
    # =====================================================

    add_section(
        "EDUCATION",
        data.get(
            "education",
            ""
        )
    )

    # =====================================================
    # ADDITIONAL INFORMATION
    # =====================================================

    add_section(
        "ADDITIONAL INFORMATION",
        data.get(
            "skills",
            ""
        )
    )

    # =====================================================
    # DYNAMIC SECTIONS
    # =====================================================

    for section_data in data.get(
        "dynamicSections",
        []
    ):

        section_name = clean_text(
            section_data.get(
                "name",
                ""
            )
        )

        content = section_data.get(
            "content",
            ""
        )

        if section_name and clean_text(content):

            add_section(
                section_name,
                content
            )

    # =====================================================
    # SAVE DOCX
    # =====================================================

    buffer = BytesIO()

    document.save(
        buffer
    )

    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="Professional_Resume.docx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )
#####resume 3 pdf code######
@app.route("/download/professional_resume3_pdf", methods=["POST"])
def professional_resume3_pdf():

    data = request.get_json() or {}

    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=25,
        bottomMargin=25
    )

    styles = getSampleStyleSheet()
    story = []

    # =====================================================
    # STYLES
    # =====================================================

    name_style = ParagraphStyle(
        "ResumeName",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=16,
        leading=17,
        alignment=TA_CENTER,
        spaceBefore=0,
        spaceAfter=1
    )

    contact_style = ParagraphStyle(
        "Contact",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7,
        leading=8,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#333333"),
        spaceBefore=0,
        spaceAfter=3
    )

    heading_style = ParagraphStyle(
        "Heading",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=9,
        textColor=colors.black,
        spaceBefore=4,
        spaceAfter=1
    )

    body_style = ParagraphStyle(
        "Body",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=9,
        textColor=colors.HexColor("#222222"),
        spaceBefore=0,
        spaceAfter=1
    )

    # =====================================================
    # CLEAN TEXT FUNCTION
    # =====================================================

    def clean_content(content):

        if not content:
            return ""

        content = str(content)

        # Convert Windows line endings
        content = content.replace("\r\n", "\n")
        content = content.replace("\r", "\n")

        # Remove multiple blank lines
        lines = content.split("\n")

        cleaned_lines = []

        for line in lines:
            line = line.strip()

            # Skip completely empty lines
            if line:
                cleaned_lines.append(line)

        # Keep only useful lines
        content = "\n".join(cleaned_lines)

        return content.strip()

    # =====================================================
    # NAME
    # =====================================================

    name = clean_content(
        data.get("name", "JOHN DOE")
    )

    story.append(
        Paragraph(
            name.upper(),
            name_style
        )
    )

    # =====================================================
    # CONTACT
    # =====================================================

    contact_items = []

    if data.get("phone"):
        contact_items.append(
            clean_content(data["phone"])
        )

    if data.get("email"):
        contact_items.append(
            clean_content(data["email"])
        )

    if data.get("address"):
        contact_items.append(
            clean_content(data["address"])
        )

    if data.get("linkedin"):
        contact_items.append(
            clean_content(data["linkedin"])
        )

    if contact_items:

        story.append(
            Paragraph(
                " | ".join(contact_items),
                contact_style
            )
        )

    # =====================================================
    # SECTION FUNCTION
    # =====================================================

    def add_section(title, content):

        content = clean_content(content)

        if not content:
            return

        # -----------------------------
        # Section Heading
        # -----------------------------

        story.append(
            Paragraph(
                title.upper(),
                heading_style
            )
        )

        # -----------------------------
        # Horizontal Line
        # -----------------------------

        story.append(
            HRFlowable(
                width="100%",
                thickness=0.8,
                color=colors.black,
                spaceBefore=0,
                spaceAfter=1
            )
        )

        # -----------------------------
        # Clean content
        # -----------------------------

        lines = content.split("\n")

        formatted_lines = []

        for line in lines:

            line = line.strip()

            if not line:
                continue

            # Escape special characters
            line = (
                line
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
            )

            formatted_lines.append(line)

        # Join lines without blank lines
        final_content = "<br/>".join(formatted_lines)

        # -----------------------------
        # Content
        # -----------------------------

        story.append(
            Paragraph(
                final_content,
                body_style
            )
        )

    # =====================================================
    # PROFESSIONAL SUMMARY
    # =====================================================

    add_section(
        "Professional Summary",
        data.get("summary", "")
    )

    # =====================================================
    # PROFESSIONAL EXPERIENCE
    # =====================================================

    add_section(
        "Professional Experience",
        data.get("experience", "")
    )

    # =====================================================
    # EDUCATION
    # =====================================================

    add_section(
        "Education",
        data.get("education", "")
    )

    # =====================================================
    # SKILLS
    # =====================================================

    add_section(
        "Technical Skills & Core Competencies",
        data.get("skills", "")
    )

    # =====================================================
    # DYNAMIC SECTIONS
    # =====================================================

    for section in data.get("dynamicSections", []):

        section_name = clean_content(
            section.get("name", "")
        )

        content = clean_content(
            section.get("content", "")
        )

        if section_name and content:

            add_section(
                section_name,
                content
            )

    # =====================================================
    # BUILD PDF
    # =====================================================

    doc.build(story)

    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="Professional_Resume.pdf",
        mimetype="application/pdf"
    )
########docx code#####
@app.route("/download/professional_resume3_docx", methods=["POST"])
def professional_resume3_docx():

    data = request.get_json() or {}

    document = Document()

    # =====================================================
    # PAGE SETUP
    # =====================================================

    section = document.sections[0]

    section.top_margin = Inches(0.35)
    section.bottom_margin = Inches(0.35)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)

    # =====================================================
    # DEFAULT FONT
    # =====================================================

    normal_style = document.styles["Normal"]

    normal_style.font.name = "Arial"
    normal_style.font.size = Pt(8)

    # =====================================================
    # CLEAN TEXT FUNCTION
    # =====================================================

    def clean_content(content):

        if not content:
            return ""

        content = str(content)

        # Convert line endings
        content = content.replace("\r\n", "\n")
        content = content.replace("\r", "\n")

        # Remove HTML line/format tags
        content = content.replace("<br>", "\n")
        content = content.replace("<br/>", "\n")
        content = content.replace("<br />", "\n")

        content = content.replace("</div>", "\n")
        content = content.replace("</li>", "\n")

        # Convert list items to bullets
        content = content.replace("<li>", "• ")

        # Remove remaining HTML tags
        import re
        content = re.sub(r"<[^>]*>", "", content)

        # Remove multiple blank lines
        lines = content.split("\n")

        cleaned_lines = []

        for line in lines:

            line = line.strip()

            if line:
                cleaned_lines.append(line)

        # Join without empty lines
        content = "\n".join(cleaned_lines)

        return content.strip()

    # =====================================================
    # NAME
    # =====================================================

    name = clean_content(
        data.get("name", "JOHN DOE")
    )

    name_para = document.add_paragraph()

    name_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    name_para.paragraph_format.space_before = Pt(0)
    name_para.paragraph_format.space_after = Pt(1)
    name_para.paragraph_format.line_spacing = 1

    name_run = name_para.add_run(
        name.upper()
    )

    name_run.bold = True
    name_run.font.name = "Arial"
    name_run.font.size = Pt(16)

    # =====================================================
    # CONTACT
    # =====================================================

    contact_items = []

    if data.get("phone"):
        contact_items.append(
            clean_content(data["phone"])
        )

    if data.get("email"):
        contact_items.append(
            clean_content(data["email"])
        )

    if data.get("address"):
        contact_items.append(
            clean_content(data["address"])
        )

    if data.get("linkedin"):
        contact_items.append(
            clean_content(data["linkedin"])
        )

    if contact_items:

        contact_para = document.add_paragraph()

        contact_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        contact_para.paragraph_format.space_before = Pt(0)
        contact_para.paragraph_format.space_after = Pt(3)
        contact_para.paragraph_format.line_spacing = 1

        contact_run = contact_para.add_run(
            " | ".join(contact_items)
        )

        contact_run.font.name = "Arial"
        contact_run.font.size = Pt(7)

    # =====================================================
    # SECTION FUNCTION
    # =====================================================

    def add_section(title, content):

        content = clean_content(content)

        if not content:
            return

        # -------------------------------------------------
        # SECTION HEADING
        # -------------------------------------------------

        heading = document.add_paragraph()

        heading.paragraph_format.space_before = Pt(3)
        heading.paragraph_format.space_after = Pt(0)
        heading.paragraph_format.line_spacing = 1

        heading_run = heading.add_run(
            title.upper()
        )

        heading_run.bold = True
        heading_run.font.name = "Arial"
        heading_run.font.size = Pt(9)

        # -------------------------------------------------
        # SECTION LINE
        # -------------------------------------------------

        line = document.add_paragraph()

        line.paragraph_format.space_before = Pt(0)
        line.paragraph_format.space_after = Pt(1)
        line.paragraph_format.line_spacing = 1

        line_run = line.add_run(
            "_" * 100
        )

        line_run.font.name = "Arial"
        line_run.font.size = Pt(5)

        # -------------------------------------------------
        # CONTENT
        # -------------------------------------------------

        lines = content.split("\n")

        for line_text in lines:

            line_text = line_text.strip()

            if not line_text:
                continue

            paragraph = document.add_paragraph()

            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1

            run = paragraph.add_run(
                line_text
            )

            run.font.name = "Arial"
            run.font.size = Pt(8)

    # =====================================================
    # PROFESSIONAL SUMMARY
    # =====================================================

    add_section(
        "Professional Summary",
        data.get("summary", "")
    )

    # =====================================================
    # PROFESSIONAL EXPERIENCE
    # =====================================================

    add_section(
        "Professional Experience",
        data.get("experience", "")
    )

    # =====================================================
    # EDUCATION
    # =====================================================

    add_section(
        "Education",
        data.get("education", "")
    )

    # =====================================================
    # TECHNICAL SKILLS
    # =====================================================

    add_section(
        "Technical Skills & Core Competencies",
        data.get("skills", "")
    )

    # =====================================================
    # DYNAMIC SECTIONS
    # =====================================================

    for section_data in data.get("dynamicSections", []):

        section_name = clean_content(
            section_data.get("name", "")
        )

        content = clean_content(
            section_data.get("content", "")
        )

        if section_name and content:

            add_section(
                section_name,
                content
            )

    # =====================================================
    # SAVE DOCX
    # =====================================================

    buffer = BytesIO()

    document.save(buffer)

    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="Professional_Resume.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
#############

# ==============================
# Run Flask App
# ==============================

if __name__ == "__main__":
    app.run(debug=True)


