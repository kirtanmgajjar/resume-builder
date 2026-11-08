import io
import json
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.opc.constants import RELATIONSHIP_TYPE
import streamlit as st
from streamlit_local_storage import LocalStorage
from xhtml2pdf import pisa

st.set_page_config(page_title="Resume Quick Apply", page_icon="📄", layout="centered")

local_storage = LocalStorage()

# --- DOCX Generation Engine ---
def generate_resume_docx(data):
    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    PRIMARY_COLOR = RGBColor(0, 0, 0)
    DARK_NEUTRAL = RGBColor(51, 51, 51)
    BORDER_HEX = "000000"

    def add_heading(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title)
        run.font.name = "Calibri"
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = PRIMARY_COLOR
        pBdr = parse_xml(
            f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="4" w:color="{BORDER_HEX}"/></w:pBdr>'
        )
        p._p.get_or_add_pPr().append(pBdr)

    def add_hyperlink(paragraph, text, url):
        part = paragraph.part
        r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), r_id)

        run = paragraph.add_run(text)
        run.font.name = "Calibri"
        run.font.size = Pt(10)
        run.font.color.rgb = DARK_NEUTRAL
        run.font.underline = False

        r_element = run._r
        paragraph._p.remove(r_element)
        hyperlink.append(r_element)
        paragraph._p.append(hyperlink)

    # 1. Header (Centered)
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_after = Pt(2)
    r_name = tp.add_run(data["name"])
    r_name.font.name = "Calibri"
    r_name.font.size = Pt(18)
    r_name.font.bold = True
    r_name.font.color.rgb = PRIMARY_COLOR

    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_after = Pt(12)

    contact = data["contact"]
    r_cnt = cp.add_run(f"{contact['phone']}  |  {contact['email']}  |  ")
    r_cnt.font.name = "Calibri"
    r_cnt.font.size = Pt(10)
    r_cnt.font.color.rgb = DARK_NEUTRAL

    add_hyperlink(cp, contact['linkedin_text'], contact['linkedin_url'])

    # 2. Professional Summary
    add_heading("Professional Summary")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_after = Pt(8)
    p_sum.paragraph_format.line_spacing = 1.15
    r_sum = p_sum.add_run(data["summary"])
    r_sum.font.name = "Calibri"
    r_sum.font.size = Pt(10.5)
    r_sum.font.color.rgb = DARK_NEUTRAL

    # 3. Experience
    add_heading("Experience")
    for exp in data["experience"]:
        tbl = doc.add_table(rows=2, cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.columns[0].width = Inches(4.8)
        tbl.columns[1].width = Inches(2.2)

        c0_l, c0_r = tbl.rows[0].cells
        p0_l = c0_l.paragraphs[0]
        p0_l.paragraph_format.space_after = Pt(0)
        r0_l = p0_l.add_run(exp["role"])
        r0_l.font.name = "Calibri"
        r0_l.font.bold = True
        r0_l.font.size = Pt(11)

        p0_r = c0_r.paragraphs[0]
        p0_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p0_r.paragraph_format.space_after = Pt(0)
        r0_r = p0_r.add_run(exp["location"])
        r0_r.font.name = "Calibri"
        r0_r.font.size = Pt(10.5)

        c1_l, c1_r = tbl.rows[1].cells
        p1_l = c1_l.paragraphs[0]
        p1_l.paragraph_format.space_after = Pt(4)
        r1_l = p1_l.add_run(exp["company"])
        r1_l.font.name = "Calibri"
        r1_l.font.italic = True
        r1_l.font.size = Pt(10.5)

        p1_r = c1_r.paragraphs[0]
        p1_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1_r.paragraph_format.space_after = Pt(4)
        r1_r = p1_r.add_run(exp["dates"])
        r1_r.font.name = "Calibri"
        r1_r.font.size = Pt(10.5)

        for bullet in exp["bullets"]:
            bullet = bullet.strip().replace('\u00A0', ' ')
            if not bullet:
                continue

            pb = doc.add_paragraph(style='List Bullet')
            pb.paragraph_format.space_after = Pt(3)
            pb.paragraph_format.line_spacing = 1.15

            if ": " in bullet:
                prefix, description = bullet.split(": ", 1)
                rb_bold = pb.add_run(f"{prefix}: ")
                rb_bold.font.name = "Calibri"
                rb_bold.font.size = Pt(10.5)
                rb_bold.font.bold = True

                rb_desc = pb.add_run(description)
                rb_desc.font.name = "Calibri"
                rb_desc.font.size = Pt(10.5)
            else:
                rb = pb.add_run(bullet)
                rb.font.name = "Calibri"
                rb.font.size = Pt(10.5)

    # 4. Technical Skills
    add_heading("Technical Skills")
    for cat, items in data["skills"]:
        if not cat.strip() and not items.strip():
            continue
        psk = doc.add_paragraph()
        psk.paragraph_format.space_after = Pt(3)
        rc = psk.add_run(f"{cat}: ")
        rc.font.name = "Calibri"
        rc.font.bold = True
        rc.font.size = Pt(10.5)
        ri = psk.add_run(items)
        ri.font.name = "Calibri"
        ri.font.size = Pt(10.5)

    # 5. Education
    add_heading("Education")
    for ed in data["education"]:
        ped_degree = doc.add_paragraph()
        ped_degree.paragraph_format.space_after = Pt(0)
        r_deg = ped_degree.add_run(ed["degree"])
        r_deg.font.name = "Calibri"
        r_deg.font.bold = True
        r_deg.font.size = Pt(10.5)

        ped_uni = doc.add_paragraph()
        ped_uni.paragraph_format.space_after = Pt(0)
        r_uni = ped_uni.add_run(f"{ed['university']} | {ed['dates']}")
        r_uni.font.name = "Calibri"
        r_uni.font.size = Pt(10.5)

        ped_score = doc.add_paragraph()
        ped_score.paragraph_format.space_after = Pt(4)
        r_score = ped_score.add_run(f"CPI: {ed['score']}")
        r_score.font.name = "Calibri"
        r_score.font.size = Pt(10.5)

    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream


# --- PDF Generation Engine (Direct HTML to PDF) ---
def generate_resume_pdf(data):
    bullets_html = ""
    for exp in data["experience"]:
        bullet_items = ""
        for b in exp["bullets"]:
            b = b.strip().replace('\u00A0', ' ')
            if ": " in b:
                prefix, desc = b.split(": ", 1)
                bullet_items += f"<li><strong>{prefix}:</strong> {desc}</li>"
            else:
                bullet_items += f"<li>{b}</li>"

        bullets_html += f"""
        <table style="width: 100%; margin-top: 4px;">
            <tr>
                <td style="font-weight: bold; font-size: 11pt; text-align: left;">{exp['role']}</td>
                <td style="font-size: 10.5pt; text-align: right;">{exp['location']}</td>
            </tr>
            <tr>
                <td style="font-style: italic; font-size: 10.5pt; text-align: left;">{exp['company']}</td>
                <td style="font-size: 10.5pt; text-align: right;">{exp['dates']}</td>
            </tr>
        </table>
        <ul style="margin-top: 4px; margin-bottom: 6px; padding-left: 20px;">
            {bullet_items}
        </ul>
        """

    skills_html = ""
    for cat, items in data["skills"]:
        if cat.strip() or items.strip():
            skills_html += f"<p style='margin: 3px 0;'><strong>{cat}:</strong> {items}</p>"

    edu_html = ""
    for ed in data["education"]:
        edu_html += f"""
        <p style="margin: 0; font-weight: bold;">{ed['degree']}</p>
        <p style="margin: 0;">{ed['university']} | {ed['dates']}</p>
        <p style="margin: 0 0 6px 0;">CPI: {ed['score']}</p>
        """

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        @page {{
            size: letter;
            margin: 0.65in 0.75in 0.65in 0.75in;
        }}
        body {{
            font-family: Helvetica, Arial, sans-serif;
            color: #333333;
            font-size: 10.5pt;
            line-height: 1.25;
        }}
        .header {{
            text-align: center;
            margin-bottom: 12px;
        }}
        .name {{
            font-size: 18pt;
            font-weight: bold;
            color: #000000;
            margin-bottom: 4px;
        }}
        .contact {{
            font-size: 10pt;
            color: #333333;
        }}
        .contact a {{
            color: #333333;
            text-decoration: none;
        }}
        h2 {{
            font-size: 12pt;
            font-weight: bold;
            color: #000000;
            border-bottom: 1px solid #000000;
            padding-bottom: 2px;
            margin-top: 12px;
            margin-bottom: 6px;
            text-transform: none;
        }}
        li {{
            margin-bottom: 3px;
        }}
    </style>
    </head>
    <body>
        <div class="header">
            <div class="name">{data['name']}</div>
            <div class="contact">
                {data['contact']['phone']} &nbsp;|&nbsp; {data['contact']['email']} &nbsp;|&nbsp;
                <a href="{data['contact']['linkedin_url']}">{data['contact']['linkedin_text']}</a>
            </div>
        </div>

        <h2>Professional Summary</h2>
        <p style="margin: 0 0 8px 0; text-align: justify;">{data['summary']}</p>

        <h2>Experience</h2>
        {bullets_html}

        <h2>Technical Skills</h2>
        {skills_html}

        <h2>Education</h2>
        {edu_html}
    </body>
    </html>
    """

    pdf_stream = io.BytesIO()
    pisa_status = pisa.CreatePDF(io.StringIO(html_content), dest=pdf_stream)
    pdf_stream.seek(0)
    return pdf_stream


# --- Retrieve Client Local Storage ---
saved_profile_raw = local_storage.getItem("my_resume_profile")
stored = {}
if saved_profile_raw:
    try:
        stored = json.loads(saved_profile_raw)
    except Exception:
        stored = {}

st.title("⚡ Resume Quick Apply")

with st.expander("👤 Contact Details", expanded=False):
    name = st.text_input("Name", value=stored.get("name", ""))
    phone = st.text_input("Phone", value=stored.get("phone", ""))
    email = st.text_input("Email", value=stored.get("email", ""))
    linkedin_url = st.text_input("LinkedIn URL", value=stored.get("linkedin_url", ""))

with st.expander("📝 Summary (Tailor for Job)", expanded=True):
    summary = st.text_area("Summary", value=stored.get("summary", ""), height=130)

with st.expander("💼 Experience Details", expanded=False):
    exp_role = st.text_input("Role", value=stored.get("role", ""))
    exp_company = st.text_input("Company", value=stored.get("company", ""))
    exp_location = st.text_input("Location", value=stored.get("location", ""))
    exp_dates = st.text_input("Dates", value=stored.get("dates", ""))
    bullets_text = st.text_area("Bullets (1 per line)", value=stored.get("bullets", ""), height=180)

with st.expander("🛠️ Skills & Education", expanded=False):
    skills_text = st.text_area("Skills (Format: Category: Skill 1, Skill 2)", value=stored.get("skills", ""), height=130)
    ed_degree = st.text_input("Degree", value=stored.get("degree", ""))
    ed_univ = st.text_input("Institution", value=stored.get("univ", ""))
    ed_dates = st.text_input("Period", value=stored.get("ed_dates", ""))
    ed_score = st.text_input("Score", value=stored.get("score", ""))

col_save, col_clear = st.columns(2)
with col_save:
    if st.button("💾 Save Profile to This Browser", use_container_width=True):
        payload_to_store = {
            "name": name,
            "phone": phone,
            "email": email,
            "linkedin_url": linkedin_url,
            "summary": summary,
            "role": exp_role,
            "company": exp_company,
            "location": exp_location,
            "dates": exp_dates,
            "bullets": bullets_text,
            "skills": skills_text,
            "degree": ed_degree,
            "univ": ed_univ,
            "ed_dates": ed_dates,
            "score": ed_score,
        }
        local_storage.setItem("my_resume_profile", json.dumps(payload_to_store))
        st.success("Saved to your browser's local storage!")
with col_clear:
    if st.button("🗑️ Clear Local Storage", use_container_width=True):
        local_storage.deleteItem("my_resume_profile")
        st.warning("Storage cleared. Refresh page to reset fields.")

# --- Parse Structured Data ---
parsed_bullets = [b.strip() for b in bullets_text.strip().split("\n") if b.strip()]
parsed_skills = []
for line in skills_text.strip().split("\n"):
    if ":" in line:
        category, items = line.split(":", 1)
        parsed_skills.append((category.strip(), items.strip()))

resume_payload = {
    "name": name if name else "Resume",
    "contact": {
        "phone": phone,
        "email": email,
        "linkedin_text": "LinkedIn",
        "linkedin_url": linkedin_url
    },
    "summary": summary,
    "experience": [
        {
            "role": exp_role,
            "company": exp_company,
            "location": exp_location,
            "dates": exp_dates,
            "bullets": parsed_bullets
        }
    ],
    "skills": parsed_skills,
    "education": [
        {
            "degree": ed_degree,
            "university": ed_univ,
            "dates": ed_dates,
            "score": ed_score
        }
    ]
}

# --- Modal Popup for Download Options ---
@st.dialog("Export Resume")
def download_modal():
    st.write("Configure your file details before downloading:")

    default_base_name = f"{(name if name else 'My').replace(' ', '_')}_Resume"
    custom_name = st.text_input("File Name (without extension)", value=default_base_name)
    file_format = st.radio("Choose Format", ["DOCX", "PDF"], horizontal=True)

    st.write("")

    if file_format == "DOCX":
        file_bytes = generate_resume_docx(resume_payload)
        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        final_filename = f"{custom_name}.docx"
    else:
        file_bytes = generate_resume_pdf(resume_payload)
        mime_type = "application/pdf"
        final_filename = f"{custom_name}.pdf"

    st.download_button(
        label=f"⬇️ Download {final_filename}",
        data=file_bytes,
        file_name=final_filename,
        mime=mime_type,
        use_container_width=True
    )

st.divider()

if st.button("⚡ Download Resume", use_container_width=True):
    download_modal()
