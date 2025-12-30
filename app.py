from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import os
from openai import OpenAI
from docx import Document
import PyPDF2
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.lib.colors import HexColor
import io
from dotenv import load_dotenv
import json

load_dotenv()

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'your-secret-key')

# Initialize OpenAI
openai_api_key = os.getenv('OPENAI_API_KEY')
if not openai_api_key:
    raise ValueError("OPENAI_API_KEY not found")

client = OpenAI(api_key=openai_api_key)

ALLOWED_EXTENSIONS = {'pdf', 'docx', 'doc', 'txt'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_pdf(file_path):
    text = ""
    with open(file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page in pdf_reader.pages:
            text += page.extract_text()
    return text

def extract_text_from_docx(file_path):
    doc = Document(file_path)
    return '\n'.join([paragraph.text for paragraph in doc.paragraphs])

def extract_resume_data_with_ai(resume_text):
    """Use AI to extract structured data from messy resume"""
    prompt = f"""Extract structured information from this resume and return ONLY a JSON object with this exact format:

{{
  "name": "Full Name",
  "email": "email@example.com",
  "phone": "phone number",
  "location": "City, State",
  "linkedin": "linkedin url or github",
  "summary": "2-3 sentence professional summary",
  "experience": [
    {{
      "title": "Job Title",
      "company": "Company Name",
      "location": "City, State",
      "dates": "Start - End",
      "responsibilities": ["achievement 1", "achievement 2", "achievement 3"]
    }}
  ],
  "education": [
    {{
      "degree": "Degree Name",
      "institution": "University Name",
      "location": "City, State",
      "year": "Graduation Year",
      "details": "GPA, honors, etc"
    }}
  ],
  "skills": [
    {{
      "category": "Category Name",
      "items": ["skill1", "skill2", "skill3"]
    }}
  ],
  "certifications": ["cert 1", "cert 2"]
}}

Resume Text:
{resume_text}

Return ONLY the JSON object, no other text."""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a resume parser. Return only valid JSON."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=2000
        )
        
        result = response.choices[0].message.content.strip()
        # Remove markdown code blocks if present
        if result.startswith('```'):
            result = result.split('```')[1]
            if result.startswith('json'):
                result = result[4:]
        
        data = json.loads(result)
        return data
    except Exception as e:
        print(f"AI extraction error: {str(e)}")
        raise

def create_ats_resume_pdf(data):
    """Create professional ATS-friendly resume from template"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch,
        leftMargin=0.75*inch,
        rightMargin=0.75*inch
    )
    
    # Colors
    primary_color = HexColor('#1e40af')
    text_color = HexColor('#1f2937')
    light_gray = HexColor('#6b7280')
    
    # Styles
    styles = getSampleStyleSheet()
    
    name_style = ParagraphStyle(
        'Name',
        parent=styles['Heading1'],
        fontSize=22,
        textColor=primary_color,
        spaceAfter=6,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    contact_style = ParagraphStyle(
        'Contact',
        parent=styles['Normal'],
        fontSize=10,
        textColor=text_color,
        spaceAfter=12,
        alignment=TA_CENTER,
        fontName='Helvetica'
    )
    
    section_header = ParagraphStyle(
        'SectionHeader',
        parent=styles['Heading2'],
        fontSize=12,
        textColor=primary_color,
        spaceBefore=12,
        spaceAfter=6,
        fontName='Helvetica-Bold'
    )
    
    job_title_style = ParagraphStyle(
        'JobTitle',
        fontSize=11,
        textColor=text_color,
        spaceAfter=2,
        fontName='Helvetica-Bold'
    )
    
    company_style = ParagraphStyle(
        'Company',
        fontSize=10,
        textColor=light_gray,
        spaceAfter=6,
        fontName='Helvetica'
    )
    
    body_style = ParagraphStyle(
        'Body',
        fontSize=10,
        textColor=text_color,
        spaceAfter=6,
        leading=14,
        fontName='Helvetica',
        alignment=TA_JUSTIFY
    )
    
    bullet_style = ParagraphStyle(
        'Bullet',
        fontSize=10,
        textColor=text_color,
        spaceAfter=4,
        leading=13,
        fontName='Helvetica',
        leftIndent=18
    )
    
    story = []
    
    # Header - Name
    name = data.get('name', 'Name Not Provided')
    story.append(Paragraph(name.upper(), name_style))
    
    # Contact Info
    contact_parts = []
    if data.get('email'):
        contact_parts.append(data['email'])
    if data.get('phone'):
        contact_parts.append(data['phone'])
    if data.get('location'):
        contact_parts.append(data['location'])
    if data.get('linkedin'):
        contact_parts.append(data['linkedin'])
    
    if contact_parts:
        contact_text = ' | '.join(contact_parts)
        story.append(Paragraph(contact_text, contact_style))
    
    # Full-width line
    page_width = letter[0] - 1.5*inch
    line = Table([['']], colWidths=[page_width])
    line.setStyle(TableStyle([
        ('LINEBELOW', (0, 0), (-1, -1), 2, primary_color),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(line)
    story.append(Spacer(1, 0.05*inch))
    
    # Professional Summary
    if data.get('summary'):
        story.append(Paragraph('PROFESSIONAL SUMMARY', section_header))
        summary_clean = data['summary'].replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        story.append(Paragraph(summary_clean, body_style))
    
    # Work Experience
    if data.get('experience'):
        story.append(Paragraph('WORK EXPERIENCE', section_header))
        for job in data['experience']:
            story.append(Spacer(1, 0.08*inch))
            title = job.get('title', 'Position').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            story.append(Paragraph(title, job_title_style))
            
            company_line = job.get('company', 'Company')
            if job.get('location'):
                company_line += f", {job['location']}"
            if job.get('dates'):
                company_line += f" | {job['dates']}"
            company_line = company_line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            story.append(Paragraph(company_line, company_style))
            
            if job.get('responsibilities'):
                for resp in job['responsibilities']:
                    resp_clean = resp.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    story.append(Paragraph(f"• {resp_clean}", bullet_style))
    
    # Education
    if data.get('education'):
        story.append(Paragraph('EDUCATION', section_header))
        for edu in data['education']:
            story.append(Spacer(1, 0.08*inch))
            degree = edu.get('degree', 'Degree').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            story.append(Paragraph(degree, job_title_style))
            
            edu_line = edu.get('institution', 'Institution')
            if edu.get('location'):
                edu_line += f", {edu['location']}"
            if edu.get('year'):
                edu_line += f" | {edu['year']}"
            edu_line = edu_line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            story.append(Paragraph(edu_line, company_style))
            
            if edu.get('details'):
                details = edu['details'].replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                story.append(Paragraph(details, body_style))
    
    # Skills
    if data.get('skills'):
        story.append(Paragraph('SKILLS', section_header))
        for skill_category in data['skills']:
            category_name = skill_category.get('category', 'Technical Skills')
            skills_list = ', '.join(skill_category.get('items', []))
            text = f"<b>{category_name}:</b> {skills_list}"
            text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            # Re-add bold tags after escaping
            text = text.replace('&lt;b&gt;', '<b>').replace('&lt;/b&gt;', '</b>')
            story.append(Paragraph(text, body_style))
    
    # Certifications
    if data.get('certifications'):
        story.append(Paragraph('CERTIFICATIONS', section_header))
        for cert in data['certifications']:
            cert_clean = cert.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            story.append(Paragraph(f"• {cert_clean}", bullet_style))
    
    doc.build(story)
    buffer.seek(0)
    return buffer

@app.route('/')
def index():
    return render_template('hybrid_form.html')

@app.route('/upload-and-extract', methods=['POST'])
def upload_and_extract():
    """Upload resume and extract structured data with AI"""
    try:
        resume_text = ""
        
        # Handle file upload
        if 'file' in request.files:
            file = request.files['file']
            if file and file.filename and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                
                os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
                file.save(filepath)
                
                if filename.endswith('.pdf'):
                    resume_text = extract_text_from_pdf(filepath)
                elif filename.endswith(('.docx', '.doc')):
                    resume_text = extract_text_from_docx(filepath)
                
                os.remove(filepath)
        
        # Handle text input
        if not resume_text and request.form.get('text_input'):
            resume_text = request.form.get('text_input')
        
        if not resume_text:
            return jsonify({'error': 'No resume data provided'}), 400
        
        # Extract structured data with AI
        print("Extracting data with AI...")
        structured_data = extract_resume_data_with_ai(resume_text)
        print("Extraction successful:", json.dumps(structured_data, indent=2))
        
        return jsonify(structured_data)
        
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/generate-pdf', methods=['POST'])
def generate_pdf():
    """Generate PDF from structured data"""
    try:
        data = request.json
        print("Generating PDF from data...")
        
        pdf_buffer = create_ats_resume_pdf(data)
        
        return send_file(
            pdf_buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name='ATS_Resume.pdf'
        )
    except Exception as e:
        print(f"Error generating PDF: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
