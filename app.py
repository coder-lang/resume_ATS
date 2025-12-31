# app.py - Pure Python ATS Resume Generator
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import os
import io
import json
import PyPDF2
from dotenv import load_dotenv
from openai import OpenAI

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE

# ---------------------------
# App & configuration
# ---------------------------
load_dotenv()

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'your-secret-key')

OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')
if not OPENAI_API_KEY:
    raise ValueError("OPENAI_API_KEY not found")

client = OpenAI(api_key=OPENAI_API_KEY)
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'txt'}

# ---------------------------
# File parsing helpers
# ---------------------------
def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_pdf(file_path: str) -> str:
    text = ""
    with open(file_path, 'rb') as f:
        pdf_reader = PyPDF2.PdfReader(f)
        for page in pdf_reader.pages:
            text += page.extract_text() or ""
    return text

def extract_text_from_docx(file_path: str) -> str:
    doc = Document(file_path)
    return '\n'.join([p.text for p in doc.paragraphs])

# ---------------------------
# AI extraction
# ---------------------------
def extract_resume_data_with_ai(resume_text: str) -> dict:
    prompt = f"""Extract structured information from this resume and return ONLY a JSON object.

CRITICAL RULES:
- Use null for ANY missing information (never use trailing commas)
- Use empty arrays [] for missing lists
- Extract ALL work experience entries
- Keep dates, names, and text EXACTLY as written
- Never invent or assume data
- Return VALID JSON without trailing commas

JSON Format:
{{
  "name": "Full Name or null",
  "email": "email or null",
  "phone": "phone or null",
  "address": "address or null",
  "linkedin": "LinkedIn URL or null",
  "github": "GitHub URL or null",
  "portfolio": "Portfolio URL or null",
  "summary": "Professional summary or null",
  
  "education": [
    {{
      "institution": "University Name",
      "degree": "Degree Name",
      "field": "Field of Study",
      "location": "City, State or null",
      "graduation_date": "Date",
      "gpa": "GPA or null",
      "honors": ["Honor 1"]
    }}
  ],
  
  "experience": [
    {{
      "company": "Company Name",
      "position": "Job Title",
      "location": "City, State or null",
      "start_date": "Date",
      "end_date": "Date or Present",
      "achievements": ["Achievement 1", "Achievement 2"]
    }}
  ],
  
  "projects": [
    {{
      "name": "Project Name",
      "description": "Brief description",
      "technologies": ["Tech 1", "Tech 2"],
      "achievements": ["Achievement 1"]
    }}
  ],
  
  "skills": {{
    "technical": ["Skill 1", "Skill 2"],
    "languages": ["Language 1"],
    "tools": ["Tool 1"],
    "certifications": ["Cert 1"]
  }},
  
  "leadership": [
    {{
      "organization": "Org Name",
      "role": "Role Title",
      "location": "Location or null",
      "start_date": "Date",
      "end_date": "Date or Present",
      "achievements": ["Achievement 1"]
    }}
  ]
}}

Resume:
{resume_text}

Return ONLY valid JSON without trailing commas."""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a resume parser. Return only valid JSON with null for missing data. Never use trailing commas."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.2,
            max_tokens=3000
        )
        result = response.choices[0].message.content.strip()
        
        # Clean response
        if result.startswith("```"):
            parts = result.split("```")
            result = parts[1] if len(parts) > 1 else result
        if result.lower().startswith("json"):
            result = result[4:].strip()
        
        # Remove trailing commas (common AI mistake)
        import re
        result = re.sub(r',(\s*[}\]])', r'\1', result)
        
        return json.loads(result)
    except json.JSONDecodeError as je:
        print(f"JSON parsing error: {str(je)}")
        print(f"AI Response: {result[:500]}...")
        
        # Try to fix common JSON issues
        try:
            import re
            # Remove trailing commas more aggressively
            fixed = re.sub(r',(\s*[}\]])', r'\1', result)
            # Remove comments if any
            fixed = re.sub(r'//.*?\n', '\n', fixed)
            fixed = re.sub(r'/\*.*?\*/', '', fixed, flags=re.DOTALL)
            return json.loads(fixed)
        except:
            raise ValueError(f"Could not parse AI response as JSON: {str(je)}")
    except Exception as e:
        print(f"AI extraction error: {str(e)}")
        raise

# ---------------------------
# Data validation
# ---------------------------
def has_value(val):
    """Check if value has actual data."""
    if val is None:
        return False
    if isinstance(val, str):
        val = val.strip().lower()
        if not val or val in ['null', 'none', 'n/a', 'na']:
            return False
        return True
    if isinstance(val, list):
        return len([x for x in val if has_value(x)]) > 0
    if isinstance(val, dict):
        return any(has_value(v) for v in val.values())
    return bool(val)

def clean_list(items):
    """Return only valid items from list."""
    if not items:
        return []
    return [str(x).strip() for x in items if has_value(x)]

# ---------------------------
# Resume Generation Engine
# ---------------------------
class ATSResumeGenerator:
    """Generates beautiful, ATS-friendly resumes from scratch."""
    
    def __init__(self):
        self.doc = Document()
        self.setup_document_styles()
    
    def setup_document_styles(self):
        """Configure document margins and styles."""
        sections = self.doc.sections
        for section in sections:
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.7)
            section.right_margin = Inches(0.7)
    
    def add_header(self, data):
        """Add name and contact information header."""
        if not has_value(data.get('name')):
            return
        
        # Name - Large, Bold, Centered
        name_para = self.doc.add_paragraph()
        name_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        name_run = name_para.add_run(data['name'].upper())
        name_run.font.size = Pt(20)
        name_run.font.bold = True
        name_run.font.color.rgb = RGBColor(0, 0, 0)
        name_para.paragraph_format.space_after = Pt(4)
        name_para.paragraph_format.space_before = Pt(0)
        
        # Contact Info - Centered, smaller
        contact_parts = []
        if has_value(data.get('email')):
            contact_parts.append(data['email'])
        if has_value(data.get('phone')):
            contact_parts.append(data['phone'])
        if has_value(data.get('address')):
            contact_parts.append(data['address'])
        
        if contact_parts:
            contact_para = self.doc.add_paragraph()
            contact_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            contact_run = contact_para.add_run(' | '.join(contact_parts))
            contact_run.font.size = Pt(10)
            contact_para.paragraph_format.space_after = Pt(3)
        
        # Links - Centered
        links = []
        if has_value(data.get('linkedin')):
            links.append(f"LinkedIn: {data['linkedin']}")
        if has_value(data.get('github')):
            links.append(f"GitHub: {data['github']}")
        if has_value(data.get('portfolio')):
            links.append(f"Portfolio: {data['portfolio']}")
        
        if links:
            links_para = self.doc.add_paragraph()
            links_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            links_run = links_para.add_run(' | '.join(links))
            links_run.font.size = Pt(9)
            links_run.font.color.rgb = RGBColor(0, 102, 204)
            links_para.paragraph_format.space_after = Pt(8)
    
    def add_section_header(self, title):
        """Add a section header with underline."""
        # Add spacing before section
        self.add_spacing(12)
        
        para = self.doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(title.upper())
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        
        para.paragraph_format.space_after = Pt(2)
        para.paragraph_format.space_before = Pt(0)
        
        # Horizontal line
        line_para = self.doc.add_paragraph()
        line_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        line_para.paragraph_format.space_after = Pt(10)
        line_run = line_para.add_run('─' * 80)
        line_run.font.size = Pt(10)
        line_run.font.color.rgb = RGBColor(100, 100, 100)
    
    def add_summary(self, data):
        """Add professional summary."""
        if not has_value(data.get('summary')):
            return
        
        self.add_section_header('Professional Summary')
        
        para = self.doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run = para.add_run(data['summary'])
        run.font.size = Pt(10)
        para.paragraph_format.space_after = Pt(12)
        para.paragraph_format.line_spacing = 1.15
    
    def add_education(self, data):
        """Add education section."""
        education = data.get('education', [])
        if not education:
            return
        
        self.add_section_header('Education')
        
        for edu in education:
            if not has_value(edu.get('institution')):
                continue
            
            # Institution and Location
            header_para = self.doc.add_paragraph()
            header_para.paragraph_format.space_after = Pt(2)
            
            inst_run = header_para.add_run(edu['institution'])
            inst_run.font.bold = True
            inst_run.font.size = Pt(11)
            
            if has_value(edu.get('location')):
                loc_run = header_para.add_run(f" — {edu['location']}")
                loc_run.font.size = Pt(10)
                loc_run.font.italic = True
            
            # Degree and Date
            degree_para = self.doc.add_paragraph()
            degree_para.paragraph_format.space_after = Pt(2)
            
            degree_parts = []
            if has_value(edu.get('degree')):
                degree_parts.append(edu['degree'])
            if has_value(edu.get('field')):
                degree_parts.append(edu['field'])
            
            if degree_parts:
                degree_run = degree_para.add_run(', '.join(degree_parts))
                degree_run.font.size = Pt(10)
                degree_run.font.italic = True
            
            if has_value(edu.get('graduation_date')):
                date_run = degree_para.add_run(f" | {edu['graduation_date']}")
                date_run.font.size = Pt(10)
            
            # GPA
            if has_value(edu.get('gpa')):
                gpa_para = self.doc.add_paragraph()
                gpa_para.paragraph_format.space_after = Pt(2)
                gpa_run = gpa_para.add_run(f"GPA: {edu['gpa']}")
                gpa_run.font.size = Pt(10)
            
            # Honors
            honors = clean_list(edu.get('honors', []))
            if honors:
                honors_para = self.doc.add_paragraph()
                honors_para.paragraph_format.space_after = Pt(10)
                honors_run = honors_para.add_run(f"Honors: {', '.join(honors)}")
                honors_run.font.size = Pt(10)
            else:
                self.add_spacing(10)
    
    def add_experience(self, data):
        """Add work experience section."""
        experience = data.get('experience', [])
        if not experience:
            return
        
        self.add_section_header('Professional Experience')
        
        for exp in experience:
            if not has_value(exp.get('company')) or not has_value(exp.get('position')):
                continue
            
            # Company and Location
            company_para = self.doc.add_paragraph()
            company_para.paragraph_format.space_after = Pt(2)
            company_run = company_para.add_run(exp['company'])
            company_run.font.bold = True
            company_run.font.size = Pt(11)
            
            if has_value(exp.get('location')):
                loc_run = company_para.add_run(f" — {exp['location']}")
                loc_run.font.size = Pt(10)
                loc_run.font.italic = True
            
            # Position and Dates
            position_para = self.doc.add_paragraph()
            position_para.paragraph_format.space_after = Pt(4)
            position_run = position_para.add_run(exp['position'])
            position_run.font.size = Pt(10)
            position_run.font.italic = True
            
            dates = []
            if has_value(exp.get('start_date')):
                dates.append(exp['start_date'])
            if has_value(exp.get('end_date')):
                dates.append(exp['end_date'])
            
            if dates:
                date_run = position_para.add_run(f" | {' – '.join(dates)}")
                date_run.font.size = Pt(10)
            
            # Achievements
            achievements = clean_list(exp.get('achievements', []))
            for achievement in achievements:
                bullet_para = self.doc.add_paragraph(style='List Bullet')
                bullet_para.paragraph_format.left_indent = Inches(0.3)
                bullet_para.paragraph_format.space_after = Pt(3)
                bullet_para.paragraph_format.line_spacing = 1.15
                bullet_run = bullet_para.add_run(achievement)
                bullet_run.font.size = Pt(10)
            
            self.add_spacing(10)
    
    def add_projects(self, data):
        """Add projects section."""
        projects = data.get('projects', [])
        if not projects:
            return
        
        self.add_section_header('Projects')
        
        for proj in projects:
            if not has_value(proj.get('name')):
                continue
            
            # Project Name
            name_para = self.doc.add_paragraph()
            name_para.paragraph_format.space_after = Pt(2)
            name_run = name_para.add_run(proj['name'])
            name_run.font.bold = True
            name_run.font.size = Pt(11)
            
            # Description
            if has_value(proj.get('description')):
                desc_para = self.doc.add_paragraph()
                desc_para.paragraph_format.space_after = Pt(2)
                desc_run = desc_para.add_run(proj['description'])
                desc_run.font.size = Pt(10)
                desc_run.font.italic = True
            
            # Technologies
            technologies = clean_list(proj.get('technologies', []))
            if technologies:
                tech_para = self.doc.add_paragraph()
                tech_para.paragraph_format.space_after = Pt(4)
                tech_run = tech_para.add_run(f"Technologies: {', '.join(technologies)}")
                tech_run.font.size = Pt(9)
                tech_run.font.color.rgb = RGBColor(80, 80, 80)
            
            # Achievements
            achievements = clean_list(proj.get('achievements', []))
            for achievement in achievements:
                bullet_para = self.doc.add_paragraph(style='List Bullet')
                bullet_para.paragraph_format.left_indent = Inches(0.3)
                bullet_para.paragraph_format.space_after = Pt(3)
                bullet_para.paragraph_format.line_spacing = 1.15
                bullet_run = bullet_para.add_run(achievement)
                bullet_run.font.size = Pt(10)
            
            self.add_spacing(10)
    
    def add_skills(self, data):
        """Add skills section."""
        skills = data.get('skills', {})
        if not skills or not any(has_value(v) for v in skills.values()):
            return
        
        self.add_section_header('Skills & Expertise')
        
        # Technical Skills
        technical = clean_list(skills.get('technical', []))
        if technical:
            skill_para = self.doc.add_paragraph()
            skill_para.paragraph_format.space_after = Pt(4)
            label_run = skill_para.add_run('Technical Skills: ')
            label_run.font.bold = True
            label_run.font.size = Pt(10)
            skills_run = skill_para.add_run(', '.join(technical))
            skills_run.font.size = Pt(10)
        
        # Languages
        languages = clean_list(skills.get('languages', []))
        if languages:
            lang_para = self.doc.add_paragraph()
            lang_para.paragraph_format.space_after = Pt(4)
            label_run = lang_para.add_run('Languages: ')
            label_run.font.bold = True
            label_run.font.size = Pt(10)
            lang_run = lang_para.add_run(', '.join(languages))
            lang_run.font.size = Pt(10)
        
        # Tools
        tools = clean_list(skills.get('tools', []))
        if tools:
            tools_para = self.doc.add_paragraph()
            tools_para.paragraph_format.space_after = Pt(4)
            label_run = tools_para.add_run('Tools & Platforms: ')
            label_run.font.bold = True
            label_run.font.size = Pt(10)
            tools_run = tools_para.add_run(', '.join(tools))
            tools_run.font.size = Pt(10)
        
        # Certifications
        certs = clean_list(skills.get('certifications', []))
        if certs:
            cert_para = self.doc.add_paragraph()
            cert_para.paragraph_format.space_after = Pt(10)
            label_run = cert_para.add_run('Certifications: ')
            label_run.font.bold = True
            label_run.font.size = Pt(10)
            cert_run = cert_para.add_run(', '.join(certs))
            cert_run.font.size = Pt(10)
    
    def add_leadership(self, data):
        """Add leadership & activities section."""
        leadership = data.get('leadership', [])
        if not leadership:
            return
        
        self.add_section_header('Leadership & Activities')
        
        for lead in leadership:
            if not has_value(lead.get('organization')) or not has_value(lead.get('role')):
                continue
            
            # Organization and Location
            org_para = self.doc.add_paragraph()
            org_para.paragraph_format.space_after = Pt(2)
            org_run = org_para.add_run(lead['organization'])
            org_run.font.bold = True
            org_run.font.size = Pt(11)
            
            if has_value(lead.get('location')):
                loc_run = org_para.add_run(f" — {lead['location']}")
                loc_run.font.size = Pt(10)
                loc_run.font.italic = True
            
            # Role and Dates
            role_para = self.doc.add_paragraph()
            role_para.paragraph_format.space_after = Pt(4)
            role_run = role_para.add_run(lead['role'])
            role_run.font.size = Pt(10)
            role_run.font.italic = True
            
            dates = []
            if has_value(lead.get('start_date')):
                dates.append(lead['start_date'])
            if has_value(lead.get('end_date')):
                dates.append(lead['end_date'])
            
            if dates:
                date_run = role_para.add_run(f" | {' – '.join(dates)}")
                date_run.font.size = Pt(10)
            
            # Achievements
            achievements = clean_list(lead.get('achievements', []))
            for achievement in achievements:
                bullet_para = self.doc.add_paragraph(style='List Bullet')
                bullet_para.paragraph_format.left_indent = Inches(0.3)
                bullet_para.paragraph_format.space_after = Pt(3)
                bullet_para.paragraph_format.line_spacing = 1.15
                bullet_run = bullet_para.add_run(achievement)
                bullet_run.font.size = Pt(10)
            
            self.add_spacing(10)
    
    def add_spacing(self, points):
        """Add vertical spacing."""
        para = self.doc.add_paragraph()
        para.paragraph_format.space_after = Pt(points)
    
    def generate(self, data):
        """Generate complete resume."""
        self.add_header(data)
        self.add_summary(data)
        self.add_education(data)
        self.add_experience(data)
        self.add_projects(data)
        self.add_skills(data)
        self.add_leadership(data)
        
        # Save to buffer
        buffer = io.BytesIO()
        self.doc.save(buffer)
        buffer.seek(0)
        return buffer

# ---------------------------
# Flask Routes
# ---------------------------
@app.route('/')
def index():
    return render_template('hybrid_form.html')

@app.route('/upload-and-extract', methods=['POST'])
def upload_and_extract():
    try:
        resume_text = ""
        
        if 'file' in request.files:
            file = request.files['file']
            if file and file.filename and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
                file.save(filepath)
                
                if filename.lower().endswith('.pdf'):
                    resume_text = extract_text_from_pdf(filepath)
                elif filename.lower().endswith('.docx'):
                    resume_text = extract_text_from_docx(filepath)
                
                try:
                    os.remove(filepath)
                except:
                    pass
        
        if not resume_text and request.form.get('text_input'):
            resume_text = request.form.get('text_input')
        
        if not resume_text:
            return jsonify({'error': 'No resume data provided'}), 400
        
        structured_data = extract_resume_data_with_ai(resume_text)
        return jsonify(structured_data)
    
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/generate-resume', methods=['POST'])
def generate_resume():
    try:
        data = request.get_json(force=True)
        
        # Validate required fields
        if not has_value(data.get('name')):
            return jsonify({'error': 'Name is required'}), 400
        
        # Generate resume
        generator = ATSResumeGenerator()
        resume_buffer = generator.generate(data)
        
        return send_file(
            resume_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=f"{data.get('name', 'Resume').replace(' ', '_')}_ATS_Resume.docx"
        )
    
    except Exception as e:
        print(f"Error generating resume: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
