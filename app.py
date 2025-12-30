from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import os
from openai import OpenAI
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import PyPDF2
import io
from dotenv import load_dotenv
import json
import re

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
TEMPLATE_PATH = 'resume_template.docx'

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
  "address": "Full Address",
  "university": "University Name",
  "college": "College/School Name",
  "gpa": "GPA (e.g., 3.85)",
  "location": "City, State",
  "major": "Major Name",
  "graduation_date": "Month Year",
  "minor": "Minor Name (if any)",
  "honors": ["Honor 1", "Honor 2", "Honor 3"],
  "scholarships": ["Scholarship 1", "Scholarship 2"],
  "coursework": ["Course 1", "Course 2", "Course 3"],
  "experience": [
    {{
      "company": "Company Name",
      "location": "City, State",
      "position": "Job Title",
      "detail": "Additional detail about role",
      "start_date": "Mnth Yr",
      "end_date": "Mnth Yr or Present",
      "responsibilities": ["achievement 1", "achievement 2", "achievement 3"]
    }}
  ],
  "leadership": [
    {{
      "organization": "Organization Name",
      "location": "City, State",
      "position": "Role/Position",
      "detail": "Additional detail",
      "start_date": "Mnth Yr",
      "end_date": "Mnth Yr or Present",
      "responsibilities": ["responsibility 1", "responsibility 2"]
    }}
  ],
  "affiliations": ["Affiliation 1", "Affiliation 2"],
  "languages": ["Language and proficiency"],
  "computer_skills": ["Skill 1", "Skill 2", "Skill 3"],
  "interests": ["Interest 1", "Interest 2", "Interest 3"]
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

def fill_word_template(data):
    """Fill the Word template with extracted data"""
    # Load template
    doc = Document(TEMPLATE_PATH)
    
    # Helper function to replace text in paragraphs
    def replace_in_paragraph(paragraph, replacements):
        for key, value in replacements.items():
            if key in paragraph.text:
                # Replace text while preserving formatting
                inline = paragraph.runs
                for run in inline:
                    if key in run.text:
                        run.text = run.text.replace(key, str(value))
    
    # Helper function to replace text in entire document
    def replace_text_in_doc(doc, replacements):
        for paragraph in doc.paragraphs:
            replace_in_paragraph(paragraph, replacements)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        replace_in_paragraph(paragraph, replacements)
    
    # Create replacements dictionary
    replacements = {
        'Your Name': data.get('name', 'Your Name'),
        '555 Your Address, NY 10005': data.get('address', ''),
        'your-email@gmail.edu': data.get('email', ''),
        '555.555.5555': data.get('phone', ''),
        'Your University': data.get('university', 'Your University'),
        'Your College/School': data.get('college', 'Your College/School'),
        '3._ _': data.get('gpa', '3._ _'),
        'City, State': data.get('location', 'City, State'),
        'Your Major: Bachelor of XYZ, ABC': f"Your Major: {data.get('major', 'Bachelor of XYZ, ABC')}",
        'Expected Graduation: Mnth Year': f"Expected Graduation: {data.get('graduation_date', 'Mnth Year')}",
        'Your Minor: DEF': f"Your Minor: {data.get('minor', 'DEF')}" if data.get('minor') else 'Your Minor: ',
    }
    
    # Replace basic info
    replace_text_in_doc(doc, replacements)
    
    # Find and update specific sections
    for i, paragraph in enumerate(doc.paragraphs):
        # Update Honors/Awards
        if 'Honors/Awards:' in paragraph.text and data.get('honors'):
            honors_text = ', '.join(sorted(data['honors']))
            paragraph.text = f"Honors/Awards: {honors_text}"
            
        # Update Scholarships
        elif 'Scholarships:' in paragraph.text and data.get('scholarships'):
            scholarships_text = ', '.join(sorted(data['scholarships']))
            paragraph.text = f"Scholarships: {scholarships_text}"
            
        # Update Relevant Coursework
        elif 'Relevant Coursework:' in paragraph.text and data.get('coursework'):
            coursework_text = ', '.join(sorted(data['coursework']))
            paragraph.text = f"Relevant Coursework: {coursework_text}"
    
    # Clear existing experience/leadership sections and rebuild
    # Find EXPERIENCE section
    exp_start_idx = None
    leadership_start_idx = None
    skills_start_idx = None
    
    for i, paragraph in enumerate(doc.paragraphs):
        if paragraph.text.strip() == 'EXPERIENCE':
            exp_start_idx = i
        elif paragraph.text.strip() == 'LEADERSHIP & PROFESSIONAL DEVELOPMENT':
            leadership_start_idx = i
        elif paragraph.text.strip() == 'SKILLS & INTERESTS':
            skills_start_idx = i
    
    # Rebuild EXPERIENCE section
    if exp_start_idx is not None and data.get('experience'):
        # Delete old content between EXPERIENCE and LEADERSHIP
        end_idx = leadership_start_idx if leadership_start_idx else skills_start_idx
        if end_idx:
            for _ in range(end_idx - exp_start_idx - 1):
                doc.paragraphs[exp_start_idx + 1]._element.getparent().remove(doc.paragraphs[exp_start_idx + 1]._element)
        
        # Add new experience entries
        for job in data['experience']:
            # Company header
            p = doc.paragraphs[exp_start_idx].insert_paragraph_before('')
            p.add_run(f"{job.get('company', 'Company Name')}").bold = True
            p.add_run(f" {job.get('location', '')}").italic = True
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            # Position and dates
            p = doc.paragraphs[exp_start_idx].insert_paragraph_before('')
            run1 = p.add_run(f"{job.get('position', 'Position')}, {job.get('detail', '')}")
            run1.italic = True
            p.add_run(f" {job.get('start_date', 'Mnth Yr')} -- {job.get('end_date', 'Present')}")
            
            # Responsibilities as bullets
            if job.get('responsibilities'):
                for resp in job['responsibilities']:
                    p = doc.paragraphs[exp_start_idx].insert_paragraph_before(resp)
                    p.style = 'List Bullet'
            
            # Spacing
            doc.paragraphs[exp_start_idx].insert_paragraph_before('')
    
    # Rebuild LEADERSHIP section
    if leadership_start_idx is not None and data.get('leadership'):
        # Find new leadership index after rebuilding experience
        leadership_start_idx = None
        for i, paragraph in enumerate(doc.paragraphs):
            if paragraph.text.strip() == 'LEADERSHIP & PROFESSIONAL DEVELOPMENT':
                leadership_start_idx = i
                break
        
        if leadership_start_idx:
            # Delete old content
            skills_start_idx = None
            for i, paragraph in enumerate(doc.paragraphs):
                if paragraph.text.strip() == 'SKILLS & INTERESTS':
                    skills_start_idx = i
                    break
            
            if skills_start_idx:
                for _ in range(skills_start_idx - leadership_start_idx - 1):
                    doc.paragraphs[leadership_start_idx + 1]._element.getparent().remove(doc.paragraphs[leadership_start_idx + 1]._element)
            
            # Add new leadership entries
            for activity in data['leadership']:
                # Organization header
                p = doc.paragraphs[leadership_start_idx].insert_paragraph_before('')
                p.add_run(f"{activity.get('organization', 'Organization')}").bold = True
                p.add_run(f" {activity.get('location', '')}").italic = True
                
                # Position and dates
                p = doc.paragraphs[leadership_start_idx].insert_paragraph_before('')
                run1 = p.add_run(f"{activity.get('position', 'Position')}, {activity.get('detail', '')}")
                run1.italic = True
                p.add_run(f" {activity.get('start_date', 'Mnth Yr')} -- {activity.get('end_date', 'Present')}")
                
                # Responsibilities as bullets
                if activity.get('responsibilities'):
                    for resp in activity['responsibilities']:
                        p = doc.paragraphs[leadership_start_idx].insert_paragraph_before(resp)
                        p.style = 'List Bullet'
                
                # Spacing
                doc.paragraphs[leadership_start_idx].insert_paragraph_before('')
            
            # Add affiliations if present
            if data.get('affiliations'):
                p = doc.paragraphs[leadership_start_idx].insert_paragraph_before('')
                p.add_run('Other Affiliations: ').bold = True
                p.add_run(', '.join(sorted(data['affiliations'])))
    
    # Update SKILLS section
    for paragraph in doc.paragraphs:
        if 'Language:' in paragraph.text and data.get('languages'):
            paragraph.text = f"Language: {', '.join(data['languages'])}"
        elif 'Computer:' in paragraph.text and data.get('computer_skills'):
            paragraph.text = f"Computer: {', '.join(data['computer_skills'])}"
        elif 'Interests:' in paragraph.text and data.get('interests'):
            paragraph.text = f"Interests: {', '.join(sorted(data['interests']))}"
    
    # Save to buffer
    buffer = io.BytesIO()
    doc.save(buffer)
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

@app.route('/generate-word', methods=['POST'])
def generate_word():
    """Generate Word document from structured data using template"""
    try:
        data = request.json
        print("Generating Word document from template...")
        
        word_buffer = fill_word_template(data)
        
        return send_file(
            word_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='ATS_Resume.docx'
        )
    except Exception as e:
        print(f"Error generating Word document: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
