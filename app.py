# app.py
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import os
import io
import json
import PyPDF2
from dotenv import load_dotenv

from openai import OpenAI

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor

# ---------------------------
# App & configuration
# ---------------------------
load_dotenv()

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'your-secret-key')

OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')
if not OPENAI_API_KEY:
    raise ValueError("OPENAI_API_KEY not found")

client = OpenAI(api_key=OPENAI_API_KEY)

ALLOWED_EXTENSIONS = {'pdf', 'docx', 'txt'}
TEMPLATE_PATH = 'resume_template.docx'

# ---------------------------
# Helpers: uploads & parsing
# ---------------------------
def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_pdf(file_path: str) -> str:
    text = ""
    with open(file_path, 'rb') as f:
        pdf_reader = PyPDF2.PdfReader(f)
        for page in pdf_reader.pages:
            page_text = page.extract_text() or ""
            text += page_text
    return text

def extract_text_from_docx(file_path: str) -> str:
    doc = Document(file_path)
    return '\n'.join([p.text for p in doc.paragraphs])

# ---------------------------
# Helpers: AI extraction
# ---------------------------
def extract_resume_data_with_ai(resume_text: str) -> dict:
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

IMPORTANT: 
- If information is not found in the resume, use null or empty arrays. Do NOT make up information.
- Extract ALL work experience entries from the resume in chronological order (most recent first)
- For dates, use format like "Oct 2010" or "03/2017" as they appear in resume
- Preserve the actual job titles and company names exactly as written

Resume Text:
{resume_text}
Return ONLY the JSON object, no other text.
"""
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a resume parser. Return only valid JSON. Use null for missing fields, empty arrays for missing lists. Extract ALL work experience entries accurately."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=2500
        )
        result = response.choices[0].message.content.strip()
        if result.startswith("```"):
            parts = result.split("```")
            result = max(parts, key=len)
        if result.lower().startswith("json"):
            result = result[4:].strip()
        return json.loads(result)
    except json.JSONDecodeError as je:
        raise ValueError(f"AI returned invalid JSON: {je}")
    except Exception as e:
        print(f"AI extraction error: {str(e)}")
        raise

# ---------------------------
# Helpers: Data validation
# ---------------------------
def non_empty_list(items):
    """Returns only non-empty, non-null items from a list."""
    if not items:
        return []
    return [str(x).strip() for x in items if x and str(x).strip() and str(x).strip().lower() != 'null']

def has_data(value):
    """Check if a value contains actual data."""
    if value is None:
        return False
    if isinstance(value, str):
        return bool(value.strip()) and value.strip().lower() != 'null'
    if isinstance(value, list):
        return len(non_empty_list(value)) > 0
    if isinstance(value, dict):
        return any(has_data(v) for v in value.values())
    return bool(value)

# ---------------------------
# Helpers: Word template ops
# ---------------------------

def set_bullet_style(p, doc):
    """Apply bullet formatting to paragraph."""
    # Try to use built-in list styles
    for style_name in ['List Bullet', 'List Paragraph', 'ListBullet']:
        try:
            p.style = doc.styles[style_name]
            return
        except KeyError:
            continue
    
    # Fallback: manually add bullet
    if p.text.strip() and not p.text.strip().startswith('•'):
        # Add bullet with proper indentation
        p.paragraph_format.left_indent = Pt(36)
        p.paragraph_format.first_line_indent = Pt(-18)
        p.text = f'• {p.text}'

def replace_in_paragraph(paragraph, replacements: dict):
    text = paragraph.text
    replaced = False
    for key, value in replacements.items():
        if key in text:
            text = text.replace(key, str(value))
            replaced = True
    if replaced:
        # Preserve formatting by keeping the first run's format
        if paragraph.runs:
            first_run_font = paragraph.runs[0].font
            bold = first_run_font.bold
            italic = first_run_font.italic
            size = first_run_font.size
        else:
            bold = italic = size = None
            
        while paragraph.runs:
            run = paragraph.runs[0]
            run._element.getparent().remove(run._element)
        
        new_run = paragraph.add_run(text)
        if bold is not None:
            new_run.font.bold = bold
        if italic is not None:
            new_run.font.italic = italic
        if size is not None:
            new_run.font.size = size

def replace_text_in_doc(doc: Document, replacements: dict):
    for paragraph in doc.paragraphs:
        replace_in_paragraph(paragraph, replacements)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph, replacements)

def find_header_index(doc: Document, header_text: str):
    for i, p in enumerate(doc.paragraphs):
        if header_text.upper() in p.text.strip().upper():
            return i
    return None

def remove_section_completely(doc: Document, start_header: str, end_header: str = None):
    """Remove entire section including header."""
    start_idx = find_header_index(doc, start_header)
    if start_idx is None:
        return
    
    end_idx = find_header_index(doc, end_header) if end_header else len(doc.paragraphs)
    
    # Remove from start_idx to end_idx
    for _ in range(end_idx - start_idx):
        if start_idx < len(doc.paragraphs):
            target = doc.paragraphs[start_idx]
            target._element.getparent().remove(target._element)

def remove_between(doc: Document, start_header: str, end_header: str):
    """Remove content between two headers (keep both headers)."""
    start_idx = find_header_index(doc, start_header)
    end_idx = find_header_index(doc, end_header) if end_header else None
    if start_idx is None:
        return
    if end_idx is None:
        while len(doc.paragraphs) > start_idx + 1:
            target = doc.paragraphs[start_idx + 1]
            target._element.getparent().remove(target._element)
    else:
        count = end_idx - start_idx - 1
        for _ in range(max(0, count)):
            if start_idx + 1 < len(doc.paragraphs):
                target = doc.paragraphs[start_idx + 1]
                target._element.getparent().remove(target._element)

def remove_line_containing(doc: Document, text: str):
    """Remove any paragraph containing specific text."""
    to_remove = []
    for p in doc.paragraphs:
        if text in p.text:
            to_remove.append(p)
    for p in to_remove:
        p._element.getparent().remove(p._element)

def insert_before_index(doc: Document, idx: int, text: str = ""):
    if idx is None or idx >= len(doc.paragraphs):
        return doc.add_paragraph(text)
    anchor = doc.paragraphs[idx]
    return anchor.insert_paragraph_before(text)

def normalize_header(doc: Document, data: dict):
    """Fix the combined header and add name at top."""
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        # Check if this is the combined header line
        if 'Your Name' in text:
            # Replace the entire paragraph with just the name
            p.clear()
            run = p.add_run(data.get('name', 'Your Name'))
            run.font.bold = True
            run.font.size = Pt(14)
            return

def fill_word_template(data: dict) -> io.BytesIO:
    doc = Document(TEMPLATE_PATH)

    # Normalize the header
    normalize_header(doc, data)

    # Determine what sections have data
    has_education = has_data(data.get('university')) or has_data(data.get('major'))
    has_honors = len(non_empty_list(data.get('honors'))) > 0
    has_scholarships = len(non_empty_list(data.get('scholarships'))) > 0
    has_coursework = len(non_empty_list(data.get('coursework'))) > 0
    has_experience = len(data.get('experience', [])) > 0
    has_leadership = len(data.get('leadership', [])) > 0
    has_affiliations = len(non_empty_list(data.get('affiliations'))) > 0
    has_languages = len(non_empty_list(data.get('languages'))) > 0
    has_computer = len(non_empty_list(data.get('computer_skills'))) > 0
    has_interests = len(non_empty_list(data.get('interests'))) > 0
    has_skills = has_languages or has_computer or has_interests

    print(f"Data check - Experience entries: {len(data.get('experience', []))}")
    for i, exp in enumerate(data.get('experience', [])):
        print(f"  Entry {i+1}: {exp.get('company')} - {exp.get('position')}")

    # Basic contact info replacements
    contact_line = []
    if has_data(data.get('address')):
        contact_line.append(data.get('address'))
    
    replacements = {
        'Your Name': data.get('name', 'Your Name'),
        '555 Your Address, NY 10005': data.get('address', '') if has_data(data.get('address')) else '',
        'your-email@gmail.edu': data.get('email', ''),
        '555.555.5555': data.get('phone', '') if has_data(data.get('phone')) else '',
    }

    # Update contact line to combine address
    for p in doc.paragraphs:
        if '555 Your Address' in p.text or 'your-email@gmail.edu' in p.text:
            # Rebuild contact line
            contact_parts = []
            if has_data(data.get('address')):
                contact_parts.append(data.get('address'))
            contact_line_text = '\n'.join(contact_parts) if contact_parts else ''
            
            email_phone_parts = []
            if has_data(data.get('email')):
                email_phone_parts.append(data.get('email'))
            if has_data(data.get('phone')):
                email_phone_parts.append(data.get('phone'))
            email_phone_line = ' | '.join(email_phone_parts)
            
            if contact_line_text and email_phone_line:
                p.text = f"{contact_line_text}\n{email_phone_line}"
            elif email_phone_line:
                p.text = email_phone_line
            elif contact_line_text:
                p.text = contact_line_text
            else:
                p.text = ''

    # Education section replacements
    if has_education:
        edu_replacements = {
            'Your University, Your College/School': f"{data.get('university', 'Your University')}, {data.get('college', '')}" if has_data(data.get('college')) else data.get('university', 'Your University'),
            'Your University': data.get('university', 'Your University'),
            'Your College/School': data.get('college', '') if has_data(data.get('college')) else '',
            'Cumulative GPA: 3._ _': f"Cumulative GPA: {data.get('gpa', '3._ _')}" if has_data(data.get('gpa')) else '',
            '3._ _': data.get('gpa', '') if has_data(data.get('gpa')) else '',
            'City, State': data.get('location', 'City, State'),
            'Your Major: Bachelor of XYZ, ABC': data.get('major', 'Bachelor of XYZ, ABC'),
            'Bachelor of XYZ, ABC': data.get('major', 'Bachelor of XYZ, ABC'),
            'Expected Graduation: Mnth Year': f"Expected Graduation: {data.get('graduation_date', 'Mnth Year')}" if has_data(data.get('graduation_date')) else '',
            'Mnth Year': data.get('graduation_date', '') if has_data(data.get('graduation_date')) else '',
            'Your Minor: DEF': f"Your Minor: {data.get('minor', '')}" if has_data(data.get('minor')) else '',
        }
        replace_text_in_doc(doc, edu_replacements)

    replace_text_in_doc(doc, replacements)

    # Handle EDUCATION section
    if not has_education:
        remove_section_completely(doc, 'EDUCATION', 'EXPERIENCE')
    else:
        # Update education sub-sections
        if has_honors:
            honors_text = ', '.join(non_empty_list(data.get('honors')))
            for p in doc.paragraphs:
                if 'Honors/Awards:' in p.text:
                    p.text = f"Honors/Awards: {honors_text}"
        else:
            remove_line_containing(doc, 'Honors/Awards:')
        
        if has_scholarships:
            scholarships_text = ', '.join(non_empty_list(data.get('scholarships')))
            for p in doc.paragraphs:
                if 'Scholarships:' in p.text:
                    p.text = f"Scholarships: {scholarships_text}"
        else:
            remove_line_containing(doc, 'Scholarships:')
        
        if has_coursework:
            coursework_text = ', '.join(non_empty_list(data.get('coursework')))
            for p in doc.paragraphs:
                if 'Relevant Coursework:' in p.text:
                    p.text = f"Relevant Coursework: {coursework_text}"
        else:
            remove_line_containing(doc, 'Relevant Coursework:')
        
        # Remove "Your Minor" line if no minor
        if not has_data(data.get('minor')):
            remove_line_containing(doc, 'Your Minor:')

    # Section headers
    exp_header = 'EXPERIENCE'
    lead_header = 'LEADERSHIP & PROFESSIONAL DEVELOPMENT'
    skills_header = 'SKILLS & INTERESTS'

    # Handle EXPERIENCE section
    if not has_experience:
        remove_section_completely(doc, exp_header, lead_header)
    else:
        exp_idx = find_header_index(doc, exp_header)
        lead_idx = find_header_index(doc, lead_header)
        skills_idx = find_header_index(doc, skills_header)
        
        end_header = lead_header if lead_idx is not None else skills_header if skills_idx is not None else None
        remove_between(doc, exp_header, end_header)
        
        # Re-find the insertion point after removal
        insert_anchor_idx = find_header_index(doc, end_header) if end_header else None

        # Sort experience by date (most recent first) and insert in reverse order
        experiences = data.get('experience', [])
        
        for job in reversed(experiences):
            company = job.get('company', 'Company Name')
            location = job.get('location', '')
            position = job.get('position', 'Position')
            detail = job.get('detail', '')
            start_date = job.get('start_date', 'Mnth Yr')
            end_date = job.get('end_date', 'Present')

            # Company name and location line (BOLD)
            p = insert_before_index(doc, insert_anchor_idx, "")
            company_run = p.add_run(company)
            company_run.font.bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

            # Position and date line (ITALIC position, normal dates)
            p = insert_before_index(doc, insert_anchor_idx, "")
            position_text = f"{position}"
            if has_data(detail):
                position_text += f", {detail}"
            
            position_run = p.add_run(position_text)
            position_run.font.italic = True
            
            # Add date with separator
            date_text = f" | {start_date} -- {end_date}"
            if has_data(location):
                date_text = f", {location}{date_text}"
            p.add_run(date_text)

            # Responsibilities as bullet points
            for resp in non_empty_list(job.get('responsibilities')):
                bp = insert_before_index(doc, insert_anchor_idx, resp)
                set_bullet_style(bp, doc)

            # Add blank line after each job
            insert_before_index(doc, insert_anchor_idx, "")

    # Handle LEADERSHIP section
    if not has_leadership:
        remove_section_completely(doc, lead_header, skills_header)
    else:
        lead_idx = find_header_index(doc, lead_header)
        skills_idx = find_header_index(doc, skills_header)
        
        remove_between(doc, lead_header, skills_header if skills_idx is not None else None)
        insert_anchor_idx = find_header_index(doc, skills_header) if skills_idx is not None else None

        for activity in reversed(data.get('leadership', [])):
            org = activity.get('organization', 'Organization')
            location = activity.get('location', '')
            position = activity.get('position', 'Position')
            detail = activity.get('detail', '')
            start_date = activity.get('start_date', 'Mnth Yr')
            end_date = activity.get('end_date', 'Present')

            # Organization name (BOLD)
            p = insert_before_index(doc, insert_anchor_idx, "")
            org_run = p.add_run(org)
            org_run.font.bold = True

            # Position and date line
            p = insert_before_index(doc, insert_anchor_idx, "")
            position_text = f"{position}"
            if has_data(detail):
                position_text += f", {detail}"
            
            position_run = p.add_run(position_text)
            position_run.font.italic = True
            
            date_text = f" | {start_date} -- {end_date}"
            if has_data(location):
                date_text = f", {location}{date_text}"
            p.add_run(date_text)

            # Responsibilities
            for resp in non_empty_list(activity.get('responsibilities')):
                bp = insert_before_index(doc, insert_anchor_idx, resp)
                set_bullet_style(bp, doc)

            insert_before_index(doc, insert_anchor_idx, "")

        if has_affiliations:
            affs = non_empty_list(data.get('affiliations'))
            p = insert_before_index(doc, insert_anchor_idx, "")
            p.add_run('Other Affiliations: ').bold = True
            p.add_run(', '.join(affs))
            insert_before_index(doc, insert_anchor_idx, "")

    # Handle SKILLS section
    if not has_skills:
        remove_section_completely(doc, skills_header, None)
    else:
        # Update or remove individual skill lines
        if has_languages:
            languages = non_empty_list(data.get('languages'))
            for p in doc.paragraphs:
                if 'Language:' in p.text:
                    lang_run = p.add_run()
                    p.clear()
                    p.add_run('Language: ').bold = True
                    p.add_run(', '.join(languages))
                    break
        else:
            remove_line_containing(doc, 'Language:')
        
        if has_computer:
            computers = non_empty_list(data.get('computer_skills'))
            found = False
            for p in doc.paragraphs:
                if 'Computer:' in p.text:
                    p.clear()
                    p.add_run('Computer: ').bold = True
                    p.add_run(', '.join(computers))
                    found = True
                    break
            
            # If Computer: line doesn't exist, add it
            if not found and has_computer:
                skills_idx = find_header_index(doc, skills_header)
                if skills_idx is not None:
                    p = insert_before_index(doc, skills_idx + 1, "")
                    p.add_run('Computer: ').bold = True
                    p.add_run(', '.join(computers))
        else:
            remove_line_containing(doc, 'Computer:')
        
        if has_interests:
            interests = non_empty_list(data.get('interests'))
            for p in doc.paragraphs:
                if 'Interests:' in p.text:
                    p.clear()
                    p.add_run('Interests: ').bold = True
                    p.add_run(', '.join(interests))
                    break
        else:
            remove_line_containing(doc, 'Interests:')

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------
# Validation for /generate-word
# ---------------------------
REQUIRED_FIELDS = ['name', 'email']

def validate_structured_data(data: dict):
    if not isinstance(data, dict):
        raise ValueError("Payload must be a JSON object")
    missing = [k for k in REQUIRED_FIELDS if not has_data(data.get(k))]
    if missing:
        raise ValueError(f"Missing required fields: {', '.join(missing)}")

# ---------------------------
# Routes
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
                except Exception:
                    pass

        if not resume_text and request.form.get('text_input'):
            resume_text = request.form.get('text_input')

        if not resume_text:
            return jsonify({'error': 'No resume data provided'}), 400

        print("Extracting data with AI...")
        structured_data = extract_resume_data_with_ai(resume_text)
        print("Extraction successful.")
        print(f"Found {len(structured_data.get('experience', []))} experience entries")
        return jsonify(structured_data)

    except ValueError as ve:
        return jsonify({'error': str(ve)}), 400
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback; traceback.print_exc()
        return jsonify({'error': 'Unexpected server error'}), 500

@app.route('/generate-word', methods=['POST'])
def generate_word():
    try:
        data = request.get_json(force=True)
        validate_structured_data(data)
        print("Generating Word document from template...")
        word_buffer = fill_word_template(data)
        return send_file(
            word_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='ATS_Resume.docx'
        )
    except ValueError as ve:
        return jsonify({'error': str(ve)}), 400
    except Exception as e:
        print(f"Error generating Word document: {str(e)}")
        import traceback; traceback.print_exc()
        return jsonify({'error': 'Unexpected server error'}), 500

# ---------------------------
# Entrypoint
# ---------------------------
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
