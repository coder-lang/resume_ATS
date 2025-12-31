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

IMPORTANT: If information is not found in the resume, use null or empty arrays. Do NOT make up information.

Resume Text:
{resume_text}
Return ONLY the JSON object, no other text.
"""
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a resume parser. Return only valid JSON. Use null for missing fields, empty arrays for missing lists."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=2000
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

BULLET_STYLE_CANDIDATES = ['List Bullet', 'List Paragraph', 'Normal']

def set_bullet_style(p, doc):
    for style_name in BULLET_STYLE_CANDIDATES:
        try:
            p.style = doc.styles[style_name]
            return
        except KeyError:
            continue
    if p.text.strip() and not p.text.strip().startswith('•'):
        p.text = f'• {p.text}'

def replace_in_paragraph(paragraph, replacements: dict):
    text = paragraph.text
    replaced = False
    for key, value in replacements.items():
        if key in text:
            text = text.replace(key, str(value))
            replaced = True
    if replaced:
        while paragraph.runs:
            run = paragraph.runs[0]
            run._element.getparent().remove(run._element)
        paragraph.add_run(text)

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
        if header_text in p.text.strip():
            return i
    return None

def remove_section_completely(doc: Document, start_header: str, end_header: str = None):
    """Remove entire section including header."""
    start_idx = find_header_index(doc, start_header)
    if start_idx is None:
        return
    
    end_idx = find_header_index(doc, end_header) if end_header else None
    
    # Remove header itself
    header_para = doc.paragraphs[start_idx]
    header_para._element.getparent().remove(header_para._element)
    
    # Recalculate indices after header removal
    start_idx = find_header_index(doc, end_header) if end_header else None
    if start_idx is not None:
        # Remove content between (now that header is gone)
        while start_idx > 0 and (end_header is None or doc.paragraphs[start_idx - 1].text.strip() != ""):
            start_idx -= 1
            if start_idx < len(doc.paragraphs):
                target = doc.paragraphs[start_idx]
                if end_header and end_header in target.text:
                    break
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
    for p in list(doc.paragraphs):
        if text in p.text:
            p._element.getparent().remove(p._element)

def insert_before_index(doc: Document, idx: int, text: str = ""):
    if idx is None or idx >= len(doc.paragraphs):
        return doc.add_paragraph(text)
    anchor = doc.paragraphs[idx]
    return anchor.insert_paragraph_before(text)

def normalize_header(doc: Document, data: dict):
    """Fix the combined header and add name at top."""
    for i, p in enumerate(doc.paragraphs):
        if 'EDUCATION' in p.text and 'Your Name' in p.text:
            # Add name above as separate paragraph
            name_line = p.insert_paragraph_before(data.get('name', 'Your Name'))
            # Keep just EDUCATION
            p.text = 'EDUCATION'
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

    print(f"Section availability: Education={has_education}, Experience={has_experience}, Leadership={has_leadership}, Skills={has_skills}")

    # Basic replacements
    replacements = {
        'Your Name': data.get('name', 'Your Name'),
        '555 Your Address, NY 10005': data.get('address', '') if has_data(data.get('address')) else '',
        'your-email@gmail.edu': data.get('email', ''),
        '555.555.5555': data.get('phone', '') if has_data(data.get('phone')) else '',
    }

    # Education section replacements (only if has education data)
    if has_education:
        replacements.update({
            'Your University': data.get('university', 'Your University'),
            'Your College/School': data.get('college', 'Your College/School'),
            '3._ _': data.get('gpa', '') if has_data(data.get('gpa')) else '',
            'City, State': data.get('location', 'City, State'),
            'Your Major: Bachelor of XYZ, ABC': f"{data.get('major', 'Bachelor of XYZ, ABC')}",
            'Expected Graduation: Mnth Year': f"Expected Graduation: {data.get('graduation_date', 'Mnth Year')}",
            'Your Minor: DEF': f"{data.get('minor', '')}" if has_data(data.get('minor')) else '',
        })

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
        insert_anchor_idx = find_header_index(doc, end_header) if end_header else None

        for job in reversed(data.get('experience', [])):
            company = job.get('company', 'Company Name')
            location = job.get('location', '')
            position = job.get('position', 'Position')
            detail = job.get('detail', '')
            start_date = job.get('start_date', 'Mnth Yr')
            end_date = job.get('end_date', 'Present')

            p = insert_before_index(doc, insert_anchor_idx, "")
            p.add_run(f"{company}").bold = True
            if has_data(location):
                r = p.add_run(f" - {location}")
                r.italic = False
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

            p = insert_before_index(doc, insert_anchor_idx, "")
            run1 = p.add_run(f"{position}")
            if has_data(detail):
                run1.text += f", {detail}"
            run1.italic = True
            p.add_run(f" | {start_date} -- {end_date}")

            for resp in non_empty_list(job.get('responsibilities')):
                bp = insert_before_index(doc, insert_anchor_idx, resp)
                set_bullet_style(bp, doc)

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

            p = insert_before_index(doc, insert_anchor_idx, "")
            p.add_run(f"{org}").bold = True
            if has_data(location):
                r = p.add_run(f" - {location}")
                r.italic = False

            p = insert_before_index(doc, insert_anchor_idx, "")
            run1 = p.add_run(f"{position}")
            if has_data(detail):
                run1.text += f", {detail}"
            run1.italic = True
            p.add_run(f" | {start_date} -- {end_date}")

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
                    p.text = f"Language: {', '.join(languages)}"
        else:
            remove_line_containing(doc, 'Language:')
        
        if has_computer:
            computers = non_empty_list(data.get('computer_skills'))
            for p in doc.paragraphs:
                if 'Computer:' in p.text:
                    p.text = f"Computer: {', '.join(computers)}"
        else:
            remove_line_containing(doc, 'Computer:')
        
        if has_interests:
            interests = non_empty_list(data.get('interests'))
            for p in doc.paragraphs:
                if 'Interests:' in p.text:
                    p.text = f"Interests: {', '.join(interests)}"
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
        print("Extraction successful. Keys:", list(structured_data.keys()))
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
