
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
Resume Text:
{resume_text}
Return ONLY the JSON object, no other text.
"""
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

def non_empty_list(items):
    return [str(x).strip() for x in (items or []) if x and str(x).strip()]

def replace_in_paragraph(paragraph, replacements: dict):
    # Safe paragraph-level replacement (resets runs in the paragraph)
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
        if p.text.strip() == header_text:
            return i
    return None

def remove_between(doc: Document, start_header: str, end_header: str):
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
            target = doc.paragraphs[start_idx + 1]
            target._element.getparent().remove(target._element)

def insert_before_index(doc: Document, idx: int, text: str = ""):
    if idx is None or idx >= len(doc.paragraphs):
        return doc.add_paragraph(text)
    anchor = doc.paragraphs[idx]
    return anchor.insert_paragraph_before(text)

def normalize_header(doc: Document, data: dict):
    """
    Fix the combined 'EDUCATIONYour Name' paragraph by splitting into:
    - 'EDUCATION' (kept as a header paragraph)
    - A new paragraph with the user's Name
    """
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == 'EDUCATIONYour Name':
            # Put Name ABOVE the 'EDUCATION' header (as a top header line)
            name_line = p.insert_paragraph_before(data.get('name', 'Your Name'))
            # Keep 'EDUCATION' as its own header
            p.text = 'EXPERIENCE' if False else 'EDUCATION'  # keep EDUCATION
            return  # after first correction, stop
    # If not found, do nothing

def clear_placeholder_lines(doc: Document, data: dict):
    """
    If certain sections are missing, clear placeholder lines to avoid dummy text.
    """
    has_honors = bool(non_empty_list(data.get('honors')))
    has_scholarships = bool(non_empty_list(data.get('scholarships')))
    has_coursework = bool(non_empty_list(data.get('coursework')))
    has_interests = bool(non_empty_list(data.get('interests')))

    for p in doc.paragraphs:
        t = p.text.strip()
        if t.startswith('Honors/Awards:') and not has_honors:
            p.text = 'Honors/Awards:'
        elif t.startswith('Scholarships:') and not has_scholarships:
            p.text = 'Scholarships:'
        elif t.startswith('Relevant Coursework:') and not has_coursework:
            p.text = 'Relevant Coursework:'
        elif t.startswith('Interests:') and not has_interests:
            p.text = 'Interests:'

def fill_word_template(data: dict) -> io.BytesIO:
    doc = Document(TEMPLATE_PATH)

    # Startup diagnostics
    try:
        para_styles = [s.name for s in doc.styles if getattr(s, 'type', None) == WD_STYLE_TYPE.PARAGRAPH]
        print("Template paragraph styles:", para_styles)
    except Exception:
        pass

    # Normalize the header that combines EDUCATION + Name
    normalize_header(doc, data)

    # Basic replacements (address, email, phone, etc.)
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

    replace_text_in_doc(doc, replacements)

    # Update merged lines for honors/scholarships/coursework
    for paragraph in doc.paragraphs:
        t = paragraph.text
        if 'Honors/Awards:' in t and data.get('honors'):
            honors_text = ', '.join(sorted(non_empty_list(data.get('honors'))))
            paragraph.text = f"Honors/Awards: {honors_text}"
        elif 'Scholarships:' in t and data.get('scholarships'):
            scholarships_text = ', '.join(sorted(non_empty_list(data.get('scholarships'))))
            paragraph.text = f"Scholarships: {scholarships_text}"
        elif 'Relevant Coursework:' in t and data.get('coursework'):
            coursework_text = ', '.join(sorted(non_empty_list(data.get('coursework'))))
            paragraph.text = f"Relevant Coursework: {coursework_text}"

    # Clear placeholders if sections are missing
    clear_placeholder_lines(doc, data)

    # Section headers
    exp_header = 'EXPERIENCE'
    lead_header = 'LEADERSHIP & PROFESSIONAL DEVELOPMENT'
    skills_header = 'SKILLS & INTERESTS'

    exp_idx = find_header_index(doc, exp_header)
    lead_idx = find_header_index(doc, lead_header)
    skills_idx = find_header_index(doc, skills_header)

    # ---------------------------
    # Rebuild EXPERIENCE section
    # ---------------------------
    if exp_idx is not None and data.get('experience'):
        end_header = lead_header if lead_idx is not None else skills_header if skills_idx is not None else None
        remove_between(doc, exp_header, end_header)
        insert_anchor_idx = find_header_index(doc, end_header) if end_header else None

        # Iterate REVERSED so final order is correct (first item ends up closest to header)
        for job in reversed(data.get('experience', [])):
            company = job.get('company', 'Company Name')
            location = job.get('location', '')
            position = job.get('position', 'Position')
            detail = job.get('detail', '')
            start_date = job.get('start_date', 'Mnth Yr')
            end_date = job.get('end_date', 'Present')

            p = insert_before_index(doc, insert_anchor_idx, "")
            p.add_run(f"{company}").bold = True
            if location:
                r = p.add_run(f" {location}")
                r.italic = True
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

            p = insert_before_index(doc, insert_anchor_idx, "")
            run1 = p.add_run(f"{position}, {detail}" if detail else f"{position}")
            run1.italic = True
            p.add_run(f" {start_date} -- {end_date}")

            for resp in non_empty_list(job.get('responsibilities')):
                bp = insert_before_index(doc, insert_anchor_idx, resp)
                set_bullet_style(bp, doc)

            insert_before_index(doc, insert_anchor_idx, "")

    # ---------------------------
    # Rebuild LEADERSHIP section
    # ---------------------------
    lead_idx = find_header_index(doc, lead_header)
    skills_idx = find_header_index(doc, skills_header)

    if lead_idx is not None and data.get('leadership'):
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
            if location:
                r = p.add_run(f" {location}")
                r.italic = True

            p = insert_before_index(doc, insert_anchor_idx, "")
            run1 = p.add_run(f"{position}, {detail}" if detail else f"{position}")
            run1.italic = True
            p.add_run(f" {start_date} -- {end_date}")

            for resp in non_empty_list(activity.get('responsibilities')):
                bp = insert_before_index(doc, insert_anchor_idx, resp)
                set_bullet_style(bp, doc)

            insert_before_index(doc, insert_anchor_idx, "")

        affs = non_empty_list(data.get('affiliations'))
        if affs:
            p = insert_before_index(doc, insert_anchor_idx, "")
            p.add_run('Other Affiliations: ').bold = True
            p.add_run(', '.join(sorted(affs)))

    # ---------------------------
    # Update SKILLS section lines
    # ---------------------------
    languages = non_empty_list(data.get('languages'))
    computers = non_empty_list(data.get('computer_skills'))
    interests = non_empty_list(data.get('interests'))

    for paragraph in doc.paragraphs:
        t = paragraph.text
        if 'Language:' in t and languages:
            paragraph.text = f"Language: {', '.join(languages)}"
        elif 'Computer:' in t and computers:
            paragraph.text = f"Computer: {', '.join(computers)}"
        elif 'Interests:' in t and interests:
            paragraph.text = f"Interests: {', '.join(sorted(interests))}"

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
    missing = [k for k in REQUIRED_FIELDS if not str(data.get(k, '')).strip()]
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
    try:
        if os.path.exists(TEMPLATE_PATH):
            d = Document(TEMPLATE_PATH)
            styles_present = [s.name for s in d.styles if getattr(s, 'type', None) == WD_STYLE_TYPE.PARAGRAPH]
            print("Template paragraph styles at startup:", styles_present)
        else:
            print(f"Template not found at path: {TEMPLATE_PATH}")
    except Exception as e:
        print(f"Could not inspect template styles: {e}")

    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
