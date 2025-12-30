
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

ALLOWED_EXTENSIONS = {'pdf', 'docx', 'txt'}  # reliable formats
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
    """
    Use AI to extract structured data from the resume; return dict.
    """
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

        # Strip markdown fences if present
        if result.startswith("```"):
            parts = result.split("```")
            # pick the largest chunk (usually the JSON)
            result = max(parts, key=len)
        if result.lower().startswith("json"):
            result = result[4:].strip()

        data = json.loads(result)
        return data

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
    """
    Tries to apply a list/bullet style. Falls back gracefully if a style is missing.
    """
    for style_name in BULLET_STYLE_CANDIDATES:
        try:
            p.style = doc.styles[style_name]
            return
        except KeyError:
            continue
    # Visual fallback if no list style exists
    if p.text.strip() and not p.text.strip().startswith('•'):
        p.text = f'• {p.text}'

def non_empty_list(items):
    return [str(x).strip() for x in (items or []) if x and str(x).strip()]

def replace_in_paragraph(paragraph, replacements: dict):
    """
    Replace placeholders in a paragraph. This approach re-sets runs to preserve correctness
    when placeholders are split across runs (sacrifices mixed formatting within the line).
    """
    text = paragraph.text
    replaced = False
    for key, value in replacements.items():
        if key in text:
            text = text.replace(key, str(value))
            replaced = True
    if replaced:
        # Remove existing runs
        while paragraph.runs:
            run = paragraph.runs[0]
            run._element.getparent().remove(run._element)
        # Add single run with replaced text
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
    """
    Remove all paragraphs strictly between start_header and end_header.
    If end_header is None or not found, remove until the end of document.
    """
    start_idx = find_header_index(doc, start_header)
    end_idx = find_header_index(doc, end_header) if end_header else None

    if start_idx is None:
        return

    # Determine removal range
    if end_idx is None:
        # remove everything after start header
        while len(doc.paragraphs) > start_idx + 1:
            target = doc.paragraphs[start_idx + 1]
            target._element.getparent().remove(target._element)
    else:
        count = end_idx - start_idx - 1
        for _ in range(max(0, count)):
            target = doc.paragraphs[start_idx + 1]
            target._element.getparent().remove(target._element)

def insert_before_index(doc: Document, idx: int, text: str = ""):
    """
    Insert a paragraph before the paragraph at 'idx'. If idx is None,
    append at the end of the document.
    """
    if idx is None or idx >= len(doc.paragraphs):
        p = doc.add_paragraph(text)
        return p
    anchor = doc.paragraphs[idx]
    p = anchor.insert_paragraph_before(text)
    return p

def fill_word_template(data: dict) -> io.BytesIO:
    """
    Fill the Word template with extracted data and rebuild sections.
    """
    doc = Document(TEMPLATE_PATH)

    # Log style inventory once for diagnostics
    try:
        para_styles = [s.name for s in doc.styles if getattr(s, 'type', None) == WD_STYLE_TYPE.PARAGRAPH]
        print("Template paragraph styles:", para_styles)
    except Exception:
        pass

    # Replacements for header fields/placeholders present in your template
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

    # Replace simple placeholders
    replace_text_in_doc(doc, replacements)

    # Update line-level composite sections
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

    # Compute key header indices
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
        # Remove everything between EXPERIENCE and the next header (LEADERSHIP or SKILLS)
        end_header = lead_header if lead_idx is not None else skills_header if skills_idx is not None else None
        remove_between(doc, exp_header, end_header)

        # Choose insertion anchor: right before the next header (or append if none)
        insert_anchor_idx = find_header_index(doc, end_header) if end_header else None

        for job in data.get('experience', []):
            company = job.get('company', 'Company Name')
            location = job.get('location', '')
            position = job.get('position', 'Position')
            detail = job.get('detail', '')
            start_date = job.get('start_date', 'Mnth Yr')
            end_date = job.get('end_date', 'Present')

            # Company header line
            p = insert_before_index(doc, insert_anchor_idx, "")
            p.add_run(f"{company}").bold = True
            if location:
                r = p.add_run(f" {location}")
                r.italic = True
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

            # Position + dates line
            p = insert_before_index(doc, insert_anchor_idx, "")
            run1 = p.add_run(f"{position}, {detail}" if detail else f"{position}")
            run1.italic = True
            p.add_run(f" {start_date} -- {end_date}")

            # Responsibilities bullets
            for resp in non_empty_list(job.get('responsibilities')):
                bp = insert_before_index(doc, insert_anchor_idx, resp)
                set_bullet_style(bp, doc)

            # Spacer (single blank line)
            insert_before_index(doc, insert_anchor_idx, "")

    # ---------------------------
    # Rebuild LEADERSHIP section
    # ---------------------------
    # Recompute indices in case EXPERIENCE changes shifted them
    lead_idx = find_header_index(doc, lead_header)
    skills_idx = find_header_index(doc, skills_header)

    if lead_idx is not None and data.get('leadership'):
        # Remove content between LEADERSHIP and SKILLS
        remove_between(doc, lead_header, skills_header if skills_idx is not None else None)

        # Insert before SKILLS or append
        insert_anchor_idx = find_header_index(doc, skills_header) if skills_idx is not None else None

        for activity in data.get('leadership', []):
            org = activity.get('organization', 'Organization')
            location = activity.get('location', '')
            position = activity.get('position', 'Position')
            detail = activity.get('detail', '')
            start_date = activity.get('start_date', 'Mnth Yr')
            end_date = activity.get('end_date', 'Present')

            # Organization header line
            p = insert_before_index(doc, insert_anchor_idx, "")
            p.add_run(f"{org}").bold = True
            if location:
                r = p.add_run(f" {location}")
                r.italic = True

            # Position + dates line
            p = insert_before_index(doc, insert_anchor_idx, "")
            run1 = p.add_run(f"{position}, {detail}" if detail else f"{position}")
            run1.italic = True
            p.add_run(f" {start_date} -- {end_date}")

            # Responsibilities bullets
            for resp in non_empty_list(activity.get('responsibilities')):
                bp = insert_before_index(doc, insert_anchor_idx, resp)
                set_bullet_style(bp, doc)

            # Spacer
            insert_before_index(doc, insert_anchor_idx, "")

        # Add affiliations under leadership if present
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

    # Save to buffer
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
    """
    Upload resume and extract structured data with AI.
    """
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

                if filename.lower().endswith('.pdf'):
                    resume_text = extract_text_from_pdf(filepath)
                elif filename.lower().endswith('.docx'):
                    resume_text = extract_text_from_docx(filepath)

                # Clean up temp file
                try:
                    os.remove(filepath)
                except Exception:
                    pass

        # Handle text input
        if not resume_text and request.form.get('text_input'):
            resume_text = request.form.get('text_input')

        if not resume_text:
            return jsonify({'error': 'No resume data provided'}), 400

        print("Extracting data with AI...")
        structured_data = extract_resume_data_with_ai(resume_text)

        # Avoid logging PII-heavy content; log keys only
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
    """
    Generate Word document from structured data using template.
    """
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
    # Startup diagnostics: show styles present in template
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
