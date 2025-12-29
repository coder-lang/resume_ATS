from flask import Flask, render_template, request, jsonify, send_file, Response
from werkzeug.utils import secure_filename
import os
from openai import OpenAI
from docx import Document
import PyPDF2
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.enums import TA_LEFT, TA_CENTER
import io
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'default-secret-key-change-this')

# Initialize OpenAI client
openai_api_key = os.getenv('OPENAI_API_KEY')

if not openai_api_key:
    raise ValueError("OPENAI_API_KEY not found in environment variables. Please check your .env file.")

client = OpenAI(api_key=openai_api_key)

ALLOWED_EXTENSIONS = {'pdf', 'docx', 'doc'}

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
    text = []
    for paragraph in doc.paragraphs:
        text.append(paragraph.text)
    return '\n'.join(text)

def generate_ats_resume(resume_data):
    """Generate ATS-friendly resume using OpenAI"""
    prompt = f"""
    You are an expert resume writer specializing in ATS-friendly resumes. 
    Based on the following information, create a well-structured, ATS-optimized resume.
    
    Input data:
    {resume_data}
    
    Create a resume with these sections in order:
    1. Contact Information (Name, Email, Phone, Location, LinkedIn)
    2. Professional Summary (2-3 sentences)
    3. Work Experience (with bullet points for achievements)
    4. Education
    5. Skills (categorized if possible)
    6. Certifications (if applicable)
    
    Format the output as plain text with clear section headers.
    Use bullet points (•) for lists.
    Make it ATS-friendly by using standard section names and avoiding special characters.
    """
    
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {
                "role": "system",
                "content": "You are an expert resume writer specializing in ATS-optimized resumes."
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        temperature=0.7,
        max_tokens=2000
    )
    
    return response.choices[0].message.content

def generate_cover_letter(resume_data, job_description=""):
    """Generate cover letter using OpenAI"""
    prompt = f"""
    Based on the following resume information, create a professional cover letter.
    
    Resume Information:
    {resume_data}
    
    Job Description (if provided):
    {job_description}
    
    Create a compelling cover letter that:
    - Opens with enthusiasm
    - Highlights relevant experience
    - Shows how the candidate's skills match the role
    - Closes with a call to action
    
    Keep it professional and concise (3-4 paragraphs).
    """
    
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {
                "role": "system",
                "content": "You are an expert in writing professional cover letters."
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        temperature=0.7,
        max_tokens=1500
    )
    
    return response.choices[0].message.content

def create_pdf_resume(resume_text):
    """Create a PDF from resume text"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter,
                          topMargin=0.5*inch, bottomMargin=0.5*inch,
                          leftMargin=0.75*inch, rightMargin=0.75*inch)
    
    styles = getSampleStyleSheet()
    
    # Custom styles
    styles.add(ParagraphStyle(
        name='CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        textColor='#2C3E50',
        spaceAfter=12,
        alignment=TA_CENTER
    ))
    
    styles.add(ParagraphStyle(
        name='SectionHeader',
        parent=styles['Heading2'],
        fontSize=12,
        textColor='#34495E',
        spaceAfter=6,
        spaceBefore=12,
        alignment=TA_LEFT
    ))
    
    styles.add(ParagraphStyle(
        name='BodyText',
        parent=styles['Normal'],
        fontSize=10,
        leading=14,
        alignment=TA_LEFT
    ))
    
    story = []
    
    # Parse the resume text and format it
    lines = resume_text.split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            story.append(Spacer(1, 0.1*inch))
            continue
        
        # Detect headers (all caps or ending with colon)
        if line.isupper() or line.endswith(':'):
            story.append(Paragraph(line, styles['SectionHeader']))
        else:
            story.append(Paragraph(line, styles['BodyText']))
    
    doc.build(story)
    buffer.seek(0)
    return buffer

@app.route('/')
def index():
    # Read the HTML file directly and bypass any caching
    try:
        html_path = os.path.join(os.path.dirname(__file__), 'templates', 'index.html')
        print(f"Reading HTML from: {html_path}")  # Debug
        print(f"File exists: {os.path.exists(html_path)}")  # Debug
        
        with open(html_path, 'r', encoding='utf-8') as f:
            content = f.read()
            print(f"HTML file size: {len(content)} bytes")  # Debug
            
            # Create response with no-cache headers
            from flask import Response
            response = Response(content, mimetype='text/html')
            response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
            response.headers['Pragma'] = 'no-cache'
            response.headers['Expires'] = '0'
            return response
    except Exception as e:
        print(f"Error loading template: {str(e)}")  # Debug
        import traceback
        traceback.print_exc()
        return f"Error loading template: {str(e)}", 500

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files and 'text_input' not in request.form:
        return jsonify({'error': 'No file or text provided'}), 400
    
    resume_text = ""
    
    # Handle file upload
    if 'file' in request.files:
        file = request.files['file']
        if file and file.filename and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            
            os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
            file.save(filepath)
            
            # Extract text based on file type
            if filename.endswith('.pdf'):
                resume_text = extract_text_from_pdf(filepath)
            elif filename.endswith('.docx') or filename.endswith('.doc'):
                resume_text = extract_text_from_docx(filepath)
            
            os.remove(filepath)  # Clean up
    
    # Handle text input
    if 'text_input' in request.form and request.form['text_input']:
        resume_text = request.form['text_input']
    
    if not resume_text:
        return jsonify({'error': 'Could not extract text from file'}), 400
    
    return jsonify({'resume_text': resume_text})

@app.route('/generate-resume', methods=['POST'])
def generate_resume():
    data = request.json
    resume_data = data.get('resume_data', '')
    
    if not resume_data:
        return jsonify({'error': 'No resume data provided'}), 400
    
    try:
        ats_resume = generate_ats_resume(resume_data)
        print(f"Generated resume length: {len(ats_resume)}")  # Debug
        print(f"Resume preview: {ats_resume[:200]}")  # Debug
        return jsonify({'resume': ats_resume})
    except Exception as e:
        print(f"Error generating resume: {str(e)}")  # Debug
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Error: {str(e)}'}), 500

@app.route('/generate-cover-letter', methods=['POST'])
def generate_cover_letter_route():
    data = request.json
    resume_data = data.get('resume_data', '')
    job_description = data.get('job_description', '')
    
    if not resume_data:
        return jsonify({'error': 'No resume data provided'}), 400
    
    try:
        cover_letter = generate_cover_letter(resume_data, job_description)
        print(f"Generated cover letter length: {len(cover_letter)}")  # Debug
        return jsonify({'cover_letter': cover_letter})
    except Exception as e:
        print(f"Error generating cover letter: {str(e)}")  # Debug
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Error: {str(e)}'}), 500

@app.route('/download-pdf', methods=['POST'])
def download_pdf():
    data = request.json
    resume_text = data.get('resume_text', '')
    
    if not resume_text:
        return jsonify({'error': 'No resume text provided'}), 400
    
    try:
        pdf_buffer = create_pdf_resume(resume_text)
        return send_file(
            pdf_buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name='ATS_Resume.pdf'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5002))  # Render provides PORT env variable
    app.run(host='0.0.0.0', port=port, debug=False)  # host='0.0.0.0' allows external access
