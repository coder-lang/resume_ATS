from reportlab.lib.colors import HexColor, black, white
from reportlab.platypus import Table, TableStylefrom flask import Flask, render_template, request, jsonify, send_file, Response
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
from reportlab.lib.colors import HexColor, black, white
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
    """Create a beautiful, ATS-friendly PDF with centered header and full-width line"""
    try:
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer, 
            pagesize=letter,
            topMargin=0.5*inch, 
            bottomMargin=0.5*inch,
            leftMargin=0.75*inch, 
            rightMargin=0.75*inch
        )
        
        # Define colors
        primary_color = HexColor('#1e40af')  # Professional blue
        text_color = HexColor('#1f2937')     # Dark gray
        light_gray = HexColor('#6b7280')     # Light gray for dates
        
        # Get base styles
        styles = getSampleStyleSheet()
        
        # Create custom styles
        name_style = ParagraphStyle(
            'NameStyle',
            parent=styles['Heading1'],
            fontSize=22,
            textColor=primary_color,
            spaceAfter=6,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold',
            leading=26
        )
        
        contact_style = ParagraphStyle(
            'ContactStyle',
            parent=styles['Normal'],
            fontSize=10,
            textColor=text_color,
            spaceAfter=12,
            alignment=TA_CENTER,
            fontName='Helvetica'
        )
        
        section_header_style = ParagraphStyle(
            'SectionHeader',
            parent=styles['Heading2'],
            fontSize=12,
            textColor=primary_color,
            spaceBefore=12,
            spaceAfter=6,
            fontName='Helvetica-Bold',
            leading=14
        )
        
        job_title_style = ParagraphStyle(
            'JobTitle',
            parent=styles['Normal'],
            fontSize=11,
            textColor=text_color,
            spaceAfter=2,
            fontName='Helvetica-Bold',
            leading=13
        )
        
        company_style = ParagraphStyle(
            'CompanyStyle',
            parent=styles['Normal'],
            fontSize=10,
            textColor=light_gray,
            spaceAfter=6,
            fontName='Helvetica',
            leading=12
        )
        
        body_style = ParagraphStyle(
            'BodyStyle',
            parent=styles['Normal'],
            fontSize=10,
            textColor=text_color,
            spaceAfter=6,
            leading=14,
            fontName='Helvetica',
            alignment=TA_JUSTIFY
        )
        
        bullet_style = ParagraphStyle(
            'BulletStyle',
            parent=styles['Normal'],
            fontSize=10,
            textColor=text_color,
            spaceAfter=4,
            leading=13,
            fontName='Helvetica',
            leftIndent=18,
            bulletIndent=0,
            alignment=TA_LEFT
        )
        
        story = []
        
        # Clean the resume text
        resume_text = resume_text.replace('```', '').strip()
        lines = [l.strip() for l in resume_text.split('\n') if l.strip()]
        
        # Parse resume sections
        i = 0
        found_name = False
        contact_lines = []
        
        # Extract header section (name + contact)
        while i < len(lines):
            line = lines[i]
            
            # First non-empty line is the name
            if not found_name:
                # Escape special characters
                clean_name = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                story.append(Paragraph(clean_name.upper(), name_style))
                found_name = True
                i += 1
                continue
            
            # Collect contact info lines
            lower_line = line.lower()
            if '@' in line or '|' in line or 'linkedin' in lower_line or 'github' in lower_line or \
               any(char.isdigit() for char in line[:30]) or 'http' in lower_line:
                contact_lines.append(line)
                i += 1
                continue
            
            # Stop when we hit non-contact content
            break
        
        # Add contact info (combine all contact lines)
        if contact_lines:
            # Join contact lines intelligently
            if any('|' in cl for cl in contact_lines):
                # Already has separators
                contact_text = ' '.join(contact_lines)
            else:
                # Add separators
                contact_text = ' | '.join(contact_lines)
            
            clean_contact = contact_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            story.append(Paragraph(clean_contact, contact_style))
        
        # Add full-width horizontal line below header
        page_width = letter[0] - 1.5*inch  # Full width minus margins
        header_line = Table([['']], colWidths=[page_width])
        header_line.setStyle(TableStyle([
            ('LINEBELOW', (0, 0), (-1, -1), 2, primary_color),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))
        story.append(header_line)
        story.append(Spacer(1, 0.05*inch))
        
        # Process remaining content
        while i < len(lines):
            line = lines[i]
            
            # Escape special characters
            line = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            
            # Detect section headers (all caps, 2-5 words)
            if line.isupper() and 2 <= len(line.split()) <= 5:
                story.append(Paragraph(line, section_header_style))
                i += 1
                continue
            
            # Detect bullet points
            if line.startswith('•') or line.startswith('-') or line.startswith('*'):
                bullet_text = line.lstrip('•-* ').strip()
                story.append(Paragraph(f'• {bullet_text}', bullet_style))
                i += 1
                continue
            
            # Detect job titles (followed by company/date line)
            if i + 1 < len(lines):
                next_line = lines[i + 1]
                # Check if next line has dates or location indicators
                has_date = any(x in next_line for x in ['20', '19', 'Present', 'Current', '–', '-'])
                has_location = ',' in next_line and len(next_line.split(',')) >= 2
                
                if (has_date or has_location) and len(line.split()) <= 10:
                    # This is a job title
                    story.append(Spacer(1, 0.08*inch))
                    story.append(Paragraph(line, job_title_style))
                    
                    # Next line is company/dates
                    i += 1
                    if i < len(lines):
                        company_line = lines[i].replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                        story.append(Paragraph(company_line, company_style))
                    i += 1
                    continue
            
            # Regular paragraph
            story.append(Paragraph(line, body_style))
            i += 1
        
        # Build the PDF
        doc.build(story)
        buffer.seek(0)
        return buffer
        
    except Exception as e:
        print(f"Error creating PDF: {str(e)}")
        import traceback
        traceback.print_exc()
        raise

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
        print(f"Creating PDF for text of length: {len(resume_text)}")  # Debug
        pdf_buffer = create_pdf_resume(resume_text)
        print("PDF created successfully")  # Debug
        
        return send_file(
            pdf_buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name='ATS_Resume.pdf'
        )
    except Exception as e:
        print(f"Error in download-pdf route: {str(e)}")  # Debug
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Failed to create PDF: {str(e)}'}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5002))  # Render provides PORT env variable
    app.run(host='0.0.0.0', port=port, debug=False)  # host='0.0.0.0' allows external access
