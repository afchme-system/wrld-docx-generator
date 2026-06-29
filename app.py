import sys
import os
import io
import logging
from flask import Flask, request, send_file, jsonify
from functools import wraps
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ============================================================================
# INITIALIZATION
# ============================================================================

def get_drive_service():
    """Authenticate with Google Drive"""
    creds_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
    if not creds_json:
        raise ValueError("GOOGLE_CREDENTIALS_JSON environment variable not set")
    
    creds = Credentials.from_service_account_info(
        eval(creds_json),
        scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=creds)

drive_service = None
try:
    drive_service = get_drive_service()
except Exception as e:
    logger.warning(f"Drive service initialization deferred: {str(e)}")

# ============================================================================
# TEMPLATE CREATOR CLASS (Document Generation from JSON Specs)
# ============================================================================

class TemplateCreator:
    """Create new Word documents from JSON specifications"""
    
    def __init__(self):
        pass
    
    def create_document(self, title, sections, placeholders, metadata_fields=None, 
                       cover_page=True, header=True, footer=True, page_numbers=True):
        """
        Create a new Word document from specification.
        
        Args:
            title: Document title
            sections: List of section dicts with 'name' and optional 'placeholder'
            placeholders: List of placeholder strings (e.g., "[INVESTOR_NAME]")
            metadata_fields: Dict of document metadata
            cover_page: Whether to add a cover page
            header: Whether to add header
            footer: Whether to add footer
            page_numbers: Whether to add page numbers
        
        Returns:
            Binary .docx file content
        """
        doc = Document()
        
        # Set default font
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(11)
        
        # Add cover page if requested
        if cover_page:
            self._add_cover_page(doc, title, metadata_fields or {})
            doc.add_page_break()
        
        # Add table of contents placeholder
        doc.add_heading('Contents', level=1)
        doc.add_paragraph('[Table of Contents will be generated here]')
        doc.add_page_break()
        
        # Add sections
        for section in sections:
            section_name = section.get('name', 'Untitled Section')
            section_placeholder = section.get('placeholder', '')
            
            doc.add_heading(section_name, level=1)
            
            if section_placeholder:
                doc.add_paragraph(section_placeholder, style='Normal')
            else:
                doc.add_paragraph('[Content for ' + section_name + ']', style='Normal')
            
            doc.add_paragraph()  # spacing
        
        # Add placeholders section
        if placeholders:
            doc.add_page_break()
            doc.add_heading('Placeholders', level=1)
            for placeholder in placeholders:
                doc.add_paragraph(f'• {placeholder}', style='List Bullet')
        
        # Add header if requested
        if header:
            self._add_header(doc, title)
        
        # Add footer if requested
        if footer:
            self._add_footer(doc, page_numbers)
        
        # Save to bytes
        doc_bytes = io.BytesIO()
        doc.save(doc_bytes)
        doc_bytes.seek(0)
        
        logger.info(f"Created document: {title}, {len(sections)} sections, {len(placeholders)} placeholders")
        return doc_bytes.getvalue()
    
    def _add_cover_page(self, doc, title, metadata):
        """Add a professional cover page"""
        # Title
        title_para = doc.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_para.add_run(title)
        title_run.font.size = Pt(28)
        title_run.font.bold = True
        title_run.font.color.rgb = RGBColor(0, 51, 102)
        
        # Spacing
        doc.add_paragraph()
        doc.add_paragraph()
        
        # Metadata - check if it's a dict
        if metadata and isinstance(metadata, dict):
            for key, value in metadata.items():
                meta_para = doc.add_paragraph(f'{key}: {value}')
                meta_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if meta_para.runs:
                    meta_para.runs[0].font.size = Pt(11)
        
        # Date placeholder
        doc.add_paragraph()
        date_para = doc.add_paragraph('[Document Date: _______________]')
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if date_para.runs:
            date_para.runs[0].font.italic = True
    
    def _add_header(self, doc, title):
        """Add header to all sections"""
        section = doc.sections[0]
        header = section.header
        header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        header_para.text = title
        if header_para.runs:
            header_para.runs[0].font.size = Pt(10)
            header_para.runs[0].font.italic = True
    
    def _add_footer(self, doc, page_numbers=True):
        """Add footer to all sections"""
        section = doc.sections[0]
        footer = section.footer
        footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        
        if page_numbers:
            footer_para.text = "Page [#] of [##]"
        else:
            footer_para.text = "---"
        
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if footer_para.runs:
            footer_para.runs[0].font.size = Pt(10)

# Initialize creator
creator = TemplateCreator()

# ============================================================================
# ERROR HANDLER DECORATOR
# ============================================================================

def error_handler(f):
    """Decorator for consistent error handling"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        try:
            return f(*args, **kwargs)
        except ValueError as e:
            logger.warning(f"Validation error: {str(e)}")
            return jsonify({"error": str(e), "status": "validation_failed"}), 400
        except FileNotFoundError as e:
            logger.error(f"Template not found: {str(e)}")
            return jsonify({"error": str(e), "status": "template_not_found"}), 404
        except Exception as e:
            logger.error(f"Unexpected error: {str(e)}", exc_info=True)
            return jsonify({"error": str(e), "status": "error"}), 500
    return decorated_function

# ============================================================================
# ROUTES
# ============================================================================

@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({"status": "ok", "service": "WRLD Document Generator"}), 200

@app.route('/create-template', methods=['POST'])
@error_handler
def create_template():
    """
    POST /create-template
    
    Create a new Word document from JSON specification.
    
    Request body:
    {
        "title": "FoundMe Investor Funding Proposal Letter",
        "sections": [
            {"name": "Company Overview", "placeholder": "[COMPANY_OVERVIEW]"},
            {"name": "Investment Opportunity", "placeholder": "[INVESTMENT_OPPORTUNITY]"}
        ],
        "placeholders": ["[INVESTOR_NAME]", "[INVESTMENT_AMOUNT]", "[COMPANY_NAME]"],
        "output_filename": "investor-letter.docx",
        "cover_page": true,
        "header": true,
        "footer": true,
        "page_numbers": true
    }
    
    Response: Binary .docx file
    """
    
    data = request.get_json()
    if not data:
        raise ValueError("Request body must be JSON")
    
    title = data.get('title', 'Untitled Document')
    sections = data.get('sections', [])
    placeholders = data.get('placeholders', [])
    metadata_fields = data.get('metadata_fields', {})
    output_filename = data.get('output_filename', 'document.docx')
    cover_page = data.get('cover_page', True)
    header = data.get('header', True)
    footer = data.get('footer', True)
    page_numbers = data.get('page_numbers', True)
    
    if not output_filename:
        raise ValueError("output_filename is required")
    
    logger.info(f"Creating template: {title} → {output_filename}")
    
    # Create document from spec
    docx_binary = creator.create_document(
        title=title,
        sections=sections,
        placeholders=placeholders,
        metadata_fields=metadata_fields,
        cover_page=cover_page,
        header=header,
        footer=footer,
        page_numbers=page_numbers
    )
    
    # Return as downloadable file
    return send_file(
        io.BytesIO(docx_binary),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name=output_filename
    )

@app.route('/', methods=['GET'])
def root():
    """Root endpoint with API documentation"""
    return jsonify({
        "service": "WRLD Document Generator",
        "version": "1.2",
        "endpoints": {
            "POST /create-template": "Create a new Word document from JSON specification",
            "GET /health": "Health check"
        },
        "example": {
            "method": "POST",
            "endpoint": "/create-template",
            "body": {
                "title": "Investor Letter",
                "sections": [
                    {"name": "Overview", "placeholder": "[OVERVIEW]"}
                ],
                "placeholders": ["[INVESTOR_NAME]"],
                "output_filename": "investor-letter.docx"
            }
        }
    }), 200

# ============================================================================
# STARTUP
# ============================================================================

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
