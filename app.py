import sys
import os
import io
import logging
from flask import Flask, request, send_file, jsonify
from functools import wraps
from docx import Document
from docx.shared import Pt, RGBColor, Inches
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
# GENERIC DOCUMENT BUILDER (JSON-Driven)
# ============================================================================

class DocumentBuilder:
    """Build Word documents from JSON component specifications"""
    
    def __init__(self):
        self.alignment_map = {
            'left': WD_ALIGN_PARAGRAPH.LEFT,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY
        }
    
    def build(self, components, margins=None):
        """
        Build document from JSON component array.
        
        Each component:
        {
            "type": "paragraph|heading|page_break",
            "content": "text content",
            "alignment": "left|right|center|justify",
            "font_name": "Calibri",
            "font_size": 11,
            "bold": false,
            "italic": false,
            "style": "Normal|List Bullet",
            "spacing_before": 0,
            "spacing_after": 12,
            "color": "#000000"
        }
        """
        doc = Document()
        
        # Set margins if provided
        if margins:
            for section in doc.sections:
                section.top_margin = Inches(margins.get('top', 1))
                section.bottom_margin = Inches(margins.get('bottom', 1))
                section.left_margin = Inches(margins.get('left', 1))
                section.right_margin = Inches(margins.get('right', 1))
        
        # Process each component
        for comp in components:
            self._render_component(doc, comp)
        
        # Save to bytes
        doc_bytes = io.BytesIO()
        doc.save(doc_bytes)
        doc_bytes.seek(0)
        
        return doc_bytes.getvalue()
    
    def _render_component(self, doc, comp):
        """Render a single component"""
        comp_type = comp.get('type', 'paragraph').lower()
        content = comp.get('content', '')
        
        if comp_type == 'page_break':
            doc.add_page_break()
        
        elif comp_type == 'heading':
            level = comp.get('level', 1)
            para = doc.add_heading(content, level=level)
            self._apply_formatting(para, comp)
        
        elif comp_type == 'paragraph':
            style = comp.get('style', 'Normal')
            para = doc.add_paragraph(content, style=style)
            self._apply_formatting(para, comp)
        
        elif comp_type == 'list_bullet':
            para = doc.add_paragraph(content, style='List Bullet')
            self._apply_formatting(para, comp)
        
        elif comp_type == 'blank':
            doc.add_paragraph()
    
    def _apply_formatting(self, para, comp):
        """Apply formatting to paragraph"""
        # Alignment
        alignment = comp.get('alignment', 'left').lower()
        para.alignment = self.alignment_map.get(alignment, WD_ALIGN_PARAGRAPH.LEFT)
        
        # Spacing
        if 'spacing_before' in comp:
            para.paragraph_format.space_before = Pt(comp['spacing_before'])
        if 'spacing_after' in comp:
            para.paragraph_format.space_after = Pt(comp['spacing_after'])
        
        # Apply to all runs
        font_name = comp.get('font_name', 'Calibri')
        font_size = comp.get('font_size', 11)
        bold = comp.get('bold', False)
        italic = comp.get('italic', False)
        color_hex = comp.get('color', '#000000')
        
        # Parse color
        try:
            color_rgb = RGBColor(int(color_hex[1:3], 16), int(color_hex[3:5], 16), int(color_hex[5:7], 16))
        except:
            color_rgb = RGBColor(0, 0, 0)
        
        for run in para.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.italic = italic
            run.font.color.rgb = color_rgb

# Initialize builder
builder = DocumentBuilder()

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
    
    Create a Word document from JSON component specification.
    
    Request body:
    {
        "title": "Document Title",
        "output_filename": "output.docx",
        "margins": {
            "top": 1,
            "bottom": 1,
            "left": 1,
            "right": 1
        },
        "components": [
            {
                "type": "paragraph",
                "content": "Your content here",
                "alignment": "left",
                "font_name": "Calibri",
                "font_size": 11,
                "bold": false,
                "italic": false,
                "spacing_before": 0,
                "spacing_after": 12
            },
            ...
        ],
        "placeholders": []
    }
    """
    
    data = request.get_json()
    if not data:
        raise ValueError("Request body must be JSON")
    
    title = data.get('title', 'Untitled Document')
    components = data.get('components', [])
    placeholders = data.get('placeholders', [])
    output_filename = data.get('output_filename', 'document.docx')
    margins = data.get('margins', {
        'top': 1,
        'bottom': 1,
        'left': 1,
        'right': 1
    })
    
    if not output_filename:
        raise ValueError("output_filename is required")
    
    if not components:
        raise ValueError("components array is required")
    
    logger.info(f"Creating document: {title} → {output_filename} ({len(components)} components)")
    
    # Add placeholders page if provided
    if placeholders:
        components.append({"type": "page_break"})
        components.append({
            "type": "heading",
            "level": 1,
            "content": "Placeholders",
            "font_name": "Calibri",
            "font_size": 11
        })
        for placeholder in placeholders:
            components.append({
                "type": "list_bullet",
                "content": placeholder,
                "font_name": "Calibri",
                "font_size": 11
            })
    
    # Build document
    docx_binary = builder.build(components, margins)
    
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
        "version": "2.0",
        "description": "Generic JSON-driven document builder. AI generates JSON specs, app.py renders them.",
        "endpoints": {
            "POST /create-template": "Create document from JSON component spec",
            "GET /health": "Health check"
        },
        "component_types": ["paragraph", "heading", "page_break", "list_bullet", "blank"],
        "example": {
            "method": "POST",
            "endpoint": "/create-template",
            "body": {
                "title": "Formal Letter",
                "output_filename": "letter.docx",
                "margins": {"top": 1, "bottom": 1, "left": 1, "right": 1},
                "components": [
                    {
                        "type": "paragraph",
                        "content": "Your Name\nYour Address\nYour Email",
                        "alignment": "right",
                        "font_name": "Calibri",
                        "font_size": 10,
                        "spacing_after": 24
                    },
                    {
                        "type": "paragraph",
                        "content": "29 June 2026",
                        "alignment": "right",
                        "font_name": "Calibri",
                        "font_size": 11,
                        "spacing_after": 24
                    },
                    {
                        "type": "paragraph",
                        "content": "Recipient Name\nRecipient Title\nOrganization",
                        "alignment": "left",
                        "font_name": "Calibri",
                        "font_size": 11,
                        "spacing_after": 24
                    },
                    {
                        "type": "paragraph",
                        "content": "RE: Subject Line",
                        "alignment": "left",
                        "font_name": "Calibri",
                        "font_size": 11,
                        "bold": True,
                        "spacing_before": 12,
                        "spacing_after": 12
                    },
                    {
                        "type": "paragraph",
                        "content": "Dear Recipient,",
                        "alignment": "left",
                        "font_name": "Calibri",
                        "font_size": 11,
                        "spacing_after": 12
                    },
                    {
                        "type": "paragraph",
                        "content": "[OPENING_PARAGRAPH]",
                        "alignment": "left",
                        "font_name": "Calibri",
                        "font_size": 11,
                        "spacing_after": 12
                    },
                    {
                        "type": "paragraph",
                        "content": "[BODY_PARAGRAPH_1]",
                        "alignment": "left",
                        "font_name": "Calibri",
                        "font_size": 11,
                        "spacing_after": 12
                    }
                ],
                "placeholders": ["[OPENING_PARAGRAPH]", "[BODY_PARAGRAPH_1]"]
            }
        }
    }), 200

# ============================================================================
# STARTUP
# ============================================================================

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
