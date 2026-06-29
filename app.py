import sys
import os

# Ensure utils is importable
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, request, send_file, jsonify
from functools import wraps
import logging
import io
from utils.template_filler import TemplateFiller
from utils.template_creator import TemplateCreator
from utils.validators import validate_request

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize template filler and creator
filler = TemplateFiller(
    drive_folder_id=os.getenv('GOOGLE_DRIVE_TEMPLATES_FOLDER'),
    google_credentials_json=os.getenv('GOOGLE_CREDENTIALS_JSON')
)

creator = TemplateCreator()

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
        "metadata_fields": {...},
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

@app.route('/fill-template', methods=['POST'])
@error_handler
def fill_template():
    """
    POST /fill-template
    
    Request body:
    {
        "template_name": "WRLD-Template-Letter (1).docx",
        "content": {
            "date": "22 June 2026",
            "recipient_full_name": "John Smith",
            ...
        },
        "output_filename": "Letter_AFC_2026.docx"
    }
    
    Response: Binary .docx file
    """
    
    # Validate request
    data = request.get_json()
    if not data:
        raise ValueError("Request body must be JSON")
    
    template_name = data.get('template_name')
    content = data.get('content', {})
    output_filename = data.get('output_filename')
    
    # Validate required fields
    validate_request(template_name, content, output_filename)
    
    logger.info(f"Processing: {template_name} → {output_filename}")
    
    # Fill template
    docx_binary = filler.fill_template(
        template_name=template_name,
        content=content
    )
    
    # Return as downloadable file
    return send_file(
        io.BytesIO(docx_binary),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name=output_filename
    )

@app.route('/list-templates', methods=['GET'])
@error_handler
def list_templates():
    """
    GET /list-templates
    
    Returns available templates in the Drive folder
    """
    templates = filler.list_available_templates()
    return jsonify({
        "status": "ok",
        "count": len(templates),
        "templates": templates
    }), 200

@app.route('/', methods=['GET'])
def root():
    """Root endpoint with API documentation"""
    return jsonify({
        "service": "WRLD Document Generator",
        "version": "1.1",
        "endpoints": {
            "POST /create-template": "Create a new Word document from JSON specification",
            "POST /fill-template": "Fill a Word template with content JSON and return .docx",
            "GET /list-templates": "List available templates in Drive folder",
            "GET /health": "Health check"
        },
        "examples": {
            "create_template": {
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
            },
            "fill_template": {
                "method": "POST",
                "endpoint": "/fill-template",
                "body": {
                    "template_name": "WRLD-Template-Letter.docx",
                    "content": {"date": "22 June 2026"},
                    "output_filename": "Letter_AFC_2026.docx"
                }
            }
        }
    }), 200

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
