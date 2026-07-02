import logging
from flask import request, jsonify
from services import DriveService, TemplateParser

logger = logging.getLogger(__name__)

def read_template(app):
    """Register /read-template endpoint"""
    
    @app.route('/read-template', methods=['POST'])
    def read_template_handler():
        """
        Read template metadata and placeholders from Google Drive
        
        Request JSON:
        {
            "template_file_id": "Google Drive file ID"
        }
        
        Response JSON:
        {
            "file_id": "...",
            "file_name": "...",
            "placeholders": ["[DATE]", "[BODY]", ...],
            "document_type": "formal_letter",
            "sections": ["header", "recipient_block", "body"],
            "metadata": { ... }
        }
        """
        try:
            data = request.get_json()
            if not data:
                return jsonify({"error": "No JSON provided"}), 400
            
            template_file_id = data.get('template_file_id')
            if not template_file_id:
                return jsonify({"error": "template_file_id required"}), 400
            
            logger.info(f"Reading template: {template_file_id}")
            
            # Initialize Drive service
            drive = DriveService()
            
            # Get file metadata
            file_metadata = drive.get_file_metadata(template_file_id)
            file_name = file_metadata.get('name', 'template.docx')
            
            # Download template
            docx_bytes = drive.download_file(template_file_id)
            
            # Parse template
            placeholders = TemplateParser.extract_placeholders(docx_bytes)
            full_metadata = TemplateParser.get_metadata(docx_bytes, file_name)
            
            response = {
                "file_id": template_file_id,
                "file_name": file_name,
                "placeholders": placeholders,
                "document_type": full_metadata.get('subject', 'formal_letter'),
                "sections": full_metadata.get('structure', {}).get('sections', []),
                "metadata": {
                    "page_count": full_metadata.get('structure', {}).get('page_count', 1),
                    "has_letterhead": full_metadata.get('structure', {}).get('has_letterhead', False),
                    "has_table": full_metadata.get('structure', {}).get('table_count', 0) > 0
                }
            }
            
            logger.info(f"Template read successful: {len(placeholders)} placeholders")
            return jsonify(response), 200
        
        except Exception as e:
            logger.error(f"Error reading template: {e}", exc_info=True)
            return jsonify({"error": str(e)}), 500
    
    return app
