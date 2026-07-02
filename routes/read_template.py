import io
import logging
from flask import request, jsonify, send_file
from services import DriveService

logger = logging.getLogger(__name__)

def read_template(app):
    """Register /read-template endpoint"""
    
    @app.route('/read-template', methods=['POST'])
    def read_template_handler():
        """
        Download and return the raw binary content of a template from Google Drive.
        
        Request JSON:
        {
            "template_file_id": "Google Drive file ID"
        }
        
        Response: Binary .docx file
        (application/vnd.openxmlformats-officedocument.wordprocessingml.document)
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
            
            # Get file metadata to derive a sensible download filename
            file_metadata = drive.get_file_metadata(template_file_id)
            file_name = file_metadata.get('name', 'template.docx')
            
            # Download raw template bytes
            docx_bytes = drive.download_file(template_file_id)
            
            logger.info(f"Template read successful: {file_name} ({len(docx_bytes)} bytes)")
            return send_file(
                io.BytesIO(docx_bytes),
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                as_attachment=True,
                download_name=file_name
            )
        
        except Exception as e:
            logger.error(f"Error reading template: {e}", exc_info=True)
            return jsonify({"error": str(e)}), 500
    
    return app
