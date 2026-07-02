import logging
from flask import request, jsonify
from services import DriveService, TemplateFiller

logger = logging.getLogger(__name__)

def fill_template(app):
    """Register /fill-template endpoint"""
    
    @app.route('/fill-template', methods=['POST'])
    def fill_template_handler():
        """
        Fill a template with provided replacements and upload to Drive
        
        Request JSON:
        {
            "template_file_id": "Google Drive file ID",
            "replacements": {
                "[DATE]": "2 July 2026",
                "[BODY]": "Content here...",
                ...
            },
            "output_filename": "Filled-Document.docx",
            "folder_id": "Google Drive folder ID for upload"
        }
        
        Response JSON:
        {
            "file_id": "New file ID",
            "drive_link": "Shareable link",
            "file_name": "Output filename",
            "replacements_made": 5,
            "warnings": []
        }
        """
        try:
            data = request.get_json()
            if not data:
                return jsonify({"error": "No JSON provided"}), 400
            
            # Validate required fields
            template_file_id = data.get('template_file_id')
            replacements = data.get('replacements', {})
            output_filename = data.get('output_filename', 'filled-document.docx')
            folder_id = data.get('folder_id')
            
            if not template_file_id:
                return jsonify({"error": "template_file_id required"}), 400
            if not isinstance(replacements, dict):
                return jsonify({"error": "replacements must be a dictionary"}), 400
            if not output_filename:
                return jsonify({"error": "output_filename required"}), 400
            
            logger.info(f"Filling template {template_file_id} with {len(replacements)} replacements")
            
            # Initialize Drive service
            drive = DriveService()
            
            # Check folder exists if provided
            if folder_id and not drive.folder_exists(folder_id):
                return jsonify({"error": f"Folder {folder_id} not found or not accessible"}), 400
            
            # Download template
            docx_bytes = drive.download_file(template_file_id)
            
            # Fill template
            filled_bytes, fill_report = TemplateFiller.fill_template(docx_bytes, replacements)
            
            # Check for critical issues
            if fill_report['unmatched_placeholders']:
                logger.warning(f"Unmatched placeholders: {fill_report['unmatched_placeholders']}")
            
            # Upload filled document
            file_id, drive_link = drive.upload_file(
                filled_bytes,
                output_filename,
                folder_id,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
            
            response = {
                "file_id": file_id,
                "drive_link": drive_link,
                "file_name": output_filename,
                "replacements_made": fill_report['replacements_made'],
                "unmatched_placeholders": fill_report['unmatched_placeholders'],
                "extra_keys": fill_report['extra_keys'],
                "warnings": fill_report['warnings']
            }
            
            logger.info(f"Template filled and uploaded: {file_id}")
            return jsonify(response), 200
        
        except Exception as e:
            logger.error(f"Error filling template: {e}", exc_info=True)
            return jsonify({"error": str(e)}), 500
    
    return app
