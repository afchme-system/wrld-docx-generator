import os
import logging
from flask import request, jsonify
from services import DriveService, TemplateFiller

logger = logging.getLogger(__name__)

def fill_template(app):
    """Register /fill-template endpoint"""
    
    @app.route('/fill-template', methods=['POST'])
    def fill_template_handler():
        """
        Fill a template with provided replacements and upload to Drive.
        
        The template can be identified by either a direct file ID or by name.
        When using template_name, the service searches the configured templates
        folder (GOOGLE_DRIVE_TEMPLATES_FOLDER) for a matching file.
        
        Request JSON (option A — direct file ID):
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
        
        Request JSON (option B — template name):
        {
            "template_name": "WRLD-Template-Letter (1).docx",
            "replacements": { ... },
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
            
            # Resolve template identifier — accept either a direct file ID or a name
            template_file_id = data.get('template_file_id')
            template_name = data.get('template_name')
            replacements = data.get('replacements', {})
            output_filename = data.get('output_filename', 'filled-document.docx')
            folder_id = data.get('folder_id')
            
            if not template_file_id and not template_name:
                return jsonify({"error": "Either template_file_id or template_name is required"}), 400
            if not isinstance(replacements, dict):
                return jsonify({"error": "replacements must be a dictionary"}), 400
            if not output_filename:
                return jsonify({"error": "output_filename required"}), 400
            
            # Initialize Drive service
            drive = DriveService()
            
            # Resolve template_name -> file ID when a name is provided
            if template_name and not template_file_id:
                templates_folder = os.getenv('GOOGLE_DRIVE_TEMPLATES_FOLDER')
                logger.info(f"Resolving template by name: {template_name}")
                try:
                    template_file_id = drive.find_file_by_name(template_name, folder_id=templates_folder)
                except FileNotFoundError as e:
                    return jsonify({"error": str(e)}), 404
            
            logger.info(f"Filling template {template_file_id} with {len(replacements)} replacements")
            
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
