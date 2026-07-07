import logging
from flask import request, jsonify, send_file
import io
from services import DriveService, TemplateFiller

logger = logging.getLogger(__name__)

def fill_template(app):
    """Register /fill-template endpoint"""
    
    @app.route('/fill-template', methods=['POST'])
    def fill_template_handler():
        """
        Fill a template with provided replacements and return the filled document.
        Does NOT upload to Drive — that is handled downstream by n8n's own
        Google Drive node, consistent with the /create-template pattern.
        
        Request JSON:
        {
            "template_file_id": "Google Drive file ID",
            "replacements": {
                "[DATE]": "2 July 2026",
                "[BODY]": "Content here...",
                ...
            },
            "output_filename": "Filled-Document.docx"
        }
        
        Response: raw .docx binary, with Content-Disposition set to output_filename
        """
        try:
            data = request.get_json()
            if not data:
                return jsonify({"error": "No JSON provided"}), 400
            
            template_file_id = data.get('template_file_id')
            replacements = data.get('replacements', {})
            output_filename = data.get('output_filename', 'filled-document.docx')
            
            if not template_file_id:
                return jsonify({"error": "template_file_id required"}), 400
            if not isinstance(replacements, dict):
                return jsonify({"error": "replacements must be a dictionary"}), 400
            if not output_filename:
                return jsonify({"error": "output_filename required"}), 400
            
            logger.info(f"Filling template {template_file_id} with {len(replacements)} replacements")
            
            drive = DriveService()
            docx_bytes = drive.download_file(template_file_id)
            
            filled_bytes, fill_report = TemplateFiller.fill_template(docx_bytes, replacements)
            
            if fill_report['unmatched_placeholders']:
                logger.warning(f"Unmatched placeholders: {fill_report['unmatched_placeholders']}")
            
            logger.info(f"Template filled successfully: {output_filename} ({fill_report['replacements_made']} replacements)")
            
            return send_file(
                io.BytesIO(filled_bytes),
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                as_attachment=True,
                download_name=output_filename
            )
        
        except Exception as e:
            logger.error(f"Error filling template: {e}", exc_info=True)
            return jsonify({"error": str(e)}), 500
    
    return app
