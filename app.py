import os
import io
import logging
from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DocumentBuilder:
    """Render JSON specification to .docx - no formatting decisions"""
    
    alignment_map = {
        'left': WD_ALIGN_PARAGRAPH.LEFT,
        'right': WD_ALIGN_PARAGRAPH.RIGHT,
        'center': WD_ALIGN_PARAGRAPH.CENTER,
        'justify': WD_ALIGN_PARAGRAPH.JUSTIFY
    }
    
    def build(self, spec):
        """Build document from JSON spec exactly as specified"""
        doc = Document()
        
        # Set margins if provided
        if 'margins' in spec:
            m = spec['margins']
            for section in doc.sections:
                section.top_margin = Inches(m.get('top', 1))
                section.bottom_margin = Inches(m.get('bottom', 1))
                section.left_margin = Inches(m.get('left', 1))
                section.right_margin = Inches(m.get('right', 1))
        
        # Render all components
        for comp in spec.get('components', []):
            self._render(doc, comp)
        
        # Save
        doc_bytes = io.BytesIO()
        doc.save(doc_bytes)
        doc_bytes.seek(0)
        return doc_bytes.getvalue()
    
    def _render(self, doc, comp):
        """Render component as specified"""
        comp_type = comp.get('type', 'paragraph')
        content = comp.get('content', '')
        
        if comp_type == 'page_break':
            doc.add_page_break()
        elif comp_type == 'heading':
            para = doc.add_heading(content, level=comp.get('level', 1))
        elif comp_type == 'list_bullet':
            para = doc.add_paragraph(content, style='List Bullet')
        else:
            para = doc.add_paragraph(content, style=comp.get('style', 'Normal'))
        
        if comp_type != 'page_break':
            self._format(para, comp)
    
    def _format(self, para, comp):
        """Apply formatting from spec"""
        # Alignment
        alignment = comp.get('alignment', 'left').lower()
        para.alignment = self.alignment_map.get(alignment, WD_ALIGN_PARAGRAPH.LEFT)
        
        # Spacing
        if 'spacing_before' in comp:
            para.paragraph_format.space_before = Pt(comp['spacing_before'])
        if 'spacing_after' in comp:
            para.paragraph_format.space_after = Pt(comp['spacing_after'])
        
        # Font
        font_name = comp.get('font_name', 'Calibri')
        font_size = comp.get('font_size', 11)
        bold = comp.get('bold', False)
        italic = comp.get('italic', False)
        
        # Color
        color_hex = comp.get('color', '#000000')
        try:
            r = int(color_hex[1:3], 16)
            g = int(color_hex[3:5], 16)
            b = int(color_hex[5:7], 16)
            color_rgb = RGBColor(r, g, b)
        except:
            color_rgb = RGBColor(0, 0, 0)
        
        # Apply to all runs
        for run in para.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.italic = italic
            run.font.color.rgb = color_rgb

builder = DocumentBuilder()

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"}), 200

@app.route('/create-template', methods=['POST'])
def create():
    """Accept JSON spec, render to .docx"""
    try:
        spec = request.get_json()
        if not spec:
            return jsonify({"error": "No JSON"}), 400
        
        filename = spec.get('output_filename', 'document.docx')
        if not filename:
            return jsonify({"error": "output_filename required"}), 400
        
        logger.info(f"Building: {filename}")
        docx_binary = builder.build(spec)
        
        return send_file(
            io.BytesIO(docx_binary),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        logger.error(str(e), exc_info=True)
        return jsonify({"error": str(e)}), 500

@app.route('/', methods=['GET'])
def root():
    return jsonify({"service": "WRLD Document Generator", "version": "2.2"}), 200

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
