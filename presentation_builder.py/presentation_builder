"""
presentation_builder.py
 
Spec-driven PowerPoint generation, mirroring DocumentBuilder's philosophy:
no hardcoded document-type logic, no hardcoded font/spacing enums baked
into the builder. Every visual choice comes from the incoming spec, with
neutral fallback defaults only when a field is omitted.
 
Supported slide types (first version, text-only per plan -- images/charts
deferred until a real example deck is reviewed):
    - title_slide
    - section_header
    - bullet_content
    - closing_slide
 
Expected spec shape:
{
    "design": {                     # optional, all fields optional
        "slide_width_in": 13.333,
        "slide_height_in": 7.5,
        "font_name": "Calibri",
        "title_font_size": 40,
        "subtitle_font_size": 22,
        "body_font_size": 20
    },
    "slides": [
        {"type": "title_slide", "title": "...", "subtitle": "..."},
        {"type": "section_header", "title": "...", "subtitle": "..."},
        {"type": "bullet_content", "title": "...", "bullets": [
            {"text": "...", "level": 0},
            {"text": "...", "level": 1}
        ]},
        {"type": "closing_slide", "title": "...", "message": "..."}
    ]
}
"""
 
import io
 
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
 
EMU_PER_INCH = 914400
 
DEFAULT_SLIDE_WIDTH_IN = 13.333
DEFAULT_SLIDE_HEIGHT_IN = 7.5
DEFAULT_FONT_NAME = "Calibri"
DEFAULT_TITLE_FONT_SIZE = 40
DEFAULT_SUBTITLE_FONT_SIZE = 22
DEFAULT_BODY_FONT_SIZE = 20
 
 
class PresentationBuilderError(Exception):
    """Raised when a spec is malformed or a slide cannot be built."""
 
 
class PresentationBuilder:
    """
    Builds a .pptx file entirely from spec['slides'] and spec['design'].
    Reads only these two keys -- no other spec-driven document type should
    ever need to touch this class, and this class should never need to
    know about document types outside presentations.
    """
 
    SLIDE_TYPES = {"title_slide", "section_header", "bullet_content", "closing_slide"}
 
    def __init__(self, spec: dict):
        if not isinstance(spec, dict):
            raise PresentationBuilderError("spec must be a dict")
 
        self.spec = spec
        self.design = spec.get("design") or {}
        self.slides_spec = spec.get("slides") or []
 
        if not self.slides_spec:
            raise PresentationBuilderError("spec['slides'] is required and cannot be empty")
 
        self.prs = Presentation()
        self.prs.slide_width = Inches(self.design.get("slide_width_in", DEFAULT_SLIDE_WIDTH_IN))
        self.prs.slide_height = Inches(self.design.get("slide_height_in", DEFAULT_SLIDE_HEIGHT_IN))
 
        self.font_name = self.design.get("font_name", DEFAULT_FONT_NAME)
        self.title_size = self.design.get("title_font_size", DEFAULT_TITLE_FONT_SIZE)
        self.subtitle_size = self.design.get("subtitle_font_size", DEFAULT_SUBTITLE_FONT_SIZE)
        self.body_size = self.design.get("body_font_size", DEFAULT_BODY_FONT_SIZE)
 
        # Layout index 6 is the blank layout in python-pptx's default template.
        self.blank_layout = self.prs.slide_layouts[6]
 
    # ------------------------------------------------------------------ #
    # Public API
    # ------------------------------------------------------------------ #
 
    def build(self) -> bytes:
        for i, slide_spec in enumerate(self.slides_spec):
            slide_type = slide_spec.get("type")
            if slide_type not in self.SLIDE_TYPES:
                raise PresentationBuilderError(
                    f"Slide {i}: unknown type '{slide_type}'. "
                    f"Must be one of {sorted(self.SLIDE_TYPES)}"
                )
            builder_method = getattr(self, f"_build_{slide_type}")
            builder_method(slide_spec)
 
        buffer = io.BytesIO()
        self.prs.save(buffer)
        return buffer.getvalue()
 
    # ------------------------------------------------------------------ #
    # Shared helpers
    # ------------------------------------------------------------------ #
 
    @property
    def _slide_w_in(self) -> float:
        return self.prs.slide_width / EMU_PER_INCH
 
    @property
    def _slide_h_in(self) -> float:
        return self.prs.slide_height / EMU_PER_INCH
 
    def _new_slide(self):
        return self.prs.slides.add_slide(self.blank_layout)
 
    def _add_textbox(self, slide, text, left_in, top_in, width_in, height_in,
                      font_size, bold=False, align=PP_ALIGN.LEFT, color_hex=None):
        box = slide.shapes.add_textbox(
            Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in)
        )
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text or ""
        p.alignment = align
 
        if p.runs:
            run = p.runs[0]
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.name = self.font_name
            if color_hex:
                run.font.color.rgb = RGBColor.from_string(color_hex)
        return box
 
    # ------------------------------------------------------------------ #
    # Slide type builders
    # ------------------------------------------------------------------ #
 
    def _build_title_slide(self, s):
        slide = self._new_slide()
        w, h = self._slide_w_in, self._slide_h_in
        title = s.get("title", "")
        subtitle = s.get("subtitle")
 
        self._add_textbox(
            slide, title, 0.75, h / 2 - 1.0, w - 1.5, 1.4,
            self.title_size, bold=True, align=PP_ALIGN.CENTER,
        )
        if subtitle:
            self._add_textbox(
                slide, subtitle, 0.75, h / 2 + 0.4, w - 1.5, 1.0,
                self.subtitle_size, align=PP_ALIGN.CENTER,
            )
 
    def _build_section_header(self, s):
        slide = self._new_slide()
        w, h = self._slide_w_in, self._slide_h_in
        title = s.get("title", "")
        subtitle = s.get("subtitle")
 
        self._add_textbox(
            slide, title, 0.75, h / 2 - 0.8, w - 1.5, 1.2,
            self.title_size, bold=True,
        )
        if subtitle:
            self._add_textbox(
                slide, subtitle, 0.75, h / 2 + 0.5, w - 1.5, 0.9,
                self.subtitle_size,
            )
 
    def _build_bullet_content(self, s):
        slide = self._new_slide()
        w, h = self._slide_w_in, self._slide_h_in
        title = s.get("title", "")
        bullets = s.get("bullets") or []
 
        if not bullets:
            raise PresentationBuilderError(
                "bullet_content slide requires a non-empty 'bullets' list"
            )
 
        self._add_textbox(slide, title, 0.6, 0.4, w - 1.2, 1.0, self.title_size, bold=True)
 
        box = slide.shapes.add_textbox(Inches(0.9), Inches(1.6), Inches(w - 1.6), Inches(h - 2.2))
        tf = box.text_frame
        tf.word_wrap = True
 
        for i, bullet in enumerate(bullets):
            if isinstance(bullet, dict):
                text = bullet.get("text", "")
                level = bullet.get("level", 0)
            else:
                text = str(bullet)
                level = 0
 
            level = min(max(int(level), 0), 4)
            marker = "•" if level == 0 else "–"
 
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"{marker}  {text}"
            p.level = level
 
            if p.runs:
                run = p.runs[0]
                run.font.size = Pt(max(self.body_size - (2 * level), 12))
                run.font.name = self.font_name
 
    def _build_closing_slide(self, s):
        slide = self._new_slide()
        w, h = self._slide_w_in, self._slide_h_in
        title = s.get("title", "")
        message = s.get("message") or s.get("subtitle")
 
        self._add_textbox(
            slide, title, 0.75, h / 2 - 1.0, w - 1.5, 1.2,
            self.title_size, bold=True, align=PP_ALIGN.CENTER,
        )
        if message:
            self._add_textbox(
                slide, message, 0.75, h / 2 + 0.3, w - 1.5, 1.0,
                self.subtitle_size, align=PP_ALIGN.CENTER,
            )
 
 
def build_presentation(spec: dict) -> bytes:
    """Convenience entry point mirroring DocumentBuilder's module-level function."""
    return PresentationBuilder(spec).build()
 
