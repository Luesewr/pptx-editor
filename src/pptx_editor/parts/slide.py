from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.part import Part

class Slide(Part):
    content_type = PresentationML.SLIDE

