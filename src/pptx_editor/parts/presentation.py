from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.part import Part
from pptx_editor.attribute import Attribute

class Presentation(Part):
    content_type = PresentationML.PRESENTATION

class SlideSize(Attribute):
    pass