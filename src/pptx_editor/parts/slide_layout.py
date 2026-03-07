from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.part import Part

class SlideLayout(Part):
    content_type = PresentationML.SLIDE_LAYOUT

    def _parse_xml(self, file_xml):
        pass
