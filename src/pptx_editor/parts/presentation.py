from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.part import Part

class Presentation(Part):
    content_type = PresentationML.PRESENTATION

    def _parse_xml(self, file_xml):
        pass
