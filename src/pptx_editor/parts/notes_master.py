from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.part import Part

class NotesMaster(Part):
    content_type = PresentationML.NOTES_MASTER

    def _parse_xml(self, file_xml):
        pass
