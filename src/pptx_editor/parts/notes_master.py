from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.parts.xml_part import XmlPart

class NotesMaster(XmlPart):
    default_content_type = PresentationML.NOTES_MASTER
    default_base_path = '/ppt/notesMasters'
    default_part_name = 'notesMaster{i}.xml'
