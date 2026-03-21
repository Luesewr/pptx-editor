from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.part import Part

class NotesMaster(Part):
    default_content_type = PresentationML.NOTES_MASTER
    default_base_path = '/ppt/notesMasters'
    default_part_name = 'notesMaster{i}.xml'
