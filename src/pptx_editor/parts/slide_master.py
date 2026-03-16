from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.part import Part

class SlideMaster(Part):
    content_type = PresentationML.SLIDE_MASTER
    default_base_path = '/ppt/slideMasters'
    default_part_name = 'slideMaster{i}.xml'
