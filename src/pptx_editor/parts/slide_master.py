from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.parts.xml_part import XmlPart

class SlideMaster(XmlPart):
    default_content_type = PresentationML.SLIDE_MASTER
    default_base_path = '/ppt/slideMasters'
    default_part_name = 'slideMaster{i}.xml'
