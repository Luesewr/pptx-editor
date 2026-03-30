from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.parts.xml_part import XmlPart

class Slide(XmlPart):
    default_content_type = PresentationML.SLIDE
    default_base_path = '/ppt/slides'
    default_part_name = 'slide{i}.xml'
