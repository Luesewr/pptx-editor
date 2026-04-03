from pathlib import PurePosixPath

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.parts.xml_part import XmlPart

class SlideLayout(XmlPart):
    default_content_type = PresentationML.SLIDE_LAYOUT
    default_base_path = PurePosixPath('/ppt/slideLayouts')
    default_part_name = 'slideLayout{i}.xml'
