from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.part import Part

class SlideLayout(Part):
    default_content_type = PresentationML.SLIDE_LAYOUT
    default_base_path = '/ppt/slideLayouts'
    default_part_name = 'slideLayout{i}.xml'
