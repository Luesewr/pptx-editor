from pptx_editor.parts.slide_layout import SlideLayout
from pptx_editor.properties.id_list import PartIdListProperty
from pptx_editor.xml_elements.id_list import AbstractId, AbstractIdList


class SlideLayoutId(AbstractId):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'sldLayoutId'

class SlideLayoutIdList(AbstractIdList):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'sldLayoutIdLst'
    id_class = SlideLayoutId
    relationship_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'

    slide_layouts = PartIdListProperty(SlideLayout)
