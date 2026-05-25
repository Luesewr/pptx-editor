from typing import TYPE_CHECKING, TypeGuard

from pptx_editor.xml_element import XmlElement
from pptx_editor.xml_elements.text import TextBody
from pptx_editor.exceptions import PowerpointIntegrityError

if TYPE_CHECKING:
    from pptx_editor.parts.slide import Slide

class Shape(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'sp'

    @property
    def text_body(self) -> 'TextBody | None':
        text_body = self.get_attribute('txBody', 'p')

        if text_body is not None and not isinstance(text_body, TextBody):
            raise PowerpointIntegrityError('The txBody element in the shape element is not of the expected type.')

        return text_body

class NonVisualShapeProperties(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'nvSpPr'

class NonVisualGroupShapeProperties(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'nvGrpSpPr'

class GroupShapeProperties(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'grpSpPr'

class GroupShape(Shape):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'grpSp'

class Picture(Shape):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'pic'

class GraphicFrame(Shape):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'graphicFrame'

class ConnectionShape(Shape):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'cxnSp'
