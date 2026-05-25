from typing import TYPE_CHECKING, TypeGuard

from pptx_editor.attribute import Attribute
from pptx_editor.exceptions import PowerpointIntegrityError

if TYPE_CHECKING:
    from pptx_editor.parts.slide import Slide

class Shape(Attribute):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'sp'

class NonVisualGroupShapeProperties(Attribute):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'nvGrpSpPr'

class GroupShapeProperties(Attribute):
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
