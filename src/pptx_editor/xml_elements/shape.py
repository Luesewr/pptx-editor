from pptx_editor.properties.xml_element import XmlElementProperty
from pptx_editor.xml_element import XmlElement
from pptx_editor.xml_elements.text import TextBody


class Shape(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'sp'

    text_body = XmlElementProperty(TextBody)


class NonVisualShapeProperties(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'nvSpPr'


class NonVisualGroupShapeProperties(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'nvGrpSpPr'


class GroupShapeProperties(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'grpSpPr'


class GroupShape(Shape):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'grpSp'


class Picture(Shape):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'pic'


class GraphicFrame(Shape):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'graphicFrame'


class ConnectionShape(Shape):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'cxnSp'
