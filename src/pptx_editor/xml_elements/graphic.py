from pptx_editor.xml_element import XmlElement
from pptx_editor.xml_elements.table import Table
from pptx_editor.properties.xml_element import RequiredXmlElementProperty, XmlElementProperty


class GraphicData(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'graphicData'

    table = XmlElementProperty(Table)


class Graphic(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'graphic'
    default_order = ('graphicData',)

    graphic_data = RequiredXmlElementProperty(GraphicData)
