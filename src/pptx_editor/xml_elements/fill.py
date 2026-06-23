from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.xml_element import XmlElement, XmlElementProperty
from pptx_editor.xml_elements.color import Color

class Fill(XmlElement):
    is_abstract = True

    def __init_subclass__(cls, **kwargs):
        cls.is_abstract = False
        super().__init_subclass__(**kwargs)


class NoFill(Fill):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'noFill'


class BlipFill(Fill):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'blipFill'


class GradientFill(Fill):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'gradFill'


class GroupFill(Fill):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'grpFill'


class PatternFill(Fill):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'pattFill'


class SolidFill(Fill):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'solidFill'

    color: Color = XmlElementProperty(Color, nullable=False)
