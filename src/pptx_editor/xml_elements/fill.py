from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.xml_element import XmlElement, XmlElementProperty
from pptx_editor.xml_elements.color import AbstractColor

class AbstractFill(XmlElement):
    is_abstract = True

    def __init_subclass__(cls, **kwargs):
        cls.is_abstract = False
        super().__init_subclass__(**kwargs)


class NoFill(AbstractFill):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'noFill'


class BlipFill(AbstractFill):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'blipFill'


class GradientFill(AbstractFill):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'gradFill'


class GroupFill(AbstractFill):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'grpFill'


class PatternFill(AbstractFill):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'pattFill'


class SolidFill(AbstractFill):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'solidFill'

    color: AbstractColor = XmlElementProperty(AbstractColor, nullable=False)
