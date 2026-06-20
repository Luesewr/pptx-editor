from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.xml_element import XmlElement
from pptx_editor.xml_elements.color import Color

class Fill(XmlElement):
    is_abstract = True

    @property
    def color(self) -> 'Color | None':
        if not self.children:
            return None

        color_element = self.children[0]

        if not isinstance(color_element, Color):
            raise PowerpointIntegrityError('The element in the solidFill element is not of the expected type.')

        return color_element

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
