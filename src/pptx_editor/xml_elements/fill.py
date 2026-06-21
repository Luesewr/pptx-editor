from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.xml_element import XmlElement
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

    @property
    def color(self) -> 'Color | None':
        if not self.children:
            return None

        color_elements = [child for child in self.children if isinstance(child, Color)]

        if len(color_elements) > 1:
            raise PowerpointIntegrityError('The solidFill element has multiple child elements of type Color, which is not allowed.')

        if not color_elements:
            return None

        color_element = color_elements[0]

        if not isinstance(color_element, Color):
            raise PowerpointIntegrityError('The element in the solidFill element is not of the expected type.')

        return color_element

    @color.setter
    def color(self, value: 'Color | None') -> None:
        if value is not None and value.part is not self.part:
            value = value.copy()  # Create a copy of the Color element in case user tries to assign a Color element to multiple Fill elements
            value.part = self.part  # Ensure the Color element is associated with the same part as the Fill element

        current_color = self.color

        if current_color is not None:
            if value is not None:
                self.replace_element(current_color, value)
            else:
                self.remove_element(current_color)
        else:
            if value is not None:
                self.add_element(value)
