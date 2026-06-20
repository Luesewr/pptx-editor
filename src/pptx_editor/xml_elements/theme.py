from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.xml_element import XmlElement
from pptx_editor.xml_elements.color import Color


class ThemeElements(XmlElement):
    """Represents the themeElements element in a PowerPoint presentation."""
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'themeElements'

    @property
    def color_scheme(self) -> 'ColorScheme':
        color_scheme = self.get_element('clrScheme', 'a')

        if color_scheme is None:
            raise PowerpointIntegrityError("ThemeElements does not have an associated color scheme.")

        return color_scheme


class ColorScheme(XmlElement):
    """Represents the clrScheme element in a PowerPoint presentation."""
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'clrScheme'

    @property
    def colors(self) -> dict[str, Color]:
        return {element.name: element.children[0] for element in self.children if element.children and isinstance(element.children[0], Color)}
