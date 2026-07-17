from pptx_editor.xml_element import XmlElement
from pptx_editor.xml_elements.color import AbstractColor
from pptx_editor.properties.xml_element import RequiredXmlElementProperty


class ColorScheme(XmlElement):
    """Represents the clrScheme element in a PowerPoint presentation."""
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'clrScheme'
    default_order = (
        'dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3',
        'accent4', 'accent5', 'accent6', 'hlink', 'folHlink', 'extLst',
    )

    @property
    def colors(self) -> dict[str, AbstractColor]:
        return {element.name: element.children[0] for element in self.children if element.children and isinstance(element.children[0], AbstractColor)}


class ThemeElements(XmlElement):
    """Represents the themeElements element in a PowerPoint presentation."""
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'themeElements'
    default_order = ('clrScheme', 'fontScheme', 'fmtScheme', 'extLst',)

    color_scheme = RequiredXmlElementProperty(ColorScheme)
