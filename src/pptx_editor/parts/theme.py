from pathlib import PurePosixPath

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.parts.xml_part import XmlPart
from pptx_editor.xml_elements.color import Color
from pptx_editor.xml_elements.theme import ThemeElements


class Theme(XmlPart):
    default_content_type: str | None = PresentationML.THEME
    default_base_path: PurePosixPath | None = PurePosixPath("/ppt/theme")
    default_part_name: str | None = "theme{i}.xml"

    def get_color(self, scheme_color: str) -> 'Color':
        color_scheme_colors = self._theme_elements.color_scheme.colors

        if scheme_color not in color_scheme_colors:
            raise ValueError(f"Scheme color {scheme_color} not found in theme color scheme.")

        return color_scheme_colors[scheme_color]

    @property
    def _theme_elements(self) -> 'ThemeElements':
        """Get the theme elements associated with this Theme."""
        theme_elements = self.get_element('themeElements', 'a')

        if theme_elements is None:
            raise ValueError("Theme does not have an associated themeElements.")

        return theme_elements
