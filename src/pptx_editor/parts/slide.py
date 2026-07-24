from pathlib import PurePosixPath

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.parts.slide_layout import SlideLayout
from pptx_editor.parts.theme import Theme
from pptx_editor.parts.xml_part import XmlPart
from pptx_editor.xml_elements.color import ColorMapOverride, MasterColorMapping, OverrideColorMapping
from pptx_editor.xml_elements.shape import AbstractShape
from pptx_editor.xml_elements.slide_data import CommonSlideData


class Slide(XmlPart):
    default_content_type = PresentationML.SLIDE
    default_base_path = PurePosixPath('/ppt/slides')
    default_part_name = 'slide{i}'

    @property
    def shapes(self) -> list['AbstractShape']:
        return self._common_slide_data.shapes

    @property
    def slide_layout(self) -> 'SlideLayout':
        """Get the SlideLayout associated with this Slide."""
        slide_layout_part = self.get_related_part(SlideLayout)

        if slide_layout_part is None:
            raise PowerpointIntegrityError("Slide does not have an associated SlideLayout.")

        return slide_layout_part

    @property
    def theme(self) -> 'Theme':
        """Get the Theme associated with this Slide."""
        return self.slide_layout.theme

    @property
    def color_map(self) -> dict[str, str]:
        """Get the color map for this Slide, taking into account any ColorMapOverride."""
        color_map_override = self._color_map_override

        override_color_map = color_map_override.children[0]

        if isinstance(override_color_map, MasterColorMapping):
            slide_master = self.slide_layout.slide_master
            return slide_master.color_map

        if isinstance(override_color_map, OverrideColorMapping):
            return override_color_map.colors

        raise PowerpointIntegrityError("Slide does not have a valid ColorMapOverride or MasterColorMapping.")

    @property
    def _color_map_override(self) -> 'ColorMapOverride':
        """Get the ColorMapOverride associated with this Slide, if it exists."""
        color_map_override = self.get_element('clrMapOvr', 'p')

        if color_map_override is not None and not isinstance(color_map_override, ColorMapOverride):
            raise PowerpointIntegrityError('The clrMapOvr element in the slide part is not of the expected type.')

        return color_map_override

    @property
    def _common_slide_data(self) -> 'CommonSlideData':
        slide_id_list = self.get_element('cSld', 'p')

        if slide_id_list is None:
            raise PowerpointIntegrityError('The presentation part is missing the required cSld element.')

        if slide_id_list is not None and not isinstance(slide_id_list, CommonSlideData):
            raise PowerpointIntegrityError('The cSld element in the presentation part is not of the expected type.')

        return slide_id_list
