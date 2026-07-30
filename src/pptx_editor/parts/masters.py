from pathlib import PurePosixPath
from typing import TYPE_CHECKING

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.parts.xml_part import XmlPart
from pptx_editor.parts.theme import Theme
from pptx_editor.properties.part import RequiredRelatedPartProperty
from pptx_editor.properties.xml_element import RequiredXmlElementProperty
from pptx_editor.xml_elements.color import ColorMap

if TYPE_CHECKING:
    from pptx_editor.xml_elements.masters import SlideLayoutIdList
    from pptx_editor.parts.slide_layout import SlideLayout

class AbstractMaster(XmlPart):
    """Base class for SlideMaster and NotesMaster."""
    default_shared: bool = True
    is_abstract = True

    theme = RequiredRelatedPartProperty(Theme, target_type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme')

    @property
    def color_map(self) -> dict[str, str]:
        """Get the color map for this Master."""
        return self._color_map.colors

    @property
    def _color_map(self) -> 'ColorMap':
        """Get the ColorMap associated with this Master."""
        color_map = self.get_element('clrMap', 'p')

        if color_map is None or not isinstance(color_map, ColorMap):
            raise PowerpointIntegrityError('The clrMap element in the master part is missing or not of the expected type.')

        return color_map

    def __init_subclass__(cls, **kwargs):
        cls.is_abstract = False
        super().__init_subclass__(**kwargs)

class NotesMaster(AbstractMaster):
    """Represents a NotesMaster part in a PowerPoint presentation."""
    default_content_type = PresentationML.NOTES_MASTER
    default_base_path = PurePosixPath('/ppt/notesMasters')
    default_part_name = 'notesMaster{i}'


class SlideMaster(AbstractMaster):
    """Represents a SlideMaster part in a PowerPoint presentation."""
    default_content_type = PresentationML.SLIDE_MASTER
    default_base_path = PurePosixPath('/ppt/slideMasters')
    default_part_name = 'slideMaster{i}'

    _slide_layout_id_list = RequiredXmlElementProperty['SlideLayoutIdList']('SlideLayoutIdList')

    @property
    def slide_layouts(self) -> list['SlideLayout']:
        """Get the list of SlideLayout parts associated with this SlideMaster."""
        return self._slide_layout_id_list.slide_layouts

    @slide_layouts.setter
    def slide_layouts(self, value: list['SlideLayout']) -> None:
        """Set the list of SlideLayout parts associated with this SlideMaster."""
        self._slide_layout_id_list.slide_layouts = value
