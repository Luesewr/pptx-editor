from pathlib import PurePosixPath

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.parts.xml_part import XmlPart
from pptx_editor.parts.theme import Theme
from pptx_editor.xml_elements.color import ColorMap

class Master(XmlPart):
    """Base class for SlideMaster and NotesMaster."""
    is_abstract = True

    @property
    def theme(self) -> 'Theme':
        """Get the Theme associated with this Master."""
        theme_part = self.get_related_part(Theme)

        if theme_part is None:
            raise PowerpointIntegrityError("Master does not have an associated Theme.")

        return theme_part

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

class NotesMaster(Master):
    """Represents a NotesMaster part in a PowerPoint presentation."""
    default_content_type = PresentationML.NOTES_MASTER
    default_base_path = PurePosixPath('/ppt/notesMasters')
    default_part_name = 'notesMaster{i}.xml'

class SlideMaster(Master):
    """Represents a SlideMaster part in a PowerPoint presentation."""
    default_content_type = PresentationML.SLIDE_MASTER
    default_base_path = PurePosixPath('/ppt/slideMasters')
    default_part_name = 'slideMaster{i}.xml'
