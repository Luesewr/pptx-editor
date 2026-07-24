from pathlib import PurePosixPath
from typing import TYPE_CHECKING

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.parts.masters import SlideMaster
from pptx_editor.parts.xml_part import XmlPart

if TYPE_CHECKING:
    from pptx_editor.parts.theme import Theme

class SlideLayout(XmlPart):
    default_content_type = PresentationML.SLIDE_LAYOUT
    default_base_path = PurePosixPath('/ppt/slideLayouts')
    default_part_name = 'slideLayout{i}'
    default_shared: bool = True

    @property
    def slide_master(self) -> 'SlideMaster':
        """Get the SlideMaster associated with this SlideLayout."""
        slide_master_part = self.get_related_part(SlideMaster)

        if slide_master_part is None:
            raise PowerpointIntegrityError("SlideLayout does not have an associated SlideMaster.")

        return slide_master_part

    @property
    def theme(self) -> 'Theme':
        """Get the Theme associated with this SlideLayout, if it exists."""
        theme_part = self.slide_master.theme

        return theme_part
