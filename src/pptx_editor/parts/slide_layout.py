from pathlib import PurePosixPath
from typing import TYPE_CHECKING

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.parts.masters import SlideMaster
from pptx_editor.parts.xml_part import XmlPart
from pptx_editor.properties.part import RequiredRelatedPartProperty

if TYPE_CHECKING:
    from pptx_editor.parts.theme import Theme

class SlideLayout(XmlPart):
    default_content_type = PresentationML.SLIDE_LAYOUT
    default_base_path = PurePosixPath('/ppt/slideLayouts')
    default_part_name = 'slideLayout{i}'
    default_shared: bool = True

    slide_master = RequiredRelatedPartProperty(SlideMaster, target_type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster')

    @property
    def theme(self) -> 'Theme':
        """Get the Theme associated with this SlideLayout, if it exists."""
        theme_part = self.slide_master.theme

        return theme_part
