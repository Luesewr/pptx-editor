from pathlib import PurePosixPath
from typing import TYPE_CHECKING

from pptx_editor.parts.xml_part import XmlPart

if TYPE_CHECKING:
    from pptx_editor.part import Part

class Base(XmlPart):
    default_content_type = None
    default_base_path: PurePosixPath | None = PurePosixPath('/')
    default_part_name = None

    def __init__(self, base: 'Base | None', file_path: PurePosixPath | None = None, content_type: str | None = None):
        self.parts: list['Part'] = []
        super().__init__(base, file_path, content_type)

    def add_part(self, part: 'Part'):
        self.parts.append(part)
