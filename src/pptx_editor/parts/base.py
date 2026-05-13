from pathlib import PurePosixPath
from typing import TYPE_CHECKING

from pptx_editor.parts.xml_part import Part

class Base(Part):
    default_content_type = None
    default_base_path: PurePosixPath | None = PurePosixPath('/')
    default_part_name = None

    def __init__(self, *args, **kwargs):
        self.parts: list['Part'] = []
        super().__init__(self, *args, **kwargs)

    def add_part(self, part: 'Part'):
        self.parts.append(part)
