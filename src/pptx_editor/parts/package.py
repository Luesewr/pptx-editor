from pathlib import PurePosixPath

from pptx_editor.parts.xml_part import Part

class Package(Part):
    default_content_type = None
    default_base_path: PurePosixPath | None = PurePosixPath('/')
    default_part_name = None

    def __init__(self, *args, **kwargs):
        self._parts: list['Part'] = []
        super().__init__(self, *args, **kwargs)

    def _add_part(self, part: 'Part'):
        self._parts.append(part)
