from pathlib import PurePosixPath

from pptx_editor.parts.xml_part import Part

class Package(Part):
    default_content_type = None
    default_base_path: PurePosixPath | None = PurePosixPath('/')
    default_part_name = None
