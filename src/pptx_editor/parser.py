from collections import defaultdict
from pathlib import PurePosixPath
from zipfile import ZipFile
from typing import IO, TYPE_CHECKING

from pptx_editor.content_types import ContentTypes
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.part import Part, PartRegistry
from pptx_editor.relationship import Relationship

if TYPE_CHECKING:
    from pptx_editor.parts.base import Base
    from pptx_editor.parts.presentation import Presentation

class Parser:
    def __init__(self, file: IO):
        self.zip_file = ZipFile(file)
        self.parts: dict[PurePosixPath, 'Part'] = {}
        self.relationships: dict[PurePosixPath, dict[str, Relationship]] = defaultdict(dict)
        self.content_types: ContentTypes | None = None
        self.base: 'Base | None' = None

    def parse_zip_file(self, return_location: str = '/ppt/presentation.xml', return_type: type[Presentation] | None = None) -> 'Presentation':
        if return_location.lstrip('/') not in self.zip_file.namelist():
            raise PowerpointIntegrityError(f"Integrity warning: Main presentation part {return_location} not found in zip file")

        from pptx_editor.parts.presentation import Presentation

        if return_type is None:
            return_type = Presentation

        self.content_types = ContentTypes.from_file(self)

        from pptx_editor.parts.base import Base

        Base.from_file(None, self, None, None)

        powerpoint_part_path = PurePosixPath('/ppt/presentation.xml')

        if not self.has_part(powerpoint_part_path):
            raise PowerpointIntegrityError(f"Integrity warning: Main presentation part {powerpoint_part_path} not found")

        powerpoint_part = self.get_part(powerpoint_part_path)

        if powerpoint_part is None or not isinstance(powerpoint_part, return_type):
            raise PowerpointIntegrityError(f"Integrity warning: Main presentation part {powerpoint_part_path} has incorrect content type")

        return powerpoint_part

    def parse_part(self, file_path: PurePosixPath):
        content_type, is_default = self.get_content_type(file_path)

        part_cls = PartRegistry().get_part_cls(content_type)
        part = part_cls.from_file(self.base, self, file_path, content_type, is_default)

        return part

    def get_content_type(self, file_path: PurePosixPath) -> tuple[str, bool]:
        if self.content_types is None:
            raise PowerpointIntegrityError("Content types not loaded")

        override_content_type = self.content_types.get_override_content_type(file_path)

        if override_content_type is not None:
            return override_content_type, False

        default_content_type = self.content_types.get_default_content_type(file_path)

        if default_content_type is not None:
            return default_content_type, True

        raise PowerpointIntegrityError(f"Integrity warning: No content type found for file {file_path}")

    def read_file(self, file_path: PurePosixPath) -> bytes:
        if file_path.is_absolute():
            file_path = file_path.relative_to(file_path.anchor)

        return self.zip_file.read(file_path.as_posix())

    def add_part(self, file_path: PurePosixPath | None, part: 'Part'):
        if file_path:
            self.parts[file_path] = part

    def has_part(self, file_path: PurePosixPath) -> bool:
        return file_path in self.parts

    def add_relationship(self, file_path: PurePosixPath | None, relationship_id: str | None, relationship: 'Relationship'):
        if file_path and relationship_id:
            self.relationships[file_path][relationship_id] = relationship

    def has_relationship(self, file_path: PurePosixPath, relationship_id: str) -> bool:
        return relationship_id in self.relationships[file_path]

    def get_part(self, file_path: PurePosixPath):
        return self.parts.get(file_path)

    def get_relationship(self, file_path: PurePosixPath, relationship_id: str) -> 'Relationship | None':
        return self.relationships[file_path].get(relationship_id)
