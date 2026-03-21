from collections import defaultdict
from zipfile import ZipFile
from typing import IO

from pptx_editor.content_types import ContentTypes
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.part import Part, PartRegistry
from pptx_editor.relationship import Relationship

class Parser:
    def __init__(self, file: IO):
        self.zip_file = ZipFile(file)
        self.parts: dict[str, 'Part'] = {}
        self.relationships: dict[str, dict[str, Relationship]] = defaultdict(dict)
        self.content_types: ContentTypes | None = None
        self.base: Part | None = None

    def parse_zip_file(self, return_location: str = '/ppt/presentation.xml', return_type: type[Part] | None = None) -> Part:
        if return_location.lstrip('/') not in self.zip_file.namelist():
            raise PowerpointIntegrityError(f"Integrity warning: Main presentation part {return_location} not found in zip file")

        if return_type is None:
            from pptx_editor.parts.presentation import Presentation
            return_type = Presentation

        self.content_types = ContentTypes.from_file(self)

        from pptx_editor.parts.base import Base

        Base.from_file(None, self, None, None)

        powerpoint_part_path = '/ppt/presentation.xml'

        if not self.has_part(powerpoint_part_path):
            raise PowerpointIntegrityError(f"Integrity warning: Main presentation part {powerpoint_part_path} not found")

        powerpoint_part: Part | None = self.get_part(powerpoint_part_path)

        if powerpoint_part is None or not isinstance(powerpoint_part, return_type):
            raise PowerpointIntegrityError(f"Integrity warning: Main presentation part {powerpoint_part_path} has incorrect content type")

        return powerpoint_part

    def parse_part(self, file_path: str):
        content_type = self.get_content_type(file_path)

        if not content_type.endswith('xml'):
            print(f"Skipping non-XML part {file_path} with content type {content_type}")
            return None

        part_cls = PartRegistry().get_part_cls(content_type)
        part = part_cls.from_file(self.base, self, file_path, content_type)

        return part

    def get_content_type(self, file_path: str) -> str:
        if self.content_types is None:
            raise PowerpointIntegrityError("Content types not loaded")

        content_type = self.content_types.get_content_type(file_path)

        if content_type is None:
            raise PowerpointIntegrityError(f"Integrity warning: No content type found for file {file_path}")

        return content_type

    def read_file(self, file_path: str) -> bytes:
        return self.zip_file.read(file_path)

    def add_part(self, file_path: str | None, part: 'Part'):
        if file_path:
            self.parts[file_path] = part

    def has_part(self, file_path: str) -> bool:
        return file_path in self.parts

    def add_relationship(self, file_path: str | None, relationship_id: str | None, relationship: 'Relationship'):
        if file_path and relationship_id:
            self.relationships[file_path][relationship_id] = relationship

    def has_relationship(self, file_path: str, relationship_id: str) -> bool:
        return relationship_id in self.relationships[file_path]

    def get_part(self, file_path: str):
        return self.parts.get(file_path)

    def get_relationship(self, file_path: str, relationship_id: str) -> 'Relationship | None':
        return self.relationships[file_path].get(relationship_id)
