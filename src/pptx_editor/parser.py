from zipfile import ZipFile
from typing import IO

from pptx_editor.content_types import ContentTypes
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.part import ReturnPart, Part, PartRegistry
from pptx_editor.relationship import Relationship

class Parser:
    def __init__(self, file: IO):
        self.zip_file = ZipFile(file)
        self.parts: dict[str, 'Part'] = {}
        self.content_types: ContentTypes | None = None

    def parse_zip_file(self, return_location: str = '/ppt/presentation.xml', return_type: type[ReturnPart] | None = None) -> Part:
        if return_location.lstrip('/') not in self.zip_file.namelist():
            raise PowerpointIntegrityError(f"Integrity warning: Main presentation part {return_location} not found in zip file")

        if return_type is None:
            from pptx_editor.parts.presentation import Presentation
            return_type = Presentation

        self.content_types = ContentTypes.from_file(self)

        main_relations_path = '_rels/.rels'
        main_relationships = Relationship.from_file(self, main_relations_path)
        self.parse_relationship_targets(main_relationships)

        powerpoint_part_path = '/ppt/presentation.xml'

        if not self.has_part(powerpoint_part_path):
            raise PowerpointIntegrityError(f"Integrity warning: Main presentation part {powerpoint_part_path} not found")

        powerpoint_part: Part | None = self.get_part(powerpoint_part_path)

        if powerpoint_part is None or not isinstance(powerpoint_part, return_type):
            raise PowerpointIntegrityError(f"Integrity warning: Main presentation part {powerpoint_part_path} has incorrect content type")

        powerpoint_part.main_relationships = main_relationships

        return powerpoint_part

    def parse_relationship_targets(self, relationships: list[Relationship]):
        for relationship in relationships:
            relationship_target_path = relationship.get_target_file_path()
            if not self.has_part(relationship_target_path):
                self.parse_part(relationship_target_path)

    def parse_part(self, file_path: str):
        content_type = self.get_content_type(file_path)

        if not content_type.endswith('xml'):
            print(f"Skipping non-XML part {file_path} with content type {content_type}")
            return None

        part_cls = PartRegistry().get_part_cls(content_type)
        part = part_cls.from_file(self, file_path, content_type)

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

    def add_part(self, part):
        self.parts[part.file_path] = part

    def has_part(self, file_path: str) -> bool:
        return file_path in self.parts

    def get_part(self, file_path: str):
        return self.parts.get(file_path)
