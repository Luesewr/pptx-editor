from collections import defaultdict
from zipfile import ZipFile
from typing import TYPE_CHECKING, IO

from pptx_editor.content_types import ContentTypes
from pptx_editor.relationship import Relationship

if TYPE_CHECKING:
    from pptx_editor.part import Part
    from pptx_editor.parts.presentation import Presentation

class Writer:
    def __init__(self, zip_file: ZipFile):
        self.zip_file = zip_file
        self.content_types = ContentTypes()
        self.relationship_id_lookup = defaultdict(dict)
        self.reverse_relationship_id_lookup = defaultdict(dict)
        self.part_index_lookup = defaultdict(dict)
        self.reverse_part_index_lookup = defaultdict(dict)
        self.written_parts = set()

    def write_to_buffer(self, presentation: 'Presentation'):
        base = presentation.base

        base.to_file(self)

    def write_file(self, file_path: str, content: bytes):
        self.zip_file.writestr(file_path.lstrip('/'), content)

    def assign_relationship_ids(self, part: 'Part', relationships: list['Relationship']):
        for relationship in relationships:
            self.assign_relationship_id(part, relationship)

    def assign_relationship_id(self, part: 'Part', relationship: 'Relationship') -> str:
        if relationship in self.reverse_relationship_id_lookup[part]:
            return self.reverse_relationship_id_lookup[part][relationship]

        relationship_id = f"rId{len(self.relationship_id_lookup[part]) + 1}"
        self.relationship_id_lookup[part][relationship_id] = relationship
        self.reverse_relationship_id_lookup[part][relationship] = relationship_id
        return relationship_id

    def assign_part_indexes(self, relationships: list['Relationship']):
        for relationship in relationships:
            target_part = relationship.target
            part_name = target_part.part_name if target_part.part_name else target_part.default_part_name

            self.assign_part_index(part_name, target_part)

    def assign_part_index(self, part_name: str | None, part: 'Part') -> str | None:
        if part_name is None or '{i}' not in part_name:
            return part_name

        if part in self.part_index_lookup[part_name]:
            return self.part_index_lookup[part_name][part]

        index = len(self.part_index_lookup[part_name]) + 1
        indexed_part_name = part_name.format(i=index)

        self.part_index_lookup[part_name][part] = indexed_part_name
        self.reverse_part_index_lookup[part_name][indexed_part_name] = part

        return indexed_part_name

    def add_written_part(self, part: 'Part'):
        self.written_parts.add(part)

    def is_part_written(self, part: 'Part') -> bool:
        return part in self.written_parts

    def has_part_index(self, part_name: str | None, part: 'Part') -> bool:
        if part_name is None:
            return False

        return part in self.part_index_lookup[part_name]

    def get_relationship_id(self, part: 'Part', relationship: 'Relationship') -> str | None:
        return self.reverse_relationship_id_lookup[part].get(relationship)

    def get_part_index(self, part_name: str | None, part: 'Part') -> str | None:
        if part_name is None:
            return None

        return self.part_index_lookup[part_name].get(part)
