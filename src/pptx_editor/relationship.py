from io import BytesIO
import posixpath
import re
import sys

from functools import cmp_to_key
from typing import TYPE_CHECKING
from pathlib import PurePosixPath

from lxml import etree


if TYPE_CHECKING:
    from pptx_editor.parser import _OOXMLParser
    from pptx_editor.part import Part
    from pptx_editor.writer import _OOXMLWriter

class Relationship:
    __slots__ = ['target_type', 'target', 'origin']

    def __init__(self, target_type: str, target: 'Part', origin: 'Part'):
        self.target_type = sys.intern(target_type)
        self.target = target
        self.origin = origin

    @classmethod
    def _from_file(cls, parser: '_OOXMLParser', file_path: PurePosixPath, origin: 'Part'):
        relationship_xml = parser.read_file(file_path)
        relationship_tree = etree.fromstring(relationship_xml)
        part_file_path = cls._get_original_file_path(file_path)

        relationship_elements = []

        for relationship_element in relationship_tree:
            relationship_id = relationship_element.get('Id')
            relationship_target = relationship_element.get('Target')

            if relationship_id is None or relationship_target is None:
                print("Integrity warning: Relationship element missing Id or Target attribute")
                continue

            relationship = cls._from_xml(parser, relationship_element, file_path, origin)

            if relationship is None:
                continue

            parser.add_relationship(part_file_path, relationship_id, relationship)
            relationship_elements.append((PurePosixPath(relationship_target), relationship))

        sorted_elements = sorted(relationship_elements, key=cmp_to_key(cls._file_comparator))
        relationships = [relationship for _, relationship in sorted_elements]

        return relationships

    @classmethod
    def _from_xml(cls, parser: '_OOXMLParser', relationship_xml: etree._Element, file_path: PurePosixPath, origin: 'Part'):
        relationship_id = relationship_xml.get('Id')
        target_type = relationship_xml.get('Type')
        raw_target_path = relationship_xml.get('Target')
        target_mode = relationship_xml.get('TargetMode')

        if relationship_id is None or target_type is None or raw_target_path is None:
            print("Integrity warning: Relationship element missing required attributes")
            return None

        if target_mode is not None and target_mode.lower() == 'external':
            return ExternalRelationship(target_type, raw_target_path, origin)

        target_path = cls._get_target_file_path(file_path, PurePosixPath(raw_target_path))
        target = parser.parse_part_from_file(target_path)

        if target is None:
            print(f"Integrity warning: Relationship target {target_path} could not be parsed")
            return None

        return cls(target_type, target, origin)

    def is_external(self) -> bool:
        return isinstance(self, ExternalRelationship)

    def _to_xml(self, writer: '_OOXMLWriter', buffer: BytesIO):
        target_file_name = writer.assign_part_index(self.target.part_name, self.target)
        target_location = PurePosixPath(self.target.base_path) / target_file_name if self.target.base_path else PurePosixPath(target_file_name)
        relative_target_path = posixpath.relpath(target_location.as_posix(), start=(self.origin.base_path or PurePosixPath('/')).as_posix())
        relationship_id = writer.assign_relationship_id(self.origin, self)

        buffer.write(f'<Relationship Id="{relationship_id}" Type="{self.target_type}" Target="{relative_target_path}"/>'.encode('utf-8'))

    @staticmethod
    def _get_target_file_path(location: PurePosixPath, target: PurePosixPath) -> PurePosixPath:
        base_path = location.parent.parent
        current_target = PurePosixPath(target)

        path = PurePosixPath(posixpath.normpath(base_path / current_target))

        if not path.is_absolute():
            path = PurePosixPath('/') / path

        return path

    @staticmethod
    def _get_original_file_path(file_path: PurePosixPath) -> PurePosixPath:
        file_name = file_path.name.removesuffix('.rels')

        part_file_path = file_path.parent.parent / file_name

        if not part_file_path.is_absolute():
            part_file_path = PurePosixPath('/') / part_file_path

        return part_file_path

    @staticmethod
    def _file_comparator(a: tuple[PurePosixPath, 'Relationship'], b: tuple[PurePosixPath, 'Relationship']) -> int:
        a_name = a[0].name
        b_name = b[0].name
        a_index = re.match(r'(.*?)(\d+)\.[^\.]+$', a_name)
        b_index = re.match(r'(.*?)(\d+)\.[^\.]+$', b_name)

        if a_index and b_index and a_index.group(1) == b_index.group(1):
            return int(a_index.group(2)) - int(b_index.group(2))

        return 1 if a_name > b_name else (-1 if a_name < b_name else 0)

class ExternalRelationship(Relationship):
    __slots__ = ['target_type', 'target', 'origin']

    def __init__(self, target_type: str, target: str, origin: 'Part'):
        super().__init__(target_type, target, origin)

    def _to_xml(self, writer: '_OOXMLWriter', buffer: BytesIO):
        relationship_id = writer.assign_relationship_id(self.origin, self)

        buffer.write(f'<Relationship Id="{relationship_id}" Type="{self.target_type}" Target="{self.target}" TargetMode="External"/>'.encode('utf-8'))
