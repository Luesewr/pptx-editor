import posixpath
import sys

from typing import TYPE_CHECKING
from pathlib import PurePosixPath

from lxml import etree


if TYPE_CHECKING:
    from pptx_editor.parser import Parser
    from pptx_editor.part import Part
    from pptx_editor.writer import Writer
    from pptx_editor.parts.xml_part import XmlPart

class Relationship:
    __slots__ = ['target_type', 'target', 'origin']

    def __init__(self, target_type: str, target: 'Part', origin: 'XmlPart'):
        self.target_type = sys.intern(target_type)
        self.target = target
        self.origin = origin

    @classmethod
    def from_file(cls, parser: 'Parser', file_path: PurePosixPath, origin: 'XmlPart'):
        relationship_xml = parser.read_file(file_path)
        relationship_tree = etree.fromstring(relationship_xml)
        part_file_path = cls._get_original_file_path(file_path)

        relationships = []
        for relationship_element in relationship_tree:
            relationship_id = relationship_element.get('Id')
            if relationship_id is None:
                print("Integrity warning: Relationship element missing Id attribute")
                continue

            if parser.has_relationship(part_file_path, relationship_id):
                relationship = parser.get_relationship(part_file_path, relationship_id)
            else:
                relationship = cls.from_xml(parser, relationship_element, file_path, origin)
                if relationship is not None:
                    parser.add_relationship(part_file_path, relationship_id, relationship)

            if relationship is not None:
                relationships.append(relationship)

        return relationships

    @classmethod
    def from_xml(cls, parser: 'Parser', relationship_xml: etree._Element, file_path: PurePosixPath, origin: 'XmlPart'):
        relationship_id = relationship_xml.get('Id')
        target_type = relationship_xml.get('Type')
        raw_target_path = relationship_xml.get('Target')

        if relationship_id is None or target_type is None or raw_target_path is None:
            print("Integrity warning: Relationship element missing required attributes")
            return None

        target_path = cls._get_target_file_path(file_path, PurePosixPath(raw_target_path))

        target = parser.parse_part(target_path)

        if target is None:
            print(f"Integrity warning: Relationship target {target_path} could not be parsed")
            return None

        return cls(target_type, target, origin)

    def _to_xml(self, writer: 'Writer') -> etree._Element:
        target_file_name = writer.assign_part_index(self.target.part_name, self.target)
        target_location = PurePosixPath(self.target.base_path) / target_file_name if self.target.base_path else PurePosixPath(target_file_name)
        relative_target_path = posixpath.relpath(target_location.as_posix(), start=(self.origin.base_path or PurePosixPath('/')).as_posix())
        relationship_element = etree.Element('Relationship')
        relationship_element.set('Id', writer.assign_relationship_id(self.origin, self))
        relationship_element.set('Type', self.target_type)
        relationship_element.set('Target', relative_target_path)
        return relationship_element

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
