import sys

from typing import TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from pptx_editor.parser import Parser
    from pptx_editor.part import Part

class Relationship:
    __slots__ = ['target_type', 'target', 'origin']

    def __init__(self, target_type: str, target: 'Part', origin: 'Part'):
        self.target_type = sys.intern(target_type)
        self.target = target
        self.origin = origin

    @classmethod
    def from_file(cls, parser: 'Parser', file_path: str, origin: 'Part'):
        relationship_xml = parser.read_file(file_path.lstrip('/'))
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
    def from_xml(cls, parser: 'Parser', relationship_xml: etree._Element, file_path: str, origin: 'Part'):
        relationship_id = relationship_xml.get('Id')
        target_type = relationship_xml.get('Type')
        raw_target_path = relationship_xml.get('Target')

        if relationship_id is None or target_type is None or raw_target_path is None:
            print("Integrity warning: Relationship element missing required attributes")
            return None

        target_path = cls._get_target_file_path(file_path, raw_target_path)

        target = parser.parse_part(target_path)

        if target is None:
            print(f"Integrity warning: Relationship target {target_path} could not be parsed")
            return None

        return cls(target_type, target, origin)

    @staticmethod
    def _get_target_file_path(location: str, target: str):
        base_path = f'{location.rsplit("/", 1)[0]}'
        current_target = f'../{target}'

        while current_target.startswith('../'):
            base_path = base_path.rsplit('/', 1)[0]
            current_target = current_target.removeprefix('../')

        path = base_path + '/' + current_target

        if not path.startswith('/'):
            path = '/' + path

        return path

    @staticmethod
    def _get_original_file_path(file_path: str):
        file_name = file_path.rsplit('/', 1)[-1].removesuffix('.rels')

        part_file_path = file_path.rsplit('/', 2)[0] + '/' + file_name

        if not part_file_path.startswith('/'):
            part_file_path = '/' + part_file_path

        return part_file_path
