import sys

from typing import TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from pptx_editor.parser import Parser

class Relationship:
    def __init__(self, relationship_id: str, target_type: str, target: str, location: str):
        self.id = relationship_id
        self.target_type = sys.intern(target_type)
        self.target = target
        self.location = location

    @classmethod
    def from_file(cls, parser: 'Parser', file_path: str):
        relationship_xml = parser.read_file(file_path)
        relationship_tree = etree.fromstring(relationship_xml)

        relationships = []
        for relationship_element in relationship_tree:
            relationship = cls.from_xml(relationship_element, file_path)
            if relationship is not None:
                relationships.append(relationship)

        return relationships

    @classmethod
    def from_xml(cls, relationship_xml: etree._Element, location: str):
        relationship_id = relationship_xml.get('Id')
        target_type = relationship_xml.get('Type')
        target = relationship_xml.get('Target')

        if relationship_id is None or target_type is None or target is None:
            print("Integrity warning: Relationship element missing required attributes")
            return None

        return cls(relationship_id, target_type, target, location)

    def get_target_file_path(self):
        base_path = f'/{self.location.rsplit("/", 1)[0]}'
        current_target = f'../{self.target}'
        while current_target.startswith('../'):
            base_path = base_path.rsplit('/', 1)[0]
            current_target = current_target.removeprefix('../')

        return base_path + '/' + current_target
