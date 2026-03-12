import sys

from lxml import etree

from pptx_editor.parser import Parser

class Relationship:
    def __init__(self, relationship_id: str, target_type: str, target: str):
        self.id = relationship_id
        self.target_type = sys.intern(target_type)
        self.target = target

    @classmethod
    def from_xml(cls, relationship_xml: etree._Element):
        relationship_id = relationship_xml.get('Id')
        target_type = relationship_xml.get('Type')
        target = relationship_xml.get('Target')

        if relationship_id is None or target_type is None or target is None:
            print("Integrity warning: Relationship element missing required attributes")
            return None

        return cls(relationship_id, target_type, target)
