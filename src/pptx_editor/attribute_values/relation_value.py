import sys

from lxml import etree
from typing import TYPE_CHECKING

from pptx_editor.attribute_value import AttributeValue

if TYPE_CHECKING:
    from pptx_editor.parser import Parser
    from pptx_editor.part import Part

class RelationValue(AttributeValue):
    default_namespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    def __init__(self, name: str, value: str):
        q = etree.QName(name)
        self.namespace = sys.intern(q.namespace) if q.namespace else None
        self.name = sys.intern(q.localname)
        self.value = sys.intern(value)

    @classmethod
    def from_item(cls, parser: 'Parser', file_path: str | None, name: str, value: str) -> 'RelationValue':
        relation_value = cls(name, value)

        relation_value.resolve_target(parser, file_path)

        return relation_value

    def resolve_target(self, parser: 'Parser', file_path: str | None):
        if file_path is None:
            print('Integrity warning: Cannot resolve relation value without file path context')
            return

        relationship_id = self.value
        relationship = parser.get_relationship(file_path, relationship_id)
        if relationship is None:
            print(f"Integrity warning: No relationship found with id {relationship_id} in part {file_path}")
            return

        self.value = relationship
