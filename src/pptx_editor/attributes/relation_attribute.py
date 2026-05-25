from io import BytesIO
from pathlib import PurePosixPath
from xml.sax.saxutils import escape
import sys

from lxml import etree
from typing import TYPE_CHECKING

from pptx_editor.attribute import AttributeValue

if TYPE_CHECKING:
    from pptx_editor.writer import _OOXMLWriter
    from pptx_editor.parser import _OOXMLParser
    from pptx_editor.relationship import Relationship

class RelationshipValue(AttributeValue):
    default_namespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    def __init__(self, prefix: str | None, name: str, value: str):
        self.prefix = prefix if prefix else None
        self.name = sys.intern(name)
        self.value: 'str | Relationship' = sys.intern(value)

    @classmethod
    def _from_item(cls, parser: '_OOXMLParser', file_path: PurePosixPath | None, namespaces: dict[str | None, str], name: str, value: str) -> 'RelationshipValue':
        q = etree.QName(name)
        namespace = sys.intern(q.namespace) if q.namespace else None
        prefix = [pfx for pfx, uri in namespaces.items() if uri == namespace][0] if namespace is not None else None
        relation_value = cls(prefix, q.localname, value)

        relation_value._resolve_target(parser, file_path)

        return relation_value

    def _to_xml(self, writer: '_OOXMLWriter', buffer: BytesIO):
        if isinstance(self.value, str):
            escaped_value = escape(self.value, entities={'"': '&quot;', "'": '&apos;', '\n': '&#10;', '\r': '&#13;', '\t': '&#9;'})
            buffer.write(f' {self.prefix + ":" if self.prefix else ""}{self.name}="{escaped_value}"'.encode('utf-8'))
        else:
            relationship_id = writer.assign_relationship_id(self.value.origin, self.value)
            escaped_relationship_id = escape(relationship_id, entities={'"': '&quot;', "'": '&apos;', '\n': '&#10;', '\r': '&#13;', '\t': '&#9;'})
            buffer.write(f' {self.prefix + ":" if self.prefix else ""}{self.name}="{escaped_relationship_id}"'.encode('utf-8'))

    def _resolve_target(self, parser: '_OOXMLParser', file_path: PurePosixPath | None):
        if file_path is None:
            print('Integrity warning: Cannot resolve relation value without file path context')
            return

        if not isinstance(self.value, str):
            return

        relationship_id = self.value
        relationship = parser.get_relationship(file_path, relationship_id)
        if relationship is None:
            escaped_relationship_id = escape(relationship_id, entities={'"': '&quot;', "'": '&apos;', '\n': '&#10;', '\r': '&#13;', '\t': '&#9;'})
            print(f"Integrity warning: No relationship found with id {escaped_relationship_id} in part {file_path}")
            return

        self.value = relationship
