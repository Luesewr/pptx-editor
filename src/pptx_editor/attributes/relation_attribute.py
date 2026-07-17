from io import BytesIO
from pathlib import PurePosixPath
from xml.sax.saxutils import escape
import sys

from lxml import etree
from typing import TYPE_CHECKING

from pptx_editor.attribute import Attribute
from pptx_editor.relationship import Relationship

if TYPE_CHECKING:
    from pptx_editor.writer import _OOXMLWriter
    from pptx_editor.parts.xml_part import XmlPart

class RelationshipAttribute(Attribute):
    default_namespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    default_prefix = 'r'
    default_name = None
    default_element_name = None

    def __init__(self, value: 'str | Relationship', prefix: str | None = None, name: str | None = None, overwrite_prefix: bool = False):
        self.prefix = prefix if prefix or overwrite_prefix else self.default_prefix
        self.name = sys.intern(name) if name else self.default_name
        self.value: 'str | Relationship' = sys.intern(value) if isinstance(value, str) else value

    @classmethod
    def _from_item(cls, part: 'XmlPart', namespaces: dict[str | None, str], name: str, value: str) -> 'RelationshipAttribute':
        q = etree.QName(name)
        namespace = sys.intern(q.namespace) if q.namespace else None
        prefix = [pfx for pfx, uri in namespaces.items() if uri == namespace][0] if namespace is not None else None
        relation_value = cls(value, prefix, q.localname)

        relation_value._resolve_target(part)

        return relation_value

    def _to_xml(self, writer: '_OOXMLWriter', buffer: BytesIO):
        if isinstance(self.value, str):
            escaped_value = escape(self.value, entities={'"': '&quot;', "'": '&apos;', '\n': '&#10;', '\r': '&#13;', '\t': '&#9;'})
            buffer.write(f' {self.prefix + ":" if self.prefix else ""}{self.name}="{escaped_value}"'.encode('utf-8'))
        else:
            relationship_id = writer.assign_relationship_id(self.value.origin, self.value)
            escaped_relationship_id = escape(relationship_id, entities={'"': '&quot;', "'": '&apos;', '\n': '&#10;', '\r': '&#13;', '\t': '&#9;'})
            buffer.write(f' {self.prefix + ":" if self.prefix else ""}{self.name}="{escaped_relationship_id}"'.encode('utf-8'))

    def _resolve_target(self, part: 'XmlPart'):
        if not isinstance(self.value, str):
            return

        relationship_id = self.value
        relationship = part.get_relationship(relationship_id)
        if relationship is None:
            escaped_relationship_id = escape(relationship_id, entities={'"': '&quot;', "'": '&apos;', '\n': '&#10;', '\r': '&#13;', '\t': '&#9;'})
            print(f"Integrity warning: No relationship found with id {escaped_relationship_id} in part {part.part_name}")
            return

        self.value = relationship

    def __str__(self):
        return f"RelationshipAttribute(name={self.name}, value={self.value.target.part_name if isinstance(self.value, Relationship) else self.value})"
