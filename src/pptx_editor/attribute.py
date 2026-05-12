from io import BytesIO
from pathlib import PurePosixPath
import sys

from lxml import etree
from typing import TYPE_CHECKING, Iterable

from pptx_editor.attribute_value import AttributeValue, AttributeValueRegistry
from pptx_editor.attribute_values.relation_value import RelationshipValue
from pptx_editor.relationship import Relationship

if TYPE_CHECKING:
    from pptx_editor.parser import Parser
    from pptx_editor.writer import Writer

class Attribute:
    __slots__ = ['name', 'prefix', 'values', 'attributes', 'text', 'tail', 'defined_namespace']

    def __init__(self, name: str, prefix: str | None, values: Iterable['AttributeValue'], attributes: Iterable['Attribute'], text: str | None, tail: str | None = None, defined_namespace: dict[str | None, str] | None = None):
        self.name = sys.intern(name)
        self.prefix = sys.intern(prefix) if prefix else None
        self.values = values
        self.attributes = attributes
        self.text = text
        self.tail = tail
        self.defined_namespace = defined_namespace

    @classmethod
    def from_xml(cls, parser: 'Parser', file_path: PurePosixPath | None, xml: etree._Element, ns_declarations: dict[str, dict[str | None, str]] | None = None):
        q = etree.QName(xml)
        name = sys.intern(q.localname)
        prefix = xml.prefix or None
        values = tuple(cls.from_item(parser, file_path, xml.nsmap, str(key), str(value)) for key, value in xml.attrib.items())
        attributes = tuple(Attribute.from_xml(parser, file_path, child, ns_declarations) for child in xml)
        text = sys.intern(xml.text) if xml.text is not None else xml.text
        tail = sys.intern(xml.tail) if xml.tail is not None else xml.tail
        xml_path = xml.getroottree().getpath(xml)

        defined_namespace = None
        if ns_declarations is not None and xml_path in ns_declarations:
            defined_namespace = {
                pfx: uri
                for pfx, uri in ns_declarations[xml_path].items()
            }

        return cls(name, prefix, values, attributes, text, tail, defined_namespace=defined_namespace)

    @classmethod
    def from_item(cls, parser: 'Parser', file_path: PurePosixPath | None, namespaces: dict[str | None, str], name: str, value: str) -> 'AttributeValue':
        registry = AttributeValueRegistry()
        q = etree.QName(name)
        namespace = sys.intern(q.namespace) if q.namespace else None
        attribute_value_cls = registry.get_attribute_value_cls(namespace)
        return attribute_value_cls.from_item(parser, file_path, namespaces, name, value)

    def get_relationships(self) -> list['Relationship']:
        relationships = []
        for value in self.values:
            if isinstance(value, RelationshipValue) and hasattr(value, 'value') and isinstance(value.value, Relationship):
                relationships.append(value.value)

        for attribute in self.attributes:
            relationships.extend(attribute.get_relationships())

        return relationships

    def get_values(self, name: str, namespace: str | None = None) -> list[AttributeValue]:
        return [value for value in self.values if value.name == name and value.prefix == namespace]

    def to_xml(self, writer: 'Writer', buffer: BytesIO):
        buffer.write('<'.encode('utf-8'))

        if self.prefix:
            buffer.write(f"{self.prefix}:{self.name}".encode('utf-8'))
        else:
            buffer.write(self.name.encode('utf-8'))

        for prefix, uri in self.defined_namespace.items() if self.defined_namespace else []:
            if prefix:
                buffer.write(f' xmlns:{prefix}="{uri}"'.encode('utf-8'))
            else:
                buffer.write(f' xmlns="{uri}"'.encode('utf-8'))

        for value in self.values:
            value.to_xml(writer, buffer)

        if not self.attributes and self.text is None:
            buffer.write('/>'.encode('utf-8'))
            return

        buffer.write('>'.encode('utf-8'))

        if self.text is not None:
            buffer.write(self.text.encode('utf-8'))

        for attribute in self.attributes:
            attribute.to_xml(writer, buffer)

        if self.tail is not None:
            buffer.write(self.tail.encode('utf-8'))

        buffer.write(f'</{self.prefix + ":" if self.prefix else ""}{self.name}>'.encode('utf-8'))

    def pretty_print(self, indent=0):
        indent_str = ' ' * indent
        print(f"{indent_str}{self}")
        for child in self.attributes:
            child.pretty_print(indent + 2)

    def __str__(self):
        return f"Attribute(name={self.name}, namespace={self.prefix}, values={[str(value) for value in self.values]})"

    def __repr__(self):
        return self.__str__()
