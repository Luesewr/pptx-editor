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
    __slots__ = ['name', 'namespace', 'values', 'attributes', 'text', 'defined_namespace']

    def __init__(self, name: str, namespace: str | None, values: Iterable['AttributeValue'], attributes: Iterable['Attribute'], text: str | None, defined_namespace: dict[str | None, str] | None = None):
        self.name = sys.intern(name)
        self.namespace = sys.intern(namespace) if namespace else None
        self.values = values
        self.attributes = attributes
        self.text = text
        self.defined_namespace = defined_namespace

    @classmethod
    def from_xml(cls, parser, file_path, xml, ns_declarations=None):
        q = etree.QName(xml)
        name = sys.intern(q.localname)
        namespace = sys.intern(q.namespace) if q.namespace else None
        values = tuple(cls.from_item(parser, file_path, str(key), str(value)) for key, value in xml.attrib.items())
        attributes = tuple(Attribute.from_xml(parser, file_path, child, ns_declarations) for child in xml)
        text = sys.intern(xml.text) if xml.text is not None else xml.text

        xml_path = xml.getroottree().getpath(xml)

        defined_namespace = None
        if ns_declarations is not None and xml_path in ns_declarations:
            defined_namespace = {
                pfx: sys.intern(uri)
                for pfx, uri in ns_declarations[xml_path].items()
            }

        return cls(name, namespace, values, attributes, text, defined_namespace=defined_namespace)

    @classmethod
    def from_item(cls, parser: 'Parser', file_path: PurePosixPath | None, name: str, value: str) -> 'AttributeValue':
        registry = AttributeValueRegistry()
        q = etree.QName(name)
        namespace = sys.intern(q.namespace) if q.namespace else None
        attribute_value_cls = registry.get_attribute_value_cls(namespace)
        return attribute_value_cls.from_item(parser, file_path, name, value)

    def get_relationships(self) -> list['Relationship']:
        relationships = []
        for value in self.values:
            if isinstance(value, RelationshipValue) and hasattr(value, 'value') and isinstance(value.value, Relationship):
                relationships.append(value.value)

        for attribute in self.attributes:
            relationships.extend(attribute.get_relationships())

        return relationships

    def get_values(self, name: str, namespace: str | None = None) -> list[AttributeValue]:
        return [value for value in self.values if value.name == name and value.namespace == namespace]

    def to_xml(self, writer: 'Writer', namespaces: dict) -> etree._Element:
        if self.defined_namespace:
            namespaces = namespaces.copy()
            namespaces.update(self.defined_namespace)

        qname = etree.QName(self.namespace, self.name) if self.namespace else self.name
        element = etree.Element(qname, nsmap=namespaces)

        if self.text is not None:
            element.text = self.text

        for value in self.values:
            value.to_xml(element, writer)


        for attribute in self.attributes:
            child_element = attribute.to_xml(writer, namespaces)
            element.append(child_element)

        return element

    def pretty_print(self, indent=0):
        indent_str = ' ' * indent
        print(f"{indent_str}{self}")
        for child in self.attributes:
            child.pretty_print(indent + 2)

    def __str__(self):
        return f"Attribute(name={self.name}, namespace={self.namespace}, values={[str(value) for value in self.values]})"

    def __repr__(self):
        return self.__str__()
