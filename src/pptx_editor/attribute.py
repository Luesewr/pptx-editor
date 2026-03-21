import sys

from lxml import etree
from typing import TYPE_CHECKING, Iterable

from pptx_editor.attribute_value import AttributeValue, AttributeValueRegistry

if TYPE_CHECKING:
    from pptx_editor.parser import Parser

class Attribute:
    __slots__ = ['name', 'namespace', 'values', 'attributes']

    def __init__(self, name: str, namespace: str | None, values: Iterable['AttributeValue'], attributes: Iterable['Attribute']):
        self.name = sys.intern(name)
        self.namespace = sys.intern(namespace) if namespace else None
        self.values = values
        self.attributes = attributes

    @classmethod
    def from_xml(cls, parser: 'Parser', file_path: str | None, xml: etree._Element) -> 'Attribute':
        q = etree.QName(xml)
        name = sys.intern(q.localname)
        namespace = sys.intern(q.namespace) if q.namespace else None
        values = tuple(cls.from_item(parser, file_path, str(key), str(value)) for key, value in xml.attrib.items())
        attributes = tuple(Attribute.from_xml(parser, file_path, child) for child in xml)
        return cls(name, namespace, values, attributes)

    @classmethod
    def from_item(cls, parser: 'Parser', file_path: str | None, name: str, value: str) -> 'AttributeValue':
        registry = AttributeValueRegistry()
        q = etree.QName(name)
        namespace = sys.intern(q.namespace) if q.namespace else None
        attribute_value_cls = registry.get_attribute_value_cls(namespace)
        return attribute_value_cls.from_item(parser, file_path, name, value)

    def get_values(self, name: str, namespace: str | None = None) -> list[AttributeValue]:
        return [value for value in self.values if value.name == name and value.namespace == namespace]

    def pretty_print(self, indent=0):
        indent_str = ' ' * indent
        print(f"{indent_str}{self}")
        for child in self.attributes:
            child.pretty_print(indent + 2)

    def __str__(self):
        return f"Attribute(name={self.name}, namespace={self.namespace}, values={[str(value) for value in self.values]})"

    def __repr__(self):
        return self.__str__()
