from io import BytesIO
from pathlib import PurePosixPath
import sys

from lxml import etree
from typing import TYPE_CHECKING, Iterable
from xml.sax.saxutils import escape

from pptx_editor.attribute_value import AttributeValue, AttributeValueRegistry
from pptx_editor.attribute_values.relation_value import RelationshipValue
from pptx_editor.relationship import Relationship
from pptx_editor.singleton import SingletonMeta

if TYPE_CHECKING:
    from pptx_editor.parser import _OOXMLParser
    from pptx_editor.writer import _OOXMLWriter

class AttributeRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}

    def register(self, namespace: str, name: str, attribute_cls):
        self._registry[(namespace, name)] = attribute_cls

    def get_attribute_cls(self, namespace: str | None, name: str | None = None) -> type['Attribute']:
        return self._registry.get((namespace, name), Attribute)

class Attribute:
    default_namespace: str | None = None
    default_name: str | None = None

    __slots__ = ['name', 'prefix', 'values', 'attributes', 'text', 'tail', 'defined_namespace']

    def __init__(self, name: str, prefix: str | None, values: tuple['AttributeValue', ...], attributes: Iterable['Attribute'], text: str | None, tail: str | None = None, defined_namespace: dict[str | None, str] | None = None):
        self.name = sys.intern(name)
        self.prefix = sys.intern(prefix) if prefix else None
        self.values = values
        self.attributes = attributes
        self.text = text
        self.tail = tail
        self.defined_namespace = defined_namespace

    def get_relationships(self) -> list['Relationship']:
        relationships = []
        for value in self.values:
            if isinstance(value, RelationshipValue) and hasattr(value, 'value') and isinstance(value.value, Relationship):
                relationships.append(value.value)

        for attribute in self.attributes:
            relationships.extend(attribute.get_relationships())

        return relationships

    def get_attribute(self, name: str, prefix: str | None = None) -> 'Attribute | None':
        for attribute in self.attributes:
            if attribute.name == name and attribute.prefix == prefix:
                return attribute

        return None

    def get_attributes(self, name: str, prefix: str | None = None) -> list['Attribute']:
        return [attribute for attribute in self.attributes if attribute.name == name and attribute.prefix == prefix]

    def get_value(self, name: str, prefix: str | None = None) -> 'AttributeValue | None':
        for value in self.values:
            if value.name == name and value.prefix == prefix:
                return value

        return None

    def get_values(self, name: str, prefix: str | None = None) -> list[AttributeValue]:
        return [value for value in self.values if value.name == name and value.prefix == prefix]

    @classmethod
    def _from_xml(cls, parser: '_OOXMLParser', file_path: PurePosixPath | None, xml: etree._Element, ns_declarations: dict[str, dict[str | None, str]] | None = None):
        from pptx_editor.parts.xml_part import XmlPart

        q = etree.QName(xml)
        name = sys.intern(q.localname)
        prefix = xml.prefix or None
        values = tuple(cls._from_item(parser, file_path, xml.nsmap, str(key), str(value)) for key, value in xml.attrib.items())
        attributes = tuple(XmlPart._parse_xml(parser, file_path, child, ns_declarations) for child in xml)
        text = sys.intern(xml.text) if xml.text is not None else xml.text
        tail = sys.intern(xml.tail) if xml.tail is not None else xml.tail
        xml_path = xml.getroottree().getpath(xml)

        defined_namespace = None
        if ns_declarations is not None and xml_path in ns_declarations:
            defined_namespace = dict(ns_declarations[xml_path].items())

        return cls(name, prefix, values, attributes, text, tail, defined_namespace=defined_namespace)

    @classmethod
    def _from_item(cls, parser: '_OOXMLParser', file_path: PurePosixPath | None, namespaces: dict[str | None, str], name: str, value: str) -> 'AttributeValue':
        registry = AttributeValueRegistry()
        q = etree.QName(name)
        namespace = sys.intern(q.namespace) if q.namespace else None
        attribute_value_cls = registry.get_attribute_value_cls(namespace)
        return attribute_value_cls._from_item(parser, file_path, namespaces, name, value)

    def _to_xml(self, writer: '_OOXMLWriter', buffer: BytesIO):
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
            value._to_xml(writer, buffer)

        if not self.attributes and self.text is None:
            buffer.write('/>'.encode('utf-8'))
            return

        buffer.write('>'.encode('utf-8'))

        if self.text is not None:
            buffer.write(escape(self.text).encode('utf-8'))

        for attribute in self.attributes:
            attribute._to_xml(writer, buffer)

        if self.tail is not None:
            buffer.write(escape(self.tail).encode('utf-8'))

        buffer.write(f'</{self.prefix + ":" if self.prefix else ""}{self.name}>'.encode('utf-8'))

    @classmethod
    def _register(cls):
        registry = AttributeRegistry()

        if cls.default_namespace is not None and cls.default_name is not None:
            registry.register(cls.default_namespace, cls.default_name, cls)

    def __init_subclass__(cls, **kwargs) -> None:
        super().__init_subclass__(**kwargs)

        class_exceptions = []
        missing_namespace = not hasattr(cls, 'default_namespace') or cls.default_namespace is None
        missing_name = not hasattr(cls, 'default_name') or cls.default_name is None

        if cls.__name__ not in class_exceptions and missing_namespace:
            raise ValueError(f"AttributeValue subclass {cls.__name__} must define a default_namespace class attribute")
        if cls.__name__ not in class_exceptions and missing_name:
            raise ValueError(f"AttributeValue subclass {cls.__name__} must define a default_name class attribute")

        if cls.__name__ not in class_exceptions:
            cls._register()

    def __str__(self):
        return f"{self.__class__.__name__}(name={escape(self.name)}, namespace={escape(self.prefix) if self.prefix else None}, values={[str(value) for value in self.values]})"

    def __repr__(self):
        return self.__str__()
