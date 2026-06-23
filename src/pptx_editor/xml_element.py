import sys

from io import BytesIO
from itertools import chain
from pathlib import PurePosixPath
from typing import TYPE_CHECKING, TypeVar, Generic
from xml.sax.saxutils import escape

from lxml import etree

from pptx_editor.attribute import Attribute, AttributeRegistry
from pptx_editor.attributes.relation_attribute import RelationshipAttribute
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.relationship import Relationship
from pptx_editor.singleton import SingletonMeta

if TYPE_CHECKING:
    from pptx_editor.parser import _OOXMLParser
    from pptx_editor.writer import _OOXMLWriter
    from pptx_editor.parts.xml_part import XmlPart

T = TypeVar('T', bound='XmlElement')

class XmlElementRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}

    def register(self, namespace: str, name: str, attribute_cls):
        self._registry[(namespace, name)] = attribute_cls

    def get_attribute_cls(self, namespace: str | None, name: str | None = None) -> type['XmlElement']:
        return self._registry.get((namespace, name), XmlElement)

class XmlElementProperty(Generic[T]):
    def __init__(self, element_type: type[T], nullable: bool = True):
        self.element_type = element_type
        self.nullable = nullable

    def __get__(self, instance: T | None, owner: type[T], nullable_override: bool | None = None) -> T | None:
        if instance is None:
            raise AttributeError("XmlElementProperty can only be accessed from an instance.")

        for child in instance.children:
            if isinstance(child, self.element_type):
                return child

        if not self.nullable and not nullable_override:
            raise PowerpointIntegrityError(f"Expected a child of type {self.element_type.__name__} in {instance.__class__.__name__}, but none was found.")

        return None

    def __set__(self, instance: T, value: T | None) -> None:
        existing_element = self.__get__(instance, type(instance), nullable_override=True)

        if value is None and not self.nullable:
            raise PowerpointIntegrityError(f"Cannot set a non-nullable XmlElementProperty to None in {instance.__class__.__name__}.")

        if value is not None and value.part is not instance.part:
            value = value.copy()
            value.part = instance.part

        if existing_element is not None:
            if value is not None:
                instance.replace_element(existing_element, value)
            else:
                instance.remove_element(existing_element)
        elif value is not None:
            instance.children = (*instance.children, value)

class XmlElement:
    default_namespace: str | None = None
    default_prefix: str | None = None
    default_name: str | None = None

    __slots__ = ['name', 'prefix', 'attributes', 'children', 'text', 'tail', 'part', 'namespaces']

    def __init__(self, name: str | None = None, prefix: str | None = None, attributes: tuple['Attribute', ...] = (), children: tuple['XmlElement', ...] = (), text: str | None = None, tail: str | None = None, part: 'XmlPart | None' = None, namespaces: dict[str | None, str] | None = None):
        self.name = sys.intern(name) if name else self.default_name
        self.prefix = sys.intern(prefix) if prefix else self.default_prefix
        self.attributes = attributes
        self.children = children
        self.text = text
        self.tail = tail
        self.part = part
        self.namespaces = namespaces

    def get_relationships(self) -> list['Relationship']:
        relationships = []
        for value in self.attributes:
            if isinstance(value, RelationshipAttribute) and hasattr(value, 'value') and isinstance(value.value, Relationship):
                relationships.append(value.value)

        for child in self.children:
            relationships.extend(child.get_relationships())

        return relationships

    def get_element(self, name: str, prefix: str | None = None) -> 'XmlElement | None':
        for element in self.children:
            if element.name == name and element.prefix == prefix:
                return element

        return None

    def get_elements(self, name: str, prefix: str | None = None) -> list['XmlElement']:
        return [element for element in self.children if element.name == name and element.prefix == prefix]

    def replace_element(self, old_element: 'XmlElement', new_element: 'XmlElement') -> None:
        index = next((i for i, obj in enumerate(self.children) if obj is old_element), None)

        if index is None:
            raise ValueError('Old element is not a child of this element.')

        self.children = tuple(chain(self.children[:index], (new_element,), self.children[index + 1:]))

    def remove_element(self, element: 'XmlElement') -> None:
        index = next((i for i, obj in enumerate(self.children) if obj is element), None)

        if index is None:
            raise ValueError('Element is not a child of this element.')

        self.children = tuple(chain(self.children[:index], self.children[index + 1:]))

    def add_element(self, element: 'XmlElement') -> None:
        self.children = tuple(chain(self.children, (element,)))

    def get_attribute(self, name: str, prefix: str | None = None) -> 'Attribute | None':
        for value in self.attributes:
            if value.name == name and value.prefix == prefix:
                return value

        return None

    def get_attributes(self, name: str, prefix: str | None = None) -> list['Attribute']:
        return [value for value in self.attributes if value.name == name and value.prefix == prefix]

    def copy(self):
        copied_attributes = tuple(value.copy() for value in self.attributes)
        copied_children = tuple(child.copy() for child in self.children)
        return self.__class__(self.name, self.prefix, copied_attributes, copied_children, self.text, self.tail, self.part, self.namespaces.copy() if self.namespaces is not None else None)

    def insert_element_before(self, new_element: 'XmlElement', reference_element: 'XmlElement') -> None:
        index = next((i for i, obj in enumerate(self.children) if obj is reference_element), None)

        if index is None:
            raise ValueError('Reference element is not a child of this element.')

        self.children = tuple(chain(self.children[:index], (new_element,), self.children[index:]))

    def insert_element_after(self, new_element: 'XmlElement', reference_element: 'XmlElement') -> None:
        index = next((i for i, obj in enumerate(self.children) if obj is reference_element), None)

        if index is None:
            raise ValueError('Reference element is not a child of this element.')

        self.children = tuple(chain(self.children[:index + 1], (new_element,), self.children[index + 1:]))

    @classmethod
    def _from_xml(cls, parser: '_OOXMLParser', file_path: PurePosixPath | None, xml: etree._Element, ns_declarations: dict[str, dict[str | None, str]] | None = None):
        from pptx_editor.parts.xml_part import XmlPart

        q = etree.QName(xml)
        name = sys.intern(q.localname)
        prefix = xml.prefix or None
        attributes = tuple(cls._from_item(parser, file_path, xml.nsmap, str(key), str(value)) for key, value in xml.attrib.items())
        children = tuple(XmlPart._parse_xml(parser, file_path, child, ns_declarations) for child in xml)
        text = sys.intern(xml.text) if xml.text is not None else xml.text
        tail = sys.intern(xml.tail) if xml.tail is not None else xml.tail
        part = parser.get_part(file_path) if file_path is not None else None
        xml_path = xml.getroottree().getpath(xml)

        defined_namespace = []
        if ns_declarations is not None and xml_path in ns_declarations:
            defined_namespace = ns_declarations[xml_path].copy()

        return cls(name, prefix, attributes, children, text, tail, part=part, namespaces=defined_namespace)

    @classmethod
    def _from_item(cls, parser: '_OOXMLParser', file_path: PurePosixPath | None, namespaces: dict[str | None, str], name: str, value: str) -> 'Attribute':
        registry = AttributeRegistry()
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

        for prefix, uri in self.namespaces.items() if self.namespaces else []:
            if prefix:
                buffer.write(f' xmlns:{prefix}="{uri}"'.encode('utf-8'))
            else:
                buffer.write(f' xmlns="{uri}"'.encode('utf-8'))

        for value in self.attributes:
            value._to_xml(writer, buffer)

        if not self.children and self.text is None:
            buffer.write('/>'.encode('utf-8'))
            if self.tail is not None:
                buffer.write(escape(self.tail).encode('utf-8'))
            return

        buffer.write('>'.encode('utf-8'))

        if self.text is not None:
            buffer.write(escape(self.text).encode('utf-8'))

        for attribute in self.children:
            attribute._to_xml(writer, buffer)

        if self.tail is not None:
            buffer.write(escape(self.tail).encode('utf-8'))

        buffer.write(f'</{self.prefix + ":" if self.prefix else ""}{self.name}>'.encode('utf-8'))

    @classmethod
    def _register(cls):
        registry = XmlElementRegistry()

        if cls.default_namespace is not None and cls.default_name is not None:
            registry.register(cls.default_namespace, cls.default_name, cls)

    def __init_subclass__(cls, **kwargs) -> None:
        super().__init_subclass__(**kwargs)

        if getattr(cls, 'is_abstract', False):
            return

        class_exceptions = []
        missing_namespace = not hasattr(cls, 'default_namespace') or cls.default_namespace is None
        missing_prefix = not hasattr(cls, 'default_prefix') or cls.default_prefix is None
        missing_name = not hasattr(cls, 'default_name') or cls.default_name is None

        if cls.__name__ not in class_exceptions and missing_namespace:
            raise ValueError(f"AttributeValue subclass {cls.__name__} must define a default_namespace class attribute")
        if cls.__name__ not in class_exceptions and missing_name:
            raise ValueError(f"AttributeValue subclass {cls.__name__} must define a default_name class attribute")
        if cls.__name__ not in class_exceptions and missing_prefix:
            raise ValueError(f"AttributeValue subclass {cls.__name__} must define a default_prefix class attribute")

        if cls.__name__ not in class_exceptions:
            cls._register()

    def __hash__(self):
        return hash((self.name, self.prefix, self.attributes, self.children, self.text, self.tail, self.part, frozenset(self.namespaces.items()) if self.namespaces else None))

    def __eq__(self, other):
        if not isinstance(other, self.__class__):
            return False

        if self is other:
            return True

        return (
            self.name,
            self.prefix,
            self.attributes,
            self.children,
            self.text,
            self.tail,
            self.part,
            self.namespaces
        ) == (
            other.name,
            other.prefix,
            other.attributes,
            other.children,
            other.text,
            other.tail,
            other.part,
            other.namespaces
        )

    def __str__(self):
        return f"{self.__class__.__name__}(name={escape(self.name)}, namespace={escape(self.prefix) if self.prefix else None}, attributes={[str(value) for value in self.attributes]})"

    def __repr__(self):
        return self.__str__()
