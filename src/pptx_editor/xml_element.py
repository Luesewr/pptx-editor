import sys

from io import BytesIO
from typing import TYPE_CHECKING, TypeVar, TypeIs
from xml.sax.saxutils import escape

from lxml import etree

import pptx_editor.xml_elements as xml_elements
import pptx_editor.parser as parser

from pptx_editor.attributes.relation_attribute import RelationshipAttribute
from pptx_editor.attribute import Attribute
from pptx_editor.modes.xml_element import AddMode
from pptx_editor.relationship import Relationship
from pptx_editor.singleton import SingletonMeta

if TYPE_CHECKING:
    from pptx_editor.writer import _OOXMLWriter
    from pptx_editor.parts.xml_part import XmlPart
    from pptx_editor.xml_elements.null import NullElement

T = TypeVar('T', bound='XmlElement')
U = TypeVar('U', bound='Attribute')

class XmlElementRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}

    def register(self, namespace: str, name: str, element_cls: type['XmlElement']):
        self._registry[(namespace, name)] = element_cls

    def get_element_cls(self, namespace: str | None, name: str | None = None) -> type['XmlElement']:
        return self._registry.get((namespace, name), XmlElement)

class XmlElement:
    default_namespace: str | None = None
    default_prefix: str | None = None
    default_name: str | None = None
    default_order: tuple[str | tuple[str, ...], ...] | None = None

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

    def get_element(self, name: str, prefix: str | None = None) -> 'XmlElement | None':
        for element in self.children:
            if element.name == name and element.prefix == prefix:
                return element

        return None

    def get_element_by_type(self, element_type: type[T]) -> 'T | None':
        for element in self.children:
            if isinstance(element, element_type):
                return element

        return None

    def get_elements(self, name: str, prefix: str | None = None) -> list['XmlElement']:
        return [element for element in self.children if element.name == name and element.prefix == prefix]

    def get_elements_by_type(self, element_type: type[T]) -> list[T]:
        return [element for element in self.children if isinstance(element, element_type)]

    def add_element(self, element: 'XmlElement', index: int | None = None) -> None:
        if index is None:
            self.children = (*self.children, element)
        else:
            self.children = (*self.children[:index], element, *self.children[index:])

    def auto_add_element(self, element: 'XmlElement', index: int | None = None, add_mode: AddMode = AddMode.SORT) -> None:
        if self.default_order is None:
            self.add_element(element, index=index)
            return

        if add_mode == AddMode.SORT:
            self.add_element(element, index=index)
            self.sort_children()

    def replace_element(self, old_element: T, new_element: T) -> None:
        index = self.index_of_element(old_element)

        self.children = (*self.children[:index], new_element, *self.children[index + 1:])

    def auto_replace_element(self, old_element: T, new_element: T, add_mode: AddMode = AddMode.SORT) -> None:
        index = self.index_of_element(old_element)

        self.children = (*self.children[:index], new_element, *self.children[index + 1:])

        if add_mode == AddMode.SORT:
            self.sort_children()

    def remove_element(self, element: 'XmlElement') -> None:
        index = self.index_of_element(element)

        self.children = (*self.children[:index], *self.children[index + 1:])

    def get_attribute(self, name: str, prefix: str | None = None) -> 'Attribute | None':
        for value in self.attributes:
            if value.name == name and value.prefix == prefix:
                return value

        return None

    def get_attribute_by_type(self, attribute_type: type[U]) -> 'U | None':
        for value in self.attributes:
            if isinstance(value, attribute_type):
                return value

        return None

    def get_attributes(self, name: str, prefix: str | None = None) -> list['Attribute']:
        return [value for value in self.attributes if value.name == name and value.prefix == prefix]

    def get_attributes_by_type(self, attribute_type: type[U]) -> list[U]:
        return [value for value in self.attributes if isinstance(value, attribute_type)]

    def add_attribute(self, attribute: 'Attribute') -> None:
        self.attributes = (*self.attributes, attribute)

    def replace_attribute(self, old_attribute: U, new_attribute: U) -> None:
        index = self.index_of_attribute(old_attribute)

        self.attributes = (*self.attributes[:index], new_attribute, *self.attributes[index + 1:])

    def remove_attribute(self, attribute: 'Attribute') -> None:
        index = self.index_of_attribute(attribute)

        self.attributes = (*self.attributes[:index], *self.attributes[index + 1:])

    def copy(self):
        copied_attributes = tuple(value.copy() for value in self.attributes)
        copied_children = tuple(child.copy() for child in self.children)
        return self.__class__(self.name, self.prefix, copied_attributes, copied_children, self.text, self.tail, self.part, self.namespaces.copy() if self.namespaces is not None else None)

    def insert_element_before(self, new_element: 'XmlElement', reference_element: 'XmlElement') -> None:
        index = self.index_of_element(reference_element)

        self.children = (*self.children[:index], new_element, *self.children[index:])

    def insert_element_after(self, new_element: 'XmlElement', reference_element: 'XmlElement') -> None:
        index = self.index_of_element(reference_element)

        self.children = (*self.children[:index + 1], new_element, *self.children[index + 1:])

    def index_of_element(self, element: 'XmlElement') -> int:
        index = next((i for i, obj in enumerate(self.children) if obj is element), None)

        if index is None:
            raise ValueError('Element is not a child of this element.')

        return index

    def index_of_attribute(self, attribute: 'Attribute') -> int:
        index = next((i for i, obj in enumerate(self.attributes) if obj is attribute), None)

        if index is None:
            raise ValueError('Attribute is not part of this element.')

        return index

    def sort_children(self) -> None:
        if self.default_order is not None:
            self.children = tuple(sorted(self.children, key=self._default_order_key))

    def update_part_recursive(self, part: 'XmlPart') -> None:
        self.part = part

        for child in self.children:
            child.update_part_recursive(part)

    def create_if_null(self, element_type: type[T] | None = None, add_mode: AddMode = AddMode.SORT):
        return self

    @staticmethod
    def is_null(element: 'T | NullElement[T]') -> TypeIs['NullElement[T]']:
        return isinstance(element, xml_elements.null.NullElement)

    @staticmethod
    def is_not_null(element: 'T | NullElement[T]') -> TypeIs[T]:
        return not isinstance(element, xml_elements.null.NullElement)

    def _default_order_key(self, element: 'XmlElement') -> int:
        if self.default_order is None:
            return 0

        for index, name in enumerate(self.default_order):
            if isinstance(name, tuple):
                if element.name in name:
                    return index
            elif element.name == name:
                return index

        return len(self.default_order)

    @classmethod
    def _from_xml(cls, part: 'XmlPart', xml: etree._Element, ns_declarations: dict[str, dict[str | None, str]] | None = None):
        q = etree.QName(xml)
        name = sys.intern(q.localname)
        prefix = xml.prefix or None
        attributes = tuple(parser._OOXMLParser.parse_attribute_from_item(part, xml.nsmap, name, str(key), str(value)) for key, value in xml.attrib.items())
        children = tuple(parser._OOXMLParser.parse_element_from_xml(part, child, ns_declarations) for child in xml)
        text = sys.intern(xml.text) if xml.text is not None else xml.text
        tail = sys.intern(xml.tail) if xml.tail is not None else xml.tail
        xml_path = xml.getroottree().getpath(xml)

        defined_namespace = []
        if ns_declarations is not None and xml_path in ns_declarations:
            defined_namespace = ns_declarations[xml_path].copy()

        return cls(name, prefix, attributes, children, text, tail, part=part, namespaces=defined_namespace)

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

        class_exceptions = ['NullElement']
        missing_namespace = not hasattr(cls, 'default_namespace') or cls.default_namespace is None
        missing_prefix = not hasattr(cls, 'default_prefix') or cls.default_prefix is None
        missing_name = not hasattr(cls, 'default_name') or cls.default_name is None

        if cls.__name__ not in class_exceptions and missing_namespace:
            raise ValueError(f"XmlElement subclass {cls.__name__} must define a default_namespace class attribute")
        if cls.__name__ not in class_exceptions and missing_name:
            raise ValueError(f"XmlElement subclass {cls.__name__} must define a default_name class attribute")
        if cls.__name__ not in class_exceptions and missing_prefix:
            raise ValueError(f"XmlElement subclass {cls.__name__} must define a default_prefix class attribute")

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
