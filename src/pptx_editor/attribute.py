import sys

from io import BytesIO
from pathlib import PurePosixPath
from typing import TYPE_CHECKING, Generic, TypeVar
from xml.sax.saxutils import escape

from lxml import etree

from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.singleton import SingletonMeta

if TYPE_CHECKING:
    from pptx_editor.parser import _OOXMLParser
    from pptx_editor.writer import _OOXMLWriter
    from pptx_editor.xml_element import XmlElement

T = TypeVar('T', bound='Attribute')

class AttributeRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}

    def register(self, namespace: str, name: str, element_names: list[str] | None, attribute_value_cls):
        if element_names is None:
            self._registry[(namespace, name, None)] = attribute_value_cls
        else:
            for element_name in element_names:
                self._registry[(namespace, name, element_name)] = attribute_value_cls

    def get_attribute_value_cls(self, namespace: str | None, name: str, element_name: str | None = None) -> type['Attribute']:
        return self._registry.get(
            (namespace, name, element_name),
            self._registry.get(
                (namespace, name, None),
                self._registry.get(
                    (namespace, None, None),
                    Attribute
                )
            )
        )

class AttributeProperty(Generic[T]):
    def __init__(self, attribute_type: type[T], nullable: bool = True):
        self.attribute_type = attribute_type
        self.nullable = nullable

    def __get__(self, instance: 'XmlElement | None', owner: type['XmlElement'], nullable_override: bool | None = None) -> T | None:
        if instance is None:
            raise AttributeError("AttributeProperty can only be accessed from an instance.")

        attribute = instance.get_attribute_by_type(self.attribute_type)

        if attribute is not None:
            return attribute

        if not self.nullable and not nullable_override:
            raise PowerpointIntegrityError(f"Expected an attribute of type {self.attribute_type.__name__} in {instance.__class__.__name__}, but none was found.")

        return None

    def __set__(self, instance: 'XmlElement', value: T | None) -> None:
        existing_attribute = self.__get__(instance, type(instance), nullable_override=True)

        if value is None and not self.nullable:
            raise PowerpointIntegrityError(f"Cannot set a non-nullable AttributeProperty to None in {instance.__class__.__name__}.")

        if value is not None and value.part is not instance.part:
            value = value.copy()
            value.part = instance.part

        if existing_attribute is not None:
            if value is not None:
                instance.replace_attribute(existing_attribute, value)
            else:
                instance.remove_attribute(existing_attribute)
        elif value is not None:
            instance.add_attribute(value)

class AttributeStringProperty(AttributeProperty[T]):
    def __get__(self, instance: 'XmlElement | None', owner: type['XmlElement'], nullable_override: bool | None = None) -> str | None:
        attribute = super().__get__(instance, owner, nullable_override)
        return attribute.value if attribute is not None else None

    def __set__(self, instance: 'XmlElement', value: str | None) -> None:
        existing_attribute = super().__get__(instance, type(instance), nullable_override=True)

        if value is None and not self.nullable:
            raise PowerpointIntegrityError(f"Cannot set a non-nullable AttributeStringProperty to None in {instance.__class__.__name__}.")

        if existing_attribute is not None:
            if value is not None:
                existing_attribute.value = value
            else:
                instance.remove_attribute(existing_attribute)
        elif value is not None:
            attribute_instance = self.attribute_type(value)
            instance.add_attribute(attribute_instance)

class BooleanAttributeProperty(AttributeProperty[T]):
    def __get__(self, instance: 'XmlElement | None', owner: type['XmlElement'], nullable_override: bool | None = None) -> bool | None:
        attribute = super().__get__(instance, owner, nullable_override)
        if attribute is not None:
            return attribute.value == '1'
        return None

    def __set__(self, instance: 'XmlElement', value: bool | None) -> None:
        existing_attribute = super().__get__(instance, type(instance), nullable_override=True)

        if value is None and not self.nullable:
            raise PowerpointIntegrityError(f"Cannot set a non-nullable BooleanAttributeProperty to None in {instance.__class__.__name__}.")

        if existing_attribute is not None:
            if value is not None:
                existing_attribute.value = '1' if value else '0'
            else:
                instance.remove_attribute(existing_attribute)
        elif value is not None:
            attribute_instance = self.attribute_type('1' if value else '0')
            instance.add_attribute(attribute_instance)

class Attribute:
    default_namespace: str | None = None
    default_prefix: str | None = None
    default_name: str | None = None
    default_element_name: str | None = None

    __slots__ = ['name', 'value', 'prefix']

    def __init__(self, value: str, prefix: str | None = None, name: str | None = None, overwrite_prefix: bool = False):
        self.prefix = prefix if prefix or overwrite_prefix else self.default_prefix
        self.name = sys.intern(name) if name else self.default_name
        self.value = value

    def copy(self):
        return self.__class__(self.value, self.prefix, self.name, overwrite_prefix=True)

    @classmethod
    def _from_item(cls, parser: '_OOXMLParser', file_path: PurePosixPath | None, namespaces: dict[str | None, str], name: str, value: str) -> 'Attribute':
        q = etree.QName(name)
        namespace = q.namespace if q.namespace else None
        prefix = [pfx for pfx, uri in namespaces.items() if uri == namespace][0] if namespace is not None else None
        attribute_value = cls(value, prefix, q.localname)

        return attribute_value

    def _to_xml(self, writer: '_OOXMLWriter', buffer: BytesIO):
        if self.prefix:
            qname = f"{self.prefix}:{self.name}"
        else:
            qname = self.name

        escaped_value = escape(self.value, entities={'"': '&quot;', "'": '&apos;', '\n': '&#10;', '\r': '&#13;', '\t': '&#9;'})
        buffer.write(f' {qname}="{escaped_value}"'.encode('utf-8'))

    @classmethod
    def _register(cls):
        registry = AttributeRegistry()

        registry.register(cls.default_namespace, cls.default_name, cls.default_element_name, cls)

    def __init_subclass__(cls, **kwargs) -> None:
        super().__init_subclass__(**kwargs)

        class_exceptions = []
        missing_namespace = not hasattr(cls, 'default_namespace')
        missing_prefix = not hasattr(cls, 'default_prefix')
        missing_name = not hasattr(cls, 'default_name')
        missing_element_name = not hasattr(cls, 'default_element_name')

        if cls.__name__ not in class_exceptions and missing_namespace:
            raise ValueError(f"Attribute subclass {cls.__name__} must define a default_namespace class attribute")
        if cls.__name__ not in class_exceptions and missing_prefix:
            raise ValueError(f"Attribute subclass {cls.__name__} must define a default_prefix class attribute")
        if cls.__name__ not in class_exceptions and missing_name:
            raise ValueError(f"Attribute subclass {cls.__name__} must define a default_name class attribute")
        if cls.__name__ not in class_exceptions and missing_element_name:
            raise ValueError(f"Attribute subclass {cls.__name__} must define a default_element_name class attribute")

        if cls.__name__ not in class_exceptions:
            cls._register()

    def __hash__(self):
        return hash((self.name, self.value, self.prefix))

    def __eq__(self, other):
        if not isinstance(other, self.__class__):
            return False

        if self is other:
            return True

        return self.name == other.name and self.value == other.value and self.prefix == other.prefix

    def __str__(self):
        escaped_value = escape(self.value, entities={'"': '&quot;', "'": '&apos;', '\n': '&#10;', '\r': '&#13;', '\t': '&#9;'})
        return f"{self.prefix}:{self.name}={escaped_value}" if self.prefix else f"{self.name}={escaped_value}"

    def __repr__(self):
        return self.__str__()
