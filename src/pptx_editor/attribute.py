from io import BytesIO
from pathlib import PurePosixPath
import sys
from xml.sax.saxutils import escape

from lxml import etree
from typing import TYPE_CHECKING

from pptx_editor.singleton import SingletonMeta

if TYPE_CHECKING:
    from pptx_editor.parser import _OOXMLParser
    from pptx_editor.writer import _OOXMLWriter

class AttributeRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}

    def register(self, namespace: str, attribute_value_cls):
        self._registry[namespace] = attribute_value_cls

    def get_attribute_value_cls(self, namespace: str | None) -> type['Attribute']:
        return self._registry.get(namespace, Attribute)

class Attribute:
    default_namespace: str | None = None

    __slots__ = ['name', 'value', 'prefix']

    def __init__(self, name: str, value: str, prefix: str | None = None):
        self.prefix = prefix if prefix else None
        self.name = sys.intern(name)
        self.value = value

    def copy(self):
        return self.__class__(self.name, self.value, self.prefix)

    @classmethod
    def _from_item(cls, parser: '_OOXMLParser', file_path: PurePosixPath | None, namespaces: dict[str | None, str], name: str, value: str) -> 'Attribute':
        q = etree.QName(name)
        namespace = q.namespace if q.namespace else None
        prefix = [pfx for pfx, uri in namespaces.items() if uri == namespace][0] if namespace is not None else None
        attribute_value = cls(q.localname, value, prefix)

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

        if cls.default_namespace is not None:
            registry.register(cls.default_namespace, cls)

    def __init_subclass__(cls, **kwargs) -> None:
        super().__init_subclass__(**kwargs)

        class_exceptions = []
        missing_namespace = not hasattr(cls, 'default_namespace') or cls.default_namespace is None

        if cls.__name__ not in class_exceptions and missing_namespace:
            raise ValueError(f"Attribute subclass {cls.__name__} must define a default_namespace class attribute")

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
