from io import BytesIO
from typing import TYPE_CHECKING, TypeVar
from xml.sax.saxutils import escape

from lxml import etree

from pptx_editor.registries.attribute import AttributeRegistry

if TYPE_CHECKING:
    from pptx_editor.writer import _OOXMLWriter
    from pptx_editor.parts.xml_part import XmlPart

T = TypeVar('T', bound='Attribute')

class Attribute:
    default_namespace: str | None = None
    default_prefix: str | None = None
    default_name: str | None = None
    default_element_name: str | None = None

    __slots__ = ['name', 'value', 'prefix']

    def __init__(self, value: str, prefix: str | None = None, name: str | None = None, overwrite_prefix: bool = False):
        self.prefix = prefix if prefix or overwrite_prefix else self.default_prefix
        self.name = name if name else self.default_name
        self.value = value

    def copy(self):
        return self.__class__(self.value, self.prefix, self.name, overwrite_prefix=True)

    @classmethod
    def _from_item(cls, part: 'XmlPart', namespaces: dict[str | None, str], name: str, value: str) -> 'Attribute':
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
