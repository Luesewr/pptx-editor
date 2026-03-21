import sys

from lxml import etree
from typing import TYPE_CHECKING

from pptx_editor.singleton import SingletonMeta

if TYPE_CHECKING:
    from pptx_editor.parser import Parser

class AttributeValueRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}

    def register(self, namespace: str, attribute_value_cls):
        self._registry[namespace] = attribute_value_cls

    def get_attribute_value_cls(self, namespace: str | None) -> type['AttributeValue']:
        return self._registry.get(namespace, AttributeValue)

class AttributeValue:
    default_namespace: str | None = None

    __slots__ = ['name', 'namespace', 'value']

    def __init__(self, name: str, value: str):
        q = etree.QName(name)
        self.namespace = sys.intern(q.namespace) if q.namespace else None
        self.name = sys.intern(q.localname)
        self.value = sys.intern(value)

    @classmethod
    def from_item(cls, parser: 'Parser', file_path: str | None, name: str, value: str) -> 'AttributeValue':
        attribute_value = cls(name, value)

        return attribute_value

    @classmethod
    def _register(cls):
        registry = AttributeValueRegistry()

        if cls.default_namespace is not None:
            registry.register(cls.default_namespace, cls)

    def __init_subclass__(cls, **kwargs) -> None:
        super().__init_subclass__(**kwargs)

        class_exceptions = []
        missing_namespace = not hasattr(cls, 'default_namespace') or cls.default_namespace is None

        if cls.__name__ not in class_exceptions and missing_namespace:
            raise ValueError(f"AttributeValue subclass {cls.__name__} must define a default_namespace class attribute")

        if cls.__name__ not in class_exceptions:
            cls._register()

    def __str__(self):
        return f"{self.namespace}:{self.name}={self.value}" if self.namespace else f"{self.name}={self.value}"

    def __repr__(self):
        return self.__str__()
