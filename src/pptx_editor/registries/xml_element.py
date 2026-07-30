
from pptx_editor.singleton import SingletonMeta
from pptx_editor import xml_element


class XmlElementRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}
        self._name_registry = {}

    def register(self, namespace: str, name: str, element_cls: type['xml_element.XmlElement']):
        self._registry[(namespace, name)] = element_cls
        self._name_registry[element_cls.__name__] = element_cls


    def get_element_cls(self, namespace: str | None, name: str | None = None) -> type['xml_element.XmlElement']:
        return self._registry.get((namespace, name), xml_element.XmlElement)

    def get_element_cls_by_name(self, name: str) -> type['xml_element.XmlElement']:
        if name in self._name_registry:
            return self._name_registry[name]
        raise ValueError(f"XmlElement class with name '{name}' not found in registry.")
