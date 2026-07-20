
from pptx_editor.singleton import SingletonMeta
from pptx_editor import xml_element


class XmlElementRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}

    def register(self, namespace: str, name: str, element_cls: type['xml_element.XmlElement']):
        self._registry[(namespace, name)] = element_cls

    def get_element_cls(self, namespace: str | None, name: str | None = None) -> type['xml_element.XmlElement']:
        return self._registry.get((namespace, name), xml_element.XmlElement)
