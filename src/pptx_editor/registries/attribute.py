
from pptx_editor import attribute
from pptx_editor.singleton import SingletonMeta

class AttributeRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}

    def register(self, namespace: str, name: str, element_names: list[str] | None, attribute_value_cls):
        if element_names is None:
            self._registry[(namespace, name, None)] = attribute_value_cls
        else:
            for element_name in element_names:
                self._registry[(namespace, name, element_name)] = attribute_value_cls

    def get_attribute_value_cls(self, namespace: str | None, name: str, element_name: str) -> type['attribute.Attribute']:
        exact_match = self._registry.get((namespace, name, element_name))

        if exact_match is not None:
            return exact_match

        name_match = self._registry.get((namespace, name, None))
        if name_match is not None:
            self.register(namespace, name, [element_name], name_match)
            return name_match

        namespace_match = self._registry.get((namespace, None, None))
        if namespace_match is not None:
            self.register(namespace, name, None, namespace_match)
            self.register(namespace, name, [element_name], namespace_match)
            return namespace_match

        self.register(namespace, name, [element_name], attribute.Attribute)
        self.register(namespace, name, None, attribute.Attribute)
        self.register(namespace, None, None, attribute.Attribute)

        return attribute.Attribute
