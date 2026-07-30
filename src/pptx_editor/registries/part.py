from pptx_editor import part
from pptx_editor.singleton import SingletonMeta

class PartRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}
        self._name_registry = {}

    def register(self, content_type: str, part_cls: type['part.Part']):
        self._registry[content_type] = part_cls
        self._name_registry[part_cls.__name__] = part_cls

    def get_part_cls(self, content_type: str) -> type['part.Part']:
        if content_type in self._registry:
            return self._registry[content_type]
        if '+xml' in content_type:
            return self.get_part_cls('application/xml')
        return part.Part

    def get_part_cls_by_name(self, name: str) -> type['part.Part']:
        if name in self._name_registry:
            return self._name_registry[name]
        raise ValueError(f"Part class with name '{name}' not found in registry.")