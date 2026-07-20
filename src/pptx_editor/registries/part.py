from pptx_editor import part
from pptx_editor.singleton import SingletonMeta

class PartRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}

    def register(self, content_type: str, part_cls):
        self._registry[content_type] = part_cls

    def get_part_cls(self, content_type: str) -> type['part.Part']:
        if content_type in self._registry:
            return self._registry[content_type]
        if '+xml' in content_type:
            return self.get_part_cls('application/xml')
        return part.Part