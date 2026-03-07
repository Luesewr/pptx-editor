from abc import ABC, abstractmethod

from pptx_editor.singleton import SingletonMeta

class PartRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}
    
    def register(self, content_type: str, part_cls):
        self._registry[content_type] = part_cls
    
    def get_part_cls(self, content_type: str):
        return self._registry.get(content_type)

class Part(ABC):
    content_type: str
    
    def __init__(self, file_path: str):
        self.file_path = file_path

    def _parse_data(self):
        pass

    @classmethod
    def _register(cls):
        registry = PartRegistry()
        registry.register(cls.content_type, cls)

    def __init_subclass__(cls, **kwargs) -> None:
        super().__init_subclass__(**kwargs)

        if cls.content_type is None:
            raise ValueError(f"Part subclass {cls.__name__} must define a content_type class attribute")

        cls._register()

    def __str__(self) -> str:
        return f"{self.__class__.__name__}(file_path={self.file_path})"
