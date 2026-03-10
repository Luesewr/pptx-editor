import sys

from zipfile import ZipFile

from lxml import etree

from pptx_editor.attribute import Attribute
from pptx_editor.singleton import SingletonMeta

class PartRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}

    def register(self, content_type: str, part_cls):
        self._registry[content_type] = part_cls

    def get_part_cls(self, content_type: str):
        return self._registry.get(content_type, Part)

class Part():
    content_type: str

    def __init__(self, file_path: str, content_type: str):
        self.file_path = file_path
        self.content_type = sys.intern(content_type)

    def _parse_data(self, zip_file: ZipFile):
        file_xml = self._get_file_xml(zip_file)
        self._parse_xml(file_xml)

    def _parse_xml(self, xml):
        self.attributes = dict(xml.attrib)
        self.children = [Attribute(child) for child in xml]

    def _get_file_xml(self, zip_file: ZipFile):
        file_data_string = zip_file.read(self.file_path.lstrip('/'))
        file_xml = etree.fromstring(file_data_string)
        return file_xml

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
        return f"{self.content_type.split('.')[-1].removesuffix('+xml')}(file_path={self.file_path})"

    def __repr__(self) -> str:
        return self.__str__()
