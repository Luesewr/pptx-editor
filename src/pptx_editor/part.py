import re

from typing import TYPE_CHECKING

from lxml import etree

from pptx_editor.attribute import Attribute, AttributeValue
from pptx_editor.relationship import Relationship
from pptx_editor.singleton import SingletonMeta

if TYPE_CHECKING:
    from pptx_editor.parser import Parser

class PartRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}

    def register(self, content_type: str, part_cls):
        self._registry[content_type] = part_cls

    def get_part_cls(self, content_type: str) -> type['Part']:
        return self._registry.get(content_type, Part)

class Part():
    content_type: str
    default_base_path: str
    default_part_name: str

    def __init__(self, file_path: str | None = None, content_type: str | None = None):
        self.original_file_path = file_path

        if file_path is None:
            self.base_path = self.default_base_path
            self.part_name = self.default_part_name
        else:
            self.base_path = file_path.rsplit('/', 1)[0]
            self.part_name = file_path.rsplit('/', 1)[1] if '/' in file_path else file_path

        if (m := re.match(r'(^.*?)\d+(\.xml)$', self.part_name)):
            self.part_name = m.group(1) + '{i}' + m.group(2)

        self.relationships: list[Relationship] | None = None
        self.values: list[AttributeValue] | None = None
        self.attributes: list[Attribute] | None = None

        if not hasattr(self, 'content_type') and content_type:
            self.content_type = content_type

    @classmethod
    def from_file(cls, parser: 'Parser', file_path: str, content_type: str):
        if parser.has_part(file_path):
            return parser.get_part(file_path)

        part = cls(file_path, content_type)

        parser.add_part(part)
        part._parse_data(parser)
        return part

    def _parse_data(self, parser: 'Parser'):
        if self._has_relationship_file(parser):
            self._parse_relationships(parser)

        file_xml = self._get_file_xml(parser)
        self._parse_xml(file_xml)


    def _parse_xml(self, xml: etree._Element):
        self.values = [AttributeValue(str(key), str(value)) for key, value in xml.attrib.items()]
        self.attributes = [Attribute(child) for child in xml]

    def _parse_relationships(self, parser: 'Parser'):
        relationship_file_path = self._get_relationship_file_path(original=True)
        relationships = Relationship.from_file(parser, relationship_file_path)
        self.relationships = relationships
        parser.parse_relationship_targets(relationships)

    def _get_file_xml(self, parser: 'Parser'):
        file_data_string = parser.read_file(self._get_file_path(original=True).lstrip('/'))
        file_xml = etree.fromstring(file_data_string)
        return file_xml

    def _get_file_path(self, original: bool = False) -> str:
        file_path = self.original_file_path if original and self.original_file_path else None

        if file_path is None:
            if self.base_path:
                file_path = f"{self.base_path}/{self.part_name}"
            else:
                file_path = f"{self.part_name}"

        if not file_path.startswith('/'):
            file_path = '/' + file_path

        return file_path

    def _get_relationship_file_path(self, original: bool = False) -> str:
        file_path = self._get_file_path(original=original)
        path_elements = file_path.lstrip('/').split('/')
        return '/'.join(path_elements[:-1]) + '/_rels/' + path_elements[-1] + '.rels'

    def _has_relationship_file(self, parser: 'Parser') -> bool:
        relationship_file_path = self._get_relationship_file_path(original=True)
        return relationship_file_path in parser.zip_file.namelist()

    @classmethod
    def _register(cls):
        registry = PartRegistry()
        registry.register(cls.content_type, cls)

    def __init_subclass__(cls, **kwargs) -> None:
        super().__init_subclass__(**kwargs)

        class_exceptions = ['ReturnPart']
        missing_content_type = not hasattr(cls, 'content_type') or cls.content_type is None
        missing_default_base_path = not hasattr(cls, 'default_base_path') or cls.default_base_path is None
        missing_default_part_name = not hasattr(cls, 'default_part_name') or cls.default_part_name is None

        if cls.__name__ not in class_exceptions and (missing_content_type or missing_default_base_path or missing_default_part_name):
            raise ValueError(f"Part subclass {cls.__name__} must define a content_type class attribute, a default_base_path class attribute, and a default_part_name class attribute")

        if cls.__name__ not in class_exceptions:
            cls._register()

    def __str__(self) -> str:
        return f"{self.content_type.split('.')[-1].removesuffix('+xml')}(file_path={self._get_file_path()})"

    def __repr__(self) -> str:
        return self.__str__()

class ReturnPart(Part):
    def __init__(self, file_path: str, content_type: str):
        super().__init__(file_path, content_type)
        self.main_relationships: list['Relationship'] | None = None
