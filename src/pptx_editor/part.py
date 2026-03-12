import sys

from lxml import etree

from pptx_editor.attribute import Attribute, AttributeValue
from pptx_editor.parser import Parser
from pptx_editor.relationship import Relationship
from pptx_editor.singleton import SingletonMeta

class PartRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}

    def register(self, content_type: str, part_cls):
        self._registry[content_type] = part_cls

    def get_part_cls(self, content_type: str) -> type[Part]:
        return self._registry.get(content_type, Part)

class Part():
    content_type: str

    def __init__(self, file_path: str, content_type: str):
        self.file_path = file_path
        self.content_type = sys.intern(content_type)
        self.relationships: list[Relationship] | None = None
        self.values: list[AttributeValue] | None = None
        self.attributes: list[Attribute] | None = None

    @classmethod
    def from_file(cls, parser: Parser, file_path: str, content_type: str):
        if parser.has_part(file_path):
            return parser.parts[file_path]

        part = cls(file_path, content_type)
        parser.add_part(part)
        part._parse_data(parser)
        return part

    def _parse_data(self, parser: Parser):
        file_xml = self._get_file_xml(parser)
        self._parse_xml(file_xml)

        if self._has_relationship_file(parser):
            relationship_file_xml = self._get_relationship_file_xml(parser)
            self._parse_relationship_xml(relationship_file_xml)

    def _parse_xml(self, xml: etree._Element):
        self.values = [AttributeValue(str(key), str(value)) for key, value in xml.attrib.items()]
        self.attributes = [Attribute(child) for child in xml]

    def _parse_relationship_xml(self, xml: etree._Element):
        relationships = []

        for relationship_xml in xml:
            relationship = Relationship.from_xml(relationship_xml)

            if relationship is None:
                continue

            relationships.append(relationship)

        self.relationships = relationships

    def _get_file_xml(self, parser: Parser):
        file_data_string = parser.read_file(self._get_file_path())
        file_xml = etree.fromstring(file_data_string)
        return file_xml

    def _get_file_path(self):
        return self.file_path.lstrip('/')

    def _get_relationship_file_xml(self, parser: Parser):
        relationship_file_data_string = parser.read_file(self._get_relationship_file_path())
        relationship_file_xml = etree.fromstring(relationship_file_data_string)
        return relationship_file_xml

    def _get_relationship_file_path(self) -> str:
        path_elements = self._get_file_path().split('/')
        return '/'.join(path_elements[:-1]) + '/_rels/' + path_elements[-1] + '.rels'

    def _has_relationship_file(self, parser: Parser) -> bool:
        relationship_file_path = self._get_relationship_file_path()
        return relationship_file_path in parser.zip_file.namelist()

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
