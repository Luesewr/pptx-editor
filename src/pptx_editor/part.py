import re
import sys

from pathlib import PurePosixPath
from typing import TYPE_CHECKING, Any

from lxml import etree

from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.relationship import Relationship
from pptx_editor.singleton import SingletonMeta

if TYPE_CHECKING:
    from pptx_editor.parser import Parser
    from pptx_editor.parts.base import Base
    from pptx_editor.writer import Writer

class PartRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}

    def register(self, content_type: str, part_cls):
        self._registry[content_type] = part_cls

    def get_part_cls(self, content_type: str) -> type['Part']:
        return self._registry.get(content_type, Part)

class Part():
    default_content_type: str | None = None
    default_base_path: PurePosixPath | None
    default_part_name: str | None
    default_attribute_name: str | None = None

    def __init__(self, base: 'Base | None', file_path: PurePosixPath | None = None, content_type: str | None = None, is_default: bool = False):
        from pptx_editor.parts.base import Base

        if base is not None:
            base.add_part(self)
            self.base: 'Base' = base
        elif isinstance(self, Base):
            self.base = self
        else:
            raise PowerpointIntegrityError("Integrity warning: Base part must be the root part of the presentation and cannot have a parent part")

        self.part_name: str | None = None

        if file_path is None:
            self.base_path: PurePosixPath | None = self.default_base_path if self.default_base_path else None
            self.part_name: str | None = sys.intern(self.default_part_name) if self.default_part_name else None
        else:
            self.base_path: PurePosixPath | None = file_path.parent
            self.part_name: str | None = sys.intern(file_path.name)

        if self.part_name and (m := re.match(r'(^.*?)\d+(\.xml)$', self.part_name)):
            self.part_name: str | None = sys.intern(m.group(1) + '{i}' + m.group(2))

        self.relationships: list[Relationship] = []
        self.data: Any | None = None

        self.content_type: str | None = None

        if content_type is not None:
            self.content_type: str | None = sys.intern(content_type)
        elif self.default_content_type is not None:
            self.content_type: str | None = sys.intern(self.default_content_type)

        self.is_default = is_default

        self.relationships: list[Relationship] = []

    @classmethod
    def from_file(cls, base: 'Base | None', parser: 'Parser', file_path: PurePosixPath | None, content_type: str | None, is_default: bool = False) -> 'Part':
        if file_path and (part := parser.get_part(file_path)):
            return part

        part = cls(base, file_path, content_type, is_default)

        from pptx_editor.parts.base import Base

        if base is None and isinstance(part, Base):
            parser.base = part

        parser.add_part(file_path, part)
        part._parse_data(parser, file_path)
        return part

    def to_file(self, writer: 'Writer'):
        if writer.is_part_written(self) or self.part_name is None:
            return

        if writer.has_part_index(self.part_name, self):
            file_name = writer.get_part_index(self.part_name, self)
        else:
            file_name = PurePosixPath(self.part_name)

        file_path = PurePosixPath(self.base_path) / file_name if self.base_path and file_name else file_name

        if file_path and not file_path.is_absolute():
            file_path = PurePosixPath('/') / file_path

        # Write the part's XML content to the zip file
        if file_path is not None and self.data is not None:
            writer.write_file(file_path, self.data)

        writer.add_written_part(self)

        relationships = self.relationships

        self._relationships_to_xml(writer, relationships)

        for relationship in relationships:
            relationship.target.to_file(writer)

    def _parse_data(self, parser: 'Parser', file_path: PurePosixPath | None = None):
        file_path = self._get_file_path() if file_path is None else file_path

        if not file_path:
            return None

        if self._has_relationship_file(parser, file_path):
            self._parse_relationships(parser, file_path)

        file_data = parser.read_file(file_path)
        self.data = file_data

    def _get_file_path(self) -> PurePosixPath | None:
        file_path = None

        if self.base_path and self.part_name is not None:
            file_path = PurePosixPath(self.base_path) / self.part_name
        elif self.part_name is not None:
            file_path = PurePosixPath(self.part_name)

        if file_path is not None and not file_path.is_absolute():
            file_path = PurePosixPath('/') / file_path

        return file_path

    def _parse_relationships(self, parser: 'Parser', file_path: PurePosixPath | None):
        relationship_file_path = self._get_relationship_file_path(file_path=file_path)
        relationships = Relationship.from_file(parser, relationship_file_path, self)
        self.relationships = relationships

    def _get_relationship_file_path(self, file_path: PurePosixPath | None = None) -> PurePosixPath:
        file_path = (self._get_file_path() if file_path is None else file_path) or PurePosixPath('')
        path = file_path.parent / '_rels' / (file_path.name + '.rels')

        if not path.is_absolute():
            path = PurePosixPath('/') / path

        return path

    def _has_relationship_file(self, parser: 'Parser', file_path: PurePosixPath | None) -> bool:
        relationship_file_path = self._get_relationship_file_path(file_path=file_path)

        if relationship_file_path.is_absolute():
            relationship_file_path = relationship_file_path.relative_to(relationship_file_path.anchor)

        return relationship_file_path.as_posix() in parser.zip_file.namelist()

    def _relationships_to_xml(self, writer: 'Writer', relationships: list['Relationship']):
        relationships_element = etree.Element('Relationships', xmlns="http://schemas.openxmlformats.org/package/2006/relationships")

        for relationship in relationships:
            relationship_xml = relationship._to_xml(writer)
            relationships_element.append(relationship_xml)

        relationship_xml_string = etree.tostring(relationships_element, encoding='utf-8', xml_declaration=True)

        file_name = writer.assign_part_index(self.part_name, self)

        file_path = PurePosixPath(self.base_path) / PurePosixPath(file_name) if self.base_path else PurePosixPath(file_name)

        if file_path and not file_path.is_absolute():
            file_path = PurePosixPath('/') / file_path

        relationship_file_path = self._get_relationship_file_path(file_path=file_path)
        writer.write_file(relationship_file_path, relationship_xml_string)

    @classmethod
    def _register(cls):
        registry = PartRegistry()

        if cls.default_content_type is not None:
            registry.register(cls.default_content_type, cls)

    def __init_subclass__(cls, **kwargs) -> None:
        super().__init_subclass__(**kwargs)

        class_exceptions = ['Base', 'XmlPart']
        missing_content_type = not hasattr(cls, 'default_content_type') or cls.default_content_type is None
        missing_default_base_path = not hasattr(cls, 'default_base_path') or cls.default_base_path is None
        missing_default_part_name = not hasattr(cls, 'default_part_name') or cls.default_part_name is None

        if cls.__name__ not in class_exceptions and (missing_content_type or missing_default_base_path or missing_default_part_name):
            raise ValueError(f"Part subclass {cls.__name__} must define a default_content_type class attribute, a default_base_path class attribute, and a default_part_name class attribute")

        if cls.__name__ not in class_exceptions:
            cls._register()

    def __str__(self) -> str:
        return f"{(self.default_content_type or 'base').split('.')[-1].removesuffix('+xml')}(file_path={self._get_file_path()})"

    def __repr__(self) -> str:
        return self.__str__()
