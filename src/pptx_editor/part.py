from io import BytesIO
import re
import sys

from pathlib import PurePosixPath
from typing import TYPE_CHECKING, Any

from pptx_editor.relationship import Relationship
from pptx_editor.singleton import SingletonMeta

if TYPE_CHECKING:
    from pptx_editor.parser import _OOXMLParser
    from pptx_editor.parts.package import Package
    from pptx_editor.writer import _OOXMLWriter

class PartRegistry(metaclass=SingletonMeta):
    def __init__(self):
        self._registry = {}

    def register(self, content_type: str, part_cls):
        self._registry[content_type] = part_cls

    def get_part_cls(self, content_type: str) -> type['Part']:
        if content_type in self._registry:
            return self._registry[content_type]
        if '+xml' in content_type:
            return self.get_part_cls('application/xml')
        return Part

class Part():
    default_content_type: str | None = None
    default_base_path: PurePosixPath | None
    default_part_name: str | None
    default_attribute_name: str | None = None

    def __init__(self, package: 'Package', file_path: PurePosixPath | None = None, content_type: str | None = None, is_default: bool = False):
        package._add_part(self)
        self.package = package

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
    def _from_file(cls, parser: '_OOXMLParser', file_path: PurePosixPath, content_type: str | None, is_default: bool = False) -> 'Part':
        if file_path and (part := parser.get_part(file_path)):
            return part

        part = cls(parser.package, file_path, content_type, is_default)

        parser.add_part(file_path, part)
        part._parse_data(parser, file_path)
        return part

    def _to_file(self, writer: '_OOXMLWriter'):
        if writer.is_part_written(self) or self.part_name is None:
            return

        file_name = writer.assign_part_index(self.part_name, self)

        file_path = PurePosixPath(self.base_path) / file_name if self.base_path else file_name

        if not file_path.is_absolute():
            file_path = PurePosixPath('/') / file_path

        # Write the part's XML content to the zip file
        if self.data is not None:
            writer.write_file(file_path, self.data)

        writer.add_written_part(self)

        if len(self.relationships) > 0:
            self._write_relationships_file(writer)

    def _parse_data(self, parser: '_OOXMLParser', file_path: PurePosixPath):
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

    def _parse_relationships(self, parser: '_OOXMLParser', file_path: PurePosixPath):
        relationship_file_path = self._get_relationship_file_path(file_path=file_path)
        relationships = Relationship._from_file(parser, relationship_file_path, self)
        self.relationships = relationships

    def _get_relationship_file_path(self, file_path: PurePosixPath) -> PurePosixPath:
        file_path = (self._get_file_path() if file_path is None else file_path) or PurePosixPath('')
        path = file_path.parent / '_rels' / (file_path.name + '.rels')

        if not path.is_absolute():
            path = PurePosixPath('/') / path

        return path

    def _has_relationship_file(self, parser: '_OOXMLParser', file_path: PurePosixPath) -> bool:
        relationship_file_path = self._get_relationship_file_path(file_path=file_path)

        if relationship_file_path.is_absolute():
            relationship_file_path = relationship_file_path.relative_to(relationship_file_path.anchor)

        return relationship_file_path.as_posix() in parser.zip_file.namelist()

    def _write_relationships_file(self, writer: '_OOXMLWriter'):
        buffer = BytesIO()
        buffer.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'.encode('utf-8'))
        buffer.write('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'.encode('utf-8'))

        for relationship in self.relationships:
            relationship._to_xml(writer, buffer)

        buffer.write('</Relationships>'.encode('utf-8'))

        file_name = writer.assign_part_index(self.part_name, self)

        file_path = PurePosixPath(self.base_path) / PurePosixPath(file_name) if self.base_path else PurePosixPath(file_name)

        if file_path and not file_path.is_absolute():
            file_path = PurePosixPath('/') / file_path

        relationship_file_path = self._get_relationship_file_path(file_path=file_path)
        writer.write_file(relationship_file_path, buffer.getvalue())

        for relationship in self.relationships:
            relationship.target._to_file(writer)

    @classmethod
    def _register(cls):
        registry = PartRegistry()

        if cls.default_content_type is not None:
            registry.register(cls.default_content_type, cls)

    def __init_subclass__(cls, **kwargs) -> None:
        super().__init_subclass__(**kwargs)

        class_exceptions = ['Package', 'XmlPart']
        missing_content_type = not hasattr(cls, 'default_content_type') or cls.default_content_type is None
        missing_default_base_path = not hasattr(cls, 'default_base_path') or cls.default_base_path is None
        missing_default_part_name = not hasattr(cls, 'default_part_name') or cls.default_part_name is None

        if cls.__name__ not in class_exceptions and (missing_content_type or missing_default_base_path or missing_default_part_name):
            raise ValueError(f"Part subclass {cls.__name__} must define a default_content_type class attribute, a default_base_path class attribute, and a default_part_name class attribute")

        if cls.__name__ not in class_exceptions:
            cls._register()

        if cls.__name__ == 'XmlPart':
            cls._register()

    def __str__(self) -> str:
        return f"{(self.default_content_type or 'package').split('.')[-1].removesuffix('+xml')}(file_path={self._get_file_path()})"

    def __repr__(self) -> str:
        return self.__str__()
