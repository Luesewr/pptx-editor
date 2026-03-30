import re

import sys
from typing import TYPE_CHECKING, Any

from lxml import etree

from pptx_editor.attribute import Attribute, AttributeValue
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
    default_base_path: str
    default_part_name: str | None
    default_attribute_name: str | None = None

    def __init__(self, base: 'Base | None', file_path: str | None = None, content_type: str | None = None):
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
            self.base_path = sys.intern(self.default_base_path)
            self.part_name = sys.intern(self.default_part_name) if self.default_part_name else None
        else:
            self.base_path = sys.intern(file_path.rsplit('/', 1)[0])
            self.part_name = sys.intern(file_path.rsplit('/', 1)[1] if '/' in file_path else sys.intern(file_path))

        if self.part_name and (m := re.match(r'(^.*?)\d+(\.xml)$', self.part_name)):
            self.part_name = sys.intern(m.group(1) + '{i}' + m.group(2))

        self.relationships: list[Relationship] = []
        self.data: Any | None = None

        self.content_type: str | None = None

        if content_type is not None:
            self.content_type = sys.intern(content_type)
        elif self.default_content_type is not None:
            self.content_type = sys.intern(self.default_content_type)

    @classmethod
    def from_file(cls, base: 'Base | None', parser: 'Parser', file_path: str | None, content_type: str | None):
        if file_path and parser.has_part(file_path):
            return parser.get_part(file_path)

        part = cls(base, file_path, content_type)

        if base is None:
            parser.base = part

        parser.add_part(file_path, part)
        part._parse_data(parser, file_path)
        return part

    def to_file(self, writer: 'Writer'):
        if writer.is_part_written(self):
            return

        if writer.has_part_index(self.part_name, self):
            file_name = writer.get_part_index(self.part_name, self)
        else:
            file_name = self.part_name
            # if file_name is None:
            #     raise PowerpointIntegrityError("Integrity warning: Part has no file path and cannot be written to file")

        file_path = f"{self.base_path}/{file_name}" if self.base_path else file_name

        if file_path and not file_path.startswith('/'):
            file_path = '/' + file_path

        # Write the part's XML content to the zip file
        if file_path is not None and self.data is not None:
            writer.write_file(file_path, self.data)

        writer.add_written_part(self)

    def _parse_data(self, parser: 'Parser', file_path: str | None = None):
        file_path = self._get_file_path() if file_path is None else file_path

        if not file_path:
            return None

        file_data = parser.read_file(file_path.lstrip('/'))
        self.data = file_data

    def _get_file_path(self) -> str:
        file_path = ''

        if self.base_path and self.part_name is not None:
            file_path = f"{self.base_path}/{self.part_name}"
        elif self.part_name is not None:
            file_path = f"{self.part_name}"

        if not file_path.startswith('/'):
            file_path = '/' + file_path

        return file_path

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
