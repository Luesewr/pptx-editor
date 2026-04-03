from pathlib import PurePosixPath
from typing import TYPE_CHECKING

from lxml.etree import _Element
from lxml import etree

if TYPE_CHECKING:
    from pptx_editor.parser import Parser
    from pptx_editor.writer import Writer

class ContentTypes():
    def __init__(self):
        self.defaults = {}
        self.overrides = {}

    @staticmethod
    def from_file(parser: 'Parser', file_path: PurePosixPath = PurePosixPath('[Content_Types].xml')) -> 'ContentTypes':
        content_types = ContentTypes()
        content_types._parse_content_types(parser, file_path)
        return content_types

    def to_file(self, writer: 'Writer'):
        root = etree.Element('Types', xmlns="http://schemas.openxmlformats.org/package/2006/content-types")

        for extension, content_type in self.defaults.items():
            default_element = etree.SubElement(root, 'Default')
            default_element.set('Extension', extension)
            default_element.set('ContentType', content_type)

        for part_name, content_type in self.overrides.items():
            override_element = etree.SubElement(root, 'Override')
            override_element.set('PartName', part_name.as_posix())
            override_element.set('ContentType', content_type)

        content_types_xml_string = etree.tostring(root, encoding='utf-8', xml_declaration=True)
        writer.write_file(PurePosixPath('[Content_Types].xml'), content_types_xml_string)

    def get_override_content_type(self, file_path: PurePosixPath) -> str | None:
        if file_path in self.overrides:
            return self.overrides[file_path]

        return None

    def get_default_content_type(self, file_path: PurePosixPath) -> str | None:
        extension = file_path.suffix.lstrip('.')

        if extension in self.defaults:
            return self.defaults[extension]

        return None

    def _parse_content_types(self, parser: 'Parser', file_path: PurePosixPath):
        content_types = parser.read_file(file_path)
        root_element = etree.fromstring(content_types)
        self._parse_content_types_tree(root_element)

    def _parse_content_types_tree(self, tree: _Element):
        for child in tree:
            self._parse_content_type(child)

    def _parse_content_type(self, content_type: _Element):
        q = etree.QName(content_type)

        if q.localname == 'Default':
            self._parse_default_content_type(content_type)
        elif q.localname == 'Override':
            self._parse_override_content_type(content_type)

    def _parse_default_content_type(self, content_type: _Element):
        content_type_value = content_type.get('ContentType')
        extension = content_type.get('Extension')

        if content_type_value is None:
            print("Integrity warning: Default element missing ContentType attribute")
            return

        if extension is None:
            print("Integrity warning: Default element missing Extension attribute")
            return

        self.defaults[extension] = content_type_value

    def _parse_override_content_type(self, content_type: _Element):
        content_type_value = content_type.get('ContentType')
        content_part_name = content_type.get('PartName')

        if content_type_value is None:
            print("Integrity warning: Override element missing ContentType attribute")
            return

        if content_part_name is None:
            print("Integrity warning: Override element missing PartName attribute")
            return

        self.overrides[PurePosixPath(content_part_name)] = content_type_value
