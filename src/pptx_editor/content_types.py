from lxml.etree import _Element
from lxml import etree

from pptx_editor.parser import Parser
from pptx_editor.part import PartRegistry, Part

class ContentTypes():
    def __init__(self):
        self.parts = []

    @staticmethod
    def from_file(parser: Parser):
        content_types = ContentTypes()
        content_types._parse_content_types(parser)
        return content_types

    def _parse_content_types(self, parser: Parser):
        content_types = parser.read_file('[Content_Types].xml')
        root_element = etree.fromstring(content_types)
        self.parts = self._parse_content_types_tree(parser, root_element)
        self._parse_content_data(parser)

    def _parse_content_types_tree(self, parser: Parser, tree: _Element) -> list[Part]:
        parts = []

        for child in tree:
            part = self._parse_content_type(parser, child)

            if not part:
                continue

            parts.append(part)

        return parts

    def _parse_content_type(self, parser: Parser, content_type: _Element) -> Part | None:
        q = etree.QName(content_type)

        part = None

        if q.localname == 'Default':
            part = self._parse_default_content_type(parser, content_type)
        elif q.localname == 'Override':
            part = self._parse_override_content_type(parser, content_type)

        return part

    def _parse_default_content_type(self, parser: Parser, content_type: _Element):
        pass

    def _parse_override_content_type(self, parser: Parser, content_type: _Element) -> Part | None:
        content_type_value = content_type.get('ContentType')
        content_part_name = content_type.get('PartName')

        if content_type_value is None:
            print("Integrity warning: Override element missing ContentType attribute")
            return None

        part_cls = PartRegistry().get_part_cls(content_type_value)

        part = part_cls.from_file(parser, content_part_name, content_type_value)

        return part

    def _parse_content_data(self, parser: Parser):
        for part in self.parts:
            part._parse_data(parser)
