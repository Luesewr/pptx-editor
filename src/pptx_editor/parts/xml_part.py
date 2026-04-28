from pathlib import PurePosixPath
from typing import TYPE_CHECKING

from lxml import etree

from pptx_editor.part import Part
from pptx_editor.relationship import Relationship
from pptx_editor.attribute import Attribute

if TYPE_CHECKING:
    from pptx_editor.parts.base import Base
    from pptx_editor.parser import Parser
    from pptx_editor.writer import Writer

class XmlPart(Part):
    default_content_type: str | None = "application/xml"
    default_base_path: PurePosixPath | None
    default_part_name: str | None
    default_attribute_name: str | None = None

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.data: Attribute | None = None

    def to_file(self, writer: 'Writer'):
        if writer.is_part_written(self):
            return

        body_relationships = self.data.get_relationships() if self.data else []
        self.relationships = list(dict.fromkeys(body_relationships + self.relationships))

        file_name = writer.assign_part_index(self.part_name, self)
        file_path = PurePosixPath(self.base_path) / file_name if self.base_path else file_name

        if not file_path.is_absolute():
            file_path = PurePosixPath('/') / file_path

        writer.assign_relationship_ids(self, self.relationships)
        writer.assign_relation_part_indexes(self.relationships)

        # Write the part's XML content to the zip file
        if self.data is not None:
            part_xml = self.data.to_xml(writer, {})
            part_xml_string = etree.tostring(part_xml, encoding='utf-8', xml_declaration=True, standalone=True)
            writer.write_file(file_path, part_xml_string)

        writer.add_written_part(self)

        if len(self.relationships) > 0:
            self._relationships_to_xml(writer)

    def _parse_data(self, parser: 'Parser', file_path: PurePosixPath | None = None):
        if self._has_relationship_file(parser, file_path):
            self._parse_relationships(parser, file_path)

        file_xml = self._get_file_xml(parser, file_path)

        if file_xml is not None:
            self._parse_xml(parser, file_path, file_xml)

    def _parse_xml(self, parser: 'Parser', file_path: PurePosixPath | None, xml: etree._Element):
        self.data = Attribute.from_xml(parser, file_path, xml)

    def _get_file_xml(self, parser: 'Parser', file_path: PurePosixPath | None):
        file_path = self._get_file_path() if file_path is None else file_path

        if not file_path:
            return None

        file_data_string = parser.read_file(file_path)
        file_xml = etree.fromstring(file_data_string)
        return file_xml
