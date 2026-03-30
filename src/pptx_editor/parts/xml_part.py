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
    default_base_path: str
    default_part_name: str | None
    default_attribute_name: str | None = None

    def __init__(self, base: 'Base | None', file_path: str | None = None, content_type: str | None = None):
        super().__init__(base, file_path, content_type)

        self.relationships: list[Relationship] = []
        self.data: Attribute | None = None

    def to_file(self, writer: 'Writer'):
        if writer.is_part_written(self):
            return

        body_relationships = self.data.get_relationships() if self.data else []
        relationships = list(dict.fromkeys(body_relationships + self.relationships))

        writer.assign_relationship_ids(self, relationships)
        writer.assign_part_indexes(relationships)

        file_name = writer.assign_part_index(self.part_name, self)

        file_path = f"{self.base_path}/{file_name}" if self.base_path else file_name

        if file_path and not file_path.startswith('/'):
            file_path = '/' + file_path

        # Write the part's XML content to the zip file
        if file_path is not None and self.data is not None:
            part_xml = self.data.to_xml(writer, {})
            part_xml_string = etree.tostring(part_xml, encoding='utf-8', xml_declaration=True)
            writer.write_file(file_path, part_xml_string)

        writer.add_written_part(self)

        self._relationships_to_xml(writer, relationships)

        for relationship in relationships:
            relationship.target.to_file(writer)

    def _parse_data(self, parser: 'Parser', file_path: str | None = None):
        if self._has_relationship_file(parser, file_path):
            self._parse_relationships(parser, file_path)

        file_xml = self._get_file_xml(parser, file_path)

        if file_xml is not None:
            self._parse_xml(parser, file_path, file_xml)

    def _parse_xml(self, parser: 'Parser', file_path: str | None, xml: etree._Element):
        self.data = Attribute.from_xml(parser, file_path, xml)

    def _parse_relationships(self, parser: 'Parser', file_path: str | None):
        relationship_file_path = self._get_relationship_file_path(file_path=file_path)
        relationships = Relationship.from_file(parser, relationship_file_path, self)
        self.relationships = relationships

    def _get_file_xml(self, parser: 'Parser', file_path: str | None):
        file_path = self._get_file_path() if file_path is None else file_path

        if not file_path.lstrip('/'):
            return None

        file_data_string = parser.read_file(file_path.lstrip('/'))
        file_xml = etree.fromstring(file_data_string)
        return file_xml

    def _get_relationship_file_path(self, file_path: str | None = None) -> str:
        file_path = (self._get_file_path() if file_path is None else file_path) or ''
        path_elements = file_path.lstrip('/').split('/')
        path = ('/'.join(path_elements[:-1]) + '/_rels/' + path_elements[-1] + '.rels').lstrip('/')

        if not path.startswith('/'):
            path = '/' + path

        return path

    def _has_relationship_file(self, parser: 'Parser', file_path: str | None) -> bool:
        relationship_file_path = self._get_relationship_file_path(file_path=file_path).lstrip('/')
        return relationship_file_path in parser.zip_file.namelist()

    def _relationships_to_xml(self, writer: 'Writer', relationships: list['Relationship']):
        relationships_element = etree.Element('Relationships', xmlns="http://schemas.openxmlformats.org/package/2006/relationships")

        for relationship in relationships:
            relationship_xml = relationship._to_xml(writer)
            relationships_element.append(relationship_xml)

        relationship_xml_string = etree.tostring(relationships_element, encoding='utf-8', xml_declaration=True)

        file_name = writer.assign_part_index(self.part_name, self)

        file_path = f"{self.base_path}/{file_name}" if self.base_path else file_name

        if file_path and not file_path.startswith('/'):
            file_path = '/' + file_path

        relationship_file_path = self._get_relationship_file_path(file_path=file_path)
        writer.write_file(relationship_file_path, relationship_xml_string)
