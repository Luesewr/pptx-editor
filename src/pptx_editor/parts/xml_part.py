from io import BytesIO
from pathlib import PurePosixPath
from typing import TYPE_CHECKING

from lxml import etree

from pptx_editor.part import Part
from pptx_editor.attribute import Attribute

if TYPE_CHECKING:
    from pptx_editor.parser import _OOXMLParser
    from pptx_editor.writer import _OOXMLWriter

class XmlPart(Part):
    default_content_type: str | None = "application/xml"
    default_base_path: PurePosixPath | None
    default_part_name: str | None
    default_attribute_name: str | None = None

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.data: Attribute | None = None

    def to_file(self, writer: '_OOXMLWriter'):
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
            buffer = BytesIO()
            buffer.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'.encode('utf-8'))
            self.data.to_xml(writer, buffer)
            writer.write_file(file_path, buffer.getvalue())

        writer.add_written_part(self)

        if len(self.relationships) > 0:
            self.write_relationships_file(writer)

    def _parse_data(self, parser: '_OOXMLParser', file_path: PurePosixPath):
        if self._has_relationship_file(parser, file_path):
            self._parse_relationships(parser, file_path)

        file_xml, ns_declarations = self._get_file_xml(parser, file_path)

        if file_xml is not None:
            self._parse_xml(parser, file_path, file_xml, ns_declarations)
        else:
            self.data = None

    def _parse_xml(self, parser, file_path, xml, ns_declarations):
        self.data = Attribute.from_xml(parser, file_path, xml, ns_declarations)

    def _get_file_xml(self, parser, file_path):
        file_path = self._get_file_path() if file_path is None else file_path
        if not file_path:
            return None, None

        file_data_bytes = parser.read_file(file_path)

        ns_declarations = {}
        pending_ns = []

        context = etree.iterparse(
            BytesIO(file_data_bytes),
            events=('start-ns', 'start'),
        )

        for event, data in context:
            data: etree._Element
            if event == 'start-ns':
                pending_ns.append(data)
            elif event == 'start' and pending_ns:
                data_path = data.getroottree().getpath(data)

                ns_declarations[data_path] = {
                    prefix: uri for prefix, uri in pending_ns
                }
                pending_ns = []

        return context.root, ns_declarations
