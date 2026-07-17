from io import BytesIO
from pathlib import PurePosixPath
from typing import TYPE_CHECKING

from lxml import etree

import pptx_editor.parser as parser
from pptx_editor.part import Part
from pptx_editor.xml_element import XmlElement

if TYPE_CHECKING:
    from pptx_editor.writer import _OOXMLWriter

class XmlPart(Part):
    default_content_type: str | None = "application/xml"
    default_base_path: PurePosixPath | None
    default_part_name: str | None
    default_attribute_name: str | None = None

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self._data: XmlElement | bytes | None = None
        self.docinfo: 'etree.DocInfo' | None = None
        self._is_data_parsed: bool = False
        self.unlock_relationships: bool = False

    def get_element(self, name: str, prefix: str | None = None) -> 'XmlElement | None':
        return self.data.get_element(name, prefix)

    def get_element_by_type(self, element_type: type['XmlElement']) -> 'XmlElement | None':
        return self.data.get_element_by_type(element_type)

    def get_elements(self, name: str, prefix: str | None = None) -> list['XmlElement']:
        return self.data.get_elements(name, prefix)

    def get_elements_by_type(self, element_type: type['XmlElement']) -> list['XmlElement']:
        return self.data.get_elements_by_type(element_type)

    @property
    def data(self) -> 'XmlElement':
        if not self._is_data_parsed:
            file_xml, ns_declarations = self._get_xml()

            if file_xml is None:
                raise ValueError("Part does not contain data")

            self._data = parser._OOXMLParser.parse_element_from_xml(self, file_xml, ns_declarations)
            self.docinfo = file_xml.getroottree().docinfo

            self._is_data_parsed = True
            self.unlock_relationships = True

        if not isinstance(self._data, XmlElement):
            raise ValueError("Part data is not an XmlElement")

        return self._data

    def _to_file(self, writer: '_OOXMLWriter'):
        if writer.is_part_written(self):
            return

        file_name = writer.assign_part_index(self.part_name, self)
        file_path = PurePosixPath(self.base_path) / file_name if self.base_path else file_name

        if not file_path.is_absolute():
            file_path = PurePosixPath('/') / file_path

        # Write the part's XML content to the zip file
        if self._data is not None:
            buffer = BytesIO()

            if self._is_data_parsed:
                if self.docinfo and self.docinfo.standalone is not None:
                    buffer.write(f'<?xml version="1.0" encoding="UTF-8" standalone="{"yes" if self.docinfo.standalone else "no"}"?>\n'.encode('utf-8'))

                self._data._to_xml(writer, buffer)
            else:
                buffer.write(self._data)
                self.relationships.sort(key=lambda r: r.original_id.removeprefix('rId'))
            writer.write_file(file_path, buffer.getvalue())

        writer.assign_relation_part_indexes(self.relationships)

        writer.add_written_part(self)

        if len(self.relationships) > 0:
            self._write_relationships_file(writer)

    def _get_xml(self) -> tuple[etree._Element | None, dict[str, dict[str | None, str]] | None]:
        if self._data is None:
            return None, None

        ns_declarations = {}
        pending_ns = []

        context = etree.iterparse(
            BytesIO(self._data),
            events=('start-ns', 'start'),
        )

        declaration_entries = []

        for event, data in context:
            data: etree._Element
            if event == 'start-ns':
                pending_ns.append(data)
            elif event == 'start' and pending_ns:
                declaration_entries.append((data, pending_ns.copy()))
                pending_ns = []

        for data, pending_ns in declaration_entries:
            data_path = data.getroottree().getpath(data)
            ns_declarations[data_path] = dict(pending_ns)


        return context.root, ns_declarations
