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
    default_base_path: PurePosixPath | None = None
    default_part_name: str | None = None
    default_extension: str | None = 'xml'

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

    def copy(self):
        new_part = self.__class__(self.main_part, self._get_file_path(), self.content_type, self.is_default)
        new_part.relationships = [r.copy(part=new_part) for r in self.relationships]

        data = self.data

        relationship_map = {r: new_r for r, new_r in zip(self.relationships, new_part.relationships)}

        new_part._data = data.copy(part=new_part, relationship_map=relationship_map)
        new_part.docinfo = self.docinfo
        new_part._is_data_parsed = self._is_data_parsed
        new_part.unlock_relationships = self.unlock_relationships

        return new_part

    @property
    def data(self) -> 'XmlElement':
        if not self._is_data_parsed:
            data_element, doc_info = parser._OOXMLParser.parse_part_from_xml(self)

            if data_element is None:
                raise ValueError("Part does not contain data")

            self._data = data_element
            self.docinfo = doc_info

            self._is_data_parsed = True
            self.unlock_relationships = True

        if not isinstance(self._data, XmlElement):
            raise ValueError("Part data is not an XmlElement")

        return self._data

    def _to_file(self, writer: '_OOXMLWriter'):
        if writer.is_part_written(self):
            return

        file_name = writer.assign_part_index(self)
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
            writer.write_file(file_path, buffer.getvalue())

        writer.add_written_part(self)

        writer.assign_relation_part_indexes(self.relationships)

        if len(self.relationships) > 0:
            self._write_relationships_file(writer)
