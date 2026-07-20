from collections import defaultdict
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
            data_element, doc_info = self._get_xml()

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

    def _get_xml(self) -> tuple[etree._Element | None, etree.DocInfo | None]:
        if self._data is None:
            return None, None

        context = etree.iterparse(
            BytesIO(self._data),
            events=('start-ns', 'end-ns', 'start', 'end'),
        )

        root_element = None
        path_stack = []
        child_stack = []
        ns_stack = []
        ns_queue = {}

        ns_map = defaultdict(list)

        for event, value in context:

            if event == "start-ns":
                prefix, uri = value
                ns_queue[prefix] = uri
                ns_map[uri].append(prefix)

            elif event == "start":
                path_stack.append(value)
                ns_stack.append(ns_queue)
                child_stack.append([])
                ns_queue = {}

            elif event == "end":
                namespace, prefix, name = self._process_tag(value.tag, ns_map)

                element = path_stack.pop()
                declared_namespaces = ns_stack.pop()
                children = tuple(child_stack.pop())

                parsed_attributes = tuple(
                    parser._OOXMLParser.parse_attribute_from_item(self, element.nsmap, name, str(key), str(value))
                    for key, value
                    in element.attrib.items()
                )

                parsed_element = parser._OOXMLParser.parse_element_from_xml(self, element, parsed_attributes, children, declared_namespaces)

                if len(path_stack) > 0:
                    child_stack[-1].append(parsed_element)
                else:
                    root_element = parsed_element

            elif event == "end-ns":
                if value is None:
                    continue

                prefix, uri = value
                if uri in ns_map and prefix in ns_map[uri]:
                    ns_map[uri].remove(prefix)


        return root_element, context.root.getroottree().docinfo

    @staticmethod
    def _process_tag(raw_tag, ns_map: dict[str, list[str]]) -> tuple[str | None, str | None, str]:
        if raw_tag.startswith("{"):
            uri, tag = raw_tag[1:].split("}", 1)
            prefixes = ns_map.get(uri)
            if prefixes:
                prefix = prefixes[-1]
                return uri, (prefix if prefix else None), tag
            return uri, None, tag
        return None, None, raw_tag
